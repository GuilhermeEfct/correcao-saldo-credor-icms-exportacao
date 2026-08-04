# -*- coding: utf-8 -*-
"""
routes.py — Rotas da ferramenta "Correção do Saldo Credor de ICMS proporcional às
Exportações".

Padrão HUB (maio/2026):
  - NÃO cria o app Flask. Exporta um Blueprint que o HUB registra.
  - NÃO faz login próprio: autenticação e permissão herdadas do HUB.
  - Processamento em SEGUNDO PLANO (lote + protocolo + polling), porque a
    plataforma derruba requisição acima de ~240 s.
  - Download de USO ÚNICO, com expiração.

🔴 "Não há oportunidade" NÃO É ERRO. Os desfechos "sem saldo credor" e "sem
exportação" percorrem o caminho de SUCESSO: job termina `concluido`, a rota
devolve 200 com o resultado completo e o download do Excel fica disponível. Só
`logica.ErroDeNegocio` e falha inesperada viram `status="erro"`.

Como o HUB registra:
    from correcao_saldo_credor_icms.routes import bp, init_app, set_auth_provider
    init_app(app)
    set_auth_provider(meu_auth)      # (request) -> (ok: bool, contexto: dict)
    app.register_blueprint(bp)
"""

from __future__ import annotations

import logging
import os
import re
import secrets
import tempfile
import threading
import time
import uuid
from functools import wraps

from flask import Blueprint, g, jsonify, request, send_file
from werkzeug.utils import secure_filename

import logica

TOOL_ID = "correcao-saldo-credor-icms-exportacao"
# Chave em users.json -> permissions. É a do GRUPO, não uma chave própria da
# ferramenta: quem tem acesso ao card do grupo ICMS tem acesso à ferramenta. Mesmo
# desenho do C170/C175 com `pis_cofins`. Chave própria para card interno já custou
# dois deploys na Auditoria de Fretes.
PERMISSAO = "icms"

bp = Blueprint(TOOL_ID, __name__, url_prefix="/tools/%s" % TOOL_ID)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
WORK_DIR = os.path.join(tempfile.gettempdir(), TOOL_ID)
os.makedirs(WORK_DIR, exist_ok=True)

log = logging.getLogger(TOOL_ID)

# Teto de UMA requisição (um arquivo por vez no envio em lote) e do lote inteiro.
# O relatório de apuração é pequeno (dezenas a centenas de KB por
# estabelecimento); a folga existe para quem manda tudo num .zip.
MAX_CONTEUDO_MB = int(os.environ.get("CORRECAO_SALDO_MAX_MB", "200"))
MAX_CONTEUDO = MAX_CONTEUDO_MB * 1024 * 1024
MAX_LOTE_MB = int(os.environ.get("CORRECAO_SALDO_MAX_LOTE_MB", "1024"))
MAX_LOTE = MAX_LOTE_MB * 1024 * 1024

DOWNLOAD_TTL = 600                      # 10 min de validade do link
JOB_TTL = 3600                          # jobs expiram após 1 h

# Teto de jobs pesados simultâneos: evita que vários lotes ao mesmo tempo
# estourem a memória da instância.
_proc_sema = threading.Semaphore(int(os.environ.get("CORRECAO_SALDO_MAX_JOBS", "2")))

EXTENSOES_OK = (".csv", ".txt", ".zip")

MSG_LIMITE = (
    "O envio excedeu o limite de %d MB por requisição. Mande os arquivos em lote (a tela "
    "faz isso sozinha, um por vez) ou compacte em .zip — CSV comprime cerca de 10x e a "
    "ferramenta lê o zip."
)
MSG_LIMITE_LOTE = (
    "O lote chegou a %d MB e o teto é %d MB. Divida a análise ou peça ao admin para "
    "elevar CORRECAO_SALDO_MAX_LOTE_MB."
)

_JOBS = {}
_DOWNLOADS = {}
_LOCK = threading.Lock()


# =============================================================================
# Autenticação e permissão herdadas do HUB (fail-closed)
# =============================================================================
_auth_provider = None


def set_auth_provider(fn):
    """O HUB registra aqui uma função (request) -> (ok: bool, contexto: dict)."""
    global _auth_provider
    _auth_provider = fn


def init_app(app):
    app.config.setdefault("CORRECAO_SALDO_WORK_DIR", WORK_DIR)
    return bp


def _contexto_usuario():
    """(ok, contexto) vindo do HUB. Sem provider registrado, NEGA.

    Este módulo não tem caminho próprio de liberação — nem por `TESTING`. A liberação
    de desenvolvimento mora no `test_app.py`, que declara não ir para produção e
    registra um provider explícito.
    """
    if _auth_provider is None:
        return False, {}
    try:
        ok, ctx = _auth_provider(request)
        return bool(ok), (ctx or {})
    except Exception:              # noqa: BLE001 — provider quebrado nega acesso
        log.exception("auth_provider falhou")
        return False, {}


def requer_permissao(f):
    """Autenticação -> 401; permissão do grupo -> 403. Em JSON: página HTML de erro
    quebraria o `res.json()` da tela.

    A checagem NEGA por padrão e só concede quando o dict concede. A forma anterior
    (`if isinstance(p, dict) and not p.get(...)`) tinha fail-open: contexto sem a chave
    `permissions`, ou com tipo inesperado, pulava o `if` inteiro e deixava entrar
    qualquer usuário autenticado.
    """
    @wraps(f)
    def _wrap(*a, **kw):
        ok, ctx = _contexto_usuario()
        if not ok:
            return jsonify({"ok": False, "erro": "Não autenticado."}), 401
        permissoes = ctx.get("permissions")
        if not isinstance(permissoes, dict):
            permissoes = {}
        if not (ctx.get("is_admin") or permissoes.get(PERMISSAO)):
            return jsonify({"ok": False,
                            "erro": "Sem permissão para esta ferramenta."}), 403
        g.correcao_saldo_user = (ctx.get("login") or ctx.get("email")
                                 or ctx.get("usuario") or "desconhecido")
        return f(*a, **kw)
    return _wrap


def _usuario():
    return getattr(g, "correcao_saldo_user", "") or ""


@bp.errorhandler(413)
def _payload_grande(_erro):
    """JSON quando o corpo estoura MAX_CONTENT_LENGTH. Sem isto o Flask devolve a
    página HTML padrão e a tela mostraria "falha de comunicação" em vez do motivo."""
    return jsonify({"ok": False, "erro": MSG_LIMITE % MAX_CONTEUDO_MB,
                    "limite_mb": MAX_CONTEUDO_MB}), 413


# =============================================================================
# Área de trabalho do lote
# =============================================================================
def _dir_lote(protocolo):
    return os.path.join(WORK_DIR, "lote_" + protocolo)


def _apagar_lote(protocolo):
    pasta = _dir_lote(protocolo)
    try:
        for nome in os.listdir(pasta):
            try:
                os.remove(os.path.join(pasta, nome))
            except OSError:
                pass
        os.rmdir(pasta)
    except OSError:
        pass


def _job_do_usuario(protocolo):
    """Job existente E do próprio usuário — protocolo de terceiro não vaza."""
    job = _JOBS.get(protocolo)
    if not job or job.get("user") != _usuario():
        return None
    return job


def _limpar_expirados():
    agora = time.time()
    with _LOCK:
        vencidos = [k for k, v in _JOBS.items() if agora - v["ts"] > JOB_TTL]
        for pid in vencidos:
            _JOBS.pop(pid, None)
        for tk in [k for k, v in _DOWNLOADS.items() if agora > v["expira"]]:
            item = _DOWNLOADS.pop(tk, None)
            if item:                    # planilha nunca baixada não fica no disco
                try:
                    os.remove(item["arquivo"])
                except OSError:
                    pass
    for pid in vencidos:
        _apagar_lote(pid)


# =============================================================================
# Rotas
# =============================================================================
@bp.route("/limites", methods=["GET"])
@requer_permissao
def limites():
    """Tetos de envio, para a tela avisar ANTES de subir os arquivos."""
    return jsonify({"ok": True, "limite_mb": MAX_CONTEUDO_MB,
                    "limite_lote_mb": MAX_LOTE_MB, "envio_em_lote": True,
                    "extensoes": list(EXTENSOES_OK)})


@bp.route("/lote", methods=["POST"])
@requer_permissao
def lote_abrir():
    """Abre um protocolo com a empresa DECLARADA pelo analista (CNPJ + razão social).

    Padrão "CNPJ primeiro" (§4.4): o analista declara qual empresa está analisando e a
    ferramenta confere os arquivos contra essa declaração.
    """
    _limpar_expirados()
    cnpj = re.sub(r"\D", "", request.form.get("cnpj", "") or "")
    razao_social = (request.form.get("razao_social", "") or "").strip()
    if len(cnpj) != 14:
        return jsonify({"ok": False,
                        "erro": "Informe e consulte o CNPJ da empresa antes de "
                                "calcular."}), 400

    protocolo = uuid.uuid4().hex[:12]
    os.makedirs(_dir_lote(protocolo), exist_ok=True)
    with _LOCK:
        _JOBS[protocolo] = {
            "user": _usuario(), "status": "recebendo", "pct": 0,
            "msg": "Aguardando arquivos…", "erro": None, "ts": time.time(),
            "iniciado_at": time.time(), "cnpj": cnpj, "razao_social": razao_social,
            "arquivos": [], "bytes": 0, "arquivo": None,
        }
    return jsonify({"ok": True, "protocolo": protocolo,
                    "limite_mb": MAX_CONTEUDO_MB, "limite_lote_mb": MAX_LOTE_MB}), 201


@bp.route("/lote/<protocolo>/arquivo", methods=["POST"])
@requer_permissao
def lote_arquivo(protocolo):
    """Recebe UM arquivo do lote e grava no disco, sem carregar na memória."""
    with _LOCK:
        job = _job_do_usuario(protocolo)
        if not job:
            return jsonify({"ok": False,
                            "erro": "Protocolo não encontrado ou expirado."}), 404
        if job["status"] != "recebendo":
            return jsonify({"ok": False, "erro": "Este lote já foi iniciado."}), 409
        acumulado, indice = job["bytes"], len(job["arquivos"])

    # Teto por requisição checado AQUI, não via MAX_CONTENT_LENGTH: o HUB não define
    # esse valor globalmente — e não pode, porque todas as ferramentas dividem o mesmo
    # app e um teto global cortaria o upload de todas. Sem esta checagem, /limites
    # anunciava 200 MB que nada aplicava.
    if (request.content_length or 0) > MAX_CONTEUDO:
        return jsonify({"ok": False, "erro": MSG_LIMITE % MAX_CONTEUDO_MB,
                        "limite_mb": MAX_CONTEUDO_MB}), 413

    fs = request.files.get("arquivo")
    if fs is None or not fs.filename:
        return jsonify({"ok": False, "erro": "Nenhum arquivo enviado."}), 400

    # a extensão decide como o arquivo é lido (zip x texto), então é preservada
    ext = os.path.splitext(fs.filename)[1].lower()
    if ext not in EXTENSOES_OK:
        return jsonify({"ok": False,
                        "erro": "O arquivo %s não é .csv, .txt nem .zip." %
                                os.path.basename(fs.filename)}), 400

    seguro = secure_filename(fs.filename) or ("parte_%03d%s" % (indice, ext))
    if not seguro.lower().endswith(ext):
        seguro += ext
    destino = os.path.join(_dir_lote(protocolo), "%03d_%s" % (indice, seguro))
    try:
        fs.save(destino)                # streaming direto para o disco
    except OSError as e:
        return jsonify({"ok": False, "erro": "Falha ao gravar o arquivo: %s" % e}), 500

    tamanho = os.path.getsize(destino)
    # segunda checagem, sobre o tamanho REAL: o header Content-Length pode vir ausente
    # ou mentiroso, e aí a checagem anterior não teria pegado nada
    if tamanho > MAX_CONTEUDO:
        os.remove(destino)
        return jsonify({"ok": False, "erro": MSG_LIMITE % MAX_CONTEUDO_MB,
                        "limite_mb": MAX_CONTEUDO_MB}), 413
    if acumulado + tamanho > MAX_LOTE:
        os.remove(destino)
        return jsonify({"ok": False,
                        "erro": MSG_LIMITE_LOTE % ((acumulado + tamanho) // (1024 * 1024),
                                                   MAX_LOTE_MB),
                        "limite_lote_mb": MAX_LOTE_MB}), 413

    with _LOCK:
        job = _job_do_usuario(protocolo)
        if not job:
            os.remove(destino)
            return jsonify({"ok": False,
                            "erro": "Protocolo expirado durante o envio."}), 404
        job["arquivos"].append(destino)
        job["bytes"] += tamanho
        job["ts"] = time.time()
        job["msg"] = "Recebidos %d arquivo(s)…" % len(job["arquivos"])
        recebidos, total_bytes = len(job["arquivos"]), job["bytes"]

    return jsonify({"ok": True, "recebidos": recebidos, "bytes": total_bytes})


@bp.route("/lote/<protocolo>/identificar", methods=["POST"])
@requer_permissao
def lote_identificar(protocolo):
    """De quem são os relatórios já enviados — sem processá-los.

    É CONFERÊNCIA, não identificação da empresa: diz quantos estabelecimentos vieram,
    quais arquivos não são apuração, e barra o lote quando os relatórios são de outra
    empresa. Quem declara a empresa é o analista, no campo CNPJ da tela.
    """
    with _LOCK:
        job = _job_do_usuario(protocolo)
        if not job:
            return jsonify({"ok": False,
                            "erro": "Protocolo não encontrado ou expirado."}), 404
        if job["status"] != "recebendo":
            return jsonify({"ok": False, "erro": "Este lote já foi iniciado."}), 409
        if not job["arquivos"]:
            return jsonify({"ok": False,
                            "erro": "Nenhum arquivo recebido neste lote."}), 400
        caminhos = list(job["arquivos"])

    try:
        info = logica.identificar_estabelecimentos(caminhos)
    except logica.ErroDeNegocio as erro:
        return jsonify({"ok": False, "erro": str(erro)}), 400
    except Exception:                       # noqa: BLE001
        log.exception("Falha ao identificar a empresa (protocolo %s)", protocolo)
        return jsonify({"ok": False,
                        "erro": "Não foi possível ler os relatórios enviados. Confira "
                                "se são os relatórios de Apuração de ICMS."}), 500

    if not info["estabelecimentos"]:
        return jsonify({"ok": False,
                        "erro": "Nenhum relatório de Apuração de ICMS (ICMSProprio) foi "
                                "reconhecido nos arquivos enviados."}), 400

    with _LOCK:
        job = _job_do_usuario(protocolo)
        if job:
            job["identificado"] = info["matriz"]
            job["ts"] = time.time()

    return jsonify({"ok": True, "matriz": info["matriz"],
                    "estabelecimentos": info["estabelecimentos"],
                    "raizes": info["raizes"], "erro_raizes": info["erro"],
                    "nao_reconhecidos": [re.sub(r"^\d{3}_", "", n)
                                         for n in info["nao_reconhecidos"]]})


@bp.route("/lote/<protocolo>/iniciar", methods=["POST"])
@requer_permissao
def lote_iniciar(protocolo):
    """Fecha o lote e dispara o processamento em segundo plano."""
    cnpj_form = re.sub(r"\D", "", request.form.get("cnpj", "") or "")
    razao_form = (request.form.get("razao_social", "") or "").strip()

    with _LOCK:
        job = _job_do_usuario(protocolo)
        if not job:
            return jsonify({"ok": False,
                            "erro": "Protocolo não encontrado ou expirado."}), 404
        if job["status"] != "recebendo":
            return jsonify({"ok": False, "erro": "Este lote já foi iniciado."}), 409
        if not job["arquivos"]:
            return jsonify({"ok": False,
                            "erro": "Nenhum arquivo recebido neste lote."}), 400
        caminhos = list(job["arquivos"])
        cnpj = cnpj_form or job["cnpj"]
        razao_social = razao_form or job["razao_social"]

        # O CNPJ é DECLARADO pelo analista (padrão "CNPJ primeiro", §4.4) e os arquivos
        # são conferidos contra essa declaração. Recusar aqui é o que faz a regra valer:
        # se o backend adotasse a matriz encontrada nos arquivos, qualquer chamada direta
        # à API produziria planilha com empresa adivinhada — e um lote trocado sairia
        # coerente consigo mesmo e errado em relação ao caso.
        if len(cnpj) != 14:
            return jsonify({"ok": False,
                            "erro": "Informe e consulte o CNPJ da empresa antes de "
                                    "calcular."}), 400

        job.update(cnpj=cnpj, razao_social=razao_social, status="processando", pct=0,
                   msg="Na fila…", ts=time.time(), iniciado_at=time.time())

    threading.Thread(target=_rodar_job,
                     args=(protocolo, caminhos, cnpj, razao_social),
                     daemon=True).start()
    return jsonify({"ok": True, "protocolo": protocolo, "status": "processando"}), 202


@bp.route("/status/<protocolo>", methods=["GET"])
@requer_permissao
def status(protocolo):
    with _LOCK:
        job = _job_do_usuario(protocolo)
        if not job:
            return jsonify({"ok": False,
                            "erro": "Processamento não encontrado ou expirado."}), 404
        job["ts"] = time.time()
        resp = {"ok": True, "status": job["status"], "pct": job["pct"],
                "msg": job["msg"],
                "decorrido_s": int(time.time() - job["iniciado_at"])}
        if job["status"] == "concluido":
            # 200 com o resultado inteiro — inclusive quando NENHUM
            # estabelecimento qualificou: aquilo é conclusão, não falha
            resp["download_token"] = job.get("download_token")
            resp["resultado"] = job.get("resultado")
        if job["status"] == "erro":
            resp["erro"] = job["erro"]
    return jsonify(resp)


@bp.route("/download/<token>", methods=["GET"])
@requer_permissao
def download(token):
    """Download de USO ÚNICO e com EXPIRAÇÃO."""
    with _LOCK:
        item = _DOWNLOADS.get(token)
        if not item or item["user"] != _usuario():
            return jsonify({"ok": False, "erro": "Link inválido ou já utilizado."}), 404
        if item["usado"] or time.time() > item["expira"]:
            _DOWNLOADS.pop(token, None)
            return jsonify({"ok": False, "erro": "Link expirado ou já utilizado."}), 410
        item["usado"] = True
        arquivo = item["arquivo"]

    return send_file(
        arquivo, as_attachment=True,
        download_name="Correcao_Saldo_Credor_ICMS_Exportacao.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


# =============================================================================
# Execução do job
# =============================================================================
def _payload_para_tela(resultado: dict) -> dict:
    """O que a tela precisa, arredondado. A explicação vai literal — é a MESMA
    string do Excel e do JSON, nunca uma segunda redação."""
    def _r(v):
        return None if v is None else round(v, 2)

    estabs = []
    for e in resultado["estabelecimentos"]:
        estabs.append({
            "cnpj": e["cnpj"], "cnpj_fmt": e["cnpj_fmt"], "uf": e["uf"],
            "mes_ref": e["mes_ref"], "faturamento": _r(e["faturamento"]),
            "exportacao": _r(e["exportacao"]), "percentual": e["percentual"],
            "saldo_credor": _r(e["saldo_credor"]), "correcao": _r(e["correcao"]),
            "status": e["status"], "rotulo_status": e["rotulo_status"],
            "explicacao": e["explicacao"],
            "serie": [{"mes": r["mes"], "faturamento": _r(r["faturamento"]),
                       "exportacao": _r(r["exportacao"]),
                       "saldo_credor": _r(r["saldo_credor"])} for r in e["serie"]],
        })
    return {
        "empresa": resultado["empresa"],
        "estabelecimentos": estabs,
        "explicacao_consolidada": resultado["explicacao_consolidada"],
        "status_consolidado": resultado["status_consolidado"],
        "totais": {**resultado["totais"],
                   "correcao": _r(resultado["totais"]["correcao"])},
        "periodo": resultado["periodo"],
        "avisos": resultado["avisos"],
        # arquivo que não é apuração não pode desaparecer em silêncio: "nenhuma
        # oportunidade" e "mandei o arquivo errado" são coisas diferentes
        "nao_reconhecidos": [re.sub(r"^\d{3}_", "", n)
                             for n in resultado["nao_reconhecidos"]],
    }


def _rodar_job(protocolo, arquivos, cnpj, razao_social):
    def progress(pct, msg):
        with _LOCK:
            j = _JOBS.get(protocolo)
            if j:
                j["pct"], j["msg"], j["ts"] = pct, msg, time.time()

    try:
        with _proc_sema:                # o trabalho pesado roda aqui
            resultado = logica.processar_saldo_credor(
                arquivos, cnpj, razao_social, progress=progress)
            progress(92, "Gerando a planilha auditável…")
            saida = os.path.join(WORK_DIR, "%s.xlsx" % protocolo)
            logica.gerar_excel_auditavel(resultado, saida)
    except logica.ErroDeNegocio as erro:
        # falha REAL: arquivo ilegível, CNPJ divergente, seção ausente
        with _LOCK:
            j = _JOBS.get(protocolo)
            if j:
                j.update(status="erro", erro=str(erro), ts=time.time())
        _apagar_lote(protocolo)
        _limpar_expirados()
        return
    except Exception:                   # noqa: BLE001
        log.exception("Erro inesperado em %s (protocolo %s)", TOOL_ID, protocolo)
        with _LOCK:
            j = _JOBS.get(protocolo)
            if j:
                j.update(status="erro", ts=time.time(),
                         erro="Erro inesperado ao processar os arquivos. Confira se são "
                              "os relatórios de Apuração de ICMS e tente de novo; se "
                              "persistir, avise o admin.")
        _apagar_lote(protocolo)
        _limpar_expirados()
        return

    # SUCESSO — incluindo os desfechos "sem saldo credor" e "sem exportação"
    token = secrets.token_urlsafe(16)
    with _LOCK:
        _DOWNLOADS[token] = {"arquivo": saida, "user": _JOBS.get(protocolo, {}).get("user"),
                             "expira": time.time() + DOWNLOAD_TTL, "usado": False}
        j = _JOBS.get(protocolo)
        if j:
            j.update(status="concluido", pct=100, msg="Concluído.", arquivo=saida,
                     download_token=token, resultado=_payload_para_tela(resultado),
                     ts=time.time())
    _apagar_lote(protocolo)
    _limpar_expirados()
