/* screen.js — JavaScript da ferramenta "Correção do Saldo Credor de ICMS
 * proporcional às Exportações".
 *
 * O HUB chama initCorrecaoSaldoCredorIcms() quando a tela é montada. Todo o código
 * fica escopado ao elemento #tool-correcao-saldo-credor-icms, e cada função e id
 * carrega o prefixo `cscie` para não colidir com outras ferramentas.
 *
 * Base das rotas: o HUB registra o blueprint em
 * /tools/correcao-saldo-credor-icms-exportacao. Se montar em outro prefixo, defina
 * window.HUB_TOOL_BASE.
 */
(function () {
  "use strict";

  function initCorrecaoSaldoCredorIcms(rootParam) {
    var root = rootParam || document.getElementById("tool-correcao-saldo-credor-icms");
    if (!root) return;
    // flag no próprio elemento: se o HUB remontar a tela, o novo nó é inicializado
    // (uma flag de módulo deixaria a segunda montagem sem eventos)
    if (root.getAttribute("data-cscie-init") === "1") return;
    root.setAttribute("data-cscie-init", "1");

    var BASE = window.HUB_TOOL_BASE || "/tools/correcao-saldo-credor-icms-exportacao";
    var API = window.API || "";
    var $ = function (sel) { return root.querySelector(sel); };

    var elBtnLimpar = $("#cscieBtnLimpar");
    var elCnpj = $("#cscieCnpj");
    var elRazao = $("#cscieRazao");
    var elCnpjMsg = $("#cscieCnpjMsg");
    var elArquivos = $("#cscieArquivos");
    var elDropzone = $("#cscieDropzone");
    var elLista = $("#cscieLista");
    var elBtn = $("#cscieBtnCalcular");
    var elBtnAviso = $("#cscieBtnAviso");
    var elProgresso = $("#cscieProgresso");
    var elProtocolo = $("#cscieProtocolo");
    var elDecorrido = $("#cscieDecorrido");
    var elBarFill = $("#cscieBarFill");
    var elProgMsg = $("#cscieProgressoMsg");
    var elResultado = $("#cscieResultado");
    var elVeredito = $("#cscieVeredito");
    var elVereditoValor = $("#cscieVereditoValor");
    var elVereditoTexto = $("#cscieVereditoTexto");
    var elBtnCopiar = $("#cscieBtnCopiar");
    var elAvisos = $("#cscieAvisos");
    var elKpis = $("#cscieKpis");
    var elTabelaCorpo = $("#cscieTabelaCorpo");
    var elTabelaRodape = $("#cscieTabelaRodape");
    var elSeletor = $("#cscieSeletor");
    var elSerieCorpo = $("#cscieSerieCorpo");
    var elDownload = $("#cscieDownload");
    var elErro = $("#cscieErro");
    var elErroMsg = $("#cscieErroMsg");

    var LIMITE_MB = 200;         // teto de UMA requisição (um arquivo)
    var LIMITE_LOTE_MB = 1024;   // teto do lote inteiro
    var estado = {cnpj: "", razao: "", manual: false, timer: null,
                  resultado: null, token: null};

    // ---------------------------------------------------------------- helpers
    function cscieHeaders(extra) {
      var h = {}, k;
      // o HUB expõe authHeaders() (Bearer). Fora dele, a sessão vai por cookie.
      try {
        if (typeof window.authHeaders === "function") {
          var a = window.authHeaders() || {};
          for (k in a) if (Object.prototype.hasOwnProperty.call(a, k)) h[k] = a[k];
        }
      } catch (e) { /* sem helper: segue com credentials same-origin */ }
      if (extra) {
        for (k in extra) if (Object.prototype.hasOwnProperty.call(extra, k)) h[k] = extra[k];
      }
      return h;
    }

    function cscieFetch(url, opcoes) {
      opcoes = opcoes || {};
      opcoes.headers = cscieHeaders(opcoes.headers);
      opcoes.credentials = "same-origin";
      return fetch(url, opcoes);
    }

    // Resposta não-JSON (413/502/504 devolvem HTML do load balancer) não pode virar
    // "Unexpected token '<'": aqui o motivo real chega à tela.
    function cscieLerJson(resp) {
      if (resp.status === 401) {
        if (typeof window.logout === "function") { window.logout(); }
        throw new Error("Sessão expirada. Entre novamente para continuar.");
      }
      var ct = resp.headers.get("content-type") || "";
      if (ct.indexOf("application/json") !== -1) return resp.json();
      return resp.text().then(function (txt) {
        if (resp.status === 413) {
          throw new Error("O envio excedeu o limite de " + LIMITE_MB + " MB aceito pelo " +
            "servidor. Compacte os arquivos em .zip (CSV comprime cerca de 10x e a " +
            "ferramenta lê o zip) ou envie menos arquivos por vez.");
        }
        throw new Error("O servidor respondeu " + resp.status + " sem JSON. Provável falta " +
          "de memória ou tempo esgotado — tente um envio menor ou avise o admin. Início " +
          "da resposta: " + txt.slice(0, 160).replace(/\s+/g, " "));
      });
    }

    function brl(v) {
      if (v === null || v === undefined) return "—";
      return "R$ " + Number(v).toLocaleString("pt-BR",
        {minimumFractionDigits: 2, maximumFractionDigits: 2});
    }

    function pctFmt(f) {
      if (f === null || f === undefined) return "—";
      return (Number(f) * 100).toLocaleString("pt-BR",
        {minimumFractionDigits: 2, maximumFractionDigits: 2}) + "%";
    }

    function formatarMB(mb) {
      return (mb >= 1024 ? (mb / 1024).toFixed(2) + " GB" : mb.toFixed(1) + " MB");
    }

    function td(texto, classe) {
      var c = document.createElement("td");
      c.textContent = texto;
      if (classe) c.className = classe;
      return c;
    }

    // ------------------------------------------------------------- limites
    cscieFetch(BASE + "/limites")
      .then(function (r) { return r.ok ? r.json() : null; })
      .then(function (j) {
        if (!j) return;
        if (j.limite_mb) LIMITE_MB = j.limite_mb;
        if (j.limite_lote_mb) LIMITE_LOTE_MB = j.limite_lote_mb;
        atualizarLista();
      })
      .catch(function () { /* mantém os padrões */ });

    // ------------------------------------------------------- CNPJ (passo 1)
    function mascaraCnpj(valor) {
      var d = (valor || "").replace(/\D/g, "").slice(0, 14);
      if (d.length > 12) return d.replace(/^(\d{2})(\d{3})(\d{3})(\d{4})(\d{0,2})/, "$1.$2.$3/$4-$5");
      if (d.length > 8) return d.replace(/^(\d{2})(\d{3})(\d{3})(\d{0,4})/, "$1.$2.$3/$4");
      if (d.length > 5) return d.replace(/^(\d{2})(\d{3})(\d{0,3})/, "$1.$2.$3");
      if (d.length > 2) return d.replace(/^(\d{2})(\d{0,3})/, "$1.$2");
      return d;
    }

    function razaoDe(dados) {
      if (!dados) return "";
      return dados.razao_social || dados.razaoSocial || dados.nome_empresarial ||
        dados.nome || (dados.company && dados.company.name) || dados.fantasia || "";
    }

    function mostrarRazao(texto, vazia) {
      elRazao.textContent = texto;
      elRazao.className = "cscie-razao" + (vazia ? " cscie-razao-vazia" : "");
    }

    function avisarCnpj(msg) {
      elCnpjMsg.hidden = !msg;
      elCnpjMsg.textContent = msg || "";
    }

    // rota do próprio HUB — a ferramenta não faz consulta externa por conta
    function buscarRazao(digitos) {
      return cscieFetch(API + "/api/cnpj/" + digitos)
        .then(cscieLerJson)
        .then(function (dados) {
          if (dados && dados.erro) throw new Error(dados.erro);
          var razao = razaoDe(dados);
          if (!razao) throw new Error("A consulta não devolveu a razão social.");
          return razao;
        });
    }

    var consultaEmCurso = 0;

    // Digitar o CNPJ é OPCIONAL: serve para sobrescrever a empresa que a ferramenta
    // identifica nos relatórios (ver identificarEmpresa).
    function consultarCnpj() {
      var digitos = elCnpj.value.replace(/\D/g, "");
      estado.cnpj = "";
      estado.razao = "";
      estado.manual = digitos.length > 0;
      if (digitos.length !== 14) {
        mostrarRazao("—", true);
        avisarCnpj(digitos.length ? "O CNPJ precisa de 14 dígitos." : "");
        atualizarBotao();
        return;
      }
      var minha = ++consultaEmCurso;
      mostrarRazao("consultando…", true);
      avisarCnpj("");
      buscarRazao(digitos)
        .then(function (razao) {
          if (minha !== consultaEmCurso) return;      // resposta velha: descarta
          estado.cnpj = digitos;
          estado.razao = razao;
          mostrarRazao(razao, false);
          atualizarBotao();
        })
        .catch(function (e) {
          if (minha !== consultaEmCurso) return;
          mostrarRazao("—", true);
          avisarCnpj("Não foi possível consultar o CNPJ: " +
            (e && e.message ? e.message : "falha na consulta") + ".");
          atualizarBotao();
        });
    }

    elCnpj.addEventListener("input", function () {
      elCnpj.value = mascaraCnpj(elCnpj.value);
      consultarCnpj();
    });

    // ---------------------------------------------------- arquivos (passo 2)
    elArquivos.addEventListener("change", atualizarLista);
    elDropzone.addEventListener("dragover", function (e) { e.preventDefault(); });
    elDropzone.addEventListener("drop", function (e) {
      e.preventDefault();
      elArquivos.files = e.dataTransfer.files;
      atualizarLista();
    });

    function atualizarLista() {
      var fs = elArquivos.files;
      if (!fs || !fs.length) {
        elLista.textContent = "";
        atualizarBotao();
        return;
      }
      var nomes = [], total = 0, maior = 0, nomeMaior = "";
      for (var i = 0; i < fs.length; i++) {
        total += fs[i].size;
        if (fs[i].size > maior) { maior = fs[i].size; nomeMaior = fs[i].name; }
        if (i < 12) nomes.push(fs[i].name);
      }
      var mb = total / 1048576;
      var txt = fs.length + " arquivo(s) · " + formatarMB(mb);
      if (fs.length <= 12) txt += ": " + nomes.join(", ");
      elLista.textContent = txt;

      // cada arquivo sobe numa requisição própria: o que limita é o maior arquivo
      // (por requisição) e a soma (por lote)
      if (maior / 1048576 > LIMITE_MB) {
        falhar("O arquivo \"" + nomeMaior + "\" tem " + formatarMB(maior / 1048576) +
          " e o limite por arquivo é " + LIMITE_MB + " MB. Compacte-o em .zip.");
      } else if (mb > LIMITE_LOTE_MB) {
        falhar("Os arquivos somam " + formatarMB(mb) + " e o teto do lote é " +
          LIMITE_LOTE_MB + " MB. Divida a análise em períodos menores.");
      } else {
        elErro.hidden = true;
      }
      atualizarBotao();
    }

    // O cálculo depende dos relatórios, não de digitação: a empresa é identificada
    // neles. Um CNPJ digitado pela metade é o único caso que trava, porque é
    // ambíguo — não se sabe se o analista quis sobrescrever ou não.
    function atualizarBotao() {
      var temArquivos = elArquivos.files && elArquivos.files.length > 0;
      var digitos = elCnpj.value.replace(/\D/g, "").length;
      var parcial = digitos > 0 && digitos < 14;
      elBtn.disabled = !temArquivos || parcial;
      if (parcial) {
        elBtnAviso.textContent = "Complete o CNPJ ou apague o campo para a ferramenta " +
          "identificar a empresa nos relatórios.";
      } else if (temArquivos) {
        elBtnAviso.textContent = "Pronto para calcular.";
      } else {
        elBtnAviso.textContent = "Envie os relatórios para habilitar o cálculo.";
      }
    }

    // ------------------------------------------------------------- calcular
    elBtn.addEventListener("click", function () {
      var fs = elArquivos.files;
      if (!fs || !fs.length) return;
      var total = 0;
      for (var i = 0; i < fs.length; i++) total += fs[i].size;

      limparSaida();
      elBtn.disabled = true;
      elProgresso.hidden = false;
      elProtocolo.textContent = "—";
      enviarLote(fs, total).catch(function (e) {
        falhar(e && e.message ? e.message
          : "Não foi possível falar com o servidor. Verifique a conexão e tente de novo.");
      });
    });

    // Envio em LOTE: abre o protocolo, sobe um arquivo por requisição e só então
    // manda processar. Um POST único não sobrevive ao corte de ~240 s da plataforma.
    function enviarLote(fs, totalBytes) {
      var abre = new FormData();
      abre.append("cnpj", estado.cnpj);
      abre.append("razao_social", estado.razao);

      elProgMsg.textContent = "Preparando o envio…";
      return cscieFetch(BASE + "/lote", {method: "POST", body: abre})
        .then(cscieLerJson)
        .then(function (d) {
          if (!d.ok) throw new Error(d.erro || "Não foi possível abrir o lote.");
          var protocolo = d.protocolo;
          elProtocolo.textContent = protocolo;

          function proximo(i) {
            if (i >= fs.length) return Promise.resolve(protocolo);
            var fd = new FormData();
            fd.append("arquivo", fs[i]);
            elProgMsg.textContent = "Enviando arquivo " + (i + 1) + " de " + fs.length +
              " — " + fs[i].name + " (" + formatarMB(fs[i].size / 1048576) + ")";
            elBarFill.style.width = Math.round(60 * i / fs.length) + "%";
            return cscieFetch(BASE + "/lote/" + protocolo + "/arquivo",
                              {method: "POST", body: fd})
              .then(cscieLerJson)
              .then(function (r) {
                if (!r.ok) throw new Error(r.erro || "Falha ao enviar " + fs[i].name + ".");
                return proximo(i + 1);
              });
          }

          return proximo(0).then(function () {
            elProgMsg.textContent = "Enviados " + fs.length + " arquivo(s), " +
              formatarMB(totalBytes / 1048576) + ". Identificando a empresa…";
            elBarFill.style.width = "60%";
            return identificarEmpresa(protocolo);
          }).then(function () {
            elProgMsg.textContent = "Iniciando o processamento…";
            var fd = new FormData();
            fd.append("cnpj", estado.cnpj || "");
            fd.append("razao_social", estado.razao || "");
            return cscieFetch(BASE + "/lote/" + protocolo + "/iniciar",
                              {method: "POST", body: fd})
              .then(cscieLerJson)
              .then(function (r) {
                if (!r.ok) throw new Error(r.erro || "Não foi possível iniciar o processamento.");
                acompanhar(protocolo, 0);
              });
          });
        });
    }

    // Descobre de quem são os relatórios e preenche o bloco Empresa. A matriz
    // (ordem 0001) sobe para o topo como identificação da empresa; os demais
    // estabelecimentos aparecem na tabela do resultado.
    function identificarEmpresa(protocolo) {
      return cscieFetch(BASE + "/lote/" + protocolo + "/identificar", {method: "POST"})
        .then(cscieLerJson)
        .then(function (d) {
          if (!d.ok) throw new Error(d.erro || "Não foi possível identificar a empresa.");
          // raízes diferentes no mesmo lote: a planilha sairia com a razão social de
          // uma empresa e o saldo de outra
          if (d.erro_raizes) throw new Error(d.erro_raizes);

          var n = (d.estabelecimentos || []).length;
          var matriz = d.matriz || {};
          var ehMatriz = matriz.cnpj && matriz.cnpj.slice(8, 12) === "0001";
          var resumo = n + (n === 1 ? " estabelecimento identificado"
                                    : " estabelecimentos identificados") +
            " nos relatórios · " + (ehMatriz ? "matriz " : "referência ") +
            (matriz.cnpj_fmt || "—");
          if (d.nao_reconhecidos && d.nao_reconhecidos.length) {
            resumo += " · " + d.nao_reconhecidos.length + " arquivo(s) fora da análise: " +
              d.nao_reconhecidos.join(", ");
          }
          avisarCnpj(resumo);

          // analista digitou o CNPJ: a escolha dele manda, e a validação de raiz
          // no backend continua conferindo os arquivos contra ele
          if (estado.manual && estado.cnpj) return null;
          if (!matriz.cnpj) return null;

          elCnpj.value = matriz.cnpj_fmt || "";
          estado.cnpj = matriz.cnpj;
          mostrarRazao("consultando…", true);
          consultaEmCurso += 1;
          return buscarRazao(matriz.cnpj)
            .then(function (razao) {
              estado.razao = razao;
              mostrarRazao(razao, false);
            })
            .catch(function (e) {
              // sem razão social o cálculo segue: o que ela alimenta é a
              // identificação do relatório, não a conta
              estado.razao = "";
              mostrarRazao("—", true);
              avisarCnpj(resumo + " · não foi possível consultar a razão social (" +
                (e && e.message ? e.message : "falha na consulta") +
                "); o cálculo segue sem ela.");
            });
        });
    }

    // setTimeout RECURSIVO, nunca setInterval: com setInterval duas consultas se
    // sobrepõem quando a resposta demora, e a antiga sobrescreve a nova.
    function acompanhar(protocolo, falhasSeguidas) {
      estado.timer = window.setTimeout(function () {
        cscieFetch(BASE + "/status/" + protocolo)
          .then(cscieLerJson)
          .then(function (d) {
            if (!d.ok) { falhar(d.erro || "Protocolo expirado."); return; }
            // o envio já ocupou os primeiros 60% da barra
            elBarFill.style.width = (60 + Math.round(0.4 * (d.pct || 0))) + "%";
            elProgMsg.textContent = d.msg || "";
            elDecorrido.textContent = (d.decorrido_s || 0) + " s";
            if (d.status === "concluido") {
              // inclui os desfechos SEM oportunidade: são conclusão, não falha
              elProgresso.hidden = true;
              estado.token = d.download_token;
              renderizar(d.resultado);
              elBtn.disabled = false;
              return;
            }
            if (d.status === "erro") { falhar(d.erro || "Erro no processamento."); return; }
            acompanhar(protocolo, 0);
          })
          .catch(function (e) {
            // erro transitório não derruba o acompanhamento; insistência sim
            if (falhasSeguidas >= 4) {
              falhar((e && e.message ? e.message + " " : "") +
                "Perdi o contato com o servidor durante o processamento (protocolo " +
                protocolo + ").");
              return;
            }
            acompanhar(protocolo, falhasSeguidas + 1);
          });
      }, 3000);
    }

    // ------------------------------------------------------------ resultado
    var CLASSE_VEREDITO = {
      CALCULADO: "cscie-veredito cscie-veredito-calculado",
      SEM_EXPORTACAO: "cscie-veredito cscie-veredito-sem-exportacao",
      SEM_SALDO_CREDOR: "cscie-veredito cscie-veredito-sem-saldo"
    };
    var CLASSE_TAG = {
      CALCULADO: "cscie-tag cscie-tag-calculado",
      SEM_EXPORTACAO: "cscie-tag cscie-tag-sem-exportacao",
      SEM_SALDO_CREDOR: "cscie-tag cscie-tag-sem-saldo"
    };

    function renderizar(res) {
      estado.resultado = res;
      elResultado.hidden = false;
      elErro.hidden = true;

      // ---- faixa de veredito: a explicação é o resultado
      elVeredito.className = CLASSE_VEREDITO[res.status_consolidado] || "cscie-veredito";
      var total = res.totais.correcao;
      // sem cálculo, NENHUM valor monetário aqui: um "R$ 0,00" seria lido como
      // "calculamos e deu zero", que é afirmação diferente e errada
      elVereditoValor.hidden = (total === null || total === undefined);
      elVereditoValor.textContent = elVereditoValor.hidden ? "" : brl(total);
      elVereditoTexto.textContent = res.explicacao_consolidada;

      // ---- avisos (inclui arquivo não reconhecido)
      var avisos = (res.avisos || []).slice();
      if (res.nao_reconhecidos && res.nao_reconhecidos.length) {
        avisos.push(res.nao_reconhecidos.length + " arquivo(s) não foram reconhecidos " +
          "como relatório de Apuração de ICMS e ficaram fora da análise: " +
          res.nao_reconhecidos.join(", ") + ".");
      }
      elAvisos.hidden = avisos.length === 0;
      if (avisos.length) {
        elAvisos.innerHTML = "";
        var titulo = document.createElement("strong");
        titulo.textContent = "Atenção";
        var ul = document.createElement("ul");
        avisos.forEach(function (a) {
          var li = document.createElement("li");
          li.textContent = a;
          ul.appendChild(li);
        });
        elAvisos.appendChild(titulo);
        elAvisos.appendChild(ul);
      }

      montarKpis(res);
      montarTabela(res);
      montarSeletor(res);
      elBtnCopiar.textContent = "Copiar explicação";
    }

    function kpi(rotulo, valor, destaque) {
      var div = document.createElement("div");
      div.className = "cscie-kpi" + (destaque ? " cscie-kpi-destaque" : "");
      var r = document.createElement("span");
      r.className = "cscie-kpi-rotulo";
      r.textContent = rotulo;
      var v = document.createElement("div");
      v.className = "cscie-kpi-valor";
      v.textContent = valor;
      div.appendChild(r);
      div.appendChild(v);
      return div;
    }

    function montarKpis(res) {
      elKpis.innerHTML = "";
      var calculados = res.estabelecimentos.filter(function (e) {
        return e.status === "CALCULADO";
      });
      // KPI só existe quando houve cálculo
      if (!calculados.length) { elKpis.hidden = true; return; }
      elKpis.hidden = false;

      if (calculados.length === 1) {
        var e = calculados[0];
        var suf = " (" + e.mes_ref + ")";
        elKpis.appendChild(kpi("Faturamento total" + suf, brl(e.faturamento)));
        elKpis.appendChild(kpi("Faturamento de exportação" + suf, brl(e.exportacao)));
        elKpis.appendChild(kpi("% de exportação" + suf, pctFmt(e.percentual)));
        elKpis.appendChild(kpi("Saldo credor de ICMS" + suf, brl(e.saldo_credor)));
      } else {
        // com vários estabelecimentos, cada um tem o SEU mês: os totais são a soma
        // do último mês de cada um. Não há percentual agregado — ele não existe na
        // metodologia, e exibir um convidaria a interpretá-lo como "a" proporção.
        var soma = function (campo) {
          return calculados.reduce(function (acc, x) { return acc + (x[campo] || 0); }, 0);
        };
        elKpis.appendChild(kpi("Faturamento total somado", brl(soma("faturamento"))));
        elKpis.appendChild(kpi("Exportação somada", brl(soma("exportacao"))));
        elKpis.appendChild(kpi("Saldo credor somado", brl(soma("saldo_credor"))));
      }
      elKpis.appendChild(kpi("Correção", brl(res.totais.correcao), true));
      if (calculados.length > 1) {
        var nota = document.createElement("p");
        nota.className = "cscie-kpi-nota";
        nota.textContent = "Somas do último mês de apuração de cada estabelecimento " +
          "calculado — cada um tem o seu mês, listado na tabela.";
        elKpis.appendChild(nota);
      }
    }

    function montarTabela(res) {
      elTabelaCorpo.innerHTML = "";
      elTabelaRodape.innerHTML = "";
      res.estabelecimentos.forEach(function (e, i) {
        var tr = document.createElement("tr");
        tr.appendChild(td(e.cnpj_fmt));
        tr.appendChild(td(e.uf));
        tr.appendChild(td(e.mes_ref));
        tr.appendChild(td(brl(e.faturamento), "cscie-num"));
        tr.appendChild(td(brl(e.exportacao), "cscie-num"));
        tr.appendChild(td(pctFmt(e.percentual),
          "cscie-num" + (e.percentual === null ? " cscie-vazio" : "")));
        tr.appendChild(td(brl(e.saldo_credor), "cscie-num"));

        var tdStatus = document.createElement("td");
        var tag = document.createElement("span");
        tag.className = CLASSE_TAG[e.status] || "cscie-tag";
        tag.textContent = e.rotulo_status;
        var ver = document.createElement("button");
        ver.type = "button";
        ver.className = "cscie-ver-motivo";
        ver.textContent = "ver motivo";
        tdStatus.appendChild(tag);
        tdStatus.appendChild(ver);
        tr.appendChild(tdStatus);

        tr.appendChild(td(brl(e.correcao),
          "cscie-num" + (e.correcao === null ? " cscie-vazio" : "")));
        elTabelaCorpo.appendChild(tr);

        // linha expansível com a explicação COMPLETA — a mesma string do Excel
        var trMotivo = document.createElement("tr");
        trMotivo.className = "cscie-motivo";
        trMotivo.hidden = true;
        var tdMotivo = document.createElement("td");
        tdMotivo.colSpan = 9;
        // a célula tem a largura da TABELA, que é maior que o container quando há
        // rolagem horizontal — sem este invólucro grudado à esquerda, o fim da
        // explicação ficaria fora da área visível
        var caixaMotivo = document.createElement("div");
        caixaMotivo.textContent = e.explicacao;
        tdMotivo.appendChild(caixaMotivo);
        trMotivo.appendChild(tdMotivo);
        elTabelaCorpo.appendChild(trMotivo);

        ver.addEventListener("click", function () {
          trMotivo.hidden = !trMotivo.hidden;
          ver.textContent = trMotivo.hidden ? "ver motivo" : "ocultar motivo";
        });
      });

      var tr = document.createElement("tr");
      tr.appendChild(td("TOTAL"));
      for (var i = 0; i < 6; i++) tr.appendChild(td(""));
      tr.appendChild(td(res.totais.n_calculados + " de " + res.totais.n_estabelecimentos +
        " qualificaram"));
      tr.appendChild(td(brl(res.totais.correcao), "cscie-num"));
      elTabelaRodape.appendChild(tr);
    }

    function montarSeletor(res) {
      elSeletor.innerHTML = "";
      res.estabelecimentos.forEach(function (e, i) {
        var opt = document.createElement("option");
        opt.value = String(i);
        opt.textContent = e.cnpj_fmt + " — " + e.uf;
        elSeletor.appendChild(opt);
      });
      elSeletor.onchange = function () { montarSerie(Number(elSeletor.value)); };
      montarSerie(0);
    }

    // O backend devolve a série de todos: trocar de estabelecimento é só trocar a
    // visão, sem novo upload.
    function montarSerie(indice) {
      var e = (estado.resultado.estabelecimentos || [])[indice];
      elSerieCorpo.innerHTML = "";
      if (!e) return;
      e.serie.forEach(function (r) {
        var ehRef = r.mes === e.mes_ref;
        var tr = document.createElement("tr");
        if (ehRef) tr.className = "cscie-linha-ref";
        tr.appendChild(td((ehRef ? "◄ " : "") + r.mes));
        tr.appendChild(td(brl(r.faturamento), "cscie-num"));
        tr.appendChild(td(brl(r.exportacao), "cscie-num"));
        tr.appendChild(td(r.faturamento ? pctFmt(r.exportacao / r.faturamento) : "—",
          "cscie-num"));
        // célula VAZIA (não zero) quando o mês não tem escrituração
        tr.appendChild(td(r.saldo_credor === null ? "" : brl(r.saldo_credor),
          "cscie-num" + (r.saldo_credor === null ? " cscie-vazio" : "")));
        elSerieCorpo.appendChild(tr);
      });
    }

    // ------------------------------------------------- copiar e baixar
    elBtnCopiar.addEventListener("click", function () {
      var texto = elVereditoTexto.textContent || "";
      var avisar = function (msg) {
        elBtnCopiar.textContent = msg;
        window.setTimeout(function () {
          elBtnCopiar.textContent = "Copiar explicação";
        }, 2500);
      };
      if (navigator.clipboard && navigator.clipboard.writeText) {
        navigator.clipboard.writeText(texto)
          .then(function () { avisar("Copiado!"); })
          .catch(function () { avisar("Não consegui copiar"); });
      } else {
        // navegador sem Clipboard API (ou página sem contexto seguro). O campo
        // auxiliar fica fora de fluxo: dentro do fluxo ele empurraria a tela.
        var ta = document.createElement("textarea");
        ta.value = texto;
        ta.setAttribute("aria-hidden", "true");
        ta.style.cssText = "position:fixed;top:0;left:-9999px;opacity:0;height:1px";
        root.appendChild(ta);
        ta.select();
        try { avisar(document.execCommand("copy") ? "Copiado!" : "Não consegui copiar"); }
        catch (e) { avisar("Não consegui copiar"); }
        root.removeChild(ta);
      }
    });

    // Baixa via fetch para levar o cabeçalho de autenticação e para que um erro
    // (link expirado, já usado) apareça como mensagem, não como página crua.
    elDownload.addEventListener("click", function (ev) {
      ev.preventDefault();
      if (!estado.token) return;
      var alvo = elDownload;
      alvo.textContent = "Gerando o download…";
      cscieFetch(BASE + "/download/" + estado.token)
        .then(function (resp) {
          var ct = resp.headers.get("content-type") || "";
          if (!resp.ok || ct.indexOf("json") !== -1) return cscieLerJson(resp)
            .then(function (d) { throw new Error(d.erro || "Falha no download."); });
          return resp.blob();
        })
        .then(function (blob) {
          var url = URL.createObjectURL(blob);
          var a = document.createElement("a");
          a.href = url;
          a.download = "Correcao_Saldo_Credor_ICMS_Exportacao.xlsx";
          root.appendChild(a);
          a.click();
          root.removeChild(a);
          window.setTimeout(function () { URL.revokeObjectURL(url); }, 4000);
          // o link é de uso único por segurança: dizer isso é melhor que deixar o
          // segundo clique falhar sem explicação
          estado.token = null;
          alvo.textContent = "Planilha baixada";
          alvo.classList.add("cscie-btn-secondary");
        })
        .catch(function (e) {
          alvo.textContent = "Baixar planilha auditável";
          falhar(e && e.message ? e.message : "Falha no download.");
        });
    });

    // ---------------------------------------------------------------- estado
    function limparSaida() {
      // interrompe o acompanhamento anterior: sem isto, dois cálculos sobrepostos
      // disputam a mesma tela e o antigo sobrescreve o novo
      if (estado.timer) { window.clearTimeout(estado.timer); estado.timer = null; }
      estado.token = null;
      elResultado.hidden = true;
      elErro.hidden = true;
      elProgresso.hidden = true;
      elBarFill.style.width = "0%";
      elProgMsg.textContent = "";
      elDecorrido.textContent = "0 s";
      elDownload.textContent = "Baixar planilha auditável";
      elDownload.classList.remove("cscie-btn-secondary");
    }

    function falhar(msg) {
      if (estado.timer) { window.clearTimeout(estado.timer); estado.timer = null; }
      elProgresso.hidden = true;
      elErro.hidden = false;
      elErroMsg.textContent = msg;
      elBtn.disabled = false;
      atualizarBotao();
    }

    // Volta a ferramenta ao estado inicial — o analista passa de uma empresa para
    // a outra sem recarregar a página, e nada da análise anterior fica na tela.
    function limparTudo() {
      limparSaida();
      consultaEmCurso += 1;          // invalida consulta de CNPJ em voo
      estado.cnpj = "";
      estado.razao = "";
      estado.manual = false;
      estado.resultado = null;
      elCnpj.value = "";
      elArquivos.value = "";
      elLista.textContent = "";
      mostrarRazao("—", true);
      avisarCnpj("");
      elKpis.innerHTML = "";
      elKpis.hidden = true;
      elAvisos.innerHTML = "";
      elAvisos.hidden = true;
      elTabelaCorpo.innerHTML = "";
      elTabelaRodape.innerHTML = "";
      elSerieCorpo.innerHTML = "";
      elSeletor.innerHTML = "";
      elVereditoTexto.textContent = "";
      elVereditoValor.textContent = "";
      atualizarBotao();
      elCnpj.focus();
    }

    elBtnLimpar.addEventListener("click", limparTudo);

    mostrarRazao("—", true);
    atualizarBotao();
  }

  window.initCorrecaoSaldoCredorIcms = initCorrecaoSaldoCredorIcms;
  if (document.readyState !== "loading") {
    if (document.getElementById("tool-correcao-saldo-credor-icms")) {
      initCorrecaoSaldoCredorIcms();
    }
  } else {
    document.addEventListener("DOMContentLoaded", function () {
      if (document.getElementById("tool-correcao-saldo-credor-icms")) {
        initCorrecaoSaldoCredorIcms();
      }
    });
  }
})();
