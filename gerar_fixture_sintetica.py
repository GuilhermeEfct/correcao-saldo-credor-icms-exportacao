#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Gerador da fixture sintetica para a ferramenta `correcao-saldo-credor-icms-exportacao`.

Produz 4 arquivos CSV no formato do relatorio `ICMSProprio` (Apuracao de ICMS),
um por estabelecimento, cobrindo os quatro cenarios do teste de aceitacao.

NAO CONTEM DADO REAL DE CLIENTE. Pode ser commitado no repositorio.

Uso:
    python gerar_fixture_sintetica.py [pasta_destino]

Cenarios gerados
----------------
A) 12345678000190 - RS  -> CALCULA. Ultimo mes DEZ/2025.
       faturamento de venda ......... R$   900.000,00
       faturamento de exportacao .... R$   100.000,00  (CFOP 7101)
       faturamento total ............ R$ 1.000.000,00
       transferencias (fora do calc.) R$ 5.000.000,00  (CFOP 5151)
       % exportacao ................. 10,00%
       saldo credor ................. R$   500.000,00
       CORRECAO ..................... R$    50.000,00   <-- resultado esperado

B) 12345678000271 - SP  -> SEM EXPORTACAO no ultimo mes (tem saldo credor).
       Exportou em AGO/2025, nao em DEZ/2025. Saldo credor DEZ/2025 = R$ 250.000,00.

C) 12345678000352 - RS  -> SEM SALDO CREDOR no ultimo mes.
       Saldo zera a partir de OUT/2025. Testa a mensagem terminal e a frase de contexto.

D) 12345678000433 - ES  -> Ultimo mes ANTERIOR aos demais (NOV/2025).
       Testa que o mes de referencia e por estabelecimento, nao do lote.
       Tem saldo credor e exportacao em NOV/2025 -> CALCULA R$ 12.000,00.

Total esperado da correcao: R$ 62.000,00 (A + D).
"""
import os
import sys

MESES = ["JUL/2025", "AGO/2025", "SET/2025", "OUT/2025", "NOV/2025", "DEZ/2025"]


def ptbr(v):
    """1234567.89 -> '1.234.567,89'"""
    return f"{v:,.2f}".replace(",", "\x00").replace(".", ",").replace("\x00", ".")


def linha(rotulo, valores):
    return rotulo + ";" + ";".join(ptbr(v) for v in valores)


# ---------------------------------------------------------------------------
# Definicao dos estabelecimentos
#   cfops: {"NNNN - Descricao": [valor por mes...]}   (Secao 19 - saidas)
#   saldo: [valor por mes...]                          (Secao 22 - E110)
#   meses: rotulos cobertos por este estabelecimento
# ---------------------------------------------------------------------------
ESTABELECIMENTOS = [
    {
        "cnpj": "12345678000190",
        "uf": "RS",
        "arquivo": "FIXTURE_ICMSProprio_A_calcula.csv",
        "meses": MESES,
        "cfops": {
            # JUL      AGO      SET      OUT      NOV      DEZ
            "5101 - Venda prod do estab":
                [700000.0, 750000.0, 800000.0, 850000.0, 880000.0, 900000.0],
            "7101 - Venda prod do estab":
                [0.0, 40000.0, 0.0, 0.0, 0.0, 100000.0],
            # transferencia: NAO entra no denominador do percentual
            "5151 - Transf prod":
                [4000000.0, 4200000.0, 4500000.0, 4700000.0, 4800000.0, 5000000.0],
        },
        "saldo": [300000.0, 350000.0, 400000.0, 430000.0, 470000.0, 500000.0],
    },
    {
        "cnpj": "12345678000271",
        "uf": "SP",
        "arquivo": "FIXTURE_ICMSProprio_B_sem_exportacao.csv",
        "meses": MESES,
        "cfops": {
            "6102 - Venda mercad adq/receb terc":
                [500000.0, 500000.0, 500000.0, 500000.0, 500000.0, 500000.0],
            # exportou em AGO/2025, mas NAO no ultimo mes -> trava
            "7101 - Venda prod do estab":
                [0.0, 80000.0, 0.0, 0.0, 0.0, 0.0],
        },
        "saldo": [150000.0, 180000.0, 200000.0, 220000.0, 240000.0, 250000.0],
    },
    {
        "cnpj": "12345678000352",
        "uf": "RS",
        "arquivo": "FIXTURE_ICMSProprio_C_sem_saldo.csv",
        "meses": MESES,
        "cfops": {
            "5102 - Venda mercad adq/receb terc":
                [300000.0, 300000.0, 300000.0, 300000.0, 300000.0, 300000.0],
            "7101 - Venda prod do estab":
                [0.0, 0.0, 20000.0, 0.0, 0.0, 30000.0],
        },
        # saldo zera a partir de OUT/2025 -> ultimo mes com saldo = SET/2025
        "saldo": [90000.0, 60000.0, 45000.0, 0.0, 0.0, 0.0],
    },
    {
        "cnpj": "12345678000433",
        "uf": "ES",
        "arquivo": "FIXTURE_ICMSProprio_D_mes_anterior.csv",
        # termina em NOV/2025: um mes antes dos demais
        "meses": MESES[:-1],
        "cfops": {
            "5101 - Venda prod do estab":
                [180000.0, 190000.0, 195000.0, 198000.0, 176000.0],
            "7101 - Venda prod do estab":
                [0.0, 0.0, 0.0, 0.0, 24000.0],
        },
        "saldo": [80000.0, 90000.0, 95000.0, 98000.0, 100000.0],
    },
]


CAB_S1 = [
    "ICMSProprio - 1. Resumo ICMS - [ E110 - VL_TOT_DEBITOS ] - Valor total dos débitos",
    "[ E110 - VL_TOT_CREDITOS ] - Valor total dos créditos",
    "[ E110 -  VL_SLD_CREDOR_ANT ] - Valor total de \"Saldo credor do período anterior\"",
    "[ E110 -  VL_SLD_CREDOR_TRANSPORTAR ] - Valor total de \"Saldo credor a transportar "
    "para o período seguinte\"",
]


def montar(est):
    m = est["meses"]
    n = len(m)
    cab = "DESCRIÇÃO;" + ";".join(m)
    saidas = [sum(v[i] for v in est["cfops"].values()) for i in range(n)]
    saldo_ant = [0.0] + est["saldo"][:-1]

    out = []
    # ---- Secao 1 (Resumo) -------------------------------------------------
    out += CAB_S1
    out.append(cab)
    # esta linha DEVE fechar com a soma da Secao 19 (check de integridade 3.6b)
    out.append(linha("Valor Operacional - Saídas/Prestações", saidas))
    out.append(linha("( = ) Total dos débitos   ➤  Débito", [v * 0.02 for v in saidas]))
    out.append(linha("( = ) Total dos Créditos   ➤  Crédito", [v * 0.05 for v in saidas]))
    out.append(linha("  ➦  ( + ) Saldo Credor do Período Anterior", saldo_ant))
    out.append(linha("            💲  Total de ICMS a Recolher", [0.0] * n))
    out.append(linha("  ➥  Saldo Credor a Transportar - Período Seguinte", est["saldo"]))
    out += ["", ""]

    # ---- Secao 19 (saidas por CFOP) --------------------------------------
    out.append("ICMSProprio - 19. Valor Operacional por CFOP - Saídas/Prestações - "
               "[ C190 - VL_OPR ] - Valor da operação")
    out.append(cab)
    for rot in sorted(est["cfops"]):
        out.append(linha(rot, est["cfops"][rot]))
    out += ["", ""]

    # ---- Secao 20 (entradas por CFOP) — presente, sem devolucao de export --
    out.append("ICMSProprio - 20. Valor Operacional por CFOP - Entradas/Aquisições - "
               "[ C190 - VL_OPR ] - Valor da operação")
    out.append(cab)
    out.append(linha("1101 - Compra p/indust ou prod rural", [v * 0.6 for v in saidas]))
    out += ["", ""]

    # ---- Secao 21 (saldo credor anterior por CNPJ) ------------------------
    out.append("ICMSProprio - 21. Abertura Saldo Credor por CNPJ - "
               "[ E110 VL_SLD_CREDOR_ANT] - Valor do saldo credor período anterior")
    out.append(cab)
    out.append(linha(f"{est['cnpj']} - {est['uf']}", saldo_ant))
    out += ["", ""]

    # ---- Secao 22 (saldo a transportar por CNPJ) -> FONTE DO SALDO CREDOR --
    out.append("ICMSProprio - 22. Abertura Saldo a Transportar por CNPJ - "
               "[ E110 VL_SLD_CREDOR_TRANSPORTAR] - Valor do saldo credor à transportar "
               "para o período seguinte")
    out.append(cab)
    out.append(linha(f"{est['cnpj']} - {est['uf']}", est["saldo"]))
    out += ["", ""]

    # ---- Secao 23 (ICMS a recolher por CNPJ) ------------------------------
    out.append("ICMSProprio - 23. Abertura ICMS a Recolher por CNPJ - "
               "[ E110 VL_ICMS_RECOLHER] - Valor do ICMS a recolher")
    out.append(cab)
    out.append(linha(f"{est['cnpj']} - {est['uf']}", [0.0] * n))
    out.append("")
    return "\r\n".join(out)


def main():
    destino = sys.argv[1] if len(sys.argv) > 1 else "fixtures"
    os.makedirs(destino, exist_ok=True)
    for est in ESTABELECIMENTOS:
        caminho = os.path.join(destino, est["arquivo"])
        # utf-8-sig: mesmo encoding do relatorio real exportado pelo sistema fiscal
        with open(caminho, "w", encoding="utf-8-sig", newline="") as fh:
            fh.write(montar(est))
        print(f"gerado: {caminho}")

    print("\nResultado esperado ao processar os 4 arquivos juntos:")
    print("  A 12.345.678/0001-90 RS  DEZ/2025  Calculado ................. R$ 50.000,00")
    print("  B 12.345.678/0002-71 SP  DEZ/2025  Sem exportacao no ultimo mes ....... —")
    print("  C 12.345.678/0003-52 RS  DEZ/2025  Sem saldo credor no ultimo mes ..... —")
    print("  D 12.345.678/0004-33 ES  NOV/2025  Calculado ................. R$ 12.000,00")
    print("  TOTAL ...................................................... R$ 62.000,00")
    print("\nO estabelecimento C deve trazer a frase de contexto:")
    print('  "O ultimo mes em que houve saldo credor foi SET/2025, com R$ 45.000,00."')
    print("\nO estabelecimento D deve citar NOV/2025, nao DEZ/2025.")


if __name__ == "__main__":
    main()
