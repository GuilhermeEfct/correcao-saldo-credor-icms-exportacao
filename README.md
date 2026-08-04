# Correção do Saldo Credor de ICMS proporcional às Exportações

Sub-página do **HUB EFCT**, grupo **ICMS**. Slug `correcao-saldo-credor-icms-exportacao`.

Sobre o **último mês de apuração** de cada estabelecimento, calcula o percentual do
faturamento de exportação em relação ao faturamento total e aplica esse percentual ao
**saldo credor de ICMS acumulado** (registro E110, `VL_SLD_CREDOR_TRANSPORTAR`). Quando
não há oportunidade, **explica o motivo em prosa** — e isso é resultado, não erro.

---

## Janela de 60 meses

A análise abrange os **últimos 60 meses** contados do **mês de corte da extração** (o mês
mais recente do lote) — o prazo dentro do qual a recuperação de crédito não está
prescrita. Mês fora da janela **não aparece em lugar nenhum**: nem na série de
conferência, nem na aba de exportações, nem no período informado no topo do relatório.
Não é "marcar em amarelo", é omitir — dado prescrito exibido ao lado do apurável convida a
somar os dois.

A contagem é de **calendário**, não de colunas presentes: se a série tiver buraco, pegar
"as últimas 60 colunas" alcançaria mês já prescrito.

Consequências que valem conhecer:

- Estabelecimento **sem nenhuma escrituração** na janela sai da análise — mas com **aviso
  na tela**, nunca em silêncio. Se nenhum sobrar, é erro de negócio com mensagem humana.
- A âncora é o mês de corte, **não a data de hoje**, para o mesmo arquivo produzir o mesmo
  relatório daqui a um ano — exigência de documento auditável. O efeito colateral é que
  uma extração antiga passaria como se nada estivesse prescrito, então a ferramenta
  **avisa** quando o mês de corte tem mais de 12 meses. O aviso não altera número algum.
- No caso real de referência, os arquivos trazem 65 meses (JAN/2021 a MAI/2026) e a
  análise usa 60 (JUN/2021 a MAI/2026). A remessa de exportação de ABR/2021 (R$ 3.241,00,
  CFOP 5501) fica fora, e o total exportado na janela é R$ 1.266.558,77 em vez dos
  R$ 1.269.799,77 dos 65 meses.

## O que ela faz, em três desfechos

O cálculo é feito exclusivamente sobre o último mês de apuração de cada estabelecimento,
**dentro da janela de 60 meses**. Não existe mês alternativo, retrocesso na série, média do
período nem soma de meses.

Nesse mês, duas condições são avaliadas em cascata **curto-circuitada**:

| # | Condição no último mês de apuração | Desfecho | Devolve valor? |
|---|---|---|---|
| 1 | Saldo credor de ICMS = 0 | `SEM_SALDO_CREDOR` — **termina aqui**, nem verifica exportação | não |
| 2 | Saldo credor > 0, exportação = 0 | `SEM_EXPORTACAO` | não |
| 3 | Saldo credor > 0 **e** exportação > 0 | `CALCULADO`: `saldo × (exportação ÷ faturamento)` | sim |

**Por que o último mês, e só ele:** o saldo credor é um *estoque* — o que interessa é o
que está disponível agora. E o percentual precisa ser medido no mesmo mês do estoque,
senão se cruzam grandezas de períodos diferentes. Escolher outro mês da série abriria
margem para seleção conveniente: num caso real de 65 meses, o resultado varia mais de
100× conforme o mês escolhido.

**Por que o teste do saldo credor é terminal:** sem saldo credor não existe base sobre a
qual aplicar proporção nenhuma, então é irrelevante se houve exportação. Por isso a
explicação do desfecho 1 **não apresenta a exportação como obstáculo** — mencioná-la
sugeriria, falsamente, que ela era o impedimento.

**`None` não é zero.** Nos desfechos 1 e 2, correção e percentual voltam como `None` e
aparecem na tela e na planilha como `—` / célula em branco. Um `R$ 0,00` numa coluna de
valor seria lido como "calculamos e deu zero", que é uma afirmação diferente e errada.

---

## 🔴 "Sem oportunidade" percorre o caminho de SUCESSO

Este é o requisito que o código mais tende a violar, porque o caminho natural leva ao
erro. Em todas as camadas:

| Camada | Comportamento |
|---|---|
| `logica.py` | retorna o dicionário normalmente (nunca levanta exceção) |
| job assíncrono | `status = "concluido"` |
| rota HTTP | **200** com o resultado completo |
| tela | renderiza a explicação em destaque, sem tarja vermelha nem `alert` |
| botão Calcular | segue habilitado |
| download do Excel | **disponível** — o relatório que documenta a não-aplicabilidade tem valor para o cliente |

`ErroDeNegocio` está reservada a falha real: arquivo ilegível, nenhuma apuração
reconhecida, CNPJ raiz divergente, nenhum mês identificado. Um analista que recebe
"não foi possível gerar" não sabe se o arquivo estava errado, se o sistema falhou ou se
o cliente realmente não se aplica — e volta a conferir à mão, que é justamente o
trabalho que a ferramenta deveria eliminar.

**E o mesmo vale para o caminho de erro: ele tem de se explicar.** A primeira versão da
mensagem de "não identifiquei os períodos" dizia apenas *"confira se é o relatório de
Apuração de ICMS"* — jogava a suspeita no arquivo do analista. E errou: o arquivo estava
correto, e era a ferramenta que não sabia ler rótulo de mês com ano de 2 dígitos. Quem
recebeu aquela mensagem ficou sem saber se a ferramenta funciona.

Agora cada erro de leitura diz **o que encontrou** e **o que esperava**, e distingue três
situações:

| Situação | O que a mensagem informa |
|---|---|
| Formato de mês desconhecido | as seções reconhecidas, as primeiras colunas do cabeçalho, os formatos aceitos, e que **isto parece variação de formato, não problema no arquivo** — com o pedido de encaminhar ao admin |
| Arquivo sem o cabeçalho de meses | as seções reconhecidas e a suspeita de truncamento |
| Arquivo que não é apuração | que não achou nenhuma seção `ICMSProprio - N.` |
| Nenhum arquivo reconhecido no lote | **os nomes** dos arquivos recusados e o que a ferramenta procura |

---

## Entrada

Relatório de **Apuração de ICMS** (`ICMSProprio`) exportado em CSV, **um arquivo por
estabelecimento**. Aceita `.csv`, `.txt` e `.zip` (classificados pelo conteúdo, não pelo
nome). Separador `;`, encoding `utf-8-sig` com fallback `latin-1`, números pt-BR.

Seções lidas, numa **única passada linha a linha**:

| Seção | Para que serve |
|---|---|
| **19** — Valor Operacional por CFOP - Saídas | numerador (exportação) e denominador (faturamento), por mês |
| **22** — Abertura Saldo a Transportar por CNPJ | saldo credor por mês (E110), CNPJ e UF do estabelecimento |
| **21** — Abertura Saldo Credor por CNPJ | fallback do CNPJ |
| **20** — Valor Operacional por CFOP - Entradas | apenas CFOPs de devolução/retorno de exportação |
| **1** — Resumo ICMS | check de integridade contra a Seção 19 |

### Três armadilhas do formato, todas reais

1. **O número de meses varia por estabelecimento** — filial aberta depois tem menos
   colunas. No caso de referência os arquivos tinham 65, 65, 65, 52, 45, 33 e 21 meses.
   Nunca assuma número fixo de colunas.
2. **O alinhamento é pelo rótulo do mês** (`MAI/2026`), nunca pela posição da coluna.
3. **Os rótulos são pt-BR** (`JAN`…`DEZ`) e a ordenação é por `(ano, mês)` — ordem
   alfabética destruiria a série.

E uma consequência importante: **o mês de referência é a última coluna do arquivo
daquele estabelecimento**, não o último mês do lote e não o último mês com movimento.
No caso de referência, um estabelecimento termina em ABR/2026 enquanto os outros seis
vão até MAI/2026, e outro tem faturamento R$ 0,00 no seu último mês — que ainda assim é
o mês dele.

---

## Identificação da empresa — "CNPJ primeiro"

A tela começa pela **Empresa** (passo 1), antes de qualquer upload: o analista informa o
CNPJ com máscara, a razão social é consultada em `GET /api/cnpj/<cnpj>` do HUB, e o botão
principal só habilita depois da consulta dar certo. É o padrão de produto do §4.4 do
team-kit, e vale nas duas pontas: **`POST /iniciar` recusa com 400 sem CNPJ de 14
dígitos**. Sem essa recusa a regra valeria só na tela — uma chamada direta à API
produziria planilha com empresa adivinhada a partir dos arquivos.

**Por que declarar é mais forte que adivinhar.** O CNPJ está dentro dos relatórios (Seções
21/22/23 trazem `<14 dígitos> - <UF>`), e é tentador extraí-lo em vez de pedir. Mas aí quem
diz qual é a empresa é o próprio arquivo, e um **lote trocado produz uma planilha coerente
consigo mesma e errada em relação ao caso** — num documento que pode virar anexo
processual. Com o CNPJ declarado, o analista afirma qual empresa está analisando e a
ferramenta confere os arquivos contra essa afirmação.

`POST /lote/{protocolo}/identificar` continua existindo, agora como **conferência**: conta
os estabelecimentos encontrados, lista os arquivos que não são apuração e **barra o lote
quando os relatórios são de outra empresa**. Junto com o `validar_cnpj_raiz` do
processamento, são segunda e terceira linha — não a única linha.

## Decisões técnicas não-óbvias

- **Leitura em fluxo**, reaproveitada da ferramenta-irmã `credito-icms-uso-consumo-exportacao`
  (`listar_partes`, `abrir_linhas`, `_encoding_do_fluxo`): cada arquivo é lido linha a
  linha, inclusive membros de `.zip`, e o pico de memória não acompanha o tamanho da
  entrada.
- **`_num()` devolve `None` para célula vazia**, e não `0.0`. A diferença importa: mês sem
  escrituração é ausência de dado, não saldo zero — e aparece em branco na planilha.
- **O percentual não é arredondado no cálculo** (só formatado com 2 casas). A irmã
  arredonda a proporção; aqui isso faria o número da tela divergir da fórmula viva do
  Excel ao recalcular.
- **Onde esta ferramenta divergiu da irmã, de propósito:** a irmã usa a proporção do
  período inteiro, com todos os CFOPs de saída no denominador e só o grupo 7 no
  numerador. Aqui a proporção é **de um único mês**, o numerador inclui a **exportação
  indireta** e o denominador tem **apenas CFOPs de venda + exportação**. A irmã também
  usa `setInterval` no polling; aqui é `setTimeout` recursivo, como manda o padrão do
  team-kit (com `setInterval`, duas consultas se sobrepõem quando a resposta demora e a
  antiga sobrescreve a nova).
- **Excel em modo normal, não `write_only`.** O bloco de veredito exige célula mesclada
  com `wrap_text`, e `WriteOnlyWorksheet` não tem `merge_cells`. O volume é limitado
  (um estabelecimento rende ~65 linhas na aba de conferência), então não há o problema de
  memória que obrigou a irmã a usar `write_only`.
- **As células de insumo da aba `Cálculo` referenciam a aba de conferência**
  (`='Série Mensal (conferência)'!C71`), em vez de repetir o número. O revisor clica na
  célula e chega ao mês de referência; e existe uma única fonte de verdade.
- **Arquivo duplicado é descartado, não somado.** Se o mesmo CNPJ vem em dois arquivos, o
  primeiro vale e o segundo é descartado com aviso na tela. Somar dobraria faturamento e
  saldo; fundir séries parciais seria adivinhação.
- **Grupo 7 não é sinônimo de exportação.** A regra original tratava todo CFOP `7xxx`
  como exportação direta — mas o grupo é "saída para o exterior", que é mais amplo:
  inclui devolução de compra (`7201`/`7202`), anulação de valor (`7205`-`7207`) e saída
  não especificada (`79xx`). Nada disso é faturamento de exportação nem gera o crédito.
  Num relatório real, um estabelecimento cujas únicas saídas do mês eram `7202` e `7949`
  produzia **100% de exportação** e devolvia o **saldo credor inteiro** como oportunidade
  — erro para cima, a pior direção. Por isso a lista é explícita
  (`7101`, `7102`, `7105`, `7106`, `7127`, `7501`), como a de venda já era.
- **Quando a saída ao exterior não qualifica, a explicação diz isso.** Sem essa frase o
  analista lê "não houve exportação", abre o relatório, vê CFOP do grupo 7 no mês e vai
  conferir à mão para entender a contradição. A explicação nomeia os CFOPs encontrados e
  o motivo de não entrarem na proporção.
- **Devolução de exportação tem piso em zero.** CFOPs `1503-1506`, `2503-2506`, `3201`,
  `3202` e `3211` das entradas são abatidos da exportação do mês em que ocorreram. Sem o
  piso, um retorno maior que a exportação do mês produziria exportação negativa e
  percentual negativo na aba de conferência. Quando o piso é acionado, a aba `Exportações`
  registra em qual mês e estabelecimento.
- **A conferência dos relatórios é uma passada leve e separada** (`identificar_estabelecimentos`),
  que para no primeiro CNPJ de cada arquivo em vez de montar a série inteira. Roda entre o
  upload e o cálculo, sobre os arquivos **já enviados** — o analista não sobe nada duas vezes.
- **O rótulo do mês vem em dois formatos** e os dois são aceitos: `JAN/2021` e `jan/21`
  (ano de 2 dígitos vira 20XX). Com só o formato de 4 dígitos, um relatório exportado no
  outro formato era recusado como se estivesse errado. Colunas de preenchimento (`;;;;`) no
  fim da linha do cabeçalho são descartadas — o alinhamento dos valores é posicional, então
  a lista de colunas guarda a posição e a lista de meses guarda só o que é mês.
- **A permissão é a do GRUPO (`icms`), não uma chave própria da ferramenta.** Quem tem
  acesso ao card do grupo ICMS tem acesso à ferramenta, como o C170/C175 fazem com
  `pis_cofins`.
- **O gate de permissão nega por padrão.** A forma intuitiva
  (`if isinstance(p, dict) and not p.get(chave)`) tem *fail-open*: contexto sem a chave
  `permissions`, ou com tipo inesperado, pula o `if` inteiro e deixa entrar qualquer usuário
  autenticado. A checagem certa monta o dict vazio quando o tipo não serve e exige
  `is_admin or permissions.get(chave)`. Sem provider registrado, o módulo nega — não existe
  caminho próprio de liberação, nem por `TESTING`: a liberação de desenvolvimento mora no
  `test_app.py`, que declara não ir para produção.
- **O teto por requisição é aplicado na rota**, com checagem dupla (`request.content_length`
  antes de gravar e o tamanho real depois, para header ausente ou mentiroso). Não via
  `MAX_CONTENT_LENGTH`: o HUB não define esse valor globalmente, e não pode — todas as
  ferramentas dividem o mesmo app e um teto global cortaria o upload de todas.
- **Processamento assíncrono** desde o nascimento: o Render corta requisição em ~240 s, e
  aumentar o `--timeout` do gunicorn não resolve. `POST /lote` → N × `POST /lote/{p}/arquivo`
  → `POST /lote/{p}/identificar` → `POST /lote/{p}/iniciar` (202) → `GET /status/{p}` a cada
  3 s → `GET /download/{token}`. Os arquivos aqui são pequenos, mas o cliente pode subir um
  `.zip` grande ou muitos estabelecimentos.

---

## Avisos importantes

- **O valor apurado é estimativa da parcela proporcional às exportações**, não crédito
  habilitado à transferência. Ver as ressalvas metodológicas abaixo.
- A ferramenta **não lê C100/C170/C190** e não faz consulta a API externa. A razão social
  vem da rota `GET /api/cnpj/<cnpj>` **do próprio HUB**.
- **Divergência de CNPJ raiz aborta a análise.** A planilha sai com a razão social vinda
  da tela e o saldo credor vindo dos arquivos: se forem de empresas diferentes, o
  documento — que pode virar anexo processual — fica silenciosamente errado. É o pior
  tipo de falha: não quebra, só mente. Estabelecimentos diferentes da **mesma** raiz são
  o caso normal e passam juntos.
- O **link de download é de uso único** e expira em 10 minutos, por exigência de
  segurança. A tela avisa depois de baixar; para baixar de novo, é recalcular.
- Se a Seção 19 divergir da Seção 1 em algum mês, a ferramenta **avisa e segue** — o
  arquivo pode estar truncado, e é melhor o analista saber antes de usar o número.

---

## Ressalvas metodológicas

As notas abaixo vão no rodapé da aba `Cálculo` do Excel (as duas condicionais
entram só quando se aplicam). A redação canônica está em `logica.py`, para ser revisada
num só lugar.

1. **Critério do mês de referência.** O cálculo é feito exclusivamente sobre o último mês
   de apuração de cada estabelecimento, sem retrocesso na série histórica. O mês
   utilizado consta na coluna "Último mês de apuração".

2. **Período analisado.** A análise abrange os últimos 60 meses contados do mês de corte
   da extração, prazo dentro do qual a recuperação de crédito não está prescrita. Meses
   anteriores foram omitidos de todo o relatório — inclusive da série de conferência e das
   operações de exportação —, e não integram nenhum total apresentado.

3. **Natureza do saldo utilizado.** O valor utilizado é o saldo credor a transportar
   apurado no registro E110 (`VL_SLD_CREDOR_TRANSPORTAR`), que corresponde ao saldo da
   apuração ordinária. **Não equivale, por si, a crédito acumulado formalmente gerado e
   apropriado** nos termos do art. 25, §1º, I, da LC 87/96 — cuja constituição depende do
   rito da legislação de cada estado (por exemplo, a sistemática de crédito acumulado do
   RICMS/SP ou regime especial no RS). O valor aqui apurado é a **estimativa da parcela
   proporcional às exportações**, e não um crédito já habilitado à transferência.
   *Quem mantiver esta ferramenta precisa saber disso: é a distinção que separa o número
   entregue de uma promessa que a EFCT não pode fazer.*

4. **Composição do faturamento total.** Considera exclusivamente CFOPs de venda somados
   aos de exportação. Transferências entre estabelecimentos da própria empresa
   (5151/5152/5153/6151/6152/6153), remessas, bonificações e devoluções **não** integram o
   denominador — transferência não é faturamento, e incluí-la derrubaria o percentual sem
   justificativa técnica. É a pergunta nº 1 que o cliente faz.

5. **Exportação indireta** *(condicional)*. As operações classificadas nos CFOPs
   5501/5502/6501/6502 são remessas com fim específico de exportação e foram computadas
   pelo valor de face. O direito ao crédito depende da efetiva exportação no prazo legal
   pelo destinatário, o que não é verificável a partir da apuração de ICMS.

6. **Devolução de exportação** *(condicional)*. Operações de devolução/retorno de
   exportação identificadas nas entradas são abatidas do faturamento de exportação do mês
   em que ocorreram. Quando o valor devolvido supera a exportação do próprio mês, o
   resultado é limitado a zero — nunca a um valor negativo.

7. **Origem dos dados.** Seções 19 e 22 do relatório de apuração de ICMS (`ICMSProprio`),
   correspondentes aos registros C190 e E110 da EFD ICMS/IPI, com o período coberto e o
   mês de corte da extração declarados na própria nota — para que uma execução futura com
   SPED mais recente possa ser comparada com esta.

**Base legal:** art. 25, §1º, I, da LC 87/96 — o crédito acumulado decorrente de
exportação é de livre disposição (transferível a terceiros), diferente do acumulado por
outras causas. Daí a necessidade de demonstrar **origem** e **proporção**.

---

## Saída — planilha auditável

`.xlsx` com **tipografia e paleta oficiais do Grupo EFCT** (seções 4.1 e 4.2 do
team-kit) e **fórmulas reais** em toda célula calculada.

- **Bebas Neue** nos títulos das abas e no valor do veredito.
- **Exo 2** no corpo, nos labels e nos cabeçalhos de coluna — cobre quase toda a planilha.
- **Libre Baskerville Italic** na ressalva de natureza do saldo, que é a de peso jurídico.
- Cores **somente** da paleta oficial: petróleo `#001C26`, oliva `#B3BC2B`, ciano
  `#4DBDF5`, creme `#FEFEDF`, texto `#111827`, texto secundário `#6B7280`, borda sutil
  `#E5E7EB`. Nenhuma cor descontinuada (`#3B82F6`, `#1F2937`).

> O `.xlsx` não embute fonte: em máquina sem as famílias instaladas o Excel substitui e o
> arquivo continua legível. Para fidelidade total, instale **Bebas Neue**, **Exo 2** e
> **Libre Baskerville** (são as mesmas do HUB, gratuitas no Google Fonts). O corpo do
> texto segue a escala de planilha (9-18 pt), não a escala em px da seção 4.2, que é da tela.

Três abas:

1. **`Cálculo`** — bloco Empresa (razão social + CNPJ formatado), período coberto, mês de
   corte, bloco de veredito com a explicação consolidada em prosa, uma linha por
   estabelecimento e o bloco de ressalvas. `% Exportação` e `Correção` são fórmulas
   (`=IF($I15<>"Calculado","",…)`), então as linhas não calculadas ficam **em branco** — e
   o TOTAL usa `=IF(COUNT(J15:J21)=0,"",SUM(J15:J21))`, que soma só os calculados e não
   mostra zero quando nada qualificou.
2. **`Série Mensal (conferência)`** — material de auditoria, **não integra o cálculo**.
   Uma seção por estabelecimento, uma linha por mês, com a linha do mês de referência
   marcada com `◄` e destacada. Sem coluna de correção, nem real nem hipotética: sugerir o
   resultado de outros meses convidaria à seleção conveniente, que é o que a metodologia
   evita. Célula de saldo vazia = mês sem escrituração.
3. **`Exportações`** — prova documental da existência e do volume da exportação, uma linha
   por (estabelecimento × CFOP × mês). Devoluções ficam em bloco próprio, **fora** do
   total exportado: somá-las ao total faria a aba divergir da série mensal quando o piso
   em zero é acionado. Quando o desfecho é "sem exportação no último mês", esta aba é a
   prova de que a exportação existiu em **outros** meses — o que sustenta a conversa com o
   cliente.

---

## Como rodar local

```bash
pip install flask openpyxl
python test_app.py
# abre http://127.0.0.1:5000
```

`test_app.py` monta a mesma tela (`screen.html/css/js`), registra o blueprint com
`TESTING=True` (libera a permissão localmente) e inclui um **stub** de
`GET /api/cnpj/<cnpj>` — essa rota é do HUB, e o stub só existe para o fluxo
"CNPJ primeiro" poder ser exercitado offline. Nem `test_app.py` nem `test_regressao.py`
vão para o HUB.

Teste com a pasta `fixtures/` — **sem digitar nada**: escolha os 4 arquivos e clique em
Calcular. A ferramenta identifica a matriz `12.345.678/0001-90` e os quatro
estabelecimentos:

| Arquivo | Estabelecimento | Último mês | Desfecho esperado |
|---|---|---|---|
| `..._A_calcula.csv` | `12345678000190` — RS | DEZ/2025 | **Calculado — R$ 50.000,00** (10,00%) |
| `..._B_sem_exportacao.csv` | `12345678000271` — SP | DEZ/2025 | Sem exportação no último mês |
| `..._C_sem_saldo.csv` | `12345678000352` — RS | DEZ/2025 | Sem saldo credor no último mês |
| `..._D_mes_anterior.csv` | `12345678000433` — ES | **NOV/2025** | **Calculado — R$ 12.000,00** (12,00%) |
| | | | **TOTAL: R$ 62.000,00** |

### Regressão

```bash
python test_regressao.py                        # só a fixture sintética
python test_regressao.py "C:\pasta\do\sped"     # + o caso real de referência
```

São 72 verificações, que cobrem os doze testes de aceitação do briefing e a janela de 60 meses. A fixture
sintética é versionada (não tem dado de cliente); os relatórios reais **não** ficam no
repositório — quem os tem passa a pasta como argumento, e sem o argumento o script
anuncia que os testes do caso real não foram executados, em vez de omiti-los em silêncio.

O que a fixture prova, além do total: **A** valida a fórmula e a exclusão das
transferências (se o percentual sair `1,67%` em vez de `10,00%`, o CFOP 5151 entrou
indevidamente no denominador); **B** exportou em AGO/2025 mas não em DEZ/2025 (se
calcular algo, está retrocedendo na série); **C** tem exportação de R$ 30.000,00 em
DEZ/2025 **e** saldo zero — a explicação não pode culpar a exportação; **D** termina em
NOV/2025 e tem de citar NOV/2025.

### Caso real de referência (sem identificar a empresa)

Sete estabelecimentos, 65 meses. Todos caem em não-cálculo — **nenhuma correção
calculada** — e é justamente isso que valida a regra:

| Estab. | UF | Último mês | Faturamento do mês | Exportação | Saldo credor | Status |
|---|---|---|---|---|---|---|
| …0102 | RS | MAI/2026 | 4.343.256,97 | 0,00 | 1.814.875,92 | Sem exportação |
| …0374 | ES | MAI/2026 | 1.332.773,85 | 0,00 | 18.527,16 | Sem exportação |
| …0455 | ES | MAI/2026 | 843.668,47 | 0,00 | 1.343,85 | Sem exportação |
| …0536 | RS | MAI/2026 | 13.400.344,83 | 0,00 | **0,00** | **Sem saldo credor** |
| …0617 | SP | **ABR/2026** | 2.992.360,75 | 0,00 | 921.945,17 | Sem exportação |
| …0706 | RS | MAI/2026 | 0,00 | 0,00 | 56.935,58 | Sem exportação |
| …0889 | RS | MAI/2026 | 34.233,00 | 0,00 | 1.072.174,23 | Sem exportação |

A empresa exportou R$ 1.269.799,77 ao longo dos 65 meses dos arquivos (CFOPs 7101,
6501 e 5501), distribuídos em 10 meses, todos no `…0102` — nenhum deles o último de
apuração. Dentro da janela de 60 meses são 9 meses e R$ 1.266.558,77 (a remessa de
ABR/2021 está prescrita). A aba
`Exportações` mostra essas operações: elas provam que a empresa exporta, ainda que não no
mês de referência.

Números que denunciariam metodologia errada e que o teste de regressão proíbe:
**R$ 40.897,08** (retrocesso na série até o último mês com exportação),
**R$ 1.610.515,70** (soma dos meses com exportação) e cerca de **R$ 8.240,72**
(percentual acumulado do período aplicado ao saldo do último mês).

O check Seção 1 × Seção 19 fecha em **R$ 0,00** nos sete arquivos — é o melhor teste de
que o parser não perde nem duplica linha.

---

## Estrutura

```
logica.py                 funções puras: parser, cascata, templates, Excel
routes.py                 Blueprint /tools/correcao-saldo-credor-icms-exportacao
screen.html               a sub-página (div #tool-correcao-saldo-credor-icms)
screen.css                CSS escopado, prefixo cscie-
screen.js                 JS escopado, prefixo cscie, IIFE
requirements_extra.txt    openpyxl==3.1.5
HUB_METADATA.json         metadados de integração
README.md                 este arquivo
fixtures/                 4 CSV sintéticos, sem dado de cliente
gerar_fixture_sintetica.py  regenera a fixture
test_app.py               servidor local de teste (NÃO vai para o HUB)
test_regressao.py         os 12 testes de aceitação (NÃO vai para o HUB)
```

Integração no HUB:

```python
from correcao_saldo_credor_icms.routes import bp, init_app, set_auth_provider
init_app(app)
set_auth_provider(meu_auth)      # (request) -> (ok: bool, contexto: dict)
app.register_blueprint(bp)
```

A permissão granular é `correcao_saldo_credor_icms_exportacao` (`opt-in`, padrão
`false`). Toda rota autentica (401) e depois checa a permissão (403), em JSON — página
HTML de erro quebraria o `res.json()` da tela.

---

*Desenvolvido por Guilherme — ICMS/IPI. Dúvidas de metodologia: ricardo@efct.com.br.*
