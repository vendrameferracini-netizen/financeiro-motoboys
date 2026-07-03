App Financeiro de Motoboys e Socios

Abra index.html no navegador.

Organizacao atual:
- Dashboard
- Motoboys
- Socios
- Entrada de Pacotes da Base
- Lancamento Diario
- Descontos
- Fechamentos
- Recibos
- Relatorios
- Auditoria
- Configuracoes

Fonte historica:
- A planilha Fechamento 1a de junho (1).xlsx foi convertida para workbook-data.js e agora e a fonte oficial.
- O app usa a planilha como fonte principal dos dados importados.
- Foram carregadas 69 abas e 1.152 colunas.
- Todas as abas importadas continuam disponiveis para auditoria.
- As abas de motoboy viraram fichas individuais, com historico financeiro e fechamento.
- A sincronizacao preserva lancamentos manuais do app em localStorage e atualiza apenas a base importada da planilha.

Socios:
- GIL, SALES e GUILHERME aparecem em destaque.
- Cada socio tem pagina individual com Entrada de Pacotes da Base, controle de descontos, resumo financeiro, fechamento quinzenal e relatorios.
- A aba GUILHERME usa os dados importados da planilha "GUILHERME M".
- Aliases de gestores: GM, G M, GUILHERME M e GUILHERME sao unificados em GUILHERME quando a celula identifica apenas esse gestor.
- A varredura tambem aceita G.M., GM. e Guilherme M. como GUILHERME.
- Celulas ambiguas com mais de um gestor no mesmo texto nao sao atribuidas a um socio especifico.
- A logica do GIL nao foi alterada.

Entrada de Pacotes da Base:
- Campos: data, socio/base, ML, Shopee, total de pacotes, valor ML, valor Shopee, total a pagar, observacao e responsavel.
- Calculo automatico: ML x R$ 8,00 e Shopee x R$ 5,00.
- Esse controle e independente do pagamento dos motoboys.
- O fechamento da base e separado por socio e por quinzena.

Motoboys:
- Lancamento diario com data, motoboy, ML, Shopee, Avulso, valores unitarios, total bruto, responsavel e observacao.
- Fechamento semanal e quinzenal por motoboy.
- Calculo: liquido = bruto - vales - extravios - outros descontos + bonificacoes.
- Recibos podem ser impressos/exportados e marcados como pagos.

Descontos:
- Central unica para Vale, Extravio, Ocorrencia e Outro desconto.
- Cada lancamento aparece no historico do socio responsavel e no historico do motoboy.
- Descontos entram automaticamente no fechamento do motoboy.

Formatacao:
- Todos os valores financeiros sao exibidos em moeda brasileira com Intl.NumberFormat pt-BR.

Auditoria:
- Relatorio de Sincronizacao mostra registros existentes, encontrados, novos, atualizados, removidos, iguais, ignorados e duplicidades evitadas.
- Relatorio de Sincronizacao lista as abas alteradas em relacao a importacao anterior.
- Total de abas encontradas e importadas.
- Total de colunas encontradas e importadas.
- Lista de abas com erro.
- Lista de colunas nao importadas.
- Divergencias encontradas.
- Comparacao entre totais importados, dashboard, relatorios e fichas individuais.
- Descontos ignorados por falta de nome valido.
- Descontos ignorados por nome numerico.
- Descontos ignorados por valor invalido.
- Celulas ignoradas por serem cabecalho ou total.
- Divergencia entrada x saida.

Validacao tecnica:
- App carregado com workbook-data.js sem erro.
- 69 abas encontradas.
- 69 abas importadas.
- 1.152 colunas encontradas/importadas.
- Todos os IDs usados pelo JavaScript existem no HTML.
- Fechamentos gerados apenas com nomes reais de motoboys.
- BASE criado como responsavel separado para descontos consolidados ou sem socio.
- Dashboard exibe comparativo de pacotes e graficos de barras com dados validos.
- Importacao reconstruida a partir da planilha embutida, com chaves unicas para saidas, entradas de base e descontos.
- Descontos mostram aba original, linha original, coluna original e observacao de origem.
- Totais de bloco/relatorio nao sao importados como novos descontos; eles aparecem somente na auditoria de conferencia.
- Auditoria mostra blocos de desconto encontrados, descontos importados, descontos ignorados, duplicidades removidas, totais por responsavel, totais por motoboy e celulas usadas como origem.
- Auditoria mostra tambem os aliases unificados, incluindo GM -> GUILHERME, com motoboy, tipo, valor e origem da celula.
- Auditoria completa GUILHERME mostra aba, linha, coluna, motoboy, valor, tipo, observacao, status e motivo.
- A auditoria separa Total financeiro GUILHERME e Registros GUILHERME rastreados para evitar soma dupla entre aba do socio e fichas de motoboy.
- Dashboard reproduz os totais oficiais da aba FINANCEIRO: Resumo Operacional, Controle por Gestor e Resumo de Caixa.
- Auditoria APP x PLANILHA compara saida de pacotes, folha bruta, descontos, valor liquido, entrada de pacotes, diferenca de pacotes, pagar para base, caixa final e totais por gestor.
- Validacao final deste ajuste: GM nao cria gestor separado, descontos GM entram em GUILHERME, nao ha desconto duplicado e totais/subtotais nao viram item de desconto.
