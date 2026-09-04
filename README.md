# Relatório Gerencial PAD Saúde+

Sistema de relatório gerencial em HTML, CSS e JavaScript.

## Importação
Aceita:
- CSV
- XLSX
- XLS

Todos os indicadores e gráficos são calculados com os registros importados.

## Colunas reconhecidas
O sistema reconhece variações de:
- Data / Data Atendimento
- Horário Início / Hora Início / Início
- Horário Fim / Hora Fim / Fim
- Paciente / Nome
- Especialidade
- Status

## Correções desta versão
- Datas Excel e datas brasileiras tratadas corretamente.
- CSV UTF-8 e Windows-1252.
- Filtros de data, status e especialidade.
- Gráfico diário/semanal baseado nos registros.
- Especialidades calculadas do arquivo.
- Status calculado do arquivo.
- Tempo médio somente quando início e fim são válidos.
- Tabela sem dados fictícios.
- Nenhum número 1.248 ou porcentagem fixa no painel.

## Atualização da tabela
A coluna Status foi substituída por Tempo de Atendimento, calculado por registro a partir dos horários de início e fim. O importador foi reforçado para reconhecer diferentes nomes de colunas de horário.

## Mapeamento do CSV Telemedicina
- **Data:** `data_fim` → `data_ultima_finalização_medico` → `data_turno` → `data_convite`
- **Horário Início:** `data_ultimo_entrada_medico_vc`
- **Horário Fim:** `data_ultima_finalização_medico`
- **Tempo de Atendimento:** `t4_duração_consulta` (campo oficial da consulta)
- **Paciente:** `nome`
- **Especialidade:** `especialidade`
- **Situação:** `situação_consulta`

## Tabela Últimos Atendimentos
A tabela foi simplificada para exibir apenas Data, Paciente, Especialidade e Tempo de Atendimento. Os horários de início e fim continuam sendo usados internamente para calcular o tempo.

## Correção v6
A tabela voltou a exibir Data, Horário Início, Horário Fim, Paciente, Especialidade e Tempo de Atendimento. Registros sem horário/tempo calculável são ocultados da tabela, em vez de mostrar `—`.

## Correção v7
Ao selecionar Status, Especialidade ou alterar o período, a tabela é atualizada imediatamente. O título muda para `Atendimentos Filtrados`. Registros filtrados sem horário/tempo permanecem na tabela e exibem `—`.

## Correção v8 — Realizado x Agendado
O gráfico Atendimentos por Dia agora compara duas séries: **Agendado**, baseado em data_turno/data agendada, e **Realizado**, baseado em data_fim e/ou registros efetivamente finalizados. As duas barras são exibidas lado a lado por data.

## PDF
Adicionado o botão **Salvar em PDF**. O sistema captura visualmente o painel inteiro como ele aparece na tela e salva em PDF A4 horizontal.

## Legibilidade v11
Os textos que estavam em cinza claro foram escurecidos para melhorar contraste e leitura, mantendo a identidade visual do painel.

## Negrito na tabela
Os textos da tabela de Atendimentos Filtrados foram colocados em negrito para maior destaque e leitura.

## Negrito geral v14
Os textos do painel, filtros, KPIs, títulos, legendas, gráficos, período, tabela e rodapé foram reforçados em negrito para melhorar a visibilidade.

## PDF v15 — gráficos de rosca
Os gráficos de rosca de Especialidade e Status foram convertidos de `conic-gradient` para SVG inline. Isso garante que apareçam também na captura feita pelo botão Salvar em PDF.
