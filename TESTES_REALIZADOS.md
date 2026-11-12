# Testes realizados

## Núcleo

- autotestes das regras de normalização, duração, associação bibliográfica e score;
- teste de deduplicação e limite máximo de dez chaves;
- teste de rotação após resposta `quotaExceeded`;
- teste de continuação com a segunda chave sem repetir trabalho concluído;
- verificação de sintaxe de `core.js`, `xlsx-lite.js` e `app.js`.

## Interface

- inicialização da aplicação autocontida em navegador Chromium;
- importação por seletor de arquivos;
- preenchimento de duas chaves;
- simulação das respostas de `search.list` e `videos.list`;
- simulação de cota encerrada na primeira chave;
- acompanhamento de progresso e métricas;
- conclusão de três títulos, com três suspeitas e uma troca automática;
- download do relatório pelo botão da interface.

## Planilhas

- leitura da planilha de referência fornecida: 696 títulos reconhecidos na aba `Catálogo Completo`;
- detecção automática das colunas de título e autor;
- geração das abas `Resumo`, `Ranking por Views`, `Vídeos Suspeitos` e `Catálogo Completo`;
- verificação estrutural do ZIP interno do XLSX;
- releitura do relatório gerado;
- verificação de freeze panes, autofiltros e hyperlinks.

## Observação

Os testes de consulta foram feitos com respostas simuladas da API para evitar usar credenciais reais ou consumir cota. Não foi executada uma busca real no YouTube neste ambiente.
