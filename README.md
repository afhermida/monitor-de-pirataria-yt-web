# Monitor de Audiolivros no YouTube — versão web local

Aplicação web estática para pesquisar no YouTube possíveis disponibilizações não autorizadas de audiolivros listados em um catálogo Excel. A interface substitui o uso de terminal e roda diretamente no navegador.

## Início rápido

1. Salve `MONITOR_YOUTUBE_OFFLINE.html` no computador.
2. Abra o arquivo com um navegador moderno, preferencialmente Chrome, Edge ou Firefox.
3. Selecione a planilha `.xlsx`.
4. Informe de uma a dez chaves da YouTube Data API v3.
5. Clique em **Iniciar monitoramento**.
6. Revise os resultados e baixe o relatório Excel.

A aplicação não precisa de Python, `pip`, terminal, servidor local ou instalador específico para Windows, macOS ou Linux. Ela precisa de internet somente para consultar a API do YouTube.

## O que está incluído

- leitura de catálogos `.xlsx` no próprio navegador;
- detecção automática da aba e das colunas de título e autor;
- mapeamento manual, caso os cabeçalhos sejam diferentes;
- até dez chaves de API;
- troca automática em caso de cota, chave inválida, restrição incompatível ou API desativada;
- perfis de classificação **Equilibrado** e **Compatível com a planilha original**;
- exclusão de canais autorizados por ID ou nome;
- filtros de região, idioma, data, duração, score e quantidade de resultados;
- cache local e checkpoint transportável em JSON;
- pausa, retomada e reinício;
- painel de resultados durante a execução;
- exportação de Excel com as abas `Resumo`, `Ranking por Views`, `Vídeos Suspeitos` e `Catálogo Completo`;
- nenhum download de áudio ou vídeo.

## Formato da planilha de entrada

O formato mínimo é:

| Título | Autor |
|---|---|
| Dom Casmurro | Machado de Assis |
| O Alienista | Machado de Assis |

A aplicação procura automaticamente aliases comuns, como `Título`, `Livro`, `Title`, `Autor`, `Autora` e `Author`. Se não reconhecer a estrutura, use o painel de mapeamento para escolher a aba, a linha de cabeçalho e as duas colunas.

O arquivo `catalogo_exemplo.xlsx` pode ser usado para um primeiro teste.

## Chaves da API

As chaves podem ser digitadas individualmente ou importadas de um TXT. O formato recomendado é uma chave por linha:

```text
CHAVE_1
CHAVE_2
CHAVE_3
```

Linhas iniciadas por `#` são comentários. Valores repetidos são eliminados e somente os dez primeiros valores únicos são usados.

### Privacidade das chaves

As chaves permanecem na memória da página e não entram no relatório Excel, no cache ou no checkpoint. Elas são enviadas somente nas chamadas feitas diretamente pelo navegador à API do Google/YouTube. Ainda assim, trate-as como credenciais: não compartilhe capturas de tela com os campos revelados e não guarde chaves reais em arquivos públicos.

### Restrições recomendadas

Restrinja a chave à **YouTube Data API v3**. Uma página aberta como `file://` não possui um domínio HTTP normal; por isso, uma chave limitada por **referenciador HTTP** pode ser recusada. Quando a política da organização exigir restrição por domínio, publique os mesmos arquivos estáticos em uma intranet HTTPS e autorize somente esse domínio. Não remova restrições de segurança sem avaliação da equipe responsável pelo Google Cloud.

### Cota e várias chaves

A interface aceita até dez chaves e muda automaticamente para a próxima quando a atual fica indisponível. Isso não significa necessariamente dez cotas independentes: chaves pertencentes ao mesmo projeto Google Cloud normalmente compartilham a cota do projeto.

O limite local automático é de 95 novos títulos por chave configurada. Esse limite serve apenas como margem operacional da aplicação. O campo **Limite local de buscas** pode ser alterado; `0` remove apenas o limite local, nunca a cota real do Google.

## Canais autorizados

É possível informar um canal por linha. Prefira o ID, porque nomes podem mudar ou se repetir:

```text
id:UCxxxxxxxxxxxxxxxxxxxxxx
nome:Canal oficial da editora
```

Linhas iniciadas por `#` são ignoradas. Um ID que começa com `UC` também pode ser informado sem o prefixo `id:`.

## Perfis de classificação

### Equilibrado

É o perfil recomendado. Exige correspondência bibliográfica mais forte entre a obra e o vídeo e penaliza resumos, resenhas e associações frágeis. Foi criado para reduzir casos como um livro de filosofia sendo associado a outro texto do mesmo autor ou vídeos apenas relacionados ao tema.

### Compatível com a planilha original

Mantém uma classificação mais permissiva, próxima da lógica reconstruída a partir da planilha de referência.

O score combina sinais como expressão de obra completa no título, tipo de canal, duração, canal não autorizado, correspondência do título e termos encontrados na descrição. O resultado é apenas um indicador para revisão humana.

## Cache, pausa e checkpoint

A aplicação tenta salvar resultados no IndexedDB do navegador. Quando o navegador não permite persistência para arquivos locais, ela usa cache apenas durante a sessão e mostra esse estado no topo da página.

- **Pausar:** interrompe a consulta atual e preserva o que já foi processado.
- **Continuar:** retoma a partir do próximo título.
- **Checkpoint:** exporta um JSON com catálogo, resultados e configurações, sem as chaves de API.
- **Apagar cache:** remove somente o cache da aplicação naquele navegador/perfil.

O cache é associado ao título, autor, consulta e parâmetros relevantes. Alterar os critérios de score recalcula os resultados sem refazer a busca; alterar região, idioma, data ou quantidade de resultados pode exigir novas consultas.

## Relatório de saída

O relatório contém:

- **Resumo:** métricas, parâmetros e uso das chaves;
- **Ranking por Views:** uma linha por obra suspeita, usando o vídeo mais visualizado;
- **Vídeos Suspeitos:** todos os vídeos acima do score mínimo;
- **Catálogo Completo:** situação de cada título, inclusive pendências.

As URLs do YouTube são exportadas como hyperlinks. O relatório não contém as chaves de API.

## Dados enviados pela aplicação

A ferramenta não possui backend próprio. O arquivo Excel e os resultados ficam no navegador. Para executar uma pesquisa, o navegador envia diretamente à API do YouTube:

- a chave ativa;
- a consulta formada a partir do título e do autor;
- filtros como região, idioma, data e quantidade de resultados;
- IDs de vídeos retornados, para obter duração, visualizações e metadados.

Use a ferramenta somente com catálogos e credenciais cuja utilização seja autorizada pela organização.

## Limitações

- A busca depende do índice e da ordem de relevância do YouTube no momento da consulta.
- Um score alto não prova infração, e um score baixo não prova autorização.
- Licenças, domínio público, territorialidade, canais oficiais e permissões contratuais exigem revisão humana.
- Vídeos removidos, privados ou alterados depois da consulta podem deixar o relatório desatualizado.
- A aplicação não envia notificações nem pedidos de remoção.
- Não há execução em segundo plano depois que a página é fechada.

## Solução de problemas

### A página não abre corretamente

Salve o arquivo com extensão `.html` e abra-o em um navegador, não no visualizador de texto ou no preview do e-mail. Se estiver usando o ZIP, extraia-o antes. O arquivo principal, porém, é autocontido e também pode ser baixado separadamente.

### A chave é recusada imediatamente

Confirme que a YouTube Data API v3 está habilitada no projeto, que a chave não expirou e que as restrições de aplicação aceitam o contexto local. Uma restrição por domínio costuma ser incompatível com `file://`.

### Todas as chaves acabam juntas

Elas podem pertencer ao mesmo projeto e compartilhar a mesma cota. A interface mostra a troca, mas não cria cota adicional.

### O navegador mostra “cache apenas nesta sessão”

Exporte um checkpoint antes de fechar a página. Outra opção é servir os arquivos da pasta `codigo-fonte` em uma intranet HTTPS, o que normalmente oferece um armazenamento persistente mais previsível.

### O catálogo foi mapeado incorretamente

Escolha manualmente a aba, a linha de cabeçalho, a coluna de título e a coluna de autor; depois clique em **Aplicar seleção**.

## Código-fonte

A pasta `codigo-fonte` contém a versão modular:

- `index.html` — estrutura da interface;
- `styles.css` — apresentação responsiva;
- `app.js` — interação, cache, importação e exportação;
- `core.js` — busca, rotação de chaves e classificação;
- `xlsx-lite.js` — leitura e criação de XLSX;
- `vendor/jszip.min.js` — compactação ZIP usada pelo formato XLSX.

`MONITOR_YOUTUBE_OFFLINE.html` incorpora todos esses arquivos e é a versão indicada para usuários finais.

## Documentação oficial relacionada

- YouTube Data API — introdução e cota: https://developers.google.com/youtube/v3/getting-started
- `search.list`: https://developers.google.com/youtube/v3/docs/search/list
- `videos.list`: https://developers.google.com/youtube/v3/docs/videos/list
- CORS nas APIs do Google: https://developers.google.com/api-client-library/javascript/features/cors

## Versão

`2.0.0-web`, empacotada em 12 de novembro de 2026.
