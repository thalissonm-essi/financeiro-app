# Abastece

App de abastecimento e manutenção dos veículos. Mesma estrutura do Anota (index.html instalável + Google Apps Script com planilha), projeto separado.

## Instalar (uma vez)
1. Crie uma Planilha Google nova chamada "Abastece".
2. Na planilha: Extensões > Apps Script. Apague o conteúdo e cole `Codigo.gs`.
3. Configurações do projeto > Propriedades do script: adicione `ANTHROPIC_API_KEY` (a mesma do Anota serve).
4. No editor, rode a função `ativar` e autorize. O log mostra a **chave do app**.
5. Implantar > Nova implantação > App da Web. Executar como: você. Quem pode acessar: qualquer pessoa. Copie a URL que termina em `/exec`.
6. Em `index.html`, troque `COLE_AQUI_A_URL_DO_APPS_SCRIPT` por essa URL.
7. Publique a pasta como o Anota (GitHub Pages), abra no Safari do iPhone, Compartilhar > Adicionar à Tela de Início, e cole a chave.

Ao mudar o `Codigo.gs` depois: Implantar > Gerenciar implantações > editar > Nova versão (a URL continua a mesma).
