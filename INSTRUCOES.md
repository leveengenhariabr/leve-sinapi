# SINAPI automático (GitHub) — passo a passo

O GitHub baixa a planilha da Caixa todo mês, separa só Mato Grosso e guarda os arquivos. O LeveApp lê esses arquivos (aba **Pesquisa SINAPI**). Ele guarda os **12 meses mais recentes** de MT. Para outro estado, mês ou ano, você pede na hora (parte 5).

Esta pasta já vem com **outubro/2026 pronto** em `dados/MT`, então a aba funciona logo depois da parte 3.

## 1. Criar o repositório (uma vez)
Use a mesma conta do GitHub do CNO.
1. Entre em https://github.com e faça login.
2. Clique no **+** (canto superior direito) › **New repository**.
3. Em "Repository name" digite exatamente `leve-sinapi`.
4. Marque **Public** (público). Pode: são dados abertos da Caixa, sem nada da empresa.
5. Clique **Create repository**.

## 2. Enviar os arquivos
1. Na página do repositório vazio, clique em **uploading an existing file**.
2. Abra a pasta `github-sinapi` (que veio no zip), selecione **tudo que está dentro dela** (Ctrl+A), e arraste para a página do GitHub. Isso inclui as pastas `dados` e `entrada` e os arquivos `sinapi_gerar.py`, `sinapi_atualizar.py`, `.gitignore`.
   - Se o GitHub não aceitar pastas arrastadas, envie primeiro os arquivos soltos e depois entre em **Add file › Upload files** de novo para cada pasta.
3. Espere terminar o envio e clique **Commit changes**.
4. Falta o arquivo da rotina automática. Clique em **Add file › Create new file**. No campo do nome digite exatamente `.github/workflows/atualizar-sinapi.yml` (ao digitar a barra o GitHub cria as pastas).
5. Abra o arquivo `atualizar-sinapi.yml.txt` desta pasta com o Bloco de Notas, copie tudo, cole na área grande do GitHub e clique **Commit changes**.

## 3. Ligar o LeveApp ao repositório
O endereço dos arquivos já vem preenchido em `js/config.js` com o seu usuário do GitHub. Confira o nome: https://raw.githubusercontent.com/leveengenhariabr/leve-sinapi/main/dados/MT/indice.json deve abrir no navegador mostrando texto com `"meses"`. Se abrir, está certo.

## 4. Deixar a atualização automática funcionar
1. No repositório, aba **Settings** › **Actions** › **General**.
2. Role até "Workflow permissions", marque **Read and write permissions** e clique **Save**.
3. Aba **Actions**. Se pedir, clique em "I understand my workflows, go ahead and enable them".
4. Clique em **Atualizar SINAPI** na lista da esquerda › **Run workflow** › **Run workflow** (deixe o campo "Mês" vazio).
5. Espere 2 a 5 minutos. Bolinha verde = deu certo. Depois disso roda sozinho do dia 10 ao dia 25 de cada mês.

## 5. Outro estado, outro mês ou outro ano (sob demanda)
No LeveApp, em **Pesquisa SINAPI**, escolha o **Estado**. Se ainda não houver dados dele, o app mostra o que fazer. Só é trazido o mês que você pedir.

**Jeito A — pelo GitHub (a rotina baixa da Caixa; funciona bem de 2025 em diante):**
1. Aba **Actions** › **Atualizar SINAPI** › **Run workflow**.
2. Em "Estado" digite a sigla (ex.: `GO`). Em "Mês" digite `AAAA-MM` (ex.: `2025-03`), ou `12meses` para os últimos 12 meses. Clique **Run workflow**.
3. Em 2 a 5 minutos aparece no LeveApp. Meses/estados trazidos assim ficam guardados (não são apagados).

**Jeito B — você mesmo baixa o zip e joga na pasta `entrada` (obrigatório para meses até 2024, pois a Caixa bloqueia o download automático):**
- De 2025 em diante: baixe `SINAPI-AAAA-MM-formato-xlsx.zip` no site da Caixa. Sem prefixo = Mato Grosso. Para outro estado, renomeie com a sigla na frente: `GO_SINAPI-2025-03-formato-xlsx.zip`.
- Até 2024: baixe DOIS zips do estado/mês (um Desonerado e um NaoDesonerado), ex.: `SINAPI_ref_Insumos_Composicoes_MT_202305_Desonerado.zip` e `..._NaoDesonerado.zip`. Não renomeie. O mês só aparece quando os dois chegarem.
- No GitHub: pasta `entrada` › **Add file › Upload files** › arraste os zips › **Commit changes**. A rotina roda sozinha e os zips somem da pasta.

## 6. Outras bases de preços (GOINFRA e SICRO)

Elas **não têm download automático** (os sites não permitem). Você baixa as planilhas e solta na pasta `entrada`; o GitHub reconhece pelo **conteúdo** (não pelo nome) e gera tudo sozinho.

**GOINFRA – Obras Civis (Goiás).** Por mês, 6 arquivos (.xls): Composições, Materiais e Mão de obra, cada um **com** e **sem** desoneração. Envie os 6 juntos.

**SICRO (DNIT) – por estado e mês.** Baixe do site do DNIT (SICRO › Relatórios › Centro-Oeste › estado › ano › mês) estes arquivos e **não renomeie** (o nome começa com a sigla do estado, ex.: `MT_07-2026_...`):
- Relatório Analítico de Composições de Custos
- Relatório Sintético de Materiais
- Relatório Sintético de Equipamentos (e o “com desoneração”)
- Relatório Sintético de Mão de Obra (e o “com desoneração”)
- Relatório Analítico de Mão de Obra (opcional; dá o salário-base)
Os demais relatórios do SICRO podem ir junto: são descartados.

Como enviar: GitHub › repositório **leve-sinapi** › pasta **entrada** › **Add file** › **Upload files** › arraste os arquivos › **Commit changes**. Em 2 a 5 minutos (o SICRO demora um pouco mais) a aba **Actions** mostra o processamento e o mês aparece no LeveApp.
Se faltar algum arquivo do conjunto, o log diz quais faltam e os arquivos esperam na pasta.

## Se algo der errado
- Bolinha vermelha em Actions: clique nela e mande um print da mensagem.
- LeveApp diz "Os arquivos do SINAPI ainda não foram gerados": o endereço em `js/config.js` está errado ou o repositório não está público.
