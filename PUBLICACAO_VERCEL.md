# Publicação do MARÉ na Vercel e no Supabase

1. Crie um projeto Supabase dedicado. No painel **Connect**, copie a URL do **Session pooler** e substitua o marcador de senha pela senha administrativa do banco. O script solicita a URL com entrada oculta. Em automação, a URL pode ser fornecida por `MARE_DATABASE_ADMIN_URL`.
2. Com Python 3.12, instale as dependências com `python -m pip install "Flask>=3.1,<4" "psycopg[binary]>=3.2,<4"` e execute `python migrar_supabase.py` na pasta que contém o SQLite e o modelo Word. Informe uma senha exclusiva para a função `mare_aplicacao`. A importação inclui transações já confirmadas no WAL, preserva IDs, recusa um destino preenchido e força RLS nas tabelas do esquema privado `mare`.
3. Importe o repositório na Vercel. Configure `MARE_DATABASE_URL` com a URL do **Transaction pooler**, trocando o usuário para `mare_aplicacao.<referência-do-projeto>` e usando a senha dessa função. Codifique caracteres reservados da senha na URL. Configure também `MARE_WEB_USER` e `MARE_WEB_PASSWORD` para o acesso compartilhado da equipe. As variáveis são segredos do servidor. A conexão administrativa é exclusiva da migração.
4. Valide a exigência de login, a listagem, os apontamentos e o download DOCX em um deploy de teste antes de disponibilizar o endereço à equipe.

O GitHub recebe o código e os recursos visuais. O banco SQLite e o modelo Word ficam fora do repositório. O modelo usado pela Vercel vem da tabela privada `mare.modelos`. A política RLS permite o acesso à função da aplicação, que atende à equipe autenticada. As permissões de cada tabela limitam as operações disponíveis.

O arquivo `ESQUEMA_SUPABASE.sql` cria as tabelas vazias com RLS no Editor SQL. A migração pode usar essas tabelas vazias e cria a função de acesso da aplicação após a importação.

O servidor Windows usa SQLite por padrão. A versão web usa o PostgreSQL e baixa o DOCX preenchido. Para PDF, abra o DOCX no Word e exporte pelo próprio Word. Antes de trocar a operação para a nuvem, encerre os cadastros no SQLite para que a cópia de migração represente o estado final.
