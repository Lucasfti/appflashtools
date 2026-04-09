Flash Tool
Este repositório contém o código-fonte do Flash Tool, um aplicativo desktop desenvolvido em Python (com customtkinter) para automatizar a geração de scripts de UPDATE, ROLLBACK e gerenciar a injeção de dados no banco Oracle (SIPPES).

Abaixo estão as instruções para configurar o ambiente de desenvolvimento e realizar o build do executável.

1. Configuração do Ambiente Virtual
É recomendado utilizar um ambiente virtual para isolar as dependências do projeto. Abra o terminal na raiz do projeto e execute:

Criação do ambiente virtual:

PowerShell
python -m venv venv
Permissão de execução (Apenas Windows/PowerShell):
Caso o Windows bloqueie a ativação do script, libere a execução para o usuário atual:

PowerShell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
Ativação do ambiente:

PowerShell
venv\Scripts\activate
2. Instalação de Dependências
Com o ambiente virtual ativado (indicado por (venv) no terminal), instale as bibliotecas listadas no projeto:

PowerShell
pip install -r requirements.txt
3. Validação do Ambiente
Para confirmar que os pacotes críticos foram instalados corretamente e estão acessíveis, execute o teste rápido abaixo:

PowerShell
python -c "import tkinter, customtkinter, pandas, oracledb, pyperclip; print('Dependências instaladas com sucesso.')"
Se o terminal retornar a mensagem de sucesso sem erros, o ambiente está pronto.

4. Geração do Executável (Build)
Para empacotar o script Python e suas dependências em um programa executável para os usuários finais, utilize o PyInstaller.

Execute o comando abaixo no terminal:

PowerShell
pyinstaller --noconsole --onedir --clean --collect-all customtkinter --collect-all cryptography --collect-all tkinter -n "Flash Tool" meu_gerador.py
Detalhamento dos parâmetros utilizados:
--noconsole: Oculta a janela de prompt de comando (CMD) ao executar o aplicativo.

--onedir: Gera uma pasta contendo o executável e todas as dependências externas.

--clean: Limpa o cache do PyInstaller antes de iniciar a compilação, evitando conflitos.

--collect-all: Força a coleta manual de pacotes e arquivos de dados críticos para bibliotecas de interface e segurança (customtkinter, cryptography, tkinter).

-n "Flash Tool": Define o nome do arquivo gerado.

Resultado:
Ao término do processo, acesse a pasta dist/Flash Tool gerada na raiz do projeto. O aplicativo estará pronto para uso no arquivo Flash Tool.exe.