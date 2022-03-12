# ExcelAutomation

## Aula sobre a biblioteca xlwings para automação Excel em Python

Nessa aula desenvolvi com meus alunos dois projetos:

- app.py
- app_multithread.py

### app.py

Primeiros comandos do xlwings para criar uma planilha de caixa, com formatação condicional

### app_multithread.py

Primeira aplicação utilizando threads no python para abrir várias planilhas do Excel e gravar números aleatórios.
Falamos também sobre as restrições impostas pelo GIL (Global Interpreter Locker)

### Para executar:

No cmd:

> git clone https://github.com/rafaeltedesco/ExcelAutomation.git

> cd ExcelAutomation

Para abrir com VisualCode

> code .

Para inicializar o ambiente virtual e instalar as dependências:

> pipenv shell

Caso o interpretador não seja identificado corretamente no VisualCode:

> CTRL + SHIFT + P

Procurar por "Select Interpreter" ou "Selecionar Interpretador"

Filtrar nos ambientes virtuais pelo que inicia com "ExcelAutomation"

Executar com <code>CTRL + F5</code> ou <code>F5</code> para o modo debug com o arquivo app.py ou app_multrithread selecionado
