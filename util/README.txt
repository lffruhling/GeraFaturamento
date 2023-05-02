Versão Mínima do Python: 3.10.*^
Versão Mínima do PIP: 23.0.1^

##################
Utilização do Venv
    python -m venv venv
    .\venv\Scripts\activate

##################
Atualização do PIP
    python -m pip install --trusted-host pypi.python.org --trusted-host files.pythonhosted.org --trusted-host pypi.org --upgrade pip

##################
Instalação via Git SubModule, após iniciar o projeto e vicular o mesmo com o git, executar o comando abaixo.
    https://gitlab.coop.sicredi.in/central_sul/coop0313/python/util.git

##################
Arquivos importados e utilizados nos projetos RPA em Python

Arquivo legado.py => Contém as funções responsáveis por abrir e manipular os sistemas legados,
tais como, abrir o ACclient, abrir e logar no SIAT,SIAC, etc. Também possui as funções
responsáveis por copiar a tela e localizar mensagens para as automações se localizarem.

Arquivo funcoes.py => Possui funcoes genéricas que são utilizadas pelo arquivo legado.py e
também pelas automações. Funções como recuperar a chave de criptografia, realizar conexão
ao banco de dados das credenciais, validar dias, etc.

Arquivo criptografia.py => Este arquivo de fato tem as funções que cria/recupera a chave 
de criptografia utilizadas nas credenciais. Também criptografa e descriptografa as 
credenciais para quando forem chamadas para logar nos sistemas.

Arquivo contstantes.py => Utilizado para carregar valores que serão utilizados pelos arquivos
citados acima.

O arquivo requeriments.txt traz a lista de bibliotecas externas que devem ser adicionadas
ao python.

pip install -r util/requeriments.txt

###########################
Ajustes na Pyscreeze
Acessar venv/Lib/pyscreeze/__init__.py
Adicionar a linha 21: import numpy as np
Alterar na linha 167: img_cv = cv2.imdecode(np.fromfile(img, dtype=np.uint8), LOAD_GRAYSCALE)
Alterar na linha 169: img_cv = cv2.imdecode(np.fromfile(img, dtype=np.uint8), LOAD_COLOR)

###########################
Ajuste Legado para capturar a tela.

Abrir o Legado, clicar com o botão direito, acessar a aba Fonte e setar o tamanho para 10 x 18
