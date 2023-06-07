**README**

# Certificate Generator

Este é um programa em Python que gera certificados personalizados a partir de um modelo em formato DOCX e uma lista de participantes em um arquivo CSV.

## Requisitos

Certifique-se de ter os seguintes requisitos instalados antes de executar o programa:

- Python 3.x
- Biblioteca `docx`
- Biblioteca `csv`
- Biblioteca `smtplib`
- Biblioteca `subprocess`
- Biblioteca `os`
- Biblioteca `win32com.client`

## Instalação

1. Clone o repositório ou faça o download do código-fonte para o seu computador.

2. Certifique-se de ter o Python 3.x instalado. Caso não tenha, você pode fazer o download no site oficial do Python (https://www.python.org).

3. Instale as bibliotecas necessárias executando o seguinte comando em seu terminal:
```
pip install python-docx
pip install pywin32
```

## Uso

Siga as etapas abaixo para utilizar o programa:

1. Coloque o arquivo CSV contendo a lista de participantes no mesmo diretório do código-fonte. Certifique-se de que o arquivo CSV esteja formatado corretamente e que a coluna contendo os nomes dos participantes seja intitulada "Aluno".

2. Certifique-se de ter um modelo de certificado em formato DOCX chamado "certificadopadrao.docx" no mesmo diretório do código-fonte. Este será o modelo utilizado para gerar os certificados personalizados. No modelo do certificado, coloque a palavra "name" no local onde o nome do participante deve ser inserido. Por exemplo, "Certificado para name" ou "name".

3. Abra um terminal e navegue até o diretório do código-fonte.

4. Execute o seguinte comando para gerar os certificados:
```
python certificate_generator.py
```
Isso processará o arquivo CSV e gerará um certificado personalizado para cada participante listado, substituindo a palavra "name" pelo nome correspondente.

5. Os certificados serão salvos em um diretório chamado "certificados" no mesmo diretório do código-fonte.

## Observações

- Certifique-se de ter permissão para criar diretórios no local onde o programa está sendo executado, para que o diretório "certificados" possa ser criado e os certificados possam ser salvos corretamente.

- O programa usa o pacote `pywin32` para converter os documentos do Word em PDF. Certifique-se de que você tenha o pacote `pywin32` instalado corretamente e funcionando em seu sistema.

- Verifique se todas as dependências e bibliotecas foram instaladas corretamente antes de executar o programa. Caso contrário, você pode instalá-las usando o comando `pip install nome_da_biblioteca`.

## Contribuição

Contribuições são bem-vindas! Se você encontrar algum problema, tiver sugestões ou quiser adicionar recursos ao programa, sinta-se à vontade para criar uma solicitação de pull.

## Licença

Este projeto está licenciado sob a licença MIT. Consulte o arquivo LICENSE para obter mais informações.
