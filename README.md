# Removedor de Acentos

Aplicação desktop desenvolvida em **Python** para remover acentos de arquivos de forma automatizada, utilizando uma interface gráfica simples e intuitiva. Este projeto nasceu de uma necessidade encontrada durante minha atuação com suporte e implantação de sistemas ERP. 

O projeto permite processar arquivos **XML, XLS, XLSX e CSV**, convertendo textos como:

```text
João → Joao
São Paulo → Sao Paulo
Informação → Informacao
Ação → Acao
```

A aplicação foi desenvolvida com foco em automação de tarefas relacionadas a arquivos exportados de sistemas, ERPs e bases de dados que exigem textos sem caracteres acentuados.

---

## Funcionalidades

* Remoção automática de acentos.
* Interface gráfica desenvolvida com `Tkinter`.
* Suporte aos formatos:

  * XML
  * XLS
  * XLSX
  * CSV
* Conversão entre formatos de planilhas.
* Detecção do formato real do arquivo.
* Tratamento de arquivos com extensão incorreta.
* Suporte a arquivos Excel XML / SpreadsheetML.
* Detecção automática de codificação de arquivos CSV.
* Detecção automática do separador do CSV.
* Preservação de fórmulas e formatação no processamento direto de arquivos XLSX.
* Barra de progresso durante o processamento.
* Processamento em thread para evitar travamentos da interface.
* Confirmação antes de substituir arquivos existentes.
* Mensagens de erro e sucesso.
* Geração automática do nome do arquivo de saída.

---

## Tecnologias utilizadas

* Python
* Tkinter
* Pandas
* OpenPyXL
* xlrd
* xlwt
* XML ElementTree

---

## Estrutura do funcionamento

O fluxo principal da aplicação é:

```text
Selecionar arquivo
       ↓
Identificar extensão
       ↓
Identificar formato real do arquivo
       ↓
Processar conteúdo
       ↓
Remover acentos
       ↓
Escolher formato de saída
       ↓
Salvar novo arquivo
```

---

## Formatos suportados

### XML

Processa textos, atributos e conteúdo textual dos elementos XML.

Exemplo:

```xml
<cliente nome="João">
    <cidade>São Paulo</cidade>
</cliente>
```

Resultado:

```xml
<cliente nome="Joao">
    <cidade>Sao Paulo</cidade>
</cliente>
```

---

### XLSX

Arquivos XLSX válidos são processados diretamente com `openpyxl`.

Esse caminho foi desenvolvido para evitar a reconstrução completa da planilha através do Pandas e, consequentemente, reduzir o risco de perda de formatação.

Fórmulas são preservadas.

---

### XLS

Arquivos no formato Excel antigo podem ser lidos utilizando `xlrd` e gerados utilizando `xlwt`.

O formato XLS possui limitações próprias, incluindo:

* Máximo de 65.536 linhas por planilha.
* Máximo de 256 colunas por planilha.

---

### CSV

A aplicação tenta identificar automaticamente:

* UTF-8 com BOM
* UTF-8
* CP1252
* Latin-1

Também tenta identificar automaticamente o separador:

```text
,
;
TAB
|
```

---

## Detecção do formato real

Uma das funcionalidades do projeto é verificar o conteúdo do arquivo antes de tentar processá-lo.

Isso é importante porque alguns sistemas legados podem gerar arquivos como:

```text
relatorio.xls
```

mesmo que o conteúdo real seja:

```text
XML / SpreadsheetML
```

ou HTML.

Nesse caso, simplesmente utilizar a extensão `.xls` pode causar erros durante a leitura.

A aplicação analisa a assinatura e o conteúdo inicial do arquivo para identificar formatos como:

```text
XLSX
XLS
SpreadsheetML
HTML
```

---

## Instalação

Clone o repositório:

```bash
git clone https://github.com/SEU-USUARIO/removedor-acentos-arquivos.git
```

Entre na pasta:

```bash
cd removedor-acentos-arquivos
```

Recomenda-se utilizar um ambiente virtual:

```bash
python -m venv venv
```

No Windows:

```bash
venv\Scripts\activate
```

Instale as dependências:

```bash
pip install -r requirements.txt
```

---

## Dependências


```text
openpyxl
pandas
xlrd
xlwt
```

O `tkinter` normalmente já acompanha instalações do Python para Windows.

---

## Execução

Execute:

```bash
python main.py
```

ou
basta baixar apenas o executável que esta na pasta dist

A aplicação abrirá uma interface gráfica onde será possível:

1. Selecionar o arquivo de entrada.
2. Escolher o formato de saída.
3. Definir o local do arquivo.
4. Processar o arquivo.
5. Salvar uma nova versão sem acentos.

---

## Exemplo de utilização

Arquivo original:

```text
clientes.xlsx
```

Após o processamento:

```text
clientes_sem_acentos.xlsx
```

Exemplo de transformação:

| Antes         | Depois        |
| ------------- | ------------- |
| João da Silva | Joao da Silva |
| São Paulo     | Sao Paulo     |
| Informação    | Informacao    |
| José          | Jose          |
| Paraná        | Parana        |

---

## Interface

A aplicação possui uma interface gráfica com:

* Seleção do arquivo de entrada
* Seleção do formato de saída
* Seleção do arquivo de destino
* Indicador de processamento
* Mensagens de status
* Tratamento visual de erros e conclusão

> Sugestão: adicione aqui uma captura de tela da aplicação.

Exemplo:

```markdown
![Interface da aplicação](screenshots/interface.png)
```

---

## Limitações

* CSV não suporta múltiplas planilhas. Quando a origem possui várias planilhas, somente a primeira é exportada para CSV.
* O formato XLS possui limitações de linhas e colunas.
* A conversão entre formatos pode não preservar todos os recursos avançados de uma planilha original.
* Arquivos protegidos por senha ou criptografados não são processados.
* O processamento de XLSX é otimizado para preservar fórmulas e formatação quando a entrada e saída são ambas XLSX.

---

## Possíveis melhorias futuras

* [ ] Modularizar Codigo
* [ ] Adicionar processamento em lote de vários arquivos.
* [ ] Permitir arrastar e soltar arquivos.
* [ ] Adicionar suporte a arquivos `.ods`.
* [ ] Criar testes automatizados.
* [ ] Criar instalador para Windows.
* [ ] Adicionar logs de processamento.
* [ ] Permitir escolher quais colunas devem ser processadas.
* [ ] Criar modo de processamento em lote.
* [ ] Melhorar o relatório de erros.
* [ ] Adicionar ícone próprio à aplicação.

---

## Objetivo do projeto

O projeto foi criado como uma ferramenta de automação para situações em que arquivos precisam ser padronizados antes de serem importados para outros sistemas.

Além da utilidade prática, o projeto demonstra conceitos de:

* Python
* Automação
* Manipulação de arquivos
* Processamento de dados
* Tkinter
* Pandas
* Excel
* XML
* Tratamento de exceções
* Threads
* Desenvolvimento de aplicações desktop

---

## Licença

Este projeto pode ser distribuído sob a licença MIT.

Consulte o arquivo `LICENSE` para mais informações.

---

## Autor

Desenvolvido por **Cosme**.

