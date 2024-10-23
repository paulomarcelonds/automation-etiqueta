# Geração de Etiquetas em PDF com QR Code

Este projeto permite a geração de etiquetas em formato PDF a partir de dados contidos em um arquivo Excel. Cada etiqueta contém informações detalhadas sobre produtos, incluindo um QR Code que representa uma lista de números seriais. 

## Funcionalidades

- **Leitura de Arquivo Excel**: O código lê um arquivo Excel que contém os dados necessários para a geração das etiquetas.
- **Formatação de Data**: As datas são convertidas para o formato `dd/mm/yyyy`.
- **Criação de QR Code**: Gera um QR Code que contém uma lista de números seriais.
- **Geração de PDF**: Produz um arquivo PDF contendo as etiquetas formatadas.

## Requisitos

- Python 3.12
- Bibliotecas:
  - `pandas`
  - `reportlab`
  - `qrcode`

Você pode instalar as bibliotecas necessárias usando o seguinte comando:

```bash
pip install pandas reportlab qrcode

Uso
Para usar o script, siga os passos abaixo:

Prepare um arquivo Excel com as seguintes colunas:

CAIXA
NOME
DATA
CD
CIDADE
COD._ITEM
DESCRICAO
N._Nfe
LOTE
SERIAL

Atualize o caminho do arquivo Excel no código:

excel_file = r"Caminho\para\seu\arquivo.xlsx"

Execute o script. O PDF será gerado na mesma pasta onde o script é executado.

python nome_do_seu_script.py

Funções
format_date(date_str): Converte uma string de data para o formato dd/mm/yyyy.
create_qr_code(serials): Gera um QR Code a partir de uma lista de números seriais.
generate_label_pdf(excel_file, output_pdf): Lê um arquivo Excel e gera um PDF com as etiquetas formatadas.
main(): Função principal que executa o processo de geração do PDF.

Exemplo de Saída

Após a execução do script, um arquivo PDF chamado etiquetas_saida.pdf será gerado, contendo as etiquetas formatadas com informações dos produtos e os QR Codes correspondentes.

Contribuição
Sinta-se à vontade para contribuir com melhorias ou correções. Para relatar problemas, utilize a seção de issues do repositório.

Licença
Este projeto é licenciado sob a MIT License.

Sinta-se à vontade para ajustar qualquer parte para melhor se adequar ao seu projeto!
