# AutoCert Empresarial V2.0

## Descrição

O **AutoCert Empresarial V2.0** é um software de **Automação** desenvolvido para a criação de Certificados Digitais de Qualidade para peças oriundas de fabricação própria. Ele é projetado para aumentar a eficiência e reduzir custos no processo de certificação de produtos, permitindo que as empresas gerem certificados em grande escala de forma automatizada e prática.

### Principais Funcionalidades:

- **Geração de Certificados**: O software permite gerar automaticamente certificados de qualidade para uma quantidade **ilimitada de itens simultaneamente**, a partir de um arquivo Excel previamente importado.
- **Automação Completa**: Automatiza o processo de preenchimento dos certificados com base nos dados fornecidos, eliminando a necessidade de fazer manualmente.
- **Eficiência e Escalabilidade**: Capaz de processar e certificar um número massivo de produtos de uma só vez, **aumentando a eficiência operacional** e **reduzindo a necessidade de contratação de mão de obra futura**.
- **Economia de Custos**: Reduz significativamente o tempo e os custos associados à geração de certificados, **aumentando a lucratividade** da empresa.
- **Compactação Automática**: Gera arquivos PDF dos certificados e os compacta em um arquivo `.zip`, facilitando o compartilhamento e armazenamento.
- **Integração com Excel**: Suporte completo para planilhas Excel (.xlsx) e execução de macros VBA, garantindo flexibilidade e compatibilidade com sistemas empresariais já existentes.

## Requisitos

- **Python 3.x**
- **Bibliotecas Python**:
  - `tkinter`
  - `pandas`
  - `openpyxl`
  - `pywin32`
- **Microsoft Excel** (necessário para a execução de macros e manipulação de templates)

`pip install pandas openpyxl pywin32`

`python AUTOCERT_EMPRESARIAL_V2.0.pyw`

## Como Usar

- Carrega a planilha de dados clicando no botão "IMPORTE O DOC. SAÍDA".
- Verifica se os produtos e quantidades estão listados corretamente.
- Preenche a Data de Faturamento e o Número da Nota Fiscal.
- Clica em "GERAR CERTIFICADOS DIGITAIS" para iniciar a geração dos certificados.

## Estrutura de Diretórios

- AUTOCERT_EMPRESARIAL_V2.0.pyw: Script principal que contém a lógica do software.
- TEMPLATE_CERTIFICADO_EMPRESARIAL.xltm: Template Excel usado para gerar os certificados.
- Certificados_Gerados/: Diretório onde os certificados gerados em PDF serão salvos.

## Benefícios

- Aumento da Eficiência: Com a automação, o software acelera a criação de certificados, permitindo que um grande volume de itens seja certificado em muito menos tempo.
- Redução de Custos: Ao eliminar a necessidade de trabalho manual, a empresa pode economizar tanto em tempo quanto em recursos humanos, aumentando a lucratividade.
- Escalabilidade: O AutoCert Empresarial pode ser facilmente escalado para certificar quantidades ainda maiores de produtos sem perda de desempenho.







