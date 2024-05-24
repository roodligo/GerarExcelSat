# Conversor de XML para Excel para Produtos SAT

Aplicativo de console em C# lê arquivos XML, extrai informações de produtos, incluindo detalhes fiscais, e gera uma planilha Excel.

## Funcionalidades

- Lê arquivos XML de uma pasta especificada
- Extrai informações de produtos, incluindo:
  - Nome
  - Código
  - Quantidade
  - Preço
  - Preço Total
  - NCM
  - CFOP
  - Unidade
  - Valor Unitário
  - ICMS CST
  - Alíquota ICMS
  - Valor ICMS
  - IPI CST
  - Alíquota IPI
  - Valor IPI
  - PIS CST
  - Alíquota PIS
  - Valor PIS
  - COFINS CST
  - Alíquota COFINS
  - Valor COFINS
- Consolida produtos únicos
- Exporta dados para uma planilha Excel com valores decimais formatados (vírgula e duas casas decimais)
- Exibe progresso com uma barra de progresso no console

## Pré-requisitos

- .NET 6.0 SDK ou superior
- Pacote NuGet ShellProgressBar
- Pacote NuGet ClosedXML