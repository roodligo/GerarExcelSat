# Conversor de XML para Excel para Produtos SAT

Aplicativo de console em C# l� arquivos XML, extrai informa��es de produtos, incluindo detalhes fiscais, e gera uma planilha Excel.

## Funcionalidades

- L� arquivos XML de uma pasta especificada
- Extrai informa��es de produtos, incluindo:
  - Nome
  - C�digo
  - Quantidade
  - Pre�o
  - Pre�o Total
  - NCM
  - CFOP
  - Unidade
  - Valor Unit�rio
  - ICMS CST
  - Al�quota ICMS
  - Valor ICMS
  - IPI CST
  - Al�quota IPI
  - Valor IPI
  - PIS CST
  - Al�quota PIS
  - Valor PIS
  - COFINS CST
  - Al�quota COFINS
  - Valor COFINS
- Consolida produtos �nicos
- Exporta dados para uma planilha Excel com valores decimais formatados (v�rgula e duas casas decimais)
- Exibe progresso com uma barra de progresso no console

## Pr�-requisitos

- .NET 6.0 SDK ou superior
- Pacote NuGet ShellProgressBar
- Pacote NuGet ClosedXML