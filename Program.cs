using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using ClosedXML.Excel;
using ShellProgressBar;

class Program
{
    static void Main()
    {
        Console.WriteLine("Digite o caminho da pasta contendo os arquivos XML:");
        string folderPath = Console.ReadLine();

        if (!Directory.Exists(folderPath))
        {
            Console.WriteLine("O caminho especificado não existe.");
            return;
        }

        var xmlFiles = Directory.GetFiles(folderPath, "*.xml");

        if (xmlFiles.Length == 0)
        {
            Console.WriteLine("Nenhum arquivo XML encontrado na pasta especificada.");
            return;
        }

        List<Product> products = new List<Product>();

        Console.WriteLine("Iniciando leitura dos arquivos XML...");

        var options = new ProgressBarOptions
        {
            ForegroundColor = ConsoleColor.Yellow,
            ForegroundColorDone = ConsoleColor.Green,
            BackgroundColor = ConsoleColor.DarkGray,
            ProgressCharacter = '─'
        };

        using (var progressBar = new ProgressBar(xmlFiles.Length, "Lendo arquivos XML...", options))
        {
            foreach (var file in xmlFiles)
            {
                XDocument doc = XDocument.Load(file);
                var productElements = doc.Descendants("det");

                foreach (var element in productElements)
                {
                    var prodElement = element.Element("prod");
                    var impostoElement = element.Element("imposto");

                    var icmsElement = impostoElement.Descendants("ICMS").FirstOrDefault()?.Elements().FirstOrDefault();
                    var ipiElement = impostoElement.Descendants("IPI").FirstOrDefault()?.Elements().FirstOrDefault();
                    var pisElement = impostoElement.Descendants("PIS").FirstOrDefault()?.Elements().FirstOrDefault();
                    var cofinsElement = impostoElement.Descendants("COFINS").FirstOrDefault()?.Elements().FirstOrDefault();

                    var product = new Product
                    {
                        Name = prodElement.Element("xProd")?.Value,
                        Code = prodElement.Element("cProd")?.Value,
                        Quantity = ParseDecimal(prodElement.Element("qCom")?.Value),
                        Price = ParseDecimal(prodElement.Element("vUnCom")?.Value),
                        TotalPrice = ParseDecimal(prodElement.Element("vProd")?.Value),
                        NCM = prodElement.Element("NCM")?.Value,
                        CFOP = prodElement.Element("CFOP")?.Value,
                        Unit = prodElement.Element("uCom")?.Value,
                        UnitValue = ParseDecimal(prodElement.Element("vUnCom")?.Value),
                        ICMSCST = icmsElement?.Element("CST")?.Value,
                        ICMSRate = ParseDecimal(icmsElement?.Element("pICMS")?.Value),
                        ICMSValue = ParseDecimal(icmsElement?.Element("vICMS")?.Value),
                        IPICST = ipiElement?.Element("CST")?.Value,
                        IPIRate = ParseDecimal(ipiElement?.Element("pIPI")?.Value),
                        IPIValue = ParseDecimal(ipiElement?.Element("vIPI")?.Value),
                        PISCST = pisElement?.Element("CST")?.Value,
                        PISRate = ParseDecimal(pisElement?.Element("pPIS")?.Value),
                        PISValue = ParseDecimal(pisElement?.Element("vPIS")?.Value),
                        COFINSCST = cofinsElement?.Element("CST")?.Value,
                        COFINSRate = ParseDecimal(cofinsElement?.Element("pCOFINS")?.Value),
                        COFINSValue = ParseDecimal(cofinsElement?.Element("vCOFINS")?.Value),
                    };

                    products.Add(product);
                }

                progressBar.Tick();
            }
        }

        Console.WriteLine("Leitura dos arquivos XML concluída. Consolidando produtos...");

        var uniqueProducts = products
            .GroupBy(p => p.Code)
            .Select(g => g.First())
            .ToList();

        Console.WriteLine("Consolidação concluída. Exportando para Excel...");

        string excelPath = Path.Combine(folderPath, "produtos.xlsx");
        ExportToExcel(uniqueProducts, excelPath);

        Console.WriteLine($"Exportação concluída. Planilha salva em {excelPath}");
    }

    static decimal ParseDecimal(string value)
    {
        if (decimal.TryParse(value, NumberStyles.Any, CultureInfo.InvariantCulture, out var result))
        {
            return result;
        }
        return 0;
    }

    static void ExportToExcel(List<Product> products, string filePath)
    {
        var workbook = new XLWorkbook();
        var worksheet = workbook.Worksheets.Add("Produtos");

        worksheet.Cell(1, 1).Value = "Nome";
        worksheet.Cell(1, 2).Value = "Código";
        worksheet.Cell(1, 3).Value = "Quantidade";
        worksheet.Cell(1, 4).Value = "Preço Unitário";
        worksheet.Cell(1, 5).Value = "Preço Total";
        worksheet.Cell(1, 6).Value = "NCM";
        worksheet.Cell(1, 7).Value = "CFOP";
        worksheet.Cell(1, 8).Value = "Unidade";
        worksheet.Cell(1, 9).Value = "Valor Unitário";
        worksheet.Cell(1, 10).Value = "ICMS CST";
        worksheet.Cell(1, 11).Value = "ICMS Alíquota";
        worksheet.Cell(1, 12).Value = "ICMS Valor";
        worksheet.Cell(1, 13).Value = "IPI CST";
        worksheet.Cell(1, 14).Value = "IPI Alíquota";
        worksheet.Cell(1, 15).Value = "IPI Valor";
        worksheet.Cell(1, 16).Value = "PIS CST";
        worksheet.Cell(1, 17).Value = "PIS Alíquota";
        worksheet.Cell(1, 18).Value = "PIS Valor";
        worksheet.Cell(1, 19).Value = "COFINS CST";
        worksheet.Cell(1, 20).Value = "COFINS Alíquota";
        worksheet.Cell(1, 21).Value = "COFINS Valor";

        var culture = new CultureInfo("pt-BR");

        for (int i = 0; i < products.Count; i++)
        {
            worksheet.Cell(i + 2, 1).Value = products[i].Name;
            worksheet.Cell(i + 2, 2).Value = products[i].Code;
            worksheet.Cell(i + 2, 3).Value = products[i].Quantity.ToString("N2", culture);
            worksheet.Cell(i + 2, 4).Value = products[i].Price.ToString("N2", culture);
            worksheet.Cell(i + 2, 5).Value = products[i].TotalPrice.ToString("N2", culture);
            worksheet.Cell(i + 2, 6).Value = products[i].NCM;
            worksheet.Cell(i + 2, 7).Value = products[i].CFOP;
            worksheet.Cell(i + 2, 8).Value = products[i].Unit;
            worksheet.Cell(i + 2, 9).Value = products[i].UnitValue.ToString("N2", culture);
            worksheet.Cell(i + 2, 10).Value = products[i].ICMSCST;
            worksheet.Cell(i + 2, 11).Value = products[i].ICMSRate.ToString("N2", culture);
            worksheet.Cell(i + 2, 12).Value = products[i].ICMSValue.ToString("N2", culture);
            worksheet.Cell(i + 2, 13).Value = products[i].IPICST;
            worksheet.Cell(i + 2, 14).Value = products[i].IPIRate.ToString("N2", culture);
            worksheet.Cell(i + 2, 15).Value = products[i].IPIValue.ToString("N2", culture);
            worksheet.Cell(i + 2, 16).Value = products[i].PISCST;
            worksheet.Cell(i + 2, 17).Value = products[i].PISRate.ToString("N2", culture);
            worksheet.Cell(i + 2, 18).Value = products[i].PISValue.ToString("N2", culture);
            worksheet.Cell(i + 2, 19).Value = products[i].COFINSCST;
            worksheet.Cell(i + 2, 20).Value = products[i].COFINSRate.ToString("N2", culture);
            worksheet.Cell(i + 2, 21).Value = products[i].COFINSValue.ToString("N2", culture);
        }

        workbook.SaveAs(filePath);
    }
}
class Product
{
    public string Name { get; set; }
    public string Code { get; set; }
    public decimal Quantity { get; set; }
    public decimal Price { get; set; }
    public decimal TotalPrice { get; set; }
    public string NCM { get; set; }
    public string CFOP { get; set; }
    public string Unit { get; set; }
    public decimal UnitValue { get; set; }
    public string ICMSCST { get; set; }
    public decimal ICMSValue { get; set; }
    public decimal ICMSRate { get; set; }
    public string IPICST { get; set; }
    public decimal IPIValue { get; set; }
    public decimal IPIRate { get; set; }
    public string PISCST { get; set; }
    public decimal PISValue { get; set; }
    public decimal PISRate { get; set; }
    public string COFINSCST { get; set; }
    public decimal COFINSValue { get; set; }
    public decimal COFINSRate { get; set; }
}

