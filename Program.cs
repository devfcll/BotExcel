using OfficeOpenXml;
using System;

class Program
{
    static void Main(string[] args)
    {
        // Solicita os dados da pessoa ao usuário
        Console.WriteLine("Digite o nome:");
        string nome = Console.ReadLine();

        Console.WriteLine("Digite a idade:");
        int idade = int.Parse(Console.ReadLine());

        Console.WriteLine("Digite o e-mail:");
        string email = Console.ReadLine();

        Console.WriteLine("Digite o telefone:");
        string telefone = Console.ReadLine();

        Console.WriteLine("Digite o endereço:");
        string endereco = Console.ReadLine();

        // Cria um novo arquivo Excel e adiciona os dados da pessoa a ele
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        using (var package = new ExcelPackage())
        {
            // Cria uma planilha e define o nome
            ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Pessoas");

            // Adiciona os cabeçalhos das colunas
            worksheet.Cells[1, 1].Value = "Nome";
            worksheet.Cells[1, 2].Value = "Idade";
            worksheet.Cells[1, 3].Value = "E-mail";
            worksheet.Cells[1, 4].Value = "Telefone";
            worksheet.Cells[1, 5].Value = "Endereço";

            // Adiciona os dados da pessoa na linha seguinte
            worksheet.Cells[2, 1].Value = nome;
            worksheet.Cells[2, 2].Value = idade;
            worksheet.Cells[2, 3].Value = email;
            worksheet.Cells[2, 4].Value = telefone;
            worksheet.Cells[2, 5].Value = endereco;

            // Salva o arquivo Excel
            package.SaveAs(new System.IO.FileInfo("D://Planilhas//pessoas.xlsx"));
        }

        Console.WriteLine("Dados salvos com sucesso no arquivo pessoas.xlsx!");
        Console.ReadKey();
    }
}
