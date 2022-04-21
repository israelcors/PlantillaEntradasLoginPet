using ExcelDataReader;
using LeerExcel.Models;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text.Json;
namespace LeerExcel
{
    class Program
    {
        static void Main(string[] args)
        {
            ////Ruta del fichero Excel
            string filePath = @"E:\Users\Israel\Downloads\pruebaEntradas.xlsx";

            //using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
            //{
            //    System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            //    using (var reader = ExcelReaderFactory.CreateReader(stream))
            //    {

            //        var result = reader.AsDataSet();

            //        // Ejemplos de acceso a datos
            //        DataTable table = result.Tables[0];
            //        DataRow row = table.Rows[0];
            //        string cell = row[0].ToString();
            //    }
            //}
            Plantilla plantillaLinea = new Plantilla();

            List<Plantilla> plantilla = new List<Plantilla>();
            string HoraLeida = "";
            int cantDiasTotales = 0;
            int cantDiasTrabajados = 0;
            int cantDiasNoTrabajados = 0;
            string NombreDia="", NombreDiaPivot="X";
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            FileInfo existingFile = new FileInfo(filePath);
            using (ExcelPackage package = new ExcelPackage(existingFile))
            {
                //get the first worksheet in the workbook
                for (int i = 0; i < package.Workbook.Worksheets.Count; i++)
                {
                    plantillaLinea = new Plantilla();
                    Console.WriteLine("nombre Hoja: "+ package.Workbook.Worksheets[i].Name);
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[i];
                    plantillaLinea.CodEmpleado = Int32.Parse(package.Workbook.Worksheets[i].Name);
                    plantillaLinea.Nombre = worksheet.Cells[5, 1].Value?.ToString().Trim();
                    int colCount = 15; //worksheet.Dimension.End.Column;  //get Column Count
                    int rowCount = 51; //worksheet.Dimension.End.Row;     //get row count
                    cantDiasTotales = 0;
                    cantDiasTrabajados = 0;
                    cantDiasNoTrabajados = 0;
                    for (int row = 7; row <= rowCount; row++)
                    {
                        NombreDia = worksheet.Cells[row, 1].Value?.ToString().Trim();
                        Console.WriteLine("Procesando dia {0}", NombreDia);
                        if (NombreDiaPivot == NombreDia)
                        {
                            NombreDia = NombreDia+"_2";
                        }
                        else { 
                            cantDiasTotales += 1;
                            HoraLeida = worksheet.Cells[row, 12].Value?.ToString().Trim();
                            Console.WriteLine("horaLeida: " + HoraLeida);
                            if (HoraLeida == null)
                            {
                                cantDiasNoTrabajados += 1;
                            }
                            else
                            {
                                cantDiasTrabajados += 1;
                            }
                        }
                        
                        for (int col = 12; col <= colCount; col++)
                        {
                            NombreDiaPivot= worksheet.Cells[row, 1].Value?.ToString().Trim();

                        }
                    }
                    plantillaLinea.TotalDias = cantDiasTotales;
                    plantillaLinea.DiasNoTrabajados = cantDiasNoTrabajados;
                    plantillaLinea.DiasTrabajados = cantDiasTrabajados;
                    plantilla.Add(plantillaLinea);
                }
               
            }
            var json = JsonSerializer.Serialize(plantilla);
            Console.WriteLine(json);
            Console.ReadKey();
        }

        static void Otro()
        {
            var Entradas = new[]
            {
                new {
                    Id = "101", Name = "C++"
                }
            };

            // Creating an instance
            // of ExcelPackage
            ExcelPackage excel = new ExcelPackage();

            // name of the sheet
            var workSheet = excel.Workbook.Worksheets.Add("Resumen");

            // setting the properties
            // of the work sheet 
            workSheet.TabColor = System.Drawing.Color.Black;
            workSheet.DefaultRowHeight = 12;

            // Setting the properties
            // of the first row
            workSheet.Row(1).Height = 20;
            workSheet.Row(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            workSheet.Row(1).Style.Font.Bold = true;

            // Header of the Excel sheet
            workSheet.Cells[1, 1].Value = "S.No";
            workSheet.Cells[1, 2].Value = "Id";
            workSheet.Cells[1, 3].Value = "Name";

            // Inserting the article data into excel
            // sheet by using the for each loop
            // As we have values to the first row 
            // we will start with second row
            int recordIndex = 2;

            foreach (var article in Entradas)
            {
                workSheet.Cells[recordIndex, 1].Value = (recordIndex - 1).ToString();
                workSheet.Cells[recordIndex, 2].Value = article.Id;
                workSheet.Cells[recordIndex, 3].Value = article.Name;
                recordIndex++;
            }

            // By default, the column width is not 
            // set to auto fit for the content
            // of the range, so we are using
            // AutoFit() method here. 
            workSheet.Column(1).AutoFit();
            workSheet.Column(2).AutoFit();
            workSheet.Column(3).AutoFit();

            // file name with .xlsx extension 
            string p_strPath = "H:\\geeksforgeeks.xlsx";

            if (File.Exists(p_strPath))
                File.Delete(p_strPath);

            // Create excel file on physical disk 
            FileStream objFileStrm = File.Create(p_strPath);
            objFileStrm.Close();

            // Write content to excel file 
            File.WriteAllBytes(p_strPath, excel.GetAsByteArray());
            //Close Excel package
            excel.Dispose();
            Console.ReadKey();

        }
    }
}
