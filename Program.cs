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
            string filePath = @"C:\Users\Israel\source\repos\LeerExcel\bin\Release\net5.0\detalleEntrada.xlsx";
            //string filePath = "";
            //filePath = args[0];

            int Feriados = 0;
            int Vacaciones = 0;
            int Enfermedad_corta = 0;
            int Ausente_sin_aviso = 0;
            int Entro_tarde = 0;
            int Accidente = 0;

            Plantilla plantillaLinea = new Plantilla();

            List<Plantilla> plantilla = new List<Plantilla>();
            string HoraLeida = "";
            string Observacion = "";
            string Novedades = "";
            int cantDiasTotales = 0;
            int cantDiasTrabajados = 0;
            int cantDiasNoTrabajados = 0;
            string NombreDia = "", NombreDiaPivot = "X", nombrePlanilla = "";
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
             FileInfo existingFile = new FileInfo(filePath);
            using (ExcelPackage package = new ExcelPackage(existingFile))
            {
                //get the first worksheet in the workbook
                for (int i = 0; i < package.Workbook.Worksheets.Count; i++)
                
                {
                    if (package.Workbook.Worksheets[i].Name != "Resumen")
                    { 
                        NombreDiaPivot = "X";
                        plantillaLinea = new Plantilla();
                        Console.WriteLine("nombre Hoja: " + package.Workbook.Worksheets[i].Name);
                        nombrePlanilla = package.Workbook.Worksheets[i].Name;
                        ExcelWorksheet worksheet = package.Workbook.Worksheets[i];

                        Feriados = 0;
                        Vacaciones = 0;
                        Enfermedad_corta = 0;
                        Ausente_sin_aviso = 0;
                        Entro_tarde = 0;
                        Accidente = 0;

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
                            if (EsDiaSemana(NombreDia) && NombreDia.ToUpper() != "DOMINGO")
                            {
                                Console.WriteLine("Procesando dia {0}", NombreDia);
                                if (NombreDiaPivot == NombreDia)
                                {
                                    NombreDia = NombreDia + "_2";
                                }
                                else
                                {
                                    cantDiasTotales += 1;
                                    HoraLeida = worksheet.Cells[row, 12].Value?.ToString().Trim();
                                    Observacion = worksheet.Cells[row, 21].Value?.ToString().Trim();
                                    Novedades = worksheet.Cells[row, 20].Value?.ToString().Trim();
                                    Console.WriteLine("horaLeida: " + HoraLeida);
                                    if (HoraLeida == null && (Novedades == null || Observacion == null))
                                    {
                                        cantDiasNoTrabajados += 1;
                                        switch (Novedades)
                                        {
                                            case "Vacaciones":
                                                Vacaciones += 1;
                                                break;
                                            case "Enfermedad Corta":
                                                Enfermedad_corta += 1;
                                                break;
                                            case "Feriado":
                                                Feriados += 1;
                                                break;
                                            case "Ausente sin aviso":
                                                Ausente_sin_aviso += 1;
                                                break;
                                            case "Accidente":
                                                Accidente += 1;
                                                break;
                                            case "Entro tarde":
                                                Entro_tarde += 1;
                                                break;

                                        }
                                    }
                                    else
                                    {
                                        cantDiasTrabajados += 1;
                                    }
                                }


                            NombreDiaPivot = worksheet.Cells[row, 1].Value?.ToString().Trim();

                            }
                        }
                        plantillaLinea.TotalDias = cantDiasTotales;
                        plantillaLinea.DiasNoTrabajados = cantDiasNoTrabajados;
                        plantillaLinea.DiasTrabajados = cantDiasTrabajados;
                        plantillaLinea.Vacaciones = Vacaciones;
                        plantillaLinea.Enfermedad_corta = Enfermedad_corta;
                        plantillaLinea.Feriados = Feriados;
                        plantillaLinea.Ausente_sin_aviso = Ausente_sin_aviso;
                        plantillaLinea.Accidente = Accidente;
                        plantillaLinea.Entro_tarde = Entro_tarde;

                        plantilla.Add(plantillaLinea);
                    }
                }
                
            }
            var json = JsonSerializer.Serialize(plantilla);
            Console.WriteLine(json);

            GuardaNuevaHojaResumen(filePath, plantilla);

            Console.WriteLine("Proceso terminado, apriete una tecla para finalizar...");
            Console.ReadKey();
        }


        static bool EsDiaSemana(string nombreDia)
        {
            if (nombreDia is null)
                return false;
            if ((nombreDia == "LUNES") ||
                    (nombreDia == "MARTES") ||
                    (nombreDia == "MIÉRCOLES") ||
                    (nombreDia == "JUEVES") ||
                    (nombreDia == "VIERNES") ||
                    (nombreDia == "SÁBADO") ||
                    (nombreDia == "DOMINGO"))
                    {
                        return true;
                    }
            else
            {
                return false;
            }
        }

        static void GuardaNuevaHojaResumen(string filePath,List<Plantilla> plantilla)
        {
            using (ExcelPackage excel = new ExcelPackage(filePath))
            {

                //((Excel.Worksheet)this.Application.ActiveWorkbook.Sheets[4]).Delete();
                try
                {
                    ExcelWorksheet worksheetA1 = excel.Workbook.Worksheets.SingleOrDefault(x => x.Name == "Resumen");
                    if (worksheetA1 != null)
                    { 
                        excel.Workbook.Worksheets.Delete(worksheetA1);
                        Console.WriteLine("Hoja eliminada");
                        excel.Save();
                    }
                    else
                    {
                        excel.Save();
                    }
                }
                catch (Exception e)
                {
                    Console.WriteLine("Error eliminando Hoja: ",e.Message);
                }
                


               


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
                workSheet.Cells[1, 3].Value = "Nombre";
                workSheet.Cells[1, 4].Value = "DiasTotales";
                workSheet.Cells[1, 5].Value = "DiasTrabajados";
                workSheet.Cells[1, 6].Value = "DiasNOTrabajados";
                workSheet.Cells[1, 7].Value = "Vacaciones";
                workSheet.Cells[1, 8].Value = "Enfermedad_corta";
                workSheet.Cells[1, 9].Value = "Feriados";
                workSheet.Cells[1, 10].Value = "Ausente_sin_aviso";
                workSheet.Cells[1, 11].Value = "Accidente";
                workSheet.Cells[1, 12].Value = "Entro_tarde";


                // Inserting the article data into excel
                // sheet by using the for each loop
                // As we have values to the first row 
                // we will start with second row
                int recordIndex = 2;

                foreach (var Plantilla in plantilla)
                {
                    workSheet.Cells[recordIndex, 1].Value = (recordIndex - 1).ToString();
                    workSheet.Cells[recordIndex, 2].Value = Plantilla.CodEmpleado;
                    workSheet.Cells[recordIndex, 3].Value = Plantilla.Nombre;
                    workSheet.Cells[recordIndex, 4].Value = Plantilla.TotalDias;
                    workSheet.Cells[recordIndex, 5].Value = Plantilla.DiasTrabajados;
                    workSheet.Cells[recordIndex, 6].Value = Plantilla.DiasNoTrabajados;
                    workSheet.Cells[recordIndex, 7].Value = Plantilla.Vacaciones;
                    workSheet.Cells[recordIndex, 8].Value = Plantilla.Enfermedad_corta;
                    workSheet.Cells[recordIndex, 9].Value = Plantilla.Feriados;
                    workSheet.Cells[recordIndex, 10].Value = Plantilla.Ausente_sin_aviso;
                    workSheet.Cells[recordIndex, 11].Value = Plantilla.Accidente;
                    workSheet.Cells[recordIndex, 12].Value = Plantilla.Entro_tarde;
                    recordIndex++;

                }

                // By default, the column width is not 
                // set to auto fit for the content
                // of the range, so we are using
                // AutoFit() method here. 
                workSheet.Column(1).AutoFit();
                workSheet.Column(2).AutoFit();
                workSheet.Column(3).AutoFit();
                workSheet.Column(4).AutoFit();
                workSheet.Column(5).AutoFit();
                workSheet.Column(6).AutoFit();
                workSheet.Column(7).AutoFit();
                workSheet.Column(8).AutoFit();
                workSheet.Column(9).AutoFit();
                workSheet.Column(10).AutoFit();
                workSheet.Column(11).AutoFit();
                workSheet.Column(12).AutoFit();

                excel.Save();
              
            }


        }
    }
}
