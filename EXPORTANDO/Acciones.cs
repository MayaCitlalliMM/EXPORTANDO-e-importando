using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Net;

namespace EXPORTANDO
{
    internal class Acciones
    {
        private List<Alumno> alumnoList = new List<Alumno>
        {
           
        };

        Correo correo = new Correo();

        public List<Alumno> MostrarAlumnos()
        {
            try
            {
                return alumnoList;
            }
            catch (Exception ex)
            {
                correo.EnviarCorreo(ex.ToString());
                throw;
            }
        }

        public bool ExportarExcel()
        {
            try
            {
                var rutaEscritorio = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                var rutaArchivo = Path.Combine(rutaEscritorio, "Alumnos.xlsx");

                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add("Alumnos");

                    // Encabezados
                    worksheet.Cell(1, 0).Value = "Nombre";
                    worksheet.Cell(1, 2).Value = "Edad";
                    worksheet.Cell(1, 3).Value = "Carrera";
                    worksheet.Cell(1, 4).Value = "Matricula";
                    worksheet.Cell(1, 5).Value = "Fecha de Nacimiento";

                    // Datos
                    for (int i = 0; i < alumnoList.Count; i++)
                    {
                        var alumno = alumnoList[i];
                        worksheet.Cell(i + 2, 1).Value = alumno.Nombre;
                        worksheet.Cell(i + 2, 2).Value = alumno.Edad;
                        worksheet.Cell(i + 2, 3).Value = alumno.Carrera;
                        worksheet.Cell(i + 2, 4).Value = alumno.Matricula;
                        worksheet.Cell(i + 2, 5).Value = alumno.Fechanacimiento.ToShortDateString();
                    }

                    workbook.SaveAs(rutaArchivo);
                }

                return true;
            }
            catch (Exception ex)
            {
                correo.EnviarCorreo(ex.ToString());
                return false;
            }
        }

      
        public bool Importar()
        {
            try
            {
                var desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                var filePath = Path.Combine(desktopPath, "Alumnos.xlsx");

                if (!File.Exists(filePath))
                {
                    return false;
                }

                using (var workbook = new XLWorkbook(filePath))
                {
                    var worksheet = workbook.Worksheet("Alumnos");

                    var rows = worksheet.RowsUsed().Skip(1); // Saltar los encabezados

                    alumnoList.Clear(); // Limpiar la lista antes de importar nuevos datos

                    foreach (var row in rows)
                    {
                        var nombre = row.Cell(1).Value.ToString();

                        // Usar TryParse para evitar excepciones en la conversión
                        int edad = 0;
                        int.TryParse(row.Cell(2).Value.ToString(), out edad); // Intentar convertir a entero

                        var carrera = row.Cell(3).Value.ToString();

                        int matricula = 0;
                        int.TryParse(row.Cell(4).Value.ToString(), out matricula); // Intentar convertir a entero

                        DateTime fechaIngreso;
                        DateTime.TryParse(row.Cell(5).Value.ToString(), out fechaIngreso); // Intentar convertir a DateTime

                        // Agregar el Alumno solo si la conversión fue exitosa
                        alumnoList.Add(new Alumno(nombre, edad, carrera, matricula, fechaIngreso));
                    }
                }

                return true;
            }
            catch (Exception ex)
            {
                correo.EnviarCorreo(ex.ToString());
                return false;
            }
        }
     
    }
}

