using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace EXPORTANDO
{
    internal class Acciones
    {
        private List<Alumno> alumnoList = new List<Alumno>
        {
            new Alumno("Ana", 20, "LADD", 25247, DateTime.Today),
            new Alumno("Juan", 18, "ISC", 10007, DateTime.Today),
            new Alumno("Jose", 20, "LDG", 20025, DateTime.Today),
            new Alumno("Angela", 30, "IDM", 10250, DateTime.Today)
        };

        public List<Alumno> MostrarAlumnos()
        {
            return alumnoList;
        }

        public bool ExportarExcel()
        {
            try
            {
                // Definir el nombre del archivo y la ruta en el escritorio
                string nombreArchivo = "Listado.Alumnos.xlsx";
                string rutaEscritorio = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                string rutaCompleta = Path.Combine(rutaEscritorio, nombreArchivo);

                // Crear un nuevo libro de trabajo
                using (var wb = new XLWorkbook())
                {
                    // Agregar una hoja de trabajo
                    var ws = wb.AddWorksheet("Alumnos");

                    // Definir los encabezados de las columnas
                    var encabezados = new[]
                    {
                        "Nombre", "Edad", "Carrera", "Matrícula", "Fecha"
                    };

                    // Insertar los encabezados en la primera fila
                    ws.Cell(1, 1).InsertData(encabezados);

                    // Aplicar formato a los encabezados
                    for (int i = 1; i <= encabezados.Length; i++)
                    {
                        var cell = ws.Cell(1, i);
                        cell.Style.Font.Bold = true;
                        cell.Style.Fill.BackgroundColor = XLColor.LightGray;
                        cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        cell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                    }

                    // Insertar los datos de los alumnos
                    var datos = alumnoList.Select(a => new object[]
                    {
                        a.Nombre,
                        a.Edad,
                        a.Carrera,
                        a.Matricula,
                        a.Fechanacimiento.ToShortDateString()
                    }).ToList();

                    ws.Cell(2, 1).InsertData(datos);

                    // Ajustar el tamaño de las columnas al contenido
                    ws.Columns().AdjustToContents();

                    // Guardar el archivo en la ruta especificada
                    wb.SaveAs(rutaCompleta);
                }

                // Confirmar que el archivo se ha guardado correctamente
                return true;
            }
            catch (Exception)
            {
                // Manejar cualquier error que ocurra durante el proceso
                return false;
            }
        }
    }

}

