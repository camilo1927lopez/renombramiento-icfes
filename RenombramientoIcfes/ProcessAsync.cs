using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;

namespace RenombramientoIcfes
{
    internal class ProcessAsync
    {
        private string NombreArchivo { set; get; }
        private string NombreCarpeta { set; get; }
        private int Inicial { get; set; }
        private int Final { get; set; }
        private int Identificador { get; set; }

        public ProcessAsync(string nombreArchivo, string nombreCarpeta, int inicial, int final, int identificador)
        {
            NombreArchivo = nombreArchivo;
            NombreCarpeta = nombreCarpeta;
            Inicial = inicial;
            Final = final;
            Identificador = identificador;
        }

        public async Task Procesar()
        {
            try
            {
                string HoraDigitalizada = DateTime.Now.Day.ToString() + DateTime.Now.Month + DateTime.Now.Year + DateTime.Now.Hour + DateTime.Now.Minute + DateTime.Now.Second + DateTime.Now.Millisecond;
                bool Proceso = true;
                //bool Proceso = Validaciones(HoraDigitalizada);
                //JPGFiles = GetJPGFiles(NombreCarpeta);
                string RutaMetadataEncontrados = Path.Combine(NombreCarpeta + "\\ImagenesDigitalizadas_" + HoraDigitalizada, "Metadata_" + HoraDigitalizada + "_" + Identificador + ".txt");
                string RutaMetadataNoEncontrados = Path.Combine(NombreCarpeta + "\\ImagenesDigitalizadas_" + HoraDigitalizada, "MetadataNoEcontrados_" + HoraDigitalizada + "_" + Identificador + ".txt");
                GenerarArchivo(new List<string>(), RutaMetadataEncontrados, ";");
                GenerarArchivo(new List<string>(), RutaMetadataNoEncontrados, ";");
                EscribirLineaArchivo("RegistroUsuario|CodigoSitio|NombreDeLaImagen", RutaMetadataEncontrados);
                EscribirLineaArchivo("RegistroUsuario|CodigoSitio|NombreDeLaImagen|CodigoDeBarras|Folio", RutaMetadataNoEncontrados);

                string RutaErroresDocumento = Path.Combine(NombreCarpeta, "ErroresDocumento_" + HoraDigitalizada + ".txt");

                GenerarArchivo(new List<string>(), RutaErroresDocumento, ";");


                //if (JPGFiles.Count <= 0)
                //{
                //    throw new Exception($"En la ruta especificada no hay imagenes para renombrar");
                //}

                if (!GetJPGFile("*", NombreCarpeta))
                {
                    throw new Exception($"En la ruta especificada no hay imagenes para renombrar");
                }
                if (NombreArchivo.Length <= 0)
                {
                    throw new Exception($"Porfavor seleccione la ruta del documento");
                }

                if (NombreCarpeta.Length <= 0)
                {
                    throw new Exception($"Porfavor seleccione la ruta de digitalizacion");
                }

                Microsoft.Office.Interop.Excel.Application xl = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel.Workbook workbook = xl.Workbooks.Open(NombreArchivo);
                Microsoft.Office.Interop.Excel.Worksheet sheet = workbook.Sheets[1];


                int numRows = this.Final;
                int numColumns = 26;     // according to your sample

                for (int rowIndex = this.Inicial; rowIndex <= numRows; rowIndex++)  // assuming the data starts at 1,1
                {
                    string numeroRegistro = string.Empty;
                    string sitio = string.Empty;
                    string codigoBarras = string.Empty;
                    string bloque = string.Empty;
                    string salon = string.Empty;

                    Range cell = (Range)sheet.Cells[rowIndex, 4];
                    if (cell.Value != null)
                    {
                        numeroRegistro = cell.Value.ToString().Replace("\n", " ");
                    }
                    cell = (Range)sheet.Cells[rowIndex, 13];
                    if (cell.Value != null)
                    {
                        sitio = cell.Value.ToString().Replace("\n", " ");
                    }
                    cell = (Range)sheet.Cells[rowIndex, 25];
                    if (cell.Value != null)
                    {
                        codigoBarras = cell.Value.ToString().Replace("\n", " ");
                    }
                    cell = (Range)sheet.Cells[rowIndex, 16];
                    if (cell.Value != null)
                    {
                        bloque = cell.Value.ToString().Replace("\n", " ");
                    }
                    cell = (Range)sheet.Cells[rowIndex, 17];
                    if (cell.Value != null)
                    {
                        salon = cell.Value.ToString().Replace("\n", " ");
                    }

                    #region validaciones


                    if (string.IsNullOrWhiteSpace(numeroRegistro))
                    {
                        EscribirLineaArchivo("Fila: " + rowIndex + " El campo NUMERO_REGISTRO se encuentra nulo", RutaErroresDocumento);
                    }
                    if (string.IsNullOrWhiteSpace(sitio))
                    {
                        EscribirLineaArchivo("Fila: " + rowIndex + " El campo CONSECUTIVO_CIUDAD se encuentra nulo", RutaErroresDocumento);
                    }
                    if (string.IsNullOrWhiteSpace(codigoBarras))
                    {
                        EscribirLineaArchivo("Fila: " + rowIndex + " El campo CODIGO_BARRAS se encuentra nulo", RutaErroresDocumento);
                    }
                    if (string.IsNullOrWhiteSpace(bloque))
                    {
                        EscribirLineaArchivo("Fila: " + rowIndex + " El campo BLOQUE se ecuentra nulo", RutaErroresDocumento);
                    }
                    if (string.IsNullOrWhiteSpace(salon))
                    {
                        EscribirLineaArchivo("Fila: " + rowIndex + " El campo NOMBRE_SALON se encuentra nulo", RutaErroresDocumento);
                    }

                    #endregion

                    var SeparadorSitio = sitio.Split("-"[0]);
                    String CodSitio = SeparadorSitio[0];

                    if (GetJPGFile(codigoBarras, NombreCarpeta))
                    {
                        String NuevoNombre = CodSitio + "_S" + salon + "_B" + bloque + "_F" + codigoBarras.Substring(13, 3);

                        if (!Directory.Exists(Path.Combine(NombreCarpeta + "\\ImagenesDigitalizadas_" + HoraDigitalizada, sitio)))
                        {
                            // Si no existe, crearla
                            Directory.CreateDirectory(Path.Combine(NombreCarpeta + "\\ImagenesDigitalizadas_" + HoraDigitalizada, sitio));
                        }
                        //JPGFiles.Remove(codigoBarras + ".jpg");

                        File.Copy(NombreCarpeta + "\\" + codigoBarras + ".jpg", NombreCarpeta + "\\ImagenesDigitalizadas_" + HoraDigitalizada + "\\" + sitio + "\\" + NuevoNombre + ".jpg", true);
                        EscribirLineaArchivo(numeroRegistro + "|" + sitio + "|" + NuevoNombre, RutaMetadataEncontrados);

                    }
                    else
                    {
                        String NoEncontrado = CodSitio + "_S" + salon + "_B" + bloque + "_F" + codigoBarras.Substring(13, 3);
                        EscribirLineaArchivo(numeroRegistro + "|" + sitio + "|" + NoEncontrado + "|" + codigoBarras + "|" + codigoBarras.Substring(13, 3), RutaMetadataNoEncontrados);
                    }


                }
                xl.Quit();

            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        private bool GetJPGFile(string codebar, string folderPath)
        {
            List<string> pngFiles = new List<string>();

            if (!Directory.Exists(folderPath))
            {
                Console.WriteLine("La ruta especificada no existe.");
                throw new Exception("La ruta especificada no existe imágenes no existe");
            }
            var consulta = Directory.GetFiles(folderPath, codebar + ".jpg");
            return (consulta.Length != 0);

        }

        private void EscribirLineaArchivo(string linea, string ruta)
        {

            //FileStream fs = new FileStream(ruta, FileMode.Open, FileAccess.Read);
            //using (StreamReader reader = new StreamReader(fs))
            //{
            using (StreamWriter writer = new StreamWriter(ruta, true))
            {
                writer.WriteLine(linea);
            }

            //}


        }

        public bool GenerarArchivo(List<string> lista, string RutaArchivo, string separador)
        {
            try
            {
                if (!Directory.Exists(Path.GetDirectoryName(RutaArchivo)))
                    Directory.CreateDirectory(Path.GetDirectoryName(RutaArchivo));


                using (StreamWriter sw = new StreamWriter(RutaArchivo))
                {
                    for (int i = 0; i < lista.Count; i++)
                    {
                        sw.WriteLine(string.Join(separador, lista[i]));
                    }
                }

                return true;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        public static implicit operator Task(ProcessAsync v)
        {
            throw new NotImplementedException();
        }

        public static implicit operator ProcessAsync(Task<Task> v)
        {
            throw new NotImplementedException();
        }
    }


}
