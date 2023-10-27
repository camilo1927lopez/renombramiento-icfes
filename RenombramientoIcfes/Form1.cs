using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using Microsoft.Office.Interop.Excel;
using System.Threading;

namespace RenombramientoIcfes
{
    public partial class Form1 : Form
    {
        List<String> ErroresDatosNull = new List<String>();
        List<String> JPGFiles = new List<string>();
        List<String> MetadataEncontrados = new List<string>();
        List<String> MetadataNoEncontrados = new List<string>();
        int contador = 2;

        public Form1()
        {
            InitializeComponent();
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

        public static List<string> GetJPGFiles(string folderPath)
        {
            List<string> pngFiles = new List<string>();

            if (!Directory.Exists(folderPath))
            {
                Console.WriteLine("La ruta especificada no existe.");
                return pngFiles;
            }

            foreach (string file in Directory.GetFiles(folderPath, "*.jpg"))
            {
                pngFiles.Add(Path.GetFileName(file));
            }

            return pngFiles;
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

        //private Boolean Validaciones(string HoraDigitalizada)
        //{
        //    try
        //    {
        //        string RutaErroresDocumento = Path.Combine(Digitalizadas.CarpetaSeleccionado, "ErroresDocumento_" + HoraDigitalizada + ".txt");

        //        ErroresDatosNull.Clear();
        //        GenerarArchivo(new List<string>(), RutaErroresDocumento, ";");
        //        bool validacion = true;

        //        ISheet objWorksheet;
        //        string strPath = ArchivoCargado.ArchivoSeleccionado;
        //        using (var fStream = new FileStream(strPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
        //        {
        //            fStream.Position = 0;
        //            XSSFWorkbook objWorkbook = new XSSFWorkbook(fStream);
        //            objWorksheet = objWorkbook.GetSheetAt(0);
        //            IRow objHeader = objWorksheet.GetRow(0);
        //            int countCells = objHeader.LastCellNum;
        //            if (countCells != 26)
        //            {
        //                throw new Exception($"El documento no cumple con la estructuta, {Environment.NewLine} el archivo debe tener 25 columnas y se estan enviando {countCells}");
        //            }
        //            for (int i = (objWorksheet.FirstRowNum + 1); i <= objWorksheet.LastRowNum; i++)
        //            {
        //                IRow objRow = objWorksheet.GetRow(i);
        //                string numeroRegistro = objRow.GetCell(3).ToString().Replace("\n", " ");
        //                string sitio = objRow.GetCell(12).ToString().Replace("\n", " ");
        //                string codigoBarras = objRow.GetCell(24).ToString().Replace("\n", " ");
        //                string bloque = objRow.GetCell(15).ToString().Replace("\n", " ");
        //                string salon = objRow.GetCell(16).ToString().Replace("\n", " ");

        //                if (string.IsNullOrWhiteSpace(numeroRegistro))
        //                {
        //                    EscribirLineaArchivo("Fila: " + i + " El campo NUMERO_REGISTRO se encuentra nulo", RutaErroresDocumento);
        //                    validacion = false;
        //                }
        //                if (string.IsNullOrWhiteSpace(sitio))
        //                {
        //                    EscribirLineaArchivo("Fila: " + i + " El campo CONSECUTIVO_CIUDAD se encuentra nulo", RutaErroresDocumento);
        //                    validacion = false;
        //                }
        //                if (string.IsNullOrWhiteSpace(codigoBarras))
        //                {
        //                    EscribirLineaArchivo("Fila: " + i + " El campo CODIGO_BARRAS se encuentra nulo", RutaErroresDocumento);
        //                    validacion = false;
        //                }
        //                if (string.IsNullOrWhiteSpace(bloque))
        //                {
        //                    EscribirLineaArchivo("Fila: " + i + " El campo BLOQUE se ecuentra nulo", RutaErroresDocumento);
        //                    validacion = false;
        //                }
        //                if (string.IsNullOrWhiteSpace(salon))
        //                {
        //                    EscribirLineaArchivo("Fila: " + i + " El campo NOMBRE_SALON se encuentra nulo", RutaErroresDocumento);
        //                    validacion = false;
        //                }
        //            }
        //            objWorkbook.Clear();
        //            objWorkbook.Close();
        //            fStream.Close();
        //            fStream.Dispose();
        //        }

        //        return validacion;

        //    }
        //    catch (Exception ex)
        //    {

        //        throw new Exception(ex.Message);
        //    }
        //}

        //private void button1_Click(object sender, EventArgs e)
        //{

        //    try
        //    {
        //        string HoraDigitalizada = DateTime.Now.Day.ToString() + DateTime.Now.Month + DateTime.Now.Year + DateTime.Now.Hour + DateTime.Now.Minute + DateTime.Now.Second + DateTime.Now.Millisecond;
        //        bool Proceso = true;
        //        //bool Proceso = Validaciones(HoraDigitalizada);
        //        //JPGFiles = GetJPGFiles(Digitalizadas.CarpetaSeleccionado);
        //        string RutaMetadataEncontrados = Path.Combine(Digitalizadas.CarpetaSeleccionado + "\\ImagenesDigitalizadas_" + HoraDigitalizada, "Metadata_" + HoraDigitalizada + ".txt");
        //        string RutaMetadataNoEncontrados = Path.Combine(Digitalizadas.CarpetaSeleccionado + "\\ImagenesDigitalizadas_" + HoraDigitalizada, "MetadataNoEcontrados_" + HoraDigitalizada + ".txt");
        //        GenerarArchivo(new List<string>(), RutaMetadataEncontrados, ";");
        //        GenerarArchivo(new List<string>(), RutaMetadataNoEncontrados, ";");
        //        EscribirLineaArchivo("RegistroUsuario|CodigoSitio|NombreDeLaImagen", RutaMetadataEncontrados);
        //        EscribirLineaArchivo("RegistroUsuario|CodigoSitio|NombreDeLaImagen|CodigoDeBarras|Folio", RutaMetadataNoEncontrados);

        //        string RutaErroresDocumento = Path.Combine(Digitalizadas.CarpetaSeleccionado, "ErroresDocumento_" + HoraDigitalizada + ".txt");

        //        ErroresDatosNull.Clear();
        //        //GenerarArchivo(new List<string>(), RutaErroresDocumento, ";");


        //        //if (JPGFiles.Count <= 0)
        //        //{
        //        //    throw new Exception($"En la ruta especificada no hay imagenes para renombrar");
        //        //}

        //        if (!GetJPGFile("*", Digitalizadas.CarpetaSeleccionado))
        //        {
        //            throw new Exception($"En la ruta especificada no hay imagenes para renombrar");
        //        }
        //        if (ArchivoCargado.ArchivoSeleccionado.Length <= 0)
        //        {
        //            throw new Exception($"Porfavor seleccione la ruta del documento");
        //        }

        //        if (Digitalizadas.CarpetaSeleccionado.Length <= 0)
        //        {
        //            throw new Exception($"Porfavor seleccione la ruta de digitalizacion");
        //        }

        //        if (Proceso)
        //        {

        //            ISheet objWorksheet;
        //            string strPath = ArchivoCargado.ArchivoSeleccionado;
        //            using (var fStream = new FileStream(strPath, FileMode.Open, FileAccess.Read))
        //            {
        //                fStream.Position = 0;
        //                XSSFWorkbook objWorkbook = new XSSFWorkbook(fStream);
        //                objWorksheet = objWorkbook.GetSheetAt(0);
        //                IRow objHeader = objWorksheet.GetRow(0);
        //                int countCells = objHeader.LastCellNum;
        //                if (countCells != 26)
        //                {
        //                    throw new Exception($"El documento no cumple con la estructuta, {Environment.NewLine} el archivo debe tener 25 columnas y se estan enviando {countCells}");
        //                }
        //                for (int i = (objWorksheet.FirstRowNum + 1); i <= objWorksheet.LastRowNum; i++)
        //                {
        //                    IRow objRow = objWorksheet.GetRow(i);

        //                    string numeroRegistro = objRow.GetCell(3).ToString().Replace("\n", " ").Replace("\t", string.Empty);
        //                    string sitio = objRow.GetCell(12).ToString().Replace("\n", " ").Replace("\t", string.Empty);
        //                    string codigoBarras = objRow.GetCell(24).ToString().Replace("\n", " ").Replace("\t", string.Empty);
        //                    string bloque = objRow.GetCell(15).ToString().Replace("\n", " ").Replace("\t", string.Empty);
        //                    string salon = objRow.GetCell(16).ToString().Replace("\n", " ").Replace("\t",string.Empty);

        //                    #region validaciones


        //                    if (string.IsNullOrWhiteSpace(numeroRegistro))
        //                    {
        //                        EscribirLineaArchivo("Fila: " + i + " El campo NUMERO_REGISTRO se encuentra nulo", RutaErroresDocumento);
        //                    }
        //                    if (string.IsNullOrWhiteSpace(sitio))
        //                    {
        //                        EscribirLineaArchivo("Fila: " + i + " El campo CONSECUTIVO_CIUDAD se encuentra nulo", RutaErroresDocumento);
        //                    }
        //                    if (string.IsNullOrWhiteSpace(codigoBarras))
        //                    {
        //                        EscribirLineaArchivo("Fila: " + i + " El campo CODIGO_BARRAS se encuentra nulo", RutaErroresDocumento);
        //                    }
        //                    if (string.IsNullOrWhiteSpace(bloque))
        //                    {
        //                        EscribirLineaArchivo("Fila: " + i + " El campo BLOQUE se ecuentra nulo", RutaErroresDocumento);
        //                    }
        //                    if (string.IsNullOrWhiteSpace(salon))
        //                    {
        //                        EscribirLineaArchivo("Fila: " + i + " El campo NOMBRE_SALON se encuentra nulo", RutaErroresDocumento);
        //                    }

        //                    #endregion

        //                    var SeparadorSitio = sitio.Split("-"[0]);
        //                    String CodSitio = SeparadorSitio[0];

        //                    if (GetJPGFile(codigoBarras, Digitalizadas.CarpetaSeleccionado))
        //                    {
        //                        String NuevoNombre = CodSitio + "_S" + salon + "_B" + bloque + "_F" + codigoBarras.Substring(13, 3);

        //                        if (!Directory.Exists(Path.Combine(Digitalizadas.CarpetaSeleccionado + "\\ImagenesDigitalizadas_" + HoraDigitalizada, sitio)))
        //                        {
        //                            // Si no existe, crearla
        //                            Directory.CreateDirectory(Path.Combine(Digitalizadas.CarpetaSeleccionado + "\\ImagenesDigitalizadas_" + HoraDigitalizada, sitio));
        //                        }
        //                        //JPGFiles.Remove(codigoBarras + ".jpg");

        //                        File.Copy(Digitalizadas.CarpetaSeleccionado + "\\" + codigoBarras + ".jpg", Digitalizadas.CarpetaSeleccionado + "\\ImagenesDigitalizadas_" + HoraDigitalizada + "\\" + sitio + "\\" + NuevoNombre + ".jpg", true);
        //                        EscribirLineaArchivo(numeroRegistro + "|" + sitio + "|" + NuevoNombre, RutaMetadataEncontrados);

        //                    }
        //                    else
        //                    {
        //                        String NoEncontrado = CodSitio + "_S" + salon + "_B" + bloque + "_F" + codigoBarras.Substring(13, 3);
        //                        EscribirLineaArchivo(numeroRegistro + "|" + sitio + "|" + NoEncontrado + "|" + codigoBarras + "|" + codigoBarras.Substring(13, 3), RutaMetadataNoEncontrados);
        //                    }
        //                }
        //                objWorkbook.Clear();
        //                objWorkbook.Close();
        //                fStream.Close();
        //                fStream.Dispose();
        //            }

        //            //string RutaFaltantes = Path.Combine(Renombradas.CarpetaSeleccionado, "HojasFaltantes.txt");




        //            contador = 0;
        //            //if (JPGFiles.Count != 0)
        //            //{
        //            //GenerarArchivo(JPGFiles, RutaFaltantes, ";");
        //            MessageBox.Show("Se Han renombrado las hojas con exito\n" +
        //            "En la  ruta: " + Digitalizadas.CarpetaSeleccionado, "Proceso Terminado", MessageBoxButtons.OK, MessageBoxIcon.Information);

        //            //}
        //            //else {
        //            //    MessageBox.Show("Se Han renombrado las hojas con exito\n", "Proceso Terminado", MessageBoxButtons.OK, MessageBoxIcon.Information);
        //            //}

        //            JPGFiles.Clear();
        //            MetadataEncontrados.Clear();
        //            MetadataNoEncontrados.Clear();
        //        }
        //        else
        //        {
        //            MessageBox.Show("La base de datos tiene inconsistencias\n" + " Verifica los errores en la siguiente ruta: " + Digitalizadas.CarpetaSeleccionado + "\\ImagenesDigitalizadas_" + HoraDigitalizada + "\\ErroresDocumento_" + HoraDigitalizada + ".txt", "Error de BD", MessageBoxButtons.OK, MessageBoxIcon.Error);
        //            JPGFiles.Clear();
        //            contador = 0;
        //            MetadataEncontrados.Clear();
        //            MetadataNoEncontrados.Clear();
        //        }

        //    }
        //    catch (Exception ex)
        //    {

        //        MessageBox.Show(ex.Message, "Error al renombrar los archivos", MessageBoxButtons.OK, MessageBoxIcon.Error);
        //    }

        //}


        private void button1_Click(object sender, EventArgs e)
        {

            try
            {
                Renombrar.Enabled = false;
                string HoraDigitalizada = DateTime.Now.Day.ToString() + DateTime.Now.Month + DateTime.Now.Year + DateTime.Now.Hour + DateTime.Now.Minute + DateTime.Now.Second + DateTime.Now.Millisecond;
                bool Proceso = true;
                //bool Proceso = Validaciones(HoraDigitalizada);
                JPGFiles = GetJPGFiles(Digitalizadas.CarpetaSeleccionado);
                string RutaMetadataEncontrados = Path.Combine(Digitalizadas.CarpetaSeleccionado + "\\ImagenesDigitalizadas_" + HoraDigitalizada, "Metadata_" + HoraDigitalizada + ".txt");
                string RutaMetadataNoEncontrados = Path.Combine(Digitalizadas.CarpetaSeleccionado + "\\ImagenesDigitalizadas_" + HoraDigitalizada, "MetadataNoEcontrados_" + HoraDigitalizada + ".txt");
                GenerarArchivo(new List<string>(), RutaMetadataEncontrados, ";");
                GenerarArchivo(new List<string>(), RutaMetadataNoEncontrados, ";");
                EscribirLineaArchivo("RegistroUsuario|CodigoSitio|NombreDeLaImagen", RutaMetadataEncontrados);
                EscribirLineaArchivo("RegistroUsuario|CodigoSitio|NombreDeLaImagen|CodigoDeBarras|Folio", RutaMetadataNoEncontrados);

                string RutaErroresDocumento = Path.Combine(Digitalizadas.CarpetaSeleccionado, "ErroresDocumento_" + HoraDigitalizada + ".txt");

                ErroresDatosNull.Clear();
                GenerarArchivo(new List<string>(), RutaErroresDocumento, ";");


                if (JPGFiles.Count <= 0)
                {
                    throw new Exception($"En la ruta especificada no hay imagenes para renombrar");
                }

                if (!GetJPGFile("*", Digitalizadas.CarpetaSeleccionado))
                {
                    throw new Exception($"En la ruta especificada no hay imagenes para renombrar");
                }
                if (ArchivoCargado.ArchivoSeleccionado.Length <= 0)
                {
                    throw new Exception($"Porfavor seleccione la ruta del documento");
                }

                if (Digitalizadas.CarpetaSeleccionado.Length <= 0)
                {
                    throw new Exception($"Porfavor seleccione la ruta de digitalizacion");
                }
                

                if (Proceso)
                {

                    Microsoft.Office.Interop.Excel.Application xl = new Microsoft.Office.Interop.Excel.Application();
                    Microsoft.Office.Interop.Excel.Workbook workbook = xl.Workbooks.Open(ArchivoCargado.ArchivoSeleccionado);
                    Microsoft.Office.Interop.Excel.Worksheet sheet = workbook.Sheets[1];


                    int numRows = sheet.UsedRange.Rows.Count;
                    int numColumns = sheet.UsedRange.Columns.Count;
                    /*int numColumns = 26;*/     // according to your sample

                    if (numColumns != 26)
                    {
                        throw new Exception($"El documento no tiene las 26 columnas requeridas por defecto, Verificar el documento cargado");
                    }

                    for (int rowIndex = 2; rowIndex <= numRows; rowIndex++)  // assuming the data starts at 1,1
                    {
                        string numeroRegistro = string.Empty;
                        string sitio = string.Empty;
                        string codigoBarras = string.Empty;
                        string bloque = string.Empty;
                        string salon = string.Empty;
                        string CodSitio = string.Empty;
                        string CodigoBarrasFolio = string.Empty;




                        Range cell = (Range)sheet.Cells[rowIndex, 4];
                        if (!string.IsNullOrWhiteSpace(cell.Value))
                        {
                            numeroRegistro = cell.Value.ToString().Replace("\n", " ").Replace("\t", string.Empty);
                        }
                        cell = (Range)sheet.Cells[rowIndex, 13];
                        if (cell.Value == null)
                        {
                            sitio = "";
                        }
                        else {
                            if (!string.IsNullOrWhiteSpace(cell.Value.ToString()))
                            {
                                var opcion = cell.Value.ToString().Replace("\n", " ").Replace("\t", string.Empty);
                                sitio = opcion.ToString();
                            }
                        }
                        
                        cell = (Range)sheet.Cells[rowIndex, 25];
                        if (!string.IsNullOrWhiteSpace(cell.Value))
                        {
                            codigoBarras = cell.Value.ToString().Replace("\n", " ").Replace("\t", string.Empty);
                        }
                        cell = (Range)sheet.Cells[rowIndex, 16];
                        if (!string.IsNullOrWhiteSpace(cell.Value))
                        {
                            bloque = cell.Value.ToString().Replace("\n", " ").Replace("\t", string.Empty);
                        }
                        cell = (Range)sheet.Cells[rowIndex, 17];
                        if (!string.IsNullOrWhiteSpace(cell.Value))
                        {
                            salon = cell.Value.ToString().Replace("\n", " ").Replace("\t", string.Empty);
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
                        if (!sitio.Contains('-'))
                        {
                            EscribirLineaArchivo("Fila: " + rowIndex + " El campo CONSECUTIVO_CIUDAD tiene una novedad en su estructura (No contiene el -)", RutaErroresDocumento);
                        }
                        if (string.IsNullOrWhiteSpace(codigoBarras))
                        {
                            EscribirLineaArchivo("Fila: " + rowIndex + " El campo CODIGO_BARRAS se encuentra nulo", RutaErroresDocumento);
                        }
                        if (codigoBarras.Length != 21)
                        {
                            EscribirLineaArchivo("Fila: " + rowIndex + " El campo CODIGO_BARRAS no cuenta con los 21 digitos obligatorios", RutaErroresDocumento);
                        }
                        if (string.IsNullOrWhiteSpace(bloque))
                        {
                            EscribirLineaArchivo("Fila: " + rowIndex + " El campo BLOQUE se ecuentra nulo", RutaErroresDocumento);
                        }
                        if (string.IsNullOrWhiteSpace(salon))
                        {
                            EscribirLineaArchivo("Fila: " + rowIndex + " El campo NOMBRE_SALON se encuentra nulo", RutaErroresDocumento);
                        }

                        //if (numeroRegistro.Length <= 0 || sitio.Length <= 0 || codigoBarras.Length <= 0 || bloque.Length <= 0 || salon.Length <= 0)
                        //{
                        //    throw new Exception($"El documento tiene incosistencias. Verificar en el documento de errores");


                        //}
                        //if (!sitio.Contains('-'))
                        //{
                        //    throw new Exception($"La columna de CONSECUTIVO_CIUDAD tiene una novedad en su estructura, Verificar en el documento de errores");
                        //}

                        //if (codigoBarras.Length != 21)
                        //{
                        //    throw new Exception($"La columna de CODIGO_BARRAAS tiene una novedad en su estructura, Verificar en el documento de errores");
                        //}

                        if (!sitio.Contains('-'))
                        {
                            CodSitio = sitio;
                        }
                        else {

                            var SeparadorSitio = sitio.Split("-"[0]);
                            CodSitio = SeparadorSitio[0];
                        }

                        if (codigoBarras.Length != 21)
                        {
                            CodigoBarrasFolio = "SinFolio";
                        }
                        else {

                            CodigoBarrasFolio = codigoBarras.Substring(13, 3);
                        }



                        #endregion




                        if (GetJPGFile(codigoBarras, Digitalizadas.CarpetaSeleccionado))
                        {
                            String NuevoNombre = CodSitio + "_S" + salon + "_B" + bloque + "_F" + CodigoBarrasFolio;

                            if (!Directory.Exists(Path.Combine(Digitalizadas.CarpetaSeleccionado + "\\ImagenesDigitalizadas_" + HoraDigitalizada, sitio)))
                            {
                                //Si no existe, crearla
                                Directory.CreateDirectory(Path.Combine(Digitalizadas.CarpetaSeleccionado + "\\ImagenesDigitalizadas_" + HoraDigitalizada, sitio));
                            }
                            JPGFiles.Remove(codigoBarras + ".jpg");

                            File.Copy(Digitalizadas.CarpetaSeleccionado + "\\" + codigoBarras + ".jpg", Digitalizadas.CarpetaSeleccionado + "\\ImagenesDigitalizadas_" + HoraDigitalizada + "\\" + sitio + "\\" + NuevoNombre + ".jpg", true);
                            EscribirLineaArchivo(numeroRegistro + "|" + sitio + "|" + NuevoNombre, RutaMetadataEncontrados);

                        }
                        else
                        {
                            String NoEncontrado = CodSitio + "_S" + salon + "_B" + bloque + "_F" + CodigoBarrasFolio;
                            EscribirLineaArchivo(numeroRegistro + "|" + sitio + "|" + NoEncontrado + "|" + codigoBarras + "|" + CodigoBarrasFolio, RutaMetadataNoEncontrados);
                        }


                    }
                    xl.Quit();
                 




                    contador = 0;
                    JPGFiles.Clear();
                    Renombrar.Enabled = true;
                    MetadataEncontrados.Clear();
                    MetadataNoEncontrados.Clear();
                    Renombrar.Enabled = true;

                    FileInfo fileinfo = new FileInfo(RutaErroresDocumento);
                    if (fileinfo.Length < 1)
                    {
                        File.Delete(RutaErroresDocumento);
                    }
                    else {

                        File.Delete(RutaMetadataEncontrados);
                        File.Delete(RutaMetadataNoEncontrados);
                        Directory.Delete(Digitalizadas.CarpetaSeleccionado + "\\ImagenesDigitalizadas_" + HoraDigitalizada,true);
                    }

                    if (fileinfo.Length < 1)
                    {
                        MessageBox.Show("Se han renombrado las hojas con exito\n", "Proceso Terminado", MessageBoxButtons.OK, MessageBoxIcon.Information);

                    }
                    else
                    {

                        MessageBox.Show("El proceso ha terminado, pero existen algunos errores.\n" +
                            "por favor verificar con el documento de errores encontrado en la siguiente ruta: \n" +
                            Digitalizadas.CarpetaSeleccionado, "Proceso Terminado", MessageBoxButtons.OK, MessageBoxIcon.Information);

                    }







                }
                else
                {
                    MessageBox.Show("La base de datos tiene inconsistencias\n" + " Verifica los errores en la siguiente ruta: " + Digitalizadas.CarpetaSeleccionado + "\\ImagenesDigitalizadas_" + HoraDigitalizada + "\\ErroresDocumento_" + HoraDigitalizada + ".txt", "Error de BD", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    JPGFiles.Clear();
                    contador = 0;
                    MetadataEncontrados.Clear();
                    MetadataNoEncontrados.Clear();
                }

            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message, "Error al renombrar los archivos", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

        //private void button1_Click(object sender, EventArgs e)
        //{

        //    try
        //    {
        //        string HoraDigitalizada = DateTime.Now.Day.ToString() + DateTime.Now.Month + DateTime.Now.Year + DateTime.Now.Hour + DateTime.Now.Minute + DateTime.Now.Second + DateTime.Now.Millisecond;

        //        JPGFiles = GetJPGFiles(Digitalizadas.CarpetaSeleccionado);


        //        if (!GetJPGFile("*", Digitalizadas.CarpetaSeleccionado))
        //        {
        //            throw new Exception($"En la ruta especificada no hay imagenes para renombrar");
        //        }
        //        if (ArchivoCargado.ArchivoSeleccionado.Length <= 0)
        //        {
        //            throw new Exception($"Porfavor seleccione la ruta del documento");
        //        }

        //        if (Digitalizadas.CarpetaSeleccionado.Length <= 0)
        //        {
        //            throw new Exception($"Porfavor seleccione la ruta de digitalizacion");
        //        }

        //        Microsoft.Office.Interop.Excel.Application xl = new Microsoft.Office.Interop.Excel.Application();
        //        Microsoft.Office.Interop.Excel.Workbook workbook = xl.Workbooks.Open(ArchivoCargado.ArchivoSeleccionado);
        //        Microsoft.Office.Interop.Excel.Worksheet sheet = workbook.Sheets[1];


        //        int numRows = sheet.UsedRange.Rows.Count;

        //        xl.Quit();


        //        int hilos = 100;

        //        int aproximado = numRows / hilos;

        //        Task[] tasks = new Task[10];

        //        int inicial = 2;

        //        List<Task> tareas = new List<Task>();

        //        for (int i = 0; i < hilos; i++)
        //        {
        //            int final = inicial + aproximado;
        //            if (i == hilos - 1)
        //            {
        //                final = numRows;
        //            }

        //            int inicialNuevo = new int();
        //            inicialNuevo = int.Parse(inicial.ToString());
        //            int indicadorNuevo = new int();
        //            indicadorNuevo = int.Parse(i.ToString());
        //            var task = Task.Run(async () => await new ProcessAsync(ArchivoCargado.ArchivoSeleccionado, Digitalizadas.CarpetaSeleccionado, inicialNuevo, int.Parse(final.ToString()), indicadorNuevo).Procesar());
        //            tareas.Add(task);

        //            inicial = inicial + aproximado + 1;
        //        }

        //        Task.WaitAll(tareas.ToArray());

        //    }
        //    catch (Exception ex)
        //    {

        //        MessageBox.Show(ex.Message, "Error al renombrar los archivos", MessageBoxButtons.OK, MessageBoxIcon.Error);
        //    }

        //}

        private void ArchivoCargado_OnSeleccionArchivo(object sender, WinFormsControlLibrary.usrSeleccionarArchivo.SeleccionArchivoEventArgs e)
        {


            bool abierto = Functions.FileHelper.IsFileLocked(ArchivoCargado.ArchivoSeleccionado);
            if (abierto)
            {
                MessageBox.Show("El documento se encuentra abierto, debes cerrarlo para realizar el proceso.\n" +
                    "Porfavor selecciona de nuevo el documento", "Documento Abierto", MessageBoxButtons.OK, MessageBoxIcon.Error);
                ArchivoCargado.Limpiar();
                Renombrar.Enabled = false;
            }
        }

        private void Digitalizadas_OnSeleccionCarpeta(object sender, WinFormsControlLibrary.usrSeleccionarCarpeta.SeleccionCarpetaEventArgs e)
        {

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
    }
}
