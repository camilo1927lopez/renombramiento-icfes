using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RenombramientoIcfes.Helper
{
    public class ExcelHelper
    {
        public DataSet LoadToDataSet(string archivo, bool tieneEncabezado = true)
        {
            DataSet ds = new DataSet();
           
            try
            {
                using (var pck = new OfficeOpenXml.ExcelPackage())
                {
                    using (var stream = File.OpenRead(archivo))
                    {
                        pck.Load(stream);
                    }


                    foreach (ExcelWorksheet ws in pck.Workbook.Worksheets)
                    {
                        DataTable tbl = new DataTable();
                        if (ws.Dimension == null)
                        {
                            throw new Exception(string.Format("Contiene hojas vacías", ws.Name));
                        }
                        int totalColumns = ws.Dimension.End.Column;

                        for (int i = 1; i <= totalColumns; i++)
                        {
                            tbl.Columns.Add(tieneEncabezado ? ws.Cells[1, i].Text : string.Format("Column {0}", i.ToString()));
                        }

                        var startRow = tieneEncabezado ? 2 : 1;
                        for (int rowNum = startRow; rowNum <= ws.Dimension.End.Row; rowNum++)
                        {
                            var wsRow = ws.Cells[rowNum, 1, rowNum, totalColumns];
                            DataRow row = tbl.Rows.Add();
                            foreach (var cell in wsRow)
                            {
                                row[cell.Start.Column - 1] = cell.Text;
                            }
                        }
                        tbl.TableName = ws.Name;
                        ds.Tables.Add(tbl);
                    }

  
                }

               
            }
            catch (Exception ex)
            {
               
                string appName = typeof(ExcelHelper).Name;
                throw new Exception(string.Format("{0}, Error cargando el archivo de excel '{1}', {2}", appName, archivo, Environment.NewLine + ex.Message));
            }
            return ds;
        }
    }
}
