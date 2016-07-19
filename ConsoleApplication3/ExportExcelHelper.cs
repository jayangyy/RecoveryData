using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Web;
using System.Linq;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.Util;

namespace 网络考试V2
{
    public class ExportExcelHelper
    {
        public enum NpoiCellFormatType
        {
            //百分比
            PerType=0
        }
        /// <summary>
        /// 导出Excel
        /// </summary>
        /// <param name="dataTable">导出数据</param>
        /// <param name="exportType">导出类型:0 xlsx 其它:xls</param>
        /// <param name="list">合并单元格数据</param>
        /// <returns></returns>
        public static Tuple<string, byte[]> GetMemoryStream(DataTable dataTable, int exportType, List<NPOI.SS.Util.CellRangeAddress> list = null,
            IEnumerable<string> formatDic =null, IEnumerable<string> ignoredColumns = null, string filter = null)
        {
            using (dataTable)
            {
                IWorkbook workbook = null;
                if (exportType == 0)
                    workbook = new XSSFWorkbook();
                else
                    workbook = new HSSFWorkbook();

                ISheet sheet = workbook.CreateSheet();
                //创建表头
                IRow headerRow = sheet.CreateRow(0);
                int index = 0;
                foreach (DataColumn column in dataTable.Columns)
                {
                    if (ignoredColumns != null && ignoredColumns.Contains(column.Caption)) continue;
                    headerRow.CreateCell(index++).SetCellValue(column.Caption);
                }
                //创建内容数据行
                var rows = string.IsNullOrWhiteSpace(filter) ? dataTable.Select() : dataTable.Select(filter);
                for (int i = 0; i < rows.Length; i++)
                {
                    DataRow dataRow = rows[i];
                    IRow row = sheet.CreateRow(i + 1);
                    index = 0;
                    foreach (DataColumn column in dataTable.Columns)
                    {
                        var columnName = column.Caption;
                        if (ignoredColumns != null && ignoredColumns.Contains(columnName)) continue;
                        var isNumeric = column.DataType.IsValueType;
                        var cellType = isNumeric ? CellType.Numeric : CellType.String;
                        var cell = row.CreateCell(index++, cellType);
                        
                        cell.CellStyle.VerticalAlignment = VerticalAlignment.Top;
                        cell.CellStyle.Alignment = HorizontalAlignment.Left;
                        #region 设置单元格格式
                        if (formatDic!=null)
                        {
                            if(formatDic.Contains(columnName))
                            {
                                ICellStyle cellStyle = workbook.CreateCellStyle();
                                //IDataFormat format = workbook.CreateDataFormat();
                                //cellStyle.DataFormat = HSSFDataFormat.GetBuiltinFormat("0.00%"); //单元格格式为“0.00”来表示,"￥#,##0"美元显示,"0.00%"百分比显示  
                                //cell.CellStyle = cellStyle;
                                //HSSFCellStyle cellStyle = workbook.CreateCellStyle() as HSSFCellStyle;
                                //HSSFDataFormat format = workbook.CreateDataFormat()  as HSSFDataFormat;
                                //cellStyle.DataFormat = format.GetFormat("0.00%");
                                HSSFFont ffont = (HSSFFont)workbook.CreateFont();
                                //ffont.FontHeight = 12 * 12;
                                ffont.FontName = "宋体";
                                ffont.Color = HSSFColor.Black.Index;
                                cellStyle.SetFont(ffont);
                                cellStyle.DataFormat = (short)10;
                                cell.CellStyle = cellStyle;
                                if(i==1000)
                                {
                                    string s = "";
                                }
                                
                            }
                        }
                        #endregion
                        if (isNumeric)
                        {
                            if(dataRow[column]==null||string.IsNullOrWhiteSpace(dataRow[column].ToString()))
                            {
                                cell.SetCellValue(dataRow[column].ToString());
                            }
                            else
                            {
                               
                                if (column.DataType.ToString() == "System.DateTime") {
                                    cell.SetCellType(CellType.String);
                                    cell.SetCellValue(dataRow[column].ToString());
                                } else {
                                    cell.SetCellValue(Convert.ToDouble(dataRow[column])); }
                            }
                        }
                        else
                        {
                            cell.SetCellValue(dataRow[column].ToString());
                        }
                    }
                }
                if (list != null && list.Count > 0)
                {
                    list.ForEach(t => sheet.AddMergedRegion(t));
                }
                using (MemoryStream ms = new MemoryStream())
                {
                    workbook.Write(ms);
                    ms.Flush();
                    return Tuple.Create<string, byte[]>(exportType == 0 ? "xlsx" : "xls", ms.ToArray());
                }

            }
        }
        /// <summary>
        /// 导出Excel
        /// </summary>
        /// <param name="dataTable">导出数据</param>
        /// <param name="exportType">导出类型:0 xlsx 其它:xls</param>
        /// <param name="list">合并单元格数据</param>
        /// <returns></returns>
        public static Tuple<string, byte[]> GetMemoryStream<T>(List<T> dataTable, int exportType, List<NPOI.SS.Util.CellRangeAddress> list = null, bool careteCaption = false)
        {
            return new Tuple<string, byte[]>("r", new byte[] { });

          
                //IWorkbook workbook = null;
                //if (exportType == 0)
                //    workbook = new XSSFWorkbook();
                //else
                //    workbook = new HSSFWorkbook();

                //ISheet sheet = workbook.CreateSheet();
                ////创建表头
                //IRow headerRow = sheet.CreateRow(0);
                //foreach (DataColumn column in dataTable.Columns)
                //{
                //    headerRow.CreateCell(column.Ordinal).SetCellValue(column.Caption);
                //}
                ////创建内容数据行
                //for (int i = 0; i < dataTable.Rows.Count; i++)
                //{
                //    DataRow dataRow = dataTable.Rows[i];
                //    IRow row = sheet.CreateRow(i + 1);

                //    foreach (DataColumn column in dataTable.Columns)
                //    {
                //        var cell = row.CreateCell(column.Ordinal);
                //        cell.CellStyle.VerticalAlignment = VerticalAlignment.Top;
                //        cell.CellStyle.Alignment = HorizontalAlignment.Left;
                //        cell.SetCellValue(dataRow[column].ToString());
                //    }
                //}
                //if (list != null && list.Count > 0)
                //{
                //    list.ForEach(t => sheet.AddMergedRegion(t));
                //}
                //using (MemoryStream ms = new MemoryStream())
                //{
                //    workbook.Write(ms);
                //    ms.Flush();
                //    return Tuple.Create<string, byte[]>(exportType == 0 ? "xlsx" : "xls", ms.ToArray());
                //}
        }
        /// <summary>
        /// 导出Excel
        /// </summary>
        /// <param name="dataTable">导出数据</param>
        /// <param name="exportType">导出类型:0 xlsx 其它:xls</param>
        /// <param name="list">合并单元格数据</param>
        /// <returns></returns>
        public static Tuple<string, byte[]> GetMemoryStream<T>(List<T> dataTable, int exportType, List<NPOI.SS.Util.CellRangeAddress> list = null, string[] captions = null)
        {
            return new Tuple<string, byte[]>("r", new byte[] { });
            //using (dataTable)
            //{
            //    IWorkbook workbook = null;
            //    if (exportType == 0)
            //        workbook = new XSSFWorkbook();
            //    else
            //        workbook = new HSSFWorkbook();

            //    ISheet sheet = workbook.CreateSheet();
            //    //创建表头
            //    IRow headerRow = sheet.CreateRow(0);
            //    foreach (DataColumn column in dataTable.Columns)
            //    {
            //        headerRow.CreateCell(column.Ordinal).SetCellValue(column.Caption);
            //    }
            //    //创建内容数据行
            //    for (int i = 0; i < dataTable.Rows.Count; i++)
            //    {
            //        DataRow dataRow = dataTable.Rows[i];
            //        IRow row = sheet.CreateRow(i + 1);

            //        foreach (DataColumn column in dataTable.Columns)
            //        {
            //            var cell = row.CreateCell(column.Ordinal);
            //            cell.CellStyle.VerticalAlignment = VerticalAlignment.Top;
            //            cell.CellStyle.Alignment = HorizontalAlignment.Left;
            //            cell.SetCellValue(dataRow[column].ToString());
            //        }
            //    }
            //    if (list != null && list.Count > 0)
            //    {
            //        list.ForEach(t => sheet.AddMergedRegion(t));
            //    }
            //    using (MemoryStream ms = new MemoryStream())
            //    {
            //        workbook.Write(ms);
            //        ms.Flush();
            //        return Tuple.Create<string, byte[]>(exportType == 0 ? "xlsx" : "xls", ms.ToArray());
            //    }

            //}


        }
        public void ImportExtend()
        {



            //HSSFWorkbook wk = new HSSFWorkbook(file.InputStream);
            //ISheet sheet = wk.GetSheetAt(0);
            //List<Merchandise> list = new List<Merchandise>();

            //for (int i = 1; i <= sheet.LastRowNum; i++)
            //{
            //    rowNum = i;
            //    try
            //    {
            //        IRow row = sheet.GetRow(i);
            //        if (row != null)
            //        {
            //            var merchandise = new Merchandise
            //            {
            //                Id = Convert.ToInt32(row.GetCell(0).ToString()),
            //                Title = row.GetCell(2).ToString(),
            //                Alias = row.GetCell(3).ToString(),
            //                MerchandiseInfo = new Product
            //                {
            //                    Type = new TypeProperty
            //                    {
            //                        TypeName = row.GetCell(5).ToString()
            //                    }
            //                },
            //                Price = Convert.ToDecimal(row.GetCell(11).ToString()),
            //                InventoryQuantity = Convert.ToInt32(row.GetCell(12).ToString())
            //            };
            //            list.Add(merchandise);
            //        }
            //    }

        }

        //public static void ExportExcel(HttpResponse Response, string filename = "", Tuple<string, byte[]> excel = null)
        //{
        //    Response.AddHeader("Content-Disposition", $"attachment; filename={filename}{DateTime.Now.ToString("yyyyMMddHHmmssfff")}.{excel.Item1}");
        //    Response.ContentType = "application/excel";
        //    Response.BinaryWrite(excel.Item2);
        //    Response.Flush();
        //    Response.Close();
        //    Response.End();
        //}

        public static string ExportGridToExcel(DataTable grid, String fileName, List<NPOI.SS.Util.CellRangeAddress> list = null,
            IEnumerable<string> formatDic = null, IEnumerable<string> ignoredColumns = null, string filter = null)
        {
            //try
            //{
            //    var ms = ExportExcelHelper.GetMemoryStream(grid, 1, list, formatDic, ignoredColumns, filter);
            //    ExportExcelHelper.ExportExcel(HttpContext.Current.Response, fileName, ms);
            //    return null;
            //}
            //catch (Exception ex)
            //{
            //    return ex.Message;
            //}
            return "";
        }

    }
}
