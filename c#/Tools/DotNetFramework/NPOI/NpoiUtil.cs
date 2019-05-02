using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Data;
using System.IO;

namespace DotNetFramework.NPOI
{
    public class NpoiUtil
    {
        public static void ModifyExcelDemo(string excelPath)
        {
            IWorkbook workbook = null;
            string excelType = Path.GetExtension(excelPath);
            using (FileStream stream = new FileStream(excelPath, FileMode.Open, FileAccess.Read))
            {
                switch (excelType.ToLower())
                {
                    case ".xls":
                        workbook = new HSSFWorkbook(stream);
                        break;
                    case ".xlsx":
                        workbook = new XSSFWorkbook(stream);
                        break;
                }
            }

            ISheet sheet = workbook.GetSheetAt(0);
            for (int i = 0; i < 4; i++)
            {
                IRow row = sheet.GetRow(i);
                for (int j = 0; j < 3; j++)
                {
                    ICell cell = row.GetCell(j);
                    cell.SetCellValue(i * j);
                }
            }

            var savePath = excelPath;
            if (excelType.ToLower() == ".xlsx")  //xlsx格式直接写不知为什么没有生效，只好另存为-Copy
            {
                savePath = savePath.Replace(".xlsx", "-Copy.xlsx");
            }
            using (FileStream stream = new FileStream(savePath, FileMode.OpenOrCreate, FileAccess.Write))
            {
                workbook.Write(stream);
                stream.Close();
            }
        }

        public static DataTable ExcelToDataTable(string excelPath)
        {
            if (!File.Exists(excelPath))
            {
                throw new Exception(string.Format("\"{0}\" doesn't exist", excelPath));
            }

            DataTable dataTable = null;
            try
            {
                IWorkbook workbook = null;
                ISheet worksheet = null;
                string first_sheet_name = "";

                using (FileStream fs = new FileStream(excelPath, FileMode.Open, FileAccess.Read))
                {
                    workbook = WorkbookFactory.Create(fs);
                    worksheet = workbook.GetSheetAt(0);
                    first_sheet_name = worksheet.SheetName;

                    dataTable = new DataTable(first_sheet_name);

                    for (int rowIndex = 0; rowIndex <= worksheet.LastRowNum; rowIndex++)
                    {
                        DataRow dataRow = null;
                        IRow row = worksheet.GetRow(rowIndex);
                        IRow row2 = null;
                        IRow row3 = null;

                        if (rowIndex == 0)
                        {
                            row2 = worksheet.GetRow(rowIndex + 1);
                            row3 = worksheet.GetRow(rowIndex + 2);
                        }

                        if (row != null) //null is when the row only contains empty cells 
                        {
                            if (rowIndex > 0) dataRow = dataTable.NewRow();

                            int colIndex = 0;
                            //Leer cada Columna de la fila
                            foreach (ICell cell in row.Cells)
                            {
                                object valorCell = null;
                                string cellType = "";
                                string[] cellType2 = new string[2];

                                if (rowIndex == 0) //Asumo que la primera fila contiene los titlos:
                                {
                                    for (int i = 0; i < 2; i++)
                                    {
                                        ICell cell2 = null;
                                        if (i == 0) { cell2 = row2.GetCell(cell.ColumnIndex); }
                                        else { cell2 = row3.GetCell(cell.ColumnIndex); }

                                        if (cell2 != null)
                                        {
                                            switch (cell2.CellType)
                                            {
                                                case CellType.Blank: break;
                                                case CellType.Boolean: cellType2[i] = "System.Boolean"; break;
                                                case CellType.String: cellType2[i] = "System.String"; break;
                                                case CellType.Numeric:
                                                    if (HSSFDateUtil.IsCellDateFormatted(cell2)) { cellType2[i] = "System.DateTime"; }
                                                    else
                                                    {
                                                        cellType2[i] = "System.Double";  //valorCell = cell2.NumericCellValue;
                                                    }
                                                    break;

                                                case CellType.Formula:
                                                    bool continuar = true;
                                                    switch (cell2.CachedFormulaResultType)
                                                    {
                                                        case CellType.Boolean: cellType2[i] = "System.Boolean"; break;
                                                        case CellType.String: cellType2[i] = "System.String"; break;
                                                        case CellType.Numeric:
                                                            if (HSSFDateUtil.IsCellDateFormatted(cell2)) { cellType2[i] = "System.DateTime"; }
                                                            else
                                                            {
                                                                try
                                                                {
                                                                    //DETERMINAR SI ES BOOLEANO
                                                                    if (cell2.CellFormula == "TRUE()") { cellType2[i] = "System.Boolean"; continuar = false; }
                                                                    if (continuar && cell2.CellFormula == "FALSE()") { cellType2[i] = "System.Boolean"; continuar = false; }
                                                                    if (continuar) { cellType2[i] = "System.Double"; continuar = false; }
                                                                }
                                                                catch { }
                                                            }
                                                            break;
                                                    }
                                                    break;
                                                default:
                                                    cellType2[i] = "System.String"; break;
                                            }
                                        }
                                    }

                                    //Resolver las diferencias de Tipos
                                    if (cellType2[0] == cellType2[1]) { cellType = cellType2[0]; }
                                    else
                                    {
                                        if (cellType2[0] == null) cellType = cellType2[1];
                                        if (cellType2[1] == null) cellType = cellType2[0];
                                        if (cellType == "") cellType = "System.String";
                                    }

                                    //Obtener el nombre de la Columna
                                    string colName = "Column_{0}";
                                    try { colName = cell.StringCellValue; }
                                    catch { colName = string.Format(colName, colIndex); }

                                    //Verificar que NO se repita el Nombre de la Columna
                                    foreach (DataColumn col in dataTable.Columns)
                                    {
                                        if (col.ColumnName == colName) colName = string.Format("{0}_{1}", colName, colIndex);
                                    }

                                    //Agregar el campos de la tabla:
                                    DataColumn codigo = new DataColumn(colName, System.Type.GetType(cellType));
                                    dataTable.Columns.Add(codigo); colIndex++;
                                }
                                else
                                {
                                    //Las demas filas son registros:
                                    switch (cell.CellType)
                                    {
                                        case CellType.Blank: valorCell = DBNull.Value; break;
                                        case CellType.Boolean: valorCell = cell.BooleanCellValue; break;
                                        case CellType.String: valorCell = cell.StringCellValue; break;
                                        case CellType.Numeric:
                                            if (HSSFDateUtil.IsCellDateFormatted(cell)) { valorCell = cell.DateCellValue; }
                                            else { valorCell = cell.NumericCellValue; }
                                            break;
                                        case CellType.Formula:
                                            switch (cell.CachedFormulaResultType)
                                            {
                                                case CellType.Blank: valorCell = DBNull.Value; break;
                                                case CellType.String: valorCell = cell.StringCellValue; break;
                                                case CellType.Boolean: valorCell = cell.BooleanCellValue; break;
                                                case CellType.Numeric:
                                                    if (HSSFDateUtil.IsCellDateFormatted(cell)) { valorCell = cell.DateCellValue; }
                                                    else { valorCell = cell.NumericCellValue; }
                                                    break;
                                            }
                                            break;
                                        default: valorCell = cell.StringCellValue; break;
                                    }
                                    //Agregar el nuevo Registro
                                    if (cell.ColumnIndex <= dataTable.Columns.Count - 1) dataRow[cell.ColumnIndex] = valorCell;
                                }
                            }
                        }
                        if (rowIndex > 0) dataTable.Rows.Add(dataRow);
                    }
                    dataTable.AcceptChanges();
                }

            }
            catch (Exception ex)
            {
                throw ex;
            }
            return dataTable;
        }

        /// <summary>
        /// 未解决：
        /// xls，创建完打开会报“文件错误，可能某些数字格式已丢失”，点击“确定”可打开
        /// xlsx,创建完打开会报“...发现不可读取的内容。是否恢复此工作簿的内容？...”点击“是”修复后可打开
        /// 打开保存一次后不再报错
        /// </summary>
        /// <param name="dtSource"></param>
        /// <param name="excelPath"></param>
        public static void CreateExcelByDataTable(DataTable dtSource, string excelPath)
        {
            if (dtSource == null || dtSource.Rows.Count <= 0)
            {
                throw new ArgumentException("dtSource cannot be empty");
            }

            try
            {
                IWorkbook workbook = null;
                ISheet worksheet = null;

                using (FileStream stream = new FileStream(excelPath, FileMode.Create, FileAccess.ReadWrite))
                {
                    string excelType = Path.GetExtension(excelPath);
                    switch (excelType.ToLower())
                    {
                        case ".xls":
                            workbook = new HSSFWorkbook();
                            break;
                        case ".xlsx":
                            workbook = new XSSFWorkbook();
                            break;
                    }

                    worksheet = workbook.CreateSheet();


                    int rowCount = 0;
                    //初始化首行表头
                    if (dtSource.Columns.Count > 0)
                    {
                        int iCol = 0;
                        IRow excelHeaderRow = worksheet.CreateRow(rowCount);
                        foreach (DataColumn dataColumn in dtSource.Columns)
                        {
                            ICell cell = excelHeaderRow.CreateCell(iCol, CellType.String);
                            cell.SetCellValue(dataColumn.ColumnName);
                            iCol++;
                        }
                        rowCount++;
                    }


                    ICellStyle _doubleCellStyle = workbook.CreateCellStyle();
                    _doubleCellStyle.DataFormat = workbook.CreateDataFormat().GetFormat("#,##0.###");

                    ICellStyle _intCellStyle = workbook.CreateCellStyle();
                    _intCellStyle.DataFormat = workbook.CreateDataFormat().GetFormat("#,##0");

                    ICellStyle _boolCellStyle = workbook.CreateCellStyle();
                    _boolCellStyle.DataFormat = workbook.CreateDataFormat().GetFormat("BOOLEAN");

                    ICellStyle _dateCellStyle = workbook.CreateCellStyle();
                    _dateCellStyle.DataFormat = workbook.CreateDataFormat().GetFormat("dd-MM-yyyy");

                    ICellStyle _dateTimeCellStyle = workbook.CreateCellStyle();
                    _dateTimeCellStyle.DataFormat = workbook.CreateDataFormat().GetFormat("dd-MM-yyyy HH:mm:ss");

                    //将数据写入excel
                    foreach (DataRow dataRow in dtSource.Rows)
                    {
                        IRow excelRow = worksheet.CreateRow(rowCount);
                        int colCount = 0;
                        foreach (DataColumn dataColumn in dtSource.Columns)
                        {
                            ICell cell = null;
                            object cellValue = dataRow[colCount];
                            if (cellValue != DBNull.Value)
                            {
                                switch (dataColumn.DataType.ToString())
                                {
                                    case "System.Boolean":
                                        cell = excelRow.CreateCell(colCount, CellType.Boolean);
                                        if (Convert.ToBoolean(cellValue))
                                        {
                                            cell.SetCellFormula("TRUE()");
                                        }
                                        else
                                        {
                                            cell.SetCellFormula("FALSE()");
                                        }
                                        cell.CellStyle = _boolCellStyle;
                                        break;
                                    case "System.String":
                                        cell = excelRow.CreateCell(colCount, CellType.String);
                                        cell.SetCellValue(Convert.ToString(cellValue));
                                        break;
                                    case "System.Int32":
                                        cell = excelRow.CreateCell(colCount, CellType.Numeric);
                                        cell.SetCellValue(Convert.ToInt32(cellValue));
                                        cell.CellStyle = _intCellStyle;
                                        break;
                                    case "System.Int64":
                                        cell = excelRow.CreateCell(colCount, CellType.Numeric);
                                        cell.SetCellValue(Convert.ToInt64(cellValue));
                                        cell.CellStyle = _intCellStyle;
                                        break;
                                    case "System.Decimal":
                                        cell = excelRow.CreateCell(colCount, CellType.Numeric);
                                        cell.SetCellValue(Convert.ToDouble(cellValue));
                                        cell.CellStyle = _doubleCellStyle;
                                        break;
                                    case "System.Double":
                                        cell = excelRow.CreateCell(colCount, CellType.Numeric);
                                        cell.SetCellValue(Convert.ToDouble(cellValue));
                                        cell.CellStyle = _doubleCellStyle;
                                        break;

                                    case "System.DateTime":
                                        cell = excelRow.CreateCell(colCount, CellType.Numeric);
                                        cell.SetCellValue(Convert.ToDateTime(cellValue));
                                        DateTime cDate = Convert.ToDateTime(cellValue);
                                        if (cDate != null && cDate.Hour > 0)
                                        {
                                            cell.CellStyle = _dateTimeCellStyle;
                                        }
                                        else
                                        {
                                            cell.CellStyle = _dateCellStyle;
                                        }
                                        break;
                                    default:
                                        break;
                                }
                            }

                            colCount++;
                        }
                        rowCount++;
                    }

                    workbook.Write(stream);
                    stream.Close();
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

    }
}
