using System;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using MSExcel = Microsoft.Office.Interop.Excel;

namespace ExcelLib
{
    public static class Excel
    {
        public static EData1And2[] EData1 { get; set; }
        public static EData3[] EData3 { get; set; }
        public static EData4[] EData4 { get; set; }
        //Преобразование из XLS к DataTable
        private static DataTable ParseTable(string path)
        {
            DataTable Tabla = null;
            try
            {
                if (System.IO.File.Exists(path))
                {

                    IWorkbook workbook = null;  //IWorkbook determina si es xls o xlsx              
                    ISheet worksheet = null;
                    string first_sheet_name = "";

                    using (FileStream FS = new FileStream(path, FileMode.Open, FileAccess.Read))
                    {
                        workbook = WorkbookFactory.Create(FS);          //Abre tanto XLS como XLSX
                        worksheet = workbook.GetSheetAt(0);    //Obtener Hoja por indice
                        first_sheet_name = worksheet.SheetName;         //Obtener el nombre de la Hoja

                        Tabla = new DataTable(first_sheet_name);
                        Tabla.Rows.Clear();
                        Tabla.Columns.Clear();

                        // Leer Fila por fila desde la primera
                        for (int rowIndex = 0; rowIndex <= worksheet.LastRowNum; rowIndex++)
                        {
                            DataRow NewReg = null;
                            IRow row = worksheet.GetRow(rowIndex);
                            IRow row2 = null;
                            IRow row3 = null;

                            if (rowIndex == 0)
                            {
                                row2 = worksheet.GetRow(rowIndex + 1); //Si es la Primera fila, obtengo tambien la segunda para saber el tipo de datos
                                row3 = worksheet.GetRow(rowIndex + 2); //Y la tercera tambien por las dudas
                            }

                            if (row != null) //null is when the row only contains empty cells 
                            {
                                if (rowIndex > 0) NewReg = Tabla.NewRow();

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
                                        foreach (DataColumn col in Tabla.Columns)
                                        {
                                            if (col.ColumnName == colName) colName = string.Format("{0}_{1}", colName, colIndex);
                                        }

                                        //Agregar el campos de la tabla:
                                        DataColumn codigo = new DataColumn();//colName, System.Type.GetType(cellType));
                                        Tabla.Columns.Add(codigo); colIndex++;
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
                                        if (cell.ColumnIndex <= Tabla.Columns.Count - 1) NewReg[cell.ColumnIndex] = valorCell;
                                    }
                                }
                            }
                            if (rowIndex > 0) Tabla.Rows.Add(NewReg);
                        }
                        Tabla.AcceptChanges();
                    }
                }
                else
                {
                    throw new Exception("ERROR 404: El archivo especificado NO existe.");
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return Tabla;
        }
        //Парсит Excel1 и Excel2
        public static EData1And2[] ParseEx(string path)
        {
            var table = ParseTable(path);

            string[,] list = new string[table.Columns.Count, table.Rows.Count];
            for (int i = 0; i < table.Columns.Count; i++)
            {
                for (int j = 0; j < table.Rows.Count; j++)
                {
                    list[i, j] = table.Rows[j][i].ToString();
                }
            }

            int lenght = 0;
            for (int j = 0; j < table.Rows.Count; j++)
            {
                var parsedValue = 0;
                if (list[0, j] != String.Empty && int.TryParse(list[0, j], out parsedValue) && parsedValue > lenght)
                {
                    lenght = parsedValue;
                }
            }


            EData1 = new EData1And2[lenght];
            for (int i = 0; i < EData1.Length; i++)
            {
                EData1[i] = new EData1And2();
            }

            var nPossition = 0;
            var k = 0;
            string[] tabHeader = new string[table.Columns.Count];
            for (int j = 0; j < table.Rows.Count; j++)
            {
                if (list[0, j] != "№")
                {
                    var value = 0;
                    if (int.TryParse(list[0, j], out value))
                    {
                        EData1[k].Index = value;

                        for (int i = 1; i < table.Columns.Count; i++)
                        {
                            if (list[i, j] != String.Empty)
                            {
                                if (i == 1)
                                {
                                    string[] chandev = list[i, j].Split('/');
                                    string[] chandev2 = chandev[1].Split('R');
                                    string[] chandev3 = chandev2[0].Split('K');
                                    EData1[k].Input.Channel = int.Parse(chandev3[1]);
                                    EData1[k].Input.Device = "R" + chandev2[1];
                                    EData1[k].Comment = tabHeader[i] + " " + chandev[0] + ", ";
                                }
                                else
                                {
                                    string[] chandev = list[i, j].Split('/');
                                    string[] chandev2 = chandev[1].Split('R');
                                    string[] chandev3 = chandev2[0].Split('K');
                                    EData1[k].Output.Channel = int.Parse(chandev3[1]);
                                    EData1[k].Output.Device = "R" + chandev2[1];

                                    string[] chaldev = list[i, j].Split('/');
                                    EData1[k].Comment += tabHeader[i] + " " + chaldev[0];
                                    break;
                                }
                            }
                        }
                        k++;
                    }
                }
                else
                {
                    for (int i = 0; i < table.Columns.Count; i++)
                    {
                        if (list[i, j] != String.Empty)
                        {
                            tabHeader[i] = list[i, j];
                        }
                    }
                    nPossition = j;
                }
            }
            return EData1;
        }
        //Парсим Excel3
        public static EData3[] ParseEx3(string path)
        {
            var table = ParseTable(path);

            string[,] list = new string[table.Columns.Count, table.Rows.Count];
            for (int i = 0; i < table.Columns.Count; i++)
            {
                for (int j = 0; j < table.Rows.Count; j++)
                {
                    list[i, j] = table.Rows[j][i].ToString();
                }
            }

            int lenght = 0;
            for (int j = 0; j < table.Rows.Count; j++)
            {
                var parsedValue = 0;
                if (list[0, j] != String.Empty && int.TryParse(list[0, j], out parsedValue) && parsedValue > lenght)
                {
                    lenght = parsedValue;
                }
            }

            int counter = 0;
            int size = 0;
            int[] intSize = new int[lenght];
            for (int i = 0; i < intSize.Length; i++)
            {
                while (size == 0)
                {
                    intSize[i] = size = list[4, counter].ToCharArray().Where(x => x == '/').Count();
                    counter++;
                }
                size = 0;
            }

            EData3 = new EData3[lenght];
            for (int i = 0; i < EData3.Length; i++)
            {
                EData3[i] = new EData3(intSize[i] + 1);
            }
            int k = 0;
            for (int i = 0; i < table.Rows.Count; i++)
            {
                if (list[0,i] != String.Empty && char.IsDigit(Convert.ToChar(list[0, i])))
                {
                    EData3[k].Index = Convert.ToInt32(list[0, i]);
                    switch (list[1,i])
                    {
                        case "3":
                            EData3[k].MultMode = multMode.DCVoltage;
                            break;
                        case "2":
                            EData3[k].MultMode = multMode.Resistance;
                            break;
                        case "1":
                            EData3[k].MultMode = multMode.DiodeTest;
                            break;
                        case "0":
                            EData3[k].MultMode = 0;
                            break;
                    }
                    EData3[k].CurrSource= Convert.ToInt16(list[2, i]);

                    var power = list[3, i].Split('/');
                    EData3[k].VoltSource.Power1 = Convert.ToInt32(power[0]);   
                    //EData3[k].VoltSource.Power2 = Convert.ToInt32(power[1]);
 
                    ////////////////
                    /// Parse K5R1
                    ///////////////
                    
                    k++;
                }
            }
            return EData3;
        }
    }
}
