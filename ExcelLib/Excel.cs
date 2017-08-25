using System;
using System.Data;
using System.IO;
using System.Linq;
using MyClass.WriteToExcel;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace ExcelLib
{
     public struct MyStruct
    {
        
    }
    public static class Excel
    {
        public static EData1And2[] EData { get; set; }
        public static EData3[] EData3 { get; set; }
        public static EData4[] EData4 { get; set; }
        public static BPPPTest[] BPPPTest { get; set; }

        //Преобразование из XLS к DataTable
        private static DataTable ParseTable(string path)
        {
            DataTable table;
            if (File.Exists(path))
            {

                IWorkbook workbook; //IWorkbook determina si es xls o xlsx              
                ISheet worksheet;
                string first_sheet_name;

                using (FileStream fs = new FileStream(path, FileMode.Open, FileAccess.Read))
                {
                    workbook = WorkbookFactory.Create(fs); //Abre tanto XLS como XLSX
                    worksheet = workbook.GetSheetAt(0); //Obtener Hoja por indice
                    first_sheet_name = worksheet.SheetName; //Obtener el nombre de la Hoja

                    table = new DataTable(first_sheet_name);
                    table.Rows.Clear();
                    table.Columns.Clear();

                    // Leer Fila por fila desde la primera
                    for (int rowIndex = 0; rowIndex <= worksheet.LastRowNum; rowIndex++)
                    {
                        DataRow newReg = null;
                        IRow row = worksheet.GetRow(rowIndex);
                        IRow row2 = null;
                        IRow row3 = null;

                        if (rowIndex == 0)
                        {
                            row2 =
                                worksheet.GetRow(
                                    rowIndex +
                                    1); //Si es la Primera fila, obtengo tambien la segunda para saber el tipo de datos
                            row3 = worksheet.GetRow(rowIndex + 2); //Y la tercera tambien por las dudas
                        }

                        if (row != null) //null is when the row only contains empty cells 
                        {
                            if (rowIndex > 0) newReg = table.NewRow();

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
                                        ICell cell2;
                                        if (i == 0)
                                        {
                                            cell2 = row2.GetCell(cell.ColumnIndex);
                                        }
                                        else
                                        {
                                            cell2 = row3.GetCell(cell.ColumnIndex);
                                        }

                                        if (cell2 != null)
                                        {
                                            switch (cell2.CellType)
                                            {
                                                case CellType.Blank: break;
                                                case CellType.Boolean:
                                                    cellType2[i] = "System.Boolean";
                                                    break;
                                                case CellType.String:
                                                    cellType2[i] = "System.String";
                                                    break;
                                                case CellType.Numeric:
                                                    if (DateUtil.IsCellDateFormatted(cell2))
                                                    {
                                                        cellType2[i] = "System.DateTime";
                                                    }
                                                    else
                                                    {
                                                        cellType2[i] =
                                                            "System.Double"; //valorCell = cell2.NumericCellValue;
                                                    }
                                                    break;

                                                case CellType.Formula:
                                                    bool continuar = true;
                                                    switch (cell2.CachedFormulaResultType)
                                                    {
                                                        case CellType.Boolean:
                                                            cellType2[i] = "System.Boolean";
                                                            break;
                                                        case CellType.String:
                                                            cellType2[i] = "System.String";
                                                            break;
                                                        case CellType.Numeric:
                                                            if (DateUtil.IsCellDateFormatted(cell2))
                                                            {
                                                                cellType2[i] = "System.DateTime";
                                                            }
                                                            else
                                                            {
                                                                try
                                                                {
                                                                    //DETERMINAR SI ES BOOLEANO
                                                                    if (cell2.CellFormula == "TRUE()")
                                                                    {
                                                                        cellType2[i] = "System.Boolean";
                                                                        continuar = false;
                                                                    }
                                                                    if (continuar && cell2.CellFormula == "FALSE()")
                                                                    {
                                                                        cellType2[i] = "System.Boolean";
                                                                        continuar = false;
                                                                    }
                                                                    if (continuar)
                                                                    {
                                                                        cellType2[i] = "System.Double";
                                                                    }
                                                                }
                                                                catch
                                                                {
                                                                    // ignored
                                                                }
                                                            }
                                                            break;
                                                    }
                                                    break;
                                                default:
                                                    cellType2[i] = "System.String";
                                                    break;
                                            }
                                        }
                                    }

                                    //Resolver las diferencias de Tipos
                                    if (cellType2[0] == cellType2[1])
                                    {
                                    }
                                    else
                                    {
                                        if (cellType2[0] == null) cellType = cellType2[1];
                                        if (cellType2[1] == null) cellType = cellType2[0];
                                        if (cellType == "")
                                        {
                                        }
                                    }

                                    //Obtener el nombre de la Columna
                                    string colName = "Column_{0}";
                                    try
                                    {
                                        colName = cell.StringCellValue;
                                    }
                                    catch
                                    {
                                        colName = string.Format(colName, colIndex);
                                    }

                                    //Verificar que NO se repita el Nombre de la Columna
                                    foreach (DataColumn col in table.Columns)
                                    {
                                        if (col.ColumnName == colName)
                                            colName = string.Format("{0}_{1}", colName, colIndex);
                                    }

                                    //Agregar el campos de la tabla:
                                    DataColumn codigo = new DataColumn(); //colName, System.Type.GetType(cellType));
                                    table.Columns.Add(codigo);
                                    colIndex++;
                                }
                                else
                                {
                                    //Las demas filas son registros:
                                    switch (cell.CellType)
                                    {
                                        case CellType.Blank:
                                            valorCell = DBNull.Value;
                                            break;
                                        case CellType.Boolean:
                                            valorCell = cell.BooleanCellValue;
                                            break;
                                        case CellType.String:
                                            valorCell = cell.StringCellValue;
                                            break;
                                        case CellType.Numeric:
                                            if (DateUtil.IsCellDateFormatted(cell))
                                            {
                                                valorCell = cell.DateCellValue;
                                            }
                                            else
                                            {
                                                valorCell = cell.NumericCellValue;
                                            }
                                            break;
                                        case CellType.Formula:
                                            switch (cell.CachedFormulaResultType)
                                            {
                                                case CellType.Blank:
                                                    valorCell = DBNull.Value;
                                                    break;
                                                case CellType.String:
                                                    valorCell = cell.StringCellValue;
                                                    break;
                                                case CellType.Boolean:
                                                    valorCell = cell.BooleanCellValue;
                                                    break;
                                                case CellType.Numeric:
                                                    if (DateUtil.IsCellDateFormatted(cell))
                                                    {
                                                        valorCell = cell.DateCellValue;
                                                    }
                                                    else
                                                    {
                                                        valorCell = cell.NumericCellValue;
                                                    }
                                                    break;
                                            }
                                            break;
                                        default:
                                            valorCell = cell.StringCellValue;
                                            break;
                                    }
                                    //Agregar el nuevo Registro
                                    if (cell.ColumnIndex <= table.Columns.Count - 1)
                                        newReg[cell.ColumnIndex] = valorCell;
                                }
                            }
                        }
                        if (rowIndex > 0) table.Rows.Add(newReg);
                    }
                    table.AcceptChanges();
                }
            }
            else
            {
                throw new Exception("ERROR 404: El archivo especificado NO existe.");
            }
            return table;
        }

        /// <summary>Convierte un DataTable en un archivo de Excel (xls o Xlsx) y lo guarda en disco.</summary>
        /// <param name="pDatos">Datos de la Tabla a guardar. Usa el nombre de la tabla como nombre de la Hoja</param>
        /// <param name="pFilePath">Ruta del archivo donde se guarda.</param>
        private static void SaveTable(DataTable pDatos, string pFilePath)
        {
            try
            {
                if (pDatos != null && pDatos.Rows.Count > 0)
                {
                    IWorkbook workbook = null;
                    ISheet worksheet = null;

                    using (FileStream stream = new FileStream(pFilePath, FileMode.Create, FileAccess.ReadWrite))
                    {
                        string Ext = System.IO.Path.GetExtension(pFilePath); //<-Extension del archivo
                        switch (Ext.ToLower())
                        {
                            case ".xls":
                                HSSFWorkbook workbookH = new HSSFWorkbook();
                                NPOI.HPSF.DocumentSummaryInformation dsi = NPOI.HPSF.PropertySetFactory.CreateDocumentSummaryInformation();
                                dsi.Company = "Cutcsa"; dsi.Manager = "Departamento Informatico";
                                workbookH.DocumentSummaryInformation = dsi;
                                workbook = workbookH;
                                break;

                            case ".xlsx": workbook = new XSSFWorkbook(); break;
                        }

                        worksheet = workbook.CreateSheet(pDatos.TableName); //<-Usa el nombre de la tabla como nombre de la Hoja

                        //CREAR EN LA PRIMERA FILA LOS TITULOS DE LAS COLUMNAS
                        int iRow = 0;
                        if (pDatos.Columns.Count > 0)
                        {
                            int iCol = 0;
                            IRow fila = worksheet.CreateRow(iRow);
                            foreach (DataColumn columna in pDatos.Columns)
                            {
                                ICell cell = fila.CreateCell(iCol, CellType.String);
                                cell.SetCellValue(columna.ColumnName);
                                iCol++;
                            }
                            iRow++;
                        }

                        //FORMATOS PARA CIERTOS TIPOS DE DATOS
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

                        //AHORA CREAR UNA FILA POR CADA REGISTRO DE LA TABLA
                        foreach (DataRow row in pDatos.Rows)
                        {
                            IRow fila = worksheet.CreateRow(iRow);
                            int iCol = 0;
                            foreach (DataColumn column in pDatos.Columns)
                            {
                                ICell cell = null; //<-Representa la celda actual                               
                                object cellValue = row[iCol]; //<- El valor actual de la celda

                                switch (column.DataType.ToString())
                                {
                                    case "System.Boolean":
                                        if (cellValue != DBNull.Value)
                                        {
                                            cell = fila.CreateCell(iCol, CellType.Boolean);

                                            if (Convert.ToBoolean(cellValue)) { cell.SetCellFormula("TRUE()"); }
                                            else { cell.SetCellFormula("FALSE()"); }

                                            cell.CellStyle = _boolCellStyle;
                                        }
                                        break;

                                    case "System.String":
                                        if (cellValue != DBNull.Value)
                                        {
                                            cell = fila.CreateCell(iCol, CellType.String);
                                            cell.SetCellValue(Convert.ToString(cellValue));
                                        }
                                        break;

                                    case "System.Int32":
                                        if (cellValue != DBNull.Value)
                                        {
                                            cell = fila.CreateCell(iCol, CellType.Numeric);
                                            cell.SetCellValue(Convert.ToInt32(cellValue));
                                            cell.CellStyle = _intCellStyle;
                                        }
                                        break;
                                    case "System.Int64":
                                        if (cellValue != DBNull.Value)
                                        {
                                            cell = fila.CreateCell(iCol, CellType.Numeric);
                                            cell.SetCellValue(Convert.ToInt64(cellValue));
                                            cell.CellStyle = _intCellStyle;
                                        }
                                        break;
                                    case "System.Decimal":
                                        if (cellValue != DBNull.Value)
                                        {
                                            cell = fila.CreateCell(iCol, CellType.Numeric);
                                            cell.SetCellValue(Convert.ToDouble(cellValue));
                                            cell.CellStyle = _doubleCellStyle;
                                        }
                                        break;
                                    case "System.Double":
                                        if (cellValue != DBNull.Value)
                                        {
                                            cell = fila.CreateCell(iCol, CellType.Numeric);
                                            cell.SetCellValue(Convert.ToDouble(cellValue));
                                            cell.CellStyle = _doubleCellStyle;
                                        }
                                        break;

                                    case "System.DateTime":
                                        if (cellValue != DBNull.Value)
                                        {
                                            cell = fila.CreateCell(iCol, CellType.Numeric);
                                            cell.SetCellValue(Convert.ToDateTime(cellValue));

                                            //Si No tiene valor de Hora, usar formato dd-MM-yyyy
                                            DateTime cDate = Convert.ToDateTime(cellValue);
                                            if (cDate != null && cDate.Hour > 0) { cell.CellStyle = _dateTimeCellStyle; }
                                            else { cell.CellStyle = _dateCellStyle; }
                                        }
                                        break;
                                    default:
                                        break;
                                }
                                iCol++;
                            }
                            iRow++;
                        }

                        workbook.Write(stream);
                        stream.Close();
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        //Парсит Excel1 и Excel2
        /// <summary>
        /// Парсит таблицу типа 1, возвращает список тестов в указанном формате
        /// </summary>
        /// <param name="path"></param>
        /// <returns></returns>
        public static EData1And2[] ParseEx(string path)
        {
            try
            {
                var table = ParseTable(path);

                string[,] list = new string[table.Columns.Count, table.Rows.Count];

                if (table.Columns.Count == 0 && table.Rows.Count == 0)
                {
                    return null;
                }

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
                    int parsedValue;
                    if (list[0, j] != String.Empty && int.TryParse(list[0, j], out parsedValue) && parsedValue > lenght)
                    {
                        lenght = parsedValue;
                    }
                }


                EData = new EData1And2[lenght];
                for (int i = 0; i < EData.Length; i++)
                {
                    EData[i] = new EData1And2();
                }

                var k = 0;
                string[] tabHeader = new string[table.Columns.Count];
                for (int j = 0; j < table.Rows.Count; j++)
                {
                    if (list[0, j] != "№")
                    {
                        int value;
                        if (int.TryParse(list[0, j], out value))
                        {
                            EData[k].Index = value;

                            for (int i = 1; i < table.Columns.Count; i++)
                            {
                                if (list[i, j] != String.Empty)
                                {
                                    if (i == 1)
                                    {
                                        string[] chandev = list[i, j].Split('/');
                                        string[] chandev2 = chandev[1].Split('R');
                                        string[] chandev3 = chandev2[0].Split('K');
                                        EData[k].Input.Channel = int.Parse(chandev3[1]);
                                        EData[k].Input.Device = "R" + chandev2[1];
                                        EData[k].Comment = tabHeader[i] + " " + chandev[0] + ", ";
                                    }
                                    else
                                    {
                                        string[] chandev = list[i, j].Split('/');
                                        string[] chandev2 = chandev[1].Split('R');
                                        string[] chandev3 = chandev2[0].Split('K');
                                        EData[k].Output.Channel = int.Parse(chandev3[1]);
                                        EData[k].Output.Device = "R" + chandev2[1];

                                        string[] chaldev = list[i, j].Split('/');
                                        EData[k].Comment += tabHeader[i] + " " + chaldev[0];
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
                    }
                }
                return EData;
            }
            catch (Exception ex)
            {
                return null;
            }
        }

        //Парсим Excel3
        /// <summary>
        /// Парсит таблицу типа 3, возвращает список тестов в указанном формате
        /// </summary>
        /// <param name="path"></param>
        /// <returns></returns>
        public static EData3[] ParseEx3(string path)
        {
            try
            {
                var table = ParseTable(path);

                string[,] list = new string[table.Columns.Count, table.Rows.Count];

                if (table.Columns.Count == 0 && table.Rows.Count == 0)
                {
                    return null;
                }

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
                    int parsedValue;
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
                int counter2 = 0;
                int[] intSize2 = new int[lenght];
                for (int i = 0; i < intSize2.Length; i++)
                {
                    while (true)
                    {
                        if (list[2, counter2] != String.Empty && list[2, counter2] != "источник питания,I mA")
                        {
                            intSize2[i] = list[2, counter2].ToCharArray().Where(x => x == '/').Count();
                            counter2++;
                            break;
                        }
                        counter2++;
                    }

                }

                EData3 = new EData3[lenght];
                for (int i = 0; i < EData3.Length; i++)
                {
                    EData3[i] = new EData3(intSize[i] + 1, intSize2[i] + 1);
                }
                int k = 0;
                for (int i = 0; i < table.Rows.Count; i++)
                {
                    int num;
                    bool isNum = int.TryParse(list[0, i], out num);
                    if (list[0, i] != String.Empty && isNum)
                    {
                        EData3[k].Index = Convert.ToInt32(list[0, i]);
                        switch (list[1, i])
                        {
                            case "3":
                                EData3[k].MultMode = MultMode.DCVoltage;
                                break;
                            case "2":
                                EData3[k].MultMode = MultMode.Resistance;
                                break;
                            case "1":
                                EData3[k].MultMode = MultMode.DiodeTest;
                                break;
                            case "0":
                                EData3[k].MultMode = 0;
                                break;
                        }
                        if (EData3[k].CurrSource.Length != 1)
                        {
                            string[] curr = list[2, i].Split('/');
                            EData3[k].CurrSource[0].CurrSource = Convert.ToInt32(curr[0]);
                            EData3[k].CurrSource[1].CurrSource = Convert.ToInt32(curr[1]);
                        }
                        else
                        {
                            EData3[k].CurrSource[0].CurrSource = Convert.ToInt32(list[2, i]);
                        }


                        var voltage = list[3, i].Split('/');

                        if (voltage[0] == "-")
                        {
                            EData3[k].VoltSupply.V1 = 0;
                        }
                        else
                        {
                            EData3[k].VoltSupply.V1 = Convert.ToInt32(voltage[0]);
                        }

                        if (voltage[1] == "-")
                        {
                            EData3[k].VoltSupply.V2 = 0;
                        }
                        else
                        {
                            EData3[k].VoltSupply.V2 = Convert.ToInt32(voltage[1]);
                        }

                        var device = list[4, i].Split('/');
                        for (int j = 0; j < EData3[k].Input.Length; j++)
                        {
                            var device2 = device[j].Split('R');
                            var device3 = device2[0].Split('K');
                            EData3[k].Input[j].Device = @"R" + device2[1];
                            if (device3[0] != String.Empty)
                            {
                                EData3[k].Input[j].Channel = Convert.ToInt32(device3[0]);
                            }
                            else
                            {
                                EData3[k].Input[j].Channel = Convert.ToInt32(device3[1]);
                            }
                        }
                        //Комментарий
                        EData3[k].Comment = list[5, i];

                        //Контроль
                        switch (list[6, i])
                        {
                            case "напряжение":
                                EData3[k].Control = Control.Напряжение;
                                break;
                            case "сопротивление":
                                EData3[k].Control = Control.Сопротивление;
                                break;
                            case "индикация":
                                EData3[k].Control = Control.Индикация;
                                break;
                            case "падение напряжение БК":
                                EData3[k].Control = Control.ПадениеНапряженияБк;
                                break;
                            case "падение напряжение БЭ":
                                EData3[k].Control = Control.ПадениеНапряженияБэ;
                                break;
                            case "падение напряжение КБ":
                                EData3[k].Control = Control.ПадениеНапряженияКб;
                                break;
                            case "падение напряжение ЭБ":
                                EData3[k].Control = Control.ПадениеНапряженияЭб;
                                break;
                            case "падение напряжение ЭК":
                                EData3[k].Control = Control.ПадениеНапряженияЭк;
                                break;
                        }

                        if (list[7, i] != String.Empty)
                        {
                            double data;
                            if (!Double.TryParse(list[7, i], out data))
                            {
                                var min = list[7, i].Split(' ');
                                EData3[k].ValMin = Convert.ToDouble(min[0]);
                                EData3[k].ValUnit = min[1];
                            }
                            else
                            {
                                EData3[k].ValMin = data;
                            }
                        }
                        else
                        {
                            EData3[k].ValMin = 0;
                        }

                        if (list[8, i] != String.Empty)
                        {
                            if (list[8, i] != "∞")
                            {
                                double data2;
                                if (!Double.TryParse(list[8, i], out data2) && list[8, i] != String.Empty)
                                {
                                    var max = list[8, i].Split(' ');
                                    EData3[k].ValMax = Convert.ToDouble(max[0]);
                                }
                                else
                                {
                                    EData3[k].ValMax = data2;
                                }
                            }
                            else
                            {
                                EData3[k].ValMax = Double.MaxValue;
                            }
                        }
                        else
                        {
                            EData3[k].ValMax = 0;
                        }
                        k++;
                    }
                }
                return EData3;
            }
            catch (Exception ex)
            {
                return null;
            }
        }

        //Парсим Excel4
        /// <summary>
        /// Парсит таблицу типа 4, возвращает список тестов в указанном формате
        /// </summary>
        /// <param name="path"></param>
        /// <returns></returns>
        public static EData4[] ParseEx4(string path)
        {
            try
            {
                var table = ParseTable(path);

                string[,] list = new string[table.Columns.Count, table.Rows.Count];

                if (table.Columns.Count == 0 && table.Rows.Count == 0)
                {
                    return null;
                }

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
                    int parsedValue;
                    if (list[0, j] != String.Empty && int.TryParse(list[0, j], out parsedValue) && parsedValue > lenght)
                    {
                        lenght = parsedValue;
                    }
                }


                EData4 = new EData4[lenght];
                for (int i = 0; i < EData4.Length; i++)
                {
                    EData4[i] = new EData4();
                }
                int k = 0;
                for (int i = 0; i < table.Rows.Count; i++)
                {
                    int value;
                    if (list[0,i] != String.Empty && int.TryParse(list[0,i], out value))
                    {
                        EData4[k].Index = value;

                        var indata = list[1, i].Split('r');
                        var indata2 = indata[0].Split('k');
                        EData4[k].Input.Device = "R" + indata[1];
                        EData4[k].Input.Channel = Convert.ToInt32(indata2[1]);

                        var outdata = list[2, i].Split('r');
                        var outdata2 = outdata[0].Split('k');
                        EData4[k].Output.Device = "R" + outdata[1];
                        EData4[k].Output.Channel = Convert.ToInt32(outdata2[1]);
                        k++;
                    }
                }
                return EData4;
            }
            catch
            {
                return null;
            }
        }

        //Парсим Excel5
        /// <summary>
        /// Используется для блока проверки печатных плат, возвращает тесты ввиде массива (BPPP)
        /// </summary>
        /// <param name="path"></param>
        /// <returns></returns>
        public static BPPPTest[] ParseBPPP(string path)
        {
            try
            {
                var table = ParseTable(path);

                string[,] list = new string[table.Columns.Count, table.Rows.Count];

                if (table.Columns.Count == 0 && table.Rows.Count == 0)
                {
                    return null;
                }

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
                    int parsedValue;
                    if (list[0, j] != String.Empty && int.TryParse(list[0, j], out parsedValue) && parsedValue > lenght)
                    {
                        lenght = parsedValue;
                    }
                }

                int counter = 0;
                int[] intSize = new int[lenght];
                for (int i = 0; i < intSize.Length; i++)
                {
                    while (true)
                    {
                        if ((list[1, counter] != String.Empty))
                        {
                            intSize[i] = list[1, counter].ToCharArray().Where(x => x == '/').Count();
                            counter++;
                            break;
                        }
                        counter++;
                    }
                }
                int counter2 = 0;
                int[] intSize2 = new int[lenght];
                for (int i = 0; i < intSize2.Length; i++)
                {
                    while (true)
                    {
                        if (list[2, counter2] != String.Empty)
                        {
                            intSize2[i] = list[2, counter2].ToCharArray().Where(x => x == '/').Count();
                            counter2++;
                            break;
                        }
                        counter2++;
                    }

                }

                BPPPTest = new BPPPTest[lenght];
                for (int i = 0; i < BPPPTest.Length; i++)
                {
                    BPPPTest[i] = new BPPPTest(intSize[i] + 1, intSize2[i] + 1);
                }
                int k = 0;
                for (int i = 0; i < table.Rows.Count; i++)
                {
                    int value;
                    if (list[0, i] != String.Empty && int.TryParse(list[0, i], out value))
                    {
                        BPPPTest[k].Index = value;
                        BPPPTest[k].Max = Double.NegativeInfinity; //list[3, i]
                        BPPPTest[k].Min = Double.NegativeInfinity; //list[4, i]
                        BPPPTest[k].Value = Double.NegativeInfinity; //list[5, i]
                        BPPPTest[k].Comment = list[6, i] +" "+ list[7, i];
                        BPPPTest[k].Range = Convert.ToInt32(list[8, i]);

                        var indata = list[1, i].Split('/');
                        for (int j = 0; j < BPPPTest[k].Input.Length; j++)
                        {
                            var indata2 = indata[j].Split('r');
                            var indata3 = indata2[0].Split('k');
                            BPPPTest[k].Input[j].Device = "R" + indata2[1];
                            BPPPTest[k].Input[j].Channel = Convert.ToInt32(indata3[1]);
                        }

                        var outdata = list[2, i].Split('/');
                        for (int j = 0; j < BPPPTest[k].Output.Length; j++)
                        {
                            var outdata2 = outdata[j].Split('r');
                            var outdata3 = outdata2[0].Split('k');
                            BPPPTest[k].Output[j].Device = "R" + outdata2[1];
                            BPPPTest[k].Output[j].Channel = Convert.ToInt32(outdata3[1]);
                        }
                        k++;
                    }
                }
                return BPPPTest;
            }
            catch(Exception ex)
            {
                return null;
            }
        }

        //Сохраняем Excel5(BPPPTest)
        /// <summary>
        /// Используется для блока проверки печатных плат, сохраняет тесты в виде таблицы xls (BPPP)
        /// </summary>
        /// <param name="data"></param>
        /// <param name="path"></param>
        /// <returns></returns>
        public static bool SaveBPPP(BPPPTest[] data, string path)
        {
            DataTable _myDataTable = new DataTable();


            _myDataTable.Columns.Add(new DataColumn("Номер проверки"));
            _myDataTable.Columns.Add(new DataColumn("A"));
            _myDataTable.Columns.Add(new DataColumn("B"));
            _myDataTable.Columns.Add(new DataColumn("Максимальные допустимые значания"));
            _myDataTable.Columns.Add(new DataColumn(String.Empty));
            _myDataTable.Columns.Add(new DataColumn("Измеренные значения"));
            _myDataTable.Columns.Add(new DataColumn("Комментарии"));
            _myDataTable.Columns.Add(new DataColumn(String.Empty));

            string[] myResult = { "", "", "", "MIN", "MAX", "", "A", "B" };
            DataRow row1 = _myDataTable.NewRow();

            DataRow _row = _myDataTable.NewRow();
            for (int i = 0; i < myResult.Length; i++)
            {
                row1[i] = myResult[i];
            }
            _myDataTable.Rows.Add(row1);

            for (int i = 0; i < data.Length; i++)
            {
                _row[i] = data[i];
            }
            _myDataTable.Rows.Add(_row);

            //dt.Rows[0][0] = @"Номер проверки";
            //dt.Rows[0][1] = @"A";
            //dt.Rows[0][2] = @"B";
            //dt.Rows[0][3] = @"Максимальные допустимые значания";
            //dt.Rows[0][5] = @"Измеренные значения";
            //dt.Rows[0][6] = @"Комментарии";

            //dt.Rows[1][3] = @"MIN";
            //dt.Rows[1][4] = @"MAX";
            //dt.Rows[1][6] = @"A";
            //dt.Rows[1][7] = @"B";

            //for (int j = 0; j < data.Length; j++)
            //{
            //    // create a DataRow using .NewRow()
            //    DataRow row = _myDataTable.NewRow();

            //    //// iterate over all columns to fill the row
            //    //for (int i = 0; i < ele; i++)
            //    //{
            //    //    row[i] = datar[i, j];
            //    //}

            //    // add the current row to the DataTable
            //    _myDataTable.Rows.Add(row);
            //}

            //DataTable dt = new DataTable();
            //dt.Rows.Clear();
            //dt.Columns.Clear();
            //dt.Rows.a

            //// Declare variables for DataColumn and DataRow objects.
            //DataColumn column;
            //DataRow row;

            //dt.Rows[0][0] = @"Номер проверки";
            //dt.Rows[0][1] = @"A";
            //dt.Rows[0][2] = @"B";
            //dt.Rows[0][3] = @"Максимальные допустимые значания";
            //dt.Rows[0][5] = @"Измеренные значения";
            //dt.Rows[0][6] = @"Комментарии";

            //dt.Rows[1][3] = @"MIN";
            //dt.Rows[1][4] = @"MAX";
            //dt.Rows[1][6] = @"A";
            //dt.Rows[1][7] = @"B";

            //int k = 0;
            //int z = 0;
            //for (int i = 2; i < dt.Rows.Count; i++)
            //{
            //    for (int j = 0; j < dt.Columns.Count; j++)
            //    {
            //        //dt.Rows[i][0] = data[k].Index;
            //        // dt.Rows[i][1] = data[k].
            //    }
            //}
            //dt.Columns.Add("Номер проверки", typeof(string));
            //dt.Columns.Add("A", typeof(string));
            //dt.Columns.Add("B", typeof(string));
            //dt.Columns.Add("Максимальные допустимые значения", typeof(string));
            //dt.Columns.Add("Измеренные значения", typeof(string));
            //dt.Columns.Add("Комментарий", typeof(string));
            //dt.Rows.Add();


            //using (FileStream stream = new FileStream(path, FileMode.Create, FileAccess.Write))
            //{
            //    IWorkbook wb = new XSSFWorkbook();
            //    ISheet sheet = wb.CreateSheet("Sheet1");
            //    ICreationHelper cH = wb.GetCreationHelper();
            //    for (int i = 0; i < dt.Rows.Count; i++)
            //    {
            //        IRow row = sheet.CreateRow(i);
            //        for (int j = 0; j < 3; j++)
            //        {
            //            ICell cell = row.CreateCell(j);
            //            cell.SetCellValue(cH.CreateRichTextString(dt.Rows[i].ItemArray[j].ToString()));
            //        }
            //    }
            //    wb.Write(stream);
            //}

            SaveTable(_myDataTable, path);

            return false;
        }
    }
}

