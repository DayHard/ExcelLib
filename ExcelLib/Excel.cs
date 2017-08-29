using System;
using System.Data;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using NPOI.SS.UserModel;
using MSExcel = Microsoft.Office.Interop.Excel;

namespace ExcelLib
{
    public static class Excel
    {

        public static EData1And2[] EData { get; set; }
        public static EData3[] EData3 { get; set; }
        public static EData4[] EData4 { get; set; }
        public static BPPPTest[] BPPPTest { get; set; }

        /// <summary>
        /// Преобразование из XLS к DataTable
        /// </summary>
        /// <param name="path"></param>
        /// <returns></returns>
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

        /// <summary>
        /// Парсит таблицу типа 1 и 2, возвращает список тестов в указанном формате
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
                                        list[i, j] = list[i, j].ToLower();
                                        string[] chandev = list[i, j].Split('/');
                                        string[] chandev2 = chandev[1].Split('r');
                                        string[] chandev3 = chandev2[0].Split('k');
                                        EData[k].Input.Channel = int.Parse(chandev3[1]);
                                        EData[k].Input.Device = "r" + chandev2[1];
                                        EData[k].Comment = tabHeader[i] + " " + chandev[0] + ", ";
                                    }
                                    else
                                    {
                                        list[i, j] = list[i, j].ToLower();
                                        string[] chandev = list[i, j].Split('/');
                                        string[] chandev2 = chandev[1].Split('r');
                                        string[] chandev3 = chandev2[0].Split('k');
                                        EData[k].Output.Channel = int.Parse(chandev3[1]);
                                        EData[k].Output.Device = "r" + chandev2[1];

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

                        list[4, i] = list[4, i].ToLower();
                        var device = list[4, i].Split('/');
                        for (int j = 0; j < EData3[k].Input.Length; j++)
                        {
                            var device2 = device[j].Split('r');
                            var device3 = device2[0].Split('k');
                            EData3[k].Input[j].Device = @"r" + device2[1];
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
                    if (list[0, i] != String.Empty && int.TryParse(list[0, i], out value))
                    {
                        EData4[k].Index = value;

                        list[1, i] = list[1, i].ToLower();
                        var indata = list[1, i].Split('r');
                        var indata2 = indata[0].Split('k');
                        EData4[k].Input.Device = "r" + indata[1];
                        EData4[k].Input.Channel = Convert.ToInt32(indata2[1]);

                        list[2, i] = list[2, i].ToLower();
                        var outdata = list[2, i].Split('r');
                        var outdata2 = outdata[0].Split('k');
                        EData4[k].Output.Device = "r" + outdata[1];
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

                        if (list[3, i] != String.Empty)
                            BPPPTest[k].Min = Convert.ToDouble(list[3, i]); //list[3, i]
                        else BPPPTest[k].Min = Double.NegativeInfinity;

                        if (list[4, i] != String.Empty)
                            BPPPTest[k].Max = Convert.ToDouble(list[4, i]); //list[4, i]
                        else BPPPTest[k].Max = Double.NegativeInfinity;

                        if (list[5, i] != String.Empty)
                            BPPPTest[k].Value = Convert.ToDouble(list[5, i]); //list[5, i]
                        else BPPPTest[k].Value = Double.NegativeInfinity;

                        BPPPTest[k].Comment = list[6, i] + " " + list[7, i];
                        BPPPTest[k].Range = Convert.ToInt32(list[8, i]);

                        list[1,i] = list[1, i].ToLower();
                        var indata = list[1, i].Split('/');
                        for (int j = 0; j < BPPPTest[k].Input.Length; j++)
                        {
                            var indata2 = indata[j].Split('r');
                            var indata3 = indata2[0].Split('k');
                            BPPPTest[k].Input[j].Device = "r" + indata2[1];
                            BPPPTest[k].Input[j].Channel = Convert.ToInt32(indata3[1]);
                        }

                        list[2, i] = list[2, i].ToLower();
                        var outdata = list[2, i].Split('/');
                        for (int j = 0; j < BPPPTest[k].Output.Length; j++)
                        {
                            var outdata2 = outdata[j].Split('r');
                            var outdata3 = outdata2[0].Split('k');
                            BPPPTest[k].Output[j].Device = "r" + outdata2[1];
                            BPPPTest[k].Output[j].Channel = Convert.ToInt32(outdata3[1]);
                        }
                        BPPPTest[k].Result = list[9, i];
                        k++;
                    }
                }
                return BPPPTest;
            }
            catch (Exception ex)
            {
                return null;
            }
        }

        /// <summary>
        /// Используется для блока проверки печатных плат, сохраняет тесты в виде таблицы xls (BPPP)
        /// </summary>
        /// <param name="data"></param>
        /// <param name="path"></param>
        /// <returns></returns>
        public static string SaveBPPP(BPPPTest[] data, string path)
        {
            try
            {
                string status = Wrapper.SaveBPPPWrapper(data, path);
                return status;
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
            finally
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }
    }

    /// <summary>
    /// Не использовать данный класс.
    /// </summary>
    // Необходим для очистки COM объектов
    public static class Wrapper
    {
        public static string SaveBPPPWrapper(BPPPTest[] data, string path)
        {
            MSExcel.Application exApp = new MSExcel.Application();

            exApp.Visible = false;
            exApp.ScreenUpdating = false;
            // Отключае запрос на перезапись файла (разрешает перезапись)
            exApp.DisplayAlerts = false;

            MSExcel.Workbook xlWorkBook;
            MSExcel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;

            xlWorkBook = exApp.Workbooks.Add(misValue);
            xlWorkSheet = (MSExcel.Worksheet)xlWorkBook.Worksheets.Item[1];

            //При добавлении столбцов, изменить размер массива(ниже при  MSExcel.Range r  тоже!)
            string[,] arr = new string[data.Length + 3, 10];
            arr[0, 0] = "Номер проверки";
            arr[0, 1] = "A";
            arr[0, 2] = "B";
            arr[0, 3] = "Максимальные допустимые значения";
            arr[0, 4] = " ";
            arr[1, 3] = "MIN";
            arr[1, 4] = "MAX";
            arr[0, 5] = "Измеренные значения";
            arr[0, 6] = "Комментарии";
            arr[1, 6] = "A";
            arr[0, 7] = " ";
            arr[1, 7] = "B";
            arr[0, 8] = "Диапазон";
            arr[1, 8] = "Ом";
            arr[0, 9] = "Результат";


            bool[] successColor = new bool[data.Length + 3];
            for (int i = 0; i < successColor.Length; i++)
            {
                successColor[i] = false;
            }
            int k = 0;
            for (int i = 2; i < data.Length + 2; i++)
            {
                arr[i, 0] = data[k].Index.ToString();

                for (int j = 0; j < data[k].Input.Length; j++)
                {
                    if (j > 0) arr[i, 1] += "/";
                    arr[i, 1] += "k" + data[k].Input[j].Channel + data[k].Input[j].Device;
                }

                for (int j = 0; j < data[k].Output.Length; j++)
                {
                    if (j > 0) arr[i, 2] += "/";
                    arr[i, 2] += "k" + data[k].Output[j].Channel + data[k].Output[j].Device;
                }

                // ReSharper disable once SpecifyACultureInStringConversionExplicitly
                arr[i, 3] = data[k].Min.ToString();
                // ReSharper disable once SpecifyACultureInStringConversionExplicitly
                arr[i, 4] = data[k].Max.ToString();
                // ReSharper disable once SpecifyACultureInStringConversionExplicitly
                arr[i, 5] = data[k].Value.ToString();
                var tcom = data[k].Comment.Split(' ');
                arr[i, 6] = tcom[0];
                arr[i, 7] = tcom[1];
                arr[i, 8] = data[k].Range.ToString();
                arr[i, 9] = data[k].Result;
                if (data[k].Result == "PASSED")
                    successColor[k] = true;
                k++;
            }
            try
            {
                MSExcel.Range r = xlWorkSheet.Range[xlWorkSheet.Cells[1, 1], xlWorkSheet.Cells[data.Length + 2, 10]];
                r.Value = arr;
                xlWorkSheet.Columns.EntireColumn.AutoFit();

                //Пример изменения цвета ячейки
                //xlWorkSheet.Cells[1, 1].Interior.Color = MSExcel.XlRgbColor.rgbRed;

                for (int i = 0; i < successColor.Length - 3; i++)
                {
                    if (successColor[i])
                        xlWorkSheet.Cells[i + 3, 10].Font.Color = MSExcel.XlRgbColor.rgbGreen;
                    else xlWorkSheet.Cells[i + 3, 10].Font.Color = MSExcel.XlRgbColor.rgbRed;
                }
                //var columnHeadingsRange = xlWorkSheet.Range[
                //    xlWorkSheet.Cells[COLUMN_HEADING_ROW, FIRST_COL],
                //    xlWorkSheet.Cells[COLUMN_HEADING_ROW, LAST_COL]];

                //columnHeadingsRange.Interior.Color = MSExcel.XlRgbColor.rgbSkyBlue;

                //columnHeadingsRange.Font.Color = MSExcel.XlRgbColor.rgbWhite;

                xlWorkBook.SaveAs(path, MSExcel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue,
                MSExcel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);

                return "Success";
            }
            catch (Exception ex)
            {
                switch (ex.Message)
                {
                    default:
                        return "Unknown error: " + ex.Message;
                }
            }
            finally
            {
                xlWorkBook.Close(true, misValue, misValue);
                exApp.Quit();

                Marshal.ReleaseComObject(xlWorkSheet);
                Marshal.ReleaseComObject(xlWorkBook);
                Marshal.ReleaseComObject(exApp);
            }
        }
    }
}
