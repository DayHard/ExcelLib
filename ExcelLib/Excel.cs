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

        public static DAQTest[] DAQTest { get; set; }
        public static IndTest[] IndTest { get; set; }
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
            try
            {
                if (File.Exists(path))
                {

                    IWorkbook workbook; //IWorkbook determina si es xls o xlsx              
                    ISheet worksheet;
                    string first_sheet_name;
                    
                    using (FileStream fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
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
            }
            catch (Exception ex)
            {
                return null;
            }
            return table;
        }

        /// <summary>
        /// Парсит таблицу типа 1 и 2, возвращает список тестов в указанном формате. Устаревший метод - работает лишь для кривонаписанных Виталиком файлов Excel
        /// </summary>
        /// <param name="path">Путь файла</param>
        /// <returns></returns>
        public static DAQTest[] ParseDAQ_Old(string path)
        {
            try
            {
                var table = ParseTable(path);

                if (table == null) return null;

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


                DAQTest = new DAQTest[lenght];
                for (int i = 0; i < DAQTest.Length; i++)
                {
                    DAQTest[i] = new DAQTest();
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
                            DAQTest[k].Index = value;
                            DAQTest[k].Result = list[12, j];
                            DAQTest[k].Value = list[13, j];

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
                                        DAQTest[k].Input.Channel = int.Parse(chandev3[1]);
                                        DAQTest[k].Input.Device = "r" + chandev2[1];
                                        DAQTest[k].Comment = tabHeader[i] + " " + chandev[0] + ", ";
                                    }
                                    else
                                    {
                                        list[i, j] = list[i, j].ToLower();
                                        string[] chandev = list[i, j].Split('/');
                                        string[] chandev2 = chandev[1].Split('r');
                                        string[] chandev3 = chandev2[0].Split('k');
                                        DAQTest[k].Output.Channel = int.Parse(chandev3[1]);
                                        DAQTest[k].Output.Device = "r" + chandev2[1];

                                        string[] chaldev = list[i, j].Split('/');
                                        DAQTest[k].Comment += tabHeader[i] + " " + chaldev[0];
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
                return DAQTest;
            }
            catch (Exception ex)
            {
                return null;
            }
        }

        /// <summary>
        /// Парсит таблицу типа 1 и 2, возвращает список тестов в указанном формате
        /// </summary>
        /// <param name="path"></param>
        /// <returns>Возвращает массив объектов</returns>
        public static DAQTest[] ParseDAQ(string path)
        {
            try
            {
                var table = ParseTable(path);

                if (table == null) return null;

                string[,] list = new string[table.Columns.Count, table.Rows.Count];

                if (table.Columns.Count == 0 || table.Rows.Count == 0)
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
                DAQTest = new DAQTest[table.Rows.Count - 1];
                for (int i = 0; i < DAQTest.Length; i++)
                {
                    DAQTest[i] = new DAQTest();
                }

                int k = 0;
                for (int i = 1; i < table.Rows.Count; i++)
                {
                    DAQTest[k].Index = Convert.ToInt32(list[0, i]);
                    DAQTest[k].Input.Device = list[1, i].ToUpper();
                    DAQTest[k].Input.Channel = Convert.ToInt32(list[2, i]);
                    DAQTest[k].Output.Device = list[3, i].ToUpper();
                    DAQTest[k].Output.Channel = Convert.ToInt32(list[4, i]);
                    DAQTest[k].Comment = list[5, i];
                    DAQTest[k].Value = list[6, i];
                    DAQTest[k].Result = list[7, i];
                    k++;
                }
             
                return DAQTest;
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
        public static IndTest[] ParseInd(string path)
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

                IndTest = new IndTest[lenght];
                for (int i = 0; i < IndTest.Length; i++)
                {
                    IndTest[i] = new IndTest(intSize[i] + 1, intSize2[i] + 1);
                }
                int k = 0;
                for (int i = 0; i < table.Rows.Count; i++)
                {
                    int num;
                    bool isNum = int.TryParse(list[0, i], out num);
                    if (list[0, i] != String.Empty && isNum)
                    {
                        IndTest[k].Index = Convert.ToInt32(list[0, i]);
                        switch (list[1, i])
                        {
                            case "3":
                                IndTest[k].MultMode = MultMode.DCVoltage;
                                break;
                            case "2":
                                IndTest[k].MultMode = MultMode.Resistance;
                                break;
                            case "1":
                                IndTest[k].MultMode = MultMode.DiodeTest;
                                break;
                            case "0":
                                IndTest[k].MultMode = 0;
                                break;
                        }
                        if (IndTest[k].CurrSource.Length != 1)
                        {
                            string[] curr = list[2, i].Split('/');
                            IndTest[k].CurrSource[0].CurrSource = Convert.ToInt32(curr[0]);
                            IndTest[k].CurrSource[1].CurrSource = Convert.ToInt32(curr[1]);
                        }
                        else
                        {
                            IndTest[k].CurrSource[0].CurrSource = Convert.ToInt32(list[2, i]);
                        }


                        var voltage = list[3, i].Split('/');

                        if (voltage[0] == "-")
                        {
                            IndTest[k].VoltSupply.V1 = 0;
                        }
                        else
                        {
                            IndTest[k].VoltSupply.V1 = Convert.ToInt32(voltage[0]);
                        }

                        if (voltage[1] == "-")
                        {
                            IndTest[k].VoltSupply.V2 = 0;
                        }
                        else
                        {
                            IndTest[k].VoltSupply.V2 = Convert.ToInt32(voltage[1]);
                        }

                        list[4, i] = list[4, i].ToLower();
                        var device = list[4, i].Split('/');
                        for (int j = 0; j < IndTest[k].Input.Length; j++)
                        {
                            var device2 = device[j].Split('r');
                            var device3 = device2[0].Split('k');
                            IndTest[k].Input[j].Device = @"r" + device2[1];
                            if (device3[0] != String.Empty)
                            {
                                IndTest[k].Input[j].Channel = Convert.ToInt32(device3[0]);
                            }
                            else
                            {
                                IndTest[k].Input[j].Channel = Convert.ToInt32(device3[1]);
                            }
                        }
                        //Комментарий
                        IndTest[k].Comment = list[5, i];

                        //Контроль
                        switch (list[6, i])
                        {
                            case "напряжение":
                            default:
                                IndTest[k].Control = Control.Напряжение;
                                break;
                            case "сопротивление":
                                IndTest[k].Control = Control.Сопротивление;
                                break;
                            case "индикация":
                                IndTest[k].Control = Control.Индикация;
                                break;
                            case "падение напряжения БК":
                                IndTest[k].Control = Control.ПадениеНапряженияБк;
                                break;
                            case "падение напряжения БЭ":
                                IndTest[k].Control = Control.ПадениеНапряженияБэ;
                                break;
                            case "падение напряжения КБ":
                                IndTest[k].Control = Control.ПадениеНапряженияКб;
                                break;
                            case "падение напряжения ЭБ":
                                IndTest[k].Control = Control.ПадениеНапряженияЭб;
                                break;
                            case "падение напряжения ЭК":
                                IndTest[k].Control = Control.ПадениеНапряженияЭк;
                                break;
                        }

                        if (list[7, i] != String.Empty)
                        {
                            double data;
                            if (!Double.TryParse(list[7, i], out data))
                            {
                                var min = list[7, i].Split(' ');
                                IndTest[k].ValMin = Convert.ToDouble(min[0]);
                                IndTest[k].ValUnit = min[1];
                            }
                            else
                            {
                                IndTest[k].ValMin = data;
                            }
                        }
                        else
                        {
                            IndTest[k].ValMin = 0;
                        }

                        if (list[8, i] != String.Empty)
                        {
                            if (list[8, i] != "∞")
                            {
                                double data2;
                                if (!Double.TryParse(list[8, i], out data2) && list[8, i] != String.Empty)
                                {
                                    var max = list[8, i].Split(' ');
                                    IndTest[k].ValMax = Convert.ToDouble(max[0]);
                                }
                                else
                                {
                                    IndTest[k].ValMax = data2;
                                }
                            }
                            else
                            {
                                IndTest[k].ValMax = Double.MaxValue;
                            }
                        }
                        else
                        {
                            IndTest[k].ValMax = 0;
                        }
                        k++;
                    }
                }
                return IndTest;
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
        /// <returns>Возвращает BPPPTest[]</returns>
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
                            BPPPTest[k].Delay = Convert.ToUInt16(list[3, i]); //list[3, i]
                        else BPPPTest[k].Delay = 0;

                        if (list[4, i] != String.Empty)
                            BPPPTest[k].Min = Convert.ToDouble(list[4, i]); //list[4, i]
                        else BPPPTest[k].Min = Double.NegativeInfinity;

                        if (list[5, i] != String.Empty)
                            BPPPTest[k].Max = Convert.ToDouble(list[5, i]); //list[5, i]
                        else BPPPTest[k].Max = Double.NegativeInfinity;

                        if (list[6, i] != String.Empty)
                            BPPPTest[k].Value = Convert.ToDouble(list[6, i]); //list[6, i]
                        else BPPPTest[k].Value = Double.NegativeInfinity;

                        BPPPTest[k].Comment = list[7, i] + " " + list[8, i];
                        BPPPTest[k].Range = Convert.ToInt32(list[9, i]);

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
                        BPPPTest[k].Result = list[10, i];
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
        /// <returns>Возращает статус сохранения</returns>
        public static string SaveBPPP(BPPPTest[] data, string path)
        {
            try
            {
                string status = Wrapper.SaveBpppWrapper(data, path);
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

        /// <summary>
        /// Используется для блока проверки печатных плат, сохраняет тесты в виде таблицы xls (DAQ)
        /// </summary>
        /// <param name="data"></param>
        /// <param name="path"></param>
        /// <returns>Возращает статус сохранения</returns>
        public static string SaveDAQ_Old(DAQTest[] data, string path)
        {
            try
            {
                string status = Wrapper.SaveDaq_Old_Wrapper(data, path);
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
        /// <summary>
        /// Используется для сохранения файла ошибок, сохраняет ошибочные тесты в виде таблицы xls(DAQ)
        /// </summary>
        /// <param name="data"></param>
        /// <param name="path"></param>
        /// <returns>Возвращает статус сохранения</returns>
        public static string SaveDAQ(DAQTest[] data, string path)
        {
            try
            {
                string status = Wrapper.SaveDaqWrapper(data, path);
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
        public static string SaveBpppWrapper(BPPPTest[] data, string path)
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
            string[,] arr = new string[data.Length + 3, 11];
            arr[0, 0] = "Номер проверки";
            arr[0, 1] = "A";
            arr[0, 2] = "B";
            arr[0, 3] = "Задержка, мс";
            arr[0, 4] = "Максимальные допустимые значения";
            arr[0, 5] = " ";
            arr[1, 4] = "MIN";
            arr[1, 5] = "MAX";
            arr[0, 6] = "Измеренные значения";
            arr[0, 7] = "Комментарии";
            arr[1, 7] = "A";
            arr[0, 8] = " ";
            arr[1, 8] = "B";
            arr[0, 9] = "Диапазон";
            arr[1, 9] = "Ом";
            arr[0, 10] = "Результат";


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

                arr[i, 3] = data[k].Delay.ToString();
                // ReSharper disable once SpecifyACultureInStringConversionExplicitly
                arr[i, 4] = data[k].Min.ToString();
                // ReSharper disable once SpecifyACultureInStringConversionExplicitly
                arr[i, 5] = data[k].Max.ToString();
                // ReSharper disable once SpecifyACultureInStringConversionExplicitly
                arr[i, 6] = data[k].Value.ToString();
                var tcom = data[k].Comment.Split(' ');
                arr[i, 7] = tcom[0];
                arr[i, 8] = tcom[1];
                arr[i, 9] = data[k].Range.ToString();
                arr[i, 10] = data[k].Result;
                if (data[k].Result == "PASSED")
                    successColor[k] = true;
                k++;
            }
            try
            {
                MSExcel.Range r = xlWorkSheet.Range[xlWorkSheet.Cells[1, 1], xlWorkSheet.Cells[data.Length + 2, 11]];
                r.Value = arr;
                xlWorkSheet.Columns.EntireColumn.AutoFit();

                //Пример изменения цвета ячейки
                //xlWorkSheet.Cells[1, 1].Interior.Color = MSExcel.XlRgbColor.rgbRed;

                for (int i = 0; i < successColor.Length - 3; i++)
                {
                    if (successColor[i])
                        xlWorkSheet.Cells[i + 3, 11].Font.Color = MSExcel.XlRgbColor.rgbGreen;
                    else xlWorkSheet.Cells[i + 3, 11].Font.Color = MSExcel.XlRgbColor.rgbRed;
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

        public static string SaveDaq_Old_Wrapper(DAQTest[] data, string path)
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
            string[,] arr = new string[data.Length + 3 + 16, 14];
            
            // Без этих элементов некорректно отрабатывает метод ParseTable
            arr[0, 0] = "ТЕСТЫ ЦЕПЕЙ";
            arr[0, 1] = " ";
            arr[0, 2] = " ";
            arr[0, 3] = " ";
            arr[0, 4] = " ";
            arr[0, 5] = " ";
            arr[0, 6] = " ";
            arr[0, 7] = " ";
            arr[0, 8] = " ";
            arr[0, 9] = " ";
            arr[0, 10] = " ";
            arr[0, 11] = " ";
            arr[0, 12] = " ";
            arr[0, 13] = " ";

            arr[1, 0] = " ";
            arr[2, 1] = "входы";
            arr[2, 12] = "Результат";
            arr[2, 13] = "Значение";
            //Шапка 1
            arr[3, 0] = "№";
            arr[3, 1] = "XP1";
            arr[3, 2] = "XS1";
            arr[3, 3] = "XS2";
            arr[3, 4] = "XS3";
            arr[3, 5] = "XS4";
            arr[3, 6] = "XS5";
            arr[3, 7] = "XS6";
            arr[3, 8] = "XS7";
            arr[3, 9] = "XS8";
            arr[3, 10] = "XS9";
            arr[3, 11] = "XS10";
            //Шапка 2
            arr[72, 0] = "№";
            arr[72, 1] = "XP2";
            arr[72, 2] = "XS1";
            arr[72, 3] = "XS2";
            arr[72, 4] = "XS3";
            arr[72, 5] = "XS4";
            arr[72, 6] = "XS5";
            arr[72, 7] = "XS6";
            arr[72, 8] = "XS7";
            arr[72, 9] = "XS8";
            arr[72, 10] = "XS9";
            arr[72, 11] = "XS10";
            //Шапка 3
            arr[91, 0] = "№";
            arr[91, 1] = "XP3";
            arr[91, 2] = "XS1";
            arr[91, 3] = "XS2";
            arr[91, 4] = "XS3";
            arr[91, 5] = "XS4";
            arr[91, 6] = "XS5";
            arr[91, 7] = "XS6";
            arr[91, 8] = "XS7";
            arr[91, 9] = "XS8";
            arr[91, 10] = "XS9";
            arr[91, 11] = "XS10";
            //Шапка 4
            arr[172, 0] = "№";
            arr[172, 1] = "XP4";
            arr[172, 2] = "XS1";
            arr[172, 3] = "XS2";
            arr[172, 4] = "XS3";
            arr[172, 5] = "XS4";
            arr[172, 6] = "XS5";
            arr[172, 7] = "XS6";
            arr[172, 8] = "XS7";
            arr[172, 9] = "XS8";
            arr[172, 10] = "XS9";
            arr[172, 11] = "XS10";
            //Шапка 5
            arr[231, 0] = "№";
            arr[231, 1] = "XP4";
            arr[231, 2] = "XS1";
            arr[231, 3] = "XS2";
            arr[231, 4] = "XS3";
            arr[231, 5] = "XS4";
            arr[231, 6] = "XS5";
            arr[231, 7] = "XS6";
            arr[231, 8] = "XS7";
            arr[231, 9] = "XS8";
            arr[231, 10] = "XS9";
            arr[231, 11] = "XS10";
            //Шапка 6
            arr[248, 0] = "№";
            arr[248, 1] = "XS2";
            arr[248, 2] = "XS1";
            arr[248, 3] = "XS2";
            arr[248, 4] = "XS3";
            arr[248, 5] = "XS4";
            arr[248, 6] = "XS5";
            arr[248, 7] = "XS6";
            arr[248, 8] = "XS7";
            arr[248, 9] = "XS8";
            arr[248, 10] = "XS9";
            arr[248, 11] = "XS10";
            //Шапка 7
            arr[303, 0] = "№";
            arr[303, 1] = "XS3";
            arr[303, 2] = "XS1";
            arr[303, 3] = "XS2";
            arr[303, 4] = "XS3";
            arr[303, 5] = "XS4";
            arr[303, 6] = "XS5";
            arr[303, 7] = "XS6";
            arr[303, 8] = "XS7";
            arr[303, 9] = "XS8";
            arr[303, 10] = "XS9";
            arr[303, 11] = "XS10";
            //Шапка 8
            arr[308, 0] = "№";
            arr[308, 1] = "XS4";
            arr[308, 2] = "XS1";
            arr[308, 3] = "XS2";
            arr[308, 4] = "XS3";
            arr[308, 5] = "XS4";
            arr[308, 6] = "XS5";
            arr[308, 7] = "XS6";
            arr[308, 8] = "XS7";
            arr[308, 9] = "XS8";
            arr[308, 10] = "XS9";
            arr[308, 11] = "XS10";
            //Шапка 9
            arr[317, 0] = "№";
            arr[317, 1] = "XS6";
            arr[317, 2] = "XS1";
            arr[317, 3] = "XS2";
            arr[317, 4] = "XS3";
            arr[317, 5] = "XS4";
            arr[317, 6] = "XS5";
            arr[317, 7] = "XS6";
            arr[317, 8] = "XS7";
            arr[317, 9] = "XS8";
            arr[317, 10] = "XS9";
            arr[317, 11] = "XS10";
            //Шапка 10
            arr[321, 0] = "№";
            arr[321, 1] = "XS1";
            //arr[321, 2] = "XS1";
            //arr[321, 3] = "XS2";
            arr[321, 4] = "XS3";
            //Шапка 11
            arr[323, 0] = "№";
            arr[323, 1] = "XS10";
            arr[323, 2] = "XS1";
            arr[323, 6] = "XS5";
            arr[323, 11] = "XS10";
            //Шапка 12
            arr[328, 0] = "№";
            arr[328, 1] = "XS3";
            arr[328, 2] = "XP3";
            //Шапка 13
            arr[330, 0] = "№";
            arr[330, 1] = "XP1";
            arr[330, 2] = "XS10";
            //Шапка 14
            arr[332, 0] = "№";
            arr[332, 1] = "XS4";
            arr[332, 2] = "XS5";
            //Шапка 15
            arr[334, 0] = "№";
            arr[334, 1] = "XP1";
            arr[334, 2] = "XS10";
            //Шапка 16
            arr[337, 0] = "№";
            arr[337, 1] = "XS1";
            arr[337, 2] = "XS3";

            bool[] successColor = new bool[data.Length + 3 + 16];
            for (int i = 0; i < successColor.Length; i++)
            {
                successColor[i] = false;
            }
            int k = 0;
            int headRow = 0;
            int posColunm = 0;
            for (int i = 3; i < data.Length + 3 + 16; i++)
            {
                if (arr[i, 0] != "№")
                {
                    arr[i, 0] = data[k].Index.ToString();
                    var com = data[k].Comment.Split(',');
                    var com1 = com[0].Split(' ');
                    var com2 = com[1].Split(' ');
                    arr[i, 1] = com1[1].ToUpper() + "/K" + data[k].Input.Channel + " " + data[k].Input.Device.ToUpper();

                    for (int j = 2; j < 12; j++)
                    {
                        if (com2[1] == arr[headRow, j])
                        {
                            posColunm = j;
                            break;
                        }
                    }

                    arr[i, posColunm] = com2[2].ToUpper() + "/K" + data[k].Output.Channel + data[k].Output.Device.ToUpper();
                    arr[i, 12] = data[k].Result;
                    arr[i, 13] = data[k].Value;
                    if (data[k].Result == "PASSED")
                    {
                        successColor[i] = true;
                    }
                    k++;
                }
                else
                {
                    headRow = i;
                }
            }

            try
            {
                MSExcel.Range r = xlWorkSheet.Range[xlWorkSheet.Cells[1, 1], xlWorkSheet.Cells[data.Length + 3 + 16, 14]];
                r.Value = arr;
                xlWorkSheet.Columns.EntireColumn.AutoFit();

                for (int i = 4; i < successColor.Length; i++)
                {
                    if (successColor[i])
                    {
                        xlWorkSheet.Cells[i + 1, 13].Font.Color = MSExcel.XlRgbColor.rgbGreen;
                        xlWorkSheet.Cells[i + 1, 14].Font.Color = MSExcel.XlRgbColor.rgbGreen;
                    }

                    else
                    {
                        xlWorkSheet.Cells[i + 1, 13].Font.Color = MSExcel.XlRgbColor.rgbRed;
                        xlWorkSheet.Cells[i + 1, 14].Font.Color = MSExcel.XlRgbColor.rgbRed;
                    }
                }

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

        public static string SaveDaqWrapper(DAQTest[] data, string path)
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
            string[,] arr = new string[data.Length + 2, 8];
            for (int i = 0; i < 7; i++)
                arr[0, i] = " ";
            arr[0, 1] = "Входы";
            arr[0, 3] = "Выходы";
            arr[1, 0] = "№";
            arr[1, 1] = "Устройство";
            arr[1, 2] = "Канал";
            arr[1, 3] = "Устройство";
            arr[1, 4] = "Канал";
            arr[1, 5] = "Комментарий";
            arr[1, 6] = "Значение";
            arr[0, 7] = " ";
            arr[1, 7] = "Результат";

            int k = 0;
            bool[] color = new bool[data.Length];
            for (int i = 2; i < data.Length + 2; i++)
            {
                arr[i, 0] = data[k].Index.ToString();
                arr[i, 1] = data[k].Input.Device;
                arr[i, 2] = data[k].Input.Channel.ToString();
                arr[i, 3] = data[k].Output.Device;
                arr[i, 4] = data[k].Output.Channel.ToString();
                arr[i, 5] = data[k].Comment;
                arr[i, 6] = data[k].Value;
                arr[i, 7] = data[k].Result;
                if (data[k].Result == "PASSED")
                    color[k] = true;
                k++;
            }
            try
            {
                MSExcel.Range r = xlWorkSheet.Range[xlWorkSheet.Cells[1, 1], xlWorkSheet.Cells[data.Length + 2, 8]];
                r.Value = arr;
                xlWorkSheet.Columns.EntireColumn.AutoFit();

                for (int i = 0; i < color.Length; i++)
                {
                    if (color[i])
                    {
                        xlWorkSheet.Cells[i + 3, 7].Font.Color = MSExcel.XlRgbColor.rgbGreen;
                        xlWorkSheet.Cells[i + 3, 8].Font.Color = MSExcel.XlRgbColor.rgbGreen;
                    }
                    else
                    {
                        xlWorkSheet.Cells[i + 3, 7].Font.Color = MSExcel.XlRgbColor.rgbRed;
                        xlWorkSheet.Cells[i + 3, 8].Font.Color = MSExcel.XlRgbColor.rgbRed;
                    }
                }

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
