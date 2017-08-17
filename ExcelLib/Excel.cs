using System;
using System.Data;
using System.IO;
using System.Linq;
using NPOI.SS.UserModel;

namespace ExcelLib
{
    public static class Excel
    {
        public static EData1And2[] EData { get; set; }
        public static EData3[] EData3 { get; set; }
        public static EData4[] EData4 { get; set; }
        public static EData5[] EData5 { get; set; }

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

        //Парсит Excel1 и Excel2
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
        public static EData5[] ParseEx5(string path)
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


                EData5 = new EData5[lenght];
                for (int i = 0; i < EData5.Length; i++)
                {
                    EData5[i] = new EData5();
                }
                int k = 0;
                for (int i = 0; i < table.Rows.Count; i++)
                {
                    int value;
                    if (list[0, i] != String.Empty && int.TryParse(list[0, i], out value))
                    {
                        EData5[k].Index = value;
                        EData5[k].Max = Double.NegativeInfinity; //list[3, i]
                        EData5[k].Min = Double.NegativeInfinity; //list[4, i]
                        EData5[k].Value = Double.NegativeInfinity; //list[5, i]
                        EData5[k].Comment = list[6, i] +" "+ list[7, i];

                        var indata = list[1, i].Split('r');
                        var indata2 = indata[0].Split('k');
                        EData5[k].Input.Device = "R" + indata[1];
                        EData5[k].Input.Channel = Convert.ToInt32(indata2[1]);

                        var outdata = list[2, i].Split('r');
                        var outdata2 = outdata[0].Split('k');
                        EData5[k].Output.Device = "R" + outdata[1];
                        EData5[k].Output.Channel = Convert.ToInt32(outdata2[1]);
                        k++;
                    }
                }
                return EData5;
            }
            catch
            {
                return null;
            }
        }
    }
}

