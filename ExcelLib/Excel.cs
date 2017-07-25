using System;
using System.IO;
using System.Runtime.InteropServices;

namespace ExcelLib
{
    public static class Excel
    {
        public static EData1And2[] EData1 { get; set; }
        public static EData1And2[] EData2 { get; set; }
        public static EData3[] EData3 { get; set; }
        public static EData4[] EData4 { get; set; }


        //private static void ParseWrapperEx1And2(string path)
        //{
        //    Microsoft.Office.Interop.Excel.Application objWorkExcel = new Microsoft.Office.Interop.Excel.Application(); //открыть эксель
        //    Microsoft.Office.Interop.Excel.Workbook objWorkBook = objWorkExcel.Workbooks.Open(path, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing); //открыть файл
        //    Microsoft.Office.Interop.Excel.Worksheet objWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)objWorkBook.Sheets[1]; //получить 1 лист
        //    Microsoft.Office.Interop.Excel.Range lastCell = null;
        //    try
        //    {
        //        lastCell = objWorkSheet.Cells.SpecialCells(Microsoft.Office.Interop.Excel.XlCellType
        //            .xlCellTypeLastCell); //1 ячейку
        //        string[,] list =
        //            new string[lastCell.Column, lastCell.Row]; // массив значений с листа равен по размеру листу
        //        for (int i = 0; i < lastCell.Column; i++) //по всем колонкам
        //        for (int j = 0; j < lastCell.Row; j++) // по всем строкам
        //            list[i, j] = objWorkSheet.Cells[j + 1, i + 1].Text.ToString(); //считываем текст в строку

        //        EData1 = new EData1And2[lastCell.Row - 1];
        //        for (int i = 0; i < EData1.Length; i++)
        //        {
        //            EData1[i] = new EData1And2();
        //        }

        //        //Отображение
        //        for (int i = 1; i < lastCell.Row; i++) //по всем колонкам
        //        {
        //            if (list[0, i].Length == list[0, i].Count(char.IsDigit))
        //            {
        //                EData1[i - 1].Index = Convert.ToInt32(list[0, i]);
        //            }
        //            string[] inputData = list[1, i].Split('R');
        //            string[] outputData = list[2, i].Split('R');
        //            if (inputData.Length == 2)
        //            {
        //                string[] inData = inputData[0].Split('K');
        //                EData1[i - 1].Input.Channel = Convert.ToInt32(inData[1]);
        //                EData1[i - 1].Input.Device = "R" + inputData[1];
        //            }
        //            if (outputData.Length == 2)
        //            {
        //                string[] outData = outputData[0].Split('K');
        //                EData1[i - 1].Output.Channel = Convert.ToInt32(outData[1]);
        //                EData1[i - 1].Output.Device = "R" + outputData[1];
        //            }
        //            EData1[i - 1].Comment = Convert.ToString(list[3, i]);

        //            // Если хотябы одно поле null зануляем текущую строку
        //            if (EData1[i - 1].Index <= 0 || EData1[i - 1].Comment == String.Empty || EData1[i - 1].Input.Channel == 0
        //                || EData1[i - 1].Input.Device == String.Empty || EData1[i - 1].Output.Channel == 0 ||
        //                EData1[i - 1].Output.Device == String.Empty)
        //            {
        //                EData1[i - 1].Index = 0;

        //                EData1[i - 1].Input.Channel = 0;
        //                EData1[i - 1].Input.Device = null;
        //                EData1[i - 1].Output.Channel = 0;
        //                EData1[i - 1].Output.Device = null;
        //                EData1[i - 1].Comment = null;
        //            }
        //        }
        //    }
        //    finally
        //    {
        //        objWorkBook.Close();
        //        objWorkExcel.Quit();

        //        Marshal.ReleaseComObject(objWorkSheet);
        //        Marshal.ReleaseComObject(objWorkBook);
        //        Marshal.ReleaseComObject(objWorkExcel);
        //        if (lastCell != null) Marshal.ReleaseComObject(lastCell);

        //    }
        //}
        ////Вызываемый метод парса таблицы типа 1 и 2
        //public static object ParseEx1And2(string path)
        //{
        //    try
        //    {
        //        if (!File.Exists(path))
        //            return null;

        //        ParseWrapperEx1And2(path);

        //    }
        //    finally
        //    {
        //        GC.Collect();
        //        GC.WaitForPendingFinalizers();
        //    }
        //    return EData1;
        //}
        private static void ParseWrapperEx1And2(string path)
        {
            Microsoft.Office.Interop.Excel.Application objWorkExcel = new Microsoft.Office.Interop.Excel.Application(); //открыть эксель
            Microsoft.Office.Interop.Excel.Workbook objWorkBook = objWorkExcel.Workbooks.Open(path, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing); //открыть файл
            Microsoft.Office.Interop.Excel.Worksheet objWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)objWorkBook.Sheets[1]; //получить 1 лист
            Microsoft.Office.Interop.Excel.Range lastCell = null;
            //Microsoft.Office.Interop.Excel.Range lastCell = objWorkSheet.get_Range("R1", "R340C14");
            try
            {
                lastCell = objWorkSheet.Cells.SpecialCells(Microsoft.Office.Interop.Excel.XlCellType
                    .xlCellTypeLastCell); //1 ячейку
                string[,] list =
                    new string[lastCell.Column, lastCell.Row]; // массив значений с листа равен по размеру листу
                for (int i = 0; i < lastCell.Column; i++) //по всем колонкам
                    for (int j = 0; j < lastCell.Row; j++) // по всем строкам
                        list[i, j] = objWorkSheet.Cells[j + 1, i + 1].Text.ToString(); //считываем текст в строку

                int lenght = 0;
                for (int j = 0; j < lastCell.Row; j++)
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
                string[] tabHeader = new string[lastCell.Column];                      
                for (int j = 0; j < lastCell.Row; j++)
                {
                    if (list[0, j] != "№")
                    {
                        var value = 0;
                        if (int.TryParse(list[0, j], out value))
                        {
                            EData1[k].Index = value;

                            for (int i = 1; i < lastCell.Column; i++)
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
                        for (int i = 0; i < lastCell.Column; i++)
                        {
                            if (list[i,j] != String.Empty)
                            {
                                tabHeader[i] = list[i, j];
                            }
                        }
                        nPossition = j;
                    }
                }

            }
            catch(Exception ex)
            {
                EData1 = null;
            }
            finally
            {
                objWorkBook.Close();
                objWorkExcel.Quit();

                Marshal.ReleaseComObject(objWorkSheet);
                Marshal.ReleaseComObject(objWorkBook);
                Marshal.ReleaseComObject(objWorkExcel);
                if (lastCell != null) Marshal.ReleaseComObject(lastCell);

            }
        }
        //Вызываемый метод парса таблицы типа 1 и 2
        public static object ParseEx1And2(string path)

        {
            try
            {
                if (!File.Exists(path))
                    return null;

                ParseWrapperEx1And2(path);

            }
            catch
            {
                return null;
            }
            finally
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            return EData1;
        }

    }
}
