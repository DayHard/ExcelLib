using System;
using System.IO;
using System.Runtime.InteropServices;

namespace ExcelLib
{
    public static class Excel
    {
        public static ExcelData[] EData;
        public static object Parse(string path)
        {
            if (!File.Exists(path))
                return null;

            Microsoft.Office.Interop.Excel.Application objWorkExcel = new Microsoft.Office.Interop.Excel.Application(); //открыть эксель
            Microsoft.Office.Interop.Excel.Workbook objWorkBook = objWorkExcel.Workbooks.Open(path, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing); //открыть файл
            Microsoft.Office.Interop.Excel.Worksheet objWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)objWorkBook.Sheets[1]; //получить 1 лист
            Microsoft.Office.Interop.Excel.Range lastCell = null;
            try
            {
                lastCell = objWorkSheet.Cells.SpecialCells(Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeLastCell); //1 ячейку
                string[,] list =
                    new string[lastCell.Column, lastCell.Row]; // массив значений с листа равен по размеру листу
                for (int i = 0; i < lastCell.Column; i++) //по всем колонкам
                for (int j = 0; j < lastCell.Row; j++) // по всем строкам
                    list[i, j] = objWorkSheet.Cells[j + 1, i + 1].Text.ToString(); //считываем текст в строку
                EData = new ExcelData[lastCell.Row - 1];
                for (int i = 0; i < EData.Length; i++)
                {
                    EData[i] = new ExcelData();
                }
                //Отображение
                for (int i = 1; i < lastCell.Row; i++) //по всем колонкам
                {
                    EData[i - 1].Index = Convert.ToInt32(list[0, i]);
                    string[] inputData = list[1, i].Split('R');
                    string[] outputData = list[2, i].Split('R');
                    EData[i - 1].Input.Channel = inputData[0];
                    EData[i - 1].Input.Device = "R" + inputData[1];
                    EData[i - 1].Output.Channel = outputData[0];
                    EData[i - 1].Output.Device = "R" + outputData[1];
                    EData[i - 1].Comment = Convert.ToString(list[3, i]);
                }
            }
            catch
            {
                return null;
            }
            finally
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();

                objWorkBook.Close();
                objWorkExcel.Quit();

                Marshal.ReleaseComObject(objWorkSheet);
                Marshal.ReleaseComObject(objWorkBook);
                Marshal.ReleaseComObject(objWorkExcel);
                if (lastCell != null) Marshal.ReleaseComObject(lastCell);
            }
            return EData;
        }

        //private static void CloseProcess()
        //{
        //    Process[] list;
        //    list = Process.GetProcessesByName("EXCEL");
        //    foreach (Process proc in list)
        //    {
        //        proc.Kill();
        //    }
        //}
    }
}
