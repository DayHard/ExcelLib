using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;

namespace MyClass.WriteToExcel
{
    public abstract class ExcelFileWriter<T>
    {
        private Excel.Application _excelApplication = null;
        private Excel.Workbooks _workBooks = null;
        private Excel._Workbook _workBook = null;
        private object _value = Missing.Value;
        private Excel.Sheets _excelSheets = null;
        private Excel._Worksheet _excelSheet = null;
        private Excel.Range _excelRange = null;
        private Excel.Font _excelFont = null;
        /// <summary>
        /// User have to set the names of header in the derived class
        /// </summary>
        public abstract object[] Headers { get;}
        /// <summary>
        /// user have to parse the data from the list and pass each data along with the
        /// column and row name to the base fun, FillExcelWithData().
        /// </summary>
        /// <param name="list"></param>
        public abstract void FillRowData(List<T> list);

        /// <summary>
        /// get the data of object which will be saved to the excel sheet
        /// </summary>
        public abstract object[,] ExcelData { get;}
        /// <summary>
        /// get the no of columns
        /// </summary>
        public abstract int ColumnCount { get;}
        /// <summary>
        /// get the now of rows to fill
        /// </summary>
        public abstract int RowCount { get;}

        /// <summary>
        /// user can override this to make the headers not be in bold.
        /// by default it is true
        /// </summary>
        protected virtual bool BoldHeaders
        {
            get
            {
                return true;
            }
        }

        /// <summary>
        /// api through which data from the list can be write to an excel
        /// kind of a Template Method Pattern is used here
        /// </summary>
        /// <param name="fileName"></param>
        /// <param name="holdingsList"></param>
        public void WriteDateToExcel(string fileName, List<T> list, string startColumn, string endColumn)
        {
            this.ActivateExcel();                        
            this.FillRowData(list);
            this.FillExcelWithData();
            this.FillHeaderColumn(Headers, startColumn, endColumn);
            this.SaveExcel(fileName);
        }
        
        /// <summary>
        /// activate the excel application
        /// </summary>
        protected virtual void ActivateExcel()
        {
            _excelApplication = new Excel.Application();
            _workBooks = (Excel.Workbooks)_excelApplication.Workbooks;
            _workBook = (Excel._Workbook)(_workBooks.Add(_value));
            _excelSheets = (Excel.Sheets)_workBook.Worksheets;
            _excelSheet = (Excel._Worksheet)(_excelSheets.get_Item(1)); 
        }

        /// <summary>
        /// fill the header columns for the range specified and make it bold if specified
        /// </summary>
        /// <param name="headers"></param>
        /// <param name="startColumn"></param>
        /// <param name="endColumn"></param>
        protected void FillHeaderColumn(object[] headers, string startColumn, string endColumn)
        {                       
            _excelRange = _excelSheet.get_Range(startColumn, endColumn);
            _excelRange.set_Value(_value, headers);
            if (BoldHeaders == true)
            {
                this.BoldRow(startColumn, endColumn);
            }
            _excelRange.EntireColumn.AutoFit();
        }
        /// <summary>
        /// Fill the excel sheet with data along with the position specified
        /// </summary>
        /// <param name="columnrow"></param>
        /// <param name="data"></param>
        private void FillExcelWithData()
        {
            _excelRange = _excelSheet.get_Range("A1", _value);
            _excelRange = _excelRange.get_Resize(RowCount + 1, ColumnCount);
            _excelRange.set_Value(Missing.Value, ExcelData);
            _excelRange.EntireColumn.AutoFit();
        }
        /// <summary>
        /// save the excel sheet to the location with file name
        /// </summary>
        /// <param name="fileName"></param>
        protected virtual void SaveExcel(string fileName)
        {
            _workBook.SaveAs(fileName, _value, _value,
                _value, _value, _value, Excel.XlSaveAsAccessMode.xlNoChange,
                _value, _value, _value, _value, null);
            _workBook.Close(false, _value, _value);
            _excelApplication.Quit();
        }
        /// <summary>
        /// make the range of rows bold
        /// </summary>
        /// <param name="row1"></param>
        /// <param name="row2"></param>
        private void BoldRow(string row1, string row2)
        {            
            _excelRange = _excelSheet.get_Range(row1, row2);
            _excelFont = _excelRange.Font;
            _excelFont.Bold = true;
        }
    }
  }
