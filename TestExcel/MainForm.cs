using System;
using System.Windows.Forms;
using ExcelLib;


namespace TestExcel
{
    public partial class MainForm : Form
    {
        //private string path = @"C:\Users\Destiny\Desktop\тесты готово\7064\1_2.xls";
        //private string path2 = @"C:\Users\Destiny\Desktop\тесты готово\7064\2.xls";
        //private string path3 = @"C:\Users\Destiny\Desktop\тесты готово\7064\3_1.xls";
        //private string path4 = @"C:\Users\Destiny\Desktop\тесты готово\7064\4.xls";
        //private string path = @"C:\Users\Destiny\Desktop\тесты готово\7194\1 и 2.xls";
        //private string path3 = @"C:\Users\Destiny\Desktop\тесты готово\7194\3.xls";
        //private string path3 = @"C:\Users\Destiny\Desktop\тесты готово\7194\3_1.xls";
        private string path4 = @"C:\Users\Destiny\Desktop\тесты готово\7194\#4.xls";
        public MainForm()
        {
            InitializeComponent();
            //btnAutoClick_Click(null, null);
        }

        private void btnClickMe_Click(object sender, EventArgs e)
        {
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                 var data2 = Excel.ParseBPPP(openFileDialog.FileName);
                 var status = Excel.SaveBPPP(data2, @"C:\Users\Destiny\Desktop\1.xlsx");             
            }
        }

        private void btnAutoClick_Click(object sender, EventArgs e)
        {
            //var data  = Excel.ParseEx(path);
            //var data2 = Excel.ParseEx3(path3);
           // var data4 = Excel.ParseEx4(path4);           
        }
    }
}
