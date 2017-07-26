using System;
using System.Windows.Forms;
using ExcelLib;

namespace TestExcel
{
    public partial class MainForm : Form
    {
        //private string path = @"C:\Users\Destiny\Desktop\тесты готово\7064\1_2.xls";
        //private string path2 = @"C:\Users\Destiny\Desktop\тесты готово\7064\2.xls";
        private string path3 = @"C:\Users\Destiny\Desktop\тесты готово\7064\3_1.xls";
        //private string path4 = @"C:\Users\Destiny\Desktop\тесты готово\7064\4.xls";
        public MainForm()
        {
            InitializeComponent();
        }

        private void btnClickMe_Click(object sender, EventArgs e)
        {
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                //var data = Excel.ParseEx1And2(openFileDialog.FileName);
            }
        }

        private void btnAutoClick_Click(object sender, EventArgs e)
        {
            //var data  = Excel.ParseEx1(path);

            //var data2 = Excel.ParseEx1(path2);

            var data3 = Excel.ParseEx3(path3);
        }
    }
}
