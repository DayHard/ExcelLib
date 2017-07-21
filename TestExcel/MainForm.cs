using System;
using System.Windows.Forms;
using ExcelLib;

namespace TestExcel
{
    public partial class MainForm : Form
    {
        //private string path = @"c:\\Users\\Destiny\\Desktop\\a.xls";
        private string path = @"C:\Users\Destiny\Desktop\тесты готово\7064\1_1.xlsx";
        public MainForm()
        {
            InitializeComponent();
        }

        private void btnClickMe_Click(object sender, EventArgs e)
        {
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                var data = Excel.ParseEx1And2(openFileDialog.FileName);
            }
        }

        private void btnAutoClick_Click(object sender, EventArgs e)
        {
            var data = Excel.ParseEx1And2(path);
        }
    }
}
