using System;
using System.Windows.Forms;
using ExcelLib;

namespace TestExcel
{
    public partial class MainForm : Form
    {
        private string path = @"c:\\Users\\Destiny\\Desktop\\a.xls";
        public MainForm()
        {
            InitializeComponent();
        }

        private void btnClickMe_Click(object sender, EventArgs e)
        {
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                var data = Excel.Parse(openFileDialog.FileName);
            }
        }

        private void btnAutoClick_Click(object sender, EventArgs e)
        {
            var data = Excel.Parse(path);
        }
    }
}
