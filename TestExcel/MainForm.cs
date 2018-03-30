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
            //IndTest[] IndTest = null;
            //if (openFileDialog.ShowDialog() == DialogResult.OK)
            //{
            //    IndTest = Excel.ParseInd(openFileDialog.FileName);
            //}
            ////if (saveFileDialog.ShowDialog() == DialogResult.OK)
            //{
            //    var status = Excel.SaveBPPP(IndTest, saveFileDialog.FileName);
            //}

            BPPPTest[] BPPPTest = null;
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                BPPPTest = Excel.ParseBPPP(openFileDialog.FileName);
            }
            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                var status = Excel.SaveBPPP(BPPPTest, saveFileDialog.FileName);
            }

            //DAQTest[] DAQTest = null;
            //if (openFileDialog.ShowDialog() == DialogResult.OK)
            //{
            //    DAQTest = Excel.ParseDAQ(openFileDialog.FileName);
            //}
            //if (saveFileDialog.ShowDialog() == DialogResult.OK)
            //{
            //    var status = Excel.SaveDAQ(DAQTest, saveFileDialog.FileName);
            //}
        }

        private void btnAutoClick_Click(object sender, EventArgs e)
        {
            //var data  = Excel.ParseEx(path);
            //var data2 = Excel.ParseEx3(path3);
           // var data4 = Excel.ParseEx4(path4);           
        }
    }
}
