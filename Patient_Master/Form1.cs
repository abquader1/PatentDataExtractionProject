using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;


namespace Patient_Master
{
    public partial class frmMainWindow : Form
    {
        Excel.Application xlApp=null;
        Word.Application wdApp=null;
        public frmMainWindow()
        {
            InitializeComponent();
        }

        private void frmMainWindow_Load(object sender, EventArgs e)
        {
            this.AcceptButton = btnProcess;
            txtExcelFile.Focus();
            txtExcelFile.Text = @"D:\Freelancer\Freelancer\Patient_Master\resultsExample.xls";
            txtWordFile.Text = @"D:\Freelancer\Freelancer\Patient_Master\ReportTemplate.docx";
        }

        private void btnProcess_Click(object sender, EventArgs e)
        {
            clsScrapper cScrapper = new clsScrapper();
            Console.Write(Assembly.GetExecutingAssembly().Location);
            clsMap oMap=  cScrapper.ScrapData(Application.ExecutablePath, "https://patents.google.com/patent/US20200088579?oq=patent:US20200088579","123.png");
            Excel.Workbook wk=null;
            Excel.Worksheet sh=null;
            try
            {
                if (string.IsNullOrEmpty(txtExcelFile.Text))
                {
                    MessageBox.Show("Please select result excel file");
                    txtExcelFile.Focus();
                    return;
                }
                if (string.IsNullOrEmpty(txtWordFile.Text))
                {
                    MessageBox.Show("Please select Word Template file");
                    txtWordFile.Focus();
                    return;
                }
                xlApp = new Excel.Application();
                wk = xlApp.Workbooks.Open(txtExcelFile.Text);
                sh = wk.Worksheets[1];
                Excel.Range rng = sh.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
                rng = sh.get_Range("A1", rng);
                object[,] data = rng.Value2;
                for (int r = 0; r < data.GetLength(0); r++)
                {
                    for (int c = 0; c < data.GetLength(1); c++)
                    {

                    }
                }
                
                wk.Close(false);
                xlApp.Quit();
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            catch (Exception)
            {

                throw;
            }
            finally
            {
                Marshal.ReleaseComObject(wk);
                Marshal.ReleaseComObject(sh);
                Marshal.ReleaseComObject(xlApp);
            }
        }

        private void btnBrowseExcel_Click(object sender, EventArgs e)
        {
            OpenFileDialog fdlg = new OpenFileDialog();
            fdlg.Title = "Excel File Dialog";
            fdlg.InitialDirectory = @"c:\";
            fdlg.Filter = "All files (*.*)|*.*|All files (*.*)|*.*";
            fdlg.FilterIndex = 2;
            fdlg.RestoreDirectory = true;
            if (fdlg.ShowDialog() == DialogResult.OK)
            {
                txtExcelFile.Text = fdlg.FileName;
            }
        }

        private void btnBrowseWord_Click(object sender, EventArgs e)
        {
            OpenFileDialog fdlg = new OpenFileDialog();
            fdlg.Title = "Word File Dialog";
            fdlg.InitialDirectory = @"c:\";
            fdlg.Filter = "All files (*.*)|*.*|All files (*.*)|*.*";
            fdlg.FilterIndex = 2;
            fdlg.RestoreDirectory = true;
            if (fdlg.ShowDialog() == DialogResult.OK)
            {
                txtWordFile.Text = fdlg.FileName;
            }
            
        }
    }
}
