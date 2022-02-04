using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;
using Microsoft.WindowsAPICodePack.Dialogs;
using System.Text.RegularExpressions;
using System.Diagnostics;

namespace Patient_Master
{
    public partial class frmMainWindow : Form
    {
        int ssPoint = 0;
        int eePoint = 0;
        Excel.Application xlApp=null;
        Word.Application wdApp=null;
        LogWriter logger = new LogWriter();
        public frmMainWindow()
        {
            InitializeComponent();
        }

        private void frmMainWindow_Load(object sender, EventArgs e)
        {
            logger.LogWrite(":frmMainWindow_Load: Begin");
            this.AcceptButton = btnProcess;
            btnClear_Click(sender, e);
            txtExcelFile.Focus();
            txtExcelFile.Text = Properties.Settings.Default.excelPath;
            txtWordFile.Text = Properties.Settings.Default.wordPath;
            txtOutputPath.Text = Properties.Settings.Default.outputPath;
            logger.LogWrite(":frmMainWindow_Load: Ends");
            //txtExcelFile.Text = @"D:\Freelancer\Freelancer\Patient_Master\test.xls";
            //txtWordFile.Text = @"D:\Freelancer\Freelancer\Patient_Master\test.docx";
            //txtOutputPath.Text = @"D:\Freelancer\Freelancer\Patient_Master\output";

        }

        private void btnProcess_Click(object sender, EventArgs e)
        {
            logger.LogWrite(":btnProcess_Click: Begin");

            Excel.Workbook wk=null;
            Excel.Worksheet sh=null;
            Word.Document doc = null;
            bool isInsideError = false;
            clsScrapper cScrapper = null;
            try
            {
                lblStatus.Text = "Processing....";
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
                if (string.IsNullOrEmpty(txtOutputPath.Text))
                {
                    MessageBox.Show("Please select Output Path");
                    txtOutputPath.Focus();
                    return;
                }
                int numberOfRecords=0;
                if (!string.IsNullOrEmpty(txtNumberOfRecords.Text))
                {
                    if(int.TryParse(txtNumberOfRecords.Text,out numberOfRecords))
                    {
                        //must be positive
                        if (numberOfRecords < 1)
                        {
                            MessageBox.Show("Record count cannot be less than 1");
                            txtNumberOfRecords.Focus();
                            return;
                        }
                    }
                    else
                    {
                        MessageBox.Show("Please enter valid record count");
                        txtNumberOfRecords.Focus();
                        return;
                    }
                }
                else
                {
                    numberOfRecords = 1000000;
                }
                if(chkExcel.Checked==false && chkWord.Checked == false)
                {
                    MessageBox.Show("Please select atleast one output type");
                    chkExcel.Focus();
                    return;
                }
                Properties.Settings.Default.excelPath = txtExcelFile.Text;
                Properties.Settings.Default.Save();
                Properties.Settings.Default.wordPath = txtWordFile.Text;
                Properties.Settings.Default.Save();
                Properties.Settings.Default.outputPath = txtOutputPath.Text;
                Properties.Settings.Default.Save();

                Dictionary<string, string> assigneeVsPatents = new Dictionary<string, string>();
                Dictionary<string, string> termsVsPatents = new Dictionary<string, string>();
                
                if (chkWord.Checked == true)
                {
                    logger.LogWrite(":btnProcess_Click: Initializing Word Process");
                    wdApp = new Word.Application();
                    wdApp.Visible = false;
                    wdApp.DisplayAlerts = Word.WdAlertLevel.wdAlertsNone;
                    doc = wdApp.Documents.Open(txtWordFile.Text);
                    logger.LogWrite(":btnProcess_Click: Document " + doc.Name + " Opened successfully!!!");
                }

                //CopySection("bkpatentIdStart", "bkpatentIdEnd", doc, "2. Testing");
                logger.LogWrite(":btnProcess_Click: Initializing Excel input");
                xlApp = new Excel.Application();
                xlApp.Visible = false;
                xlApp.DisplayAlerts = false;
                wk = xlApp.Workbooks.Open(txtExcelFile.Text);
                logger.LogWrite(":btnProcess_Click: Excel " + wk.Name + " Opened successfully!!!");
                sh = wk.Worksheets[1];
                //Excel.Worksheet shTmp = wk.Worksheets.Add();
                //shTmp.Name = "Tmp";
                //sh.Cells.Copy(shTmp.Range["A1"]);
                //Excel.Range rng = sh.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
                //rng = sh.get_Range("A1", rng);
                //rng.RemoveDuplicates(1, Excel.XlYesNoGuess.xlYes);

                Excel.Range rng = sh.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
                rng = sh.get_Range("A1", rng);
                object[,] data = rng.Value2;
                logger.LogWrite(":btnProcess_Click: Total Rows " + data.GetLength(0));
                logger.LogWrite(":btnProcess_Click: Total Columns " + data.GetLength(1));
                string fileName = wk.Name.Replace(".xls","").Replace(".xlsx","");
                int lastRow = 0;
                int availableRecords = 0;

                //List<termsMap> lMap = new List<termsMap>();

                //termsVsPatents.Add("US10560643", "image,tissue,hyperspectral");
                //termsVsPatents.Add("US9117133", "image,tissue,hyperspectral");
                //termsVsPatents.Add("US20090326383", "image,tissue,hyperspectral");
                //termsVsPatents.Add("US20160080665", "image,tissue,hyperspectral");
                //termsVsPatents.Add("US20200267336", "image,tissue,hyperspectral");
                //termsVsPatents.Add("US20200088579", "hybrid,spectral,image");
                //termsVsPatents.Add("US20180310828", "optical,image,mode");
                //termsVsPatents.Add("US20150110381", "cellular,classify,histopathology");
                //termsVsPatents.Add("US20170319073", "optical,image,mode");
                //termsVsPatents.Add("US9717417", "optical,image,mode");
                //termsVsPatents.Add("US20170079530", "optical,image,mode");
                //termsVsPatents.Add("US9962090", "optical,classify,histopathology");
                //int n = 1;
                //foreach (KeyValuePair<string, string> item in termsVsPatents)
                //{
                //    termsMap tMap = new termsMap();
                //    tMap.terms = item.Value;
                //    tMap.patent_id = item.Key;
                //    if (n > 5)
                //    {
                //        tMap.cpc = "Cpc_2";
                //    }
                //    else
                //    {
                //        tMap.cpc = "Cpc_1";
                //    }
                //    n++;
                //    lMap.Add(tMap);
                //}
                //assigneeVsPatents.Add("US10560643", "Andrew Bodkin");
                //assigneeVsPatents.Add("US9117133", "Barnes Donald Michael");
                //assigneeVsPatents.Add("US20090326383", "Carl Pennypacker");
                //assigneeVsPatents.Add("US20160080665", "Chad Leverette");
                //assigneeVsPatents.Add("US20200267336", "Clean Earth Technologies, Llc");
                //assigneeVsPatents.Add("US20200088579", "Dalton William S");
                //assigneeVsPatents.Add("US20180310828", "Dermaspect, Llc");
                //assigneeVsPatents.Add("US20150110381", "Frank Geshwind");
                //assigneeVsPatents.Add("US20170319073", "Andrew Bodkin");
                //assigneeVsPatents.Add("US9717417", "Frank Geshwind");
                //assigneeVsPatents.Add("US20170079530", "Dalton William S");
                //assigneeVsPatents.Add("US9962090", "Carl Pennypacker");
                //assigneeVsPatents.Add("US10776606", "Frank Geshwind");
                //assigneeVsPatents.Add("US20190200905", "Dalton William S");

                //WriteChartInWord(doc, assigneeVsPatents, lMap);

                if (numberOfRecords == 0)
                    lastRow = data.GetLength(0);
                else
                {
                    //Number of records should not be more than avaialbe one
                    if (numberOfRecords > data.GetLength(0))
                    {
                        lastRow = data.GetLength(0);
                    }
                    else
                    {
                        for (int i = 1; i <= data.GetLength(0); i++)
                        {
                            if (data[i, 2].ToString().ToLower().Trim() == data[i + 1, 2].ToString().ToLower().Trim())
                                continue;
                            else
                            {
                                availableRecords++;
                                if (availableRecords == numberOfRecords)
                                {
                                    lastRow = i + 2;
                                    break;
                                }
                            }
                        }
                    }
                }
                
                Excel.Workbook wkOutput = null;
                Excel.Worksheet shOutput = null;
                List<string> lstControls = new List<string>();
                if (chkAbstract.Checked)
                    lstControls.Add("Title");
                if(chkAssignee.Checked)
                    lstControls.Add("Assignee");
                //if(chkClaims.Checked)
                //    lstControls.Add("Claims");
                if(chkCPC.Checked)
                    lstControls.Add("CPC");
                if(chkCPCName.Checked)
                    lstControls.Add("CPC Name");
                //if(chkDescription.Checked)
                //    lstControls.Add("Description");
                if(chkStatus.Checked)
                    lstControls.Add("Status");
                if(chkFilingDate.Checked)
                    lstControls.Add("Filing Date");

                lstControls.Add("Patent Link");
                lstControls.Add("Date Of Anticipated Expiration");

                int oColNum = 3;
                if (chkExcel.Checked == true)
                {
                    logger.LogWrite(":btnProcess_Click: Adding output workbook");
                    wkOutput = xlApp.Workbooks.Add();
                    shOutput = wkOutput.Worksheets[1];
                    shOutput.Range["A1"].Value = "Family ID";
                    shOutput.Range["B1"].Value = "Patent ID";

                    foreach (string Colname in lstControls)
                    {
                        shOutput.Cells[1, oColNum].Value = Colname;
                        oColNum++;
                    }
                    oColNum--;
                    //shOutput.Range["B1"].Value = "Status";
                    //shOutput.Range["C1"].Value = "Assignee";
                    //shOutput.Range["D1"].Value = "CPC";
                    //shOutput.Range["E1"].Value = "CPC Name";
                    //shOutput.Range["F1"].Value = "Abstract";
                    //shOutput.Range["G1"].Value = "Claims";
                    //shOutput.Range["H1"].Value = "Filing Date";
                    //shOutput.Range["I1"].Value = "Description";
                }
                string inputQuery = data[1, 2].ToString();
                string filingDate = string.Empty;
                string title = string.Empty;
                string patentId = string.Empty;
                cScrapper = new clsScrapper(Application.StartupPath);

                //only Word
                if (chkWord.Checked==true)
                    FillBookMark("bkInputStart", "bkInputEnd", doc, inputQuery,0);

                bool firstOccurance = false;
                int outputCounter = 2;
                int lastColumn = data.GetLength(1);
                logger.LogWrite(":btnProcess_Click: Total Rows to be processed: " + lastRow);
                List<termsMap> lsTermsMap = new List<termsMap>();
                int recordPreserver = 0;
                bool firsttime = false;
                int oCounter = 2;
                //Loop for each row
                for (int r = 3; r <= lastRow; r++)
                {                    
                    termsMap termsMap = new termsMap();
                    logger.LogWrite(":btnProcess_Click: Processing Row No. " + r);
                    patentId = data[r, 1].ToString();
                    termsMap.patent_id = patentId;
                    logger.LogWrite(":btnProcess_Click: Processing Patent: " + patentId);
                    title = data[r, 2].ToString();

                    string BASE_URL = string.Format("https://patents.google.com/patent/{0}?oq=patent:{1}",patentId,patentId);                    
                    string filename = DateTime.Now.ToString("yyyyMMddHHmmss") + "_" + r.ToString() + ".png";
                    clsMap oMap = null;
                    logger.LogWrite(":btnProcess_Click: Invoking Scrapper for: " + BASE_URL);                    
                    if (chkWord.Checked)
                        oMap = cScrapper.ScrapData(BASE_URL, filename,true);
                    else
                        oMap = cScrapper.ScrapData(BASE_URL, filename, false);

                    //Check incase empty map
                    if(oMap.Abstract== "Unable to load Url")
                    {
                        //Write in Excel here
                        if (chkExcel.Checked == true)
                        {
                            //Family Id
                            if (data[r, 7] != null)
                            {
                                shOutput.Range["A" + oCounter].Value = data[r, 7].ToString();
                            }
                            shOutput.Range["B" + oCounter].Value = patentId;
                            oCounter++;
                            continue;
                        }
                    }
                    logger.LogWrite(":btnProcess_Click: Scrapping finished");
                    if (data[r, 5] != null)
                        filingDate = data[r, 5].ToString();

                    StringBuilder stitles = new StringBuilder();
                    recordPreserver = r;
                    if (chkWord.Checked == true)
                    {
                        logger.LogWrite(":btnProcess_Click: Writing Header info begin");
                        if (firstOccurance == false)
                        {
                            FillBookMark("bkpatentIdStart", "bkpatentIdEnd", doc, patentId + "-" + title, 1);
                            firstOccurance = true;
                        }
                        else
                        {
                            CopySection("bkpatentIdStart", "bkpatentIdEnd", doc, patentId + "-" + title, outputCounter - 1);
                        }
                        logger.LogWrite(":btnProcess_Click: Writing header info end");
                        stitles = new StringBuilder();
                        if(!assigneeVsPatents.ContainsKey(patentId))
                            assigneeVsPatents.Add(patentId, oMap.Current_Assignee);
                        for (int i = r; i <= lastRow; i++)
                        {
                            if (i == lastRow)
                            {
                                break;
                            }
                            if (data[i + 1, 2] != null)
                            {
                                if (title.ToLower() == data[i + 1, 2].ToString().ToLower())
                                {
                                    r++;
                                    stitles.Append(data[r, 1].ToString() + " ");
                                    if (!assigneeVsPatents.ContainsKey(patentId))
                                        assigneeVsPatents.Add(data[r, 1].ToString(), oMap.Current_Assignee);
                                }
                                else
                                {
                                    break;
                                }
                            }
                        }

                        if (stitles.Length > 0)
                        {
                            stitles.Length--;                           
                            WordWriteOperation(doc, "Additional Family Patents: " + stitles.ToString(), true);
                        }
                        StringBuilder trms = new StringBuilder();
                        StringBuilder graphTerms = new StringBuilder();
                        if (data.GetLength(1) > 10)
                        {
                            if (data[r, 11] != null)
                            {
                                trms.Append( data[r, 11].ToString() + ",");
                                graphTerms.Append(data[r, 11].ToString() + ",");
                            }
                        }
                        if (data.GetLength(1) > 11)
                        {
                            if (data[r, 12] != null)
                            {
                                trms.Append( data[r, 12].ToString() + ",");
                                graphTerms.Append(data[r, 12].ToString() + ",");
                            }
                        }
                        if (data.GetLength(1) > 12)
                        {
                            if (data[r, 13] != null)
                            {
                                trms.Append(data[r, 13].ToString() + ",");
                                graphTerms.Append(data[r, 13].ToString() + ",");
                            }
                        }
                        if (data.GetLength(1) > 13)
                        {
                            if (data[r, 14] != null)
                            {
                                trms.Append(data[r, 14].ToString() + ",");
                                graphTerms.Append(data[r, 14].ToString() + ",");
                            }
                        }
                        if (data.GetLength(1) > 14)
                        {
                            if (data[r, 15] != null)
                            {
                                trms.Append(data[r, 15].ToString());
                                graphTerms.Append(data[r, 15].ToString() + ",");
                            }
                        }
                        if (graphTerms.Length > 0)
                        {
                            graphTerms.Length--;
                            termsMap.terms = graphTerms.ToString();
                            //.Add(patentId, graphTerms.ToString());
                        }
                        else
                        {
                            termsMap.terms = string.Empty;
                            //termsVsPatents.Add(patentId, string.Empty);
                        }
                        if (trms.Length > 0)
                        {
                            //Remove last char if ,
                            if (trms.ToString().Trim().EndsWith(","))
                                trms.Length--;
                            string oTermText = trms.ToString().Replace(",", ", ");
                            ColorFont(doc, oTermText);
                        }

                        WordWriteOperation(doc, "Current Assignee: " + oMap.Current_Assignee, true);
                        WordWriteOperation(doc, "Status: " + oMap.status, true);
                        WordWriteOperation(doc, "Filing Date: " + filingDate, true);
                        WordWriteOperation(doc, "Classification: " + oMap.Classfication, true);
                        WordWriteOperation(doc, BASE_URL, true, true);
                        if(File.Exists(Path.GetTempPath() + filename))
                            InsertPicture(doc, Path.GetTempPath() + filename);
                        WordWriteOperation(doc, "Abstract: " + oMap.Abstract, true);

                        //Only first claim needs to be written
                        if (chkClaimsRequired.Checked == true && firsttime==false)
                        {
                            
                            WordWriteOperation(doc, "Claims: " + oMap.Claim, true);
                            firsttime = true;
                        }
                        else if(chkClaimsRequired.Checked==false)
                        {
                            WordWriteOperation(doc, "Claims: " + oMap.Claim, true);
                        }
                        FindAndHighlight(doc, inputQuery);
                        FindAndHighlight(doc, inputQuery);
                    }
                    //Only Excel
                    if (chkExcel.Checked == true)
                    {                        
                        string cpc = string.Empty;
                        string cpcName = string.Empty;
                        try
                        {
                            cpc = oMap.Classfication.Split(' ')[0];
                            cpcName = oMap.Classfication.Replace(cpc, "");
                        }
                        catch (Exception)
                        {
                        }
                        for (int i = recordPreserver; i <= r; i++)
                        {

                            if (data[i, 2] != null)
                            {
                                //Family Id
                                if (data[i, 7] != null)
                                {
                                    shOutput.Range["A" + oCounter].Value = data[i, 7].ToString();
                                }
                                //Patent Id
                                if (data[i, 1] != null)
                                {                                    
                                    shOutput.Range["B" + oCounter].Value = data[i, 1].ToString();   
                                }
                                for (int k = 3; k <= oColNum; k++)
                                {
                                    string header = (string)(shOutput.Cells[1, k] as Excel.Range).Text;
                                    if (!string.IsNullOrEmpty(header))
                                    {
                                        if (header == "Title")
                                            shOutput.Cells[oCounter, k].Value = title;
                                        else if (header == "Assignee")
                                            shOutput.Cells[oCounter, k].Value = oMap.Current_Assignee;
                                        //else if (header == "Claims")
                                        //    shOutput.Cells[oCounter, k].Value = oMap.Claim;
                                        //else if (header == "Description")
                                        //shOutput.Cells[oCounter, k].Value = oMap.Description;
                                        else if (header == "Status")
                                            shOutput.Cells[oCounter, k].Value = oMap.status;
                                        else if (header == "Filing Date")
                                        {
                                            shOutput.Cells[oCounter, k].Value = filingDate;
                                            shOutput.Cells[oCounter, k].NumberFormat = "mm/dd/yyyy;@";
                                        }
                                        else if (header == "CPC")
                                        {
                                            if (!string.IsNullOrEmpty(cpc))
                                                shOutput.Cells[oCounter, k].Value = cpc;
                                        }
                                        else if (header == "CPC Name")
                                        {
                                            if (!string.IsNullOrEmpty(cpcName))
                                                shOutput.Cells[oCounter, k].Value = cpcName;
                                        }
                                        else if (header == "Patent Link")
                                            shOutput.Cells[oCounter, k].Value = BASE_URL;
                                        else if (header == "Date Of Anticipated Expiration")
                                            shOutput.Cells[oCounter, k].Value = oMap.anticipationExpiry;
                                    }
                                }
                                oCounter++;
                            }
                        }
                        
                        //shOutput.Range["C" + outputCounter].Value = oMap.Current_Assignee;
                        //shOutput.Range["B" + outputCounter].Value = oMap.status;
                        //shOutput.Range["H" + outputCounter].Value = filingDate;
                        //shOutput.Range["I" + outputCounter].Value = " " + oMap.Description;
                        if (!string.IsNullOrEmpty(cpc))
                        {
                            //shOutput.Range["D" + outputCounter].Value = cpc;
                            if (cpc.Length >= 4)
                                termsMap.cpc = new string(cpc.Take(4).ToArray());
                            else
                                termsMap.cpc = cpc;
                        }
                        //if (!string.IsNullOrEmpty(cpcName))
                        //    shOutput.Range["E" + outputCounter].Value = cpcName;
                        //shOutput.Range["F" + outputCounter].Value = " " + oMap.Abstract;
                        //shOutput.Range["G" + outputCounter].Value = " " + oMap.Claim;
                    }
                    lsTermsMap.Add(termsMap);
                    outputCounter++;
                }
                //word only
                if (chkWord.Checked == true)
                {
                    FindAndBold(doc, "Additional Family Patents: ");
                    FindAndBold(doc, "Current Assignee: ");
                    FindAndBold(doc, "Status: ");
                    FindAndBold(doc, "Filing Date: ");
                    FindAndBold(doc, "Classification: ");
                    FindAndBold(doc, "Abstract: ", true);
                    if(chkClaimsRequired.Checked==true)
                        FindAndBold(doc, "Claims: ", true);
                    doc.Content.Select();
                    doc.ActiveWindow.Selection.Font.Name = "Calibri (Body)";
                    //Write Chart data in word Assignee vs Patents
                    WriteChartInWord(doc, assigneeVsPatents, lsTermsMap);
                    logger.LogWrite(":btnProcess_Click: Saving Word document begin");
                    string fName = fileName + ".docx";
                    doc.SaveAs2(Path.Combine(txtOutputPath.Text, fName), Word.WdSaveFormat.wdFormatDocumentDefault);
                    logger.LogWrite(":btnProcess_Click: Saving Word document end");
                    doc.Close();
                    wdApp.Quit();
                }
                //only Excel
                if (chkExcel.Checked == true)
                {
                    oCounter--;
                    shOutput.Range["A2:A" + oCounter].Rows.RowHeight = 25;                    
                    logger.LogWrite(":btnProcess_Click: Saving Output Excel begin");
                    string xName = fileName + ".xlsx";
                    wkOutput.SaveAs(Path.Combine(txtOutputPath.Text, xName), Excel.XlFileFormat.xlWorkbookDefault);
                    logger.LogWrite(":btnProcess_Click: Saving Output Excel end");
                    logger.LogWrite(":btnProcess_Click: Copying columns B,C,D,E,H at the end");
                    shOutput.Range["B1:G" + oCounter].Copy(sh.Cells[2, lastColumn + 1]);
                    shOutput.Range["I1:I" + oCounter].Copy(sh.Cells[2, lastColumn + 7]);
                    //Color rows
                    for (int i = 2; i <= oCounter; i++)
                    {
                        Random r = new Random();
                        int iNum;                        
                        iNum = r.Next(-30, 50);

                        string txt = (string)(shOutput.Range["A" + i] as Excel.Range).Text;
                        if (!string.IsNullOrEmpty(txt))
                        {
                            shOutput.Range["A" + i].EntireRow.Interior.Color = 15132391;
                        }
                        else
                        {
                            shOutput.Range["A" + i].EntireRow.Interior.Color = 15132391 + Convert.ToInt32((iNum / (int)Math.Abs(iNum)));
                        }
                    }
                    wkOutput.Close();
                }

                wk.Close(true);
                GC.Collect();
                GC.WaitForPendingFinalizers();
                lblStatus.Text = "Success!!!";
                MessageBox.Show("Data processed successfully");
            }
            catch (Exception ex)
            {
                //MessageBox.Show("Processing error: Please contact infoextract pvt ltd or write us at \n infoextractpvtltd@gmail.com");
                MessageBox.Show("Processing error: " + ex.Message);
                logger.LogWrite(":btnProcess_Click: Error: " + ex.Message);                
                if (chkWord.Checked == true)
                    if(wdApp!=null)
                        wdApp.Visible = true;
                if (chkExcel.Checked == true)
                    if(xlApp!=null)
                        xlApp.Visible = true;
                isInsideError = true;
                return;
            }
            finally
            {
                if (cScrapper != null)
                    cScrapper.CloseSession();

                if (isInsideError == false)
                {
                    if (xlApp != null)
                        xlApp.Quit();
                    if (wk != null)
                        Marshal.ReleaseComObject(wk);
                    if (sh != null)
                        Marshal.ReleaseComObject(sh);
                    if (xlApp != null)
                        Marshal.ReleaseComObject(xlApp);

                    KillProcess("winword");
                    KillProcess("excel");
                }
            }
        }
        private void KillProcess(string processName)
        {
            foreach (Process p in System.Diagnostics.Process.GetProcessesByName(processName))
            {
                try
                {
                    p.Kill();
                    p.WaitForExit(); // possibly with a timeout
                }
                catch (Win32Exception winException)
                {
                    // process was terminating or can't be terminated - deal with it
                }
                catch (InvalidOperationException invalidException)
                {
                    // process has already exited - might be able to let this one go
                }
            }
        }
        
        private void WriteChartInWord(Word.Document doc, Dictionary<string, string> dataAssignee, List<termsMap> termsPatentMap)
        {
            logger.LogWrite(":WriteChartInWord: begin");
            Excel.Workbook wkCharts = xlApp.Workbooks.Open(Path.Combine(Application.StartupPath, "Graphs.xlsx"));
            Excel.Worksheet sh = wkCharts.Worksheets["Graph"];
            int counter = 0;
            int lastRow = 0;
            Word.Selection oSelection = null;
            if (dataAssignee.Count > 0)
            {
                try
                {


                    logger.LogWrite(":WriteChartInWord: Assignee chart start");
                    sh.Range["A2:B10000"].ClearContents();
                    sh.Range["K2:L10000"].ClearContents();
                    counter = 2;
                    foreach (KeyValuePair<string, string> item in dataAssignee)
                    {
                        sh.Range["A" + counter].Value = item.Value;
                        sh.Range["K" + counter].Value = item.Value;
                        counter++;
                    }
                    counter--;
                    if (counter >= 2)
                    {
                        sh.Range["K2:K" + counter].RemoveDuplicates(1, Excel.XlYesNoGuess.xlNo);
                        sh.Range["L2"].Formula = "=COUNTIF($A$2:$A$10000,K2)";
                        lastRow = sh.Range["K10000"].End[Excel.XlDirection.xlUp].Row;

                        sh.Range["L2"].Copy();
                        sh.Range["L3:L" + lastRow].PasteSpecial(Excel.XlPasteType.xlPasteFormulas);
                        sh.Calculate();

                        sh.Range["L2:L" + lastRow].Copy();
                        sh.Range["L2:L" + lastRow].PasteSpecial(Excel.XlPasteType.xlPasteValues);

                        sh.Range["A2:B10000"].ClearContents();

                        sh.Range["K2:L" + lastRow].Copy(sh.Range["A2"]);

                        Excel.ChartObject chartObject2 = (Excel.ChartObject)sh.ChartObjects("assignee");
                        Excel.Chart chart = chartObject2.Chart;
                        chart.SetSourceData(sh.Range["A1:B" + lastRow]);

                        MoveEnd(doc);
                        oSelection = wdApp.Selection;
                        oSelection.Paragraphs.Add();
                        MoveEnd(doc);
                        oSelection = wdApp.Selection;
                        chart.ChartArea.Copy();
                        oSelection.PasteAndFormat(Word.WdRecoveryType.wdChartPicture);
                        logger.LogWrite(":WriteChartInWord: Assignee chart ends");
                    }
                    else
                    {
                        logger.LogWrite(":WriteChartInWord: No Assignee found");
                    }
                }
                catch (Exception ex)
                {

                    logger.LogWrite(":WriteChartInWord: Assignee Chart: " + ex.Message);
                }
            }
            if (termsPatentMap.Count > 0)
            {
                try
                {
                    logger.LogWrite(":WriteChartInWord: Term Chart starts");
                    sh = wkCharts.Worksheets["Graph1"];
                    termsPatentMap = termsPatentMap.OrderBy(x => x.cpc).ToList();
                    sh.Cells.ClearContents();
                    //write all cpcs
                    counter = 1;
                    List<string> lstCpcs = new List<string>();
                    foreach (termsMap item in termsPatentMap)
                    {
                        lstCpcs.Add(item.cpc);
                    }

                    foreach (string cpc_item in lstCpcs.Distinct().ToList())
                    {
                        string currCPC = cpc_item;
                        counter = 2;
                        
                        List<string> lstTerms = new List<string>();
                        foreach (termsMap item in termsPatentMap.Where(x => x.cpc == currCPC).ToList())
                        {
                            if (!string.IsNullOrEmpty(item.terms))
                            {
                                foreach (string trms in item.terms.Split(','))
                                {                                    
                                    lstTerms.Add(trms.Trim());
                                }
                            }                            
                            sh.Range["B" + counter].Value = item.patent_id; 
                            counter++;
                        }                                                

                        //at least one term require
                        if (counter >= 1)
                        {
                            int colCounter = 3;
                            List<string> uniqueTerms = lstTerms.Distinct().ToList();
                            //Write terms
                            foreach (string oTerm in uniqueTerms)
                            {
                                if (!string.IsNullOrEmpty(oTerm))
                                {
                                    sh.Cells[counter, colCounter].Value = oTerm;
                                    colCounter++;
                                }
                            }

                            colCounter--;
                            lastRow = sh.Range["C10000"].End[Excel.XlDirection.xlUp].Row;                            
                            Excel.Range rng = sh.Range[sh.Cells[2, 2], sh.Cells[lastRow+1, colCounter]];
                            rng.EntireColumn.AutoFit();
                            sh.Range["B1"].Value = currCPC;
                            sh.Range["B1"].Font.Bold = true;

                            Excel.Range rngMerge = sh.Range[sh.Cells[1, 2], sh.Cells[1, colCounter]];
                            rngMerge.Merge(Missing.Value);
                            rngMerge.Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                            rngMerge.Cells.VerticalAlignment = Excel.XlVAlign.xlVAlignTop;

                            for (int row = 2; row <= lastRow; row++)
                            {
                                string oPatent = (string)(sh.Range["B" + row] as Excel.Range).Text;
                                if (!string.IsNullOrEmpty(oPatent))
                                {
                                    for (int col = 3; col <= colCounter; col++)
                                    {
                                        string key = (string)(sh.Cells[lastRow, col] as Excel.Range).Text;
                                        if (!string.IsNullOrEmpty(key))
                                        {
                                            var result=termsPatentMap.Where(x => x.patent_id == oPatent && x.terms.Contains(key)).FirstOrDefault();
                                            if (result != null)
                                            {
                                                sh.Cells[row, col].Value = ".";
                                                sh.Cells[row, col].Font.Size = 36;
                                                sh.Cells[row, col].Font.Color = ColorTranslator.ToOle(Color.Blue);
                                            }
                                        }

                                    }
                                }
                            }
                            sh.Range["A2:A" + lastRow].Rows.RowHeight=20;
                           
                            //Put term heading
                            sh.Cells[lastRow + 1, 3].Value = "Terms";
                            sh.Cells[lastRow + 1, 3].Font.Italic = true;
                            rngMerge = sh.Range[sh.Cells[lastRow + 1, 3], sh.Cells[lastRow + 1, colCounter]];
                            rngMerge.Merge(Missing.Value);
                            rngMerge.Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                            rngMerge.Font.Bold = true;
                            rngMerge.Font.Italic = true;
                            rngMerge.Font.Size = 16;

                            //Put Patent heading
                            sh.Range["A1"].Value = "Patents";
                            sh.Range["A1:A" + lastRow].Merge(Missing.Value);
                            sh.Range["A1:A" + lastRow].Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                            sh.Range["A1:A" + lastRow].Cells.VerticalAlignment = Excel.XlVAlign.xlVAlignBottom;
                            sh.Range["A1:A" + lastRow].Cells.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                            sh.Range["A1:A" + lastRow].Cells.Orientation = 90;
                            sh.Range["A1:A" + lastRow].Cells.Font.Bold = true;
                            sh.Range["A1:A" + lastRow].Cells.Font.Italic = true;
                            sh.Range["A1:A" + lastRow].Cells.Font.Size = 16;

                            //Border
                            rng = sh.Range[sh.Cells[2, 3], sh.Cells[lastRow - 1, colCounter]];
                            rng.Borders[Excel.XlBordersIndex.xlDiagonalDown].LineStyle = Excel.XlLineStyle.xlLineStyleNone;
                            rng.Borders[Excel.XlBordersIndex.xlDiagonalUp].LineStyle = Excel.XlLineStyle.xlLineStyleNone;

                            rng.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                            rng.Borders[Excel.XlBordersIndex.xlEdgeTop].ColorIndex = 0;
                            rng.Borders[Excel.XlBordersIndex.xlEdgeTop].ThemeColor = 7;
                            rng.Borders[Excel.XlBordersIndex.xlEdgeTop].TintAndShade = 0.399945066682943f;
                            rng.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThin;

                            rng.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                            rng.Borders[Excel.XlBordersIndex.xlEdgeLeft].ColorIndex = 0;
                            rng.Borders[Excel.XlBordersIndex.xlEdgeLeft].ThemeColor = 7;
                            rng.Borders[Excel.XlBordersIndex.xlEdgeLeft].TintAndShade = 0.399945066682943f;
                            rng.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;

                            rng.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                            rng.Borders[Excel.XlBordersIndex.xlEdgeBottom].ColorIndex = 0;
                            rng.Borders[Excel.XlBordersIndex.xlEdgeBottom].ThemeColor = 7;
                            rng.Borders[Excel.XlBordersIndex.xlEdgeBottom].TintAndShade = 0.399945066682943f;
                            rng.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;

                            rng.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                            rng.Borders[Excel.XlBordersIndex.xlEdgeRight].ColorIndex = 0;
                            rng.Borders[Excel.XlBordersIndex.xlEdgeRight].ThemeColor = 7;
                            rng.Borders[Excel.XlBordersIndex.xlEdgeRight].TintAndShade = 0.399945066682943f;
                            rng.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;

                            rng.Borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle = Excel.XlLineStyle.xlContinuous;
                            rng.Borders[Excel.XlBordersIndex.xlInsideVertical].ColorIndex = 0;
                            rng.Borders[Excel.XlBordersIndex.xlInsideVertical].ThemeColor = 7;
                            rng.Borders[Excel.XlBordersIndex.xlInsideVertical].TintAndShade = 0.399945066682943f;
                            rng.Borders[Excel.XlBordersIndex.xlInsideVertical].Weight = Excel.XlBorderWeight.xlThin;

                            rng.Borders[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle = Excel.XlLineStyle.xlContinuous;
                            rng.Borders[Excel.XlBordersIndex.xlInsideHorizontal].ColorIndex = 0;
                            rng.Borders[Excel.XlBordersIndex.xlInsideHorizontal].ThemeColor = 7;
                            rng.Borders[Excel.XlBordersIndex.xlInsideHorizontal].TintAndShade = 0.399945066682943f;
                            rng.Borders[Excel.XlBordersIndex.xlInsideHorizontal].Weight = Excel.XlBorderWeight.xlThin;

                            sh.Cells.Rows[1].RowHeight = 21;
                            sh.Cells.Columns[1].ColumnWidth = 3.29f;
                            sh.Cells.Rows[lastRow + 1].RowHeight = 21;

                            rng = sh.Range[sh.Cells[1, 1], sh.Cells[lastRow + 1, colCounter]];
                            rng.CopyPicture(Excel.XlPictureAppearance.xlScreen, Excel.XlCopyPictureFormat.xlPicture);

                            MoveEnd(doc);
                            oSelection = wdApp.Selection;

                            int inlineShapesCount = doc.InlineShapes.Count;
                            //Add new page
                            oSelection.InsertBreak(Word.WdBreakType.wdPageBreak);
                            MoveEnd(doc);
                            oSelection = wdApp.Selection;                            
                            oSelection.Paste();
                            //Select shape
                            Word.InlineShape oShape = doc.InlineShapes[inlineShapesCount + 1];
                            oShape.Width = 515f;
                            sh.Range["A1:AZ10000"].Clear();
                            rngMerge.UnMerge();
                        }
                        else
                        {
                            logger.LogWrite(":WriteChartInWord: No Terms");
                        }
                    }
                    logger.LogWrite(":WriteChartInWord: Term Chart End");
                }
                catch (Exception ex)
                {

                    logger.LogWrite(":WriteChartInWord: Term Chart: " + ex.Message);
                }
            }

            wkCharts.Close(false);
            Marshal.ReleaseComObject(wkCharts);
            logger.LogWrite(":WriteChartInWord: ends");
        }
        private void WriteChartInWord(Word.Document doc, Dictionary<string,string> dataAssignee, Dictionary<string, string> dataTerms)
        {
            logger.LogWrite(":WriteChartInWord: begin");
            Excel.Workbook wkCharts=xlApp.Workbooks.Open(Path.Combine(Application.StartupPath , "Graphs.xlsx"));
            Excel.Worksheet sh = wkCharts.Worksheets["Graph"];
            int counter = 0;
            int lastRow = 0;
            Word.Selection oSelection = null;
            if (dataAssignee.Count > 0)
            {
                logger.LogWrite(":WriteChartInWord: Assignee chart start");
                sh.Range["A2:B10000"].ClearContents();
                sh.Range["K2:L10000"].ClearContents();
                counter = 2;
                foreach (KeyValuePair<string, string> item in dataAssignee)
                {
                    sh.Range["A" + counter].Value = item.Value;
                    sh.Range["K" + counter].Value = item.Value;
                    counter++;
                }
                counter--;
                if (counter >= 2)
                {
                    sh.Range["K2:K" + counter].RemoveDuplicates(1, Excel.XlYesNoGuess.xlNo);
                    sh.Range["L2"].Formula = "=COUNTIF($A$2:$A$10000,K2)";
                    lastRow = sh.Range["K10000"].End[Excel.XlDirection.xlUp].Row;

                    sh.Range["L2"].Copy();
                    sh.Range["L3:L" + lastRow].PasteSpecial(Excel.XlPasteType.xlPasteFormulas);
                    sh.Calculate();

                    sh.Range["L2:L" + lastRow].Copy();
                    sh.Range["L2:L" + lastRow].PasteSpecial(Excel.XlPasteType.xlPasteValues);

                    sh.Range["A2:B10000"].ClearContents();

                    sh.Range["K2:L" + lastRow].Copy(sh.Range["A2"]);

                    Excel.ChartObject chartObject2 = (Excel.ChartObject)sh.ChartObjects("assignee");
                    Excel.Chart chart = chartObject2.Chart;
                    chart.SetSourceData(sh.Range["A1:B" + lastRow]);

                    MoveEnd(doc);
                    oSelection = wdApp.Selection;
                    oSelection.Paragraphs.Add();
                    MoveEnd(doc);
                    oSelection = wdApp.Selection;
                    chart.ChartArea.Copy();
                    oSelection.PasteAndFormat(Word.WdRecoveryType.wdChartPicture);
                    logger.LogWrite(":WriteChartInWord: Assignee chart ends");
                }
                else
                {
                    logger.LogWrite(":WriteChartInWord: No Assignee found");
                }
            }
            if (dataTerms.Count > 0)
            {
                logger.LogWrite(":WriteChartInWord: Term Chart starts");
                sh = wkCharts.Worksheets["Graph1"];
                sh.Cells.ClearContents();
                counter = 1;
                int colCounter = 2;
                foreach (KeyValuePair<string, string> item in dataTerms)
                {
                    if (!string.IsNullOrEmpty(item.Value))
                    {
                        foreach (string trms in item.Value.Split(','))
                        {
                            sh.Range["A" + counter].Value = trms;
                            counter++;
                        }
                    }
                    sh.Cells[1, colCounter].Value = item.Key;
                    colCounter++;
                }
                colCounter--;
                counter--;
                //at least one term require
                if (counter >= 1 && colCounter>=2)
                {
                    sh.Range["A1:A" + counter].RemoveDuplicates(1, Excel.XlYesNoGuess.xlNo);
                    sh.Range["A1"].EntireRow.Insert();
                    lastRow = sh.Range["A10000"].End[Excel.XlDirection.xlUp].Row;
                    lastRow++;

                    sh.Range[sh.Cells[2, 2], sh.Cells[2, colCounter]].Cut(sh.Range["B" + lastRow]);

                    Excel.Range rng = sh.Range[sh.Cells[1, 1], sh.Cells[lastRow, colCounter]];
                    rng.EntireColumn.AutoFit();

                    for (int col = 2; col <= colCounter; col++)
                    {
                        string key = sh.Cells[lastRow, col].Value;
                        if (!string.IsNullOrEmpty(key))
                        {
                            string vals = dataTerms[key];
                            for (int row = 2; row <= lastRow; row++)
                            {
                                foreach (var item in vals.Split(','))
                                {
                                    if (item == sh.Range["A" + row].Value)
                                    {
                                        sh.Cells[row, col].Value = ".";
                                        sh.Cells[row, col].Font.Size = 25;
                                    }
                                }
                            }
                        }
                    }
                    rng.CopyPicture(Excel.XlPictureAppearance.xlScreen, Excel.XlCopyPictureFormat.xlPicture);
                    MoveEnd(doc);
                    oSelection = wdApp.Selection;
                    oSelection.Paragraphs.Add();
                    MoveEnd(doc);
                    oSelection = wdApp.Selection;
                    oSelection.Paste();
                }
                else
                {
                    logger.LogWrite(":WriteChartInWord: No Terms");
                }
                logger.LogWrite(":WriteChartInWord: Term Chart End");
            }
            wkCharts.Close(false);
            Marshal.ReleaseComObject(wkCharts);
            logger.LogWrite(":WriteChartInWord: ends");
        }

        private void btnBrowseExcel_Click(object sender, EventArgs e)
        {
            logger.LogWrite(":btnBrowseExcel_Click: begin");
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
            logger.LogWrite(":btnBrowseExcel_Click: ends");
        }

        private void btnBrowseWord_Click(object sender, EventArgs e)
        {
            logger.LogWrite(":btnBrowseWord_Click: begin");
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
            logger.LogWrite(":btnBrowseWord_Click: ends");
        }
        void InsertPicture(Word.Document doc,string oPath)
        {
            logger.LogWrite(":InsertPicture: begin");
            MoveEnd(doc);
            
            Word.Selection oSelection = wdApp.Selection;

            var oShape = oSelection.InlineShapes.AddPicture(oPath,false,true);
            //oShape.Width = doc.PageSetup.PageWidth;
            MoveEnd(doc);
            oSelection = wdApp.Selection;
            oSelection.Paragraphs.Add();
            MoveEnd(doc);
            logger.LogWrite(":InsertPicture: ends");
            //oShape.Width = 193;
            //oShape.Height = 25;

        }
        void ColorFont(Word.Document doc, string oText)
        {
            logger.LogWrite(":ColorFont: begin");
            MoveEnd(doc);
            Word.Selection oSelection = wdApp.Selection;
            int sP = oSelection.Range.Start;
            int eP = sP + oText.Length;
            oSelection.Text = oText;
            Word.Range rng = doc.Range(sP, eP);
            rng.Select();
            rng.Font.Bold= Convert.ToInt32(true);
            rng.Font.Color = Word.WdColor.wdColorRed;
            MoveEnd(doc);
            oSelection = wdApp.Selection;
            oSelection.Paragraphs.Add();
            MoveEnd(doc);
            oSelection = wdApp.Selection;
            oSelection.set_Style(Word.WdBuiltinStyle.wdStyleNormal);
            logger.LogWrite(":ColorFont: ends");
        }
        void NextArticle(Word.Document doc,string oText, int num)
        {
            logger.LogWrite(":NextArticle: begin");
            Word.Selection oSelection = wdApp.Selection;
            oSelection.EndKey(Word.WdUnits.wdStory);
            oSelection = wdApp.Selection;
            oSelection.Paragraphs.Add();
            oSelection = wdApp.Selection;

            int sPoint = oSelection.Range.Start;
            int ePoint = sPoint + oText.Length;
            oSelection.Text = oText;
            MoveEnd(doc);
            oSelection = wdApp.Selection;
            oSelection.Paragraphs.Add();
            Word.Range rng = doc.Range(sPoint, ePoint);
            rng.Select();

            //Word.ListTemplate oTemplate =Word.ListGalleries[Word.WdListGalleryType.wdNumberGallery].ListTemplates[1];
            //Word.ListLevel oLevel = oTemplate.ListLevels[1];
            //oLevel.NumberFormat = "%1.";
            //oLevel.TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            //oLevel.NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            //oLevel.NumberPosition = Word.InchesToPoints(0.25f);
            //oLevel.TextPosition = Utility.WdApplication.InchesToPoints(0.5f);
            //oLevel.TabPosition = 9999999;
            //oLevel.ResetOnHigher = 0;
            //oLevel.StartAt = 1;
            //oSelection.Range.ListFormat.ApplyListTemplateWithLevel(oTemplate, false, Word.WdListApplyTo.wdListApplyToWholeList, Word.WdDefaultListBehavior.wdWord10ListBehavior);

            rng.ListFormat.ApplyListTemplateWithLevel(doc.ListTemplates[1], true, Word.WdListApplyTo.wdListApplyToWholeList, Word.WdDefaultListBehavior.wdWord10ListBehavior);
            MoveEnd(doc);
            oSelection = wdApp.Selection;
            oSelection.set_Style(Word.WdBuiltinStyle.wdStyleNormal);
            logger.LogWrite(":NextArticle: ends");
        }
        void CopySection(string bkSName, string bkEName, Word.Document doc, string oText, int articleNumber)
        {
            logger.LogWrite(":CopySection: begin");
            logger.LogWrite(":CopySection: Taking bookmark reference");
            System.Threading.Thread.Sleep(2000);
            Word.Selection oSelection = wdApp.Selection;
            Word.Bookmark bmkS = doc.Bookmarks[bkSName];
            Word.Bookmark bmkE = doc.Bookmarks[bkEName];
            logger.LogWrite(":CopySection: Bookmark taken");
            oSelection.EndKey(Word.WdUnits.wdStory);
            System.Threading.Thread.Sleep(2000);
            oSelection = wdApp.Selection;
            oSelection.Paragraphs.Add();
            logger.LogWrite(":CopySection: Paragraph added");
            string articleTitle = "  " + articleNumber.ToString() + ".   " + oText;
            logger.LogWrite(":CopySection: article title build");
            int sPoint = oSelection.Range.Start;
            int ePoint = sPoint + articleTitle.Length;
            oSelection.Text = articleTitle;
            logger.LogWrite(":CopySection: article title populated");
            Word.Range rng = doc.Range(bmkS.Start, bmkE.Start);
            rng.Select();
            logger.LogWrite(":CopySection: range selected");
            System.Threading.Thread.Sleep(2000);
            oSelection = wdApp.Selection;
            oSelection.CopyFormat();
            logger.LogWrite(":CopySection: format copied");
            Word.Range targetRange = doc.Range(sPoint, ePoint);
            targetRange.Select();
            logger.LogWrite(":CopySection: target range selected");
            System.Threading.Thread.Sleep(2000);
            oSelection = wdApp.Selection;
            oSelection.PasteFormat();
            logger.LogWrite(":CopySection: format pasted");

            oSelection.EndOf(Word.WdUnits.wdStory, Word.WdMovementType.wdMove);
            logger.LogWrite(":CopySection: Move end");
            System.Threading.Thread.Sleep(2000);
            oSelection = wdApp.Selection;
            oSelection.Paragraphs.Add();
            logger.LogWrite(":CopySection: Paragraph added");
            oSelection.EndOf(Word.WdUnits.wdStory, Word.WdMovementType.wdMove);
            logger.LogWrite(":CopySection: Move end");
            System.Threading.Thread.Sleep(2000);
            oSelection = wdApp.Selection;
            oSelection.set_Style(Word.WdBuiltinStyle.wdStyleNormal);
            logger.LogWrite(":CopySection: ends");
        }
        public void FillBookMark(string bkSName, string bkEName, Word.Document doc, string oText, int articleNumber)
        {
            logger.LogWrite(":FillBookMark: begin");
            string aTitle = string.Empty;
            if (articleNumber > 0)
                aTitle = "  " + articleNumber.ToString() + ".   " + oText;
            else
                aTitle = oText;
            Word.Bookmark bmkS = doc.Bookmarks[bkSName];
            Word.Bookmark bmkE = doc.Bookmarks[bkEName];
            doc.Range(bmkS.Start, bmkE.Start).Text = aTitle;
            logger.LogWrite(":FillBookMark: ends");
        }
        void MoveEnd(Word.Document doc)
        {
            Word.Selection oSelection = wdApp.Selection;
            oSelection.EndOf(Word.WdUnits.wdStory,Word.WdMovementType.wdMove);
        }
        void WordWriteOperation(Word.Document doc, string oText, bool newPara, bool isHyperLink=false)
        {
            logger.LogWrite(":WordWriteOperation: begin");
            MoveEnd(doc);
            Word.Selection oSelection = wdApp.Selection;
            
            if (isHyperLink)
            {
                Word.Hyperlinks oLinks = doc.Hyperlinks;
                object wAddress = oText;
                object wSubAddress = oText;
                int sPoint = oSelection.Range.Start;
                int ePoint = sPoint + oText.Length;
                ssPoint = oSelection.Range.Start;
                oSelection.Text = oText;
                Word.Range rng = doc.Range(sPoint, ePoint);
                Word.Hyperlink mLink = oLinks.Add(rng, ref wAddress, ref wSubAddress);
            }
            else
            {
                ssPoint = oSelection.Range.Start;
                oSelection.Text = oText;                
            }
            if(newPara)
            {
                MoveEnd(doc);
                oSelection = wdApp.Selection;
                oSelection.Paragraphs.Add();
            }
            MoveEnd(doc);
            oSelection = wdApp.Selection;
            eePoint = oSelection.Range.Start;
            logger.LogWrite(":WordWriteOperation: ends");
        }

        private void btnClear_Click(object sender, EventArgs e)
        {
            logger.LogWrite(":btnClear_Click: begin");
            //string filename = Guid.NewGuid().ToString() + ".png";
            //clsScrapper cScrapper = new clsScrapper(Application.ExecutablePath);
            //clsMap oMap = cScrapper.ScrapData("https://patents.google.com/patent/US10560643?oq=patent:US10560643", filename);
            //cScrapper.CloseSession();
            chkExcel.Checked = true;
            chkWord.Checked = true;
            chkAbstract.Checked = true;
            chkClaims.Checked = true;
            chkDescription.Checked = true;            
            chkClaimsRequired.Checked = true;
            txtNumberOfRecords.Text = "";
            txtExcelFile.Text = txtOutputPath.Text = txtWordFile.Text = string.Empty;
            lblStatus.Text = "";
            txtNumberOfRecords.Focus();
            logger.LogWrite(":btnClear_Click: ends");
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnOutputFolder_Click(object sender, EventArgs e)
        {
            logger.LogWrite(":btnOutputFolder_Click: begin");
            CommonOpenFileDialog dialog = new CommonOpenFileDialog();
            dialog.InitialDirectory = "C:\\Users";
            dialog.IsFolderPicker = true;
            if (dialog.ShowDialog() == CommonFileDialogResult.Ok)
            {
                txtOutputPath.Text= dialog.FileName;
            }
            logger.LogWrite(":btnOutputFolder_Click: ends");
        }
        private void FindAndHighlight(Word.Document doc, string inputQuery)
        {
            logger.LogWrite(":FindAndHighlight: begin");
            try
            {
                
                string formatString = RemoveSpecialCharacters(inputQuery);
                foreach (string oText in formatString.Split(' '))
                {
                    Word.Range rng = doc.Range(ssPoint, eePoint);
                    //rng.Find.HitHighlight(oText, Word.WdColor.wdColorYellow);
                    //rng.Find.ClearFormatting();                    
                    rng.Find.MatchWholeWord = true;
                    rng.Find.MatchCase = false;
                    rng.Find.Text = oText;
                    while (rng.Find.Execute())
                    {
                        rng.Select();
                        rng.HighlightColorIndex = Word.WdColorIndex.wdYellow;
                    }

                }
            }
            catch (Exception ex)
            {
                logger.LogWrite(":FindAndHighlight: Error: " + ex.Message);
            }
            logger.LogWrite(":FindAndHighlight: ends");
        }
        private void FindAndBold(Word.Document doc, string oText, bool isSize=false)
        {
            logger.LogWrite(":FindAndBold: begin");
            string formatString = RemoveSpecialCharacters(oText);
            Word.Range rng = doc.Content;
            //rng.Find.ClearFormatting();
            rng.Find.MatchWholeWord = true;
            rng.Find.Text = formatString;            
            while (rng.Find.Execute())
            {
                rng.Select();
                rng.Font.Bold = Convert.ToInt32(true);
                if (isSize)
                    rng.Font.Size = 13;
            }
            logger.LogWrite(":FindAndBold: ends");
        }
        public string RemoveSpecialCharacters(string str)
        {
            logger.LogWrite(":RemoveSpecialCharacters: begin");
            Regex reg = new Regex("[*'\",_&#^@]");
            str = reg.Replace(str, string.Empty);

            Regex reg1 = new Regex("[ ]");
            return reg.Replace(str, "-");
            logger.LogWrite(":RemoveSpecialCharacters: ends");
        }

        private void chkClaims_CheckedChanged(object sender, EventArgs e)
        {
            if (chkClaims.Checked == false)
                chkClaimsRequired.Checked = false;
        }
    }
}
