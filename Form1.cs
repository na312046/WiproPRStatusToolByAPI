using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.VisualStudio.Services.Client;
using Microsoft.VisualStudio.Services.WebApi;
using System.IO;
using System.Windows.Forms;
using System.Text.RegularExpressions;
using System.Diagnostics;
using System.Collections;
using System.Threading;
using Microsoft.VisualStudio.Services.Common;
using Microsoft.TeamFoundation.WorkItemTracking.WebApi;
using Microsoft.TeamFoundation.WorkItemTracking.WebApi.Models;

namespace PRStatusAPITool
{
    public partial class Form1 : Form
    {

        string strQryLnk = "";
        string strfPath = "";
        string status = "";
        string strStatus = "";
        int PRNum = 0;
        string strXLSFilePath = "";
        string strCaption = "CRM :: Pull Request Status";

        public delegate void addListBox(string msg);
        public void addlog(string msg)
        {
            if (this.lstBxLog.InvokeRequired)
            {
                addListBox mydel = new addListBox(addlog);
                this.Invoke(mydel, new object[] { msg });
            }
            else
            {
                this.lstBxLog.Items.Add(msg);
            }

        }
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            clsStatus.CopyUserManual();
        }

        private void btnBrowse_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog ofd = new OpenFileDialog() { Filter = "Excel Workbook|*.xlsx;*.xls", ValidateNames = true })
            {
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    textBox1.Text = ofd.FileName;
                }

            }
        }

        private void btnStart_Click(object sender, EventArgs e)
        {
            if (!backgroundWorker1.IsBusy)
            {
                //progressBar1.Visible = true;
                lstBxLog.Items.Clear();
                backgroundWorker1.RunWorkerAsync();
            }
            else
            {
                MessageBox.Show("Busy Processing,Please Wait!!!");
            }

        }

        private void radioBtnQuery_CheckedChanged(object sender, EventArgs e)
        {
            btnBrowse.Visible = false;
            textBox1.Text = "Enter a Query";
            lstBxLog.Items.Clear();
        }

        private void radioBtnExcel_CheckedChanged(object sender, EventArgs e)
        {
            btnBrowse.Visible = true;
            textBox1.Text = "Choose an Excel Sheet with PR Links";
            lstBxLog.Items.Clear();
        }

        private void textBox1_MouseDown(object sender, MouseEventArgs e)
        {
            textBox1.Text = "";
        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            if (radioBtnQuery.Checked)
            {
                if (textBox1.Text != "" || textBox1.Text != null)
                {
                    strQryLnk = textBox1.Text.Trim();

                    if (string.IsNullOrEmpty(strQryLnk))
                    {
                        MessageBox.Show("Please enter query link.", strCaption, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    }
                    else if (!strQryLnk.ToLower().Contains("http://vstfmbs:8080/tfs/crm/engineering/_workitems#path="))
                    {
                        MessageBox.Show("Please enter valid query link.", strCaption, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    }
                    else
                    {
                        status = "";
                        addlog("Connecting to TFS and getting the query data....");
                        var prList = GetBugPrList(strQryLnk);

                        if (prList.Count > 0)
                        {
                            addlog("Saving in Excel....");
                            getStatusList(prList);
                        }
                        addlog(status);
                    }
                }
                else
                {
                    MessageBox.Show("Please enter the querylink");
                }

            }

            else if (radioBtnExcel.Checked == true)
            {
                if (textBox1.Text != "" || textBox1.Text != null)
                {
                    strfPath = textBox1.Text;
                    status = getStatusExcel(strfPath);
                    if (status == "")
                        MessageBox.Show("Please enter valid excel path and try again.", strCaption, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    else
                        addlog(status);
                }
                else
                {
                    MessageBox.Show("Please enter the filepath", strCaption, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }
            }
            else
            {
                MessageBox.Show("Please check the radiobutton");
            }

        }

        public string getStatusExcel(string fPath)
        {
            string strval = validation(fPath);
            try
            {
                if (strval == "true")
                {
                    addlog("Getting PR status....");
                    Excel.Application xlApp = new Excel.Application();
                    Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(@fPath);
                    Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
                    Excel.Range xlRange = xlWorksheet.UsedRange;

                    VssConnection connection2 = new VssConnection(new Uri($"https://dynamicscrm.visualstudio.com"), new VssAadCredential());
                    var statsAPI = new GetStatusAPI(connection2);

                    int rowCount = xlRange.Rows.Count;
                    int colCount = xlRange.Columns.Count;
                    int cCnt = 1;
                    int csCnt = 2;
                    if (rowCount > 1)
                    {
                        for (int rCnt = 1; rCnt <= rowCount; rCnt++)
                        {
                            if (rCnt == 1)
                            {
                                addlog("PR Number#               PR Status");
                            }
                            else
                            {
                                if (xlRange.Cells[rCnt, cCnt] != null && xlRange.Cells[rCnt, cCnt].Value2 != null)
                                {
                                    string PrLink = xlRange.Cells[rCnt, cCnt].Value2.ToString();
                                    strStatus = getStatus(PrLink.ToString(), statsAPI).Result.ToString();
                                    xlWorksheet.Cells[rCnt, csCnt] = strStatus;
                                    addlog(PRNum + "               " + strStatus);
                                    xlWorkbook.Save();
                                }
                                else
                                {
                                    strStatus = "Invalid PRLink";
                                    xlWorksheet.Cells[rCnt, csCnt] = strStatus;
                                    xlWorkbook.Save();
                                    addlog(strStatus);
                                }
                            }
                        }
                        xlWorkbook.Close();
                        strStatus = "The file in the same file path has been updated successfully";
                    }
                    else
                    {
                        xlWorkbook.Close();
                        strStatus = "No Data found in the given file";
                    }
                }
                else
                {
                    addlog(strval);
                }
            }
            catch (Exception ex)
            {
                strStatus = ex.Message;
                clsStatus.WriteLog("getStatusExcel()-->" + ex.Message);
            }
            return strStatus;
        }

        public void getStatusList(List<PrList> prList)
        {
            string strstats = "";
            try
            {
                if (prList.Count() > 1)
                {
                    try
                    {
                        var threadexl = new Thread((ThreadStart)(() => {
                            SaveFileDialog saveFileDialog = new SaveFileDialog();
                            saveFileDialog.Filter = "Execl files (*.xlsx)|*.xlsx";
                            saveFileDialog.Title = "Enter file name to save Query Data";
                            saveFileDialog.FilterIndex = 0;
                            saveFileDialog.RestoreDirectory = true;
                            saveFileDialog.CreatePrompt = false;
                            Form1 mainForm = new Form1();

                            DialogResult dlgRes = DialogResult.OK;
                            DialogResult dlgCancelRes = DialogResult.No;
                            do
                            {
                                dlgRes = saveFileDialog.ShowDialog();
                                if (dlgRes == DialogResult.OK)
                                    break;
                                else
                                {
                                    if (MessageBox.Show(mainForm,"Are you sure you want to cancel save operation?.", strCaption, MessageBoxButtons.YesNo, MessageBoxIcon.Information, MessageBoxDefaultButton.Button2) == DialogResult.Yes)
                                        dlgCancelRes = DialogResult.Yes;
                                }
                            } while (dlgCancelRes == DialogResult.No);

                            if (dlgRes == DialogResult.OK)
                            {
                                addlog("Writing bugs details in excel...");
                                strXLSFilePath = saveFileDialog.FileName;
                                if (File.Exists(strXLSFilePath))
                                    File.Delete(strXLSFilePath);
                                SaveInExcel(strXLSFilePath, prList);
                            }
                            else
                            {
                                addlog("You have canceled the save operation...");
                                SaveInExcel("NoPath", prList);
                            }
                        }));
                        threadexl.SetApartmentState(ApartmentState.STA);
                        threadexl.Start();
                        threadexl.Join();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                }
                else
                {
                    strStatus = "No Data Found";
                }
            }
            catch (Exception ex)
            {
                strStatus = ex.Message;
                clsStatus.WriteLog("getStatusList()-->" + ex.Message);
            }
        }

        public string validation(string fpath)
        {
            string strval = "true";
            try
            {
                if (closeOpenedExcel(fpath))
                    System.Threading.Thread.Sleep(1000);

                if (string.IsNullOrEmpty(fpath))
                {
                    strval = "No File path entered.";
                    return strval;
                }
                if (!File.Exists(fpath))
                {
                    strval = "Excel file does not exists in physical location.";
                    return strval;
                }
                string fileExt = Path.GetExtension(fpath);
                if (fileExt.ToUpper() != ".XLS" && fileExt.ToUpper() != ".XLSX")
                {
                    strval = "Please provide file in Excel format.";
                }
            }
            catch (Exception ex)
            {
                strval = ex.Message;
                clsStatus.WriteLog("validation()-->" + ex.Message);
            }
            return strval;
        }

        public async Task<string> getStatus(string prLink, GetStatusAPI statusApi)
        {
            string strStat = string.Empty;
            string[] strPRArr = null;
            if (prLink != null && prLink != "" && prLink != "NA" && (prLink.ToString().ToLower().Contains("pullrequest")))
            {
                strPRArr = System.Text.RegularExpressions.Regex.Split(prLink.ToLower(), "pullrequest/");
                if (strPRArr.Length >= 2)
                {
                    string strPRLink = strPRArr[0];
                   string  strPRID = strPRArr[1];
                        ArrayList splChars = new ArrayList();
                        splChars.Add("!"); splChars.Add("@"); splChars.Add("$");
                        splChars.Add("&"); splChars.Add("("); splChars.Add(")");
                        splChars.Add("?"); splChars.Add("#"); splChars.Add("^");
                        splChars.Add("%"); splChars.Add("/"); splChars.Add("*");
                        foreach (string chr in splChars)
                        {
                            if (strPRID.Contains(chr))
                                strPRID = strPRID.Substring(0, strPRID.IndexOf(chr));
                        }
                    if (strPRID == null || strPRID == "")
                    {
                        PRNum = 0;
                        strStat = "Invalid PRLink";
                    }
                    else
                    {
                        try
                        {
                            PRNum = Convert.ToInt32(strPRID);
                            var results1 = statusApi.GetPrStatus(PRNum).GetAwaiter().GetResult();
                            strStat = results1[0].ToString();
                        }
                        catch (Exception ex)
                        {
                            PRNum = 0;
                            strStat = ex.Message;
                            clsStatus.WriteLog("getStatus()-->" + ex.Message);
                            if (strStat == "Input string was not in a correct format.")
                                strStat = "Invalid PR NUmber";
                        }
                    }
                }
                else
                {
                    PRNum = 0;
                    strStat = "Invalid PRLink";
                }
            }
            else
            {
                PRNum = 0;
                strStat = "Invalid PRLink";
            }
            return strStat;
        }

        public static bool closeOpenedExcel(string fPath)
        {
            bool retVal = false;
            FileStream fs = null;
            try
            {
                fs = new FileStream(fPath, FileMode.Open, FileAccess.Read);
            }
            catch (Exception ex)
            {
                retVal = true;
                string fileName = Path.GetFileNameWithoutExtension(fPath).ToUpper().Trim();
                Process[] oProcess = Process.GetProcessesByName("EXCEL");
                foreach (Process item in oProcess)
                {
                    if (item.MainWindowTitle.ToUpper().Contains(fileName))
                        item.Kill();
                }
                //WriteLog("closeOpenedExcel()-->" + ex.Message);
            }
            finally
            {
                if (fs != null)
                    fs.Close();
            }
            return retVal;
        }

        private void helpToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            try
            {
                helpFrm objFrm = new helpFrm();
                objFrm.ShowDialog();
                System.Windows.Forms.Application.DoEvents();
                objFrm.Dispose();
            }
            catch (Exception ex)
            {
                clsStatus.WriteLog("helpToolStripMenuItem_Click()-->" + ex.Message);
            }
        }

        private void btnClose_Click_1(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            textBox1.Text = "";
            lstBxLog.Items.Clear();
            radioBtnExcel.Checked = true;
        }

        public void SaveInExcel(string fpath, List<PrList> prList)
        {
            try
            {
                if (fpath == "NoPath")
                {
                    addlog("Completed......");
                }
                else
                {
                    Excel.Application xlApp = new Excel.Application();
                    Excel.Workbook workbook = xlApp.Workbooks.Add(1);
                    Excel.Worksheet worksheet = (Excel.Worksheet)workbook.Sheets[1];
                    Excel.Range worksheet_range = null;

                    if (closeOpenedExcel(fpath))
                        System.Threading.Thread.Sleep(2000);

                    if (File.Exists(fpath))
                    {
                        File.Delete(fpath);
                    }

                    workbook.SaveAs(fpath, Excel.XlFileFormat.xlOpenXMLWorkbook, Type.Missing, Type.Missing, false, false, XlSaveAsAccessMode.xlNoChange, XlSaveConflictResolution.xlLocalSessionChanges, Type.Missing, Type.Missing);

                    string cellName;
                    int counter = 1;
                    cellName = "A" + counter.ToString();
                    worksheet.get_Range(cellName, cellName).Value2 = "BugID";
                    cellName = "B" + counter.ToString();
                    worksheet.get_Range(cellName, cellName).Value2 = "PR Number";
                    cellName = "C" + counter.ToString();
                    worksheet.get_Range(cellName, cellName).Value2 = "PR Status";
                    counter = counter + 1;
                    string status = "";

                    VssConnection connection2 = new VssConnection(new Uri($"https://dynamicscrm.visualstudio.com"), new VssAadCredential());
                    var comparertest = new GetStatusAPI(connection2);

                    foreach (var item in prList)
                    {
                        string PrDet = "";
                        cellName = "A" + counter.ToString();
                        var range1 = worksheet.get_Range(cellName, cellName);
                        range1.Value2 = item.BugId.ToString();
                        cellName = "B" + counter.ToString();
                        worksheet.get_Range(cellName, cellName).Value2 = item.PRLink.ToString();
                        status = item.PRStatus;
                        cellName = "C" + counter.ToString();
                        worksheet.get_Range(cellName, cellName).Value2 = status;
                        workbook.Save();
                        ++counter;
                    }

                    workbook.Close();
                    addlog("Status file has been saved in file path : " + fpath);
                }
        }
        catch (Exception ex)
        {
         MessageBox.Show(ex.Message);
        }
       }

        public List<PrList> GetBugPrList(string strQueryLnk)
        {
            var PRLists = new List<PrList>();
            string strCaption = "CRM :: Pull Request Status";

            //var collectionUri = "https://dynamicscrm.visualstudio.com";
            var conur = "http://vstfmbs:8080/tfs/CRM";
            var teamProjectName = "Engineering";
            var strError = "";

            VssConnection conn = new VssConnection(new Uri(conur), new VssCredentials());
            WorkItemTrackingHttpClient witClient = conn.GetClient<WorkItemTrackingHttpClient>();

            List<QueryHierarchyItem> queryHierarchyItems = witClient.GetQueriesAsync(teamProjectName, depth: 1).Result;

            try
            {
                string temp = System.Net.WebUtility.UrlDecode(strQueryLnk);
                string queryFullPath = Regex.Split(temp, "#path=", RegexOptions.IgnoreCase)[1]; //strQueryURL.Split(new string[] { "#path=" }, StringSplitOptions.None)[1];
                queryFullPath = queryFullPath.Substring(0, queryFullPath.ToLower().IndexOf("&_a=query"));

                string[] queryArr = Regex.Split(queryFullPath, "/", RegexOptions.IgnoreCase);

                string rootFolderName = queryArr[0];
                string strQueryName = queryArr[queryArr.Length - 1];
                int qFolderCount = queryArr.Length;

                QueryHierarchyItem myQueriesFolder = queryHierarchyItems.FirstOrDefault(qhi => qhi.Name.Equals(rootFolderName));
                if (myQueriesFolder != null)
                {
                    QueryHierarchyItem query = witClient.GetQueryAsync(teamProjectName, queryFullPath).Result;
                    WorkItemQueryResult result = witClient.QueryByIdAsync(query.Id).Result;

                    VssConnection connection2 = new VssConnection(new Uri($"https://dynamicscrm.visualstudio.com"), new VssAadCredential());
                    var apiStatus = new GetStatusAPI(connection2);

                    if (result.WorkItems.Any())
                    {
                        int skip = 0;
                        const int batchSize = 100;
                        string status = "";
                        IEnumerable<WorkItemReference> workItemRefs;
                        do
                        {
                            workItemRefs = result.WorkItems.Skip(skip).Take(batchSize);
                            if (workItemRefs.Any())
                            {
                                List<WorkItem> workItems = witClient.GetWorkItemsAsync(workItemRefs.Select(wir => wir.Id)).Result;
                                addlog("Getting PR Status.....");
                                addlog("Bug ID#               PR Number#               PR Status");
                                for (int i = 0; i < workItems.Count(); i++)
                                {
                                    PrList prList = new PrList();
                                    prList.BugId = workItems[i].Id.ToString();
                                   
                                    if (workItems[i].Fields.ContainsKey("Microsoft.CRM.PRLink"))
                                    {
                                        if (workItems[i].Fields["Microsoft.CRM.PRLink"].ToString() != "" || workItems[i].Fields["Microsoft.CRM.PRLink"].ToString() != "NA")
                                        {
                                            prList.PRLink = workItems[i].Fields["Microsoft.CRM.PRLink"].ToString();
                                            prList.PRStatus = getStatus(prList.PRLink.ToString(), apiStatus).Result;
                                            prList.PRNum = PRNum.ToString();
                                            if (prList.PRNum == "0")
                                                addlog(prList.BugId + "               " + "Invalid" + "               " + prList.PRStatus);
                                            else
                                                addlog(prList.BugId + "               " + prList.PRNum + "               " + prList.PRStatus);
                                        }
                                        else
                                        {
                                            prList.PRLink = "";
                                            addlog(prList.BugId + "               " + "Invalid" + "               " + prList.PRStatus);
                                        }
                                    }
                                    else
                                    {
                                        prList.PRLink = "";
                                        addlog(prList.BugId + "               " + "Invalid" + "               " + prList.PRStatus);
                                    }
                                    PRLists.Add(prList);
                                }

                            }
                            skip += batchSize;
                        }
                        while (workItemRefs.Count() == batchSize);
                    }
                    else
                    {
                        MessageBox.Show("No work items were returned from query.", strCaption, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        return PRLists;
                    }
                }
                else
                {
                    MessageBox.Show("The folder does not contain any query.", strCaption, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return PRLists;
                }
            }
            catch (Exception ex)
            {
                strError = ex.Message;
                clsStatus.WriteLog("GetBugPrList()-->" + ex.Message);
            }
            return PRLists;
        }

    }
}
