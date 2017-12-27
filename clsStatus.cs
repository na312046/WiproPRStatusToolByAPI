using System;
using System.Collections.Generic;
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

namespace PRStatusAPITool
{
    public static class clsStatus
    {

        public static void WriteLog(string strError)
        {
            try
            {
                string AppPath = AppDomain.CurrentDomain.BaseDirectory;
                string strLog = @"LOG\";
                string strFilePath = AppPath + strLog;

                if (!(Directory.Exists(strFilePath)))
                {
                    Directory.CreateDirectory(strFilePath);
                }
                string fn = string.Format("{0}{1}.txt", strFilePath, DateTime.Now.ToString("ddMMyyyy"));
                FileStream fs = new FileStream(fn, FileMode.Append, FileAccess.Write, FileShare.ReadWrite);

                StreamWriter writer = new StreamWriter(fs);
                //writer.Write("[ " + DateTime.Now.Hour.ToString() + ":" + DateTime.Now.Minute.ToString() + ":" + DateTime.Now.Second.ToString() + " ]");
                //writer.WriteLine(strError);
                //   writer.WriteLine("--------------------------------------------------------------------------");
                writer.WriteLine(string.Format("[ {0} ] {1}", DateTime.Now.ToString("HH:mm:ss"), strError));
                writer.Close();
                fs.Close();
            }
            finally
            {
                //nothing
            }
        }

        public static void CopyUserManual()
        {
                try
                {
                    string filePath = AppDomain.CurrentDomain.BaseDirectory + "UserManual.pdf";
                    if (!File.Exists(filePath))
                    {
                        string sourcePath = Directory.GetParent(Environment.CurrentDirectory).ToString().Replace("bin", "Docs");
                        if ((Directory.Exists(sourcePath)))
                        {
                            string strSourceFilePath = sourcePath + "\\UserManual.pdf";
                            if (File.Exists(strSourceFilePath))
                            {
                                File.Copy(strSourceFilePath, filePath);
                            }
                            else
                                WriteLog("CopyUserManual-->UserManual.pdf file does not exist in Docs folder");
                        }
                        else
                            WriteLog("CopyUserManual-->Docs folder does not exist");
                    }
                }
                catch (Exception ex)
                {
                    WriteLog("CopyUserManual-->" + ex.Message);
                }

        }

    }
}
