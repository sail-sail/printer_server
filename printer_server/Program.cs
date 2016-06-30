using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading;
using System.IO;
using System.Management;

namespace printer_server
{
    class Program
    {
        private static string execState = "stop";

        private static string dirStr = "D:/PrintDirectory/";
        private static string prn_logStr = "D:/PrintDirectory/prn_log/";

        static void Main(string[] args)
        {
            //ThreadStart threadStart = new ThreadStart(exec);
            //Thread thread = new Thread(threadStart);
            //thread.IsBackground = true;
            //execState = "running";
            //thread.Start();

            execState = "running";
            exec();
        }

        private static void printExcel(string excelPath,int sheetNum)
        {
            if (sheetNum < 1) sheetNum = 1;
            Microsoft.Office.Interop.Excel.Application xApp = new Microsoft.Office.Interop.Excel.Application();
            xApp.Visible = false;
            Microsoft.Office.Interop.Excel.Workbook xBook = xApp.Workbooks._Open(excelPath);
            try
            {
                Microsoft.Office.Interop.Excel.Worksheet xSheet = (Microsoft.Office.Interop.Excel.Worksheet)xBook.Sheets[sheetNum];
                xSheet.PrintOut(1, 1, 1, false);
            }
            catch (Exception e)
            {
                throw e;
            }
            finally
            {
                xBook.Close(false);
                xApp.Quit();
            }
        }

        public static void exec()
        {
            while (true)
            {
                if (execState.Equals("stop")) return;
                Thread.Sleep(300);
                try
                {
                    if (!Directory.Exists(dirStr))
                    {
                        Directory.CreateDirectory(dirStr);
                    }
                    if (!Directory.Exists(prn_logStr))
                    {
                        Directory.CreateDirectory(prn_logStr);
                    }
                    DirectoryInfo dirInfo = new DirectoryInfo(dirStr);
                    FileSystemInfo[] fileArr = dirInfo.GetFileSystemInfos();
                    for (int i = 0; i < fileArr.Length; i++)
                    {
                        FileInfo file = fileArr[i] as FileInfo;
                        if (file == null || file.Attributes == FileAttributes.Hidden) continue;
                        string extname = file.Name.Substring(file.Name.LastIndexOf(".")).ToLower();
                        if (extname != ".xlsx") continue;
                        if (file.Name.Substring(0, 2) == "~$") continue;
                        string[] arr = file.Name.Split('$');
                        if (arr.Length != 3) continue;
                        Thread.Sleep(100);
                        int sheetNum = 1;
                        try
                        {
                            sheetNum = int.Parse(arr[1]);
                        }
                        catch (Exception e3)
                        {
                            logErr(e3.Message);
                        }
                        
                        //默认打印机
                        string defPrtr = arr[2].Substring(0, arr[2].Length - extname.Length);
                        if (String.IsNullOrWhiteSpace(defPrtr)) continue;
                        //文件名
                        string filename = arr[0]+extname;
                        if (String.IsNullOrWhiteSpace(filename)) continue;
                        bool isHas = false;
                        StreamReader sr = null;
                        FileStream fsr = null;
                        try
                        {
                            fsr = new FileStream(prn_logStr + "/PrintFileArr.log", FileMode.OpenOrCreate, FileAccess.Read);
                            sr = new StreamReader(fsr);
                            while (true)
                            {
                                string line = sr.ReadLine();
                                if (line == null) break;
                                arr = line.Split('|');
                                if (arr.Length != 2) continue;
                                string name0 = arr[1];
                                if (String.IsNullOrEmpty(name0)) continue;
                                if (name0 == filename)
                                {
                                    isHas = true;
                                    break;
                                }
                            }
                        }
                        finally
                        {
                            if (sr!=null) sr.Close();
                            if(fsr!=null) fsr.Close();
                        }
                        if (isHas) continue;
                        try
                        {
                            log(file.FullName);
                            int reint = SetDefaultPrinter(defPrtr);
                            if (reint == 0) throw new Exception("设置默认打印机 " + defPrtr + " 失败!");
                            printExcel(file.FullName,sheetNum);
                            try
                            {
                                string logDir = prn_logStr + System.DateTime.Now.ToString("d");
                                if (!Directory.Exists(logDir))
                                {
                                    Directory.CreateDirectory(logDir);
                                }
                                file.MoveTo(logDir + "/" + file.Name);
                            }
                            finally
                            {
                                FileStream fs = new FileStream(prn_logStr + "/PrintFileArr.log", FileMode.OpenOrCreate, FileAccess.Write);
                                StreamWriter sw = new StreamWriter(fs);
                                sw.BaseStream.Seek(0, SeekOrigin.End);
                                sw.WriteLine("[" + DateTime.Now.ToString() + "]|" + filename);
                                sw.Flush();
                                sw.Close();
                                fs.Close();
                            }
                            
                        }
                        catch (Exception e2)
                        {
                            logErr(e2.Message);
                        }
                    }
                }
                catch (Exception e)
                {
                    logErr(e.Message);
                }
            }
        }

        //设置默认打印机 存在打印机并设置成功 返回1 失败0
        protected static int SetDefaultPrinter(string PrinterName)
        {

            int reint = 0;
            ManagementObjectSearcher query;
            ManagementObjectCollection queryCollection;
            string _classname = "SELECT * FROM Win32_Printer";

            query = new ManagementObjectSearcher(_classname);
            queryCollection = query.Get();

            foreach (ManagementObject mo in queryCollection)
            {
                if (string.Compare(mo["Name"].ToString(), PrinterName, true) == 0)
                {
                    mo.InvokeMethod("SetDefaultPrinter", null);
                    reint = 1;
                    break;
                }
            }
            return reint;
        }

        private static void log(string msg)
        {
            FileStream fs = new FileStream(prn_logStr + "/PrintDirectory.log", FileMode.OpenOrCreate, FileAccess.Write);
            StreamWriter sw = new StreamWriter(fs);
            sw.BaseStream.Seek(0, SeekOrigin.End);
            sw.WriteLine("[" + DateTime.Now.ToString() + " log]:" + msg);
            sw.Flush();
            sw.Close();
            fs.Close();
        }

        private static void logErr(string msg)
        {
            FileStream fs = new FileStream(prn_logStr + "/PrintDirectory.log", FileMode.OpenOrCreate, FileAccess.Write);
            StreamWriter sw = new StreamWriter(fs);
            sw.BaseStream.Seek(0, SeekOrigin.End);
            sw.WriteLine("[" + DateTime.Now.ToString() + " err]:" + msg);
            sw.Flush();
            sw.Close();
            fs.Close();
        }
    }
}
