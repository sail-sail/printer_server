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
        private static string prn_errorStr = "D:/PrintDirectory/prn_error/";

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

        private static void printExcel(string excelPath)
        {
            Microsoft.Office.Interop.Excel.Application xApp = new Microsoft.Office.Interop.Excel.Application();
            xApp.Visible = false;
            Microsoft.Office.Interop.Excel.Workbook xBook = xApp.Workbooks._Open(excelPath);
            Microsoft.Office.Interop.Excel.Worksheet xSheet = null;
            try
            {
                try
                {
                    xApp.Run("Workbook_Open");
                }
                catch (Exception) 
                { 
                    
                }
                xSheet = (Microsoft.Office.Interop.Excel.Worksheet)xBook.ActiveSheet;
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
                System.Runtime.InteropServices.Marshal.ReleaseComObject(xApp);
                xSheet = null;
                xBook = null;
                xApp = null;
                GC.Collect();
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
                    if (!Directory.Exists(prn_errorStr))
                    {
                        Directory.CreateDirectory(prn_errorStr);
                    }
                    DirectoryInfo dirInfo = new DirectoryInfo(dirStr);
                    FileSystemInfo[] fileArr = dirInfo.GetFileSystemInfos();
                    for (int i = 0; i < fileArr.Length; i++)
                    {
                        FileInfo file = fileArr[i] as FileInfo;
                        if (file == null || file.Attributes == FileAttributes.Hidden) continue;
                        string extname = file.Name.Substring(file.Name.LastIndexOf(".")).ToLower();
                        if (extname != ".xlsx" && extname != ".xlsm") continue;
                        if (file.Name.Substring(0, 2) == "~$") continue;
                        string[] arr = file.Name.Split('$');
                        Thread.Sleep(100);
                        
                        //默认打印机
                        string defPrtr = null;
                        if (arr.Length == 2)
                        {
                            defPrtr = arr[1].Substring(0, arr[1].Length - extname.Length);
                            if (!String.IsNullOrWhiteSpace(defPrtr))
                            {
                                byte[] defPrtrByte = Convert.FromBase64String(defPrtr);
                                defPrtr = Encoding.UTF8.GetString(defPrtrByte);
                            }
                        }
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
                            try
                            {
                                log(file.FullName);
                                if (!String.IsNullOrWhiteSpace(defPrtr))
                                {
                                    int reint = SetDefaultPrinter(defPrtr);
                                    if (reint == 0) throw new Exception("设置默认打印机 " + defPrtr + " 失败!");
                                }
                                printExcel(file.FullName);
                            }
                            catch (Exception e4)
                            {
                                if (File.Exists(prn_errorStr + "/" + file.Name))
                                {
                                    File.Delete(prn_errorStr + "/" + file.Name);
                                    log("File.Delete:" + prn_errorStr + "/" + file.Name);
                                }
                                file.MoveTo(prn_errorStr + "/" + file.Name);
                                throw e4;
                            }
                            try
                            {
                                string logDir = prn_logStr + System.DateTime.Now.ToString("yyyy-MM-dd");
                                if (!Directory.Exists(logDir))
                                {
                                    Directory.CreateDirectory(logDir);
                                }
                                if (File.Exists(logDir + "/" + file.Name))
                                {
                                    File.Delete(logDir + "/" + file.Name);
                                    log("File.Delete:" + logDir + "/" + file.Name);
                                }
                                file.MoveTo(logDir + "/" + file.Name);
                            }
                            finally
                            {
                                FileStream fs = new FileStream(prn_logStr + "/PrintFileArr.log", FileMode.OpenOrCreate, FileAccess.Write);
                                StreamWriter sw = new StreamWriter(fs);
                                sw.BaseStream.Seek(0, SeekOrigin.End);
                                sw.WriteLine("[" + DateTime.Now.ToString("yyyy-MM-dd HH:MM:ss") + "]|" + filename);
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
                //Console.WriteLine(mo["Name"].ToString());
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
