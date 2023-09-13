using System;
using System.Collections.Generic;
using System.IO;
using Microsoft.VisualStudio.Tools.Applications;
using Microsoft.VisualStudio.Tools.Applications.Deployment;
using System.Windows.Forms;
using Microsoft.VisualBasic.Logging;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Linq;


namespace VSTOInstaller
{
    public class VSTOInstaller : IAddInPostDeploymentAction
    {
        public void Execute(AddInPostDeploymentActionArgs args)
        {

            Uri deploymentManifestUri = args.ManifestLocation;
            string sourcePath = args.AddInPath;
            string deployPath = Path.Combine(sourcePath, "Deploy");
            string destPath = Path.Combine(@"C:\RealSMART\", args.ProductName);
            string executablePath = "";
            string databasePath = "";

            //LogWrite("sourcePath : " + sourcePath);
            //LogWrite("deployPath : " + deployPath);
            //LogWrite("destPath : " + destPath);

            /// 복사대상 파일 리스트를 만든다.
            List<string> sourceFiles = new List<string>();

            /// 배포PC에 저장될 상대경로를 만든다.
            List<string> destFiles = new List<string>();

            /// Deploy폴더 하위에 있는 파일 리스트를 긁어서 복사대상파일과 상대경로 리스트를 채운다.
            DirectoryInfo dr = new DirectoryInfo(deployPath);
            if (dr.Exists)
            {
                foreach (FileInfo f in dr.GetFiles("*.*", SearchOption.AllDirectories))
                {
                    if (!f.Name.StartsWith("~$"))
                    {
                        sourceFiles.Add(f.FullName);
                        destFiles.Add(destPath + f.FullName.Replace(deployPath, ""));

                        if (f.Extension == ".xlsx" && Path.GetDirectoryName(f.FullName) == deployPath) executablePath = destFiles[destFiles.Count - 1];        // 추가된 파일이  배포루트에 존재하는 xlsx파일이면 실행파일로 간주한다. 방금 추가된 값을 취득하면 된다.
                        if (f.Extension == ".sdb") databasePath = destFiles[destFiles.Count - 1];                                            // 추가된 파일이 sdb 파일이면 db파일로 간주한다. 1사이트에 1개의 db파일만 존재해야만 성립함.

                        //LogWrite(" - sourceFiles : " + f.FullName + " / extension : " + f.Extension);
                        //LogWrite(" - destFiles : " + destFiles[destFiles.Count - 1]);

                    }
                }
                //LogWrite("executablePath : " + executablePath);
                //LogWrite("databasePath : " + databasePath);
            }



            switch (args.InstallationStatus)
            {
                case AddInInstallationStatus.InitialInstall:
                case AddInInstallationStatus.Update:


                    for (int i = 0 ; i < sourceFiles.Count ; i++)
                    {
                        if (File.Exists(sourceFiles[i]))
                        {
                            bool deploy = true;

                            // 목적지 파일명을 검사해서 해당 디렉토리가 없으면 디렉토리를 생성시킨다.
                            if (!Directory.Exists(Path.GetDirectoryName(destFiles[i]))) Directory.CreateDirectory(Path.GetDirectoryName(destFiles[i]));

                            // 목적지 파일명이 db파일명이면서, db파일이 이미 해당위치에 존재하는 경우, db파일을 일단 백업시켜준다.  --> 백업본 남기고 덮어씌우는 방식은 일단 보류. DB파일은 배포하지 않는다.
                            if (destFiles[i] == databasePath && File.Exists(databasePath)) deploy = false; //File.Copy(databasePath, string.Format("{0}_Backup_{1}", databasePath, DateTime.Now.ToString("yyyyMMdd_HHmmss")));

                            // 업데이트모드이면서, 파일이 executable path이면 (실행용 엑셀파일) 이 파일은 배포하지 않는다.
                            // if (args.InstallationStatus == AddInInstallationStatus.Update && destFiles[i] == executablePath) deploy = false;


                            /// 파일을 복사하면서 덮어씌운다.
                            try
                            {
                                if (deploy) File.Copy(sourceFiles[i], destFiles[i], true);
                            }
                            catch (Exception)
                            {
                                if (destFiles[i] == executablePath)
                                {
                                    string updatedExecutablePath = executablePath.Replace(Path.GetExtension(executablePath), "") + "_" + args.Version + Path.GetExtension(executablePath);
                                    File.Copy(sourceFiles[i], updatedExecutablePath, true);
                                    executablePath = updatedExecutablePath;
                                }
                                else
                                {
                                    File.Copy(sourceFiles[i], Path.GetExtension(destFiles[i]) + "_Updated", true);
                                }
                            }
                        }
                    }

                    ServerDocument.RemoveCustomization(executablePath);
                    ServerDocument.AddCustomization(executablePath, deploymentManifestUri);

                    if (args.InstallationStatus == AddInInstallationStatus.Update)
                        MessageBox.Show("프로그램이 업데이트 되었습니다.\r\n프로그램을 종료하시고 새로운 실행 파일을 확인하여 다시 실행하세요.");
                    break;

                case AddInInstallationStatus.Uninstall:
                    foreach (string file in destFiles)
                    {
                        if (File.Exists(file) && file != databasePath && !file.Replace(deployPath,"").Contains("Data"))     // 파일이 데이터파일이 아니고, 배포폴더의 하위 Data폴더 내에 들어있는 파일이 아니면 삭제.
                        {
                            File.Delete(file);
                        }
                    }
                    break;
            }
        }


        //******************************************************************************************************************************************************
        // 로그를 기록함
        //******************************************************************************************************************************************************
        private static Log log = new Log();
        public static void LogWrite(string log_text)
        {
            // 로그 기록 폴더        
            log.DefaultFileLogWriter.CustomLocation = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + @"\RealSMART_Log";
            // 로그 파일 명(프로그램명_날짜)        
            log.DefaultFileLogWriter.BaseFileName = "ErrorLog_" + DateTime.Now.ToString("yyyy-MM-dd");
            // 로그 내용 기록        
            log.WriteEntry(String.Format(DateTime.Now.ToString(), "yyyy-MM-dd HH:mm:ss") + "  ===  " + log_text, TraceEventType.Information);
            // 로그 기록 닫기        
            log.DefaultFileLogWriter.Close();
        }


        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        public static void CloseAllExcel()
        {
            Process[] processlist = Process.GetProcesses();
            int pCount = processlist.Where(x => x.ProcessName.ToLower().Equals("excel")).Count();

            if (pCount > 0)
            {
                while (pCount > 0)
                {
                    Microsoft.Office.Interop.Excel.Application oExcelApp = null;
                    try
                    {
                        oExcelApp = (Microsoft.Office.Interop.Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
                    }
                    catch { }

                    if (oExcelApp != null)
                    {
                        try
                        {
                            while (oExcelApp.Workbooks.Count > 0)
                                oExcelApp.Workbooks[1].Close(false);

                            uint pID = 0;
                            oExcelApp.Quit();
                            GetWindowThreadProcessId((IntPtr)((Microsoft.Office.Interop.Excel.Application)oExcelApp).Hwnd, out pID);
                            Process.GetProcessById((int)pID).Kill();
                        }
                        catch { }
                    }

                    processlist = Process.GetProcesses();
                    pCount = processlist.Where(x => x.ProcessName.ToLower().Equals("excel")).Count();
                }
            }
        }

    }
}
