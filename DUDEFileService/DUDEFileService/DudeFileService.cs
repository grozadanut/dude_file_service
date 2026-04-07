using System;
using System.CodeDom.Compiler;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using dude;

namespace DUDEFileService
{
    public partial class DudeFileService : ServiceBase
    {
        public static readonly String ECR_FOLDER_KEY = "ecr_folder";

        private static readonly String RESULT_SUFFIX = "_result";
        private static readonly int NO_RETRY = 200;
        private static readonly int NO_RETRY_CLOSE = 6000;

        private dude.CFD_DUDE serv;
        private bool inCommand = false;

        public static bool IsFileReady(string filename)
        {
            // If the file can be opened for exclusive access it means that the file
            // is no longer locked by another process.
            try
            {
                using (FileStream inputStream = File.Open(filename, FileMode.Open, FileAccess.Read, FileShare.None))
                    return inputStream.Length > 0;
            }
            catch (Exception)
            {
                return false;
            }
        }

        public static void WaitForFile(string filename)
        {
            //This will lock the execution until the file is ready
            //TODO: Add some logic to make it async and cancelable
            while (!IsFileReady(filename)) { }
        }

        public DudeFileService()
        {
            InitializeComponent();
            if (!string.IsNullOrEmpty(Environment.GetEnvironmentVariable(ECR_FOLDER_KEY)))
                this.fileSystemWatcher1.Path = Environment.GetEnvironmentVariable(ECR_FOLDER_KEY);
            fileSystemWatcher1.EnableRaisingEvents = true;
        }

        protected override void OnStart(string[] args)
        {
            //var config = new Dictionary<string, string>();
            //foreach (var row in file.readalllines("c:\\datecs\\config.ini"))
            //    config.add(row.split('=')[0], string.Join("=", row.Split('=').Skip(1).ToArray()));
            //config["ecr_ip"];
            Start_COMServer();
        }

        protected override void OnStop()
        {
            int result = Stop_COMServer();
            eventLog1.WriteEntry("Stop_COMServer result " + result);
        }

        private void FileSystemWatcher_Created(object sender, FileSystemEventArgs e)
        {
            WaitForFile(e.FullPath);

            string ecrCommands = "";
            string ecrIp = "127.0.0.1";
            string ecrPort = "3999";

            // read commands file
            // file show have the structure:
            // first line: ecr ip address
            // second line: ecr port
            // other lines: ecr commands
            String[] lines = File.ReadAllLines(e.FullPath);
            string resultFilePath = e.FullPath + RESULT_SUFFIX;

            if (lines.Length < 3)
            {
                File.WriteAllText(resultFilePath, "-1: Fisierul de comenzi trebuie sa contina cel putin 3 linii!");
                return;
            }

            ecrIp = lines[0];
            ecrPort = lines[1];

            for (int i = 2; i < lines.Length; i++)
                ecrCommands += lines[i] + Environment.NewLine;

            int noRetries = NO_RETRY;
            while (inCommand)
            {
                if (noRetries-- < 0)
                {
                    eventLog1.WriteEntry("OpenConnection Timeout expired " + e.FullPath);
                    break;
                }
                System.Threading.Thread.Sleep(100);
            }

            int resultCode = OpenConnection(ecrIp, ecrPort);
            if (resultCode != 0)
            {
                eventLog1.WriteEntry("OpenConnection result " + resultCode + " ecrIp "+ecrIp + " ecrPort "+ecrPort);
                File.WriteAllText(resultFilePath, resultCode + ": " + serv.lastError_Message);
                return;
            }

            if (ecrCommands.StartsWith("raportmf"))
            {
                string[] command = ecrCommands.Split('&');// raportmf&start&end&directory
                DownloadAnafXML(command[1], command[2], command[3], e.FullPath, resultFilePath);
            }
            else if (ecrCommands.StartsWith("receipts"))
            {
                string[] command = ecrCommands.Split('&');// receipts&start&end
                DownloadReceipts(command[1].Trim(), command[2].Trim(), e.FullPath, resultFilePath);
            }
            else
                ExecuteScript(ecrCommands, e.FullPath, resultFilePath);

            noRetries = NO_RETRY_CLOSE;
            while (inCommand)
            {
                if (noRetries-- < 0)
                {
                    eventLog1.WriteEntry("StopConnection Timeout expired " + e.FullPath);
                    inCommand = false;
                    break;
                }
                System.Threading.Thread.Sleep(100);
            }

            StopConnection();
        }


        private void ExecuteScript(string cmd_Script, string inputFilePath, string resultFilePath)
        {
            //while (inCommand)
            //    System.Threading.Thread.Sleep(100);

            inCommand = true;

            ThreadPool.QueueUserWorkItem(delegate
            {
                try
                {
                    int result = serv.execute_Script_V1(TScriptType.DS, cmd_Script);
                    if (result != 0)
                        eventLog1.WriteEntry("Script result " + result + " for commands " + cmd_Script);
                    else // delete the executed print file
                        File.Delete(inputFilePath);

                    // always write the result file, even if success
                    // result file is read and deleted by Colibri Platform
                    File.WriteAllText(resultFilePath, result + ": " + serv.lastError_Message);
                }
                catch (Exception ex)
                {
                    inCommand = false;
                    eventLog1.WriteEntry("Script exception: " + ex.Message);
                    File.WriteAllText(resultFilePath, -145 + ": " + ex.Message);
                }
                finally
                {
                    inCommand = false;
                }
            }, null);
        }

        /**
         * DD-MM-YY hh:mm:ss DST
         */
        private void DownloadAnafXML(string startDateTime, string endDateTime, string chosenDirectory, string inputFilePath, string resultFilePath)
        {
            //while (inCommand)
            //    System.Threading.Thread.Sleep(100);

            inCommand = true;

            bool Old_Active_OnSendCommand;
            bool Old_Active_OnReceiveAnswer;
            bool Old_Active_OnWait;
            bool Old_Active_OnStatusChange;
            bool Old_Active_OnError;

            Old_Active_OnSendCommand = serv.active_OnSendCommand;
            Old_Active_OnReceiveAnswer = serv.active_OnReceiveAnswer;
            Old_Active_OnWait = serv.active_OnWait;
            Old_Active_OnStatusChange = serv.active_OnStatusChange;
            Old_Active_OnError = serv.active_OnError;

            try
            {
                int error_Code = serv.set_Download_Path(chosenDirectory.Trim());
                if (error_Code != 0)
                {
                    eventLog1.WriteEntry("DownloadAnafXML error at set_Download_Path " + error_Code);
                    inCommand = false;
                    File.WriteAllText(resultFilePath, error_Code + ": " + serv.lastError_Message);
                    return;
                }

                string chosenPath = serv.download_Path;
                if (chosenPath == null || !chosenPath.Trim().Equals(chosenDirectory.Trim()))
                {
                    inCommand = false;

                    string chosenPathC = "";
                    foreach (char c in chosenPath)
                        chosenPathC += " "+((int)c);
                    string chosenDirectoryC = "";
                    foreach (char c in chosenDirectory)
                        chosenDirectoryC += " " + ((int)c);

                    eventLog1.WriteEntry("DownloadAnafXML serv.download_Path not correct " + chosenPath +"("+ chosenPathC + ") should be " + chosenDirectory+"("+ chosenDirectoryC+")");
                    File.WriteAllText(resultFilePath, "-1: DownloadAnafXML serv.download_Path not correct " + chosenPath + " should be " + chosenDirectory);
                    return;
                }
                
                serv.DateRange_StartValue = startDateTime;
                serv.DateRange_EndValue = endDateTime;

                ThreadPool.QueueUserWorkItem(delegate
                {
                    try
                    {
                        serv.set_CommunicationEvents(false, false, false, false, true, false);
                        error_Code = serv.download_ANAF_DTRange();
                    }
                    finally
                    {
                        inCommand = false;
                        if (error_Code != 0)
                            eventLog1.WriteEntry(serv.lastError_Message);
                        else
                            File.Delete(inputFilePath);
                        serv.set_CommunicationEvents(Old_Active_OnSendCommand, Old_Active_OnWait, Old_Active_OnReceiveAnswer, Old_Active_OnStatusChange, Old_Active_OnError, false);
                        File.WriteAllText(resultFilePath, error_Code + ": " + serv.lastError_Message);
                    }
                }, null);
            }
            finally
            {

            }
        }

        /**
         * DD-MM-YY hh:mm:ss DST
         */
        private void DownloadReceipts(string startDateTime, string endDateTime, string inputFilePath, string resultFilePath)
        {
            try
            {
                string ErrorCode = null;
                string StartDate = null;
                string EndDate = null;
                string RepFirstDoc = null;
                string FirstDoc = null;
                string RepLastDoc = null;
                string LastDoc = null;
                int error_Code = execute_124_ej_Search_Documents_ByDate(startDateTime, endDateTime, "1",
                    ref ErrorCode, ref StartDate, ref EndDate, ref RepFirstDoc, ref FirstDoc, ref RepLastDoc, ref LastDoc);

                if (error_Code != 0)
                {
                    eventLog1.WriteEntry("DownloadReceipts error at execute_124_ej_Search_Documents_ByDate: " + error_Code);
                    File.WriteAllText(resultFilePath, ErrorCode + ": " + serv.lastError_Message);
                    return;
                }

                int numOfReceipts = int.Parse(LastDoc);
                int z = int.Parse(RepFirstDoc);
                string receiptLines = "";

                for (int i = int.Parse(FirstDoc); i < numOfReceipts; i++)
                {
                    string DocNumber = null;
                    string RecReport = null;
                    string RecNumber = null;
                    string Date = null;
                    string DocType = null;
                    string ZNumber = null;

                    error_Code = execute_125_ej_Set_Document_For_Reading("0", String.Format("{0}{1:0000}", z, i), "1", 
                        ref ErrorCode, ref DocNumber, ref RecReport, ref RecNumber, ref Date, ref DocType, ref ZNumber);

                    if (error_Code != 0)
                    {
                        eventLog1.WriteEntry("DownloadReceipts error at execute_125_ej_Set_Document_For_Reading: " + error_Code);
                        File.WriteAllText(resultFilePath, ErrorCode + ": " + serv.lastError_Message);
                        return;
                    }

                    string TextData = null;
                    while (execute_125_ej_Get_LineAsText("1", ref ErrorCode, ref TextData) == 0)
                    {
                        receiptLines += TextData + Environment.NewLine;
                    }
                }

                File.WriteAllText(resultFilePath, String.Format("0:{0}{1}", Environment.NewLine, receiptLines));
            }
            finally
            {

            }
        }

        int execute_124_ej_Search_Documents_ByDate(
            string input_StartDate,
            string input_EndDate,
            string DocumentType,
            ref string ErrorCode,
            ref string StartDate,
            ref string EndDate,
            ref string RepFirstDoc,
            ref string FirstDoc,
            ref string RepLastDoc,
            ref string LastDoc)
        {
            const string cmd = "124_ej_Search_Documents_ByDate";
            int Result = -1;
            ErrorCode = "-1";
            try
            {
                try
                {
                    if (!serv.connected_ToDevice) return Result;
                    if (serv.set_InputParam_ByName(cmd, "input_StartDate", input_StartDate) != 0)
                    {
                        
                        eventLog1.WriteEntry("execute_124_ej_Search_Documents_ByDate input_StartDate: " + serv.lastError_Code);
                        return Result;
                    }
                    if (serv.set_InputParam_ByName(cmd, "input_EndDate", input_EndDate) != 0)
                    {
                        eventLog1.WriteEntry("execute_124_ej_Search_Documents_ByDate input_EndDate: " + serv.lastError_Code);
                        return Result;
                    }
                    if (serv.set_InputParam_ByName(cmd, "DocumentType", DocumentType) != 0)
                    {
                        eventLog1.WriteEntry("execute_124_ej_Search_Documents_ByDate DocumentType: " + serv.lastError_Code);
                        return Result;
                    }
                    if (serv.execute_Command_ByName(cmd) != 0)
                    {
                        eventLog1.WriteEntry("execute_124_ej_Search_Documents_ByDate execute: " + serv.lastError_Code);
                        return Result;
                    }
                    serv.get_OutputParam_ByName(cmd, "StartDate", ref StartDate);
                    serv.get_OutputParam_ByName(cmd, "EndDate", ref EndDate);
                    serv.get_OutputParam_ByName(cmd, "RepFirstDoc", ref RepFirstDoc);
                    serv.get_OutputParam_ByName(cmd, "FirstDoc", ref FirstDoc);
                    serv.get_OutputParam_ByName(cmd, "RepLastDoc", ref RepLastDoc);
                    serv.get_OutputParam_ByName(cmd, "LastDoc", ref LastDoc);
                }
                catch (Exception ex)
                {
                    eventLog1.WriteEntry("execute_124_ej_Search_Documents_ByDate: " + ex.Message);
                }
            }
            finally
            {
                Result = serv.lastError_Code;
                serv.get_OutputParam_ByName(cmd, "ErrorCode", ErrorCode);
            }
            return Result;
        }

        int execute_125_ej_Set_Document_For_Reading(
            string Option,         
            string input_DocNumber,
            string input_DocType,  
            ref string ErrorCode,  
            ref string DocNumber,  
            ref string RecReport,  
            ref string RecNumber,  
            ref string Date,       
            ref string DocType,    
            ref string ZNumber)                          
        {
            const string cmd = "125_ej_Set_Document_For_Reading";
            int Result = -1;
            ErrorCode = "-1";
            if (serv == null) return Result;
            try
            {
                try
                {
                    if (!serv.connected_ToDevice) return Result;
                    if (serv.set_InputParam_ByName(cmd, "Option", Option) != 0)
                    {
                        eventLog1.WriteEntry("execute_125_ej_Set_Document_For_Reading Option: " + serv.lastError_Code);
                        return Result;
                    }
                    if (serv.set_InputParam_ByName(cmd, "input_DocNumber", input_DocNumber) != 0)
                    {
                        eventLog1.WriteEntry("execute_125_ej_Set_Document_For_Reading input_DocNumber: " + serv.lastError_Code);
                        return Result;
                    }
                    if (serv.set_InputParam_ByName(cmd, "input_DocType", input_DocType) != 0)
                    {
                        eventLog1.WriteEntry("execute_125_ej_Set_Document_For_Reading input_DocType: " + serv.lastError_Code);
                        return Result;
                    }
                    if (serv.execute_Command_ByName(cmd) != 0)
                    {
                        eventLog1.WriteEntry("execute_125_ej_Set_Document_For_Reading execute: " + serv.lastError_Code);
                        return Result;
                    }
                    serv.get_OutputParam_ByName(cmd, "DocNumber", ref DocNumber);
                    serv.get_OutputParam_ByName(cmd, "RecReport", ref RecReport);
                    serv.get_OutputParam_ByName(cmd, "RecNumber", ref RecNumber);
                    serv.get_OutputParam_ByName(cmd, "Date", ref Date);
                    serv.get_OutputParam_ByName(cmd, "DocType", ref DocType);
                    serv.get_OutputParam_ByName(cmd, "ZNumber", ref ZNumber);
                }
                catch (Exception ex)
                {
                    eventLog1.WriteEntry("execute_124_ej_Search_Documents_ByDate: " + ex.Message);
                }
            }
            finally
            {
                Result = serv.lastError_Code;
                serv.get_OutputParam_ByName(cmd, "ErrorCode", ErrorCode);
            }
            return Result;
        }

        int execute_125_ej_Get_LineAsText(
            string Option,
            ref string ErrorCode,
            ref string TextData)
        {
            const string cmd = "125_ej_Get_LineAsText";
            int Result = -1;
            ErrorCode = "-1";
            if (serv == null) return Result;
            try
            {
                try
                {
                    if (!serv.connected_ToDevice)
                    {
                        eventLog1.WriteEntry("execute_125_ej_Get_LineAsText connected_ToDevice: " + serv.lastError_Code);
                        return Result;
                    }
                    if (serv.set_InputParam_ByName(cmd, "Option", Option) != 0)
                    {
                        eventLog1.WriteEntry("execute_125_ej_Get_LineAsText Option: " + serv.lastError_Code);
                        return Result;
                    }
                    if (serv.execute_Command_ByName(cmd) != 0)
                    {
                        if (serv.lastError_Code != -100003) // no more data
                            eventLog1.WriteEntry("execute_125_ej_Get_LineAsText execute: " + serv.lastError_Code);
                        return Result;
                    }
                    serv.get_OutputParam_ByName(cmd, "TextData", ref TextData);
                }
                catch (Exception ex)
                {
                    eventLog1.WriteEntry("execute_124_ej_Search_Documents_ByDate: " + ex.Message);
                }
            }
            finally
            {
                Result = serv.lastError_Code;
                serv.get_OutputParam_ByName(cmd, "ErrorCode", ErrorCode);
            }
            return Result;
        }

        private int OpenConnection(string ecrIp, string ecrPort)
        {
            int lanPort;
            int error_Code = 0;

            if (serv == null) return -1995;
            
            if (serv.connected_ToDevice)
            {
                if (serv.close_Connection() != 0)
                {
                    eventLog1.WriteEntry(serv.lastError_Message);
                    return -1995;
                }
            }
            try
            {
                
                error_Code = serv.set_TransportType(TTransportProtocol.ctc_TCPIP);
                if (error_Code != 0) return error_Code;

                lanPort = Int32.Parse(ecrPort);
                error_Code = serv.set_TCPIP(ecrIp, (ushort)lanPort);
                if (error_Code != 0) return error_Code;

                error_Code = serv.open_Connection();
                if (error_Code != 0) return error_Code;
                return 0;
            }
            catch (Exception ex)
            {
                eventLog1.WriteEntry("Open connection failed: " + ex.Message);
                return -1995;
            }
            finally
            {
                if (error_Code != 0)
                    eventLog1.WriteEntry(serv.lastError_Message);
            }
        }

        private void StopConnection()
        {
            int error_Code = serv.close_Connection();
            if (error_Code != 0)
                eventLog1.WriteEntry(serv.lastError_Message);
        }

        private void Start_COMServer()
        {
            try
            {
                serv = new CFD_DUDE();
                eventLog1.WriteEntry("Start_COMServer");
            }
            catch (Exception t)
            {
                eventLog1.WriteEntry(t.Message);
            }
        }

        private int Stop_COMServer()
        {
            try
            {
                try
                {
                    if (serv == null)
                        return 0;

                    if (serv.connected_ToDevice) return serv.close_Connection();
                    else return 0;
                }
                catch (Exception t)
                {
                    eventLog1.WriteEntry(t.Message);
                    return -1;
                }
            }
            finally
            {
                if (serv != null)
                {
                    while (System.Runtime.InteropServices.Marshal.ReleaseComObject(serv) > 0) ;
                    //technically the final release and GC. calls are neither needed nor recommended, the framework will dispose the instances when needed,
                    //but leaving here for the sake of showing how to release the com server right away (for example when update is required)
                    while (System.Runtime.InteropServices.Marshal.FinalReleaseComObject(serv) > 0) ;
                    serv = null;
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                }
            }
        }

        private void fileSystemWatcher1_Changed(object sender, FileSystemEventArgs e)
        {

        }
    }
}
