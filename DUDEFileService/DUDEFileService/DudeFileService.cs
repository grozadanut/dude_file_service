using dude;
using System;
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

namespace DUDEFileService
{
    public partial class DudeFileService : ServiceBase
    {
        public static readonly String ECR_FOLDER_KEY = "ecr_folder";

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
            Stop_COMServer();
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

            if (lines.Length < 3)
                return;

            ecrIp = lines[0];
            ecrPort = lines[1];

            for (int i = 2; i < lines.Length; i++)
                ecrCommands += lines[i] + Environment.NewLine;

            OpenConnection(ecrIp, ecrPort);
            if (ecrCommands.StartsWith("raportmf"))
            {
                string[] command = ecrCommands.Split('&');// raportmf&start&end&directory
                DownloadAnafXML(command[1], command[2], command[3]);
            }
            else
                ExecuteScript(ecrCommands);
            StopConnection();

            // delete the executed print file
            File.Delete(e.FullPath);
        }


        private void ExecuteScript(string cmd_Script)
        {
            while (inCommand)
                System.Threading.Thread.Sleep(100);

            inCommand = true;

            ThreadPool.QueueUserWorkItem(delegate
            {
                try
                {
                    serv.execute_Script_V1(TScriptType.DS, cmd_Script);
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
        private void DownloadAnafXML(string startDateTime, string endDateTime, string chosenDirectory)
        {
            while (inCommand)
                System.Threading.Thread.Sleep(100);

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
                int error_Code = serv.set_Download_Path(chosenDirectory);
                if (error_Code != 0) return;
                serv.DateRange_StartValue = startDateTime;
                serv.DateRange_EndValue = endDateTime;

                inCommand = true;

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
                        serv.set_CommunicationEvents(Old_Active_OnSendCommand, Old_Active_OnWait, Old_Active_OnReceiveAnswer, Old_Active_OnStatusChange, Old_Active_OnError, false);
                    }
                }, null);
            }
            finally
            {

            }
        }

        private void OpenConnection(string ecrIp, string ecrPort)
        {
            int lanPort;
            int error_Code = 0;

            if (serv == null) return;
            
            if (serv.connected_ToDevice)
            {
                if (serv.close_Connection() != 0)
                {
                    eventLog1.WriteEntry(serv.lastError_Message);
                    return;
                }
            }
            try
            {
                
                error_Code = serv.set_TransportType(TTransportProtocol.ctc_TCPIP);
                if (error_Code != 0) return;

                lanPort = Int32.Parse(ecrPort);
                error_Code = serv.set_TCPIP(ecrIp, (ushort)lanPort);
                if (error_Code != 0) return;

                error_Code = serv.open_Connection();
                if (error_Code != 0) return;
            }
            catch (Exception ex)
            {
                eventLog1.WriteEntry("Open connection failed: " + ex.Message);
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
    }
}
