using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;

namespace DUDEFileService
{
    public partial class DudeFileService : ServiceBase
    {
        public DudeFileService()
        {
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {
        }

        protected override void OnStop()
        {
        }

        private void FileSystemWatcher_Created(object sender, FileSystemEventArgs e)
        {
            string ecrCommand = System.IO.File.ReadAllText(e.FullPath);

            //eventLog1.WriteEntry("File created: " + e.FullPath);
            //eventLog1.WriteEntry("content: " + ecrCommand);

            File.Delete(e.FullPath);
        }
    }
}
