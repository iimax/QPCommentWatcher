using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Permissions;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace QPCommentWatcher
{
    public partial class FrmMain : Form
    {
        private static string folder = System.Configuration.ConfigurationManager.AppSettings["whichFolder"];
        private FileSystemWatcher watcher = null;
        public const string SublimeTextPath = "C:\\Program Files\\Sublime Text 3\\sublime_text.exe";
        private string baseTitle = string.Empty;

        [DllImport("user32.dll")]
        internal static extern bool OpenClipboard(IntPtr hWndNewOwner);

        [DllImport("user32.dll")]
        internal static extern bool CloseClipboard();

        [DllImport("user32.dll")]
        internal static extern bool SetClipboardData(uint uFormat, IntPtr data);

        public FrmMain()
        {
            InitializeComponent();
            baseTitle = this.Text;
            if (string.IsNullOrWhiteSpace(folder) || !Directory.Exists(folder))
            {
                btnWatch.Enabled = false;
                UpdateTitle("Disabled(Invalid folder)");
            }
        }


        private void FrmMain_Load(object sender, EventArgs e)
        {
            if (btnWatch.Enabled)
            {
                this.groupBox1.Text = string.Format("Wathing {0}", folder);
                StartWatch();
            }
            
        }

        private void StartWatch()
        {
            Run(folder);
            UpdateTitle("Running");
            this.btnWatch.Text = "Stop";
        }

        private void btnWatch_Click(object sender, EventArgs e)
        {
            var action = this.btnWatch.Text;
            switch (action)
            {
                case "Start":
                    StartWatch();
                    break;
                default:
                    Stop();
                    break;
            }
        }

        private void FrmMain_FormClosing(object sender, FormClosingEventArgs e)
        {
            Stop();
        }

        private void UpdateTitle(string msg)
        {
            this.Text = string.Format("{0} - {1}", baseTitle, msg);
        }

        private void UpdateStatus(string msg)
        {
            lblStatus.Text = msg;
        }

        [PermissionSet(SecurityAction.Demand, Name = "FullTrust")]
        public void Run(string dir)
        {
            // Create a new FileSystemWatcher and set its properties.
            watcher = new FileSystemWatcher();
            watcher.Path = dir;
            /* Watch for changes in LastAccess and LastWrite times, and
               the renaming of files or directories. */
            //watcher.NotifyFilter = NotifyFilters.FileName | NotifyFilters.DirectoryName;
            // Only watch text files.
            watcher.Filter = "*.txt";

            // Add event handlers.
            //watcher.Changed += new FileSystemEventHandler(OnChanged);
            watcher.Created += new FileSystemEventHandler(OnChanged);
            //watcher.Deleted += new FileSystemEventHandler(OnChanged);
            //watcher.Renamed += new RenamedEventHandler(OnRenamed);

            watcher.IncludeSubdirectories = false;
            // Begin watching.
            watcher.EnableRaisingEvents = true;

        }

        // Define the event handlers.
        private void OnChanged(object source, FileSystemEventArgs e)
        {
            // Specify what is done when a file is changed, created, or deleted.
            if (e.ChangeType == WatcherChangeTypes.Created)
            {
                try
                {
                    //Clipboard.SetText(e.Name);
                    var fileName = Path.GetFileNameWithoutExtension(e.FullPath);
                    SetClipboardText(fileName);
                    UpdateStatus(fileName);
                }
                catch (Exception ex)
                {
                    UpdateStatus(ex.Message);
                }

                try
                {
                    Process.Start(SublimeTextPath, string.Format("\"{0}\"", e.FullPath));
                }
                catch (Exception ex)
                {
                    
                }
                
                
            }
            //Console.WriteLine("File: " + e.FullPath + " " + e.ChangeType);
        }

        private void OnRenamed(object source, RenamedEventArgs e)
        {
            // Specify what is done when a file is renamed.
            Console.WriteLine("File: {0} renamed to {1}", e.OldFullPath, e.FullPath);
        }

        public void Stop()
        {
            if (watcher != null)
            {
                watcher.EnableRaisingEvents = false;
            }
            btnWatch.Text = "Start";
            UpdateTitle("Sopped");
        }

        private void SetClipboardText(string text)
        {
            //https://stackoverflow.com/questions/14082942/copy-result-to-clipboard
            OpenClipboard(IntPtr.Zero);
            var yourString = text;
            var ptr = Marshal.StringToHGlobalUni(yourString);
            SetClipboardData(13, ptr);
            CloseClipboard();
            Marshal.FreeHGlobal(ptr);
        }

    }
}
