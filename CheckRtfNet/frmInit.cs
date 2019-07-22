using MaskedExtensionControl.Properties;
using Microsoft.Office.Interop.Word;
using Microsoft.Win32;
using NLog;
using System;
using System.Diagnostics;
using System.IO;
using System.Net;
using System.Security.Principal;
using System.Windows.Forms;
using Application = System.Windows.Forms.Application;
using Version = System.Version;

namespace MaskedExtensionControl
{
    public partial class frmInit : Form
    {
        public frmInit()
        {
            InitializeComponent();
        }

        public static readonly Logger Logger = LogManager.GetCurrentClassLogger();

        private void frmCapture_Load(object sender, EventArgs e)
        {
            try
            {
                if (string.IsNullOrEmpty(Program.FileParam))
                {
                    SetTitleVersion();
                    LoadAboutText();
                    SetButtonsState();
                    return;
                }

                Logger.Info("Initialize MEC: " + Program.FileParam);

                var resultMagicNumber = GetFileSignature();

                if (string.IsNullOrEmpty(resultMagicNumber))
                {
                    Logger.Info("File signature is empty. File created in Explorer.exe");
                    if (true)
                    {
                        Microsoft.Office.Interop.Word.Application ap = new Microsoft.Office.Interop.Word.Application
                        {
                            Visible = true
                        };
                        var document = ap.Documents.Open(Program.FileParam);
                    }
                }
                else
                    ProcessMagicNumberResult(resultMagicNumber);

                Program.FileParam = string.Empty;
            }
            catch (Exception ex)
            {
                Logger.Error(ex);
            }

            Application.Exit();
        }

        private void SetButtonsState()
        {
            var pathComplete = Environment.GetEnvironmentVariable("ProgramFiles") + @"\ElevenPaths\Mec";

            if (Directory.Exists(pathComplete))
            {
                btnInstall.Enabled = false;
            }
            else
                btnUninstall.Enabled = false;
            
        }


        /// <summary>
        /// Set title in App with the last version.
        /// </summary>
        private void SetTitleVersion()
        {
            var valueVersion = "About MEC " + ProductVersion;
            this.Text = valueVersion;
        }

        /// <summary>
        /// Read from resources the value for About.
        /// </summary>
        private void LoadAboutText()
        {
            richAbout.Rtf = Resources.About;
        }

        /// <summary>
        /// Process magic number.
        /// </summary>
        /// <param name="resultMagicNumber"></param>
        private static void ProcessMagicNumberResult(string resultMagicNumber)
        {
            switch (resultMagicNumber)
            {
                case "docx":
                case "doc":
                    var ap = new Microsoft.Office.Interop.Word.Application();
                    ap.Visible = true;
                    ap.Documents.Open(Program.FileParam);
                    break;
                case "rtf":
                    Logger.Info("RTF File: " + Program.FileParam);

                    var fileName = Program.FileParam;
                    var startInfo = new ProcessStartInfo
                    {
                        FileName = "wordpad.exe",
                        Arguments = "\"" + fileName + "\""
                    };

                    Process.Start(startInfo);
                    break;
                case "mht":
                    //Info - The file extension is modified to .mht
                    var renameFile = Path.ChangeExtension(Program.FileParam, ".mht");
                    File.Move(Program.FileParam, renameFile);

                    var psi = new ProcessStartInfo();
                    psi.FileName = GetDefaultBrowser();
                    psi.Arguments = "\"" + renameFile + "\"";
                    Process.Start(psi);
                    break;
                case "xml":
                    MessageBox.Show("This is an obsolete format for documents. Some malware use it to hide itself. If you think this is a legitimate document, ask the one that generated it to save it in another format.");
                    break;
                default:
                    var appByExtension = Signatures.ResourceManager.GetString(resultMagicNumber);
                    MessageBox.Show("The file you are trying to open can't be opened with Microsoft Word: " + appByExtension);
                    break;
            }
        }

        /// <summary>
        /// Return file signature.
        /// </summary>
        /// <returns>string</returns>
        private static string GetFileSignature()
        {
            byte[] buffer;

            using (var fs = new FileStream(Program.FileParam, FileMode.Open, FileAccess.Read))
            using (var reader = new BinaryReader(fs))
                buffer = reader.ReadBytes(4);

            var hex = BitConverter.ToString(buffer);

            var value = hex.Replace("-", string.Empty).ToLower();

            string output = null;

            //Info - All existing filesignature are in the files Properties\Signatures.resx

            switch (value)
            {
                case "7b5c7274":
                case "4f726967":
                case "446f6320":
                    output = "rtf";
                    break;
                case "d0cf11e0":
                    output = "doc";
                    break;
                case "504b0304":
                    output = "docx";
                    break;
                case "46726f6d":
                    output = "mht";
                    break;
                case "3c3f786d":
                    output = "xml";
                    break;
                default:
                    output = hex.Replace("-", " ");
                    break;
            }

            return output;
        }

        /// <summary>
        /// Return default browser.
        /// </summary>
        /// <returns>Return string default browser.</returns>
        public static string GetDefaultBrowser()
        {
            var browserName = "iexplore.exe";
            using (var userChoiceKey = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Windows\Shell\Associations\UrlAssociations\http\UserChoice"))
            {
                if (userChoiceKey == null) return browserName;

                object progIdValue = userChoiceKey.GetValue("Progid");

                if (progIdValue == null) return browserName;

                if (progIdValue.ToString().ToLower().Contains("chrome"))
                    browserName = "chrome.exe";
                else if (progIdValue.ToString().ToLower().Contains("firefox"))
                    browserName = "firefox.exe";
                else if (progIdValue.ToString().ToLower().Contains("safari"))
                    browserName = "safari.exe";
                else if (progIdValue.ToString().ToLower().Contains("opera"))
                    browserName = "opera.exe";
            }

            return browserName;
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void lnkEleven_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            lnkEleven.LinkVisited = true;
            Process.Start("https://www.elevenpaths.com/");
        }

        private void btnInstall_Click(object sender, EventArgs e)
        {
            try
            {
                if (!IsAdministrator())
                    MessageBox.Show("Administrator permissions are required to run the installation", "Message", MessageBoxButtons.OK);
                else
                    InstallMec();
            }
            catch (Exception ex)
            {
                Logger.Error(ex);
            }
        }

        private void InstallMec()
        {
            var pathComplete = Environment.GetEnvironmentVariable("ProgramFiles") + @"\ElevenPaths\Mec";

            if (!Directory.Exists(pathComplete))
                Directory.CreateDirectory(pathComplete);

            foreach (var file in Directory.GetFiles(Environment.CurrentDirectory))
                File.Copy(file, Path.Combine(pathComplete, Path.GetFileName(file)));

            new OptsRegedit().InitializeRegedit();

            MessageBox.Show("Installation complete, this window will close itself in few seconds.", "MEC", MessageBoxButtons.OK);

            System.Threading.Thread.Sleep(1000);

            this.Close();

            btnInstall.Enabled = false;
            btnUninstall.Enabled = true;
        }

        public static bool IsAdministrator()
        {
            WindowsIdentity identity = WindowsIdentity.GetCurrent();
            WindowsPrincipal principal = new WindowsPrincipal(identity);
            return principal.IsInRole(WindowsBuiltInRole.Administrator);
        }
        
        private void btnUninstall_Click(object sender, EventArgs e)
        {

            if (!IsAdministrator())
                MessageBox.Show("Administrator permissions are required to run the uninstall", "Message", MessageBoxButtons.OK);
            else
            {
                try
                {
                    var pathBack = Environment.GetEnvironmentVariable("ProgramFiles") + @"\ElevenPaths\Mec\";

                    string[] lines;

                    if (File.Exists(pathBack + "BackRegValues.txt"))
                        lines = File.ReadAllLines(pathBack + "BackRegValues.txt");
                    else
                        lines = File.ReadAllLines(pathBack + "DefaultRegValues.txt");

                    foreach (var item in lines)
                    {
                        if (string.IsNullOrEmpty(item))
                            return;

                        var lineValues = item.Split(',');
                        var regPath = lineValues[0];
                        var regValue = lineValues[1];

                        ModifiedRegistry(regPath, regValue);

                        Logger.Info(item);
                    }

                    Logger.Info("Uninstaller success");

                    Directory.Delete(pathBack, true);

                    this.Close();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Some elements could not be removed, please remove them manually", "MEC", MessageBoxButtons.OK);
                }
               
            }

        }

        /// <summary>
        /// Change the registry value.
        /// </summary>
        /// <param name="regPath">Path in reg.</param>
        /// <param name="regValue">Value for the path.</param>
        private static void ModifiedRegistry(string regPath, string regValue)
        {
            try
            {
                var regKey = Registry.LocalMachine.OpenSubKey(regPath, true);
                regKey.SetValue("", regValue);
            }
            catch (Exception ex)
            {
                Logger.Error(ex);
            }
        }
    }
}
