using Microsoft.Win32;
using NLog;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace MaskedExtensionControl
{
    public class OptsRegedit
    {
        public static readonly Logger Logger = LogManager.GetCurrentClassLogger();

        private const string RegKey = @"SOFTWARE\Classes\Word.Document\CurVer";

        public string regExtDocPath = Environment.GetEnvironmentVariable("ProgramFiles") + @"\ElevenPaths\Mec\Mec.exe" + " %1"; //Application.ExecutablePath + " %1";

        private RegistryKey regKeyEdit;
        private RegistryKey regKeyOpen;

        private RegistryKey regKeyEdit8;
        private RegistryKey regKeyOpen8;

        private RegistryKey regKeyEditRtf8;
        private RegistryKey regKeyOpenRtf8;

        private List<BackupRegistry> ListBackupRegValues;

        public static string result;

        /// <summary>
        /// Initialize Regedit.
        /// </summary>
        public void InitializeRegedit()
        {
            GetWordVersion();
            InitializeKeyRegistry();
            BackupRegValues();
            WriteRegistry();
        }

        /// <summary>
        /// Backup registry values.
        /// </summary>
        private void BackupRegValues()
        {
            try
            {
                ListBackupRegValues = new List<BackupRegistry>();
                ListBackupRegValues.Add(new BackupRegistry() { Path = regKeyEdit.Name.Remove(0,19), Value = regKeyEdit.GetValue("").ToString() });
                ListBackupRegValues.Add(new BackupRegistry() { Path = regKeyOpen.Name.Remove(0, 19), Value = regKeyOpen.GetValue("").ToString() });

                if (regKeyOpen8 != null)
                {
                    ListBackupRegValues.Add(new BackupRegistry() { Path = regKeyEdit8.Name.Remove(0, 19), Value = regKeyEdit8.GetValue("").ToString() });
                    ListBackupRegValues.Add(new BackupRegistry() { Path = regKeyOpen8.Name.Remove(0, 19), Value = regKeyOpen8.GetValue("").ToString() });
                }

                //save command line 12.

                if (regKeyEdit.GetValue("command") != null)
                {
                    ListBackupRegValues.Add(new BackupRegistry() { Path = regKeyEdit.Name.Remove(0, 19), Value = regKeyEdit.GetValue("command").ToString() });
                    ListBackupRegValues.Add(new BackupRegistry() { Path = regKeyOpen.Name.Remove(0, 19), Value = regKeyOpen.GetValue("command").ToString() });
                }

                //save command line 8.

                if (regKeyEdit8.GetValue("command") != null)
                {
                    ListBackupRegValues.Add(new BackupRegistry() { Path = regKeyEdit8.Name.Remove(0, 19), Value = regKeyEdit8.GetValue("command").ToString() });
                    ListBackupRegValues.Add(new BackupRegistry() { Path = regKeyOpen8.Name.Remove(0, 19), Value = regKeyOpen8.GetValue("command").ToString() });
                }

                var pathBack = Environment.GetEnvironmentVariable("ProgramFiles") + @"\ElevenPaths\Mec\";
                File.WriteAllLines(pathBack + "BackRegValues.txt", ListBackupRegValues.Select(x => x.Path + "," + x.Value ).ToArray());

            }
            catch (Exception ex)
            {
                Logger.Error(ex.Message);
            }

        }

        /// <summary>
        /// Gets the component's path from the registry. if it can't find it - returns an empty string
        /// </summary>
        /// <param name="component"></param>
        /// <returns></returns>
        private static void GetWordVersion()
        {
            var nroVersion = string.Empty;

            try
            {
                var resultValue = string.Empty;

                var keyCurrentUser = Registry.LocalMachine;

                keyCurrentUser = keyCurrentUser.OpenSubKey(RegKey + "\\", false);

                resultValue = keyCurrentUser.GetValue(string.Empty).ToString();

                nroVersion = resultValue.Split('.').Last();

                if (keyCurrentUser != null)
                    keyCurrentUser.Close();

                result = nroVersion;

            }
            catch (Exception ex)
            {
                Logger.Error(ex);
            }
        }

        private void InitializeKeyRegistry()
        {
            try
            {
                Logger.Info("Initialize Keys");

                var regEditWord = @"SOFTWARE\Classes\Word.Document." + result + @"\shell\Edit\command";
                var regOpenWord = @"SOFTWARE\Classes\Word.Document." + result + @"\shell\Open\command";

                var regEditRtf8 = @"SOFTWARE\Classes\Word.Rtf.8\shell\Edit\command";
                var regOpenRtf8 = @"SOFTWARE\Classes\Word.Rtf.8\shell\Open\command";

                regKeyEdit = Registry.LocalMachine.OpenSubKey(regEditWord, true);
                regKeyOpen = Registry.LocalMachine.OpenSubKey(regOpenWord, true);

                regKeyEditRtf8 = Registry.LocalMachine.OpenSubKey(regEditRtf8, true);
                regKeyOpenRtf8 = Registry.LocalMachine.OpenSubKey(regOpenRtf8, true);

                if (!result.Equals("12")) return;

                regKeyEdit8 = Registry.LocalMachine.OpenSubKey(regEditWord.Replace("12", "8"), true);
                regKeyOpen8 = Registry.LocalMachine.OpenSubKey(regOpenWord.Replace("12", "8"), true);
            }
            catch (Exception ex)
            {
                Logger.Error(ex);
            }
        }

        /// <summary>
        /// Write new values in the registry.
        /// </summary>
        private void WriteRegistry()
        {
            try
            {
                Logger.Info("Write new value App path: " + regExtDocPath);
                regKeyEdit.SetValue("", regExtDocPath);
                regKeyOpen.SetValue("", regExtDocPath);

                regKeyEditRtf8.SetValue("", regExtDocPath);
                regKeyOpenRtf8.SetValue("", regExtDocPath);
                
                if (result.Equals("12"))
                {
                    regKeyEdit8.SetValue("", regExtDocPath);
                    regKeyOpen8.SetValue("", regExtDocPath);
                }

                DeleteCommandReg();
            }
            catch (Exception ex)
            {
                Logger.Error(ex);
            }
        }

        /// <summary>
        /// Detele command reg.
        /// </summary>
        private void DeleteCommandReg()
        {
            try
            {
                var existValue8 = regKeyEdit8.GetValue("command");

                if (existValue8 != null)
                {
                    Logger.Info("Deleting command from Word.Document.8");
                    regKeyEdit8.DeleteValue("command");
                    regKeyOpen8.DeleteValue("command");
                }

                if (existValue8 != null)
                {
                    Logger.Info("Deleting command from Word.Document.12");
                    regKeyEdit.DeleteValue("command");
                    regKeyOpen.DeleteValue("command");
                }

                var existValueRtf = regKeyEditRtf8.GetValue("command");

                if (existValueRtf == null) return;

                Logger.Info("Deleting command from RTF");
                regKeyOpenRtf8.DeleteValue("command");
                regKeyEditRtf8.DeleteValue("command");
            }
            catch (Exception ex)
            {
                Logger.Error(ex);
            }
        }
    }

    internal class BackupRegistry
    {
        public string Path { get; set; }

        public string Value { get; set; }
    }
}

