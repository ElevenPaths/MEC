using Microsoft.Win32;
using NLog;
using System;
using System.IO;

namespace MaskedExtensionControlRegistry
{
    class Program
    {
        public static readonly Logger Logger = LogManager.GetCurrentClassLogger();

        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main(string[] args)
        {
            try
            {
                string[] lines = File.ReadAllLines("BackRegValues.txt");

                foreach (var item in lines)
                {
                    var lineValues = item.Split(',');
                    var regPath = lineValues[0];
                    var regValue = lineValues[1];

                    ModifiedRegistry(regPath, regValue);

                    Logger.Info(item);
                }
                
                Logger.Info("Uninstaller success");
            }
            catch (Exception ex)
            {
                Logger.Info(ex);
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

