using NLog;
using System;
using System.Linq;
using System.Threading;
using System.Windows.Forms;

namespace MaskedExtensionControl
{
    static class Program
    {
        public static Mutex MutexValue { get; set; }

        public static readonly Logger Logger = LogManager.GetCurrentClassLogger();

        public static string FileParam;

        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main(string[] args)
        {
            if (args == null || args.Length == 0)
                FileParam = string.Empty;
            else
            {
                foreach (var item in args.ToList())
                {
                    FileParam += item + " "; 
                }

                FileParam.Replace(" ", string.Empty);
                Logger.Info("Param: " + FileParam);
            }

            var first = false;
            MutexValue = new Mutex(true, Application.ProductName.ToString(), out first);

            if (!first) return;

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new frmInit());
            MutexValue.ReleaseMutex();
        }
    }
}
