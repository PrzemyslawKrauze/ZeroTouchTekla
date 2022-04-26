using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using Tekla.Structures.Model;

namespace ZeroTouchTekla
{
    public static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new MainForm());
        }
       private static Model activeModel = new Model();

        private static string _excelPath = String.Empty;
        private static string _sheetName = String.Empty;
        private static Dictionary<string, string> _excelDictionary = new Dictionary<string, string>();

        public static string ExcelPath { get => _excelPath; set => _excelPath = value; }
        public static string SheetName { get => _sheetName; set => _sheetName = value; }
        public static Dictionary<string, string> ExcelDictionary { get => _excelDictionary; set => _excelDictionary = value; }
        public static Model ActiveModel { get => activeModel;}
    }
}
