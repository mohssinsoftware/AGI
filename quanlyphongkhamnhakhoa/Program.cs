using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PKNK_CNPM
{
    public static class AGI_DB
    {
        public static string PathDataBase = File.ReadAllText("FilePathDataBase.txt");
    }

    internal static class GeneratedFile
    {
        public static string LogoPath
        {
            get
            {
                string logoPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Logo_AGI_page-0001-removebg-preview.png");
                if (File.Exists(logoPath))
                {
                    return logoPath;
                }

                return Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"..\..\..\Logo_AGI_page-0001-removebg-preview.png"));
            }
        }

        public static string GetOutputPath(string directory, string fileName)
        {
            if (string.IsNullOrWhiteSpace(directory))
            {
                throw new InvalidOperationException("Le dossier de destination n'est pas configuré.");
            }

            if (string.IsNullOrWhiteSpace(fileName))
            {
                throw new ArgumentException("Le nom du fichier est obligatoire.", "fileName");
            }

            foreach (char invalidCharacter in Path.GetInvalidFileNameChars())
            {
                fileName = fileName.Replace(invalidCharacter, '_');
            }

            Directory.CreateDirectory(directory);
            return Path.Combine(directory, fileName);
        }
    }

    internal static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            //Application.Run(new frmLoginPage());frmHomeScreen

            Application.Run(new frmHomeScreen());
        }
    }
}