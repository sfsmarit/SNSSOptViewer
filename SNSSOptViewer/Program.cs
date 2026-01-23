using System;
using System.IO;
using System.Windows.Forms;

namespace SNSSOptViewer
{
    static class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            string initialRoot = "";

            if (args.Length > 0)
            {
                string p = args[0].Trim().Trim('"');
                if (Directory.Exists(p))
                {
                    initialRoot = p;
                }
            }

            Application.Run(new Form1(initialRoot));
        }
    }
}
