using System;
using System.Windows.Forms;
using C2B_FBR_Connect.Forms;

namespace C2B_FBR_Connect
{
    static class Program
    {
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new MainForm());
        }
    }
}