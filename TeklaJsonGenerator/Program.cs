using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace TeklaJsonGenerator
{
    static class Program
    {
        /// <summary>
        /// Главная точка входа для приложения.
        /// </summary>
        [DllImport("user32.dll", SetLastError = true)]
        static extern IntPtr FindWindow(IntPtr lpClassName, string lpWindowName);
        [DllImport("user32.dll", SetLastError = true)]
        static extern int ShowWindow(IntPtr h, int f);

        [STAThread]
        static void Main()
        {
            try
            {
                //Process[] TeklaJsonGenerator = Process.GetProcessesByName("TeklaJsonGenerator");
                Process[] TeklaJsonGenerator = Process.GetProcessesByName("TeklaJsonGenerator.vshost");

                if (TeklaJsonGenerator.Length == 1)
                {
                    Application.EnableVisualStyles();
                    Application.SetCompatibleTextRenderingDefault(false);
                    Form1 app = new Form1();
                    app.Show(new TSMainWindowStub());
                    Application.Run(app);
                }
                else
                {
                    IntPtr h = FindWindow(IntPtr.Zero, "TeklaJsonGenerator");
                    if (IntPtr.Zero != h)
                    {
                        ShowWindow(h, 9);
                    }
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message + "\r\n\r\n" + e.StackTrace);
            }
        }

        class TSMainWindowStub : IWin32Window
        {
            public TSMainWindowStub() { }
            public IntPtr Handle
            {
                get
                {
                    Process[] tekla = Process.GetProcessesByName("TeklaStructures");
                    return 0 < tekla.Length ? tekla[0].MainWindowHandle : IntPtr.Zero;
                }
            }
        }
    }
}
