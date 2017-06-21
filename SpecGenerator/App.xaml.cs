using MetalSpec;
using System.Windows;

namespace SpecGenerator
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : Application
    {
        protected void Application_Startup(object sender, StartupEventArgs e)
        {
            MainWindow wnd = new MainWindow(e.Args);

            //wnd.Show();
        }
    }
}
