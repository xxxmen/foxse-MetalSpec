using MetalSpec.DataAdapter;
using MetalSpec.Model;
using MetalSpec.SpecGenerator;
using MetalSpec.UI.WPF;
using System.Windows;
using MetalSpec.View;
using MetalSpec.ViewModel;
using System.Collections.Generic;

namespace MetalSpec
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private IProjectsModel _projectModel;

        public MainWindow(string[] args)
        {
            InitializeComponent();
            //_projectModel = new ProjectsModel(
            //    new DataServiceStub());
            //_idm = new DocumentsModel(
            //    new DataServiceStub());

            //            stvm = new SpecTableViewModel();

            SpecTableView stv = new SpecTableView();
            stv.Show();
            // stv.DataContext = stvm;
            //stv.specHeadDataGrid.ItemsSource = new List<SpecTableViewModel>() {DataContext as SpecTableViewModel};
            //stv.specListBox.ItemsSource = stvm.Documents;
            if (args.Length > 0)
               // args = new string[] { @"C:\Users\EBabaev\Desktop\TeklaMetalSpec.json" };
                ((SpecTableViewModel)stv.Resources["sh"]).LoadSpecFile(args[0], @"D:\");

            this.Close();

            //CreateConstruction ccv = new CreateConstruction();
            //ccv.Show();
        }

        private void ShowProjectsButton_Click(object sender,
            RoutedEventArgs e)
        {
            ProjectsView view = new ProjectsView();
            view.DataContext
                = new ProjectsViewModel(_projectModel);
            view.Owner = this;
            view.Show();
        }
    }
}
