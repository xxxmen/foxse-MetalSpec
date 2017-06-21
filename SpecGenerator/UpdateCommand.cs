using System;
using System.ComponentModel;
using System.Windows.Input;

namespace MetalSpec.SpecGenerator
{
    internal class UpdateCommand : ICommand
    {
        private const int ARE_EQUAL = 0;
        private const int NONE_SELECTED = -1;
        private IProjectsViewModel _vm;
        private IDocumentsViewModel _dvm;

        public UpdateCommand(IProjectsViewModel viewModel)
        {
            _vm = viewModel;
            _vm.PropertyChanged += vm_PropertyChanged;
        }

        public UpdateCommand(IDocumentsViewModel viewModel)
        {
            _dvm = viewModel;
            _dvm.PropertyChanged += vm_PropertyChanged;
        }

        private void vm_PropertyChanged(object sender,
            PropertyChangedEventArgs e)
        {
            if (string.Compare(e.PropertyName,
                               ProjectsViewModel.
                               SELECTED_PROJECT_PROPERRTY_NAME)
                == ARE_EQUAL)
            {
                CanExecuteChanged(this, new EventArgs());
            }
        }

        public bool CanExecute(object parameter)
        {
            if (_vm.SelectedProject == null)
                return false;
            return ((ProjectViewModel)_vm.SelectedProject).ID
                   > NONE_SELECTED;
        }

        public event EventHandler CanExecuteChanged
            = delegate { };

        public void Execute(object parameter)
        {
            _vm.UpdateProject();
        }
    }
}
