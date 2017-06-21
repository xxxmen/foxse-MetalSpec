using MetalSpec.DataAdapter;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Input;

namespace MetalSpec.ViewModel
{
    public class CreateConstruction : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;

        private Construction _currentConstruction;

        public int ConstructionCounter = 0;

        public CreateConstruction()
        {
            //CurrentConstruction = ConstructionTypes.FirstOrDefault().Counstructions.FirstOrDefault();
            
        }

        public List<ConstructionType> ConstructionTypes { get; set; }

        public List<Profile> Profiles { get; set; }

        private ICommand _addCommand;

        public Construction CurrentConstruction
        {
            get
            {
                return _currentConstruction;
            }

            set
            {
                _currentConstruction = value;
                this.OnPropertyChanged("CurrentConstruction");
            }
        }

        public ICommand AddCommand
        {
            get
            {
                if (_addCommand == null)
                {
                    _addCommand = new RelayCommand(
                        param => this.NewConstruction(),
                        param => this.CanAdd()
                    );
                }
                return _addCommand;
            }
        }

        private bool CanAdd()
        {
            // Verify command can be executed here
            return true;
        }

        public void NewConstruction()
        {
            //Command execution logic
            CurrentConstruction = new Construction(true, null) { ID = ConstructionCounter + 1, Name = "Новая конструкция" };
        }

        public void OnPropertyChanged(string propertyName)
        {
            if (PropertyChanged != null)
            {
                PropertyChanged(this, new PropertyChangedEventArgs(propertyName));
            }
        }




    }
}
