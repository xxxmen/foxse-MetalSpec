using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Text;

namespace MetalSpec.DataAdapter
{
    public class ConstructionType : INotifyPropertyChanged
    {
        public int ID { get; set; }
        public string Name { get; set; }

        public bool IsModified { get; set; } = false;

        public ObservableCollection<Construction> Constructions { get; set; }

        void OnPropertyChanged(String prop)
        {
            PropertyChangedEventHandler handler = PropertyChanged;

            if (handler != null)
            {
                PropertyChanged(this, new PropertyChangedEventArgs(prop));
            }
        }
        public event PropertyChangedEventHandler PropertyChanged;
    }
}
