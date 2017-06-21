using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;

namespace MetalSpec.DataAdapter
{
    public class Construction : INotifyPropertyChanged
    {
        public int ID { get; set; }

        private string name;

        public bool IsModified { get; set; } = false;

        public Construction()
        {
        }

        public Construction(Profile owner)
        {
            Owner = owner;
        }

        [JsonIgnore]
        public Profile Owner { get; set; }

        public Construction(bool addFireResist, Profile owner)
        {
            _fireResists.Add(new FireResist(this) { Name = "" });
            Owner = owner;
        }

        public string Name
        {
            get { return name; }
            set
            {
                name = value;
                OnPropertyChanged("Name");
            }
        }

        private ObservableCollection<FireResist> _fireResists = new ObservableCollection<FireResist>();

        public ObservableCollection<FireResist> FireResists
        {
            get
            {
                return _fireResists;
            }

            set
            {
                _fireResists = value;
            }
        }

        // public string FireResist { get; set; }

        double? _count = 0;


        public double? Count
        {
            get
            {
                _count = _fireResists.Sum(i => i.Count);
                if (_count == 0 || _count == null)
                    return null;
                else
                    return Math.Round((double)_count, 2, MidpointRounding.AwayFromZero);
            }
            set
            {
                _count = value;
                OnPropertyChanged("Count");
            }
        }


        public void SaveToJson(string path = "")
        {
            try
            {
                File.WriteAllText(Path.Combine(path, ID.ToString()), JsonConvert.SerializeObject(this));
            }
            catch
            {

            }
        }

        public void ChangeProperty(string prop)
        {
            OnPropertyChanged(prop);
        }

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
