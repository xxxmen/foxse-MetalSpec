using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;

namespace MetalSpec.ViewModel
{
    public class Settings : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;

        private string fileName = ".\\config.json";

        private string _exportDirectory = Environment.GetEnvironmentVariable("USERPROFILE");

        public string ExportDirectory { get { return _exportDirectory; }
            set {
                _exportDirectory = value;
                OnPropertyChanged("ExportDirectory");
            }
        } 

        public bool StampInExcel { get; set; } = true;

        public int UnitDigits { get; set; } = 2;

        public string Units { get; set; } = "т";

        public string LastDirectory { get; set; } = Environment.GetEnvironmentVariable("USERPROFILE");

        public ObservableCollection<string> LastFiles { get; set; } = new ObservableCollection<string>();

        public void Save()
        {
            File.WriteAllText(fileName, JsonConvert.SerializeObject(this, Formatting.Indented));
        }

        public void Load()
        {
            try
            {
                if (File.Exists(fileName))
                {
                    Settings loaded = JsonConvert.DeserializeObject<Settings>(File.ReadAllText(fileName));

                    if (loaded.ExportDirectory != null)
                        ExportDirectory = loaded.ExportDirectory;
                    StampInExcel = loaded.StampInExcel;
                    UnitDigits = loaded.UnitDigits;
                    Units = (loaded.Units == null)? @"т" : loaded.Units;
                    if (loaded.LastDirectory != null)
                        LastDirectory = loaded.LastDirectory;
                    if (loaded.LastFiles != null)
                        LastFiles = loaded.LastFiles;
                }
            }catch
            {

            }
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
