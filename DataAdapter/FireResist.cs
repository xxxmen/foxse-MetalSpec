using System;
using System.ComponentModel;

namespace MetalSpec.DataAdapter
{
    public class FireResist : INotifyPropertyChanged
    {
        [Newtonsoft.Json.JsonIgnore]
        public Construction Owner { get; set; }

        public bool IsModified { get; set; } = false;

        public FireResist(Construction owner)
        {
            Owner = owner;
        }

        public FireResist()
        {

        }

        private string _strCount = "";
        [Newtonsoft.Json.JsonIgnore]
        public string StrCount {
            get { return _strCount; }
            set {
                
                double o = 0;
                if (!double.TryParse(value.Replace(".", ","), out o))
                    double.TryParse(value.Replace(",", "."), out o);
                if (o != 0)
                {
                    Count = Math.Round(o, 2, MidpointRounding.AwayFromZero);
                    _strCount = Count.ToString().Replace(",",".");
                }
                else
                {
                    Count = null;
                   // _strCount = "";
                }

                Owner.Owner.Owner.ChangeProperty("TotalConstructionsWeights");
                Owner.Owner.Owner.ChangeProperty("TotalWeight");
                Owner.Owner.Owner.Owner.ChangeProperty("TotalWeight"); 
                Owner.Owner.Owner.Owner.ChangeProperty("ProfilesTotal"); 
                OnPropertyChanged("StrCount");
            }
        }

        private double? _count;
        public double? Count
        {
            get
            {
                if (_count == 0 || _count == null)
                    return null;
                else
                    return _count;
            }
            set
            {
                _count = value;
                _strCount = (value == null)?"":value.ToString();
            }
        }

        private string name;
        public string Name
        {
            get
            {
                return name;
            }

            set
            {
                name = value;
                OnPropertyChanged("Name");
            }
        }

        void OnPropertyChanged(string prop)
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