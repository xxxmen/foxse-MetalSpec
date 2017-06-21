using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Text;

namespace MetalSpec.DataAdapter
{
    public interface IProfile
    {
        int ID { get; set; }
        string Name { get; set; }
        string Tag { get; set; }
        bool ProtectedFromCorrosion { get; set; }
        ObservableCollection<ConstructionType> ConstructionTypes { get; set; }
        double? TotalWeight { get; }
        double Weight { get; set; }
        double PaintAreaPerMeter { get; set; }
        void Update(IProfile profile);
    }

    public class Profile : IProfile, INotifyPropertyChanged
    {
        [Newtonsoft.Json.JsonIgnore]
        public Material Owner { get; set; }
        [Newtonsoft.Json.JsonIgnore]
        public ObservableCollection<string> ProfilesItems { get; set; }
        private double? totalWeight;
        public bool IsModified { get; set; } = false;
        private int id;

        public int ID
        {
            get { return id; }
            set
            {
                id = value;
                OnPropertyChanged("ID");
            }
        }
        private string name;

        public string Name
        {
            get { return name; }
            set
            {
                if (ProfilesItems.Contains(value))
                {
                    Tooltip = value;
                    CellBackgroundColor = Color.White;
                }
                else
                {
                    Tooltip = "Элемент не найден в базе";
                    CellBackgroundColor = Color.FromArgb(50, 255, 0, 0);
                }
                OnPropertyChanged("Name");
                OnPropertyChanged("CellBackgroundColor");
                OnPropertyChanged("Tooltip");

                name = value;

                if (Owner != null)
                    Owner.OrderProfiles();
            }
        }

        public string Tag { get; set; }

        private Color cellBackgroundColor;

        [Newtonsoft.Json.JsonIgnore]
        public Color CellBackgroundColor
        {
            get { return cellBackgroundColor; }
            set
            {
                cellBackgroundColor = value;
            }
        }

        private string tooltip;

        [Newtonsoft.Json.JsonIgnore]
        public string Tooltip
        {
            get { return tooltip; }
            set
            {
                // OnPropertyChanged("Tooltip");
                tooltip = value;
            }
        }

        public bool ProtectedFromCorrosion { get; set; } = true;
        public ObservableCollection<ConstructionType> ConstructionTypes { get; set; }

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

        public Profile(Material owner)
        {
            Owner = owner;
            if (Owner.Owner.Name != null)
                ProfilesItems = Owner.Owner.AllProfilesItems[Owner.Owner.Name];
            else
                ProfilesItems = new ObservableCollection<string>();
        }

        public Profile()
        {
            ProfilesItems = new ObservableCollection<string>();
        }

        public double? TotalWeight
        {
            get
            {
                totalWeight = 0;
                if (ConstructionTypes != null && ConstructionTypes != null)
                    foreach (var ct in ConstructionTypes)
                    {
                        foreach (var c in ct.Constructions)
                        {
                            if (c.Count != null)
                            {
                                totalWeight += c.Count;
                            }
                        }
                    }
                else
                    totalWeight = null;
                if (totalWeight == 0)
                    return null;
                return totalWeight;
            }
        }

        public Double TotalArea
        {
            get { return (TotalWeight == null) ? 0 : (double)TotalWeight * PaintAreaPerMeter; }
        }

        public double Weight { get; set; }
        public double PaintAreaPerMeter { get; set; }

        public void Update(IProfile profile)
        {
            ID = profile.ID;
            Name = profile.Name;
            Tag = profile.Tag;
            Weight = profile.Weight;
            PaintAreaPerMeter = profile.PaintAreaPerMeter;
            ConstructionTypes = profile.ConstructionTypes;
        }
    }

}
