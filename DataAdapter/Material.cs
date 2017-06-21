using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Drawing;
using System.Linq;

namespace MetalSpec.DataAdapter
{
    public class Material : INotifyPropertyChanged
    {
        public int ID { get; set; }
        private string name;

        public bool IsModified { get; set; } = false;

        public string Name
        {
            get { return name; }
            set
            {
                Tooltip = value;
                OnPropertyChanged("Name");
                OnPropertyChanged("Tooltip");
                name = value;
            }
        }

        public Color CellBackgroundColor { get; set; }
        public string Tooltip { get; set; }

        [Newtonsoft.Json.JsonIgnore]
        public Document Owner { get; set; }

        public Material(Document owner)
        {
            Owner = owner;
        }

        public Material()
        {
        }
        private ObservableCollection<double?> totalConstructionsWeight = new ObservableCollection<double?>();
        public ObservableCollection<double?> TotalConstructionsWeights
        {
            get
            {
                totalConstructionsWeight.Clear();

                Dictionary<string, Dictionary<string, double?>> totalConstructions = new Dictionary<string, Dictionary<string, double?>>();

                if (Profiles != null)
                    foreach (var p in Profiles)
                    {
                        if (p.ConstructionTypes != null)
                        {
                            foreach (var ct in p.ConstructionTypes)
                            {
                                if (ct.Name == null)
                                    ct.Name = "";
                                if (!totalConstructions.ContainsKey(ct.Name))
                                {
                                    totalConstructions.Add(ct.Name, new Dictionary<string, double?>());
                                }

                                foreach (var c in ct.Constructions)
                                {
                                    if (c.Name == null)
                                        c.Name = "";
                                    if (!totalConstructions[ct.Name].ContainsKey(c.Name))
                                    {
                                        totalConstructions[ct.Name].Add(c.Name, null);
                                    }
                                    if (c.Count != null)
                                    {
                                        if (totalConstructions[ct.Name][c.Name] == null)
                                            totalConstructions[ct.Name][c.Name] = 0;
                                        totalConstructions[ct.Name][c.Name] += (double)c.Count;
                                    }
                                }
                            }
                        }

                    }

                foreach (var item in totalConstructions.Values)
                {
                    foreach (var val in item.Values)
                        totalConstructionsWeight.Add(val);
                }

                return totalConstructionsWeight;
            }
        }

        internal void OrderProfiles()
        {
            if (Profiles != null && Profiles.Count > 1)
            {
                int id = Profiles.FirstOrDefault().ID;
                if (Profiles.Where(p => p.Name == null).Count() == 0)
                {
                    var ordered = Profiles.OrderBy(i => i.Name).OrderBy(i => i.Name.Length).ToList();
                    Profiles.Clear();
                    foreach (Profile prof in ordered)
                    {
                        prof.ID = id;
                        id++;
                        Profiles.Add(prof);
                    }
                }
            }
        }

        public double? TotalWeight
        {
            get
            {
                double totalWeight = 0;
                if (Profiles != null)
                    foreach (var p in Profiles)
                    {
                        if (p.TotalWeight != null)
                            totalWeight += (double)p.TotalWeight;
                    }
                if (totalWeight == 0)
                    return null;

                return totalWeight;
            }
        }
        /// <summary>
        /// ID профиля, Количество (т)
        /// </summary>
        public ObservableCollection<Profile> Profiles { get; set; }

        [Newtonsoft.Json.JsonIgnore]
        public bool IsSelected { get; set; }

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
