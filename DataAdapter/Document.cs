using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Text;

namespace MetalSpec.DataAdapter
{
    public interface IDocument
    {
        int ID { get; set; }
        string Name { get; set; }
        void Update(IDocument document);
    }

    public class Document : IDocument, INotifyPropertyChanged
    {
        [Newtonsoft.Json.JsonIgnore]
        public Dictionary<string, ObservableCollection<string>> AllProfilesItems { get; set; }

        public Document(Dictionary<string, ObservableCollection<string>> profilesItems)
        {
            AllProfilesItems = profilesItems;
        }

        public Document()
        {
        }

        public void RemoveConstruction(int constructionID)
        {
            //   int c = 5;
            if (Materials != null)
                foreach (var mater in Materials)
                {
                    if (mater.Profiles != null)
                    {
                        foreach (var prof in mater.Profiles)
                        {
                            for (int i = 0; i < prof.ConstructionTypes[0].Constructions.Count; i++)
                            {
                                if (prof.ConstructionTypes[0].Constructions[i].ID == constructionID)
                                {
                                    prof.ConstructionTypes[0].Constructions.RemoveAt(i);
                                    i--;
                                }
                                else
                                {
                                    prof.ConstructionTypes[0].Constructions[i].ID = i + 5;

                                }
                            }
                        }
                        //хак для обновления интерфейса
                        if (mater.Profiles.Count > 0 && mater.Profiles[0].ConstructionTypes.Count > 0 && mater.Profiles[0].ConstructionTypes[0].Constructions.Count > 0)
                        {
                            var oldCount = mater.Profiles[0].ConstructionTypes[0].Constructions[0].Count;
                            mater.Profiles[0].ConstructionTypes[0].Constructions[0].Count = 0;
                            mater.Profiles[0].ConstructionTypes[0].Constructions[0].Count = oldCount;
                        }
                    }
                }
            OnPropertyChanged("TotalWeight");
        }

        public bool IsModified { get; set; } = false;

        private string name;

        public string Name
        {
            get { return name; }
            set
            {
                if (Materials != null)
                    foreach (var mater in Materials)
                    {
                        if (mater.Profiles != null)
                            foreach (var prof in mater.Profiles)
                            {
                                if (prof.ProfilesItems != null)
                                {
                                    prof.ProfilesItems.Clear();
                                    foreach (var item in AllProfilesItems[value])
                                    {
                                        prof.ProfilesItems.Add(item);
                                    }
                                }
                               // prof.Name = "";
                            }

                    }
                if (AllProfilesItems != null && AllProfilesItems.ContainsKey(value))
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
                OnPropertyChanged("Tooltip");
                OnPropertyChanged("CellBackgroundColor");
                name = value;
            }
        }


        public Color CellBackgroundColor { get; set; }
        public string Tooltip { get; set; }

        public int ID { get; set; }

        public int lines { get; set; }
        public int linesMin { get; set; }

        public List<double?> ProfilesTotal
        {
            get
            {
                var result = new List<double?>();
                if (Materials != null)
                    for (int i = 0; i < Materials.Count; i++)
                    {
                        var m = Materials[i];

                        if (result.Count == 0)
                        {
                            result.AddRange(m.TotalConstructionsWeights);
                        }
                        else
                        {
                            for (int j = 0; j < m.TotalConstructionsWeights.Count; j++)
                            {
                                if (m.TotalConstructionsWeights[j] != null)
                                {
                                    if (result[j] == null)
                                        result[j] = 0;
                                    result[j] += m.TotalConstructionsWeights[j];
                                }
                            }
                        }
                    }
                return result;
            }
        }

        public int ProfilesCount
        {
            get
            {
                int result = 0;
                if (Materials != null)
                    foreach (var mat in Materials)
                    {
                        result += mat.Profiles.Count;
                    }

                return result;
            }
        }

        public double? TotalWeight
        {
            get
            {
                double total = 0;
                foreach (var item in ProfilesTotal)
                {
                    if (item != null)
                        total += (double)item;
                }

                if (total == 0)
                    return null;
                return total;
               
            }
        }

        public bool IsSelected { get; set; }

        /// <summary>
        /// Профили, попадающие под этот документ
        /// </summary>
        public ObservableCollection<Material> Materials { get; set; }

        public void Update(IDocument document)
        {
            ID = document.ID;
            Name = document.Name;
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
