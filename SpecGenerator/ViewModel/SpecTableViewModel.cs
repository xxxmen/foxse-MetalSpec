using MetalSpec.DataAdapter;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Windows.Input;
using Newtonsoft.Json;
using System.IO;
using System.Windows.Forms;

using sdc = System.Drawing.Color;
using System.Diagnostics;

namespace MetalSpec.ViewModel
{
    public class SpecTableViewModel : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;

        private Document _currentDocument;

        public Settings settings { get; set; } = new Settings();
        /// <summary>
        /// Эталон загруженного документа, с ним сравниваем на наличие изменений.
        /// </summary>
        private Specification spec;

        private string specFilePatch;
        public string SpecFilePatch
        {
            get { return specFilePatch; }
            set
            {
                specFilePatch = value;
                OnPropertyChanged("SpecFilePatch");
            }
        }


        public ObservableCollection<Sortament> DocumentItemsSource { get; set; }
        public ObservableCollection<string> DocumentStringItemsSource { get; set; }

        public ObservableCollection<SpecTableHeader> SpecTableHeaders { get; set; }

        public List<string> MaterialsItemsSource { get; set; }

        public ObservableCollection<string> FireResistItemsSource { get; set; }

        public Dictionary<string, ObservableCollection<string>> ProfilesItems { get; set; }

        public Specification NewSpec = new Specification();

        public ObservableCollection<Document> Documents { get; set; }

        private ObservableCollection<double> totalMetalMass = new ObservableCollection<double>();

        public ObservableCollection<double> TotalMetalMass
        {
            get
            {
                return totalMetalMass;
            }
        }

        public double TotalMassCount
        {
            get
            {
                return totalMetalMass.Sum();
            }
        }

        private Dictionary<string, Dictionary<string, float?>> profilesAreas = new Dictionary<string, Dictionary<string, float?>>();

        public List<Material> Materials { get; set; }

        public int ProfileCounter = 0;

        double? _paintArea = 0;
        public double? PaintArea
        {
            get { return Math.Round((_paintArea == null) ? 0 : (double)_paintArea, 2, MidpointRounding.AwayFromZero); }
            set
            {
                _paintArea = value;
                OnPropertyChanged("PaintArea");
            }
        }

        public int CurrentDocumentIndex { get; set; }

        public int CurrentMaterialIndex { get; set; }

        private int _profileCounter = 1;

        RelayCommand RemoveTagsCommand;
        RelayCommand RemoveConstructionRelayCommand;
        RelayCommand AddMaterialRelayCommand;
        RelayCommand AddProfileRelayCommand;
        RelayCommand AddFireResistRelayCommand;
        RelayCommand OpenFileRelayCommand;

        SpecTableHeader sph;

        StampData stampData = new StampData();

        public ObservableCollection<string> AvailableUnits { get; set; } = new ObservableCollection<string>() { "т", "кг" };

        public SpecTableViewModel()
        {
            settings.Load();
            //Можно задавать названия колонок переменными
            // = "ГОСТ";
            SpecFilePatch = Environment.GetEnvironmentVariable("USERPROFILE") + "\\" + DateTime.Now.ToString("d.MM.yyyy") + "_MetalSpec.json";

            #region получаем источники данных
            MetalSpecEntities mse = new MetalSpecEntities();
            var result = mse.Sortament.ToList();

            if (result != null)
            {
                foreach (var item in result)
                {
                    item.Name = item.Name.Trim();
                    item.Description = item.Description.Trim();
                    //    mse.Sortament.Attach(item);
                    //    var entry = mse.Entry(item);
                    //    entry.Property(e => e.Name).IsModified = true;
                    //    // other changed properties
                    //    mse.SaveChanges();
                }
                DocumentItemsSource = new ObservableCollection<Sortament>(result);
                DocumentStringItemsSource = new ObservableCollection<string>(DocumentItemsSource.Select(i => i.Name.Trim()).ToList());
            }

            //DocumentItemsSource = new List<string>() { "Двутавры стальные горячекатаные с параллельными гранями полок ГОСТ 26020-83", "Швеллеры стальные гнутые равнополочные ГОСТ 8278-83", };

            var querry = mse.Materials.Select(i => (i.Name.Trim() + "\r\n" + i.Description.Trim())).ToList();
            if (querry != null)
            {
                MaterialsItemsSource = querry;
            }
            else
            {
                MaterialsItemsSource = new List<string>() { "C235\r\nГОСТ 19281-89", "C345\r\nГОСТ 19281-89" };
            }


            querry = mse.FireResistTypes.Select(i => i.Name.Trim()).ToList();
            if (querry != null)
            {
                FireResistItemsSource = new ObservableCollection<string>(querry);
            }
            else
            {
                FireResistItemsSource = new ObservableCollection<string>() { "R45", "R60" };
            }

            if (DocumentItemsSource != null)
            {
                ProfilesItems = new Dictionary<string, ObservableCollection<string>>()
                {
                };
                profilesAreas = new Dictionary<string, Dictionary<string, float?>>(0);

                foreach (var doc in DocumentItemsSource)
                {
                    var docName = doc.Name.Trim();

                    ProfilesItems.Add(docName, new ObservableCollection<string>(doc.Profiles.Select(i => i.Name.Trim()).ToList()));


                    var profArera = new Dictionary<string, float?>();

                    foreach (var kvp in doc.Profiles.Select(i => new { i.Name, i.PaintArea }))
                        if (!profArera.ContainsKey(kvp.Name.Trim()))
                            profArera.Add(kvp.Name.Trim(), kvp.PaintArea);

                    if (!profilesAreas.ContainsKey(docName))
                        profilesAreas.Add(docName, profArera);
                }
            }
            else
            {

            }
            sph = new SpecTableHeader();

            sph.ConstructionTypes = new ObservableCollection<ConstructionType>() {
                new ConstructionType()
                {
                 Constructions = new ObservableCollection<Construction>()
                }
            };

            SpecTableHeaders = new ObservableCollection<SpecTableHeader>() { sph };

            #endregion закончили получать источники данных
            //parameterizedCommand = new Command(DoParameterisedCommand);

            Documents = new ObservableCollection<Document>();

            //LoadSpecFile($@"C:\Users\{Environment.UserName}\Desktop\test.json", $@"C:\Users\{Environment.UserName}\Desktop\");

            //LoadSpecFile(SpecFilePatch, $@"C:\Users\{Environment.UserName}\Desktop\");

            CurrentDocumentIndex = 0;

            RemoveTagsCommand = new RelayCommand(RemoveTags);
            RemoveConstructionRelayCommand = new RelayCommand(RemoveConstruction);
            AddMaterialRelayCommand = new RelayCommand(AddMaterial);
            AddProfileRelayCommand = new RelayCommand(AddProfile);
            AddFireResistRelayCommand = new RelayCommand(AddFireResist);
            OpenFileRelayCommand = new RelayCommand(OpenFile);
        }

        private void OpenFile(object filePath)
        {
            LoadSpecFile((string)filePath, settings.ExportDirectory);
        }

        private void AddFireResist(object construction)
        {
            Construction constr = (Construction)construction;
            if (constr.FireResists != null)
            {
                constr.FireResists.Add(new FireResist(constr) { Name = "" });
            }
        }

        private void AddProfile(object material)
        {
            Material mater = (Material)material;
            if (mater.Profiles == null)
                mater.Profiles = new ObservableCollection<Profile>();
            var ctsToAdd = new ObservableCollection<ConstructionType>();

            var profToadd = new Profile(mater)
            {
                ConstructionTypes = ctsToAdd,
                ProtectedFromCorrosion = true,
                CellBackgroundColor = sdc.White,
            };

            foreach (var ct in sph.ConstructionTypes)
            {
                var ctToAdd = new ConstructionType()
                {
                    Name = ct.Name,
                    ID = ct.ID,
                    Constructions = new ObservableCollection<Construction>()
                };
                foreach (var c in ct.Constructions)
                {
                    ctToAdd.Constructions.Add(new Construction(true, profToadd)
                    {
                        ID = c.ID,
                        Name = c.Name
                    });
                }
                ctsToAdd.Add(ctToAdd);
            }
            mater.Profiles.Add(profToadd);

            updateProfileIDs();
        }

        private void updateProfileIDs()
        {
            int i = 0;
            foreach (var doc in Documents)
                if (doc.Materials != null)
                    foreach (var mater in doc.Materials)
                        if (mater.Profiles != null)
                            foreach (var prof in mater.Profiles)
                            {
                                i++;
                                prof.ID = i;
                            }
            _profileCounter = i;
        }

        private void AddMaterial(object document)
        {
            Document doc = (Document)document;

            if (doc.Materials == null)
            {
                doc.Materials = new ObservableCollection<Material>();
            }
            doc.Materials.Add(GetNewMaterial(doc));
            updateProfileIDs();
        }

        private void RemoveConstruction(object construction)
        {
            if (MessageBox.Show("Вы действительно хотите удалить конструкцию?\r\nЭту операцию невозможно отменить!", "Удаление конструкции", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
            {
                Construction constr = (Construction)construction;

                if (sph.ConstructionTypes.Count > 0)
                    if (sph.ConstructionTypes[0].Constructions.Count > 1)
                    {
                        var consts = sph.ConstructionTypes[0].Constructions;

                        consts.Remove(constr);

                        int i = 5;
                        foreach (var cons in consts)
                        {
                            cons.ID = i;
                            i++;
                        }

                        foreach (var doc in Documents)
                        {
                            doc.RemoveConstruction(constr.ID);
                        }
                    }
                OnPropertyChanged("TotalMetalMass");
                OnPropertyChanged("TotalMassCount");
                OnPropertyChanged("ProfilesTotal");
            }
        }

        internal bool IsThisNameExist(string constructionName)
        {
            foreach (var ct in sph.ConstructionTypes)
            {
                foreach (var c in ct.Constructions)
                {
                    if (c.Name == constructionName)
                        return true;
                }
            }
            return false;
        }

        internal void RenameConstriction(string constructionName, int tag)
        {
            foreach (var doc in Documents)
            {
                if (doc.Materials != null)
                    foreach (var mate in doc.Materials)
                    {
                        if (mate.Profiles != null)
                            foreach (var prof in mate.Profiles)
                            {
                                if (prof.ConstructionTypes != null)
                                    foreach (var ct in prof.ConstructionTypes)
                                    {
                                        if (ct.Constructions != null)
                                            foreach (var c in ct.Constructions)
                                            {
                                                if (c.ID == tag)
                                                {
                                                    c.Name = constructionName;
                                                }
                                            }
                                    }
                                //Construction found = prof.ConstructionTypes.FirstOrDefault().Constructions.Where(i => i.ID == tag).FirstOrDefault();
                                //found.Name = constructionName;
                            }
                    }


            }

        }

        private void RemoveTags(object type)
        {
            Type str = type.GetType();
            string strType = str.ToString();
            switch (strType)
            {
                case "MetalSpec.DataAdapter.Document":
                    var doc = (Document)type;
                    //for (int i = 0; i < Documents.Count; i++)
                    //{
                    //    if (Documents[i].IsSelected)
                    //    {
                    //        Documents.RemoveAt(i);
                    //    }
                    //}
                    Documents.Remove(doc);
                    break;
                case "MetalSpec.DataAdapter.Material":
                    var material = (Material)type;
                    material.Owner.Materials.Remove(material);
                    break;
                case "MetalSpec.DataAdapter.Profile":
                    var profile = (Profile)type;
                    profile.Owner.Profiles.Remove(profile);
                    break;
                case "MetalSpec.DataAdapter.FireResist":
                    var fr = (FireResist)type;
                    if (fr.Owner.FireResists.Count > 1)
                    {
                        var mboxresult = MessageBox.Show("Прибавить массу к первому в списке типу огнезащиты для данного профиля?", "Удаление степени огнезащиты профиля.", MessageBoxButtons.YesNoCancel);
                        double frCount = 0;
                        switch (mboxresult)
                        {
                            case DialogResult.Yes:
                                frCount = (fr.Count == null) ? 0 : (double)fr.Count;
                                break;
                            case DialogResult.No:
                                break;
                            default:
                                return;
                        }

                        var owner = fr.Owner;
                        fr.Owner.FireResists.Remove(fr);
                        if (owner.FireResists[0].Count == null && frCount != 0)
                            owner.FireResists[0].Count = frCount;
                        else
                            owner.FireResists[0].Count += frCount;
                    }
                    break;
                default:
                    break;
            }

            updateProfileIDs();
        }

        public void RemoveTags(IList toRemove)
        {
            var collection = toRemove.Cast<Document>();
            List<Document> copy = new List<Document>(collection);

            foreach (Document tag in copy)
            {
                Documents.Remove(tag);
            }
        }


        private ICommand _addDocumentCommand;


        public Document CurrentDocument
        {
            get
            {
                return _currentDocument;
            }

            set
            {
                _currentDocument = value;
                this.OnPropertyChanged("CurrentConstruction");
            }
        }

        public ICommand AddDocumentCommand
        {
            get
            {
                if (_addDocumentCommand == null)
                {
                    _addDocumentCommand = new RelayCommand(
                        param => this.AddNewDocument(),
                        param => this.CanAdd()
                    );
                }
                return _addDocumentCommand;
            }
        }

        private ICommand _removeDocumentCommand;

        public ICommand RemoveCommand
        {
            get
            {
                if (_removeDocumentCommand == null)
                {
                    _removeDocumentCommand = RemoveTagsCommand;
                }
                return _removeDocumentCommand;
            }
        }
        private ICommand _removeConstructionCommand;
        public ICommand RemoveConstructionCommand
        {
            get
            {
                if (_removeConstructionCommand == null)
                {
                    _removeConstructionCommand = RemoveConstructionRelayCommand;
                }
                return _removeConstructionCommand;
            }
        }

        private ICommand _addMaterialCommand;

        public ICommand AddMaterialCommand
        {
            get
            {
                if (_addMaterialCommand == null)
                {
                    _addMaterialCommand = AddMaterialRelayCommand;
                }
                return _addMaterialCommand;
            }
        }

        private ICommand _addProfileCommand;

        public ICommand AddProfileCommand
        {
            get
            {
                if (_addProfileCommand == null)
                {
                    _addProfileCommand = AddProfileRelayCommand;
                }
                return _addProfileCommand;
            }
        }


        private ICommand _saveCommand;

        public ICommand SaveCommand
        {
            get
            {
                if (_saveCommand == null)
                {
                    _saveCommand = new RelayCommand(
                        param => this.SaveSpecification(),
                        param => this.CanAdd()
                    );
                }
                return _saveCommand;
            }
        }

        internal void SaveSpecification()
        {
            if (SpecFilePatch == null && SpecFilePatch != $"{Environment.GetEnvironmentVariable("USERPROFILE")}\\backup.json")
            {
                SaveAsSpecification();
                return;
            }

            SaveSpecification(SpecFilePatch);


        }

        public void SaveSpecification(string file)
        {
            bool specRevised = false;
            #region Проверка на изменения
            if (spec != null)
            {
                for (int i = 0; i < spec.Documents.Count; i++)
                {
                    if (this.Documents.Count > i)
                    {
                        if (spec.Documents[i].Name != this.Documents[i].Name)
                        {
                            this.Documents[i].IsModified = true;
                            specRevised = true;
                        }
                        else
                        {

                            for (int m = 0; m < spec.Documents[i].Materials.Count; m++)
                            {
                                if (this.Documents[i].Materials.Count > m)
                                {
                                    if (spec.Documents[i].Materials[m].Name != this.Documents[i].Materials[m].Name)
                                    {
                                        this.Documents[i].Materials[m].IsModified = true;
                                        specRevised = true;
                                    }
                                    else
                                    {
                                        for (int p = 0; p < spec.Documents[i].Materials[m].Profiles.Count; p++)
                                        {
                                            if (this.Documents[i].Materials[m].Profiles.Count > p)
                                            {
                                                if (spec.Documents[i].Materials[m].Profiles[p].Name != this.Documents[i].Materials[m].Profiles[p].Name)
                                                {
                                                    this.Documents[i].Materials[m].Profiles[p].IsModified = true;
                                                    specRevised = true;
                                                }
                                                else
                                                {
                                                    for (int c = 0; c < spec.Documents[i].Materials[m].Profiles[p].ConstructionTypes[0].Constructions.Count; c++)
                                                    {
                                                        if (this.Documents[i].Materials[m].Profiles[p].ConstructionTypes[0].Constructions.Count > c)
                                                        {
                                                            if (spec.Documents[i].Materials[m].Profiles[p].ConstructionTypes[0].Constructions[c].FireResists.Sum(f => f.Count) != this.Documents[i].Materials[m].Profiles[p].ConstructionTypes[0].Constructions[c].FireResists.Sum(f => f.Count))
                                                            {
                                                                this.Documents[i].Materials[m].Profiles[p].ConstructionTypes[0].Constructions[c].IsModified = true;
                                                                specRevised = true;
                                                            }
                                                            else
                                                            {
                                                                //нужно ли?
                                                                //this.Documents[i].Materials[m].Profiles[p].ConstructionTypes[0].Constructions[c].IsModified = false;
                                                            }

                                                        }
                                                        else
                                                        {
                                                            specRevised = true;
                                                            this.Documents[i].Materials[m].Profiles[p].IsModified = true;
                                                        }
                                                    }

                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                NewSpec.Revision = spec.Revision;

                if (specRevised)
                    NewSpec.Revision++;
            }
            #endregion



            if (Documents.Count > 0)
            {
                NewSpec.Documents = this.Documents;
                NewSpec.Cipher = stampData.Cipher;
                NewSpec.BuildingObject1 = stampData.BuildingObject1;
                NewSpec.BuildingObject2 = stampData.BuildingObject2;
                NewSpec.BuildingObject3 = stampData.BuildingObject3;
                NewSpec.BuildingName1 = stampData.BuildingName1;
                NewSpec.BuildingName2 = stampData.BuildingName2;
                NewSpec.BuildingName3 = stampData.BuildingName3;
                NewSpec.AttName6 = stampData.AttName6;
                NewSpec.AttName7 = stampData.AttName7;
                NewSpec.AttName8 = stampData.AttName8;
                NewSpec.AttName9 = stampData.AttName9;
                NewSpec.AttName10 = stampData.AttName10;
                NewSpec.AttName11 = stampData.AttName11;
                NewSpec.AttValue6 = stampData.AttValue6;
                NewSpec.AttValue7 = stampData.AttValue7;
                NewSpec.AttValue8 = stampData.AttValue8;
                NewSpec.AttValue9 = stampData.AttValue9;
                NewSpec.AttValue10 = stampData.AttValue10;
                NewSpec.AttValue11 = stampData.AttValue11;
                NewSpec.Stage = stampData.Stage;
                NewSpec.Sheets = stampData.Sheets;
                NewSpec.OrganizationName1 = stampData.OrganizationName1;
                NewSpec.OrganizationName2 = stampData.OrganizationName2;
                NewSpec.OrganizationName3 = stampData.OrganizationName3;
                NewSpec.Headers = stampData.Headers = SpecTableHeaders[0].ConstructionTypes[0].Constructions.Select(i => i.Name).ToArray();

                File.WriteAllText(file, JsonConvert.SerializeObject(NewSpec, Formatting.Indented));

                SpecFilePatch = file;
            }
        }

        private ICommand _newCommand;

        public ICommand NewCommand
        {
            get
            {
                if (_newCommand == null)
                {
                    _newCommand = new RelayCommand(
                        param => this.NewSpecification(),
                        param => this.CanAdd()
                    );
                }
                return _newCommand;
            }
        }

        private void NewSpecification()
        {
            switch (MessageBox.Show("Сохранить изменения в текущем документе?", "Новая спецификация", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question))
            {
                case DialogResult.Yes:
                    if (SpecFilePatch != null && SpecFilePatch != $"{Environment.GetEnvironmentVariable("USERPROFILE")}\\backup.json")
                        SaveSpecification();
                    else
                        SaveAsSpecification();
                    break;
                case DialogResult.No:

                    break;
                case DialogResult.Cancel:
                    return;
            }
            SpecFilePatch = null;
            Documents.Clear();
            sph.ConstructionTypes[0].Constructions.Clear();
            sph.ConstructionTypes[0].Constructions.Add(new Construction(true, null) { ID = 5, Name = "конструкция" });
            CountTotalMass();
        }

        private ICommand _openCommand;

        public ICommand OpenCommand
        {
            get
            {
                if (_openCommand == null)
                {
                    _openCommand = new RelayCommand(
                        param => this.OpenSpecification(),
                        param => this.CanAdd()
                    );
                }
                return _openCommand;
            }
        }

        private void OpenSpecification()
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();

            openFileDialog1.InitialDirectory = settings.LastDirectory;
            openFileDialog1.Filter = "json files (*.json)|*.json";
            openFileDialog1.FilterIndex = 2;
            openFileDialog1.RestoreDirectory = true;

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
                if (openFileDialog1.CheckFileExists)
                {
                    switch (MessageBox.Show("Сохранить изменения в текущем документе?", "Открытие файла", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question))
                    {
                        case DialogResult.Yes:
                            if (SpecFilePatch != null && SpecFilePatch != $"{Environment.GetEnvironmentVariable("USERPROFILE")}\\backup.json")
                                SaveSpecification();
                            else
                                SaveAsSpecification();
                            break;
                        case DialogResult.No:

                            break;
                        case DialogResult.Cancel:
                            return;
                    }
                    string[] split = openFileDialog1.FileName.Split('\\');

                    settings.LastDirectory = "";

                    for (int i = 0; i < split.Length - 1; i++)
                    {
                        settings.LastDirectory += (split[i] + "\\");
                    }


                    LoadSpecFile(openFileDialog1.FileName, settings.LastDirectory);

                    SpecFilePatch = openFileDialog1.FileName;

                    if (!settings.LastFiles.Contains(openFileDialog1.FileName))
                    {
                        if (settings.LastFiles.Count == 10)
                        {
                            settings.LastFiles.RemoveAt(9);
                        }
                    }
                    else
                    {
                        settings.LastFiles.Remove(openFileDialog1.FileName);

                    }
                    settings.LastFiles.Insert(0, openFileDialog1.FileName);
                }
        }

        internal void LoadSpecFile(string fileName, string initDir)
        {
            if (sph.ConstructionTypes.Count == 0)
                sph.ConstructionTypes.Add(new ConstructionType()
                {
                    Constructions = new ObservableCollection<Construction>()
                });
            else
                sph.ConstructionTypes[0].Constructions.Clear();

            try
            {
                SpecFilePatch = fileName;
                settings.LastDirectory = initDir;
                Documents.Clear();
                // Dictionary<string, List<string>> cts = new Dictionary<string, List<string>>();
                ObservableCollection<ConstructionType> allConstrTypes = new ObservableCollection<ConstructionType>();
                ObservableCollection<Construction> allConst = new ObservableCollection<Construction>();

                spec = JsonConvert.DeserializeObject<Specification>(File.ReadAllText(SpecFilePatch));
                stampData.Cipher = spec.Cipher;
                stampData.BuildingObject1 = spec.BuildingObject1;
                stampData.BuildingObject2 = spec.BuildingObject2;
                stampData.BuildingObject3 = spec.BuildingObject3;
                stampData.BuildingName1 = spec.BuildingName1;
                stampData.BuildingName2 = spec.BuildingName2;
                stampData.BuildingName3 = spec.BuildingName3;
                stampData.AttName6 = spec.AttName6;
                stampData.AttName7 = spec.AttName7;
                stampData.AttName8 = spec.AttName8;
                stampData.AttName9 = spec.AttName9;
                stampData.AttName10 = spec.AttName10;
                stampData.AttName11 = spec.AttName11;
                stampData.AttValue6 = spec.AttValue6;
                stampData.AttValue7 = spec.AttValue7;
                stampData.AttValue8 = spec.AttValue8;
                stampData.AttValue9 = spec.AttValue9;
                stampData.AttValue10 = spec.AttValue10;
                stampData.AttValue11 = spec.AttValue11;
                stampData.Stage = spec.Stage;
                stampData.Sheets = spec.Sheets;
                stampData.OrganizationName1 = spec.OrganizationName1;
                stampData.OrganizationName2 = spec.OrganizationName2;
                stampData.OrganizationName3 = spec.OrganizationName3;
                stampData.Headers = spec.Headers;

                foreach (var item in spec.Documents)
                {
                    bool docExistInDb = true;
                    if (DocumentItemsSource.Where(d => ((d.Name == null) ? "" : d.Name.Trim()) == ((item.Name == null) ? "" : item.Name.Trim())).Count() == 0)
                    {
                        DocumentItemsSource.Add(new Sortament() { Name = (item.Name == null) ? "" : item.Name.Trim() });
                        if (!ProfilesItems.ContainsKey((item.Name == null) ? "" : item.Name.Trim()))
                            ProfilesItems.Add((item.Name == null) ? "" : item.Name.Trim(), new ObservableCollection<string>());
                        docExistInDb = false;
                        item.CellBackgroundColor = sdc.FromArgb(50, 255, 0, 0);
                        item.Tooltip = "Элемент не найден в базе стандартов";
                    }
                    else
                    {
                        item.Tooltip = (item.Name == null) ? "" : item.Name.Trim();
                        item.CellBackgroundColor = sdc.White;
                    }

                    foreach (var material in item.Materials)
                    {
                        foreach (var profile in material.Profiles)
                        {
                            profile.ProfilesItems = ProfilesItems[(item.Name == null) ? "" : item.Name.Trim()];

                            if (!docExistInDb && !ProfilesItems[(item.Name == null) ? "" : item.Name.Trim()].Contains(profile.Name))
                            {
                                ProfilesItems[(item.Name == null) ? "" : item.Name.Trim()].Add(profile.Name);
                                profile.CellBackgroundColor = sdc.FromArgb(50, 255, 0, 0);
                                profile.Tooltip = "Элемент не найден в базе профилей";
                            }
                            else
                            {
                                profile.Tooltip = profile.Name;
                                profile.CellBackgroundColor = sdc.White;
                            }
                            foreach (var constructionType in profile.ConstructionTypes)
                            {
                                string currentCtName = constructionType.Name;
                                if (allConstrTypes.Where(i => i.Name == currentCtName).ToArray().Length == 0)
                                {
                                    allConstrTypes.Add(new ConstructionType()
                                    {
                                        Name = currentCtName,
                                        ID = constructionType.ID,
                                        Constructions = new ObservableCollection<Construction>()
                                    });
                                }

                                foreach (var construction in constructionType.Constructions)
                                {
                                    //тут страховка на случай, если прийдет пустой StampData Headers из Tekla
                                    if (stampData.Headers != null && stampData.Headers.Length > 0)
                                        construction.Name = stampData.Headers[construction.ID - 5];
                                    string currentConstructionName = construction.Name;
                                    if (allConstrTypes[0].Constructions.Where(i => i.Name == currentConstructionName).ToArray().Length == 0)
                                    {
                                        allConstrTypes[0].Constructions.Add(
                                            new Construction(profile)
                                            {
                                                Name = currentConstructionName,
                                                ID = construction.ID
                                            });
                                    }
                                    foreach (var fr in construction.FireResists)
                                    {
                                        if (!FireResistItemsSource.Contains(fr.Name))
                                        {
                                            FireResistItemsSource.Add(fr.Name);
                                        }
                                        fr.Owner = construction;
                                    }
                                }
                                //
                            }
                        }
                    }
                    Documents.Add(item);

                }

                spec = JsonConvert.DeserializeObject<Specification>(File.ReadAllText(SpecFilePatch));

                foreach (var doc in Documents)
                {
                    doc.IsModified = false;
                    doc.AllProfilesItems = ProfilesItems;
                    if (doc.Name == null) doc.Name = "";
                    foreach (var mat in doc.Materials)
                    {
                        mat.IsModified = false;
                        mat.Owner = doc;
                        foreach (var prof in mat.Profiles)
                        {
                            prof.IsModified = false;
                            if (prof.ConstructionTypes == null)
                                prof.ConstructionTypes = new ObservableCollection<ConstructionType>();
                            prof.Owner = mat;
                            prof.ProfilesItems = ProfilesItems[doc.Name];
                            prof.ProtectedFromCorrosion = true;

                            for (int i = 0; i < allConstrTypes.Count; i++)
                            {
                                if (prof.ConstructionTypes.Where(c => c.Name == allConstrTypes[i].Name).ToArray().Length == 0)
                                {
                                    var ctToAdd = new ConstructionType()
                                    {
                                        ID = allConstrTypes[i].ID,
                                        Name = allConstrTypes[i].Name,
                                        Constructions = new ObservableCollection<Construction>(),
                                        IsModified = false
                                    };
                                    foreach (var ac in allConstrTypes[i].Constructions)
                                    {
                                        var cta = new Construction(true, prof)
                                        {
                                            Name = ac.Name,
                                            ID = ac.ID,
                                            IsModified = false
                                        };

                                        ctToAdd.Constructions.Add(cta);
                                    }
                                    prof.ConstructionTypes.Insert(i, ctToAdd);
                                }
                                else
                                {
                                    for (int ac = 0; ac < allConstrTypes[i].Constructions.Count; ac++)
                                    {
                                        var aconst = allConstrTypes[i].Constructions[ac];
                                        if (prof.ConstructionTypes[0].Constructions.Where(c => c.Name == aconst.Name).ToArray().Length == 0)
                                        {
                                            var cToAdd = new Construction(true, prof) { ID = aconst.ID, Name = aconst.Name, IsModified = false };

                                            prof.ConstructionTypes[0].Constructions.Insert(ac, cToAdd);
                                        }
                                    }
                                }

                            }
                            foreach (var ct in prof.ConstructionTypes)
                            {
                                //ct.Owner = prof;
                                foreach (var c in ct.Constructions)
                                {
                                    c.Owner = prof;
                                    foreach (var fr in c.FireResists)
                                    {
                                        fr.Owner = c;
                                    }
                                }
                            }
                        }
                    }
                }

                foreach (var ct in allConstrTypes)
                {
                    // var ctToAdd = new ObservableCollection<Construction>();
                    foreach (var c in ct.Constructions)
                    {
                        var constToAdd = new Construction(null) { Name = c.Name, ID = c.ID, IsModified = false };
                        sph.ConstructionTypes[0].Constructions.Add(constToAdd);
                        //ctToAdd.Add(new Construction() { Name = c.Name, ID = c.ID });
                    }
                    // sph.ConstructionTypes.Add(new ConstructionType() { Name = ct.Name, ID = ct.ID, Constructions = ctToAdd });
                }



                DocumentStringItemsSource = new ObservableCollection<string>(DocumentItemsSource.Select(i => i.Name.Trim()).ToList());

                updateProfileIDs();

                OnPropertyChanged("TotalMetalMass");
                OnPropertyChanged("TotalMassCount");
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show($"Не могу прочитать файл. Ошибка: {ex.Message}, {ex.ToString()}");
            }
        }

        private ICommand _saveAsCommand;

        public ICommand SaveAsCommand
        {
            get
            {
                if (_saveAsCommand == null)
                {
                    _saveAsCommand = new RelayCommand(
                        param => this.SaveAsSpecification(),
                        param => this.CanAdd()
                    );
                }
                return _saveAsCommand;
            }
        }

        internal void SaveAsSpecification()
        {
            SaveFileDialog saveFileDialog1 = new SaveFileDialog();

            saveFileDialog1.InitialDirectory = settings.LastDirectory;
            saveFileDialog1.Filter = "json files (*.json)|*.json";
            saveFileDialog1.FilterIndex = 2;
            saveFileDialog1.RestoreDirectory = true;

            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    //  if (saveFileDialog1.CheckFileExists)
                    //    {
                    //     if (System.Windows.Forms.MessageBox.Show(null, "Перезаписать файл?", "Такой файл уже есть", MessageBoxButtons.YesNo) == DialogResult.Yes)
                    //     {
                    SpecFilePatch = saveFileDialog1.FileName.ToString();

                    SaveSpecification(saveFileDialog1.FileName);

                    //  }
                    //}
                }
                catch (Exception ex)
                {
                    System.Windows.MessageBox.Show("Error: Could not read file from disk. Original error: " + ex.Message);
                }
            }
        }


        private ICommand _excelExportCommand;

        public ICommand ExcelExportCommand
        {
            get
            {
                //string file = $"{Environment.GetEnvironmentVariable("USERPROFILE")}\\TeklaMetalSpec.json";

                if (_excelExportCommand == null)
                {
                    _excelExportCommand = new RelayCommand(
                        param => CreateTable(true),
                        param => CanAdd()
                    );
                }
                return _excelExportCommand;
            }
        }



        private ICommand _autocadExportCommand;

        public ICommand AutocadExportCommand
        {
            get
            {
                //string file = $"{Environment.GetEnvironmentVariable("USERPROFILE")}\\TeklaMetalSpec.json";

                if (_autocadExportCommand == null)
                {
                    _autocadExportCommand = new RelayCommand(
                        param => ExportToAutocad(),
                        param => CanAdd()
                    );
                }
                return _autocadExportCommand;
            }
        }

        private ICommand _openFileCommand;

        public ICommand OpenFileCommand
        {
            get
            {
                //string file = $"{Environment.GetEnvironmentVariable("USERPROFILE")}\\TeklaMetalSpec.json";

                if (_openFileCommand == null)
                {
                    _openFileCommand = OpenFileRelayCommand;
                }
                return _openFileCommand;
            }
        }

        private ICommand _addFireResistCommand;

        public ICommand AddFireResistCommand
        {
            get
            {
                //string file = $"{Environment.GetEnvironmentVariable("USERPROFILE")}\\TeklaMetalSpec.json";

                if (_addFireResistCommand == null)
                {
                    _addFireResistCommand = AddFireResistRelayCommand;
                }
                return _addFireResistCommand;
            }
        }

        private void ExportToAutocad()
        {
            if (Documents.Count == 0)
                return;
            var sheetData = CreateTable();
            File.WriteAllText(Environment.GetEnvironmentVariable("TEMP") + @"\metalSpec_sd.json",
                JsonConvert.SerializeObject(sheetData, Formatting.Indented));
            try
            {
                string tempDir = @"D:\AutocadSupport\";
                Directory.CreateDirectory(tempDir);
                File.Copy(@"X:\Apps\Autocad\AutoCAD2016\design\MetalSpec\spec.dwg", tempDir + @"spec.dwg", true);
                File.Copy(@"X:\Apps\Autocad\AutoCAD2016\design\MetalSpec\acaddoc.lsp", tempDir + @"acaddoc.lsp", true);
                //"C:\Program Files\Autodesk\AutoCAD 2016\acad.exe" 
                Process proc = new Process();
                proc.StartInfo.Arguments = @"/language ""ru - RU"" /product ""ACAD"" /p \\vnpmosc52\Public\Apps\Autocad\AutoCAD2016\design\profile\AutoCAD2016vnp.arg " +
                    tempDir + "spec.dwg";
                proc.StartInfo.WorkingDirectory = @"C:\Program Files\Autodesk\AutoCAD 2016\";
                proc.StartInfo.FileName = @"C:\Program Files\Autodesk\AutoCAD 2016\acad.exe";
                proc.Start();
            }
            catch
            {
                MessageBox.Show(@"Чертеж с таблицей уже открыт в AutoCAD. 
                    Данные таблицы экспортированы успешно. 
                    Выполните комманду вставки спецификации в AutoCAD для вставки экспортированной таблицы.", "Экспорт в AutoCAD");
            }
        }


        private ICommand _browseExportPathCommand;

        public ICommand BrowseExportPathCommand
        {
            get
            {
                if (_browseExportPathCommand == null)
                {
                    _browseExportPathCommand = new RelayCommand(
                        param => BrowseExportPath(),
                        param => CanAdd()
                    );
                }
                return _browseExportPathCommand;
            }
        }

        private void BrowseExportPath()
        {
            FolderBrowserDialog fbd = new FolderBrowserDialog();
            DialogResult result = fbd.ShowDialog();

            settings.ExportDirectory = fbd.SelectedPath;
        }

        private ICommand _countMassCommand;

        public ICommand CountMassCommand
        {
            get
            {
                if (_countMassCommand == null)
                {
                    _countMassCommand = new RelayCommand(
                        param => CountTotalMass(),
                        param => CanAdd()
                    );
                }
                return _countMassCommand;
            }
        }

        private void CountTotalMass()
        {
            totalMetalMass.Clear();

            PaintArea = 0;

            Dictionary<string, int> fireResistArea = new Dictionary<string, int>(0);

            Dictionary<string, List<int>> byMaterial = new Dictionary<string, List<int>>(0);

            foreach (var doc in Documents)
            {
                for (int i = 0; i < doc.ProfilesTotal.Count; i++)
                {
                    if (totalMetalMass.Count - 1 < i)
                    {
                        totalMetalMass.Add((doc.ProfilesTotal[i] != null) ? (double)doc.ProfilesTotal[i] : 0);
                    }
                    else
                    {
                        totalMetalMass[i] += ((doc.ProfilesTotal[i] != null) ? (double)doc.ProfilesTotal[i] : 0);
                    }
                }
                double? paintArea = doc.Materials.Select(m => m.Profiles.Where(p => p.Name != null && p.ProtectedFromCorrosion)
                        .Sum(p => (p.TotalWeight) * ((profilesAreas.ContainsKey(doc.Name) && profilesAreas[doc.Name].ContainsKey(p.Name)) ? profilesAreas[doc.Name][p.Name] : null))).Sum();
                if (paintArea == 0)
                    PaintArea = null;
                else
                    PaintArea += paintArea;
            }
            OnPropertyChanged("TotalMetalMass");
            OnPropertyChanged("TotalMassCount");
            OnPropertyChanged("PaintArea");
        }

        private ICommand _addConstructionCommand;

        public ICommand AddConstructionCommand
        {
            get
            {
                if (_addConstructionCommand == null)
                {
                    _addConstructionCommand = new RelayCommand(
                        param => AddConstruction(),
                        param => CanAdd()
                    );
                }
                return _addConstructionCommand;
            }
        }

        private void AddConstruction()
        {
            var allconsts = sph.ConstructionTypes[0].Constructions;

            if (allconsts.Count == 14)
            {
                MessageBox.Show("Невозможно добавить конструкцию. Ограничение 14 элементов", "Добавление конструкции");
                return;
            }
            int c = 1;
            while (allconsts.Where(i => i.Name == "конструкция " + c.ToString()).Count() > 0)
            {
                c++;
            }

            string newName = "конструкция " + c.ToString();

            allconsts.Add(new Construction(true, null) { Name = newName, ID = allconsts.Count + 5 });

            foreach (var doc in Documents)
            {
                foreach (var mater in doc.Materials)
                {
                    foreach (var prof in mater.Profiles)
                    {
                        var consts = prof.ConstructionTypes[0].Constructions;
                        int id = consts.Count + 5;
                        consts.Add(new Construction(true, prof) { Name = newName, ID = id });
                    }
                }
            }
            OnPropertyChanged("TotalMetalMass");
            OnPropertyChanged("TotalMassCount");
            OnPropertyChanged("ProfilesTotal");
        }

        private SortedDictionary<string, int> CreateTable(bool showSucsessMessage = false)
        {
            if (Documents.Count == 0)
                return null;

            if (SpecFilePatch == null)
                SpecFilePatch = $"{Environment.GetEnvironmentVariable("USERPROFILE")}\\backup.json";

            SaveSpecification(SpecFilePatch);

            var resutlt = ExcelTable.CreateTable(SpecFilePatch, stampData, settings.ExportDirectory,
               settings.StampInExcel, settings.Units, settings.UnitDigits);

            if (showSucsessMessage)
            {
                //if (MessageBox.Show("Спецификация успешно экспортирована в файл \r\n" + resutlt.FirstOrDefault().Key.Replace("_SPEC_FILE_PATH_", "")+ "\r\nОткрыть файл в Excel?", "Экспорт в Excel", MessageBoxButtons.YesNo) == DialogResult.Yes)
                //{
                    try
                    {
                        Process proc = new Process();
                        proc.StartInfo.Arguments = "\"" + resutlt.FirstOrDefault().Key.Replace("_SPEC_FILE_PATH_", "") + "\"";
                        proc.StartInfo.FileName = @"excel";
                        proc.Start();
                    }
                    catch
                    {
                        MessageBox.Show(@"Чертеж с таблицей уже открыт в AutoCAD. 
                    Данные таблицы экспортированы успешно. 
                    Выполните комманду вставки спецификации в AutoCAD для вставки экспортированной таблицы.", "Экспорт в AutoCAD");
                    }
              //  }    
            };

            return resutlt;
        }

        private bool CanAdd()
        {
            return true;
        }

        public void AddNewDocument()
        {
            var newDoc = new Document(ProfilesItems)
            {
                Materials = new ObservableCollection<Material>()
            };

            newDoc.Materials.Add(GetNewMaterial(newDoc));

            Documents.Add(newDoc);

            updateProfileIDs();

            //OnPropertyChanged("Documents");
        }

        private Material GetNewMaterial(Document owner)
        {
            var newMaterial = new Material(owner) { Profiles = new ObservableCollection<Profile>() };

            newMaterial.Profiles.Add(GetNewProfile(newMaterial));

            return newMaterial;
        }

        private Profile GetNewProfile(Material owner)
        {
            var newProfile = new Profile(owner)
            {
                ConstructionTypes = new ObservableCollection<ConstructionType>()
            };

            var constructions = new ObservableCollection<Construction>();

            foreach (var c in sph.ConstructionTypes[0].Constructions)
            {
                var newC = new Construction(newProfile)
                {
                    Name = c.Name,
                    ID = c.ID,
                    FireResists = new ObservableCollection<FireResist>()
                };

                newC.FireResists.Add(new FireResist(newC)
                {
                    Name = "",
                });


                constructions.Add(newC);
            }

            var newCt = new ConstructionType();

            newCt.Constructions = constructions;

            newProfile.ConstructionTypes.Add(newCt);

            return newProfile;
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
