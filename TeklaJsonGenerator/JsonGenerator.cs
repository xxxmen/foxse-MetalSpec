using Newtonsoft.Json.Linq;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Tekla.Structures.Catalogs;
using Tekla.Structures.Model;
using Tekla.Structures.Model.Operations;

namespace TeklaJsonGenerator
{
    class JsonGenerator
    {
        internal static void Run()
        {
            try
            {
                Operation.DisplayPrompt("Обработка…");
                PleaseWaitForm pleaseWait = new PleaseWaitForm();
                pleaseWait.Show();
                Application.DoEvents();

                Model model = new Model();
                Tekla.Structures.Model.UI.ModelObjectSelector selector = new Tekla.Structures.Model.UI.ModelObjectSelector();
                ModelObjectEnumerator selectedObjects = selector.GetSelectedObjects();
                ModelObjectEnumerator objects;
                if (selectedObjects.GetSize() > 0)
                {
                    objects = selectedObjects;
                }
                else
                {
                    objects = model.GetModelObjectSelector().GetAllObjectsWithType(new Type[] { typeof(Assembly) });
                }

                List<Detail> details = new List<Detail>();
                List<Detail> unknownDetails = new List<Detail>();

                if (objects.GetSize() > 0)
                {
                    string ru_shifr = string.Empty;
                    //model.GetProjectInfo().GetUserProperty("ru_shifr", ref ru_shifr);
                    model.GetProjectInfo().GetUserProperty("grafa_1", ref ru_shifr);
                    string ru_object_stroit_1 = string.Empty;
                    string ru_object_stroit_2 = string.Empty;
                    string ru_object_stroit_3 = string.Empty;
                    //model.GetProjectInfo().GetUserProperty("ru_objekt_stroit_1", ref ru_object_stroit_1);
                    //model.GetProjectInfo().GetUserProperty("ru_objekt_stroit_2", ref ru_object_stroit_2);
                    //model.GetProjectInfo().GetUserProperty("ru_objekt_stroit_3", ref ru_object_stroit_3);
                    model.GetProjectInfo().GetUserProperty("grafa_2", ref ru_object_stroit_1);
                    string ru_naimen_stroit_1 = string.Empty;
                    string ru_naimen_stroit_2 = string.Empty;
                    string ru_naimen_stroit_3 = string.Empty;
                    //model.GetProjectInfo().GetUserProperty("ru_naimen_stroit_1", ref ru_naimen_stroit_1);
                    //model.GetProjectInfo().GetUserProperty("ru_naimen_stroit_2", ref ru_naimen_stroit_2);
                    //model.GetProjectInfo().GetUserProperty("ru_naimen_stroit_3", ref ru_naimen_stroit_3);
                    model.GetProjectInfo().GetUserProperty("grafa_3", ref ru_naimen_stroit_1);
                    string ru_6 = string.Empty;
                    string ru_7 = string.Empty;
                    string ru_8 = string.Empty;
                    string ru_9 = string.Empty;
                    string ru_10 = string.Empty;
                    string ru_11 = string.Empty;
                    //model.GetProjectInfo().GetUserProperty("ru_6", ref ru_6);
                    //model.GetProjectInfo().GetUserProperty("ru_7", ref ru_7);
                    //model.GetProjectInfo().GetUserProperty("ru_8", ref ru_8);
                    //model.GetProjectInfo().GetUserProperty("ru_9", ref ru_9);
                    //model.GetProjectInfo().GetUserProperty("ru_10", ref ru_10);
                    //model.GetProjectInfo().GetUserProperty("ru_11", ref ru_11);
                    string ru_6_fam = string.Empty;
                    string ru_7_fam = string.Empty;
                    string ru_8_fam = string.Empty;
                    string ru_9_fam = string.Empty;
                    string ru_10_fam = string.Empty;
                    string ru_11_fam = string.Empty;
                    //model.GetProjectInfo().GetUserProperty("ru_6_fam", ref ru_6_fam);
                    //model.GetProjectInfo().GetUserProperty("ru_7_fam", ref ru_7_fam);
                    //model.GetProjectInfo().GetUserProperty("ru_8_fam", ref ru_8_fam);
                    //model.GetProjectInfo().GetUserProperty("ru_9_fam", ref ru_9_fam);
                    //model.GetProjectInfo().GetUserProperty("ru_10_fam", ref ru_10_fam);
                    //model.GetProjectInfo().GetUserProperty("ru_11_fam", ref ru_11_fam);
                    model.GetProjectInfo().GetUserProperty("razrabotal", ref ru_6_fam);
                    model.GetProjectInfo().GetUserProperty("proveril", ref ru_7_fam);
                    model.GetProjectInfo().GetUserProperty("nach_gruppy", ref ru_8_fam);
                    model.GetProjectInfo().GetUserProperty("n_control", ref ru_9_fam);
                    model.GetProjectInfo().GetUserProperty("nach_otdela", ref ru_10_fam);
                    model.GetProjectInfo().GetUserProperty("glav_eng", ref ru_11_fam);
                    string ru_stadiya = string.Empty;
                    //model.GetProjectInfo().GetUserProperty("ru_stadiya", ref ru_stadiya);
                    model.GetProjectInfo().GetUserProperty("grafa_6", ref ru_stadiya);
                    string ru_listov = string.Empty;
                    //model.GetProjectInfo().GetUserProperty("ru_listov", ref ru_listov);
                    model.GetProjectInfo().GetUserProperty("grafa_8", ref ru_listov);
                    int ru_listov_int = 0;
                    int.TryParse(ru_listov, out ru_listov_int);
                    string ru_nazvanie_org_1 = string.Empty;
                    string ru_nazvanie_org_2 = string.Empty;
                    string ru_nazvanie_org_3 = string.Empty;
                    //model.GetProjectInfo().GetUserProperty("ru_nazvanie_org_1", ref ru_nazvanie_org_1);
                    //model.GetProjectInfo().GetUserProperty("ru_nazvanie_org_2", ref ru_nazvanie_org_2);
                    //model.GetProjectInfo().GetUserProperty("ru_nazvanie_org_3", ref ru_nazvanie_org_3);
                    model.GetProjectInfo().GetUserProperty("grafa_9", ref ru_nazvanie_org_1);
                    string[] categories = new string[14];
                    model.GetProjectInfo().GetUserProperty("cm_kat_5", ref categories[0]);
                    model.GetProjectInfo().GetUserProperty("cm_kat_6", ref categories[1]);
                    model.GetProjectInfo().GetUserProperty("cm_kat_7", ref categories[2]);
                    model.GetProjectInfo().GetUserProperty("cm_kat_8", ref categories[3]);
                    model.GetProjectInfo().GetUserProperty("cm_kat_9", ref categories[4]);
                    model.GetProjectInfo().GetUserProperty("cm_kat_10", ref categories[5]);
                    model.GetProjectInfo().GetUserProperty("cm_kat_11", ref categories[6]);
                    model.GetProjectInfo().GetUserProperty("cm_kat_12", ref categories[7]);
                    model.GetProjectInfo().GetUserProperty("cm_kat_13", ref categories[8]);
                    model.GetProjectInfo().GetUserProperty("cm_kat_14", ref categories[9]);
                    model.GetProjectInfo().GetUserProperty("cm_kat_15", ref categories[10]);
                    model.GetProjectInfo().GetUserProperty("cm_kat_16", ref categories[11]);
                    model.GetProjectInfo().GetUserProperty("cm_kat_17", ref categories[12]);
                    model.GetProjectInfo().GetUserProperty("cm_kat_18", ref categories[13]);

                    bool isNonAssemblyExist = false;

                    while (objects.MoveNext())
                    {
                        ModelObject obj = objects.Current as ModelObject;
                        if (obj != null)
                        {
                            if (obj is Assembly)
                            {
                                Assembly assy = obj as Assembly;
                                if (assy.GetAssemblyType() == Assembly.AssemblyTypeEnum.STEEL_ASSEMBLY)
                                {
                                    int assyCategory = 0;
                                    bool result = assy.GetUserProperty("cm_kat", ref assyCategory);
                                    if (result)
                                    {
                                        assyCategory += 5;
                                    }

                                    List<Part> descendants = new List<Part>();
                                    TeklaUtils.GetDescendants(assy, ref descendants);

                                    foreach (Part part in descendants)
                                    {
                                        string standard = string.Empty;
                                        string profile = part.Profile.ProfileString;
                                        string material = string.Empty;
                                        part.GetReportProperty("MATERIAL", ref material);
                                        double weight = 0.0;
                                        part.GetReportProperty("WEIGHT", ref weight);
                                        weight /= 1000;
                                        double length = 0.0;
                                        part.GetReportProperty("LENGTH", ref length);
                                        int sortOrder = 0;
                                        string fireResist = string.Empty;
                                        part.GetUserProperty("FIRE_RESIST_VNP", ref fireResist);

                                        if ((part is Beam) || (part is PolyBeam))
                                        {
                                            string temp = part.Profile.ProfileString.ToUpper();

                                            if (temp.StartsWith("—") || (temp.StartsWith("PL")) || (temp.StartsWith("BL")) || (temp.StartsWith("FL")) ||
                                                (temp.StartsWith("BPL")) || (temp.StartsWith("FLT")) || (temp.StartsWith("FPL")) || (temp.StartsWith("PLT")) ||
                                                (temp.StartsWith("PLATE")) || (temp.StartsWith("TANKO")) || (temp.StartsWith("ПОЛОСА")))
                                            {
                                                details.Add(new Detail(part, assyCategory, "ГОСТ 19903-74", $"Лист {temp.Substring(temp.IndexOf("*") + 1)}", material, weight, 4, fireResist));
                                            }
                                            else if (temp.StartsWith("ПВ"))
                                            {
                                                details.Add(new Detail(part, assyCategory, "ГОСТ 8706-78", $"Лист {temp.Substring(temp.IndexOf("*") + 1)}", "C235", weight, 13, fireResist));
                                            }
                                            else if (temp.StartsWith("РИФ"))
                                            {
                                                details.Add(new Detail(part, assyCategory, "ГОСТ 8568-77", $"Лист {temp.Substring(temp.IndexOf("*") + 1)}", "C235", weight, 14, fireResist));
                                            }
                                            else if (temp.StartsWith("ЧРИФ"))
                                            {
                                                details.Add(new Detail(part, assyCategory, "ГОСТ 8568-77", $"Лист {temp.Substring(temp.IndexOf("*") + 1)}", "C235", weight, 15, fireResist));
                                            }
                                            else
                                            {
                                                LibraryProfileItem profileItem = new LibraryProfileItem();
                                                profileItem.ProfileName = part.Profile.ProfileString;
                                                profileItem.Select();
                                                ArrayList parameters = profileItem.aProfileItemUserParameters;
                                                if (parameters.Count > 0)
                                                {
                                                    foreach (ProfileItemParameter parameter in parameters)
                                                    {
                                                        if (parameter.Property == "GOST_NAME")
                                                        {
                                                            standard = parameter.StringValue;
                                                        }
                                                        if (parameter.Property == "TPL_SORT")
                                                        {
                                                            int.TryParse(parameter.StringValue, out sortOrder);
                                                        }
                                                    }

                                                    details.Add(new Detail(part, assyCategory, standard, profile, material, weight, sortOrder, fireResist));
                                                }
                                                else
                                                {
                                                    unknownDetails.Add(new Detail(part, assyCategory, "???", profile, material, weight, 999, fireResist));
                                                }
                                            }
                                        }
                                        else if (part is ContourPlate)
                                        {
                                            string temp = part.Profile.ProfileString.ToUpper();

                                            if (temp.StartsWith("—"))
                                            {
                                                details.Add(new Detail(part, assyCategory, "ГОСТ 19903-74", $"Лист {temp.Substring(1)}", material, weight, 4, fireResist));
                                            }
                                            else if ((temp.StartsWith("PL")) || (temp.StartsWith("BL")) || (temp.StartsWith("FL")))
                                            {
                                                details.Add(new Detail(part, assyCategory, "ГОСТ 19903-74", $"Лист {temp.Substring(2)}", material, weight, 4, fireResist));
                                            }
                                            else if ((temp.StartsWith("BPL")) || (temp.StartsWith("FLT")) || (temp.StartsWith("FPL")) || (temp.StartsWith("PLT")))
                                            {
                                                details.Add(new Detail(part, assyCategory, "ГОСТ 19903-74", $"Лист {temp.Substring(3)}", material, weight, 4, fireResist));
                                            }
                                            else if ((temp.StartsWith("PLATE")) || (temp.StartsWith("TANKO")))
                                            {
                                                details.Add(new Detail(part, assyCategory, "ГОСТ 19903-74", $"Лист {temp.Substring(5)}", material, weight, 4, fireResist));
                                            }
                                            else if (temp.StartsWith("ПОЛОСА"))
                                            {
                                                details.Add(new Detail(part, assyCategory, "ГОСТ 19903-74", $"Лист {temp.Substring(6)}", material, weight, 4, fireResist));
                                            }
                                            else if (temp.StartsWith("ПВ"))
                                            {
                                                details.Add(new Detail(part, assyCategory, "ГОСТ 8706-78", $"Лист {temp.Substring(2)}", "C235", weight, 13, fireResist));
                                            }
                                            else if (temp.StartsWith("РИФ"))
                                            {
                                                details.Add(new Detail(part, assyCategory, "ГОСТ 8568-77", $"Лист {temp.Substring(3)}", "C235", weight, 14, fireResist));
                                            }
                                            else if (temp.StartsWith("ЧРИФ"))
                                            {
                                                details.Add(new Detail(part, assyCategory, "ГОСТ 8568-77", $"Лист {temp.Substring(4)}", "C235", weight, 15, fireResist));
                                            }
                                            else
                                            {
                                                unknownDetails.Add(new Detail(part, assyCategory, "???", $"{part.Profile.ProfileString}", material, weight, 999, fireResist));
                                            }
                                        }
                                    }
                                }
                            }
                            else
                            {
                                if (!(obj is Grid))
                                {
                                    isNonAssemblyExist = true;
                                }
                            }
                        }
                    }

                    if (isNonAssemblyExist == true)
                    {
                        MessageBox.Show("Среди выбранных объектов есть объекты не являющиеся сборками.\r\nВыберите только сборки или не выбирайте ничего и запустите программу снова.",
                            "Сообщение", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

                        pleaseWait.Close();
                        pleaseWait.Dispose();
                        Operation.DisplayPrompt("Ошибка.");

                        return;
                    }

                    IEnumerable<StandardGroup> groups =
                        from d in details
                        orderby d.SortOrder
                        group d by d.StandardFull into g
                        select new StandardGroup()
                        {
                            Key = g.Key,
                            materials = from d in g
                                        orderby d.MaterialFull
                                        group d by d.MaterialFull into g2
                                        select new MaterialGroup()
                                        {
                                            Key = g2.Key,
                                            profiles = from d in g2
                                                       orderby d.Profile
                                                       group d by d.Profile into g3
                                                       select new ProfileGroup()
                                                       {
                                                           Key = g3.Key,
                                                           categories = from d in g3
                                                                        orderby d.Category
                                                                        group d by d.Category into g4
                                                                        select new CategoryGroup()
                                                                        {
                                                                            Key = g4.Key,
                                                                            fireResists = from d in g4
                                                                                          group d by d.FireResist into g5
                                                                                          select g5
                                                                        }
                                                       }
                                        }
                        };

                    JObject json =
                        new JObject(
                        new JProperty("ID", 0),
                        new JProperty("Name", null),
                        new JProperty("Profiles", null),
                        new JProperty("Revision", 0),
                        new JProperty("AntiCor", false),
                        new JProperty("Cipher", ru_shifr),
                        new JProperty("BuildingObject1", ru_object_stroit_1),
                        new JProperty("BuildingObject2", ru_object_stroit_2),
                        new JProperty("BuildingObject3", ru_object_stroit_3),
                        new JProperty("BuildingName1", ru_naimen_stroit_1),
                        new JProperty("BuildingName2", ru_naimen_stroit_2),
                        new JProperty("BuildingName3", ru_naimen_stroit_3),
                        new JProperty("AttName6", ru_6),
                        new JProperty("AttName7", ru_7),
                        new JProperty("AttName8", ru_8),
                        new JProperty("AttName9", ru_9),
                        new JProperty("AttName10", ru_10),
                        new JProperty("AttName11", ru_11),
                        new JProperty("AttValue6", ru_6_fam),
                        new JProperty("AttValue7", ru_7_fam),
                        new JProperty("AttValue8", ru_8_fam),
                        new JProperty("AttValue9", ru_9_fam),
                        new JProperty("AttValue10", ru_10_fam),
                        new JProperty("AttValue11", ru_11_fam),
                        new JProperty("Stage", ru_stadiya),
                        new JProperty("Sheets", ru_listov_int),
                        new JProperty("OrganzitaionName1", ru_nazvanie_org_1),
                        new JProperty("OrganzitaionName2", ru_nazvanie_org_2),
                        new JProperty("OrganzitaionName3", ru_nazvanie_org_3),
                        new JProperty("Headers", new JArray(categories)),
                        new JProperty("Documents",
                        new JArray(
                            from p in groups
                            select new JObject(
                                new JProperty("ID", 0),
                                new JProperty("Name", p.Key),
                                new JProperty("ProfilesTotal", new JArray()),
                                new JProperty("ProfilesCount", 0),
                                new JProperty("TotalWeight", 0),
                                new JProperty("IsSelected", false),
                                new JProperty("IsModified", false),
                                new JProperty("Materials",
                                new JArray(
                                    from q in p.materials
                                    select new JObject(
                                        new JProperty("ID", 0),
                                        new JProperty("Name", q.Key),
                                        new JProperty("TotalConstructionsWeight", new JArray()),
                                        new JProperty("TotalWeight", 0),
                                        new JProperty("IsSelected", false),
                                        new JProperty("IsModified", false),
                                        new JProperty("Profiles",
                                        new JArray(
                                            from r in q.profiles
                                            select new JObject(
                                                new JProperty("ID", 0),
                                                new JProperty("Name", r.Key),
                                                new JProperty("Tag", null),
                                                new JProperty("ProfileItems", new JArray()),
                                                new JProperty("Weight", 0),
                                                new JProperty("IsModified", false),
                                                new JProperty("ConstructionTypes",
                                                    new JArray(
                                                        new JObject(
                                                            new JProperty("Name", ""),
                                                            new JProperty("Constructions",
                                                            new JArray(
                                                                from s in r.categories
                                                                select new JObject(
                                                                    new JProperty("ID", s.Key),
                                                                    new JProperty("Name", ""),
                                                                    new JProperty("Count", GetConstructionWeight(s)),
                                                                    new JProperty("Profiles", null),
                                                                    new JProperty("IsModified", false),
                                                                    new JProperty("FireResists",
                                                                        new JArray(
                                                                            from t in s.fireResists
                                                                            select new JObject(
                                                                                new JProperty("Name", t.Key),
                                                                                new JProperty("Count", GetFireResistWeight(t))
                                                                                )))))))))))))))))));

                    string jsonFilePath = $"{Environment.GetEnvironmentVariable("USERPROFILE")}\\TeklaMetalSpec.json";
                    using (StreamWriter writer = new StreamWriter(jsonFilePath, false, Encoding.UTF8))
                    {
                        writer.Write(json.ToString());
                    }
                }

                pleaseWait.Close();
                pleaseWait.Dispose();
                Operation.DisplayPrompt("Завершено.");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"{ex.StackTrace}\r\n\r\n{ex.Message}");
            }
        }

        private static double GetConstructionWeight(CategoryGroup categories)
        {
            double weight = 0.0;

            foreach (IGrouping<string, Detail> fireResist in categories.fireResists)
            {
                foreach (Detail detail in fireResist)
                {
                    weight += detail.Weight;
                }
            }

            return weight;
        }

        private static double GetFireResistWeight(IGrouping<string, Detail> details)
        {
            double weight = 0.0;

            foreach (Detail detail in details)
            {
                weight += detail.Weight;
            }

            return weight;
        }
    }
}
