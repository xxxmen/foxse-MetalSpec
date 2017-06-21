using System;
using System.IO;
using System.Text.RegularExpressions;
using Tekla.Structures.Model;

namespace TeklaJsonGenerator
{
    public class Detail
    {
        private static string configFilePath = TeklaUtils.GetConfigFilePath();

        private Part _part;
        public Part Part
        {
            get { return _part; }
            set { _part = value; }
        }

        private int _category;
        public int Category
        {
            get { return _category; }
            set { _category = value; }
        }

        private string _standard;
        public string Standard
        {
            get { return _standard; }
            set { _standard = value; }
        }

        private string _standardFull;
        public string StandardFull
        {
            get { return _standardFull; }
            set { _standardFull = value; }
        }

        private string _profile;
        public string Profile
        {
            get { return _profile; }
            set { _profile = value; }
        }

        private string _material;
        public string Material
        {
            get { return _material; }
            set { _material = value; }
        }

        private string _materialFull;
        public string MaterialFull
        {
            get { return _materialFull; }
            set { _materialFull = value; }
        }

        private double _weight;
        public double Weight
        {
            get { return _weight; }
            set { _weight = value; }
        }

        private int _sortOrder;
        public int SortOrder
        {
            get { return _sortOrder; }
            set { _sortOrder = value; }
        }

        private string _fireResist;
        public string FireResist
        {
            get { return _fireResist; }
            set { _fireResist = value; }
        }

        public Detail(Part part, int category, string standard, string profile, string material, double weight, int sortOrder, string fireResist)
        {
            Part = part;
            Category = category;
            Standard = standard.Trim();
            StandardFull = GetStandardFull(standard, profile);
            Profile = GetProfile(profile).Trim();
            Material = material.Trim();
            MaterialFull = GetMaterialFull(material);
            Weight = weight;
            SortOrder = sortOrder;
            FireResist = fireResist;
        }

        private string GetProfile(string profile)
        {
            using (StreamReader reader = new StreamReader(configFilePath))
            {
                string line;
                while ((line = reader.ReadLine()) != null)
                {
                    if (line.StartsWith(profile.Trim()))
                    {
                        return line.Substring(line.IndexOf(';') + 1);
                    }
                }
            }

            return string.Empty;
        }

        private string GetStandardFull(string standard, string profile)
        {
            string standardFull = string.Empty;

            switch (standard)
            {
                case "СТО АСЧМ 20-93":
                    if (Regex.Match(profile, @"ДВУТAВР\d{1,3}К(1|2|3)").Success)
                    {
                        standardFull = "Двутавp колонный (К) по СТО АСЧМ 20-93";
                    }
                    else if (Regex.Match(profile, @"ДВУТAВР\d{1,3}Ш(1|2|3)").Success)
                    {
                        standardFull = "Двутавp широкополочный по СТО АСЧМ 20-93";
                    }
                    else if (Regex.Match(profile, @"ДВУТAВР\d{1,3}Б(1|2|3)").Success)
                    {
                        standardFull = "Двутавp нормальный (Б) по СТО АСЧМ 20-93";
                    }
                    break;
                case "ГОСТ 8239-89":
                    standardFull = "Двутавp с уклоном полок по ГОСТ 8239-89";
                    break;
                case "ГОСТ 26020-83":
                    if (Regex.Match(profile, @"Двутавр\d{1,3}Б(1|2|3)").Success)
                    {
                        standardFull = "Двутавp нормальный (Б) по ГОСТ 26020-83";
                    }
                    else if (Regex.Match(profile, @"Двутавр\d{1,3}Ш(1|2|3)").Success)
                    {
                        standardFull = "Двутавp широкополочный по ГОСТ 26020-83";
                    }
                    else if (Regex.Match(profile, @"Двутавр\d{1,3}К(1|2|3)").Success)
                    {
                        standardFull = "Двутавp колонный (К) по ГОСТ 26020-83";
                    }
                    else if (Regex.Match(profile, @"Двутавр\d{1,3}Д(Б|Ш)(1|2|3)\*?").Success)
                    {
                        standardFull = "Двутавp дополнительной серии (Д) по ГОСТ 26020-83";
                    }
                    break;
                case "ГОСТ 19425-74":
                    standardFull = "Двутавp стальной специальный по ГОСТ 19425-74*";
                    break;
                case "ГОСТ 8240-97":
                    if (Regex.Match(profile, @"Швеллер\d{1,3}(\.\d)?а?У").Success)
                    {
                        standardFull = "Швеллеp с уклоном полок по ГОСТ 8240-97";
                    }
                    else if (Regex.Match(profile, @"Швеллер\d{1,3}(\.\d)?а?П").Success)
                    {
                        standardFull = "Швеллеp с паpаллельными гpанями полок по ГОСТ 8240-97";
                    }
                    else if (Regex.Match(profile, @"Швеллер\d{1,3}(\.\d)?Э").Success)
                    {
                        standardFull = "Швеллеpы экономичные с паpаллельными гpанями полок по ГОСТ 8240-97";
                    }
                    else if (Regex.Match(profile, @"Швеллер\d{1,3}Л").Success)
                    {
                        standardFull = "Швеллеpы легкой серии с параллельными гранями полок по ГОСТ 8240-97";
                    }
                    else if (Regex.Match(profile, @"Швеллер\d{1,3}С(а|б)?").Success)
                    {
                        standardFull = "Швеллеpы специальные  по ГОСТ 8240-97";
                    }
                    break;
                case "ГОСТ 8278-83":
                    if (Material == "С245")
                    {
                        standardFull = "Гнутый равнополочный швеллер по ГОСТ 8278-83 из сталей С239-С245";
                    }
                    else if ((Material == "С255") || (Material == "С275"))
                    {
                        standardFull = "Гнутый равнополочный швеллер по ГОСТ 8278-83 из сталей С255-С275";
                    }
                    else
                    {
                        standardFull = "Гнутый равнополочный швеллер по ГОСТ 8278-83"; //Отсутствует в БД
                    }
                    break;
                case "ГОСТ 8509-93":
                    standardFull = "Уголок равнополочный по ГОСТ 8509-93";
                    break;
                case "ГОСТ 8510-86":
                    standardFull = "Уголок неравнополочный по ГОСТ 8510-86*";
                    break;
                case "ГОСТ 8510-93":
                    standardFull = "Уголок неравнополочный по ГОСТ 8510-86*";
                    break;
                case "ГОСТ 19771-93":
                    standardFull = "Уголки стальные гнутые равнополочные ГОСТ 19771-93"; //Отсутствует в БД, отсутствует в Tekla 20.0
                    break;
                case "ГОСТ 19772-93":
                    standardFull = "Уголки стальные гнутые неравнополочные ГОСТ 19772-93"; //Отсутствует в БД
                    break;
                case "ГОСТ Р 54157-2010":
                    standardFull = "Прямоугольные трубы по ГОСТ Р 54157-2010"; //Отсутствует в Tekla 20.0 (для последующих версий необходимо дописать алгоритм, т. к. в БД больше одной записи)
                    break;
                case "ГОСТ 54157-2010":
                    standardFull = "Прямоугольные трубы по ГОСТ Р 54157-2010"; //Дубликат предыдущего (проверить в последующих версиях)
                    break;
                case "ГОСТ 8732-78":
                    standardFull = "Тpубы стальные бесшовные горячедеформированные, ГОСТ 8732-78";
                    break;
                case "ГОСТ 10704-91":
                    standardFull = "Тpубы электросварные прямошовные по ГОСТ 10704-91";
                    break;
                case "ГОСТ 30245-2003":
                    if (Regex.Match(profile, @"Профиль\(кв\.\)\d{1,3}X\d{1,3}X\d{1,2}(\.\d)?").Success)
                    {
                        standardFull = "Стальные гнутые замкнутые сварные квадратные профили по ГОСТ 30245-2003";
                    }
                    else if (Regex.Match(profile, @"Профиль\(пр\.\)\d{1,3}X\d{1,3}X\d{1,2}(\.\d)?").Success)
                    {
                        standardFull = "Стальные гнутые замкнутые сварные прямоугольные профили по ГОСТ 30245-2003";
                    }
                    break;
                case "ТУ 67-2287-80":
                    standardFull = "Прямоугольные трубы по ТУ 67-2287-80";
                    break;
                case "ТУ 36-2287-80":
                    standardFull = "Квадратные трубы по ТУ 36-2287-80";
                    break;
                case "ТУ 14-2-685-86":
                    if (Regex.Match(profile, @"Тавр\d{1,2}(\.\d)?БТ(1|2|3|4)").Success)
                    {
                        standardFull = "Тавр БТ по ТУ 14-2-685-86"; //Отсутствует в БД
                    }
                    else if (Regex.Match(profile, @"Тавр\d{1,2}(\.\d)?КТ(1|2|3|4)").Success)
                    {
                        standardFull = "Тавpы колонные (КТ) по ТУ 14-2-685-86";
                    }
                    else if (Regex.Match(profile, @"Тавр\d{1,2}(\.\d)?ШТ(1|2|3|4)").Success)
                    {
                        standardFull = "Тавp ШТ по ТУ 14-2-685-86";
                    }
                    break;
                case "ГОСТ 2590-88":
                    standardFull = "Прокат стальной горячекатаный круглый ГОСТ 2590-88";
                    break;
                case "ГОСТ 2591-88":
                    standardFull = "Прокат стальной горячекатаный квадратный ГОСТ 2591-88";
                    break;
                case "ГОСТ 2590-2006":
                    standardFull = "Прокат стальной горячекатаный круглый ГОСТ 2590-88";
                    break;
                case "ГОСТ 2591-2006":
                    standardFull = "Прокат стальной горячекатаный квадратный ГОСТ 2591-88";
                    break;
                case "ГОСТ 51685-2000":
                    standardFull = "Рельсы крановые ГОСТ 51685-2000"; //Отсутствует в БД
                    break;
                case "ГОСТ Р 53866-2010":
                    standardFull = "Рельсы крановые ГОСТ Р 53866-2010"; //Отсутствует в БД, отсутствует в Tekla 20.0
                    break;
                case "ГОСТ 4121-96":
                    standardFull = "Рельсы крановые ГОСТ 4121-96"; //Отсутствует в БД, отсутствует в Tekla 20.0
                    break;
                case "ГОСТ 19903-74":
                    standardFull = "Сталь листовая горячекатанная ГОСТ 19903-74"; //Отсутствует в БД, отсутствует в Tekla 20.0
                    break;
                case "ГОСТ 8706-78":
                    standardFull = "Листы стальные просечно-вытяжные"; //Отсутствует в БД, отсутствует в Tekla 20.0
                    break;
                case "ГОСТ 8568-77":
                    standardFull = "Листы стальные с ромбическим и чечевичным рифлением"; //Отсутствует в БД, отсутствует в Tekla 20.0
                    break;
                case "ГОCТ 24045-2010":
                    standardFull = "Профили стальные листовые гнутые с трапециевидными гофрами для строительства ГОCТ 24045-2010"; //Отсутствует в БД, отсутствует в Tekla 20.0
                    break;
                default:
                    break;
            }

            return standardFull;
        }

        private string GetMaterialFull(string material)
        {
            string materialTmp = string.Empty;

            if ((material == "09Г2С") || (material == "25Г2С"))
            {
                materialTmp = "ГОСТ 27772-88";
            }
            else
            {
                materialTmp = "ГОСТ 19281-89";
            }

            return $"{material}\r\n{materialTmp}";
        }
    }
}
