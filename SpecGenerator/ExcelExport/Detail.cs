using System;
using System.IO;

namespace MetalSpec
{
    public class Detail
    {
        private static string _configFilePath;
        public static string ConfigFilePath
        {
            get { return _configFilePath; }
            set
            {
                _configFilePath = "X:\\Apps\\Tekla\\applications\\JsonGenerator\\TeklaMetalSpecConfig.txt";
            }
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

        public Detail(int category, string standard, string profile, string material, double weight, int sortOrder)
        {
            Category = category;
            Standard = standard.Trim();
            StandardFull = GetStandardFull(standard);
            Profile = profile.Trim();
            Material = material.Trim();
            MaterialFull = GetMaterialFull(material);
            Weight = weight;
            SortOrder = sortOrder;
        }

        private string GetStandardFull(string standard)
        {
            using (StreamReader reader = new StreamReader(ConfigFilePath))
            {
                string line;
                while ((line = reader.ReadLine()) != null)
                {
                    if (line.StartsWith(standard.Trim()))
                    {
                        return line.Substring(line.IndexOf(';') + 1);
                    }
                }
            }

            return string.Empty;
        }

        private string GetMaterialFull(string material)
        {
            string materialTmp = string.Empty;

            if ((material == "09Г2С") || (material == "25Г2С"))
            {
                materialTmp = "Material1";
            }
            else
            {
                materialTmp = "Material2";
            }

            using (StreamReader reader = new StreamReader(ConfigFilePath))
            {
                string line;
                while ((line = reader.ReadLine()) != null)
                {
                    if (line.StartsWith(materialTmp))
                    {
                        return $"{material}\r\n{line.Substring(line.IndexOf(';') + 1)}";
                    }
                }
            }

            return string.Empty;
        }
    }
}
