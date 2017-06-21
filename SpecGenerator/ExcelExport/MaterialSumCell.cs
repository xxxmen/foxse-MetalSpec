using System.Collections.Generic;

namespace MetalSpec
{
    class MaterialSumCell
    {
        private string _material;
        public string Material
        {
            get { return _material; }
            set { _material = value; }
        }

        private string _category5Addresses;
        public string Category5Addresses
        {
            get { return _category5Addresses; }
            set { _category5Addresses = value; }
        }

        private string _category6Addresses;
        public string Category6Addresses
        {
            get { return _category6Addresses; }
            set { _category6Addresses = value; }
        }

        private string _category7Addresses;
        public string Category7Addresses
        {
            get { return _category7Addresses; }
            set { _category7Addresses = value; }
        }

        private string _category8Addresses;
        public string Category8Addresses
        {
            get { return _category8Addresses; }
            set { _category8Addresses = value; }
        }

        private string _category9Addresses;
        public string Category9Addresses
        {
            get { return _category9Addresses; }
            set { _category9Addresses = value; }
        }

        private string _category10Addresses;
        public string Category10Addresses
        {
            get { return _category10Addresses; }
            set { _category10Addresses = value; }
        }

        private string _category11Addresses;
        public string Category11Addresses
        {
            get { return _category11Addresses; }
            set { _category11Addresses = value; }
        }

        private string _category12Addresses;
        public string Category12Addresses
        {
            get { return _category12Addresses; }
            set { _category12Addresses = value; }
        }

        private string _category13Addresses;
        public string Category13Addresses
        {
            get { return _category13Addresses; }
            set { _category13Addresses = value; }
        }

        private string _category14Addresses;
        public string Category14Addresses
        {
            get { return _category14Addresses; }
            set { _category14Addresses = value; }
        }

        private string _category15Addresses;
        public string Category15Addresses
        {
            get { return _category15Addresses; }
            set { _category15Addresses = value; }
        }

        private string _category16Addresses;
        public string Category16Addresses
        {
            get { return _category16Addresses; }
            set { _category16Addresses = value; }
        }

        private string _category17Addresses;
        public string Category17Addresses
        {
            get { return _category17Addresses; }
            set { _category17Addresses = value; }
        }

        private string _category18Addresses;
        public string Category18Addresses
        {
            get { return _category18Addresses; }
            set { _category18Addresses = value; }
        }

        public MaterialSumCell(string material)
        {
            Material = material;
        }

        public MaterialSumCell(string material
            , string category5Addresses, string category6Addresses, string category7Addresses
            , string category8Addresses, string category9Addresses, string category10Addresses
            , string category11Addresses, string category12Addresses, string category13Addresses
            , string category14Addresses, string category15Addresses, string category16Addresses
            , string category17Addresses, string category18Addresses)
        {
            Material = material;
            Category5Addresses = category5Addresses;
            Category6Addresses = category6Addresses;
            Category7Addresses = category7Addresses;
            Category8Addresses = category8Addresses;
            Category9Addresses = category9Addresses;
            Category10Addresses = category10Addresses;
            Category11Addresses = category11Addresses;
            Category12Addresses = category12Addresses;
            Category13Addresses = category13Addresses;
            Category14Addresses = category14Addresses;
            Category15Addresses = category15Addresses;
            Category16Addresses = category16Addresses;
            Category17Addresses = category17Addresses;
            Category18Addresses = category18Addresses;
        }

        internal static bool ContainsMaterial(List<MaterialSumCell> materialSumCells, string curMaterial)
        {
            for (int i = 0; i < materialSumCells.Count; i++)
            {
                if (materialSumCells[i].Material == curMaterial)
                {
                    return true;
                }
            }
            return false;
        }
    }
}
