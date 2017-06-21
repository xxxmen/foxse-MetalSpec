using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;

namespace MetalSpec.DataAdapter
{
    public interface ISpecification
    {
        int ID { get; set; }
        string Name { get; set; }
        #region StampData
        string Cipher { get; set; }
        string BuildingObject1 { get; set; }
        string BuildingObject2 { get; set; }
        string BuildingObject3 { get; set; }
        string BuildingName1 { get; set; }
        string BuildingName2 { get; set; }
        string BuildingName3 { get; set; }
        string AttName6 { get; set; }
        string AttName7 { get; set; }
        string AttName8 { get; set; }
        string AttName9 { get; set; }
        string AttName10 { get; set; }
        string AttName11 { get; set; }
        string AttValue6 { get; set; }
        string AttValue7 { get; set; }
        string AttValue8 { get; set; }
        string AttValue9 { get; set; }
        string AttValue10 { get; set; }
        string AttValue11 { get; set; }
        string Stage { get; set; }
        int Sheets { get; set; }
        string OrganizationName1 { get; set; }
        string OrganizationName2 { get; set; }
        string OrganizationName3 { get; set; }
        string[] Headers { get; set; }
        #endregion
        List<int> Profiles { get; set; }
        int Revision { get; set; }
        void Update(ISpecification specification);
    }

    /// <summary>
    /// Основной документ спецификации
    /// </summary>
    public class Specification : ISpecification
    {
        public int ID { get; set; }
        public string Name { get; set; }
        #region StampData
        public string Cipher { get; set; }
        public string BuildingObject1 { get; set; }
        public string BuildingObject2 { get; set; }
        public string BuildingObject3 { get; set; }
        public string BuildingName1 { get; set; }
        public string BuildingName2 { get; set; }
        public string BuildingName3 { get; set; }
        public string AttName6 { get; set; }
        public string AttName7 { get; set; }
        public string AttName8 { get; set; }
        public string AttName9 { get; set; }
        public string AttName10 { get; set; }
        public string AttName11 { get; set; }
        public string AttValue6 { get; set; }
        public string AttValue7 { get; set; }
        public string AttValue8 { get; set; }
        public string AttValue9 { get; set; }
        public string AttValue10 { get; set; }
        public string AttValue11 { get; set; }
        public string Stage { get; set; }
        public int Sheets { get; set; }
        public string OrganizationName1 { get; set; }
        public string OrganizationName2 { get; set; }
        public string OrganizationName3 { get; set; }
        public string[] Headers { get; set; }
        #endregion

        /// <summary>
        /// Типы конструкции в спецификации
        /// </summary>
        public List<int> Profiles { get; set; }

        public ObservableCollection<Document> Documents { get; set; }

        public int Revision { get; set; }

        public void Update(ISpecification specification)
        {
            Name = specification.Name;
            Profiles = specification.Profiles;
            Revision = specification.Revision;
        }
    }
}