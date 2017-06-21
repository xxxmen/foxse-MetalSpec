using MetalSpec.DataAdapter;
using MetalSpec.Model;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;

namespace MetalSpec.SpecGenerator
{
    public interface IDocumentViewModel : IDocument
    {
        Status EstimateStatus { get; set; }
    }

    public class DocumentViewModel : Notifier, IDocumentViewModel
    {
        private int _id;
        private string _gostID { get; set; }
        private string _materialID { get; set; }
        private string _profileID { get; set; }
        private ObservableCollection<double?> _profilesTotal { get; set; }
        private Dictionary<string, float> _constructionId { get; set; }

        private string _name;
        private double _estimate;
        private double _actual;
        private Status _estimateStatus = Status.None;

        public int ID
        {
            get { return _id; }
            set
            {
                _id = value;
                NotifyPropertyChanged("Id");
            }
        }

        public ObservableCollection<double?> ProfilesTotal
        {
            get { return _profilesTotal; }
            set
            {
                _profilesTotal = value;
                NotifyPropertyChanged("ProfilesTotal");
            }
        }

        public Dictionary<string, float> ConstructionId
        {
            get { return _constructionId; }
            set
            {
                _constructionId = value;
                NotifyPropertyChanged("ConstructionId");
            }
        }

        public string Name
        {
            get { return _name; }
            set
            {
                _name = value;
                NotifyPropertyChanged("Name");
            }
        }

        public double Estimate
        {
            get { return _estimate; }
            set
            {
                _estimate = value;
                NotifyPropertyChanged("Estimate");
            }
        }

        public double Actual
        {
            get { return _actual; }
            set
            {
                _actual = value;
                UpdateEstimateStatus();
                NotifyPropertyChanged("Actual");
            }
        }

        public Status EstimateStatus
        {
            get { return _estimateStatus; }
            set
            {
                _estimateStatus = value;
                NotifyPropertyChanged("EstimateStatus");
            }
        }

        public DocumentViewModel()
        { }

        public DocumentViewModel(IDocument document)
        {
            if (document == null)
                return;
            ID = document.ID;
            Update(document);
        }

        public void Update(IDocument document)
        {
            ID = document.ID;
            Name = document.Name;
        }

        private void UpdateEstimateStatus()
        {
            if (Actual == 0)
                EstimateStatus = Status.None;
            else if (Actual <= Estimate)
                EstimateStatus = Status.Good;
            else
                EstimateStatus = Status.Bad;
        }
    }

    public interface IDocumentsViewModel : INotifyPropertyChanged
    {
        IDocumentViewModel SelectedDocument { get; set; }
        void UpdateDocument();
    }
}
