using System.Collections.ObjectModel;
using System.Linq;
using System.Windows.Input;
using MetalSpec.DataAdapter;
using MetalSpec.Model;
using System.ComponentModel;

namespace MetalSpec.SpecGenerator
{
    public class DocumentsViewModel : Notifier, IDocumentsViewModel
    {
        public const string SELECTED_DOCUMENT_PROPERTY_NAME
            = "SelectedDocument";

        private readonly IDocumentsModel _model;
        private IDocumentViewModel _selectedDocument;
        private Status _detailsEstimateStatus
            = Status.None;
        private bool _detailsEnabled;
        private readonly ICommand _updateCommand;

        public ObservableCollection<Document>
            Documents { get { return _model.Documents; } }

        public int? SelectedValue
        {
            set
            {
                if (value == null)
                    return;
                Document document = GetDocument((int)value);
                if (SelectedDocument == null)
                {
                    SelectedDocument
                        = new DocumentViewModel(document);
                }
                else
                {
                    SelectedDocument.Update(document);
                }
                DetailsEstimateStatus =
                    SelectedDocument.EstimateStatus;
            }
        }

        public IDocumentViewModel SelectedDocument
        {
            get { return _selectedDocument; }
            set
            {
                if (value == null)
                {
                    _selectedDocument = value;
                    //DetailsEnabled = false;
                }
                else
                {
                    if (_selectedDocument == null)
                    {
                        _selectedDocument =
                            new DocumentViewModel(value);
                    }
                    _selectedDocument.Update(value);
                    DetailsEstimateStatus =
                        _selectedDocument.EstimateStatus;
                    DetailsEnabled = true;
                    NotifyPropertyChanged(
                        SELECTED_DOCUMENT_PROPERTY_NAME);
                }
            }
        }

        public Status DetailsEstimateStatus
        {
            get { return _detailsEstimateStatus; }
            set
            {
                _detailsEstimateStatus = value;
                NotifyPropertyChanged("DetailsEstimateStatus");
            }
        }

        public bool DetailsEnabled
        {
            get { return _detailsEnabled; }
            set
            {
                _detailsEnabled = value;
                NotifyPropertyChanged("DetailsEnabled");
            }
        }

        public ICommand UpdateCommand
        {
            get { return _updateCommand; }
        }

        public DocumentsViewModel(IDocumentsModel documentModel)
        {
            _model = documentModel;
            _model.DocumentUpdated +=
                model_DocumentUpdated;
            _updateCommand = new UpdateCommand(this);
        }

        public void UpdateDocument()
        {
            DetailsEstimateStatus =
                SelectedDocument.EstimateStatus;
            _model.UpdateDocument(SelectedDocument);
        }

        private void model_DocumentUpdated(object sender,
                                          DocumentEventArgs e)
        {
            GetDocument(e.Document.ID).Update(e.Document);
            if (SelectedDocument != null
                && e.Document.ID == SelectedDocument.ID)
            {
                SelectedDocument.Update(e.Document);
                DetailsEstimateStatus =
                    SelectedDocument.EstimateStatus;
            }
        }

        private Document GetDocument(int documentId)
        {
            return (from p in Documents
                    where p.ID == documentId
                    select p).FirstOrDefault();
        }
    }
}
