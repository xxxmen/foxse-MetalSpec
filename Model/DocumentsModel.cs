using MetalSpec.DataAdapter;
using MetalSpec.Model;
using System;
using System.Collections.ObjectModel;
using System.Linq;

namespace MetalSpec.SpecGenerator
{
    public class DocumentsModel : IDocumentsModel
    {
        public ObservableCollection<Document> Documents { get; set; }
        public event EventHandler<DocumentEventArgs> DocumentUpdated = delegate { };

        public DocumentsModel(IDataService dataService)
        {
            Documents = new ObservableCollection<Document>();
            foreach (Document project in dataService.GetDocuments())
            {
                Documents.Add(project);
            }
        }

        public void UpdateDocument(IDocument updatedDocument)
        {
            GetDocument(updatedDocument.ID).Update(updatedDocument);
            DocumentUpdated(this,
                new DocumentEventArgs(updatedDocument));
        }

        private Document GetDocument(int documentID)
        {
            return Documents.FirstOrDefault(
                document => document.ID == documentID);
        }
    }
}
