using MetalSpec.DataAdapter;
using System;
using System.Collections.ObjectModel;
using System.ComponentModel;

namespace MetalSpec.Model
{
    public class Notifier : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged = delegate { };

        protected void NotifyPropertyChanged(
            string propertyName)
        {
            PropertyChanged(this, new PropertyChangedEventArgs(propertyName));
        }
    }

    public interface IProjectsModel
    {
        ObservableCollection<Project> Projects { get; set; }
        event EventHandler<ProjectEventArgs> ProjectUpdated;
        void UpdateProject(IProject updatedProject);
    }

    public class ProjectEventArgs : EventArgs
    {
        public IProject Project { get; set; }
        public ProjectEventArgs(IProject project)
        {
            Project = project;
        }
    }

    public interface IDocumentsModel
    {
        ObservableCollection<Document> Documents { get; set; }
        event EventHandler<DocumentEventArgs> DocumentUpdated;
        void UpdateDocument(IDocument updatedDocument);
    }

    public class DocumentEventArgs : EventArgs
    {
        public IDocument Document { get; set; }
        public DocumentEventArgs(IDocument  document)
        {
            Document = document;
        }
    }
}
