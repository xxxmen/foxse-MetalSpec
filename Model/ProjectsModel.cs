using MetalSpec.DataAdapter;
using MetalSpec.Model;
using System;
using System.Collections.ObjectModel;
using System.Linq;

namespace MetalSpec.SpecGenerator
{
    public class ProjectsModel : IProjectsModel
    {
        public ObservableCollection<Project> Projects { get; set; }
        public event EventHandler<ProjectEventArgs> ProjectUpdated = delegate { };

        public ProjectsModel(IDataService dataService)
        {
            Projects = new ObservableCollection<Project>();
            foreach (Project project in dataService.GetProjects())
            {
                Projects.Add(project);
            }
        }

        public void UpdateProject(IProject updatedProject)
        {
            GetProject(updatedProject.ID).Update(updatedProject);
            ProjectUpdated(this,
                new ProjectEventArgs(updatedProject));
        }

        private Project GetProject(int projectId)
        {
            return Projects.FirstOrDefault(
                project => project.ID == projectId);
        }
    }
}
