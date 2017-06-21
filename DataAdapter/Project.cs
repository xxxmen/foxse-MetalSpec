namespace MetalSpec.DataAdapter
{
    public interface IProject
    {
        int ID { get; set; }
        string Name { get; set; }
        double Estimate { get; set; }
        double Actual { get; set; }
        void Update(IProject project);
    }

    public class Project : IProject
    {
        public int ID { get; set; }
        public string Name { get; set; }
        public double Estimate { get; set; }
        public double Actual { get; set; }

        public void Update(IProject project)
        {
            Name = project.Name;
            Estimate = project.Estimate;
            Actual = project.Actual;
        }
    }
}
