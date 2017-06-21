using System.Collections.ObjectModel;

namespace MetalSpec.DataAdapter
{
    public class ProfileItemSource
    {
        public string Name { get; set; }
        public ObservableCollection<string> ProfileItems { get; set; }
        public ObservableCollection<ProfileItemSource> Owner { get; set; }
    }
}
