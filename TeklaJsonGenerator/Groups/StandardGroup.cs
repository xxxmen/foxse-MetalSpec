using System.Collections.Generic;

namespace TeklaJsonGenerator
{
    public class StandardGroup
    {
        public string Key { get; set; }
        public int lines { get; set; }
        public int linesMin { get; set; }
        public IEnumerable<MaterialGroup> materials { get; set; }
    }
}
