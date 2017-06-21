using System.Collections.Generic;
using System.Linq;

namespace TeklaJsonGenerator
{
    public class MaterialGroup
    {
        public string Key { get; set; }
        public IEnumerable<ProfileGroup> profiles { get; set; }
    }
}
