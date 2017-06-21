using System.Collections.Generic;
using System.Linq;

namespace TeklaJsonGenerator
{
    public class ProfileGroup
    {
        public string Key { get; set; }
        public IEnumerable<CategoryGroup> categories { get; set; }
    }
}
