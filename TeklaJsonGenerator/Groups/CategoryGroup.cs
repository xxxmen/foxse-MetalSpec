using System.Collections.Generic;
using System.Linq;

namespace TeklaJsonGenerator
{
    public class CategoryGroup
    {
        public int Key { get; set; }
        public IEnumerable<IGrouping<string, Detail>> fireResists { get; set; }
    }
}