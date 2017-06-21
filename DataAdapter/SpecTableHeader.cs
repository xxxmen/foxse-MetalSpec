using MetalSpec.DataAdapter;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MetalSpec.DataAdapter
{
    public class SpecTableHeader
    {
        public ObservableCollection<ConstructionType> ConstructionTypes { get; set; }
    }
}
