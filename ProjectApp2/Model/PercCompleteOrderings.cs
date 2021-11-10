using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ProjectApp2.Model
{
    public class PercCompleteOrderings
    {
        public List<PercCompleteOrder> PercentCompleteAreaOrder { get; set; } = new List<PercCompleteOrder>();
        public List<PercCompleteOrder> PercentCompleteFloorOrder { get; set; } = new List<PercCompleteOrder>();
        public List<PercCompleteOrder> PercentCompleteSubZoneOrder { get; set; } = new List<PercCompleteOrder>();
        public string ErrorOnOrdering { get; set; }
    }
}
