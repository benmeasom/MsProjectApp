using ProjectApp2.Model.enums;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ProjectApp2.Model
{
    public class PercCompleteOrder
    {
        public TaskOrder OrderingItemType { get; set; }
        public string ItemValue { get; set; }
        public string UnOrderedItemValue { get; set; }
        public int? OrderNumber { get; set; }
    }
}
