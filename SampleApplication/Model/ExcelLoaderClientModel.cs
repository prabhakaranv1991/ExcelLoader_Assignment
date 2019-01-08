using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SampleApplication.Model
{
    [ExcludeFromCodeCoverage]
    public class ExcelLoaderClientModel
    {
        public string CommodityCode { get; set;}
        public string DiminishingBalanceContract { get; set; }
        public double ExpiryMonthLimit { get; set; }
        public double AllMonthLimit { get; set; }
        public double AnyOneMonthLimit { get; set; }
        public DateTime ValidFrom { get; set; }
    }
}
