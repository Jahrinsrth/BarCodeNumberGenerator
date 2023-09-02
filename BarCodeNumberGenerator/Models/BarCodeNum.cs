using System.Collections.Generic;

namespace BarCodeNumberGenerator
{
    public class BarCodeNum
    {
        public long StartingNumber { get; set; }
        public long NumberGap { get; set;}
        public long TotalNumberLength { get; set; }
        public List<long> NumberList { get; set; }
        public string ErrorMsg { get; set; }
        public bool IsValid { get; set; }
    }
}
