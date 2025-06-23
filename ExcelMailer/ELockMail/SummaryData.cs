using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelMailer.ELockMail
{
    public class SummaryData
    {
        public int TotalELocks { get; set; }
        public int AssignedELocks { get; set; }
        public int TotalOpenELock { get; set; }
        public int CurrentlyOpenELock { get; set; }
        public int ClosedELock { get; set; }
        public int CurrentlyOpenWarehouse { get; set; }
        public int TotalOpenedWarehouse { get; set; }
        public int StatusClosedCount { get; set; }
        public int StatusOngoingCount { get; set; }
        public int StatusOpeningCount { get; set; }
        public int TracologicDevices { get; set; }
        public int ImzDevices { get; set; }
    }
}
