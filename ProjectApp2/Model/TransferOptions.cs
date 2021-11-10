using ProjectApp2.Model.enums;
using System;
using System.Threading;

namespace ProjectApp2.Model
{
    public class TransferOptions
    {
        public string ExcelTasksFileName { get; set; }
        public string ExcelOrderFileName { get; set; }
        public string ExcelRateSheetFileName { get; set; }
        public string MsProjectFileName { get; set; }
        public ProjectTransferType transferType { get; set; }
        public PercCompType PercCompType { get; set; }
        public DatabaseItems DatabaseItems { get; set; }
        public IProgress<ProgressReportModel> Progress { get; set; }
        public CancellationTokenSource CancellationToken { get; set; }
    }
}
