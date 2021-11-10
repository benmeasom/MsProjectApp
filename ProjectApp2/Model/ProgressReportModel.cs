namespace ProjectApp2.Model
{
    public class ProgressReportModel
    {
        public int PercentageComplete { get; set; } = 0;
        public string ProgressStatus { get; set; } = string.Empty;
        public bool CancelEnabled { get; set; }
        public string ImportSuccessStatus { get; set; } = string.Empty;
    }
}
