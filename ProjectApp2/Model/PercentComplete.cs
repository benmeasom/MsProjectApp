using ProjectApp2.Model.Abstract;

namespace ProjectApp2.Model
{
    public class PercentComplete
    {
        public string LabourPhase { get; set; }
        public string ClientRef { get; set; }
        public decimal HeightBand { get; set; }
        public string MeasomRef { get; set; }
        public string Area { get; set; }
        public string Floor { get; set; }
        public string DrawingNumber { get; set; }
        public string SubZone { get; set; }
        public string WallNumber { get; set; }
        public string InstalledOn { get; set; }
        public string InstalledBy { get; set; }
        public decimal TotalMeasure { get; set; }
        public string UOM { get; set; }
        public decimal TotalValue { get; set; }
        public decimal TotalManHours { get; set; }
        public decimal MeasureInstalled { get; set; }
        public decimal TotalValueInstalled { get; set; }
        public decimal ManHoursEarnt { get; set; }
        public decimal PercentMeasureIntalled { get; set; }
        public decimal PercentValueInstalled { get; set; }
        public decimal PercentManHoursEarned { get; set; }
        public string WallType { get; set; }
        public decimal ProgressWeightInTask { get; set; }
        public decimal ProgressToTask { get; set; }

    }

    public class PercentCompleteWithProgress : TaskBase
    {
        public double TotalMeasure { get; set; }
        public decimal ProgressGainForTask { get; set; }
        public double TotalManhour { get; set; }

    }
    public class PercentCompleteTotal
    {
        public string Area { get; set; }
        public string Floor { get; set; }
        public string DrawingNumber { get; set; }
        public string SubZone { get; set; }
        public string WallType { get; set; }
        public string MeasomRef { get; set; }
        public string LabourPhase { get; set; }
        public decimal TotalMeasureTotal { get; set; }
        public int PhaseOrder { get; set; }
        public int? AreaOrder { get; set; }
        public int? FloorOrder { get; set; }
        public int? SubZoneOrder { get; set; }
        public decimal TotalManHoursTotal { get; set; }
        public decimal TaskProgressTotal { get; set; }
    }

}
