namespace ProjectApp2.Model.Abstract
{
    public abstract class TaskBase
    {
        public string Area { get; set; }
        public string Floor { get; set; }
        public string DrawingNumber { get; set; }
        public string SubZone { get; set; }
        public string WallType { get; set; }
        public string MeasomRef { get; set; }
        public string LabourPhase { get; set; }
    }
}
