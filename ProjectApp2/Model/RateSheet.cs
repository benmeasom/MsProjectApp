using System;

namespace ProjectApp2.Model
{
    public class RateSheet
    {
        //spec related
        public string ProjectName { get; set; }
        public string Specification { get; set; }
        public string ClientRef { get; set; }
        public decimal Adjustment { get; set; }
        //spec items related
        public string Item { get; set; }
        public string ProductCode { get; set; }
        public string Unit { get; set; }
        public decimal MaterialCost { get; set; }
        public string LabourPhase { get; set; }
        public decimal LabourCost { get; set; }
        public decimal OtherCost { get; set; }
        public decimal MaterialWasteCost { get; set; }
        public decimal MaterialMarkup { get; set; }
        public decimal LabourMarkup { get; set; }
        public decimal OtherMarkup { get; set; }
        //public decimal TotalWithoutAdjustment { get; set; }
        public decimal MaterialTotalWithoutAdjustment
        {
            get
            {
                return (MaterialCost + MaterialWasteCost + MaterialMarkup);
            }
        }

        public decimal LabourTotalWithoutAdjustment
        {
            get
            {
                return (LabourCost + LabourMarkup);
            }
        }
        public decimal OtherTotalWithoutAdjustment
        {
            get
            {
                return (OtherCost + OtherMarkup);
            }
        }

        public decimal TotalWithoutAdjustment
        {
            get
            {
                return (MaterialTotalWithoutAdjustment + LabourTotalWithoutAdjustment + OtherTotalWithoutAdjustment);
            }
        }

        public decimal MaterialAdjustment { get; set; }
        public decimal LabourAdjustment { get; set; }
        public decimal OtherAdjustment { get; set; }

        public decimal MaterialTotal
        {
            get
            {
                return Math.Round((MaterialTotalWithoutAdjustment + MaterialAdjustment), 2);
            }
        }

        public decimal LabourTotal
        {
            get
            {
                return Math.Round((LabourTotalWithoutAdjustment + LabourAdjustment), 2);
            }
        }
        public decimal OtherTotal
        {
            get
            {
                return Math.Round((OtherTotalWithoutAdjustment + OtherAdjustment), 2);
            }
        }

        //public decimal TotalCost
        //{
        //    get
        //    {
        //        return Math.Round(TotalWithoutAdjustment + Math.Round(Adjustment, 2), 2);
        //    }
        //}
    }
}
