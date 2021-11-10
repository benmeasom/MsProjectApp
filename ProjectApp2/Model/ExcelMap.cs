using ProjectApp2.Contracts;
using ProjectApp2.Model.Abstract;
using System.Collections.Generic;

namespace ProjectApp2.Model
{
    public class ExcelMap
    {
        public ExcelMap()
        {
            ExcelMapNewProjectMaterialAssignments = new List<ExcelMapNewProjectMaterialAssignment>();
            ExcelMapCostAssignments = new List<ExcelMapCostAssignment>();
            ExcelMapResources = new List<ExcelMapResource>();
            ExcelMapTasks = new List<ExcelMapTask>();
        }

        public List<ExcelMapNewProjectMaterialAssignment> ExcelMapNewProjectMaterialAssignments { get; set; }
        public List<ExcelMapCostAssignment> ExcelMapCostAssignments { get; set; }
        public List<ExcelMapResource> ExcelMapResources { get; set; }
        public List<ExcelMapTask> ExcelMapTasks { get; set; }
        public class ExcelMapCostAssignment
        {
            public string ResourceUniqueID { get; set; }
            public string TaskUniqueID { get; set; }
            public double Cost { get; set; }
        }

        //Task/Resource Uniqe Id may change while merging an excel file into an existing Ms Project, Id won't be used in that case
        public class ExcelMapNewProjectMaterialAssignment
        {
            public string ResourceUniqueID { get; set; }
            public string TaskUniqueID { get; set; }
            public string ScheduledWork { get; set; }
        }

        public class ExcelMapTask : TaskBase, IHasUniqeId
        {
            public string Id { get; set; }
            public string UniqeId { get; set; }
            public string Name { get; set; }
            public string Predecessors { get; set; }
            public string OutlineLevel { get; set; }
            public decimal PercentComplete { get; set; }
            public string TaskMode { get; set; }
            public string Type { get; set; }
            public string Milestone { get; set; }
            public double TotalManHours { get; set; }
            public double TotalMeasure { get; set; }
        }

        public class ExcelMapResource : IHasUniqeId
        {
            public string Id { get; set; }
            public string UniqeId { get; set; }
            public string Name { get; set; }
            public string Type { get; set; }
            public string MaterialLabel { get; set; }
            public string GroupName { get; set; }
            public string StandardRate { get; set; }
        }
    }
}
