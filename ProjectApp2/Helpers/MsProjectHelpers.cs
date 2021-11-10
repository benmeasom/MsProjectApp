using ProjectApp2.Contracts;
using ProjectApp2.Model;
using ProjectApp2.Model.enums;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using static ProjectApp2.Model.ExcelMap;
using Microsoft.Office.Interop.MSProject;
using System.Data;
using System.IO;
using ProjectApp2.Model.Abstract;
using System.Diagnostics;

namespace ProjectApp2.Helpers
{
    public class MsProjectHelpers
    {
        public void DoMainOperationsCreateOrUpdateProject(List<PercentCompleteTotal> percentCompleteTotal, List<RateSheet> rateSheets, TransferOptions transferOptions)

        {
            //Instantiate ExcelMap to Create Lists To Prepare Sheets of Excel Map file          
            ExcelMap excelMap = new ExcelMap();

            if (transferOptions.transferType == ProjectTransferType.NewProject)
            {
                string projectName = rateSheets.FirstOrDefault(x => x.ProjectName != string.Empty).ProjectName;
                CreateStaticResources(excelMap.ExcelMapResources);
                CreateMainFirstTask(excelMap, projectName);
                CreateExcelMaps(percentCompleteTotal, rateSheets, excelMap);
            }
            ExcelHelpers.MainOperationInitialize(excelMap, transferOptions, percentCompleteTotal);
        }

        private void CreateExcelMaps(List<PercentCompleteTotal> percentCompleteTotal, List<RateSheet> rateSheets, ExcelMap excelMap)
        {
            //for recording previous task, will be compared to new task before adding
            PercentCompleteTotal taskToCompare = new PercentCompleteTotal();

            //Below is to define which indentation new task will be according to previous task.
            string changedPropertyValue;

            ExcelMapTask lasAddedtTask = null;

            //Create MS Project Resource List to check resource before adding to Ms Project resources     
            List<MsProjectResource> msProjectResources = new List<MsProjectResource>();

            //Loop percentCompleteTotal list => list of Tasks to be in Ms Project - including summary tasks -
            //this loop will create ExcelMapTasks (and also its resources and assignments inside the loop)
            foreach (PercentCompleteTotal item in percentCompleteTotal)
            {
                //Return Changed Property Name comparing with the last task added to tasks list
                changedPropertyValue = ChangedPropertyName(item, taskToCompare);

                //Get the number of changed property
                int changedPropertyLevel = (int)Enum.Parse(typeof(TaskOrder), changedPropertyValue);

                //Create summary tasks and return final level task
                var newTask = AddTasks(excelMap, item, changedPropertyLevel, lasAddedtTask);
                lasAddedtTask = newTask;

                // get the material lines from ratesheet with the related spec && labour phase
                var resourcesForTask = rateSheets.Where(x => x.LabourPhase == newTask.Name.ToString() &&
                                                         x.Specification == item.MeasomRef).ToList();

                //Create resources (if not already created) and make resource assignments to task
                CreateResourcesAndAssignToTask(excelMap, item, newTask, msProjectResources, resourcesForTask);
                taskToCompare = item;
            }
        }

        private void CreateMainFirstTask(ExcelMap excelMap, string projectName)
        {
            var newTaskId = GetUniqeId(excelMap.ExcelMapTasks);
            if (projectName == string.Empty)
            {
                projectName = "#Unknown Project#";
            }
            excelMap.ExcelMapTasks.Add(new ExcelMapTask
            {
                Name = projectName,
                OutlineLevel = "1",
                Id = newTaskId,
                UniqeId = newTaskId,
                TaskMode = "Auto Scheduled",
                Type = "Fixed Work",
                Milestone = "No"
            });
        }

        /// <summary>
        /// This Method Will Create ExcelMapResources, ExcelMapCostAssignments, ExcelMapNewProjectMaterialAssignments
        /// </summary>
        /// <param name="excelMap"></param>
        /// <param name="item"></param>
        /// <param name="newTask"></param>
        /// <param name="msProjectResources"></param>
        /// <param name="resourcesForTask"></param>
        private static void CreateResourcesAndAssignToTask(ExcelMap excelMap, PercentCompleteTotal item,
                                    ExcelMapTask newTask, List<MsProjectResource> msProjectResources, List<RateSheet> resourcesForTask)
        {
            decimal difference = 0;
            int lastOccurance = 0;
            bool createResource = false;
            ExcelMapResource taskResource = null;

            CreateLabourAndOtherResourceAssignments(excelMap, resourcesForTask, newTask, item);

            //Process each material as resource 
            foreach (RateSheet rsrc in resourcesForTask)
            {
                string materialName = rsrc.ProductCode;
                decimal materialValue = rsrc.MaterialTotal;
                //check if the resource with given material name and value already exists 
                MsProjectResource existingResourceAndValue = msProjectResources.FirstOrDefault
                                            (p => p.RateSheetMaterialName == materialName && p.RateSheetMaterialValue == materialValue);
                if (existingResourceAndValue != null)
                {
                    taskResource = excelMap.ExcelMapResources.Find(x => x.Id == existingResourceAndValue.MsProjectResourceUniqeId.ToString());
                    goto TaskAssignment;
                }
                //check if the resource with given material name already exists 
                List<MsProjectResource> existingResources = msProjectResources.Where(p => p.RateSheetMaterialName == materialName).ToList();

                if (existingResources.Any())
                {
                    foreach (MsProjectResource resource in existingResources)
                    {
                        difference = Math.Abs(resource.RateSheetMaterialValue - materialValue);

                        if (difference <= Convert.ToDecimal(0.05))
                        {
                            taskResource = excelMap.ExcelMapResources.Find(x => x.Id == resource.MsProjectResourceUniqeId.ToString());
                            goto TaskAssignment;
                        }
                    }
                    MsProjectResource existingResource = msProjectResources.OrderByDescending(p => p.LastOccurance).FirstOrDefault
                                                    (p => p.RateSheetMaterialName == materialName);
                    lastOccurance = existingResource.LastOccurance;
                    createResource = true;
                }

                string unit = rsrc.Unit;
                string rsrcItem = rsrc.Item;
                string newResourceId;

                //If Materialname is not found lastoccurance is null, so materialname will be created without "_"
                //If Materialname is found materialname will be materialname + "_" + lastoccurance +1
                string msProjectResourceName = materialName + (createResource ? "_" + (lastOccurance + 1).ToString() : string.Empty);
                newResourceId = GetUniqeId(excelMap.ExcelMapResources);

                taskResource = new ExcelMapResource
                {
                    Name = msProjectResourceName,
                    Type = "Material",
                    MaterialLabel = unit,
                    GroupName = rsrcItem,
                    StandardRate = materialValue.ToString(),
                    Id = newResourceId,
                    UniqeId = newResourceId
                };

                excelMap.ExcelMapResources.Add(taskResource);

                MsProjectResource newResource = new MsProjectResource
                {
                    LastOccurance = lastOccurance + 1,
                    MsProjectResourceUniqeId = int.Parse(taskResource.UniqeId),
                    RateSheetMaterialName = materialName,
                    RateSheetMaterialValue = materialValue,
                    MsProjectResourceName = msProjectResourceName
                };
                msProjectResources.Add(newResource);

            TaskAssignment:
                var excelMapAssignment = new ExcelMapNewProjectMaterialAssignment
                {
                    //ResourceID = taskResource.Id,
                    ResourceUniqueID = taskResource.UniqeId,
                    //TaskId = newTask.Id,
                    TaskUniqueID = newTask.UniqeId,
                    ScheduledWork = item.TotalMeasureTotal.ToString()
                };

                excelMap.ExcelMapNewProjectMaterialAssignments.Add(excelMapAssignment);

                lastOccurance = 0;
                createResource = false;
            }
        }
        public static void CreateLabourAndOtherResourceAssignments(ExcelMap excelMap, List<RateSheet> resourcesForTask,
                                                                    ExcelMapTask newTask, PercentCompleteTotal item)
        {
            var labourTotal = resourcesForTask.Sum(i => i.LabourTotal);

            if (labourTotal > 0)
            {
                excelMap.ExcelMapCostAssignments.Add(new ExcelMapCostAssignment
                {
                    Cost = ((double)labourTotal * newTask.TotalMeasure),
                    ResourceUniqueID = ((int)StaticResources.Labour_Cost).ToString(),
                    TaskUniqueID = newTask.UniqeId
                });
            }

            if (item.TotalManHoursTotal != 0)
            {
                excelMap.ExcelMapNewProjectMaterialAssignments.Add(new ExcelMapNewProjectMaterialAssignment
                {
                    ResourceUniqueID = ((int)StaticResources.Labour_Manhour).ToString(),
                    TaskUniqueID = newTask.UniqeId,
                    ScheduledWork = item.TotalManHoursTotal.ToString()
                });
            }

            var otherTotal = resourcesForTask.Sum(i => i.OtherTotal);

            if (otherTotal > 0)
            {
                excelMap.ExcelMapCostAssignments.Add(new ExcelMapCostAssignment
                {
                    Cost = ((double)otherTotal * newTask.TotalMeasure),
                    ResourceUniqueID = ((int)StaticResources.Other_Cost).ToString(),
                    TaskUniqueID = newTask.UniqeId,
                });
            }
        }

        private static ExcelMapTask AddTasks(ExcelMap excelMap, PercentCompleteTotal item,
                                                                int changedPropertyLevel, ExcelMapTask lasAddedtTask)
        {
            bool noProdeceessor = changedPropertyLevel > 1;

            ExcelMapTask returnValue = null;

            //Loop tasks including summary tasks until changed property of previous task
            for (int i = changedPropertyLevel; i >= 1; i--)
            {
                returnValue = CreateTaskWithValues(excelMap.ExcelMapTasks, item, lasAddedtTask, i, noProdeceessor);

                //Add Tak(s) (including summary tasks) to Main Tasks List
                excelMap.ExcelMapTasks.Add(returnValue);
            }

            //above for loop will always count down to 1 and return last outline level task            
            return returnValue;
        }

        private static ExcelMapTask CreateTaskWithValues(List<ExcelMapTask> excelMapTasks, PercentCompleteTotal item,
            ExcelMapTask lasAddedtTask, int i, bool noProdeceessor)
        {
            string newTaskId;
            //Get the value of the property from current PercentCompleteTotal LINE
            var enumName = Enum.GetName(typeof(TaskOrder), i);
            var propertyValue2 = (item.GetType().GetProperty(enumName).
                                GetValue(item) ?? string.Empty).ToString();
            newTaskId = GetUniqeId(excelMapTasks);
            //Don't put prodecesor in the first outline lvl5 task            
            string predecessor = !noProdeceessor ? GetPredecessor(lasAddedtTask, i) : string.Empty;
            // **** OutlineCode level will be always +1 because top task outline level 1 is being created in the beginning 
            var outlineLevel = (Enum.GetNames(typeof(TaskOrder)).Length + 1 - i + 1).ToString();

            var returnValue = new ExcelMapTask
            {
                Name = propertyValue2,
                OutlineLevel = outlineLevel,
                Id = newTaskId,
                UniqeId = newTaskId,
                Predecessors = predecessor,
                PercentComplete = item.TaskProgressTotal,
                TaskMode = "Auto Scheduled",
                Type = "Fixed Work",
                Milestone = "No",
                TotalMeasure = (double)item.TotalMeasureTotal,
                TotalManHours = (double)item.TotalManHoursTotal
            };
            return returnValue;
        }

        private static string GetPredecessor(ExcelMapTask lasAddedtTask, int i)
        {
            var predecessor = string.Empty;
            if (i == 1 && lasAddedtTask != null)
            {
                predecessor = lasAddedtTask.Id.ToString();
            }

            return predecessor;
        }
        private static string GetUniqeId<T>(List<T> ts) where T : IHasUniqeId
        {
            string newId;
            if (ts.Any())
            {
                newId = (ts.Max(g => int.Parse(g.UniqeId)) + 1).ToString();
            }
            else
                newId = 1.ToString();
            return newId;
        }

        private static string ChangedPropertyName(PercentCompleteTotal item, PercentCompleteTotal taskToCompare)
        {
            PropertyInfo[] properties = typeof(PercentCompleteTotal).GetProperties();
            string taskToComparePropertyValue;
            string changedPropertyValue = string.Empty;

            foreach (PropertyInfo prop in properties)
            {
                //value of current PercentCompleteTotal property
                string propertyValue = (prop.GetValue(item) ?? string.Empty).ToString();
                //value of previous PercentCompleteTotal property
                taskToComparePropertyValue = (taskToCompare.GetType().GetProperty(prop.Name).
                                                        GetValue(taskToCompare) ?? string.Empty).ToString();
                //Compare property value of this and previous PercentCompleteTotal LINE
                if (propertyValue != taskToComparePropertyValue && prop.Name != "TotalMeasureTotal"
                        && prop.Name != "PhaseOrder" && prop.Name != "TotalManHoursTotal" && prop.Name != "TaskProgressTotal")
                {
                    //If Property Values are different get the name of the value changed property 
                    changedPropertyValue = prop.Name;
                    break;
                }
            }

            return changedPropertyValue;
        }
        private static void CreateStaticResources(List<ExcelMapResource> excelMapResources)
        {
            excelMapResources.AddRange(new List<ExcelMapResource>
            {
                new ExcelMapResource
                {
                    Id = ((int)StaticResources.Labour_Cost).ToString(),
                    UniqeId = ((int)StaticResources.Labour_Cost).ToString(),
                    Name = "Labour_Cost",
                    Type = "Cost"
                },
                new ExcelMapResource
                {
                    Id = ((int)StaticResources.Other_Cost).ToString(),
                    UniqeId = ((int)StaticResources.Other_Cost).ToString(),
                    Name = "Other_Cost",
                    Type = "Cost"
                },
                 new ExcelMapResource
                {
                    Id = ((int)StaticResources.Labour_Manhour).ToString(),
                    UniqeId = ((int)StaticResources.Labour_Manhour).ToString(),
                    Name = "Labour_Manhour",
                    Type = "Work"
                },
            });
        }

        public static void MainOperationMsProjectOperations(ExcelMap excelMap, TransferOptions transferOptions, List<PercentCompleteTotal> percentCompleteTotal)
        {
            Application projectApplication = new Application();
            var successStatus = string.Empty;

            CreateMapsInMsProjectApplication(projectApplication);

            projectApplication.DisplayAlerts = false;

            if (transferOptions.transferType == ProjectTransferType.NewProject)
            {
                CreateNewProjectFromMaps(projectApplication);
                projectApplication.TableEditEx(Name: "Entry", TaskTable: true, NewFieldName: "Cost", Width: 15);
                projectApplication.TableEditEx(Name: "Entry", TaskTable: true, NewFieldName: "Work", Width: 15);
                projectApplication.TableApply("Entry");
            }
            else if (transferOptions.transferType == ProjectTransferType.UpdateProject)
            {
                var newMapTasksWithProgres = UpdateProjectProgress(transferOptions, projectApplication, percentCompleteTotal);

                ExcelHelpers.DoProjectUpdateExcelOperations(newMapTasksWithProgres);

                projectApplication.FileOpenEx(Name: Globals.ExcelMapFilePaths.TasksPath, Merge: PjMergeType.pjMerge,
                                       Map: "Map_Tasks_Progress");
            }

            projectApplication.Visible = true;

            if (File.Exists(Globals.ExcelMapFilePaths.ErrorTextFile))
            {
                successStatus = "Error in Transfer, see the text file for errors !!";
                Process.Start(Globals.ExcelMapFilePaths.ErrorTextFile);
            }
            else
            {
                successStatus = "Transfer Succesfull!!";
            }

            transferOptions.Progress.Report(new ProgressReportModel
            {
                ImportSuccessStatus = successStatus
            });
        }

        private static void CreateMapsInMsProjectApplication(Application projectApplication)
        {
            //Create Tasks Map
            projectApplication.MapEdit(Name: "Map_Tasks", Create: true, OverwriteExisting: true, DataCategory: PjDataCategories.pjMapTasks,
                                        TableName: "Task_Table1", HeaderRow: true, FieldName: "ID", ExternalFieldName: "ID", ImportMethod: PjImportMethods.pjImportNew);
            projectApplication.MapEdit(Name: "Map_Tasks", DataCategory: PjDataCategories.pjMapTasks, FieldName: "Unique ID", ExternalFieldName: "Unique_ID");
            projectApplication.MapEdit(Name: "Map_Tasks", DataCategory: PjDataCategories.pjMapTasks, FieldName: "Name", ExternalFieldName: "Name");
            projectApplication.MapEdit(Name: "Map_Tasks", DataCategory: PjDataCategories.pjMapTasks, FieldName: "Predecessors", ExternalFieldName: "Predecessors");
            projectApplication.MapEdit(Name: "Map_Tasks", DataCategory: PjDataCategories.pjMapTasks, FieldName: "Outline Level", ExternalFieldName: "Outline_Level");
            projectApplication.MapEdit(Name: "Map_Tasks", DataCategory: PjDataCategories.pjMapTasks, FieldName: "% Complete", ExternalFieldName: "Percent_Complete");
            projectApplication.MapEdit(Name: "Map_Tasks", DataCategory: PjDataCategories.pjMapTasks, FieldName: "Task Mode", ExternalFieldName: "Task_Mode");
            projectApplication.MapEdit(Name: "Map_Tasks", DataCategory: PjDataCategories.pjMapTasks, FieldName: "Type", ExternalFieldName: "Type");
            projectApplication.MapEdit(Name: "Map_Tasks", DataCategory: PjDataCategories.pjMapTasks, FieldName: "Milestone", ExternalFieldName: "Milestone");

            //Create Resources Map
            projectApplication.MapEdit(Name: "Map_Resources", Create: true, OverwriteExisting: true, DataCategory: PjDataCategories.pjMapResources,
                                        TableName: "Resource_Table1", HeaderRow: true, FieldName: "Unique ID", ExternalFieldName: "Unique_ID",
                                        ImportMethod: PjImportMethods.pjImportMerge, MergeKey: "Unique ID");
            projectApplication.MapEdit(Name: "Map_Resources", DataCategory: PjDataCategories.pjMapResources, FieldName: "ID", ExternalFieldName: "ID");
            projectApplication.MapEdit(Name: "Map_Resources", DataCategory: PjDataCategories.pjMapResources, FieldName: "Name", ExternalFieldName: "Name");
            projectApplication.MapEdit(Name: "Map_Resources", DataCategory: PjDataCategories.pjMapResources, FieldName: "Type", ExternalFieldName: "Type");
            projectApplication.MapEdit(Name: "Map_Resources", DataCategory: PjDataCategories.pjMapResources, FieldName: "Material Label", ExternalFieldName: "Material_Label");
            projectApplication.MapEdit(Name: "Map_Resources", DataCategory: PjDataCategories.pjMapResources, FieldName: "Group", ExternalFieldName: "Group_Name");
            projectApplication.MapEdit(Name: "Map_Resources", DataCategory: PjDataCategories.pjMapResources, FieldName: "Standard Rate", ExternalFieldName: "Standard_Rate");

            //Create "Map_Material_Assignments"            
            projectApplication.MapEdit(Name: "Map_Material_Assignments", Create: true, OverwriteExisting: true, DataCategory: PjDataCategories.pjMapAssignments,
                                       TableName: "Assignment_Table1", HeaderRow: true, FieldName: "Resource Unique ID", ExternalFieldName: "Resource_Unique_ID",
                                       ImportMethod: PjImportMethods.pjImportMerge, MergeKey: "Resource Unique ID");
            projectApplication.MapEdit(Name: "Map_Material_Assignments", DataCategory: PjDataCategories.pjMapAssignments, FieldName: "Task Unique ID", ExternalFieldName: "Task_Unique_ID");
            projectApplication.MapEdit(Name: "Map_Material_Assignments", DataCategory: PjDataCategories.pjMapAssignments, FieldName: "Work", ExternalFieldName: "Scheduled_Work");

            //Create "Map_Labour_Assignments"
            projectApplication.MapEdit(Name: "Map_Labour_Assignments", Create: true, OverwriteExisting: true, DataCategory: PjDataCategories.pjMapAssignments,
                                       TableName: "Assignment_Table1", HeaderRow: true, FieldName: "Resource Unique ID", ExternalFieldName: "Resource_Unique_ID",
                                       ImportMethod: PjImportMethods.pjImportMerge, MergeKey: "Resource Unique ID");
            projectApplication.MapEdit(Name: "Map_Labour_Assignments", DataCategory: PjDataCategories.pjMapAssignments, FieldName: "Task Unique ID", ExternalFieldName: "Task_Unique_ID");
            projectApplication.MapEdit(Name: "Map_Labour_Assignments", DataCategory: PjDataCategories.pjMapAssignments, FieldName: "Cost", ExternalFieldName: "Cost");

            //Create "Map_Tasks_Progress"
            projectApplication.MapEdit(Name: "Map_Tasks_Progress", Create: true, OverwriteExisting: true, DataCategory: PjDataCategories.pjMapTasks,
                                       TableName: "Task_Table1", HeaderRow: true, FieldName: "Unique ID", ExternalFieldName: "Unique_ID",
                                       ImportMethod: PjImportMethods.pjImportMerge, MergeKey: "Unique ID");
            projectApplication.MapEdit(Name: "Map_Tasks_Progress", DataCategory: PjDataCategories.pjMapTasks, FieldName: "% Complete", ExternalFieldName: "Percent_Complete");
            projectApplication.MapEdit(Name: "Map_Tasks_Progress", DataCategory: PjDataCategories.pjMapTasks, FieldName: "Task Mode", ExternalFieldName: "Task_Mode");
        }

        private static void SaveSelectedProjectAsExcelFiles(TransferOptions transferOptions, Application projectApplication)
        {
            ExcelHelpers.DeleteExcelMapFiles();
            //Open the Chosen Ms Project File
            projectApplication.FileOpenEx(Name: transferOptions.MsProjectFileName);

            //Save the Chosen Ms Project File as Excels, using related Maps
            projectApplication.FileSaveAs(Name: Globals.ExcelMapFilePaths.TasksPath, PjFileFormat.pjXLSX, Map: "Map_Tasks");

            projectApplication.DisplayAlerts = false;
        }

        private static ExcelMap UpdateProjectProgress(TransferOptions transferOptions, Application projectApplication, List<PercentCompleteTotal> percentCompleteTotal)
        {
            SaveSelectedProjectAsExcelFiles(transferOptions, projectApplication);

            //The selected Project File is Saved as Excel file on above line,
            //below converts the Tasks Excel file to type ExcelMap (ExcelMap.ExcelMapsTasks is filled only)            
            var selectedProjectTasksAsExcelMapForComparison = CreateExelMapTasksFromTasksExcelFile();

            //Below is joining tasks in selected Project File(selectedProjectTasksAsExcelMapForComparison) with percentCompleteTotal (which contains progress)            

            var currentTasks = selectedProjectTasksAsExcelMapForComparison.ExcelMapTasks;
            ExcelMap newMapWithuniqeIdAndProgressOnly = new ExcelMap
            {
                ExcelMapTasks = (from cT in currentTasks
                                 join pCT in percentCompleteTotal on new
                                 {
                                     x1 = cT.Area,
                                     x2 = cT.Floor,
                                     x3 = cT.DrawingNumber,
                                     x4 = cT.SubZone,
                                     x5 = cT.WallType,
                                     x6 = cT.MeasomRef,
                                     x7 = cT.LabourPhase
                                 } equals new
                                 {
                                     x1 = pCT.Area,
                                     x2 = pCT.Floor,
                                     x3 = pCT.DrawingNumber,
                                     x4 = pCT.SubZone,
                                     x5 = pCT.WallType,
                                     x6 = pCT.MeasomRef,
                                     x7 = pCT.LabourPhase
                                 }
                                 select new ExcelMapTask
                                 {
                                     UniqeId = cT.UniqeId,
                                     PercentComplete = pCT.TaskProgressTotal,
                                     TaskMode = "Auto Scheduled"
                                 }).ToList()
            };
            return newMapWithuniqeIdAndProgressOnly;
        }

        private static ExcelMap CreateExelMapTasksFromTasksExcelFile()
        {
            DataRow nextOccuranceof2 = null;
            int index = 0;
            int index2 = 0;
            string Area, Floor, DrawingNumber, SubZone, WallType, MeasomRef, LabourPhase;
            Area = Floor = DrawingNumber = SubZone = WallType = MeasomRef = LabourPhase = string.Empty;
            bool lastLoopTrigger = false;
            bool doLoop = true;
            DataTable prectTasksDT;
            EnumerableRowCollection<DataRow> taksDTEnumarable;

            //This is the list to attach progress and create a new Ms Project file from
            ExcelMap excelMap = new ExcelMap();

            prectTasksDT = ReturnExcelTasksAsDT();
            taksDTEnumarable = prectTasksDT.AsEnumerable();

            //Outline Level 1 = Project Name
            //Outline Level 2 = Area

            var firstRowOccuranceof2 = taksDTEnumarable.Where(row => row["Outline_Level"].ToString() == "2").FirstOrDefault();
            index = prectTasksDT.Rows.IndexOf(firstRowOccuranceof2);

            //Now find the Next occurance to select the rows to process , THIS SHOULD BE A LOOP UNTIL THE END OF DATABLE
            nextOccuranceof2 = taksDTEnumarable.Skip(index + 1).Where(row => row["Outline_Level"].ToString() == "2").FirstOrDefault();
            index2 = prectTasksDT.Rows.IndexOf(nextOccuranceof2);

            if (index2 == -1)
            {
                lastLoopTrigger = true;
            }

            do
            {
                int outlineLevel = 0;
                IEnumerable<DataRow> rowsToProcess;
                int maxOutLineLevel;
                if (lastLoopTrigger)
                {
                    rowsToProcess = taksDTEnumarable.Skip(index).Take(taksDTEnumarable.Count() - index);
                }
                else
                {
                    rowsToProcess = taksDTEnumarable.Skip(index).Take(index2 - index);
                }
                maxOutLineLevel = rowsToProcess.Max(row => int.Parse(row["Outline_Level"].ToString()));

                //Populate excelMap.ExcelMapTasks here (Which will be the map file, without progress initially)
                foreach (var row in rowsToProcess)
                {
                    outlineLevel = int.Parse(row["Outline_Level"].ToString());
                    int propCount = typeof(TaskBase).GetProperties().Count();
                    var data = propCount - maxOutLineLevel + outlineLevel;

                    switch (data)
                    {
                        case 1:
                            Area = row["Name"].ToString();
                            break;
                        case 2:
                            Floor = row["Name"].ToString();
                            break;
                        case 3:
                            DrawingNumber = row["Name"].ToString();
                            break;
                        case 4:
                            SubZone = row["Name"].ToString();
                            break;
                        case 5:
                            WallType = row["Name"].ToString();
                            break;
                        case 6:
                            MeasomRef = row["Name"].ToString();
                            break;
                        case 7:
                            LabourPhase = row["Name"].ToString();
                            break;
                    }

                    if (outlineLevel == maxOutLineLevel)
                    {
                        excelMap.ExcelMapTasks.Add(new ExcelMapTask
                        {
                            Area = Area,
                            Floor = Floor,
                            DrawingNumber = DrawingNumber,
                            SubZone = SubZone,
                            WallType = WallType,
                            MeasomRef = MeasomRef,
                            LabourPhase = LabourPhase,
                            Id = row["ID"].ToString(),
                            UniqeId = row["Unique_ID"].ToString(),
                            PercentComplete = parseStringTo5DigitDecimal(row["Percent_Complete"].ToString())
                        });
                    }
                }
                if (lastLoopTrigger)
                {
                    doLoop = false;
                }
                index = index2;
                nextOccuranceof2 = taksDTEnumarable.Skip(index + 1).Where(row => row["Outline_Level"].ToString() == "2").FirstOrDefault();
                index2 = prectTasksDT.Rows.IndexOf(nextOccuranceof2);
                if (index2 == -1)
                {
                    lastLoopTrigger = true;
                }
                Area = Floor = DrawingNumber = SubZone = WallType = MeasomRef = LabourPhase = string.Empty;
            } while (doLoop);

            return excelMap;
        }
        private static decimal parseStringTo5DigitDecimal(string strVal)
        {
            if (strVal.Length > 5)
            {
                return decimal.Parse(strVal.ToString().Substring(0, 5));
            }
            else
            {
                return decimal.Parse(strVal.ToString());
            }

        }

        private static DataTable ReturnExcelTasksAsDT()
        {
            //Return the saved Excel TASKS File as DataTable
            return ExcelHelpers.ReturnExcelSheetAsDataTable(Globals.ExcelMapFilePaths.TasksPath);
        }

        private static void CreateNewProjectFromMaps(Application projectApplication)
        {
            projectApplication.FileOpenEx(Name: Globals.ExcelMapFilePaths.TasksPath, Merge: PjMergeType.pjDoNotMerge,
                                        Map: "Map_Tasks");
            projectApplication.FileOpenEx(Name: Globals.ExcelMapFilePaths.RsrcPath, Merge: PjMergeType.pjMerge,
                        Map: "Map_Resources");
            projectApplication.FileOpenEx(Name: Globals.ExcelMapFilePaths.MatAsggnmtPath, Merge: PjMergeType.pjMerge,
                        Map: "Map_Material_Assignments");
            projectApplication.FileOpenEx(Name: Globals.ExcelMapFilePaths.LabAsggnmtPath, Merge: PjMergeType.pjMerge,
                        Map: "Map_Labour_Assignments");
        }
    }
}
