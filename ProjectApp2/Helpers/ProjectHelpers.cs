using ProjectApp2.Model;
using ProjectApp2.Model.Abstract;
using ProjectApp2.Model.enums;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using System.Text;
using System.Windows;

namespace ProjectApp2.Helpers
{
    public class ProjectHelpers
    {
        public static void InitializeOperation(TransferOptions transferOptions)
        {
            string path = Path.GetTempPath();

            Globals.ExcelMapFilePaths = new ExcelMapFilePaths
            {
                TasksPath = Path.GetFullPath(path + @"\TemplateProject.xlsx"),
                RsrcPath = Path.GetFullPath(path + @"\ExcelMapResources.xlsx"),
                MatAsggnmtPath = Path.GetFullPath(path + @"\ExcelMapMatrAssgnmnts.xlsx"),
                LabAsggnmtPath = Path.GetFullPath(path + @"\ExcelMapLabrAssgnmnts.xlsx"),
                ErrorTextFile = Path.GetFullPath(path + @"\ErrorTextFile.txt")
            };

            var erPath = Globals.ExcelMapFilePaths.ErrorTextFile;
            if (File.Exists(erPath))
            {
                File.Delete(erPath);
            }

            List<RateSheet> rateSheets = null;
            //Create Task Data Table (which contains in its rows each Task in Ms Project)
            var tasksDataTable = CreateTaskDataTable(transferOptions);
            MsProjectHelpers msProjectHelpers = new MsProjectHelpers();
            tasksDataTable.Rows[0].Delete();
            tasksDataTable.AcceptChanges();

            //Below returns Ordering Sheet
            DataTable orderExcelDT = ExcelHelpers.ReturnOrderExcelAsDT(transferOptions);

            PercCompleteOrderings percCompleteOrderings = PercCompleteOrdersAsList(orderExcelDT);
            if (!string.IsNullOrEmpty(percCompleteOrderings.ErrorOnOrdering))
            {
                return;
            }

            //Below maps tasksDataTable to a list of type PercentCompleteTotal
            List<PercentCompleteTotal> percentCompleteTotal = ReturnTaskDataSetAsList(tasksDataTable, transferOptions, percCompleteOrderings);

            if (transferOptions.transferType == ProjectTransferType.NewProject)
            {
                //Below returns all sheets in Rate Excel to 1 List
                rateSheets = GetAllExcelRateSheetsAsOneList(transferOptions);
            }

            //Main Transfer to Ms Project operation using Tasks and Rate lists created above                
            msProjectHelpers.DoMainOperationsCreateOrUpdateProject(percentCompleteTotal, rateSheets, transferOptions);
        }
        private static bool ShowOrderingError(List<PercCompleteOrder> listToCheck, string itemType)
        {
            if (listToCheck.Where(ao => ao.OrderNumber != 0).Any(ao => !string.IsNullOrEmpty(ao.UnOrderedItemValue)))
            {
                MessageBox.Show($"There are unordered {itemType}s, Operation is terminating!!!", $"{itemType} Ordering Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return true;
            }
            return false;
        }

        private static PercCompleteOrderings PercCompleteOrdersAsList(DataTable orderExcelDT)
        {
            PercCompleteOrderings percCompleteOrderings = new PercCompleteOrderings();

            percCompleteOrderings.PercentCompleteAreaOrder = PercCompleteAreaOrderAsList(orderExcelDT);
            if (ShowOrderingError(percCompleteOrderings.PercentCompleteAreaOrder, "Area"))
                percCompleteOrderings.ErrorOnOrdering = "Area";

            percCompleteOrderings.PercentCompleteFloorOrder = PercCompleteFloorOrderAsList(orderExcelDT);
            if (ShowOrderingError(percCompleteOrderings.PercentCompleteFloorOrder, "Floor"))
                percCompleteOrderings.ErrorOnOrdering = "Floor";

            percCompleteOrderings.PercentCompleteSubZoneOrder = PercCompleteSubZoneOrderAsList(orderExcelDT);
            if (ShowOrderingError(percCompleteOrderings.PercentCompleteSubZoneOrder, "SubZone"))
                percCompleteOrderings.ErrorOnOrdering = "SubZone";

            return percCompleteOrderings;
        }

        private static List<PercCompleteOrder> PercCompleteAreaOrderAsList(DataTable orderExcelDT)
        {
            orderExcelDT.Rows[1].Delete();
            orderExcelDT.AcceptChanges();

            //Add Areas
            int orderNo = 1;
            var listToReturn = (from DataRow dr in orderExcelDT.Rows
                                where !string.IsNullOrEmpty(dr[0].ToString().Trim()) || !string.IsNullOrEmpty(dr[1].ToString().Trim())
                                select new PercCompleteOrder()
                                {
                                    UnOrderedItemValue = dr[0].ToString().Trim(),
                                    ItemValue = dr[1].ToString().Trim(),
                                    OrderingItemType = TaskOrder.Area,
                                    OrderNumber = orderNo++
                                }).ToList();
            return listToReturn;

        }
        private static List<PercCompleteOrder> PercCompleteFloorOrderAsList(DataTable orderExcelDT)
        {
            //Add Areas
            int orderNo = 1;
            var listToReturn = (from DataRow dr in orderExcelDT.Rows
                                where !string.IsNullOrEmpty(dr[3].ToString().Trim()) || !string.IsNullOrEmpty(dr[4].ToString().Trim())
                                select new PercCompleteOrder()
                                {
                                    UnOrderedItemValue = dr[3].ToString().Trim(),
                                    ItemValue = dr[4].ToString().Trim(),
                                    OrderingItemType = TaskOrder.Floor,
                                    OrderNumber = orderNo++
                                }).ToList();
            return listToReturn;
        }
        private static List<PercCompleteOrder> PercCompleteSubZoneOrderAsList(DataTable orderExcelDT)
        {
            //Add Areas
            int orderNo = 1;
            var listToReturn = (from DataRow dr in orderExcelDT.Rows
                                where !string.IsNullOrEmpty(dr[6].ToString().Trim()) || !string.IsNullOrEmpty(dr[7].ToString().Trim())
                                select new PercCompleteOrder()
                                {
                                    UnOrderedItemValue = dr[6].ToString().Trim(),
                                    ItemValue = dr[7].ToString().Trim(),
                                    OrderingItemType = TaskOrder.SubZone,
                                    OrderNumber = orderNo++
                                }).ToList();
            return listToReturn;
        }
        public static DataTable CreateTaskDataTable(TransferOptions transferOptions)
        {
            DataTable tasksDataTable;
            if (transferOptions.PercCompType == PercCompType.Excel)
            {
                //Below gets the PercentComplete Excel as DataSet
                tasksDataTable = ExcelHelpers.ReturnExcelSheetAsDataTable(transferOptions.ExcelTasksFileName);
            }
            else
            {
                string sqlString = "SELECT * from " + transferOptions.DatabaseItems.PercCompleteTableName;
                //Try to return data from table to see if there is an error           
                tasksDataTable = SQLHelpers.GetDTFromSqlString(transferOptions.DatabaseItems, sqlString);
            }

            return tasksDataTable;
        }

        private static List<RateSheet> GetAllExcelRateSheetsAsOneList(TransferOptions transferOptions)
        {
            List<string> ExcelRateSheetNames = ExcelHelpers.GetOleDbExcelSheetNames(transferOptions.ExcelRateSheetFileName);
            List<RateSheet> RateSheets = new List<RateSheet>();
            var strBuilder = new StringBuilder();
            foreach (string ExcelRateSheet in ExcelRateSheetNames)
            {
                var excelRateSheetDataTable = ExcelHelpers.ReturnExcelSheetAsDataTable(transferOptions.ExcelRateSheetFileName, ExcelRateSheet);
                var returnedList = ReturnExcelRateSheetDataTableAsList(excelRateSheetDataTable, ExcelRateSheet, strBuilder);
                if (returnedList != null)
                {
                    RateSheets.AddRange(returnedList);
                }
            }

            if (strBuilder.Length != 0)
            {
                var erPath = Globals.ExcelMapFilePaths.ErrorTextFile;
                File.Create(erPath).Close();
                File.WriteAllText($"{erPath}", strBuilder.ToString());
            }
            return RateSheets;
        }
        private static List<RateSheet> ReturnExcelRateSheetDataTableAsList(DataTable excelRateSheetDataTable, string excelRateSheet, StringBuilder strBuilder)
        {
            //Data preperation and finding values from datatable
            string specification = string.Empty;
            string projectName = string.Empty;

            List<RateSheet> rateSheets = null;

            try
            {
                //Find row of first Occurence of "Specification" and get its value ( from known / hardcoded/ column numbers => 19 &  26)            
                specification = excelRateSheetDataTable.Rows[RowNoDT(excelRateSheetDataTable, "Specification", 19)][26].ToString();
                projectName = excelRateSheetDataTable.Rows[RowNoDT(excelRateSheetDataTable, "Project", 0)][3].ToString();
            }
            catch
            {
                strBuilder.AppendLine($"Error in Ratesheet Worksheet {excelRateSheet.Replace("$", string.Empty)}, Rate for this Worksheet could not be created!!");
                return rateSheets;
            }

            //Find row of first Occurence of "Item"            
            int itemRowNo = RowNoDT(excelRateSheetDataTable, "Item", 0);

            //Find row of first Occurence of "Revision Notes"
            int revisionNotes = RowNoDT(excelRateSheetDataTable, "Revision Notes", 0);

            //Find row of first Occurence of "Adjustment"
            int adjustmentRowNo = RowNoDT(excelRateSheetDataTable, "Adjustment", 18);

            //Find row of first Occurence of "Partition Area" Header
            int partitionAreaHeader = RowNoDT(excelRateSheetDataTable, "Partition Area", 12);
            string partitionAreaOrg = excelRateSheetDataTable.Rows[partitionAreaHeader][15].ToString();
            decimal partitionArea = DC(new string(partitionAreaOrg.Where(c => char.IsDigit(c)).ToArray()));

            // summing 2 columns cause it might be in any 1 of the columns while other column is 0 
            decimal adjustment = DC(excelRateSheetDataTable.Rows[adjustmentRowNo][32])
                        + DC(excelRateSheetDataTable.Rows[adjustmentRowNo][38]);

            //delete unnecessary datable rows
            for (int i = excelRateSheetDataTable.Rows.Count - 1; i >= 0; i--)
            {
                DataRow dr = excelRateSheetDataTable.Rows[i];
                if (dr[0].ToString() == string.Empty
                    || i >= revisionNotes
                    || i <= itemRowNo
                    )
                {
                    dr.Delete();
                }
            }
            excelRateSheetDataTable.AcceptChanges();

            rateSheets = (from DataRow dr in excelRateSheetDataTable.Rows
                          select new RateSheet()
                          {
                              ProjectName = projectName,
                              Specification = specification,
                              ClientRef = string.Empty,
                              Adjustment = adjustment,
                              Item = dr[0].ToString().Trim(),
                              ProductCode = dr[2].ToString().Trim(),
                              Unit = dr[14].ToString().Trim(),
                              MaterialCost = DC(dr[24]) / partitionArea,
                              LabourPhase = dr[29].ToString().Trim(),
                              LabourCost = DC(dr[30]),
                              OtherCost = DC(dr[35]),
                              MaterialWasteCost = DC(dr[24]) * DC(dr[40]) / partitionArea,
                              MaterialMarkup = DC(dr[24]) * (1 + DC(dr[40])) * DC(dr[44]) / partitionArea,
                              LabourMarkup = DC(dr[30]) * DC(dr[47]),
                              OtherMarkup = DC(dr[35]) * DC(dr[50])
                          }).ToList();

            //sum all Total Values Before Adjustment
            decimal MaterialTotalWithoutAdjustment = rateSheets.Sum(x => x.MaterialTotalWithoutAdjustment);
            decimal LabourTotalWithoutAdjustment = rateSheets.Sum(x => x.LabourTotalWithoutAdjustment);
            decimal OtherTotalWithoutAdjustment = rateSheets.Sum(x => x.OtherTotalWithoutAdjustment);
            decimal TotalWithoutAdjustment = MaterialTotalWithoutAdjustment + LabourTotalWithoutAdjustment + OtherTotalWithoutAdjustment;
            decimal MaterialAdjustment = MaterialTotalWithoutAdjustment * adjustment / TotalWithoutAdjustment;
            decimal LabourAdjustment = LabourTotalWithoutAdjustment * adjustment / TotalWithoutAdjustment;
            decimal OtherAdjustment = OtherTotalWithoutAdjustment * adjustment / TotalWithoutAdjustment;

            //Assign adjustment value to list items
            foreach (RateSheet item in rateSheets)
            {
                item.MaterialAdjustment = MaterialTotalWithoutAdjustment == 0 ? 0 : item.MaterialTotalWithoutAdjustment / MaterialTotalWithoutAdjustment * MaterialAdjustment;
                item.LabourAdjustment = LabourTotalWithoutAdjustment == 0 ? 0 : item.LabourTotalWithoutAdjustment / LabourTotalWithoutAdjustment * LabourAdjustment;
                item.OtherAdjustment = OtherTotalWithoutAdjustment == 0 ? 0 : item.OtherTotalWithoutAdjustment / OtherTotalWithoutAdjustment * OtherAdjustment;
            }

            return rateSheets;
        }

        public static List<PercentCompleteTotal> ReturnTaskDataSetAsList(DataTable excelTasksDataTable, TransferOptions transferOptions,
                                                                        PercCompleteOrderings percCompleteOrderings)
        {
            var sqlString = "SELECT phasename, [order] from " + transferOptions.DatabaseItems.LabourPhaseTableName;
            DataTable phaseOrderDt = SQLHelpers.GetDTFromSqlString(transferOptions.DatabaseItems, sqlString);

            var phaseOrderList = (from row in phaseOrderDt.AsEnumerable()
                                  select new
                                  {
                                      PhaseName = row[0].ToString().Trim(),
                                      Order = row[1].ToString().Trim()
                                  }).ToList();

            List<PercentComplete> dataListSpecs = ConvertDTToList(excelTasksDataTable);

            List<PercentCompleteWithProgress> listWithProgress = GenerateProgressForTasks(dataListSpecs);

            List<PercentCompleteTotal> dataListSpecsWithWallTotals2 = GenerateTotals(listWithProgress);

            PopulateBlankValues(dataListSpecsWithWallTotals2);

            //Make a Join with Main list(dataListSpecsWithWallTotals2) and the PhaseOrder list in another list(dataListSpecsWithWallTotals)
            //          for clarity (And to use Linq Query Syntax for join)

            List<PercentCompleteTotal> dataListSpecsWithWallTotals = (from d in dataListSpecsWithWallTotals2
                                                                      join p in phaseOrderList on d.LabourPhase equals p.PhaseName
                                                                      join pcoa in percCompleteOrderings.PercentCompleteAreaOrder
                                                                      on d.Area equals pcoa.ItemValue into lpcoa
                                                                      from pcoa in lpcoa.DefaultIfEmpty()
                                                                      join pcof in percCompleteOrderings.PercentCompleteFloorOrder
                                                                      on d.Floor equals pcof.ItemValue into lpcof
                                                                      from pcof in lpcof.DefaultIfEmpty()
                                                                      join pcos in percCompleteOrderings.PercentCompleteSubZoneOrder
                                                                      on d.SubZone equals pcos.ItemValue into lpcos
                                                                      from pcos in lpcos.DefaultIfEmpty()
                                                                      select new PercentCompleteTotal
                                                                      {
                                                                          Area = d.Area,
                                                                          Floor = d.Floor,
                                                                          DrawingNumber = d.DrawingNumber,
                                                                          SubZone = d.SubZone,
                                                                          WallType = d.WallType,
                                                                          MeasomRef = d.MeasomRef,
                                                                          LabourPhase = d.LabourPhase,
                                                                          TotalMeasureTotal = d.TotalMeasureTotal,
                                                                          TaskProgressTotal = d.TaskProgressTotal,
                                                                          PhaseOrder = int.Parse(p.Order),
                                                                          TotalManHoursTotal = d.TotalManHoursTotal,
                                                                          AreaOrder = pcoa != null ? pcoa.OrderNumber : 0,
                                                                          FloorOrder = pcof != null ? pcof.OrderNumber : 0,
                                                                          SubZoneOrder = pcos != null ? pcos.OrderNumber : 0
                                                                      }).OrderBy(h => h.AreaOrder).
                                                                         ThenBy(h => h.FloorOrder).
                                                                         ThenBy(h => h.DrawingNumber).
                                                                         ThenBy(h => h.SubZoneOrder).
                                                                         ThenBy(h => h.WallType).
                                                                         ThenBy(h => h.MeasomRef).
                                                                         ThenBy(h => h.PhaseOrder)
                                                                         .ToList();

            return dataListSpecsWithWallTotals;
        }

        private static List<PercentCompleteTotal> GenerateTotals(List<PercentCompleteWithProgress> listWithProgress)
        {
            return listWithProgress.
                GroupBy(y => new
                {
                    y.Area,
                    y.Floor,
                    y.DrawingNumber,
                    y.SubZone,
                    y.WallType,
                    y.MeasomRef,
                    y.LabourPhase
                }).
            Select(h => new PercentCompleteTotal
            {
                Area = h.Key.Area,
                Floor = h.Key.Floor,
                DrawingNumber = h.Key.DrawingNumber,
                SubZone = h.Key.SubZone,
                WallType = h.Key.WallType,
                MeasomRef = h.Key.MeasomRef,
                LabourPhase = h.Key.LabourPhase,
                TotalMeasureTotal = (decimal)h.Sum(y => y.TotalMeasure),
                TaskProgressTotal = Math.Truncate(h.Sum(y => y.ProgressGainForTask) * 100) / 100,
                TotalManHoursTotal = (decimal)h.Sum(y => y.TotalManhour)
            }).ToList();
        }

        private static List<PercentCompleteWithProgress> GenerateProgressForTasks(List<PercentComplete> dataListSpecs)
        {

            //This list will get the task progress for each line. A task is made of different lines in PercentComlete file because 
            //each line is wall so a progress in wall will effect the task's progress according to its measure weight in task total measure
            return dataListSpecs.
                GroupBy(dls => new
                {
                    dls.Area,
                    dls.Floor,
                    dls.DrawingNumber,
                    dls.SubZone,
                    dls.WallType,
                    dls.MeasomRef,
                    dls.LabourPhase
                }).
            SelectMany(nl => nl, (dlst, dlssub) => new PercentCompleteWithProgress
            {
                Area = dlst.Key.Area,
                Floor = dlst.Key.Floor,
                DrawingNumber = dlst.Key.DrawingNumber,
                SubZone = dlst.Key.SubZone,
                WallType = dlst.Key.WallType,
                MeasomRef = dlst.Key.MeasomRef,
                LabourPhase = dlst.Key.LabourPhase,
                TotalMeasure = (double)dlssub.TotalMeasure,
                TotalManhour = (double)dlssub.TotalManHours,
                ProgressGainForTask = CalculateProgresGain(dlssub, dlst)
            }).ToList();
        }

        private static List<PercentComplete> ConvertDTToList(DataTable excelTasksDataTable)
        {
            //Mapping Excel Columns to object properties is done with Excel column numbers,
            //can also be done with Excel Column headers like : dr["Labour Phase"] as long as correct headers are given to datatable
            return (from DataRow dr in excelTasksDataTable.Rows
                    select new PercentComplete()
                    {
                        LabourPhase = dr[0].ToString().Trim(),
                        ClientRef = dr[1].ToString().Trim(),
                        HeightBand = DC(dr[2]),
                        MeasomRef = dr[3].ToString().Trim(),
                        Area = dr[4].ToString().Trim(),
                        Floor = dr[5].ToString().Trim(),
                        DrawingNumber = dr[6].ToString().Trim(),
                        SubZone = dr[7].ToString().Trim(),
                        WallNumber = dr[8].ToString().Trim(),
                        InstalledOn = dr[9].ToString().Trim(),
                        InstalledBy = dr[10].ToString().Trim(),
                        TotalMeasure = TotalMeasure(dr),
                        UOM = dr[12].ToString().Trim(),
                        TotalValue = DC(dr[13]),
                        TotalManHours = DC(dr[14]),
                        MeasureInstalled = MeasureInstalled(dr),
                        TotalValueInstalled = DC(dr[16]),
                        ManHoursEarnt = DC(dr[17]),
                        PercentMeasureIntalled = DC(dr[18]),
                        PercentValueInstalled = DC(dr[19]),
                        PercentManHoursEarned = DC(dr[20]),
                        WallType = dr[21].ToString().Trim()
                    }).ToList();
        }
        private static void PopulateBlankValues(List<PercentCompleteTotal> dataListSpecsWithWallTotals2)
        {
            PropertyInfo[] properties = dataListSpecsWithWallTotals2[0].GetType().GetProperties();

            foreach (var item in dataListSpecsWithWallTotals2)
            {
                foreach (PropertyInfo property in properties)
                {
                    //check if that propertry is in taskbase
                    if (typeof(TaskBase).GetProperty(property.Name) != null)
                    {
                        var propertyValue = (item.GetType().GetProperty(property.Name).
                               GetValue(item) ?? string.Empty).ToString();
                        if (propertyValue == string.Empty)
                        {
                            property.SetValue(item, $"#Unkown {property.Name}#");
                        }
                    }
                }
            }
        }

        private static decimal MeasureInstalled(DataRow dr)
        {
            string unit = dr[12].ToString().Trim();
            decimal measure = DC(dr[15]);
            decimal heightBand = DC(dr[2]);
            string wallType = dr[21].ToString().Trim();

            if (unit.ToUpper() == "M" && wallType.ToUpper() == "WALL")
            {
                return measure * heightBand / 1000;
            }
            else
            {
                return measure;
            }
        }

        private static decimal TotalMeasure(DataRow dr)
        {
            string unit = dr[12].ToString().Trim();
            decimal measure = DC(dr[11]);
            decimal heightBand = DC(dr[2]);
            string wallType = dr[21].ToString().Trim();

            if (unit.ToUpper() == "M" && wallType.ToUpper() == "WALL")
            {
                return measure * heightBand / 1000;
            }
            else
            {
                return measure;
            }
        }

        private static decimal CalculateProgresGain<T>(PercentComplete percentComplete, IGrouping<T, PercentComplete> percentCompletes)
        {
            decimal output = 0;
            if (percentComplete.TotalMeasure * (percentComplete.TotalMeasure / percentCompletes.Sum(tm => tm.TotalMeasure)) != 0)
            {
                output = percentComplete.MeasureInstalled / percentComplete.TotalMeasure * (percentComplete.TotalMeasure / percentCompletes.Sum(tm => tm.TotalMeasure));
            }
            return output;
        }

        private static decimal DC(object dataColumn)
        {
            decimal d = 0;
            decimal output = 0;
            string obj = dataColumn as String;
            if (obj == null)
            {
                obj = dataColumn.ToString();
            }

            if (obj != null)
            {
                obj = obj.Replace("£", "").ToString().Trim();
                if (obj.Contains("%"))
                {
                    obj = obj.Replace("%", "");
                    if (decimal.TryParse(obj, out d))
                    {
                        output = decimal.Parse(obj) / 100;
                    }
                }
                else
                {
                    if (decimal.TryParse(obj, out d))
                    {
                        output = decimal.Parse(obj);
                    }
                }
            }
            return output;
        }

        private static int RowNoDT(DataTable DT, string valueToSearch, int columnNumber)
        {
            DataRow dr = DT.AsEnumerable().FirstOrDefault(r => r[columnNumber].ToString() == valueToSearch);
            return DT.Rows.IndexOf(dr);
        }


        public static string[,] ReturnDTAs2DArray(DataTable dt)
        {
            var colCount = dt.Columns.Count;
            string[,] str = new string[dt.Rows.Count, colCount];

            int i = 0;
            foreach (DataRow row in dt.Rows)
            {
                for (var j = 0; j < colCount; j++)
                {
                    str[i, j] = row[j].ToString();
                }
                i++;
            }

            return str;
        }

        public static string[,] ReturnListAs2DArray<T>(List<T> listContainingStrings, string onlyColumns = "")
        {
            int i = 0;
            int j = 0;
            string[] columnNames = null;

            int rowCount = listContainingStrings.Count;
            int colCount = typeof(T).GetProperties().Count();
            string[,] myArr = new string[rowCount, colCount];

            if (!string.IsNullOrEmpty(onlyColumns))
            {
                columnNames = onlyColumns.Split(',');
            }

            foreach (var item in listContainingStrings)
            {
                foreach (PropertyInfo PI in typeof(T).GetProperties())
                {
                    if (!string.IsNullOrEmpty(onlyColumns))
                    {
                        if (columnNames.Contains(PI.Name))
                        {
                            myArr[i, j] = typeof(T).GetProperty(PI.Name).GetValue(item)?.ToString() ?? string.Empty;
                            j++;
                        }
                    }
                    else
                    {
                        myArr[i, j] = typeof(T).GetProperty(PI.Name).GetValue(item)?.ToString() ?? string.Empty;
                        j++;
                    }
                }
                j = 0;
                i++;
            }
            return myArr;
        }
    }
}
