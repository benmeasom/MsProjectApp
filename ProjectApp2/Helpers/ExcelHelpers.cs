using Microsoft.Office.Interop.MSProject;
using ProjectApp2.Model;
using ProjectApp2.Model.enums;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using static ProjectApp2.Model.ExcelMap;
using Excel = Microsoft.Office.Interop.Excel;

namespace ProjectApp2.Helpers
{
    public class ExcelHelpers
    {
        public static Excel.Workbook GetNewExelWorkbook(Excel.Application excelApp)
        {
            var wb = excelApp.Workbooks.Add();
            return wb;
        }

        public static void MainOperationInitialize(ExcelMap excelMap, TransferOptions transferOptions, List<PercentCompleteTotal> percentCompleteTotal)
        {
            DeleteExcelMapFiles();
            if (transferOptions.transferType == ProjectTransferType.NewProject)
            {
                ExportMapsToExcel(excelMap);
            }
            MsProjectHelpers.MainOperationMsProjectOperations(excelMap, transferOptions, percentCompleteTotal);
        }

        public static void ExportMapsToExcel(ExcelMap excelMap, bool onlyTaskProgress = false)
        {
            Excel.Workbook wb = null;
            Excel.Application excelApp = new Excel.Application();

            if (onlyTaskProgress)
            {
                wb = GetNewExelWorkbook(excelApp);
                ExportTasksProgressToExcelAndCloseWb(excelMap, wb, Globals.ExcelMapFilePaths.TasksPath);
            }
            else
            {
                if (excelMap.ExcelMapTasks.Count > 0)
                {
                    wb = GetNewExelWorkbook(excelApp);
                    ExportTasksToExcelAndCloseWb(excelMap, wb, Globals.ExcelMapFilePaths.TasksPath);
                }
                if (excelMap.ExcelMapResources.Count > 0)
                {
                    wb = GetNewExelWorkbook(excelApp);
                    ExportResourcesToExcelAndCloseWb(excelMap, wb, Globals.ExcelMapFilePaths.RsrcPath);
                }
                if (excelMap.ExcelMapNewProjectMaterialAssignments.Count > 0)
                {
                    wb = GetNewExelWorkbook(excelApp);
                    ExportMaterialAssignmentsToExcelAndCloseWb(excelMap, wb, Globals.ExcelMapFilePaths.MatAsggnmtPath);
                }
                if (excelMap.ExcelMapCostAssignments.Count > 0)
                {
                    wb = GetNewExelWorkbook(excelApp);
                    ExportCostAssignmentsToExcelAndCloseWb(excelMap, wb, Globals.ExcelMapFilePaths.LabAsggnmtPath);
                }
            }

            excelApp.Quit();

            GC.Collect();
            GC.WaitForPendingFinalizers();

            Marshal.ReleaseComObject(wb);
            Marshal.ReleaseComObject(excelApp);

            excelApp = null;
        }

        public static void DeleteExcelMapFiles()
        {
            foreach (PropertyInfo PI in typeof(ExcelMapFilePaths).GetProperties())
            {
                if (PI.Name != "ErrorTextFile")
                {
                    var fileFullPath = typeof(ExcelMapFilePaths).GetProperty(PI.Name).GetValue(Globals.ExcelMapFilePaths)?.ToString() ?? string.Empty;
                    if (File.Exists(fileFullPath))
                        File.Delete(fileFullPath);
                }
            }
        }

        public static void DeleteCostAssgnmntExcelMapFile()
        {
            if (File.Exists(Globals.ExcelMapFilePaths.LabAsggnmtPath))
                File.Delete(Globals.ExcelMapFilePaths.LabAsggnmtPath);
        }

        public static void DeleteMatAssgnmntExcelMapFile()
        {
            if (File.Exists(Globals.ExcelMapFilePaths.MatAsggnmtPath))
                File.Delete(Globals.ExcelMapFilePaths.MatAsggnmtPath);
        }

        public static void DeleteTasksExcelMapFile()
        {
            if (File.Exists(Globals.ExcelMapFilePaths.TasksPath))
                File.Delete(Globals.ExcelMapFilePaths.TasksPath);
        }

        //**************IMPORTANT*******************
        // Below Headers in Excel have to be in the same order with the Properties in the Model Class (like ExcelMapTask)

        private static void ExportTasksProgressToExcelAndCloseWb(ExcelMap excelMap, Excel.Workbook wb, string wbFileName)
        {
            var wSheetName = "Task_Table1";
            var hdrs = "Unique_ID,Percent_Complete,Task_Mode";
            ExportListToExcel(excelMap.ExcelMapTasks, wb, wSheetName, hdrs, 0, "UniqeId,PercentComplete,Task_Mode");
            wb.SaveAs(wbFileName);
            wb.Close(false);
        }
        private static void ExportTasksToExcelAndCloseWb(ExcelMap excelMap, Excel.Workbook wb, string wbFileName)
        {
            var wSheetName = "Task_Table1";
            var hdrs = "ID,Unique_ID,Name,Predecessors,Outline_Level,Percent_Complete,Task_Mode,Type,Milestone";
            ExportListToExcel(excelMap.ExcelMapTasks, wb, wSheetName, hdrs);
            wb.SaveAs(wbFileName);
            wb.Close(false);
        }
        private static void ExportResourcesToExcelAndCloseWb(ExcelMap excelMap, Excel.Workbook wb, string wbFileName)
        {
            var wSheetName = "Resource_Table1";
            var hdrs = "Unique_ID,ID,Name,Type,Material_Label,Group_Name,Standard_Rate";
            ExportListToExcel(excelMap.ExcelMapResources, wb, wSheetName, hdrs);
            wb.SaveAs(wbFileName);
            wb.Close(false);
        }
        private static void ExportMaterialAssignmentsToExcelAndCloseWb(ExcelMap excelMap, Excel.Workbook wb, string wbFileName)
        {
            var wSheetName = "Assignment_Table1";
            var hdrs = "Resource_Unique_ID,Task_Unique_ID,Scheduled_Work";
            ExportListToExcel(excelMap.ExcelMapNewProjectMaterialAssignments, wb, wSheetName, hdrs);
            wb.SaveAs(wbFileName);
            wb.Close(false);
        }
        private static void ExportCostAssignmentsToExcelAndCloseWb(ExcelMap excelMap, Excel.Workbook wb, string wbFileName)
        {
            var wSheetName = "Assignment_Table1";
            var hdrs = "Resource_Unique_ID,Task_Unique_ID,Cost";
            ExportListToExcel(excelMap.ExcelMapCostAssignments, wb, wSheetName, hdrs);
            wb.SaveAs(wbFileName);
            wb.Close(false);
        }

        private static void ExportListToExcel<T>(List<T> listToExport, Excel.Workbook wb, string wSheetName,
                                                    string headers, int afterWorksheet = 0, string onlyColumns = "")
        {
            var arrayToExporttoExcel = ProjectHelpers.ReturnListAs2DArray(listToExport, onlyColumns);
            Excel.Worksheet ws = CreateExcelWorksheet(wb, wSheetName, headers, afterWorksheet);
            ExportArrayToExcel(arrayToExporttoExcel, ws);
            Marshal.ReleaseComObject(ws);
        }

        private static Excel.Worksheet CreateExcelWorksheet(Excel.Workbook wb, string wSheetName, string headers, int afterWorksheet)
        {
            Excel.Worksheet ws;
            if (afterWorksheet == 0)
            {
                ws = (Excel.Worksheet)wb.Sheets[1];
            }
            else
            {
                ws = wb.Worksheets.Add(After: wb.Sheets[1]);
            }
            ws.Name = wSheetName;
            CreateHeaders(headers, ws);
            return ws;
        }

        private static void Export2DStringArrayToExcel(string[,] arrayToExporttoExcel, Excel.Workbook wb, string wSheetName,
                                                   string headers, int afterWorksheet = 0)
        {
            Excel.Worksheet ws = CreateExcelWorksheet(wb, wSheetName, headers, afterWorksheet);
            ExportArrayToExcel(arrayToExporttoExcel, ws);
            CreateHeaders(headers, ws);
            Marshal.ReleaseComObject(ws);
        }

        private static void ExportArrayToExcel(string[,] arrayToExporttoExcel, Excel.Worksheet ws)
        {
            Excel.Range rng = ws.Cells.get_Resize(arrayToExporttoExcel.GetLength(0), arrayToExporttoExcel.GetLength(1)).Offset[1];
            rng.Value2 = arrayToExporttoExcel;
        }

        private static void CreateHeaders(string hdrs, Excel.Worksheet ws)
        {
            var subhdr = hdrs.Split(',');
            for (int i = 0; i < subhdr.Length; i++)
            {
                ws.Cells[1, i + 1] = subhdr[i];
            }
        }

        public static DataTable ReturnOrderExcelAsDT(TransferOptions transferOptions)
        {
            var orderExcelDT = ReturnExcelSheetAsDataTable(transferOptions.ExcelOrderFileName, "Order");
            return orderExcelDT;
        }

        public static DataTable ReturnExcelSheetAsDataTable(string ExcelFileName, string ExcelSheetName = "")
        {
            string sSheetName = null;
            var dataTable = new DataTable();

            string sConnection = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={ExcelFileName};Mode=Read;Extended Properties=\"Excel 12.0;HDR=Yes;IMEX=1\"";

            OleDbConnection oleExcelConnection = new OleDbConnection(sConnection);

            oleExcelConnection.Open();

            DataTable dtTablesList = oleExcelConnection.GetSchema("Tables");

            if (dtTablesList.Rows.Count > 0)
            {
                sSheetName = ExcelSheetName != "" ? ExcelSheetName : dtTablesList.Rows[0]["TABLE_NAME"].ToString();
            }

            dtTablesList.Clear();
            dtTablesList.Dispose();

            if (!string.IsNullOrEmpty(sSheetName))
            {
                using (oleExcelConnection)
                {
                    if (!sSheetName.EndsWith("$"))
                        sSheetName += "$";
                    string query = string.Format("SELECT * FROM [{0}]", sSheetName);
                    OleDbDataAdapter adapter = new OleDbDataAdapter(query, oleExcelConnection);
                    adapter.Fill(dataTable);
                }
            }
            return dataTable;
        }

        public static List<string> GetOleDbExcelSheetNames(string ExcelFileName)
        {
            string connectionString = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={ExcelFileName};Mode=Read;Extended Properties=\"Excel 12.0;HDR=Yes;IMEX=1\"";
            OleDbConnection con = new OleDbConnection(connectionString);
            con.Open();
            DataTable dt = con.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);

            if (dt == null)
            {
                return null;
            }

            List<string> sheets = new List<string>();

            foreach (DataRow row in dt.Rows)
                if (row["TABLE_NAME"].ToString().Contains("$"))
                {
                    string s = row["TABLE_NAME"].ToString();
                    if (!s.Contains("Print_Titles"))
                    {
                        sheets.Add(s);
                    }
                }

            con.Close();
            return sheets;
        }

        public static void DoProjectUpdateExcelOperations(ExcelMap newMapTasksWithProgres)
        {
            Excel.Workbook wb = null;
            Excel.Application excelApp = new Excel.Application();

            wb = GetNewExelWorkbook(excelApp);

            DeleteTasksExcelMapFile();
            //Below creates ExcelMap Tasks with progres only
            ExportMapsToExcel(newMapTasksWithProgres, true);

            excelApp.Quit();

            GC.Collect();
            GC.WaitForPendingFinalizers();

            Marshal.ReleaseComObject(wb);
            Marshal.ReleaseComObject(excelApp);
        }
    }
}
