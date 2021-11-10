using NLog;
using ProjectApp2.Helpers;
using ProjectApp2.Model;
using ProjectApp2.Model.enums;
using System;
using System.ComponentModel;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;

namespace ProjectApp2.ViewModel
{
    public class TasksViewModel : INotifyPropertyChanged
    {
        private static Logger logger = LogManager.GetCurrentClassLogger();

        #region Private Fields
        private string excelTasksFileName;
        private string excelRateSheetFileName;
        private string excelOrderFileName;
        private string msProjectFileName;
        private int progressPercentage = 0;
        private string progressOutput;
        private bool transferStarted;
        private bool cancelEnabled;
        private string serverName = SQLHelpers.GetDefaultConStringItems().ServerName;
        private string databaseName = SQLHelpers.GetDefaultConStringItems().DatabaseName;
        private string tableName = "PhaseOrders";
        private bool perctgComplExcel = true;
        private bool msProjectExistingFile = true;
        private string percCompleteTableName;
        private string importSuccessStatus;
        CancellationTokenSource cts = new CancellationTokenSource();
        #endregion

        #region Properties 
        public string MsProjectFileName
        {
            get { return msProjectFileName; }
            set
            {
                msProjectFileName = value;
                RaisePropertyChanged("MsProjectFileName");
                TransferCommand.RaiseCanExecuteChanged();
                UpdateProgresCommand.RaiseCanExecuteChanged();
            }
        }
        public string ExcelTasksFileName
        {
            get
            {
                return excelTasksFileName;
            }
            set
            {
                excelTasksFileName = value;
                RaisePropertyChanged("ExcelTasksFileName");
                TransferCommand.RaiseCanExecuteChanged();
                UpdateProgresCommand.RaiseCanExecuteChanged();
            }
        }


        public string ExcelOrderFileName
        {
            get
            {
                return excelOrderFileName;
            }
            set
            {
                excelOrderFileName = value;
                RaisePropertyChanged("ExcelOrderFileName");
            }
        }


        public string ExcelRateSheetFileName
        {
            get
            {
                return excelRateSheetFileName;
            }
            set
            {
                excelRateSheetFileName = value;
                RaisePropertyChanged("ExcelRateSheetFileName");
            }
        }
        public int ProgressPercentage
        {
            get { return progressPercentage; }
            set
            {
                if (progressPercentage != value)
                {
                    progressPercentage = value;
                    RaisePropertyChanged("ProgressPercentage");
                }
            }
        }

        public string ProgressOutput
        {
            get { return progressOutput; }
            set
            {
                if (progressOutput != value)
                {
                    progressOutput = value;
                    RaisePropertyChanged("ProgressOutput");
                }
            }
        }

        public bool TransferStarted
        {
            get
            {
                return transferStarted;
            }
            set
            {
                transferStarted = value;
                RaisePropertyChanged("TransferStarted");
                TransferCommand.RaiseCanExecuteChanged();
                UpdateProgresCommand.RaiseCanExecuteChanged();
            }
        }
        public bool CancelEnabled
        {
            get
            {
                return cancelEnabled;
            }
            set
            {
                cancelEnabled = value;
                RaisePropertyChanged("CancelEnabled");
                CancelTransferCommand.RaiseCanExecuteChanged();
                UpdateProgresCommand.RaiseCanExecuteChanged();
            }
        }

        public string ServerName
        {
            get { return serverName; }
            set
            {
                if (serverName != value)
                {
                    serverName = value;
                    RaisePropertyChanged("ServerName");
                }
            }
        }

        public string DatabaseName
        {
            get { return databaseName; }
            set
            {
                if (databaseName != value)
                {
                    databaseName = value;
                    RaisePropertyChanged("DatabaseName");
                }
            }
        }

        public string TableName
        {
            get { return tableName; }
            set
            {
                if (tableName != value)
                {
                    tableName = value;
                    RaisePropertyChanged("TableName");
                }
            }
        }
        public bool PerctgComplExcel
        {
            get
            {
                return perctgComplExcel;
            }
            set
            {
                perctgComplExcel = value;
                RaisePropertyChanged("PerctgComplExcel");
            }
        }

        public bool MsProjectExistingFile
        {
            get
            {
                return msProjectExistingFile;
            }
            set
            {
                msProjectExistingFile = value;
                RaisePropertyChanged("MsProjectExistingFile");
                TransferCommand.RaiseCanExecuteChanged();
                UpdateProgresCommand.RaiseCanExecuteChanged();
            }
        }

        public string PercCompleteTableName
        {
            get
            {
                return percCompleteTableName;
            }
            set
            {
                percCompleteTableName = value;
                RaisePropertyChanged("PercCompleteTableName");
                TransferCommand.RaiseCanExecuteChanged();
                UpdateProgresCommand.RaiseCanExecuteChanged();
            }
        }

        public string ImportSuccessStatus
        {
            get
            {
                return importSuccessStatus;
            }
            set
            {
                importSuccessStatus = value;
                RaisePropertyChanged("ImportSuccessStatus");
            }
        }

        #endregion

        #region Commands        
        public DelegateCommand TransferCommand { get; set; }
        public DelegateCommand UpdateProgresCommand { get; set; }
        public DelegateCommand CancelTransferCommand { get; set; }
        public ICommand PercCompChangeCommand { get; }
        public ICommand MsPrjFileChangeCommand { get; }
        public ICommand OpenExcelTasksDialogCommand { get; }
        public ICommand OpenExcelSpecsDialogCommand { get; }
        public ICommand OpenExcelRateSheetDialogCommand { get; }
        public ICommand OpenExcelOrderDialogCommand { get; }
        public ICommand OpenMsProjectDialogCommand { get; }
        #endregion

        /// <summary>
        /// Default Constructor
        /// </summary>
        public TasksViewModel()
        {
            TransferCommand = new DelegateCommand(TransferAsync, () => TransferButtonEnabled());
            UpdateProgresCommand = new DelegateCommand(TransferAsync, () => UpdateProgressButtonEnabled());
            CancelTransferCommand = new DelegateCommand(CancelTransfer, () => CancelEnabled);
            PercCompChangeCommand = new RelayCommand(PercCompChange);
            MsPrjFileChangeCommand = new RelayCommand(MsPrjFileChange);
            OpenExcelTasksDialogCommand = new RelayCommand(OpenExcelTasksDialog);
            OpenExcelRateSheetDialogCommand = new RelayCommand(OpenExcelRateSheetDialog);
            OpenExcelOrderDialogCommand = new RelayCommand(OpenExcelOrderDialog);
            OpenMsProjectDialogCommand = new RelayCommand(OpenMsProjectDialog);
        }

        private bool UpdateProgressButtonEnabled()
        {
            return (!string.IsNullOrEmpty(ExcelTasksFileName) || !string.IsNullOrEmpty(PercCompleteTableName)) &&
                    !string.IsNullOrEmpty(ExcelRateSheetFileName) &&
                    MsProjectExistingFile && !string.IsNullOrEmpty(MsProjectFileName) &&
                   TransferStarted == false;
        }

        private bool TransferButtonEnabled()
        {
            return (!string.IsNullOrEmpty(ExcelTasksFileName) || !string.IsNullOrEmpty(PercCompleteTableName)) &&
                   !string.IsNullOrEmpty(ExcelRateSheetFileName) &&
                   !MsProjectExistingFile &&
                   TransferStarted == false;
        }

        /// <summary>
        /// OpenFileDialog to get Ms Project file name
        /// </summary>        
        private void OpenMsProjectDialog()
        {
            //MsProjectFileName = @"C:\Users\RB\Desktop\Progres\ProjectFile-ProgresAdded - WorkFile.mpp";
            string MsProjectFile = DialogHelpers.GetFileName("MsProject Files|*.mpp");
            if (!string.IsNullOrEmpty(MsProjectFile) && !MsProjectFile.Contains(".lnk"))
            {
                MsProjectFileName = MsProjectFile;
            }
        }
        /// <summary>
        /// OpenFileDialog to get Tasks Excel file name
        /// </summary>
        private void OpenExcelTasksDialog()
        {
            //ExcelTasksFileName = "D:\\M_Projects\\Microsoft Project\\Percentage Complete_020821_Filtered.xlsx";
            string ExcelFile = DialogHelpers.GetFileName("Excel Files|*.xls;*.xlsx;*.xlsm");
            if (!string.IsNullOrEmpty(ExcelFile) && !ExcelFile.Contains(".lnk"))
            {
                ExcelTasksFileName = ExcelFile;
            }
        }
        /// <summary>
        /// OpenFileDialog to get Specs Excel file name
        /// </summary>
        private void OpenExcelRateSheetDialog()
        {
            //ExcelRateSheetFileName = "D:\\M_Projects\\Microsoft Project\\BatterseaRateSheet_Sample_020821.xlsx";

            var ExcelFile = DialogHelpers.GetFileName("Excel Files|*.xls;*.xlsx;*.xlsm");
            if (!string.IsNullOrEmpty(ExcelFile) && !ExcelFile.Contains(".lnk"))
            {
                ExcelRateSheetFileName = ExcelFile;
            }
        }
        private void OpenExcelOrderDialog()
        {
            //ExcelOrderFileName = "C:\\Users\\RB\\Desktop\\PercentComplete_Order.xlsm";
            var ExcelFile = DialogHelpers.GetFileName("Excel Files|*.xls;*.xlsx;*.xlsm");
            if (!string.IsNullOrEmpty(ExcelFile) && !ExcelFile.Contains(".lnk"))
            {
                ExcelOrderFileName = ExcelFile;
            }
        }

        private void CancelTransfer(object value)
        {
            cts.Cancel();
            cts.Dispose();
            CancelEnabled = false;
            TransferStarted = false;
        }

        private void PercCompChange()
        {
            if (PerctgComplExcel)
            {
                PercCompleteTableName = string.Empty;
            }
            else
            {
                ExcelTasksFileName = string.Empty;
                PercCompleteTableName = "PercentComplete";
            }
        }

        private void MsPrjFileChange()
        {
            if (!MsProjectExistingFile)
            {
                MsProjectFileName = string.Empty;
            }
        }

        private DatabaseItems DatabaseConnectionCheck()
        {
            DatabaseItems databaseItems = new DatabaseItems()
            {
                ServerName = ServerName,
                DatabaseName = DatabaseName,
                LabourPhaseTableName = TableName,
                PercCompleteTableName = PercCompleteTableName,
            };

            try
            {
                var sqlString = "SELECT phasename, [order] from " + TableName;
                //Try to return data from table to see if there is an error               
                SQLHelpers.GetDTFromSqlString(databaseItems, sqlString);
                return databaseItems;
            }
            catch (Exception)
            {
                MessageBox.Show("Can't initialize database, please Check Database settings!!!", "Server/Database Error");
                return null;
            }
        }

        private async void TransferAsync(object value)
        {
            ProjectTransferType transferType = (ProjectTransferType)value;
            var databaseItems = DatabaseConnectionCheck();
            if (databaseItems == null)
                return;

            ProgressOutput = "Starting Transfer......";
            TransferStarted = true;

            Progress<ProgressReportModel> progress = new Progress<ProgressReportModel>();
            progress.ProgressChanged += ReportProgress;

            TransferOptions TransferOptions = ReturnTransferOptions(databaseItems, progress, transferType);

            await Task.Run(() => ProjectHelpers.InitializeOperation(TransferOptions));
        }

        private TransferOptions ReturnTransferOptions(DatabaseItems databaseItems, IProgress<ProgressReportModel> progress, ProjectTransferType transferType)
        {
            return new TransferOptions
            {
                ExcelTasksFileName = ExcelTasksFileName,
                ExcelOrderFileName = ExcelOrderFileName,
                ExcelRateSheetFileName = ExcelRateSheetFileName,
                MsProjectFileName = MsProjectFileName,
                PercCompType = PerctgComplExcel ? PercCompType.Excel : PercCompType.SQL,
                transferType = transferType,
                DatabaseItems = databaseItems,
                Progress = progress,
                CancellationToken = cts
            };
        }

        /// <summary>
        /// EventHandler to handle "progress.ProgressChanged" event for the main Transfer Async Task
        /// </summary>
        private void ReportProgress(object sender, ProgressReportModel e)
        {
            CancelEnabled = e.CancelEnabled;
            ProgressOutput = e.ProgressStatus;
            ProgressPercentage = e.PercentageComplete;
            ImportSuccessStatus = e.ImportSuccessStatus;
        }

        public event PropertyChangedEventHandler PropertyChanged;

        private void RaisePropertyChanged(string property)
        {
            if (PropertyChanged != null)
            {
                PropertyChanged(this, new PropertyChangedEventArgs(property));
            }
        }
    }
}
