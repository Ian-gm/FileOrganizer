using System.Collections.ObjectModel;
using System.Text;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Collections.Generic;
using ExcelDataReader;
using System.Data;
using System.Windows.Media.Media3D;
using static System.Runtime.InteropServices.JavaScript.JSType;
using Hardcodet.Wpf.TaskbarNotification;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using System.Windows.Media.TextFormatting;
using System.Diagnostics;
using System.ComponentModel;
using System.Windows.Forms;
using FileAway.Properties;
using System.Globalization;
using System;



namespace FileOrganizer
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {
        //public ObservableCollection<string> FileList { get; set; }
        public Dictionary<string, string> DirectoriesDictionary { get; set; }
        public Dictionary<string, string> RenamingDictionary { get; set; }

        public List<string> FileList { get; set; }
        public string TextLines { get; set; }
        public ObservableCollection<Processed> ProcessedList { get; set; }
        public DataTable excelData {  get; set; }

        private System.Threading.Timer? timer1;

        private string gateDirectory;

        private TaskbarIcon? tbi;

        public MainWindow()
        {
            
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            //FileList = new ObservableCollection<string>();
            FileList = new List<string>();
            ProcessedList = new ObservableCollection<Processed>();
            DirectoriesDictionary = new Dictionary<string, string>();
            RenamingDictionary = new Dictionary<string, string>();
            excelData = new DataTable();
            
            DirectoriesDictionary = new Dictionary<string, string>();
            RenamingDictionary = new Dictionary<string, string>();
            excelData = new DataTable();

            InitializeComponent();
            this.DataContext = this;

            //TASKBAR ICON
            tbi = new TaskbarIcon();
            tbi.Icon = System.Drawing.Icon.ExtractAssociatedIcon(System.Reflection.Assembly.GetEntryAssembly().ManifestModule.Name);
            tbi.ToolTipText = "FileAway";

            timer1 = new System.Threading.Timer(Callback, null, 0, 10000);

            var periodicTimer = new PeriodicTimer(TimeSpan.FromSeconds(5));
            
            //READ ALL .TXT FILES
            string appPath = AppContext.BaseDirectory;
            string appPathPrevious = Directory.GetParent(appPath).Parent.FullName;
            string excelPath = Path.Combine(appPathPrevious, @"data.xlsx");
            
            ReadDataExcel(excelPath);

            gateDirectory = FileAway.Properties.Settings.Default.GateFolderPath;
            
            if(gateDirectory != null && Path.Exists(gateDirectory))
            {
                ChosenFolder.Text = "Gate Folder: " + Path.GetFileName(gateDirectory);
                StatusMessage.Text = "Chosen Gate Folder: " + gateDirectory;
            }

            string[] args = Environment.GetCommandLineArgs();
            
            AddItemstoFileList(args);

            checkGateDirectory();

            this.Closing += MainWindow_Closing;
        }

        private void MainWindow_Closing(object? sender, CancelEventArgs e)
        {
            tbi.Dispose();
        }

        public class Processed : INotifyPropertyChanged
        {
            private string time;
            private string name;
            private string preset;
            public string Time
            {
                get { return this.time; }
                set
                {
                    if (this.time != value)
                    {
                        this.time = value;
                        this.NotifyPropertyChanged("Time");
                    }
                }
            }
            public string Name
            {
                get { return this.name; }
                set
                {
                    if (this.name != value)
                    {
                        this.name = value;
                        this.NotifyPropertyChanged("Name");
                    }
                }
            }
            public string Preset
            {
                get { return this.preset; }
                set
                {
                    if (this.preset != value)
                    {
                        this.preset = value;
                        this.NotifyPropertyChanged("Preset");
                    }
                    else
                    {
                        this.preset = "NO MATCH FOUND";
                    }
                }
            }
            public Processed(string fileName, string filePreset)
            {
                time = DateTime.Now.ToShortTimeString();
                name = fileName;
                if(filePreset != null)
                {
                    preset = filePreset;
                }
                else
                {
                    preset = "NO MATCH FOUND";
                }
                
            }

            public event PropertyChangedEventHandler PropertyChanged;

            public void NotifyPropertyChanged(string propName)
            {
                if (this.PropertyChanged != null)
                    this.PropertyChanged(this, new PropertyChangedEventArgs(propName));
            }
        }

        private void Callback(object? state)
        {
           checkGateDirectory();
        }

        private void checkGateDirectory()
        {
            if (gateDirectory != null && Path.Exists(gateDirectory))
            {
                string[] gateFiles = Directory.GetFiles(gateDirectory);
                AddItemstoFileList(gateFiles);
            }
        }

        private static void OnChanged(object sender, FileSystemEventArgs e)
        {
            Console.WriteLine(e.FullPath);
        }

        private void dropfiles(object sender, System.Windows.DragEventArgs e) //Esta es la función que recibe archivos por drag n drop
        {
            string[] droppedFiles = null;

            if (e.Data.GetDataPresent(System.Windows.DataFormats.FileDrop))
            {
                droppedFiles = e.Data.GetData(System.Windows.DataFormats.FileDrop, true) as string[];
            }

            if ((null == droppedFiles) || (!droppedFiles.Any())) { return; }

            AddItemstoFileList(droppedFiles);
        }

        private void ReadDataExcel(string filePath)
        {
            if (!Path.Exists(filePath))
            {
                StatusMessage.Text = "data.xls file doesn't exist";
                return;
            }

            FileStream stream;

            try
            {
                stream = File.Open(filePath, FileMode.Open, FileAccess.Read);
            }
            catch(Exception ex)
            {
                StatusMessage.Text = ex.Message;
                return;
            }
            IExcelDataReader excelReader;

            excelReader = ExcelDataReader.ExcelReaderFactory.CreateReader(stream);

            var conf = new ExcelDataSetConfiguration
            {
                ConfigureDataTable = _ => new ExcelDataTableConfiguration
                {
                    UseHeaderRow = true
                }
            };

            var dataSet = excelReader.AsDataSet(conf);

            excelData = dataSet.Tables[0];

            stream.Dispose();
        }

        private void AddItemstoFileList(string[] files)
        {
            foreach (string s in files)
            {
                if (!Path.GetExtension(s).Contains(".dll"))
                {
                    FileList.Add(s);
                }
            }
            
            OrganizeFiles();
        }

        private void OrganizeFiles()
        {
            foreach(string file in FileList)
            {
                int rowIndex = 0;
                string fileName = Path.GetFileNameWithoutExtension(file);
                string fullName = Path.GetFileName(file);
                string[] fileNamePieces;
                DateTime fileDate = DateTime.Today;
                string? rename = null;
                bool isDate = false;

                string chosenPiece = "";
                string stringDate = "";

                if (fileName.Contains("_"))
                {
                    fileNamePieces = fileName.Split('_');
                    int dateIndex = 0;

                    foreach (string fileNamePiece in fileNamePieces)
                    {
                        //isDate = DateTime.TryParseExact(fileNamePiece, "MMddyyyy", enUS,
                         //     DateTimeStyles.AdjustToUniversalout, fileDate);
                        if ( isDate ) { break; }
                        dateIndex++;
                    }

                    if (isDate)
                    {
                        if (dateIndex == 0)
                        {
                            chosenPiece = fileNamePieces[1];
                        }
                        else
                        {
                            chosenPiece = fileNamePieces[0];
                        }
                    }
                    else
                    {
                        chosenPiece = fileNamePieces[0];
                    }
                }
                else
                {
                    chosenPiece = fileName;
                }


                string finalDate = "";
                if (isDate)
                {
                    stringDate = fileDate.ToShortDateString();
                    string[] datePieces = stringDate.Split('/');
                    finalDate = datePieces[2].Substring(2) + datePieces[1] + datePieces[0];
                }
                else
                {
                    string[] datePieces = DateTime.Today.ToShortDateString().Split('/');
                    finalDate = datePieces[2].Substring(2) + datePieces[1] + datePieces[0];
                }
                

                foreach (DataRow row in excelData.Rows)
                {
                    string Keyword = row["Keyword"].ToString();

                    if (chosenPiece.Equals(Keyword))
                    {
                        string? filePath = null;

                        try
                        {
                            filePath = row["Directory"].ToString();
                        }
                        catch
                        {

                        }

                        try
                        {
                            rename = row["Preset"].ToString();
                        }
                        catch
                        {
                        
                        }

                        string ext = Path.GetExtension(file);
                        if(filePath != null && rename != null)
                        {
                            string newfile = Path.Combine(filePath, rename + ext);

                            try
                            {
                                File.Copy(file, newfile);
                            }
                            catch(Exception e)
                            { 
                            
                            }
                        }
                    }

                    rowIndex++;
                }

                string name = Path.GetFileNameWithoutExtension(file);
                rename = finalDate + "_" + rename;
                Processed mewItem = new Processed(name, rename);
                try
                {
                    System.Windows.Application.Current.Dispatcher.BeginInvoke(new Action(() => this.ProcessedList.Add(mewItem)));
                }
                catch(System.Exception e)
                {
                    StatusMessage.Text = e.ToString();
                }
            }

            ClearGateDirectory();
            FileList.Clear();
        }

        private void ClearGateDirectory()
        {
            if(gateDirectory != null && Path.Exists(gateDirectory))
            {
                string[] gateFiles = Directory.GetFiles(gateDirectory);
                foreach (string gateFile in gateFiles)
                {
                    File.Delete(gateFile);
                }
            }

        }

        private void GateFolder_Click(object sender, RoutedEventArgs e)
        {
            using (var fbd = new FolderBrowserDialog())
            {
                DialogResult result = fbd.ShowDialog();

                string chosenFolder = fbd.SelectedPath;
                
                FileAway.Properties.Settings.Default.GateFolderPath = chosenFolder;
                FileAway.Properties.Settings.Default.Save();

                gateDirectory = FileAway.Properties.Settings.Default.GateFolderPath;

                ChosenFolder.Text = "Gate Folder: " + Path.GetFileName(chosenFolder);
                StatusMessage.Text = "Chosen Gate Folder: " + chosenFolder;
            }
        }

        private void ClearProcessedList_Click(object sender, RoutedEventArgs e)
        {
            ProcessedList.Clear();
        }
    }
}