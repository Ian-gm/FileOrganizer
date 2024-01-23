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
        public Dictionary<string, string> DirectoriesDictionary { get; set; }
        public Dictionary<string, string> RenamingDictionary { get; set; }
        public DataTable excelData {  get; set; }

        private Timer? timer1;

        private string gateDirectory;

        private TaskbarIcon? tb;

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

            timer1 = new Timer(Callback, null, 0, 10000);

            var periodicTimer = new PeriodicTimer(TimeSpan.FromSeconds(5));

            //READ ALL .TXT FILES
            string appPath = AppContext.BaseDirectory;
            string appPathPrevious = Directory.GetParent(appPath).Parent.FullName;
            string excelPath = Path.Combine(appPathPrevious, @"data.xlsx");
            string icoPath = Path.Combine(appPath, @"icons\Logo.ico");

            //tb = (TaskbarIcon)FindResource("MyNotifyIcon");
            //tb.Icon = new System.Drawing.Icon(icoPath);
            TaskBar.Icon = new System.Drawing.Icon(icoPath);

            ReadDataExcel(excelPath);

            gateDirectory = Path.Combine(appPathPrevious, @"Gate");
            //gateDirectory = @"C:\Users\Ian\Desktop\Gate";

            string[] args = Environment.GetCommandLineArgs();

            /*TaskbarIcon tbi = new TaskbarIcon();
            tbi.Icon = Resources
            tbi.ToolTipText = "hello world";
            */
            AddItemstoFileList(args);

            checkGateDirectory();
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
            if (gateDirectory != null)
            {
                string[] gateFiles = Directory.GetFiles(gateDirectory);
                AddItemstoFileList(gateFiles);
            }
        }

        private static void OnChanged(object sender, FileSystemEventArgs e)
        {
            Console.WriteLine(e.FullPath);
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
                return;
            }

            FileStream stream = File.Open(filePath, FileMode.Open, FileAccess.Read);
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
        }

        private void AddItemstoFileList(string[] files)
        {
            foreach (string s in files)
            {
                if (!Path.GetFileName(s).EndsWith(".dll"))
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
                string? rename = null;

                if (fileName.Contains("_"))
                {
                    fileName = fileName.Split('_')[0];
                }
                
                foreach (DataRow row in excelData.Rows)
                {
                    string Keyword = row["Keyword"].ToString();

                    if (fileName.Equals(Keyword))
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
                Processed mewItem = new Processed(name, rename);
                Application.Current.Dispatcher.BeginInvoke(new Action(() => this.ProcessedList.Add(mewItem)));
            }

            ClearGateDirectory();
            FileList.Clear();
        }

        private void ClearGateDirectory()
        {
            string[] gateFiles = Directory.GetFiles(gateDirectory);
            foreach (string gateFile in gateFiles)
            {
                File.Delete(gateFile);
            }
        }
    }
}