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



namespace FileOrganizer
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {

        public ObservableCollection<string> FileList { get; set; }
        public Dictionary<string, string> DirectoriesDictionary { get; set; }
        public Dictionary<string, string> RenamingDictionary { get; set; }

        public DataTable excelData {  get; set; }

        public MainWindow()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            FileList = new ObservableCollection<string>();
            DirectoriesDictionary = new Dictionary<string, string>();
            RenamingDictionary = new Dictionary<string, string>();
            excelData = new DataTable();

            InitializeComponent();
            this.DataContext = this;

            //READ ALL .TXT FILES
            string appPath = AppContext.BaseDirectory;
            string appPathPrevious = Directory.GetParent(appPath).Parent.FullName;
            string excelPath = Path.Combine(appPathPrevious, @"data.xlsx");

            ReadDataExcel(excelPath);

            //string gateDirectory = Path.Combine(appPathPrevious, @"Gate");
            string gateDirectory = @"C:\Users\Ian\Desktop\Gate";

            using FileSystemWatcher watcher = new FileSystemWatcher(gateDirectory);
            watcher.NotifyFilter = NotifyFilters.Attributes
                                 | NotifyFilters.CreationTime
                                 | NotifyFilters.DirectoryName
                                 | NotifyFilters.FileName
                                 | NotifyFilters.LastAccess
                                 | NotifyFilters.LastWrite
                                 | NotifyFilters.Security
                                 | NotifyFilters.Size;
            watcher.Created += new FileSystemEventHandler(OnChanged);
            watcher.Changed += new FileSystemEventHandler(OnChanged);
            watcher.EnableRaisingEvents = true;

            string[] gateFiles = Directory.GetFiles(gateDirectory);

            string[] args = Environment.GetCommandLineArgs();

            if(args.Length > 0)
            {
                foreach( string arg in args)
                {
                    if(Path.GetExtension(arg) != ".dll")
                    {
                        gateFiles.Append(arg);
                    }
                }
            }

            AddItemstoFileList(gateFiles);

            string FileListText = "";

            TextFileNames.Text = FileListText;
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

        private void AddItemtoFileList(string file)
        {
            FileList.Add(file);

            OrganizeFiles();
        }

        private void OrganizeFiles()
        {

            foreach(string file in FileList)
            {
                int rowIndex = 0;
                string fileName = Path.GetFileNameWithoutExtension(file);
                string fullName = Path.GetFileName(file);
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
                        string? rename = null;

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
            }

            FileList.Clear();
        }
    }
}