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
using static System.Windows.Forms.AxHost;
using System.Xml.Linq;
using System.Linq.Expressions;
using System.Text.RegularExpressions;



namespace FileOrganizer
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {
        public List<string> FileList { get; set; }

        public ObservableCollection<Processed> ProcessedList { get; set; }

        public DataTable excelData {  get; set; }

        private System.Threading.Timer? timer1;

        private string gateDirectory;

        private TaskbarIcon? tbi;

        private bool Running = false;

        public MainWindow()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            FileList = new List<string>();
            ProcessedList = new ObservableCollection<Processed>();
            excelData = new DataTable();

            InitializeComponent();
            this.DataContext = this;

            
            //TASKBAR ICON
            string fileExe = System.Reflection.Assembly.GetEntryAssembly().Location;
            tbi = new TaskbarIcon();
            //tbi.Icon = System.Drawing.Icon.ExtractAssociatedIcon(fileExe);
            tbi.ToolTipText = "FileAway";
            

            //READ ALL .TXT FILES
            string appPath = AppContext.BaseDirectory;
            string appPathPrevious = Directory.GetParent(appPath).Parent.FullName;
            string excelPath = Path.Combine(appPathPrevious, @"data.xlsx");
            
            bool excelRead = ReadDataExcel(excelPath);

            if(excelRead)
            {
                timer1 = new System.Threading.Timer(Callback, null, 0, 2000);
                //var periodicTimer = new PeriodicTimer(TimeSpan.FromSeconds(5));

                gateDirectory = FileAway.Properties.Settings.Default.GateFolderPath;

                if (gateDirectory != null && Path.Exists(gateDirectory))
                {
                    ChosenFolder.Text = "Gate Folder: " + Path.GetFileName(gateDirectory);
                    StatusMessage.Text = "Chosen Gate Folder: " + gateDirectory;
                }

                /*
                string[] args = Environment.GetCommandLineArgs();
                AddItemstoFileList(args);
                */

                //checkGateDirectory();
            }

            this.Closing += MainWindow_Closing;
        }

        private void MainWindow_Closing(object? sender, CancelEventArgs e)
        {
            tbi.Dispose();
        }

        private static void OnChanged(object sender, FileSystemEventArgs e)
        {
            Console.WriteLine(e.FullPath);
        }

        public class Processed : INotifyPropertyChanged
        {
            private string time;
            private string name;
            private string preset;
            private string folder;
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
            public string Folder
            {
                get { return this.folder; }
                set
                {
                    if (this.folder != value)
                    {
                        this.folder = value;
                        this.NotifyPropertyChanged("Folder");
                    }
                    else
                    {
                        this.folder = "NO FOLDER";
                    }
                }
            }

            public Processed(string fileName, string filePreset, string fileFolder)
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

                if (fileFolder != null)
                {
                    folder = fileFolder;
                }
                else
                {
                    folder = "-";
                }
            }

            public event PropertyChangedEventHandler PropertyChanged;

            public void NotifyPropertyChanged(string propName)
            {
                if (this.PropertyChanged != null)
                    this.PropertyChanged(this, new PropertyChangedEventArgs(propName));
            }
        }

        private bool ReadDataExcel(string filePath)
        {
            if (!Path.Exists(filePath))
            {
                StatusMessage.Text = "data.xls file doesn't exist. Please add it";
                return false;
            }

            FileStream stream;

            try
            {
                stream = File.Open(filePath, FileMode.Open, FileAccess.Read);
            }
            catch (Exception ex)
            {
                StatusMessage.Text = ex.Message;
                return false;
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


            if (excelData != null)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        private void Callback(object? state)
        {
            if (!Running)
            {
                Running = true;
                checkGateDirectory();
            }
        }

        private void dropfiles(object sender, System.Windows.DragEventArgs e) //Esta es la función que recibe archivos por drag n drop
        {
            string[] droppedFiles = null;

            if (e.Data.GetDataPresent(System.Windows.DataFormats.FileDrop))
            {
                droppedFiles = e.Data.GetData(System.Windows.DataFormats.FileDrop, true) as string[];
            }

            if ((null == droppedFiles) || (!droppedFiles.Any())) { return; }

            Running = true;

            AddItemstoFileList(droppedFiles);
        }

        private void checkGateDirectory()
        {
            if (gateDirectory != null && Path.Exists(gateDirectory))
            {
                string[] gateFiles = Directory.GetFiles(gateDirectory);
                if(gateFiles.Length > 0)
                {
                    AddItemstoFileList(gateFiles);
                }
            }
            
            Running = false;
        }
        
        private void AddItemstoFileList(string[] files)
        {
            foreach (string s in files)
            {
                 FileList.Add(s);
            }
            OrganizeFiles();

            Running = false;
        }
        
        private void OrganizeFiles()
        {
            foreach(string file in FileList)
            {
                string newfile = "";
                int rowIndex = 0;
                string fileName = Path.GetFileNameWithoutExtension(file);
                string fullName = Path.GetFileName(file);
                string[] fileNamePieces;
                DateTime fileDate = DateTime.Today;
                string? rename = null;
                bool isDate = false;

                string keywordPiece = "";
                string stringDate = "";

                string? filePath = null;

                if (fileName.Contains("$"))
                {
                    fileNamePieces = fileName.Split('$');
                    int dateIndex = 0;

                    stringDate = fileNamePieces[0];
                    keywordPiece = fileNamePieces[1];

                    //ALL DATE VARIATIONS
                    /*
                    while (!isDate)
                    {
                        isDate = CustomParseDate("dddd-MMMM-yyyy", fileNamePiece, out fileDate); if (isDate) { break; }
                        isDate = CustomParseDate("ddd-MMMM-yyyy", fileNamePiece, out fileDate); if (isDate) { break; }
                        isDate = CustomParseDate("dd-MMMM-yyyy", fileNamePiece, out fileDate); if (isDate) { break; }
                        isDate = CustomParseDate("d-MMMM-yyyy", fileNamePiece, out fileDate); if (isDate) { break; }
                        isDate = CustomParseDate("dddd-MMM-yyyy", fileNamePiece, out fileDate); if (isDate) { break; }
                        isDate = CustomParseDate("ddd-MMM-yyyy", fileNamePiece, out fileDate); if (isDate) { break; }
                        isDate = CustomParseDate("dd-MMM-yyyy", fileNamePiece, out fileDate); if (isDate) { break; }
                        isDate = CustomParseDate("d-MMM-yyyy", fileNamePiece, out fileDate); if (isDate) { break; }
                        isDate = CustomParseDate("dddd-MM-yyyy", fileNamePiece, out fileDate); if (isDate) { break; }
                        isDate = CustomParseDate("ddd-MM-yyyy", fileNamePiece, out fileDate); if (isDate) { break; }

                        isDate = CustomParseDate("dddd-M-yyyy", fileNamePiece, out fileDate); if (isDate) { break; }
                        isDate = CustomParseDate("ddd-M-yyyy", fileNamePiece, out fileDate); if (isDate) { break; }

                        isDate = CustomParseDate("dd-MM-yyyy", fileNamePiece, out fileDate); if (isDate) { break; }
                        isDate = CustomParseDate("d-MM-yyyy", fileNamePiece, out fileDate); if (isDate) { break; }
                        isDate = CustomParseDate("dd-M-yyyy", fileNamePiece, out fileDate); if (isDate) { break; }
                        isDate = CustomParseDate("d-M-yyyy", fileNamePiece, out fileDate); if (isDate) { break; }

                        isDate = CustomParseDate("dddd-MMMM-yy", fileNamePiece, out fileDate); if (isDate) { break; }
                        isDate = CustomParseDate("ddd-MMMM-yy", fileNamePiece, out fileDate); if (isDate) { break; }
                        isDate = CustomParseDate("dd-MMMM-yy", fileNamePiece, out fileDate); if (isDate) { break; }
                        isDate = CustomParseDate("d-MMMM-yy", fileNamePiece, out fileDate); if (isDate) { break; }
                        isDate = CustomParseDate("dddd-MMM-yy", fileNamePiece, out fileDate); if (isDate) { break; }
                        isDate = CustomParseDate("ddd-MMM-yy", fileNamePiece, out fileDate); if (isDate) { break; }
                        isDate = CustomParseDate("dd-MMM-yy", fileNamePiece, out fileDate); if (isDate) { break; }
                        isDate = CustomParseDate("d-MMM-yy", fileNamePiece, out fileDate); if (isDate) { break; }
                        isDate = CustomParseDate("dddd-MM-yy", fileNamePiece, out fileDate); if (isDate) { break; }
                        isDate = CustomParseDate("ddd-MM-yy", fileNamePiece, out fileDate); if (isDate) { break; }
                        isDate = CustomParseDate("dddd-M-yy", fileNamePiece, out fileDate); if (isDate) { break; }
                        isDate = CustomParseDate("ddd-M-yy", fileNamePiece, out fileDate); if (isDate) { break; }

                        isDate = CustomParseDate("dd-MM-yy", fileNamePiece, out fileDate); if (isDate) { break; }

                        isDate = CustomParseDate("MMMM-yy", fileNamePiece, out fileDate); if (isDate) { break; }
                        isDate = CustomParseDate("MMMM-yyy", fileNamePiece, out fileDate); if (isDate) { break; }
                        isDate = CustomParseDate("MMMM-yyyy", fileNamePiece, out fileDate); if (isDate) { break; }
                        isDate = CustomParseDate("MMM-yy", fileNamePiece, out fileDate); if (isDate) { break; }
                        isDate = CustomParseDate("MMM-yyy", fileNamePiece, out fileDate); if (isDate) { break; }
                        isDate = CustomParseDate("MMM-yyyy", fileNamePiece, out fileDate); if (isDate) { break; }
                        isDate = CustomParseDate("MM-yyyy", fileNamePiece, out fileDate); if (isDate) { break; }
                        isDate = CustomParseDate("M-yyyy", fileNamePiece, out fileDate); if (isDate) { break; }

                        isDate = CustomParseDate("yyyyMMdd", fileNamePiece, out fileDate); if (isDate) { break; }
                        isDate = CustomParseDate("yyMMdd", fileNamePiece, out fileDate); if (isDate) { break; }
                        break;
                    }
                    */
                    
                }
                else
                {
                    stringDate = fileName;
                }


                string finalDate = "";

                bool dateFound = ComplexParseDate(stringDate, out finalDate);

                keywordPiece = keywordPiece.ToLower();

                foreach (DataRow row in excelData.Rows)
                {

                    string Keyword = "";

                    try
                    {
                        Keyword = row["Keyword"].ToString().ToLower();
                    }
                    catch
                    {
                        StatusMessage.Text = "Missing Keyword column"; 
                    }

                    if (keywordPiece.Equals(Keyword))
                    {
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

                        break;
                    }
                    rowIndex++;
                }

                string ext = Path.GetExtension(file);
                if (filePath != null && rename != null)
                {
                    if(!Path.Exists(filePath))
                    {
                        Directory.CreateDirectory(filePath);

                        StatusMessage.Text = "Created folder: " + filePath;
                    }
                    
                    newfile = Path.Combine(filePath, finalDate + "_" + rename + ext);
                    newfile = addPrefix(newfile);
                    string originalDirectory = Directory.GetParent(file).ToString();

                    try
                    {
                        File.Copy(file, newfile);
                    }
                    catch (Exception e)
                    {
                        StatusMessage.Text = "Couldn't send file to indicated folder. " + e.Message;
                    }

                    if(originalDirectory == FileAway.Properties.Settings.Default.GateFolderPath && Path.Exists(newfile))
                    {
                        File.Delete(file);
                    }
                }


                string name = Path.GetFileNameWithoutExtension(file);
                if(newfile.Length == 0)
                {
                    newfile = "NO MATCH";
                }
                string newname = Path.GetFileNameWithoutExtension(newfile);
                Processed mewItem = new Processed(name, newname, filePath);
                bool alreadyAdded = false;

                foreach (Processed item in ProcessedList)
                {
                    if (item.Name.Equals(name) && item.Preset.Equals(newname))
                    {
                        alreadyAdded = true;
                    }
                }

                if (!alreadyAdded)
                {
                    try
                    {
                        System.Windows.Application.Current.Dispatcher.BeginInvoke(new Action(() => this.ProcessedList.Add(mewItem)));
                    }
                    catch (System.Exception e)
                    {
                        StatusMessage.Text = e.ToString();
                    }
                }
            }

            FileList.Clear();
            Running = false;
        }

        private string addPrefix(string filePath)
        {
            string name = System.IO.Path.GetFileNameWithoutExtension(filePath);
            string docpath = Directory.GetParent(filePath).FullName;
            string ext = Path.GetExtension(filePath);

            if(System.IO.Path.Exists(docpath))
            {
                string[] getFiles = Directory.GetFiles(docpath);
                int largestPrefix = 0;

                foreach (string file in getFiles)
                {
                    int filePrefix = 0;
                    string checkName = name;
                    string fileName = System.IO.Path.GetFileNameWithoutExtension(file);
                    bool flag = fileName.Contains(checkName);

                    if (flag)
                    {
                        filePrefix = 1;
                        if (filePrefix > largestPrefix)
                        {
                            largestPrefix = filePrefix;
                        }

                        string[] fileNamePieces = fileName.Split('-');
                        if (fileNamePieces.Length > 1)
                        {
                            string possiblePrefix = fileNamePieces[fileNamePieces.Length - 1];

                            flag = int.TryParse(possiblePrefix, out filePrefix);
                        }
                        if (filePrefix >= largestPrefix)
                        {
                            largestPrefix = filePrefix + 1;
                        }
                    }
                }
                string prefix = "";

                if (largestPrefix > 0)
                {
                    prefix = "-" + largestPrefix.ToString();
                }

                string filename = name + prefix + ext;
                string finalPath = System.IO.Path.Combine(docpath, filename);

                return finalPath;
            }
            else
            {
                return filePath;
            }
        }
        
        private bool ComplexParseDate(string input, out string dateString)
        {
            bool isDate;
            bool dateFound = false;
            DateTime fileDate = DateTime.Now;

            int yearNumber = -1;
            int monthNumber = -1;
            int dayNumber = -1;

            DateTime foundYear;
            DateTime foundMonth = DateTime.Now;
            DateTime foundDay;

            bool isYear = false;
            bool isMonth = false;
            bool isDay = false;

            //SEPARATE WORDS AND NUMBERS
            var words = new List<string> { string.Empty };
            for (var i = 0; i < input.Length; i++)
            {
                char newChar = input[i];
                if(char.IsLetter(newChar) || char.IsNumber(newChar))
                {
                    words[words.Count - 1] += newChar;
                }
                if (i + 1 < input.Length && char.IsLetter(input[i]) != char.IsLetter(input[i + 1]))
                {
                    words.Add(string.Empty);
                }
            }
            
            foreach(string word in words)
            {
                int value = 0;


                if (!int.TryParse(word, out value))
                {
                    isMonth = DateTime.TryParseExact(word, "MMM", CultureInfo.InvariantCulture, DateTimeStyles.None, out foundMonth);
                    isMonth = DateTime.TryParseExact(word, "MMMM", CultureInfo.InvariantCulture, DateTimeStyles.None, out foundMonth);
                }
                else
                {
                    if(word.Length == 1)
                    {
                        dayNumber = value;
                    }
                    else if(value > 31 || word.Length >= 3 || value == 0)
                    {
                        yearNumber = value;
                    }
                    else if(yearNumber != 0)
                    {
                        dayNumber = value;
                    }
                }

                if (isMonth)
                {
                    monthNumber = foundMonth.Month;
                }
            }


            //CHECK WHAT'S MISSING
            if(yearNumber < 0 && monthNumber < 0 && dayNumber < 0)
            {
                dateFound = false;
            }
            else
            {
                dateFound = true;
            }

            if(yearNumber < 0)
            {
                yearNumber = DateTime.Now.Year;
            }

            if (monthNumber < 0)
            {
                monthNumber = DateTime.Now.Month;
            }

            if (dayNumber < 0)
            {
                dayNumber = DateTime.Now.Day;
            }

            dateString = $"{yearNumber:D2}" + "-" + $"{monthNumber:D2}" + "-" + $"{dayNumber:D2}";

            return dateFound;
        }

        private bool CustomParseDate(string style, string datestring, out DateTime fileDate)
        {
            bool isDate;
            fileDate = DateTime.Now;
            string[] stylePiece = style.Split('-');

            if(stylePiece.Length == 2)
            {
                //style = stylePiece[0] + stylePiece[1];
                isDate = DateTime.TryParseExact(datestring, style, CultureInfo.InvariantCulture, DateTimeStyles.None, out fileDate); if (isDate) { return true; }
                
                style = stylePiece[0] + " " + stylePiece[1];
                isDate = DateTime.TryParseExact(datestring, style, CultureInfo.InvariantCulture, DateTimeStyles.None, out fileDate); if (isDate) { return true; }

                style = stylePiece[0] + "-" + stylePiece[1];
                isDate = DateTime.TryParseExact(datestring, style, CultureInfo.InvariantCulture, DateTimeStyles.None, out fileDate); if (isDate) { return true; }

                style = stylePiece[0] + "." + stylePiece[1];
                isDate = DateTime.TryParseExact(datestring, style, CultureInfo.InvariantCulture, DateTimeStyles.None, out fileDate); if (isDate) { return true; }


                //style = stylePiece[1] + stylePiece[0];
                isDate = DateTime.TryParseExact(datestring, style, CultureInfo.InvariantCulture, DateTimeStyles.None, out fileDate); if (isDate) { return true; }

                style = stylePiece[1] + " " + stylePiece[0];
                isDate = DateTime.TryParseExact(datestring, style, CultureInfo.InvariantCulture, DateTimeStyles.None, out fileDate); if (isDate) { return true; }

                style = stylePiece[1] + "-" + stylePiece[0];
                isDate = DateTime.TryParseExact(datestring, style, CultureInfo.InvariantCulture, DateTimeStyles.None, out fileDate); if (isDate) { return true; }

                style = stylePiece[1] + "." + stylePiece[0];
                isDate = DateTime.TryParseExact(datestring, style, CultureInfo.InvariantCulture, DateTimeStyles.None, out fileDate); if (isDate) { return true; }

            }
            else if (stylePiece.Length == 3)
            {
                //style = stylePiece[0] + stylePiece[1] + stylePiece[2];
                isDate = DateTime.TryParseExact(datestring, style, CultureInfo.InvariantCulture, DateTimeStyles.None, out fileDate); if (isDate) { return true; }

                style = stylePiece[0] + " " + stylePiece[1] + " " + stylePiece[2];
                isDate = DateTime.TryParseExact(datestring, style, CultureInfo.InvariantCulture, DateTimeStyles.None, out fileDate); if (isDate) { return true; }

                style = stylePiece[0] + "-" + stylePiece[1] + "-" + stylePiece[2];
                isDate = DateTime.TryParseExact(datestring, style, CultureInfo.InvariantCulture, DateTimeStyles.None, out fileDate); if (isDate) { return true; }

                style = stylePiece[0] + "." + stylePiece[1] + "." + stylePiece[2];
                isDate = DateTime.TryParseExact(datestring, style, CultureInfo.InvariantCulture, DateTimeStyles.None, out fileDate); if (isDate) { return true; }


                //style = stylePiece[1] + stylePiece[0] + stylePiece[2];
                isDate = DateTime.TryParseExact(datestring, style, CultureInfo.InvariantCulture, DateTimeStyles.None, out fileDate); if (isDate) { return true; }

                style = stylePiece[1] + " " + stylePiece[0] + " " + stylePiece[2];
                isDate = DateTime.TryParseExact(datestring, style, CultureInfo.InvariantCulture, DateTimeStyles.None, out fileDate); if (isDate) { return true; }

                style = stylePiece[1] + "-" + stylePiece[0] + "-" + stylePiece[2];
                isDate = DateTime.TryParseExact(datestring, style, CultureInfo.InvariantCulture, DateTimeStyles.None, out fileDate); if (isDate) { return true; }

                style = stylePiece[1] + "." + stylePiece[0] + "." + stylePiece[2];
                isDate = DateTime.TryParseExact(datestring, style, CultureInfo.InvariantCulture, DateTimeStyles.None, out fileDate); if (isDate) { return true; }


                //style = stylePiece[2] + stylePiece[1] + stylePiece[0];
                isDate = DateTime.TryParseExact(datestring, style, CultureInfo.InvariantCulture, DateTimeStyles.None, out fileDate); if (isDate) { return true; }

                style = stylePiece[2] + " " + stylePiece[1] + " " + stylePiece[0];
                isDate = DateTime.TryParseExact(datestring, style, CultureInfo.InvariantCulture, DateTimeStyles.None, out fileDate); if (isDate) { return true; }

                style = stylePiece[2] + "-" + stylePiece[1] + "-" + stylePiece[0];
                isDate = DateTime.TryParseExact(datestring, style, CultureInfo.InvariantCulture, DateTimeStyles.None, out fileDate); if (isDate) { return true; }

                style = stylePiece[2] + "." + stylePiece[1] + "." + stylePiece[0];
                isDate = DateTime.TryParseExact(datestring, style, CultureInfo.InvariantCulture, DateTimeStyles.None, out fileDate); if (isDate) { return true; }
            }
            else
            {
                isDate = DateTime.TryParseExact(datestring, style, CultureInfo.InvariantCulture, DateTimeStyles.None, out fileDate); if (isDate) { return true; }
            }


            return false;
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

        private void FolderButton_Click(object sender, RoutedEventArgs e)
        {
            string folderPath = "";

            var rowItem = (sender as System.Windows.Controls.Button).DataContext as Processed;

            folderPath = rowItem.Folder;


            if (Directory.Exists(folderPath))
            {
                Process.Start("explorer.exe", string.Format("/select, \"{0}\"", folderPath));
                StatusMessage.Text = "Opened: " + folderPath;
            }
            else
            {
                StatusMessage.Text = "That folder doesn't exist";
            }
        }
    }
}