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
using System.Diagnostics.Tracing;
using Microsoft.VisualBasic.Devices;

namespace FileOrganizer
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    /// 

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

            /*
            string fileExe = System.Reflection.Assembly.GetEntryAssembly().Location;
            tbi = new TaskbarIcon();
            //tbi.Icon = System.Drawing.Icon.ExtractAssociatedIcon(fileExe);
            tbi.ToolTipText = "FileAway";
            */

            

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
                    this.Dispatcher.Invoke(() => { StatusMessage.Text = "Chosen Gate Folder: " + gateDirectory; });
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
            //tbi.Dispose();
            
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
            foreach (string file in FileList)
            {
                bool failedDirectory = false;
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
                    keywordPiece = fileNamePieces[1].Trim();
                }
                else
                {
                    stringDate = fileName;
                    keywordPiece = fileName;
                }


                string finalDate = "";

                //stringDate = stringDate.Replace(" ", string.Empty);
                bool dateFound = ComplexParseDate(stringDate, out finalDate);

                keywordPiece = keywordPiece.ToLower();

                string Keyword = "";

                foreach (DataRow row in excelData.Rows)
                {

                    try
                    {
                        Keyword = row["Keyword"].ToString();
                    }
                    catch
                    {
                        this.Dispatcher.Invoke(() =>
                        {
                            StatusMessage.Text = "Missing Keyword column";
                        });
                    }

                    if (keywordPiece.Contains(Keyword.ToLower().Trim()))
                    {
                        try
                        {
                            filePath = row["Directory"].ToString();
                            string quote = '"'.ToString();
                            filePath = filePath.Replace(quote, string.Empty);
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
                    if (!Path.Exists(filePath))
                    {
                        try
                        {
                            Directory.CreateDirectory(filePath);
                        }
                        catch (Exception e)
                        {
                            failedDirectory = true;
                            /*
                            this.Dispatcher.Invoke(() =>
                            {
                                StatusMessage.Text = "The directory on keyword: '" + Keyword.Trim() + "' is not valid";
                            });
                            */
                        }

                        if (Path.Exists(filePath))
                        {
                            this.Dispatcher.Invoke(() =>
                            {
                                StatusMessage.Text = "Created folder: " + filePath;
                            });
                        }
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
                        this.Dispatcher.Invoke(() =>
                        {
                            StatusMessage.Text = "Couldn't send file to indicated folder. " + e.Message;
                        });
                    }

                    if (originalDirectory == FileAway.Properties.Settings.Default.GateFolderPath && Path.Exists(newfile))
                    {
                        File.Delete(file);
                    }
                }


                string name = Path.GetFileNameWithoutExtension(file);
                if (newfile.Length == 0)
                {
                    newfile = "COULDN'T FIND KEYWORD";
                }
                else if (failedDirectory)
                {
                    newfile = "COULDN'T CREATE DIRECTORY";
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
                        this.Dispatcher.Invoke(() => { StatusMessage.Text = e.ToString(); });
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
                if (char.IsLetter(newChar) || char.IsNumber(newChar)) //SI ES NUMERO O LETRA AGREGAR A WORD
                {
                    words[words.Count - 1] += newChar;
                }

                if (i + 1 < input.Length) //SI NO ESTAMOS AL BORDE DEL STRING
                {
                    char nextChar = input[i + 1];
                    bool newCharbool = false;
                    bool nextCharbool = false;
                    bool prevCharbool = false;


                    //CHECKEAR SI ESTE CHAR Y EL SIGUIENTE SON DE DISTINTO TIPO
                    if (char.IsLetter(newChar))
                    {
                        if (char.IsUpper(newChar))
                        {
                            newCharbool = true;
                            nextCharbool = true;
                        }
                        else if (char.IsLower(newChar))
                        {
                            newCharbool = true;
                            nextCharbool = char.IsLower(nextChar);
                        }
                    }
                    else if (char.IsDigit(newChar))
                    {
                        newCharbool = true;
                        nextCharbool = char.IsDigit(nextChar);
                    }
                    else
                    {
                        newCharbool = false;
                        nextCharbool = char.IsLetter(nextChar) | char.IsDigit(nextChar);
                    }

                    //SI LA WORD TIENE DATA Y LOS CHARS SON DISTINTOS ENTONCES PASAR A LA SIGUIENTE
                    if (words[words.Count - 1].Length != 0 && newCharbool ^ nextCharbool)
                    {
                        words.Add(string.Empty);
                    }

                    //CHECKEAR SI EL CHAR PREVIO NO ERA UN SÍMBOLO Y SI LOS DOS SIGUIENTES SON SIMBOLOS
                    if (i - 1 >= 0)
                    {
                        char prevChar = input[i - 1];
                        prevCharbool = !(char.IsLetter(prevChar) | char.IsDigit(prevChar) | char.IsWhiteSpace(prevChar));
                        newCharbool = !(char.IsLetter(newChar) | char.IsDigit(newChar) | char.IsWhiteSpace(newChar));
                        nextCharbool = !(char.IsLetter(nextChar) | char.IsDigit(nextChar) | char.IsWhiteSpace(nextChar));

                        if (prevCharbool & newCharbool & nextCharbool)
                        {
                            words[words.Count - 1] += "x";
                        }
                    }
                }
            }

            //CREATE A NEW SUBLIST IF THERE'S NON-MONTH BETWEEN ITEMS OR MORE THAN 1 SYMBOL CHAR
            List<List<string>> subwords = new List<List<string>>();
            subwords.Add(new List<string>());

            for (var i = 0; i < words.Count; i++)
            {
                string word = words[i];
                bool wordisnumber = true;
                bool wordisMonth = false;

                string nextWord = "";
                bool nextWordisnumber = true;
                bool nextWordisMonth = false;

                if (i + 1 < words.Count)
                {
                    nextWord = words[i + 1];
                }

                wordisnumber = int.TryParse(word, out _);
                if (!wordisnumber)
                {
                    wordisMonth = DateTime.TryParseExact(word, "MMM", CultureInfo.InvariantCulture, DateTimeStyles.None, out _);
                    wordisMonth = wordisMonth || DateTime.TryParseExact(word, "MMMM", CultureInfo.InvariantCulture, DateTimeStyles.None, out _);
                    if (wordisMonth)
                    {
                        wordisnumber = true;
                    }
                }

                nextWordisnumber = int.TryParse(nextWord, out _);
                if (!nextWordisnumber)
                {
                    nextWordisMonth = DateTime.TryParseExact(nextWord, "MMM", CultureInfo.InvariantCulture, DateTimeStyles.None, out _);
                    nextWordisMonth = nextWordisMonth || DateTime.TryParseExact(nextWord, "MMMM", CultureInfo.InvariantCulture, DateTimeStyles.None, out _);
                    if (nextWordisMonth)
                    {
                        nextWordisnumber = true;
                    }
                }

                //if this word is a number or month, add it
                if (wordisnumber)
                {
                    subwords[subwords.Count - 1].Add(word);
                }

                //if next word isn't a number or month, make new sublist
                if (wordisnumber & !nextWordisnumber)
                {
                    subwords.Add(new List<string>());
                }
            }

            //Clean the last subwords if it's empty
            if (subwords[subwords.Count - 1].Count == 0)
            {
                subwords.RemoveAt(subwords.Count - 1);
            }

            //LEGEND: {YEAR, MONTH, DAY}
            List<int> dateData = new List<int> { -1, -1, -1 };

            foreach (List<string> swords in subwords)
            {
                List<int> newdateData = getDate(swords);

                bool hadYear = dateData[0] != -1;
                bool hadMonth = dateData[1] != -1;
                bool hadDay = dateData[2] != -1;

                int hadSum = Convert.ToInt32(hadYear) + Convert.ToInt32(hadMonth) + Convert.ToInt32(hadDay);

                bool hasYear = newdateData[0] != -1;
                bool hasMonth = newdateData[1] != -1;
                bool hasDay = newdateData[2] != -1;

                int hasSum = Convert.ToInt32(hasYear) + Convert.ToInt32(hasMonth) + Convert.ToInt32(hasDay);

                if (hasSum > hadSum)
                {
                    dateData = newdateData;
                }
                else if (hasSum == hadSum)
                {
                    if (!hadYear && hasYear)
                    {
                        dateData = newdateData;
                    }
                    else
                    {
                        if (dateData[0] < newdateData[0]) //LARGER YEAR
                        {
                            dateData = newdateData;
                        }
                        else if (dateData[0] == newdateData[0])
                        {
                            if (dateData[1] < newdateData[1]) //LARGER MONTH
                            {
                                dateData = newdateData;
                            }
                            else if (dateData[1] == newdateData[1])
                            {
                                if (dateData[2] < newdateData[2]) //LARGER DAY
                                {
                                    dateData = newdateData;
                                }
                            }
                        }
                    }
                }
            }

            if (dateData[0] == -1)
            {
                dateData[0] = DateTime.Now.Year;
            }
            if (dateData[1] == -1)
            {
                dateData[1] = DateTime.Now.Month;
            }
            if (dateData[2] == -1)
            {
                dateData[2] = DateTime.Now.Day;
            }

            dateString = $"{dateData[0]%100:D2}" + "-" + $"{dateData[1]:D2}" + "-" + $"{dateData[2]:D2}";

            return true;
        }

        private List<int> getDate(List<string> input)
        {
            List<int> finalDate = new List<int>{ -1, -1, -1};

            //LEGEND: 0 = year | 1 = month | 2 = day | 3 = longdate

            bool isYear = false;
            bool isMonth = false;
            bool isLongDate = false;

            DateTime getMonth = DateTime.Now;

            List<string> firstParse = new List<string>();

            for (int i = 0; i < input.Count; i++)
            {
                string word = input[i];
                char firstchar = word[0];
                int value = -1;

                if (char.IsLetter(firstchar))
                {
                    isMonth = DateTime.TryParseExact(word, "MMM", CultureInfo.CurrentCulture, DateTimeStyles.None, out getMonth);
                    //isMonth = DateTime.TryParseExact(word, "MMM")
                    if (!isMonth)
                    {
                        isMonth = DateTime.TryParseExact(word, "MMMM", CultureInfo.InvariantCulture, DateTimeStyles.None, out getMonth);
                    }

                    if (isMonth)
                    {
                        finalDate[1] = getMonth.Month;
                    }
                }
                else if (word.Length >= 6)
                {
                    DateTime longDate = DateTime.Now;

                    int wordYear = -1;
                    int wordMonth = -1;
                    int wordDay = -1;

                    if (word.Length == 6)
                    {
                        wordYear = int.Parse(word.Substring(0, 2));
                        wordMonth = int.Parse(word.Substring(2, 2));
                        wordDay = int.Parse(word.Substring(4, 2));
                        if (wordMonth > 12 | wordDay > 31)
                        {
                            wordDay = int.Parse(word.Substring(0, 2));
                            wordMonth = int.Parse(word.Substring(2, 2));
                            wordYear = int.Parse(word.Substring(4, 2));
                        }
                        //isLongDate = DateTime.TryParseExact(word, "yyMMdd", CultureInfo.InvariantCulture, DateTimeStyles.None, out longDate);
                    }
                    else if (word.Length == 8)
                    {
                        wordYear = int.Parse(word.Substring(0, 4));
                        wordMonth = int.Parse(word.Substring(4, 2));
                        wordDay = int.Parse(word.Substring(6, 2));
                        if (wordMonth > 12 | wordDay > 31 | wordYear < 1900)
                        {
                            wordDay = int.Parse(word.Substring(0, 2));
                            wordMonth = int.Parse(word.Substring(2, 2));
                            wordYear = int.Parse(word.Substring(4, 4));
                        }

                        if (wordMonth > 12 | wordDay > 31 | wordYear < 1900)
                        {
                            wordMonth = int.Parse(word.Substring(0, 2));
                            wordDay = int.Parse(word.Substring(2, 2));
                            wordYear = int.Parse(word.Substring(4, 4));
                        }
                    }

                    if (wordYear != -1 && wordMonth != -1 && wordDay != -1)
                    {
                        if (wordMonth <= 12 && wordDay <= 31)
                        {
                            isLongDate = true;

                            finalDate[0] = wordYear;
                            finalDate[1] = wordMonth;
                            finalDate[2] = wordDay;

                            //return finalDate;
                        }
                    }

                }
                else if (word.Length == 4)
                {
                    isYear = true;
                    finalDate[0] = int.Parse(word);
                }
                else if(int.TryParse(word, out value))
                {
                    if (value > 31)
                    {
                        if (!isYear)
                        {
                            isYear = true;
                            finalDate[0] = value;
                        }
                        else
                        {
                            firstParse.Add(word);
                        }
                    }
                    else if (value > 12)
                    {
                        if (isYear)
                        {
                            finalDate[2] = value;
                        }
                        else
                        {
                            firstParse.Add(word);
                        }
                    }
                    else if (isYear && isMonth)
                    {
                        finalDate[2] = value;
                    }
                    else
                    {
                        firstParse.Add(word);
                    }
                }
            }

            for (int i = 0; i < firstParse.Count; i++)
            {
                string word = firstParse[i];
                int value = int.Parse(word);

                if (value > 31)
                {
                    if (!isYear)
                    {
                        isYear = true;
                        finalDate[0] = value;
                    }
                }
                else if (value > 12)
                {
                    if (isYear)
                    {
                        if (finalDate[2] == -1) { finalDate[2] = value; }
                    }
                }
                else
                {
                    if (isYear && isMonth)
                    {
                        if (finalDate[2] == -1) { finalDate[2] = value; }
                    }
                    else if (isYear)
                    {
                        if (finalDate[1] == -1)
                        { 
                            isMonth = true;
                            finalDate[1] = value;
                        }
                    }
                }
                
            }

            if (finalDate[0] != -1 && finalDate[1] != -1 && finalDate[2] != -1)
            {
                return finalDate;
            }
            
            for (int i = 0; i < input.Count; i++)
            {
                string word = input[i];
                int value = -1;

                bool isNumber = int.TryParse(word, out value);

                if (isNumber)
                {
                    if (input.Count == 3)
                    {
                        if (!isYear && i == 0)
                        {
                            if (finalDate[0] == -1) { finalDate[0] = value; }
                        }
                        if (!isMonth && i == 1)
                        {
                            if (finalDate[1] == -1) { finalDate[1] = value; }
                        }
                        if (i == 2)
                        {
                            if (finalDate[2] == -1) { finalDate[2] = value; }
                        }
                    }
                    else if (input.Count == 2)
                    {
                        if (!isMonth && i == 0)
                        {
                            if (finalDate[1] == -1) { finalDate[1] = value; }
                        }
                        else if (isMonth && i == 0)
                        {
                            if (finalDate[2] == -1) { finalDate[2] = value; }
                        }
                        else if (i == 1)
                        {
                            if (finalDate[2] == -1) { finalDate[2] = value; }
                        }
                    }
                }
            }
            

            return finalDate;
        }

        private List<int> gDate(List<string> input)
        {
            List<int> finalDate = new List<int> { -1, -1, -1 };

            //LEGEND: 0 = year | 1 = month | 2 = day | 3 = longdate

            bool isYear = false;
            bool isMonth = false;
            bool isLongDate = false;

            DateTime getMonth = DateTime.Now;

            List<string> firstParse = new List<string>();

            for (int i = 0; i < input.Count; i++)
            {
                string word = input[0];
                char firstchar = word[0];
                int value = -1;

                if (char.IsLetter(firstchar))
                {
                    isMonth = DateTime.TryParseExact(word, "MMM", CultureInfo.CurrentCulture, DateTimeStyles.None, out getMonth);
                    //isMonth = DateTime.TryParseExact(word, "MMM")
                    if (!isMonth)
                    {
                        isMonth = DateTime.TryParseExact(word, "MMMM", CultureInfo.InvariantCulture, DateTimeStyles.None, out getMonth);
                    }

                    if (isMonth)
                    {
                        finalDate[1] = getMonth.Month;

                        //GET THE REST VALUES
                        if(i == 0)
                        {
                            if(input.Count == 2)
                            {
                               
                            }
                        }
                    }
                }
                else if (word.Length >= 6)
                {
                    DateTime longDate = DateTime.Now;

                    int wordYear = -1;
                    int wordMonth = -1;
                    int wordDay = -1;

                    if (word.Length == 6)
                    {
                        wordYear = int.Parse(word.Substring(0, 2));
                        wordMonth = int.Parse(word.Substring(2, 2));
                        wordDay = int.Parse(word.Substring(4, 2));
                        if (wordMonth > 12 | wordDay > 31)
                        {
                            wordDay = int.Parse(word.Substring(0, 2));
                            wordMonth = int.Parse(word.Substring(2, 2));
                            wordYear = int.Parse(word.Substring(4, 2));
                        }
                        //isLongDate = DateTime.TryParseExact(word, "yyMMdd", CultureInfo.InvariantCulture, DateTimeStyles.None, out longDate);
                    }
                    else if (word.Length == 8)
                    {
                        wordYear = int.Parse(word.Substring(0, 4));
                        wordMonth = int.Parse(word.Substring(4, 2));
                        wordDay = int.Parse(word.Substring(6, 2));
                        if (wordMonth > 12 | wordDay > 31 | wordYear < 1900)
                        {
                            wordDay = int.Parse(word.Substring(0, 2));
                            wordMonth = int.Parse(word.Substring(2, 2));
                            wordYear = int.Parse(word.Substring(4, 4));
                        }

                        if (wordMonth > 12 | wordDay > 31 | wordYear < 1900)
                        {
                            wordMonth = int.Parse(word.Substring(0, 2));
                            wordDay = int.Parse(word.Substring(2, 2));
                            wordYear = int.Parse(word.Substring(4, 4));
                        }
                    }

                    if (wordYear != -1 && wordMonth != -1 && wordDay != -1)
                    {
                        if (wordMonth <= 12 && wordDay <= 31)
                        {
                            isLongDate = true;

                            finalDate[0] = wordYear;
                            finalDate[1] = wordMonth;
                            finalDate[2] = wordDay;

                            return finalDate;
                        }
                    }

                }
                else if (word.Length == 4)
                {
                    isYear = true;
                    finalDate[0] = int.Parse(word);
                }
            }

            //NO OBVIOUS MATCHES
            if(input.Count == 0)
            {
                int Year = -1;
                int Month = -1;
                int Day = -1;

                Year = int.Parse(input[0]);
                Month = int.Parse(input[1]);
                Day = int.Parse(input[2]);

                if (Month > 12 | Day > 31)
                {
                    Year = int.Parse(input[2]);
                    Month = int.Parse(input[1]);
                    Day = int.Parse(input[0]);
                }

                if (Month > 12 | Day > 31)
                {
                    Year = int.Parse(input[2]);
                    Month = int.Parse(input[0]);
                    Day = int.Parse(input[1]);
                }
            }




            return finalDate;
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
                this.Dispatcher.Invoke(() => { StatusMessage.Text = "Chosen Gate Folder: " + chosenFolder; });
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
                this.Dispatcher.Invoke(() => { StatusMessage.Text = "Opened: " + folderPath; });
            }
            else
            {
                StatusMessage.Text = "That folder doesn't exist";
            }
        }
    }
}