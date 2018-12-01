//------------------------------------------------------------------------------
// <copyright file="MainWindow.xaml.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
//------------------------------------------------------------------------------

namespace Microsoft.Samples.Kinect.BodyBasics
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel;
    using System.Diagnostics;
    using System.Globalization;
    using System.IO;
    using System.Windows;
    using System.Windows.Media;
    using System.Windows.Media.Imaging;
    using Microsoft.Win32;
    using System.Threading;
    using Microsoft.Kinect;
    using Microsoft.Kinect.Tools;
    using Excel = Microsoft.Office.Interop.Excel;
    using System.Reflection;
    
    /// <summary>
    /// Interaction logic for MainWindow
    /// </summary>
    public partial class MainWindow : Window, INotifyPropertyChanged
    {
        /// <summary> Indicates if a playback is currently in progress </summary>
        private bool isPlaying = false;

        private string lastFile = string.Empty;
        /// <summary> Number of playback iterations </summary>
        private uint loopCount = 0;

        /// <summary> Delegate to use for placing a job with no arguments onto the Dispatcher </summary>
        private delegate void NoArgDelegate();
        
        /// <summary>
        /// Delegate to use for placing a job with a single string argument onto the Dispatcher
        /// </summary>
        /// <param name="arg">string argument</param>
        private delegate void OneArgDelegate(string arg);

        /// <summary>
        /// Current record/playback status text to display
        /// </summary>
        private string recordPlayStatusText = string.Empty;

        /// <summary>
        /// Radius of drawn hand circles
        /// </summary>
        private const double HandSize = 30;

        /// <summary>
        /// Thickness of drawn joint lines
        /// </summary>
        private const double JointThickness = 3;

        /// <summary>
        /// Thickness of clip edge rectangles
        /// </summary>
        private const double ClipBoundsThickness = 10;

        /// <summary>
        /// Constant for clamping Z values of camera space points from being negative
        /// </summary>
        private const float InferredZPositionClamp = 0.1f;

        /// <summary>
        /// Brush used for drawing hands that are currently tracked as closed
        /// </summary>
        private readonly Brush handClosedBrush = new SolidColorBrush(Color.FromArgb(128, 255, 0, 0));

        /// <summary>
        /// Brush used for drawing hands that are currently tracked as opened
        /// </summary>
        private readonly Brush handOpenBrush = new SolidColorBrush(Color.FromArgb(128, 0, 255, 0));

        /// <summary>
        /// Brush used for drawing hands that are currently tracked as in lasso (pointer) position
        /// </summary>
        private readonly Brush handLassoBrush = new SolidColorBrush(Color.FromArgb(128, 0, 0, 255));

        /// <summary>
        /// Brush used for drawing joints that are currently tracked
        /// </summary>
        private readonly Brush trackedJointBrush = new SolidColorBrush(Color.FromArgb(255, 68, 192, 68));

        /// <summary>
        /// Brush used for drawing joints that are currently inferred
        /// </summary>        
        private readonly Brush inferredJointBrush = Brushes.Yellow;

        /// <summary>
        /// Pen used for drawing bones that are currently inferred
        /// </summary>        
        private readonly Pen inferredBonePen = new Pen(Brushes.Gray, 1);

        /// <summary>
        /// Drawing group for body rendering output
        /// </summary>
        private DrawingGroup drawingGroup;

        /// <summary>
        /// Drawing image that we will display
        /// </summary>
        private DrawingImage imageSource;

        /// <summary>
        /// Active Kinect sensor
        /// </summary>
        private KinectSensor kinectSensor = null;

        /// <summary>
        /// Coordinate mapper to map one type of point to another
        /// </summary>
        private CoordinateMapper coordinateMapper = null;

        /// <summary>
        /// Reader for body frames
        /// </summary>
        private BodyFrameReader bodyFrameReader = null;

        /// <summary>
        /// Array for the bodies
        /// </summary>
        
        /// <summary>
        /// Current status text to display
        /// </summary>
        private string statusText = null;

        /// <summary>
        /// Initializes a new instance of the MainWindow class.
        /// </summary>
        public MainWindow()
        {
            // one sensor is currently supported
            this.kinectSensor = KinectSensor.GetDefault();

            // get the coordinate mapper
            this.coordinateMapper = this.kinectSensor.CoordinateMapper;

            // get the depth (display) extents
            FrameDescription frameDescription = this.kinectSensor.DepthFrameSource.FrameDescription;
            
            // Create the drawing group we'll use for drawing
            this.drawingGroup = new DrawingGroup();

            // Create an image source that we can use in our image control
            this.imageSource = new DrawingImage(this.drawingGroup);

            // use the window object as the view model in this simple example
            this.DataContext = this;

            // initialize the components (controls) of the window
            this.InitializeComponent();
        }

        /// <summary>
        /// INotifyPropertyChangedPropertyChanged event to allow window controls to bind to changeable data
        /// </summary>
        public event PropertyChangedEventHandler PropertyChanged;

        /// <summary>
        /// Gets the bitmap to display
        /// </summary>
        public ImageSource ImageSource
        {
            get
            {
                return this.imageSource;
            }
        }

        /// <summary>
        /// Gets or sets the current status text to display
        /// </summary>
        public string StatusText
        {
            get
            {
                return this.statusText;
            }

            set
            {
                if (this.statusText != value)
                {
                    this.statusText = value;

                    // notify any bound elements that the text has changed
                    if (this.PropertyChanged != null)
                    {
                        this.PropertyChanged(this, new PropertyChangedEventArgs("StatusText"));
                    }
                }
            }
        }

        
        /// <summary>
        /// Execute shutdown tasks
        /// </summary>
        /// <param name="sender">object sending the event</param>
        /// <param name="e">event arguments</param>
        private void MainWindow_Closing(object sender, CancelEventArgs e)
        {
            if (this.bodyFrameReader != null)
            {
                // BodyFrameReader is IDisposable
                this.bodyFrameReader.Dispose();
                this.bodyFrameReader = null;
            }

            if (this.kinectSensor != null)
            {
                this.kinectSensor.Close();
                this.kinectSensor = null;
            }
        }

        /////////////////////////////////////////////////////////////////////////////////////////////////////////below to edited
        /// <summary>
        /// Gets or sets the current status text to display for the record/playback features
        /// </summary>
        public string RecordPlaybackStatusText
        {
            get
            {
                return this.recordPlayStatusText;
            }

            set
            {
                if (this.recordPlayStatusText != value)
                {
                    this.recordPlayStatusText = value;

                    // notify any bound elements that the text has changed
                    if (this.PropertyChanged != null)
                    {
                        this.PropertyChanged(this, new PropertyChangedEventArgs("RecordPlaybackStatusText"));
                    }
                }
            }
        }
        /// <summary>
        /// Handles the user clicking on the Play button
        /// </summary>
        /// <param name="sender">object sending the event</param>
        /// <param name="e">event arguments</param>
        private void PlayButton_Click(object sender, RoutedEventArgs e)
        {
            string filePath = this.OpenFileForPlayback();

            if (!string.IsNullOrEmpty(filePath))
            {
                this.lastFile = filePath;
                this.isPlaying = true;
               // this.RecordPlaybackStatusText = Properties.Resources.PlaybackInProgressText;
                this.UpdateState();

                // Start running the playback asynchronously
                OneArgDelegate playback = new OneArgDelegate(this.PlaybackClip);
                playback.BeginInvoke(filePath, null, null);
            }
        }
        
        /// <summary>
        /// Plays back a .xef file to the Kinect sensor
        /// </summary>
        /// <param name="filePath">Full path to the .xef file that should be played back to the sensor</param>
        private void PlaybackClip(string filePath)
        {
            using (KStudioClient client = KStudio.CreateClient())
            {
                client.ConnectToService();

                // Create the playback object
                using (KStudioPlayback playback = client.CreatePlayback(filePath))
                {
                    playback.LoopCount = this.loopCount;
                    playback.Start();

                    while (playback.State == KStudioPlaybackState.Playing)
                    {
                        Thread.Sleep(500);
                    }
                }

                client.DisconnectFromService();
            }

            // Update the UI after the background playback task has completed
            this.isPlaying = false;
            this.Dispatcher.BeginInvoke(new NoArgDelegate(UpdateState));
        }

        /// <summary>
        /// Enables/Disables the record and playback buttons in the UI
        /// </summary>
        private void UpdateState()
        {
            if (this.isPlaying)
            {
                this.PlayButton.IsEnabled = false;
            }
            else
            {
                this.RecordPlaybackStatusText = string.Empty;
                this.PlayButton.IsEnabled = true;
            }
        }

        /// <summary>
        /// Launches the OpenFileDialog window to help user find/select an event file for playback
        /// </summary>
        /// <returns>Path to the event file selected by the user</returns>
        private string OpenFileForPlayback()
        {
            string fileName = string.Empty;

            OpenFileDialog dlg = new OpenFileDialog();
            dlg.FileName = this.lastFile;
            //dlg.DefaultExt = Properties.Resources.XefExtension; // Default file extension
            //dlg.Filter = Properties.Resources.EventFileDescription + " " + Properties.Resources.EventFileFilter; // Filter files by extension 
            bool? result = dlg.ShowDialog();

            if (result == true)
            {
                fileName = dlg.FileName;
            }

            return fileName;
        }

        /// <summary>
        /// Launches the SaveFileDialog window to help user create a new recording file
        /// </summary>
        /// <returns>File path to use when recording a new event file</returns>
        private string SaveRecordingAs()
        {
            string fileName = string.Empty;

            SaveFileDialog dlg = new SaveFileDialog();
            dlg.FileName = "recordAndPlaybackBasics.xef";
           // dlg.DefaultExt = Properties.Resources.XefExtension;
            dlg.AddExtension = true;
            //dlg.Filter = Properties.Resources.EventFileDescription + " " + Properties.Resources.EventFileFilter;
            dlg.CheckPathExists = true;
            bool? result = dlg.ShowDialog();

            if (result == true)
            {
                fileName = dlg.FileName;
            }

            return fileName;
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            ListItemSelected.Items.Clear();
            string[] files = Directory.GetFiles("E:\\TCS Research Internship\\TCS_VSCode\\Gesture files");

            foreach (string file in files)
            {
                string file_name = file;
                string result = string.Empty;
                file_name = file_name.Remove(0,52);
                ListItemSelected.Items.Add(file_name);
            }
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
         
        //create new xls file
        string fileTest = "E:\\TCS Research Internship\\TCS_VSCode\\"+ this.UserIDEntered.Text+".xlsx";
            string user_id = this.UserIDEntered.Text;
           
            Excel.Application oApp;
            Excel.Worksheet oSheet;
            Excel.Workbook oBook;

            oApp = null;
            oApp = new Excel.Application(); // create Excell App
            oApp.DisplayAlerts = false; // turn off alerts

            if (File.Exists(fileTest))
            {
                oBook = (Excel.Workbook)(oApp.Workbooks._Open(fileTest, System.Reflection.Missing.Value,
                   System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                   System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                   System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                   System.Reflection.Missing.Value, System.Reflection.Missing.Value)); // open the existing excel file

                int numberOfWorkbooks = oApp.Workbooks.Count; // get number of workbooks (optional)

                oSheet = (Excel.Worksheet)oBook.Worksheets[1]; // defines in which worksheet, do you want to add data
                oSheet.Name = "WorkSheet 1"; // define a name for the worksheet (optinal)

                int numberOfSheets = oBook.Worksheets.Count; // get number of worksheets (optional)
            }
            else
            {
                oApp = new Excel.Application();
                oBook = oApp.Workbooks.Add();
                oSheet = (Excel.Worksheet)oBook.Worksheets.get_Item(1);
                int numberOfWorkbooks = oApp.Workbooks.Count; // get number of workbooks (optional)
                oSheet.Name = "WorkSheet 1"; // define a name for the worksheet (optinal)

                int numberOfSheets = oBook.Worksheets.Count; // get number of worksheets (optional)
            }


            oSheet.get_Range("A1", "C1").Font.Bold = true;
            oSheet.get_Range("A1", "C1").VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;

            oSheet.Cells[1, 1] = "User ID";
            oSheet.Cells[1, 2] = "Gesture Name";
            oSheet.Cells[1, 3] = "Response";

            int gesture_no = this.ListItemSelected.SelectedIndex+1;
            
                oSheet.Cells[gesture_no+1, 2] = this.ListItemSelected.SelectedItem ;
                oSheet.Cells[gesture_no+1, 1] = this.UserIDEntered.Text;
                oSheet.Cells[gesture_no+1, 3] = this.ComboSelected.Text;
                oBook.SaveAs(fileTest);
                oBook.Close();
                oApp.Quit();

                MessageBox.Show("Response Saved.");
         
        }

        private void Button_Click_3(object sender, RoutedEventArgs e)
        {
            //create new xls file
            string fileTest = "E:\\TCS Research Internship\\TCS_VSCode\\" + this.UserIDEntered.Text + ".xlsx";
            string user_id = this.UserIDEntered.Text;

            Excel.Application oApp;
            Excel.Worksheet oSheet;
            Excel.Workbook oBook;

            oApp = null;
            oApp = new Excel.Application(); // create Excell App
            oApp.DisplayAlerts = false; // turn off alerts

            if (File.Exists(fileTest))
            {
                oBook = (Excel.Workbook)(oApp.Workbooks._Open(fileTest, System.Reflection.Missing.Value,
                   System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                   System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                   System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                   System.Reflection.Missing.Value, System.Reflection.Missing.Value)); // open the existing excel file

                int numberOfWorkbooks = oApp.Workbooks.Count; // get number of workbooks (optional)

                oSheet = (Excel.Worksheet)oBook.Worksheets[1]; // defines in which worksheet, do you want to add data
                oSheet.Name = "WorkSheet 1"; // define a name for the worksheet (optinal)

                int numberOfSheets = oBook.Worksheets.Count; // get number of worksheets (optional)
            }
            else
            {
                oApp = new Excel.Application();
                oBook = oApp.Workbooks.Add();
                oSheet = (Excel.Worksheet)oBook.Worksheets.get_Item(1);
                int numberOfWorkbooks = oApp.Workbooks.Count; // get number of workbooks (optional)
                oSheet.Name = "WorkSheet 1"; // define a name for the worksheet (optinal)

                int numberOfSheets = oBook.Worksheets.Count; // get number of worksheets (optional)
            }


            oSheet.get_Range("A1", "C1").Font.Bold = true;
            oSheet.get_Range("A1", "C1").VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;

            oSheet.Cells[1, 1] = "User ID";
            oSheet.Cells[1, 2] = "Gesture Name";
            oSheet.Cells[1, 3] = "Response";

            int gesture_no = this.ListItemSelected.SelectedIndex + 1;

            oSheet.Cells[gesture_no + 1, 2] = this.ListItemSelected.SelectedIndex + 1;
            oSheet.Cells[gesture_no + 1, 1] = this.UserIDEntered.Text;
            oSheet.Cells[gesture_no + 1, 3] = this.ComboSelected.Text;
            oBook.SaveAs(fileTest);
            oBook.Close();
            oApp.Quit();
            
            this.ListItemSelected.SelectedIndex = this.ListItemSelected.SelectedIndex + 1;
        }

        private void Button_Click_2(object sender, RoutedEventArgs e)
        {
            //create new xls file
            string fileTest = "E:\\TCS Research Internship\\TCS_VSCode\\" + this.UserIDEntered.Text + ".xlsx";
            string user_id = this.UserIDEntered.Text;

            Excel.Application oApp;
            Excel.Worksheet oSheet;
            Excel.Workbook oBook;

            oApp = null;
            oApp = new Excel.Application(); // create Excell App
            oApp.DisplayAlerts = false; // turn off alerts

            if (File.Exists(fileTest))
            {
                oBook = (Excel.Workbook)(oApp.Workbooks._Open(fileTest, System.Reflection.Missing.Value,
                   System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                   System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                   System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                   System.Reflection.Missing.Value, System.Reflection.Missing.Value)); // open the existing excel file

                int numberOfWorkbooks = oApp.Workbooks.Count; // get number of workbooks (optional)

                oSheet = (Excel.Worksheet)oBook.Worksheets[1]; // defines in which worksheet, do you want to add data
                oSheet.Name = "WorkSheet 1"; // define a name for the worksheet (optinal)

                int numberOfSheets = oBook.Worksheets.Count; // get number of worksheets (optional)
            }
            else
            {
                oApp = new Excel.Application();
                oBook = oApp.Workbooks.Add();
                oSheet = (Excel.Worksheet)oBook.Worksheets.get_Item(1);
                int numberOfWorkbooks = oApp.Workbooks.Count; // get number of workbooks (optional)
                oSheet.Name = "WorkSheet 1"; // define a name for the worksheet (optinal)

                int numberOfSheets = oBook.Worksheets.Count; // get number of worksheets (optional)
            }


            oSheet.get_Range("A1", "C1").Font.Bold = true;
            oSheet.get_Range("A1", "C1").VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;

            oSheet.Cells[1, 1] = "User ID";
            oSheet.Cells[1, 2] = "Gesture Number";
            oSheet.Cells[1, 3] = "Response";

            int gesture_no = this.ListItemSelected.SelectedIndex + 1;

            oSheet.Cells[gesture_no + 1, 2] = this.ListItemSelected.SelectedIndex + 1;
            oSheet.Cells[gesture_no + 1, 1] = this.UserIDEntered.Text;
            oSheet.Cells[gesture_no + 1, 3] = this.ComboSelected.Text;
            oBook.SaveAs(fileTest);
            oBook.Close();
            oApp.Quit();
            

            this.ListItemSelected.SelectedIndex = this.ListItemSelected.SelectedIndex - 1;
        }

        private void ComboSelected_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {

        }

        
            
    }
}

