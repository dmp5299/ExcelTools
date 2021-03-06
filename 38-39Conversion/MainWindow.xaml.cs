﻿using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using _38_39Conversion._38ConversionFiles;
using _38_39Conversion.XmlGenerationFiles;
using System.ComponentModel;
using System.Threading;
using _38_39Conversion.ExcelObjects;
using _38_39Conversion.Interfaces;
using _38_39Conversion.CustomExceptions;
using _38_39Conversion.ContentaObjects;
using ContentaDataExport;
using ContentaDataExport.Utils;
using PCMClientLib;
using PCMPortalLib;
using ContentaDataExport.ContentaClasses;

namespace _38_39Conversion
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public BackgroundWorker worker;
        public BackgroundWorker worker1;
        PCMClientLib.PCMConnection conn;

        public MainWindow()
        {
            InitializeComponent();
            initializeBackgroundWorker();
        }

        private void GenerateContentaReport_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                DataConnection dataConn = new DataConnection(ContentaUtils.getCookie());
                checkContentaConnOptions(dataConn);
                IPCMcommand command = connectToDB(dataConn.Host, dataConn.Socket, dataConn.Database);
                ContentaModule m = new ContentaModule();
                m.getContentaObjects(command);
            }
            catch(Exception ex)
            {
                System.Windows.MessageBox.Show(ex.Message);
            }
        }

        private IPCMcommand connectToDB(string host, string socket, string database)
        {
            IPCMcommand command = null;
            try
            {
                conn = new PCMConnection();
                command = ContentaConnection.getCommandObject(conn,host, socket, database);
            }
            catch (Exception ex)
            {
                throw new Exception("Error connecting to CSDB: " + ex.Message);
            }
            return command;
        }

        private void checkContentaConnOptions(DataConnection data)
        {
            string errorMessage = "";
            foreach (var prop in data.GetType().GetProperties())
            {
                if (prop.GetValue(data, null) == "")
                {
                    errorMessage += "Missing " + prop.Name + " connection option\n";
                }
            }
            if(errorMessage.Length > 0)
            {
                throw new Exception(errorMessage);
            }
        }

        private void ContentaOptions_Click(object sender, RoutedEventArgs e)
        {
            ContentaOptions options = new ContentaOptions();
            options.Show();
        }

        private void Close_Click(object sender, RoutedEventArgs e)
        {
            System.Windows.Application.Current.Shutdown();
        }

        private void _38Conversion_Click(object sender, RoutedEventArgs e)
        {
            _411GenerationGrid.Visibility = Visibility.Hidden;
            _38ConversionGrid.Visibility = Visibility.Visible;
        }

        private void _Xml411Generation_Click(object sender, RoutedEventArgs e)
        {
            _38ConversionGrid.Visibility = Visibility.Hidden;
            _411GenerationGrid.Visibility = Visibility.Visible;
        }

        private void Generate411s_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (!string.IsNullOrEmpty(FilePath_411Text.Text))
                {
                    build_411s.IsEnabled = false;
                    ExcelParser parser = new ExcelParser(FilePath_411Text.Text);
                    parser.getExcelData(FilePath_411Text.Text);
                    int count = parser._411s.Count;
                    XmlGenerationStatus.Maximum = count;
                    worker.RunWorkerAsync(parser._411s);
                    
                }
                else
                    throw new InputException("Enter excel file path");
            }
            catch(InputException a)
            {
                System.Windows.Forms.MessageBox.Show(a.Message);
            }
            catch(Exception ex)
            {
                build_411s.IsEnabled = true;
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }
        }

        private void Convert_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if((bool)ConvertTo39_Checkbox.IsChecked || (bool)Clean38s_Checkbox.IsChecked)
                {
                    if (!string.IsNullOrEmpty(FilePathText.Text))
                    {
                        //string[] files = Directory.GetFiles(FilePathText.Text, "*.xlsx| ");
                        List<object> workerArguments = new List<object>();
                        var files = Directory.EnumerateFiles(FilePathText.Text, "*.*", SearchOption.AllDirectories)
                        .Where(s => s.EndsWith(".xlsx") || s.EndsWith(".xls") && !s.Contains("39"));
                        workerArguments.Add(files);
                        workerArguments.Add(ConvertTo39_Checkbox.IsChecked);
                        workerArguments.Add(Clean38s_Checkbox.IsChecked);
                        _38ConversionStatus.Maximum = files.Count();
                        if (files.Count() > 0)
                        {
                            convert38s.IsEnabled = false;
                            worker1.RunWorkerAsync(workerArguments);
                        }
                        else
                        {
                            throw new InputException("There are no .xlsx files in the selected directory");
                        }

                    }
                    else
                    {
                        throw new InputException("You must enter a path");
                    }
                }
                else
                {
                    throw new InputException("You must check Clean38s or Convert39s");
                }
            }
            catch(InputException a)
            {
                System.Windows.Forms.MessageBox.Show(a.Message);
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        private void Browse38Excel_Click(object sender, RoutedEventArgs e)
        {
            using (var fbd = new FolderBrowserDialog())
            {
                DialogResult result = fbd.ShowDialog();
                if ((result == System.Windows.Forms.DialogResult.OK) && !string.IsNullOrWhiteSpace(fbd.SelectedPath))
                {
                    FilePathText.Text = fbd.SelectedPath;
                }
            }
        }

        private void Browse411Excel_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();

            // Set filter for file extension and default file extension 
            dlg.Filter = "excel files (*.xlsx)|*.xlsx";


            // Display OpenFileDialog by calling ShowDialog method 
            Nullable<bool> result = dlg.ShowDialog();


            // Get the selected file name and display in a TextBox 
            if (result == true)
            {
                // Open document 
                string filename = dlg.FileName;
                FilePath_411Text.Text = filename;
            }
        }

        private void Window_ContentRendered(object sender, EventArgs e)
        {
            initializeBackgroundWorker();
        }

        private void initializeBackgroundWorker()
        {
            worker = new BackgroundWorker();
            worker.WorkerReportsProgress = true;
            worker.DoWork += worker_dowork;
            worker.ProgressChanged += worker_ProgressChanged;
            worker.RunWorkerCompleted += new RunWorkerCompletedEventHandler(
            backgroundWorker1_RunWorkerCompleted);

            worker1 = new BackgroundWorker();
            worker1.WorkerReportsProgress = true;
            worker1.DoWork += _38Converter_dowork;
            worker1.ProgressChanged += _38Converter_ProgressChanged;
            worker1.RunWorkerCompleted += new RunWorkerCompletedEventHandler(
            _38Converter_RunWorkerCompleted);
        }

        void worker_dowork(object sender, DoWorkEventArgs e)
        {
            BackgroundWorker worker = sender as BackgroundWorker;
            try
            {
                ExcelParser.get411Dms(worker,(List<_411Module>)e.Argument);
            }
            catch(Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        private void backgroundWorker1_RunWorkerCompleted(
            object sender, RunWorkerCompletedEventArgs e)
        { 
            if (e.Error != null)
            {
                XmlGenerationStatus.Value = 0;
                System.Windows.Forms.MessageBox.Show(e.Error.Message);
            }
            build_411s.IsEnabled = true;
        }

        void worker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            XmlGenerationStatus.Value = e.ProgressPercentage;
            
        }

        void _38Converter_dowork(object sender, DoWorkEventArgs e)
        {
            BackgroundWorker worker = sender as BackgroundWorker;
            try
            {
                List<IDictionary<string, object>> all38Data = new List<IDictionary<string, object>>();
                int i = 0;
                List<object> arguments = (List<object>)e.Argument;
                Boolean convert = (bool)arguments[1];
                Boolean clean = (bool)arguments[2];
                foreach (string file in (IEnumerable<string>)arguments[0])
                {
                    if (!file.Contains("39"))
                    {
                        I38Data _xlsConversionObject = new Dash38Xls();
                        I38Data _xlsxConversionObject = new Dash38Xlsx();
                        if (System.IO.Path.GetExtension(file).Equals(".xlsx"))
                        {
                            if(convert)
                            {
                                Dash39.build39File(_xlsxConversionObject.parseThirtyEightFile(file, clean));
                            }
                            else
                            {
                                _xlsxConversionObject.parseThirtyEightFile(file,clean);
                            }
                        }
                        else
                        {
                            if (convert)
                            {
                                Dash39.build39File(_xlsConversionObject.parseThirtyEightFile(file,clean));
                            }
                            else
                            {
                                _xlsConversionObject.parseThirtyEightFile(file,clean);
                            }
                        }
                    }
                    worker1.ReportProgress(i + 1);
                    //Thread.Sleep(5);
                    i++;
                }
                System.Windows.Forms.MessageBox.Show("done");
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        void _38Converter_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            _38ConversionStatus.Value = e.ProgressPercentage;
        }

        private void _38Converter_RunWorkerCompleted(
            object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Error != null)
            {
                _38ConversionStatus.Value = 0;
                System.Windows.Forms.MessageBox.Show(e.Error.Message);
            }
            convert38s.IsEnabled = true;
        }
    }

}
