using System;
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

namespace _38_39Conversion
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public BackgroundWorker worker;
        public BackgroundWorker worker1;


        public MainWindow()
        {
            InitializeComponent();

            initializeBackgroundWorker();
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
                    throw new ArgumentException("Enter excel file path");
            }
            catch(ArgumentException a)
            {
                System.Windows.Forms.MessageBox.Show(a.Message);
            }
            catch(IOException i)
            {
                System.Windows.Forms.MessageBox.Show(i.Message);
            }
            catch(Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }
        }

        private void Convert_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (!string.IsNullOrEmpty(FilePathText.Text))
                {

                    convert38s.IsEnabled = false;
                    //string[] files = Directory.GetFiles(FilePathText.Text, "*.xlsx| ");
                    var files = Directory.EnumerateFiles(FilePathText.Text, "*.*", SearchOption.AllDirectories)
                    .Where(s => s.EndsWith(".xlsx") || s.EndsWith(".xls" ) && !s.Contains("39"));
                    _38ConversionStatus.Maximum = files.Count();
                    if (files.Count() > 0)
                    {
                        worker1.RunWorkerAsync(files);
                    }
                    else
                    {
                        throw new IOException("There are no .xlsx files in the selected directory");
                    }
                }
                else
                {
                    throw new ArgumentException("You must enter a path");
                }
            }
            catch(ArgumentException a)
            {
                System.Windows.Forms.MessageBox.Show(a.Message);
            }
            catch (IOException i)
            {
                System.Windows.Forms.MessageBox.Show(i.Message);
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
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
            catch (IOException i)
            {
                throw new IOException(i.Message);
            }

        }

        private void backgroundWorker1_RunWorkerCompleted(
            object sender, RunWorkerCompletedEventArgs e)
        {
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
                foreach (string file in (IEnumerable<string>)e.Argument)
                {
                    if (!file.Contains("39"))
                    {
                        if (System.IO.Path.GetExtension(file).Equals(".xlsx"))
                        {
                            Dash39.build39File(Dash38.parseThirtyEightFile(file));
                        }
                        else
                        {
                            Dash39.build39File(Dash38.parseThirtyEightXlsFile(file));
                        }
                    }
                    worker1.ReportProgress(i + 1);
                    Thread.Sleep(100);
                    i++;
                }
                System.Windows.Forms.MessageBox.Show("done");
            }
            catch (IOException i)
            {
                throw new IOException(i.Message);
            }

        }

        void _38Converter_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            _38ConversionStatus.Value = e.ProgressPercentage;
        }

        private void _38Converter_RunWorkerCompleted(
            object sender, RunWorkerCompletedEventArgs e)
        {
            convert38s.IsEnabled = true;

        }
    }

}
