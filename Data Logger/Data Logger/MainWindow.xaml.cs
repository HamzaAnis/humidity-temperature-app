using System;
using System.Collections.Generic;
using System.Data;
using System.IO.Ports;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using ClosedXML.Attributes;
using ClosedXML.Excel;

namespace Data_Logger
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
       
        DataTable[] table = new DataTable[16];
        private String portName = "";
        private SerialPort port;
        public MainWindow()
        {
            InitializeComponent();
            this.Closing += MainWindow_Closing;

            declare();
            setPortNames();

        }
        private void MainWindow_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            if (port.IsOpen)
            {
                port.Close();
            }
            System.Windows.Application.Current.Shutdown();
        }
        public void declare()
        {
            running = false;
            zoneLabel1.Content = "Temp=0°C\nHumidity=0";
            zoneLabel2.Content = "Temp=0°C\nHumidity=0";
            zoneLabel3.Content = "Temp=0°C\nHumidity=0";
            zoneLabel4.Content = "Temp=0°C\nHumidity=0";
            zoneLabel5.Content = "Temp=0°C\nHumidity=0";
            zoneLabel6.Content = "Temp=0°C\nHumidity=0";
            zoneLabel7.Content = "Temp=0°C\nHumidity=0";
            zoneLabel8.Content = "Temp=0°C\nHumidity=0";
            zoneLabel9.Content = "Temp=0°C\nHumidity=0";
            zoneLabel10.Content = "Temp=0°C\nHumidity=0";
            zoneLabel11.Content = "Temp=0°C\nHumidity=0";
            zoneLabel12.Content = "Temp=0°C\nHumidity=0";
            zoneLabel13.Content = "Temp=0°C\nHumidity=0";
            zoneLabel14.Content = "Temp=0°C\nHumidity=0";
            zoneLabel15.Content = "Temp=0°C\nHumidity=0";
            zoneLabel16.Content = "Temp=0°C\nHumidity=0";
            for (int i = 0; i < 16; i++)
            {
                table[i] = new DataTable();
                table[i].Columns.Add("Dosage", typeof(int));
                table[i].Columns.Add("Drug", typeof(string));
                table[i].Columns.Add("Patient", typeof(string));
                table[i].Columns.Add("Date", typeof(DateTime));
            }
        }
        public void setPortNames()
        {
            string[] ports = SerialPort.GetPortNames();
            foreach (var port in ports)
            {
                portsSelectComboBox.Items.Add(port);
            }

            if (ports.Length > 0)
            {
                portsSelectComboBox.SelectedIndex = 0;
            }

        }
        private void Button_Click(object sender, RoutedEventArgs e)
        {

        }
        private Boolean running;
        public void readPort()
        {
            try
            {
               port = new SerialPort(portName,
                    9600, Parity.None, 8, StopBits.One);
                port.Open();
                int count = 0;
                while (port.IsOpen)
                {
                    String line = port.ReadLine();
                    String time = port.ReadLine();
                    String[] split = line.Split('|');
                    if (split.Length == 2)
                    {
                        String[] temperatures = split[0].Split(',');
                        String[] humidity = split[1].Split(',');
                        if (temperatures.Length >= 16 && humidity.Length >= 16)
                        {
                            this.Dispatcher.InvokeAsync(new Action(() =>
                            {
                                this.Dispatcher.BeginInvoke(new Action(() =>
                                {
                                    zoneLabel1.Content = "Temp=" + temperatures[0] + "°C\nHumidity=" + humidity[0];
                                    zoneLabel2.Content = "Temp=" + temperatures[1] + "°C\nHumidity=" + humidity[1];
                                    zoneLabel3.Content = "Temp=" + temperatures[2] + "°C\nHumidity=" + humidity[2];
                                    zoneLabel4.Content = "Temp=" + temperatures[3] + "°C\nHumidity=" + humidity[3];
                                    zoneLabel5.Content = "Temp=" + temperatures[4] + "°C\nHumidity=" + humidity[4];
                                    zoneLabel6.Content = "Temp=" + temperatures[5] + "°C\nHumidity=" + humidity[5];
                                    zoneLabel7.Content = "Temp=" + temperatures[6] + "°C\nHumidity=" + humidity[6];
                                    zoneLabel8.Content = "Temp=" + temperatures[7] + "°C\nHumidity=" + humidity[7];
                                    zoneLabel9.Content = "Temp=" + temperatures[8] + "°C\nHumidity=" + humidity[8];
                                    zoneLabel10.Content = "Temp=" + temperatures[9] + "°C\nHumidity=" + humidity[9];
                                    zoneLabel11.Content = "Temp=" + temperatures[10] + "°C\nHumidity=" + humidity[10];
                                    zoneLabel12.Content = "Temp=" + temperatures[11] + "°C\nHumidity=" + humidity[11];
                                    zoneLabel13.Content = "Temp=" + temperatures[12] + "°C\nHumidity=" + humidity[12];
                                    zoneLabel14.Content = "Temp=" + temperatures[13] + "°C\nHumidity=" + humidity[13];
                                    zoneLabel15.Content = "Temp=" + temperatures[14] + "°C\nHumidity=" + humidity[14];
                                    zoneLabel16.Content = "Temp=" + temperatures[15] + "°C\nHumidity=" + humidity[15];
                                }));

                            }));
                        }
                        Console.WriteLine(temperatures.Length);
                        Console.WriteLine(humidity.Length);
                        Console.WriteLine();
                        count++;
                    }
                    Console.WriteLine("Temperatures" + line);
                    Console.WriteLine("Time" + time);
                    Console.WriteLine();
                }
            }
            catch (Exception exception)
            {
                Console.WriteLine(exception);
            }
        }
        private void startButton_Click(object sender, RoutedEventArgs e)
        {
            if (!running)
            {
                startButton.Content = "Running";
                portName = (string) portsSelectComboBox.SelectedValue;
                Thread thr = new Thread(new ThreadStart(readPort));
                thr.Start();
            }
            running = true;
        }
        private void saveButton_Click(object sender, RoutedEventArgs e)
        {
            MessageBox.Show((string)portsSelectComboBox.SelectedValue);
        }
        public void Create(String filePath)
        {
            String[] sensors =
            {
                "Sensor 1", "Sensor 2", "Sensor 3", "Sensor 4", "Sensor 5", "Sensor 6", "Sensor 7", "Sensor 8",
                "Sensor 9", "Sensor 10", "Sensor 11", "Sensor 12", "Sensor 13", "Sensor 14", "Sensor 15", "Sensor 16"
            };
            // From a DataTable
            
            using (var wb = new XLWorkbook())
            {
                for (int i = 0; i < 16; i++)
                {

                    var ws = wb.Worksheets.Add(sensors[i]);

                    ws.Range(1, 1, 1, 4).Merge().AddToNamed("Titles");
                    ws.Cell(1, 1).InsertTable(table[i]);

                    // Prepare the style for the titles

                    ws.Columns().AdjustToContents();


                }
                wb.SaveAs(filePath);
            }

        }
    }
}
