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
using DocumentFormat.OpenXml.Wordprocessing;

namespace Data_Logger
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    { 
        DataTable[] table = new DataTable[16];
        DateTime before;
        private String portName = "";
        private SerialPort port;
        private bool firstTime = true;
        public MainWindow()
        {
            InitializeComponent();
            this.Closing += MainWindow_Closing;
            before = DateTime.Now;
            declare();
            setPortNames();

        }
        private void MainWindow_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            Create();
            if (port!=null) 
                if(port.IsOpen)
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
                table[i].Columns.Add("Date", typeof(string));
                table[i].Columns.Add("Time", typeof(string));
                table[i].Columns.Add("Temperature", typeof(string));
                table[i].Columns.Add("Humidity", typeof(string));
                table[i].Columns.Add("Remarks", typeof(string));
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
                    String temp1 = port.ReadLine();
                    String temp2 = port.ReadLine();
                    String line="";
                    String time="";
                    if (temp1.Contains(":"))
                    {
                        time = temp1;
                        line = temp2;
                    }else if (temp2.Contains(":"))
                    {
                        time = temp2;
                        line = temp1;
                    }
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
                            DateTime now=DateTime.Now;
                            int minutes = (int)now.Subtract(before).TotalMinutes;
                            Console.WriteLine("Minutes are : "+minutes);
                            if (minutes >= 5|| firstTime)
                            {
                                for (int i = 0; i < 16; i++)
                                {
                                    table[i].Rows.Add((string)time, (string)temperatures[i], (string)humidity[i], "");
                                }
                                before=DateTime.Now;
                                firstTime = false;
                            }
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

        private void TextBlock_Clicked(object sender, RoutedEventArgs e)
        {
            MessageBox.Show(
                "Humidity and Temperature logger\nDeveloped by: Hamza Anis\nVersion: 0.1\nContact: hamzaanis9514@gmail.com\nPhone: 00923420011719");
        }
        private void saveButton_Click(object sender, RoutedEventArgs e)
        {
            Create();
        }
        public void Create()
        {
            String filePath ="logs/" +DateTime.Now.ToString("dd-MM-yyyy_h-mm-ss-tt")+".xlsx";
           // filePath = "Nice.xlsx";
            String[] sensors =
            {
                "Zone 1", "Zone 2", "Zone 3", "Zone 4", "Zone 5", "Zone 6", "Zone 7", "Zone 8",
                "Zone 9", "Zone 10", "Zone 11", "Zone 12", "Zone 13", "Zone 14", "Zone 15", "Zone 16"
            };
            // From a DataTable
            
            using (var wb = new XLWorkbook())
            {
                for (int i = 0; i < 16; i++)
                {

                    var ws = wb.Worksheets.Add(sensors[i]);

                    ws.Range(1, 1, 1, 5).Merge().AddToNamed("Titles");
                    ws.Cell(1, 1).InsertTable(table[i]);

                    // Prepare the style for the titles

                    ws.Columns().AdjustToContents();


                }
                wb.SaveAs(filePath);
            }

            for (int i = 0; i < 16; i++)
            {
                table[i].Clear();
            }

        }
    }
}
