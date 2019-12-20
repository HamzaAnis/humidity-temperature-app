using System;
using System.Collections.Generic;
using System.Data;
using System.IO.Ports;
using System.Linq;
using System.Text;
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

        public MainWindow()
        {
            InitializeComponent();
            declare();
            setPortNames();

        }

        public void declare()
        {
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

        private void startButton_Click(object sender, RoutedEventArgs e)
        {
            MessageBox.Show((string) portsSelectComboBox.SelectedValue);
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
