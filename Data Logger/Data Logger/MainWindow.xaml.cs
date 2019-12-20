using System;
using System.Collections.Generic;
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

            setPortNames();
            zoneLabel1.Content = "Temp=80°C\nHumidity=20";
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
        }
    }
}
