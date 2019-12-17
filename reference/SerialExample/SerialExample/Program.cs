using System;
using System.IO.Ports;
namespace SerialPortExample
{
    class SerialPortProgram
    {
        // Create the serial port with basic settings
        private SerialPort port = new SerialPort("COM4",
            9600, Parity.None, 8, StopBits.One);

        [STAThread]
        static void Main(string[] args)
        {
            // Instatiate this class
            new SerialPortProgram();
        }

        private SerialPortProgram()
        {
            Console.WriteLine("Incoming Data:");

            // Begin communications
            port.Open();
            int count = 0;
            while (true)
            {
                String a = port.ReadLine();
                if (a!=String.Empty)
                {
                    Console.WriteLine(count+" : "+a);
                    count++;
                }
            }

            // Enter an application loop to keep this thread alive
        }

    }
}