using System;
using System.IO.Ports;
using System.Xml.Linq;
using System.Linq;
using System.IO;
using FTD2XX_NET;
using System.Threading.Tasks;

namespace SerialControl
{
    class Program
    {
        
        static readonly byte[] masks = new byte[]
        {
            0x01, //CH-1
            0x02, //CH-2
            0x04, //CH-3
            0x08, //CH-4
            0x10, //CH-5
            0x20, //CH-6
            0x40, //CH-7
            0x80  //CH-8
        };

        static void Main(string[] args)
        {
            SerialPortSettings serialPortSettings = GetSerialPortSettingsFromConfig();
            UInt32 ftdiIndex;
            initFTDI(serialPortSettings.ComPort, out ftdiIndex);
            if (args.Length != 2)
            {
                Console.WriteLine("Usage: RelayControl <channel (int 1-8)> <state (int 0=OFF 1=ON)>");
                return;
            }

            int arg0;
            int arg1;
            if (!int.TryParse(args[0], out arg0) || !int.TryParse(args[1], out arg1) || arg0 < 1 || arg0 > 8 || arg1 < 0 || arg1 > 1)
            {
                Console.WriteLine($"Invalid argument passed. Channel must be an integer between 1 and 8. State must be an integer 0 or 1");
                return;
            }
            arg0--;
            byte state = getRelayState(ftdiIndex);
            if (arg1 == 1)
            {
                state |= masks[arg0];
            }
            else if (arg1 == 0)
            {
                byte imask = (byte)~masks[arg0];
                state &= imask;
            }
            byte[] message = { state };
            string port = "COM" + serialPortSettings.ComPort.ToString();

            try
            {
                using (SerialPort serialPort = new SerialPort(port, serialPortSettings.BaudRate, serialPortSettings.Parity, serialPortSettings.DataBits, serialPortSettings.StopBits))
                {
                    serialPort.Open();
                    if (serialPort.IsOpen)
                    {
                        serialPort.Write(message, 0, message.Length);
                        System.Threading.Thread.Sleep(10);
                        serialPort.Close();
                    }
                    else
                    {
                        Console.WriteLine("Failed to open the serial port.");
                    }
                }
            }
            catch (UnauthorizedAccessException ex)
            {
                Console.WriteLine($"Access to the COM port is denied. {ex.Message}");
            }
            catch (ArgumentException ex)
            {
                Console.WriteLine($"The port name does not begin with 'COM', or the file type of the port is not supported. {ex.Message}");
            }
            catch (IOException ex)
            {
                Console.WriteLine($"The COM port is in an invalid state. {ex.Message}");
            }
            catch (InvalidOperationException ex)
            {
                Console.WriteLine($"The specified port on the current instance of the SerialPort is already open. {ex.Message}");
            }
        }

        private static SerialPortSettings GetSerialPortSettingsFromConfig()
        {
            var exeLocation = System.Reflection.Assembly.GetExecutingAssembly().Location;
            var exeDirectory = System.IO.Path.GetDirectoryName(exeLocation);
            var configFilePath = System.IO.Path.Combine(exeDirectory, "config.xml");

            var config = XDocument.Load(configFilePath);

            var comPort = int.Parse(config.Descendants("COMPort").First().Value);
            var baudRate = int.Parse(config.Descendants("BaudRate").First().Value);
            var parity = (Parity)Enum.Parse(typeof(Parity), config.Descendants("Parity").First().Value);
            var dataBits = int.Parse(config.Descendants("DataBits").First().Value);
            var stopBits = (StopBits)Enum.Parse(typeof(StopBits), config.Descendants("StopBits").First().Value);

            return new SerialPortSettings(comPort, baudRate, parity, dataBits, stopBits);
        }

        private static byte getRelayState(UInt32 deviceIndex)
        {
            FTDI.FT_STATUS ftStatus = FTDI.FT_STATUS.FT_OK;
            FTDI newFTDI = new FTDI();
            ftStatus = newFTDI.OpenByIndex(deviceIndex);
            if (ftStatus != FTDI.FT_STATUS.FT_OK) Console.WriteLine($"FTDI Error: {ftStatus}");
            byte[] buffer = new byte[1];
            UInt32 readBytes = 0;
            ftStatus = newFTDI.Read(buffer, 1, ref readBytes);
            if (ftStatus != FTDI.FT_STATUS.FT_OK) Console.WriteLine($"FTDI Error: {ftStatus}");
            ftStatus = newFTDI.Close();
            if (ftStatus != FTDI.FT_STATUS.FT_OK) Console.WriteLine($"FTDI Error: {ftStatus}");
            return buffer[0];
        }

        private static FTDI.FT_STATUS initFTDI(int comport, out UInt32 index)
        {
            UInt32 ftdiDeviceCount = 0;
            UInt32 deviceIndex = 1000; //big number
            index = deviceIndex;
            FTDI.FT_STATUS ftStatus = FTDI.FT_STATUS.FT_OK;
            FTDI newFTDI = new FTDI();
            ftStatus = newFTDI.GetNumberOfDevices(ref ftdiDeviceCount);
            FTDI.FT_DEVICE_INFO_NODE[] ftdiDeviceList = new FTDI.FT_DEVICE_INFO_NODE[ftdiDeviceCount];
            ftStatus = newFTDI.GetDeviceList(ftdiDeviceList);
            for (UInt32 i = 0; i < ftdiDeviceCount; i++)
            {
                //Open each FTDI device found and check if the comport matches what was passed
                ftStatus = newFTDI.OpenByIndex(i);
                if (ftStatus != FTDI.FT_STATUS.FT_OK) Console.WriteLine($"Error opening FTDI device at index {i}");
                string com;
                newFTDI.GetCOMPort(out com);
                com = com.Substring(3);
                if (int.Parse(com) == comport)
                {
                    //If the comport matches, break from the loop before closing the device
                    deviceIndex = i;
                    index = i;
                    break;
                }
                ftStatus = newFTDI.Close();
                if (ftStatus != FTDI.FT_STATUS.FT_OK) return ftStatus;
            }
            if (deviceIndex != 1000)
            {
                //if the loop was exited because an FTDI with the right comport was found, set bitbang mode and close
                ftStatus = newFTDI.SetBitMode(0xFF, FTDI.FT_BIT_MODES.FT_BIT_MODE_ASYNC_BITBANG); //mask sets all pins to output
                if (ftStatus != FTDI.FT_STATUS.FT_OK) return ftStatus;
                ftStatus = newFTDI.Close();
                if (ftStatus != FTDI.FT_STATUS.FT_OK) return ftStatus;
            } else Console.WriteLine($"No FTDI device found at COMPORT {comport}");
            return FTDI.FT_STATUS.FT_OK;
        }
    }

    class SerialPortSettings
    {
        public int ComPort { get; }
        public int BaudRate { get; }
        public Parity Parity { get; }
        public int DataBits { get; }
        public StopBits StopBits { get; }

        public SerialPortSettings(int COMPort, int baudRate, Parity parity, int dataBits, StopBits stopBits)
        {
            ComPort = COMPort;
            BaudRate = baudRate;
            Parity = parity;
            DataBits = dataBits;
            StopBits = stopBits;
        }
    }
}
