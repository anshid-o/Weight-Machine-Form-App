using System;
using System.Drawing;
using System.IO;
using System.IO.Ports;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace WeightGetter
{
    public partial class myForm : Form
    {
        private string lastLine = string.Empty;
        private const int WM_HOTKEY = 0x0312;
        private const int HOTKEY_ID = 9000;
        private SerialPort serialPort1;

        [DllImport("Gdi32.dll", EntryPoint = "CreateRoundRectRgn")]
        private static extern IntPtr CreateRoundRectRgn
        (
            int nLeftRect,     // x-coordinate of upper-left corner
            int nTopRect,      // y-coordinate of upper-left corner
            int nRightRect,    // x-coordinate of lower-right corner
            int nBottomRect,   // y-coordinate of lower-right corner
            int nWidthEllipse, // width of ellipse
            int nHeightEllipse // height of ellipse
        );

        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, IntPtr ProcessId);

        [DllImport("user32.dll")]
        private static extern bool AttachThreadInput(uint idAttach, uint idAttachTo, bool fAttach);

        [DllImport("kernel32.dll")]
        private static extern uint GetCurrentThreadId();

        public Point mouseLocation;

        public myForm()
        {
            InitializeComponent();

            this.FormBorderStyle = FormBorderStyle.None;
            Region = System.Drawing.Region.FromHrgn(CreateRoundRectRgn(0, 0, Width, Height, 20, 20));
            InitializeRichTextBox();
            RegisterHotKey(Handle, HOTKEY_ID, MOD_CONTROL, (int)Keys.Q);
            InitializeSerialPort();
        }

        private void InitializeRichTextBox()
        {
            richTextBox1.ScrollBars = RichTextBoxScrollBars.Both;
            richTextBox1.WordWrap = false;
            richTextBox1.ReadOnly = false; // Allow editing
            richTextBox1.SelectionStart = 0;
            richTextBox1.SelectionLength = 0;
            richTextBox1.BorderStyle = BorderStyle.None; // No border
            richTextBox1.Margin = new Padding(10); // Padding for the RichTextBox
        }

        private void InitializeSerialPort()
        {
            serialPort1 = new SerialPort
            {
                PortName = "COM3", // Update to your actual serial port name
                BaudRate = 9600,   // Update to match your weight machine
                Parity = Parity.None,
                DataBits = 8,
                StopBits = StopBits.One,
                Handshake = Handshake.None,
                ReadTimeout = 500,
                WriteTimeout = 500
            };

            try
            {
                // Check if the port is already open and close it before opening again
                if (serialPort1.IsOpen)
                {
                    serialPort1.Close();
                }

                serialPort1.Open();

                if (serialPort1.IsOpen)
                {
                    richTextBox1.AppendText("Connection Successful!" + Environment.NewLine);
                }
            }
            catch (UnauthorizedAccessException)
            {
                richTextBox1.AppendText("Access to the port is denied." + Environment.NewLine);
                try
                {
                    serialPort1.Open();
                    if (serialPort1.IsOpen)
                    {
                        richTextBox1.AppendText("Connection Successful!" + Environment.NewLine);
                    }
                    else
                    {
                        richTextBox1.AppendText("Connection Unsuccessful!" + Environment.NewLine);
                    }
                }
                catch (Exception ex)
                {
                    richTextBox1.AppendText($"Connection failed: {ex.Message}" + Environment.NewLine);
                }


            }
            catch (IOException)
            {
                richTextBox1.AppendText("The port is in an invalid state. Make sure the device is connected and not in use by another application." + Environment.NewLine);
            }
            catch (ArgumentException)
            {
                richTextBox1.AppendText("The port name does not begin with 'COM' or is not valid." + Environment.NewLine);
            }
            catch (InvalidOperationException)
            {
                richTextBox1.AppendText("The specified port is already open." + Environment.NewLine);
            }
            catch (Exception ex)
            {
                richTextBox1.AppendText($"Connection failed: {ex.Message}" + Environment.NewLine);
            }

            serialPort1.DataReceived += new SerialDataReceivedEventHandler(SerialPort1_DataReceived);
        }

        private void SerialPort1_DataReceived(object sender, SerialDataReceivedEventArgs e)
        {
            try
            {
                string rawData = serialPort1.ReadLine(); // Adjust according to your weight machine's data format
                string parsedData = ParseWeightData(rawData);

                Invoke(new Action(() =>
                {
                    richTextBox1.AppendText(parsedData + Environment.NewLine); // Update the RichTextBox with the received data
                    lastLine = parsedData; // Update lastLine with the latest received data
                }));
            }
            catch (Exception ex)
            {
                Invoke(new Action(() =>
                {
                    richTextBox1.AppendText($"Error reading serial port: {ex.Message}" + Environment.NewLine);
                }));
            }
        }

        private string ParseWeightData(string rawData)
        {
            // Remove the leading '+' character
            if (rawData.StartsWith("+"))
            {
                rawData = rawData.Substring(1);
            }

            // Remove leading zeros
            rawData = rawData.TrimStart('0');

            // Remove the last four characters, assuming they are " G U"
            if (rawData.Length >= 4)
            {
                rawData = rawData.Substring(0, rawData.Length - 4);
            }

            // Remove any remaining trailing spaces
            rawData = rawData.Trim();

            return rawData;
        }

        private void richTextBox1_TextChanged(object sender, EventArgs e)
        {
            for (int i = richTextBox1.Lines.Length - 1; i >= 0; i--)
            {
                if (!string.IsNullOrWhiteSpace(richTextBox1.Lines[i]))
                {
                    lastLine = richTextBox1.Lines[i];
                    break;
                }
            }
        }

        protected override void WndProc(ref Message m)
        {
            base.WndProc(ref m);

            if (m.Msg == WM_HOTKEY && m.WParam.ToInt32() == HOTKEY_ID)
            {
                IntPtr hForegroundWindow = GetForegroundWindow();
                uint dwThreadId = GetWindowThreadProcessId(hForegroundWindow, IntPtr.Zero);
                uint dwThisThreadId = GetCurrentThreadId();

                if (AttachThreadInput(dwThisThreadId, dwThreadId, true))
                {
                    SetForegroundWindow(hForegroundWindow);
                    AttachThreadInput(dwThisThreadId, dwThreadId, false);
                }

                SendTextToActiveWindow(lastLine);
            }
        }

        private void SendTextToActiveWindow(string text)
        {
            IntPtr hForegroundWindow = GetForegroundWindow();
            uint dwThreadId = GetWindowThreadProcessId(hForegroundWindow, IntPtr.Zero);
            uint dwThisThreadId = GetCurrentThreadId();

            if (AttachThreadInput(dwThisThreadId, dwThreadId, true))
            {
                SetForegroundWindow(hForegroundWindow);
                AttachThreadInput(dwThisThreadId, dwThreadId, false);
            }
            System.Threading.Thread.Sleep(500);
            foreach (char c in text)
            {
                SendKey(c);
                System.Threading.Thread.Sleep(10); // Small delay between keystrokes
            }
        }

        private void SendKey(char keyChar)
        {
            INPUT[] inputs = new INPUT[2];

            inputs[0].type = INPUT_KEYBOARD;
            inputs[0].U.ki.wVk = 0;
            inputs[0].U.ki.wScan = (ushort)keyChar;
            inputs[0].U.ki.dwFlags = KEYEVENTF_UNICODE;
            inputs[0].U.ki.time = 0;
            inputs[0].U.ki.dwExtraInfo = IntPtr.Zero;

            inputs[1].type = INPUT_KEYBOARD;
            inputs[1].U.ki.wVk = 0;
            inputs[1].U.ki.wScan = (ushort)keyChar;
            inputs[1].U.ki.dwFlags = KEYEVENTF_UNICODE | KEYEVENTF_KEYUP;
            inputs[1].U.ki.time = 0;
            inputs[1].U.ki.dwExtraInfo = IntPtr.Zero;

            SendInput((uint)inputs.Length, inputs, Marshal.SizeOf(typeof(INPUT)));
        }

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, int vk);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool UnregisterHotKey(IntPtr hWnd, int id);

        [DllImport("user32.dll")]
        private static extern uint SendInput(uint nInputs, INPUT[] pInputs, int cbSize);

        private const uint MOD_CONTROL = 0x0002;
        private const int INPUT_KEYBOARD = 1;
        private const int KEYEVENTF_UNICODE = 0x0004;
        private const int KEYEVENTF_KEYUP = 0x0002;

        [StructLayout(LayoutKind.Sequential)]
        private struct INPUT
        {
            public int type;
            public InputUnion U;
        }

        [StructLayout(LayoutKind.Explicit)]
        private struct InputUnion
        {
            [FieldOffset(0)]
            public MOUSEINPUT mi;
            [FieldOffset(0)]
            public KEYBDINPUT ki;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct MOUSEINPUT
        {
            public int dx;
            public int dy;
            public int mouseData;
            public int dwFlags;
            public int time;
            public IntPtr dwExtraInfo;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct KEYBDINPUT
        {
            public ushort wVk;
            public ushort wScan;
            public int dwFlags;
            public int time;
            public IntPtr dwExtraInfo;
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            UnregisterHotKey(Handle, HOTKEY_ID);
            if (serialPort1.IsOpen)
            {
                serialPort1.Close();
            }
        }

        private void mouseDown(object sender, MouseEventArgs e)
        {
            mouseLocation = new Point(-e.X, -e.Y);
        }

        private void mouseMove(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                Point point = Control.MousePosition;
                point.Offset(mouseLocation.X, mouseLocation.Y);
                Location = point;
            }
        }

        private void close_window(object sender, MouseEventArgs e)
        {
            Close();
        }

        private void minimize_window(object sender, MouseEventArgs e)
        {
            WindowState = FormWindowState.Minimized;
        }

        private void reset_richTextBox(object sender, MouseEventArgs e)
        {
            richTextBox1.Clear();
        }

    }
}