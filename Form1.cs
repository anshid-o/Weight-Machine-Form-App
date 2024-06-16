using System;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace WeightGetter
{
    public partial class myForm : Form
    {
        private string sourceFileName = @"C:\Users\91964\OneDrive\Desktop\sendData.txt"; // Update with your source file path
        private string lastLine = string.Empty;
        private FileSystemWatcher fileWatcher;
        private long lastPosition = 0;
        private const int WM_HOTKEY = 0x0312;
        private const int HOTKEY_ID = 9000;

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

        public Point mouseLocation;
        public myForm()
        {
            InitializeComponent();
            this.FormBorderStyle = FormBorderStyle.None;
            Region = System.Drawing.Region.FromHrgn(CreateRoundRectRgn(0, 0, Width, Height, 20, 20));
            InitializeRichTextBox();
            LoadInitialData();
            RegisterHotKey(Handle, HOTKEY_ID, MOD_CONTROL, (int)Keys.Q);
            StartFileWatcher();
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

        private void LoadInitialData()
        {
            try
            {
                string[] lines = File.ReadAllLines(sourceFileName);
                richTextBox1.Lines = lines;
                lastLine = lines.Length > 0 ? lines[lines.Length - 1] : string.Empty;

                // Set lastPosition to the length of the file
                lastPosition = new FileInfo(sourceFileName).Length;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading data: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void StartFileWatcher()
        {
            fileWatcher = new FileSystemWatcher(Path.GetDirectoryName(sourceFileName))
            {
                Filter = Path.GetFileName(sourceFileName),
                NotifyFilter = NotifyFilters.LastWrite | NotifyFilters.Size
            };
            fileWatcher.Changed += OnFileChanged;
            fileWatcher.EnableRaisingEvents = true;
        }

        private void OnFileChanged(object sender, FileSystemEventArgs e)
        {
            try
            {
                // Introduce a delay to allow the file write operation to complete
                System.Threading.Thread.Sleep(100);

                using (FileStream fs = new FileStream(sourceFileName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                using (StreamReader sr = new StreamReader(fs))
                {
                    // Move the stream position to the last read position
                    fs.Seek(lastPosition, SeekOrigin.Begin);

                    // Read new lines from the last position
                    string newLine;
                    bool isFirstLine = true;
                    while ((newLine = sr.ReadLine()) != null)
                    {
                        lastLine = newLine;

                        // Append the new line to the RichTextBox
                        richTextBox1.Invoke((MethodInvoker)(() =>
                        {
                            if (!isFirstLine)
                            {
                                richTextBox1.AppendText(Environment.NewLine);
                            }
                            richTextBox1.AppendText(lastLine);
                            isFirstLine = false;
                        }));
                    }

                    // Update the last read position
                    lastPosition = fs.Position;
                }
            }
            catch (IOException ioEx)
            {
                Console.WriteLine($"IOException in OnFileChanged: {ioEx}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Exception in OnFileChanged: {ex}");
            }
        }

        protected override void WndProc(ref Message m)
        {
            base.WndProc(ref m);
            if (m.Msg == WM_HOTKEY && m.WParam.ToInt32() == HOTKEY_ID)
            {
                SendTextToActiveWindow(lastLine);
            }
        }

        private void SendTextToActiveWindow(string text)
        {
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
            [FieldOffset(0)]
            public HARDWAREINPUT hi;
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

        [StructLayout(LayoutKind.Sequential)]
        private struct HARDWAREINPUT
        {
            public int uMsg;
            public short wParamL;
            public short wParamH;
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            UnregisterHotKey(Handle, HOTKEY_ID);
            fileWatcher.Dispose();
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

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void richTextBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void panel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void panel2_Paint(object sender, PaintEventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void pictureBox2_Click(object sender, EventArgs e)
        {

        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {

        }

        private void reset_richTextBox(object sender, MouseEventArgs e)
        {
            richTextBox1.Clear();
        }
    }
}
