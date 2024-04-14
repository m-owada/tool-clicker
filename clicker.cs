using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;

[assembly: AssemblyVersion("1.0.0.0")]
[assembly: AssemblyFileVersion ("1.0.0.0")]
[assembly: AssemblyInformationalVersion("1.0")]
[assembly: AssemblyTitle("")]
[assembly: AssemblyDescription("")]
[assembly: AssemblyProduct("Clicker")]
[assembly: AssemblyCompany("")]
[assembly: AssemblyCopyright("Copyright (C) 2024 m-owada.")]
[assembly: AssemblyTrademark("")]
[assembly: AssemblyCulture("")]

class Program
{
    [STAThread]
    static void Main()
    {
        try
        {
            Application.EnableVisualStyles();
            Application.ThreadException += (sender, e) =>
            {
                throw new Exception(e.Exception.Message);
            };
            Application.Run(new MainForm());
        }
        catch(Exception e)
        {
            MessageBox.Show(e.Message, e.Source, MessageBoxButtons.OK, MessageBoxIcon.Error);
            Application.Exit();
        }
    }
}

class MainForm : Form
{
    private ComboBox comboBox1 = new ComboBox();
    private ComboBox comboBox2 = new ComboBox();
    private CheckBox checkBox1 = new CheckBox();
    private NumericUpDown numericUpDown1 = new NumericUpDown();
    private Button button1 = new Button();
    private ToolStripStatusLabel toolStripStatusLabel1 = new ToolStripStatusLabel();
    private ToolStripStatusLabel toolStripStatusLabel2 = new ToolStripStatusLabel();
    private CustomTimer timer = new CustomTimer();
    private MouseClick mouseClick = new MouseClick();
    private MouseHook mouseHook = new MouseHook();
    private KeyHook keyHook = new KeyHook();
    private long clickCount = 0;
    
    public MainForm()
    {
        // 設定ファイル
        var config = new Config();
        
        // フォーム
        this.Text = "Clicker";
        this.Location = new Point(config.X, config.Y);
        this.Size = new Size(480, 90);
        this.MaximizeBox = false;
        this.MinimizeBox = true;
        this.MinimumSize = this.Size;
        this.StartPosition = FormStartPosition.Manual;
        this.FormBorderStyle = FormBorderStyle.FixedSingle;
        this.Icon = Icon.ExtractAssociatedIcon(Process.GetCurrentProcess().MainModule.FileName);
        this.FormClosing += Form_Closing;
        
        // ラベル
        this.Controls.Add(CreateLabel(10, 10, "操作"));
        
        // コンボボックス1
        comboBox1.Location = new Point(45, 10);
        comboBox1.Size = new Size(80, 20);
        comboBox1.DropDownStyle = ComboBoxStyle.DropDownList;
        comboBox1.Items.AddRange(new string[] {"左クリック", "右クリック", "中央クリック"});
        comboBox1.SelectedIndex = config.Operation;
        this.Controls.Add(comboBox1);
        
        // ラベル
        this.Controls.Add(CreateLabel(135, 10, "間隔"));
        
        // 数値入力1
        numericUpDown1.Location = new Point(170, 11);
        numericUpDown1.Size = new Size(70, 20);
        numericUpDown1.Maximum = 9999999;
        numericUpDown1.Minimum = 1;
        numericUpDown1.Value = (numericUpDown1.Minimum <= config.Interval && config.Interval <= numericUpDown1.Maximum) ? config.Interval : 1000;
        numericUpDown1.ThousandsSeparator = true;
        numericUpDown1.TextAlign = HorizontalAlignment.Right;
        numericUpDown1.Enter += Enter_numericUpDown1;
        numericUpDown1.Leave += Leave_numericUpDown1;
        numericUpDown1.ValueChanged += ValueChanged_numericUpDown1;
        this.Controls.Add(numericUpDown1);
        
        // ラベル
        this.Controls.Add(CreateLabel(245, 10, "ms"));
        
        // ラベル
        this.Controls.Add(CreateLabel(270, 10, "中止"));
        
        // コンボボックス2
        comboBox2.Location = new Point(305, 10);
        comboBox2.Size = new Size(50, 20);
        comboBox2.DropDownStyle = ComboBoxStyle.DropDownList;
        comboBox2.Items.AddRange(new string[] {"Esc", "F1", "F2", "F3", "F4", "F5", "F6", "F7", "F8", "F9", "F10", "F11", "F12"});
        comboBox2.SelectedIndex = config.Cancel;
        this.Controls.Add(comboBox2);
        
        // チェックボックス1
        checkBox1.Location = new Point(365, 12);
        checkBox1.Size = new Size(65, 20);
        checkBox1.Text = "最前面";
        checkBox1.AutoSize = true;
        checkBox1.CheckedChanged += CheckedChanged_checkBox1;
        checkBox1.Checked = config.TopMost;
        this.Controls.Add(checkBox1);
        
        // ボタン1
        button1.Location = new Point(425, 10);
        button1.Size = new Size(40, 20);
        button1.Text = "開始";
        button1.Click += Click_button1;
        this.Controls.Add(button1);
        
        // ステータスバー
        var statusStrip1 = new StatusStrip();
        statusStrip1.SizingGrip = false;
        toolStripStatusLabel1.AutoSize = false;
        toolStripStatusLabel1.BorderSides = ToolStripStatusLabelBorderSides.Right;
        toolStripStatusLabel1.BorderStyle = Border3DStyle.Flat;
        toolStripStatusLabel1.Size = new System.Drawing.Size(400, 20);
        toolStripStatusLabel1.TextAlign = ContentAlignment.MiddleLeft;
        toolStripStatusLabel2.AutoSize = false;
        toolStripStatusLabel2.Size = new System.Drawing.Size(75, 20);
        toolStripStatusLabel2.TextAlign = ContentAlignment.MiddleCenter;
        statusStrip1.Items.Add(toolStripStatusLabel1);
        statusStrip1.Items.Add(toolStripStatusLabel2);
        toolStripStatusLabel1.Text = string.Empty;
        toolStripStatusLabel2.Text = string.Empty;
        SetToolStripStatusLabel2();
        this.Controls.Add(statusStrip1);
        
        // タイマー
        timer.Tick += Tick_timer;
        TimerStop();
        
        // マウスフック
        mouseHook.MouseClickEvent += Click_mouseHook;
        
        // キーボードフック
        keyHook.KeyDownEvent += Down_keyHook;
    }
    
    private Label CreateLabel(int x, int y, string text)
    {
        var label = new Label();
        label.Location = new Point(x, y);
        label.Text = text;
        label.Size = new Size(label.PreferredWidth, 20);
        label.TextAlign = ContentAlignment.MiddleLeft;
        return label;
    }
    
    private void Form_Closing(object sender, FormClosingEventArgs e)
    {
        var config = new Config();
        if(this.WindowState == FormWindowState.Normal)
        {
            config.X = this.Location.X;
            config.Y = this.Location.Y;
        }
        config.Operation = comboBox1.SelectedIndex;
        config.Interval = numericUpDown1.Value;
        config.Cancel = comboBox2.SelectedIndex;
        config.TopMost = checkBox1.Checked;
        config.Save();
        timer.Dispose();
        mouseHook.Dispose();
        keyHook.Dispose();
    }
    
    private void Enter_numericUpDown1(object sender, EventArgs e)
    {
        numericUpDown1.Select(0, numericUpDown1.Text.Length);
    }
    
    private void Leave_numericUpDown1(object sender, EventArgs e)
    {
        numericUpDown1.Text = numericUpDown1.Value.ToString();
    }
    
    private void ValueChanged_numericUpDown1(object sender, EventArgs e)
    {
        SetToolStripStatusLabel2();
    }
    
    private void CheckedChanged_checkBox1(object sender, EventArgs e)
    {
        this.TopMost = checkBox1.Checked;
    }
    
    private void Click_button1(object sender, EventArgs e)
    {
        AllEnabled(false);
        toolStripStatusLabel1.Text = comboBox1.Text + "で実行します。中止する場合は" + comboBox2.Text + "キーを押してください。";
        mouseHook.Start();
        keyHook.Start();
    }
    
    private void Tick_timer(object sender, EventArgs e)
    {
        if(timer.Enabled)
        {
            switch(comboBox1.Text)
            {
                case "左クリック":
                    mouseClick.Left();
                    break;
                case "右クリック":
                    mouseClick.Right();
                    break;
                case "中央クリック":
                    mouseClick.Middle();
                    break;
            }
            clickCount++;
            SetToolStripStatusLabel1();
        }
    }
    
    private void Click_mouseHook(object sender, MouseEventArgs e)
    {
        if(IsMouseClick(e.Button) && !button1.Enabled && !timer.Enabled)
        {
            TimerStart();
        }
    }
    
    private void Down_keyHook(object sender, KeyHook.KeyEventArgs e)
    {
        if(IsKeyDown(e.KeyCode))
        {
            TimerStop();
        }
    }
    
    private void SetToolStripStatusLabel1()
    {
        if(timer.Enabled)
        {
            toolStripStatusLabel1.Text = "クリック中..." + clickCount.ToString("N0");
        }
    }
    
    private void SetToolStripStatusLabel2()
    {
        toolStripStatusLabel2.Text = TimeSpan.FromMilliseconds(Decimal.ToDouble(numericUpDown1.Value)).ToString(@"hh\:mm\:ss\.fff");
    }
    
    private void TimerStart()
    {
        timer.Start(decimal.ToDouble(numericUpDown1.Value));
        clickCount = 0;
        SetToolStripStatusLabel1();
    }
    
    private void TimerStop()
    {
        timer.Stop();
        mouseHook.Stop();
        keyHook.Stop();
        AllEnabled(true);
        toolStripStatusLabel1.Text = "操作内容を設定して" + button1.Text + "ボタンをクリックしてください。";
    }
    
    private bool IsMouseClick(MouseButtons button)
    {
        return ((comboBox1.Text == "左クリック"   && button == MouseButtons.Left  ) ||
                (comboBox1.Text == "右クリック"   && button == MouseButtons.Right ) ||
                (comboBox1.Text == "中央クリック" && button == MouseButtons.Middle));
    }
    
    private bool IsKeyDown(int keyCode)
    {
        return ((comboBox2.Text == "Esc" && keyCode ==  27) ||
                (comboBox2.Text == "F1"  && keyCode == 112) ||
                (comboBox2.Text == "F2"  && keyCode == 113) ||
                (comboBox2.Text == "F3"  && keyCode == 114) ||
                (comboBox2.Text == "F4"  && keyCode == 115) ||
                (comboBox2.Text == "F5"  && keyCode == 116) ||
                (comboBox2.Text == "F6"  && keyCode == 117) ||
                (comboBox2.Text == "F7"  && keyCode == 118) ||
                (comboBox2.Text == "F8"  && keyCode == 119) ||
                (comboBox2.Text == "F9"  && keyCode == 120) ||
                (comboBox2.Text == "F10" && keyCode == 121) ||
                (comboBox2.Text == "F11" && keyCode == 122) ||
                (comboBox2.Text == "F12" && keyCode == 123));
    }
    
    private void AllEnabled(bool enabled)
    {
        comboBox1.Enabled = enabled;
        numericUpDown1.Enabled = enabled;
        comboBox2.Enabled = enabled;
        checkBox1.Enabled = enabled;
        button1.Enabled = enabled;
    }
}

class Config
{
    public readonly string FileName = "config.xml";
    public int X { get; set; }
    public int Y { get; set; }
    public int Operation { get; set; }
    public decimal Interval { get; set; }
    public int Cancel { get; set; }
    public bool TopMost { get; set; }
    
    public Config()
    {
        Load();
    }
    
    public void Load()
    {
        var xml = GetDocument().Element("config");
        X = GetValue(xml, "x", 10);
        Y = GetValue(xml, "y", 10);
        Operation = GetValue(xml, "operation", 0);
        Interval = GetValue(xml, "interval", 1000.0M);
        Cancel = GetValue(xml, "cancel", 0);
        TopMost = GetValue(xml, "topmost", true);
    }
    
    public void Save()
    {
        var xml = GetDocument().Element("config");
        SetValue(xml, "x", X);
        SetValue(xml, "y", Y);
        SetValue(xml, "operation", Operation);
        SetValue(xml, "interval", Interval);
        SetValue(xml, "cancel", Cancel);
        SetValue(xml, "topmost", TopMost);
        xml.Save(FileName);
    }
    
    private XDocument GetDocument()
    {
        if(File.Exists(FileName))
        {
            return XDocument.Load(FileName);
        }
        else
        {
            var xml = new XDocument(
                new XDeclaration("1.0", "utf-8", string.Empty),
                new XElement("config",
                    new XElement("x", "10"),
                    new XElement("y", "10"),
                    new XElement("operation", "0"),
                    new XElement("interval", "1000.0"),
                    new XElement("cancel", "0"),
                    new XElement("topmost", "True")
                )
            );
            xml.Save(FileName);
            return xml;
        }
    }
    
    private int GetValue(XElement node, string name, int init)
    {
        var val = init;
        if(node.Element(name) != null)
        {
            Int32.TryParse(node.Element(name).Value, out val);
        }
        return val;
    }
    
    private decimal GetValue(XElement node, string name, decimal init)
    {
        var val = init;
        if(node.Element(name) != null)
        {
            Decimal.TryParse(node.Element(name).Value, out val);
        }
        return val;
    }
    
    private bool GetValue(XElement node, string name, bool init)
    {
        var val = init;
        if(node.Element(name) != null)
        {
            Boolean.TryParse(node.Element(name).Value, out val);
        }
        return val;
    }
    
    private void SetValue(XElement node, string name, string val)
    {
        if(node.Element(name) == null)
        {
            node.Add(new XElement(name, val));
        }
        else
        {
            node.Element(name).Value = val;
        }
    }
    
    private void SetValue(XElement node, string name, int val)
    {
        SetValue(node, name, val.ToString());
    }
    
    private void SetValue(XElement node, string name, decimal val)
    {
        SetValue(node, name, val.ToString());
    }
    
    private void SetValue(XElement node, string name, bool val)
    {
        SetValue(node, name, val.ToString());
    }
}

class CustomTimer : IDisposable
{
    public bool Enabled { get; private set; }
    public event EventHandler Tick;
    
    public CustomTimer()
    {
        Stop();
    }
    
    public void Start(double interval)
    {
        Enabled = true;
        var nextAt = DateTime.Now.AddMilliseconds(interval);
        Task.Run(() =>
        {
            while(Enabled)
            {
                while(true)
                {
                    var rest = (nextAt - DateTime.Now).TotalMilliseconds;
                    if(rest > 16)
                    {
                        Thread.Sleep((int)(rest - 16));
                    }
                    else if(rest > 0)
                    {
                        Thread.SpinWait(50);
                    }
                    else
                    {
                        break;
                    }
                }
                OnTickEvent();
                nextAt = nextAt.AddMilliseconds(interval);
            }
        });
    }
    
    protected void OnTickEvent()
    {
        if(Tick != null)
        {
            Tick(this, new EventArgs());
        }
    }
    
    public void Stop()
    {
        Enabled = false;
    }
    
    public void Dispose()
    {
        Stop();
    }
}

class MouseClick
{
    [DllImport("user32.dll")]
    private static extern void SendInput(int nInputs, ref Input pInputs, int cbsize);
    
    [DllImport("user32.dll")]
    private static extern IntPtr GetMessageExtraInfo();
    
    [DllImport("user32.dll")]
    private static extern bool SetCursorPos(int x, int y);
    
    [StructLayout(LayoutKind.Sequential)]
    private struct Input
    {
        public int Type;
        public InputUnion ui;
    }
    
    [StructLayout(LayoutKind.Explicit)]
    private struct InputUnion
    {
        [FieldOffset(0)]
        public MouseInput Mouse;
        [FieldOffset(0)]
        public KeyboardInput Keyboard;
        [FieldOffset(0)]
        public HardwareInput Hardware;
    }
    
    [StructLayout(LayoutKind.Sequential)]
    private struct MouseInput
    {
        public int X;
        public int Y;
        public int Data;
        public int Flags;
        public int Time;
        public IntPtr ExtraInfo;
    }
    
    [StructLayout(LayoutKind.Sequential)]
    private struct KeyboardInput
    {
        public short VirtualKey;
        public short ScanCode;
        public int Flags;
        public int Time;
        public IntPtr ExtraInfo;
    }
    
    [StructLayout(LayoutKind.Sequential)]
    private struct HardwareInput
    {
        public int uMsg;
        public short wParamL;
        public short wParamH;
    }
    
    public void Left()
    {
        LeftDown();
        LeftUp();
    }
    
    public void LeftDown()
    {
        MouseSendInput(0x0002);
    }
    
    public void LeftUp()
    {
        MouseSendInput(0x0004);
    }
    
    public void Right()
    {
        RightDown();
        RightUp();
    }
    
    public void RightDown()
    {
        MouseSendInput(0x0008);
    }
    
    public void RightUp()
    {
        MouseSendInput(0x0010);
    }
    
    public void Middle()
    {
        MiddleDown();
        MiddleUp();
    }
    
    public void MiddleDown()
    {
        MouseSendInput(0x0020);
    }
    
    public void MiddleUp()
    {
        MouseSendInput(0x0040);
    }
    
    private void MouseSendInput(int flags)
    {
        var input = new Input();
        input.Type = 0;
        input.ui.Mouse.Flags = flags;
        input.ui.Mouse.Data = 0;
        input.ui.Mouse.X = 0;
        input.ui.Mouse.Y = 0;
        input.ui.Mouse.Time = 0;
        input.ui.Mouse.ExtraInfo = GetMessageExtraInfo();
        SendInput(1, ref input, Marshal.SizeOf(input));
    }
}

class MouseHook : IDisposable
{
    [DllImport("user32.dll")]
    private static extern IntPtr SetWindowsHookEx(int idHook, MouseHookCallback lpfn, IntPtr hMod, uint dwThreadId);
    
    [DllImport("user32.dll")]
    private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, uint msg, ref MSLLHOOKSTRUCT msllhookstruct);
    
    [DllImport("user32.dll")]
    private static extern bool UnhookWindowsHookEx(IntPtr hhk);
    
    [DllImport("kernel32.dll")]
    private static extern IntPtr GetModuleHandle(string lpModuleName);
    
    [StructLayout(LayoutKind.Sequential)]
    private struct POINT
    {
        public int x;
        public int y;
    }
    
    [StructLayout(LayoutKind.Sequential)]
    private struct MSLLHOOKSTRUCT
    {
        public POINT pt;
        public uint mouseData;
        public uint flags;
        public uint time;
        public IntPtr dwExtraInfo;
    }
    
    private delegate IntPtr MouseHookCallback(int nCode, uint msg, ref MSLLHOOKSTRUCT msllhookstruct);
    private MouseHookCallback mouseHookProc = null;
    
    private IntPtr hookPtr = IntPtr.Zero;
    
    public event MouseEventHandler MouseClickEvent;
    
    public void Start()
    {
        using(var curProcess = Process.GetCurrentProcess())
        {
            using(var curModule = curProcess.MainModule)
            {
                Stop();
                mouseHookProc = MouseHookProc;
                hookPtr = SetWindowsHookEx(14, mouseHookProc, GetModuleHandle(curModule.ModuleName), 0);
            }
        }
    }
    
    private IntPtr MouseHookProc(int nCode, uint msg, ref MSLLHOOKSTRUCT s)
    {
        if(nCode >= 0)
        {
            switch(msg)
            {
                case 0x0201:
                    OnMouseClickEvent(MouseButtons.Left, 1, s.pt.x, s.pt.y, 0);
                    break;
                case 0x0204:
                    OnMouseClickEvent(MouseButtons.Right, 1, s.pt.x, s.pt.y, 0);
                    break;
                case 0x0207:
                    OnMouseClickEvent(MouseButtons.Middle, 1, s.pt.x, s.pt.y, 0);
                    break;
            }
        }
        return CallNextHookEx(hookPtr, nCode, msg, ref s);
    }
    
    protected void OnMouseClickEvent(MouseButtons button, int clicks, int x, int y, int delta)
    {
        if(MouseClickEvent != null)
        {
            MouseClickEvent(this, new MouseEventArgs(button, clicks, x, y, delta));
        }
    }
    
    public void Stop()
    {
        Dispose();
    }
    
    public void Dispose()
    {
        if(hookPtr != IntPtr.Zero)
        {
            UnhookWindowsHookEx(hookPtr);
            mouseHookProc = null;
            hookPtr = IntPtr.Zero;
        }
    }
}

class KeyHook : IDisposable
{
    [DllImport("user32.dll")]
    private static extern IntPtr SetWindowsHookEx(int idHook, KeyHookCallback lpfn, IntPtr hMod, uint dwThreadId);
    
    [DllImport("user32.dll")]
    private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, uint msg, ref KBDLLHOOKSTRUCT kbdllhookstruct);
    
    [DllImport("user32.dll")]
    private static extern bool UnhookWindowsHookEx(IntPtr hhk);
    
    [DllImport("kernel32.dll")]
    private static extern IntPtr GetModuleHandle(string lpModuleName);
    
    [Flags]
    private enum KBDLLHOOKSTRUCTFlags : uint
    {
        LLKHF_EXTENDED = 0x01,
        LLKHF_INJECTED = 0x10,
        LLKHF_ALTDOWN = 0x20,
        LLKHF_UP = 0x80,
    }
    
    [StructLayout(LayoutKind.Sequential)]
    private struct KBDLLHOOKSTRUCT
    {
        public uint vkCode;
        public uint scanCode;
        public KBDLLHOOKSTRUCTFlags flags;
        public uint time;
        public UIntPtr dwExtraInfo;
    }
    
    private delegate IntPtr KeyHookCallback(int nCode, uint msg, ref KBDLLHOOKSTRUCT kbdllhookstruct);
    private KeyHookCallback keyHookProc = null;
    
    private IntPtr hookPtr = IntPtr.Zero;
    
    public delegate void KeyEventHandler(object sender, KeyEventArgs e);
    public event KeyEventHandler KeyDownEvent;
    public event KeyEventHandler KeyUpEvent;
    
    public void Start()
    {
        using(var curProcess = Process.GetCurrentProcess())
        {
            using(var curModule = curProcess.MainModule)
            {
                Stop();
                keyHookProc = KeyHookProc;
                hookPtr = SetWindowsHookEx(13, keyHookProc, GetModuleHandle(curModule.ModuleName), 0);
            }
        }
    }
    
    private IntPtr KeyHookProc(int nCode, uint msg, ref KBDLLHOOKSTRUCT s)
    {
        if(nCode >= 0)
        {
            switch(msg)
            {
                case 0x0100:
                case 0x0104:
                    OnKeyDownEvent((int)s.vkCode);
                    break;
                case 0x0101:
                case 0x0105:
                    OnKeyUpEvent((int)s.vkCode);
                    break;
            }
        }
        return CallNextHookEx(hookPtr, nCode, msg, ref s);
    }
    
    protected void OnKeyDownEvent(int keyCode)
    {
        if(KeyDownEvent != null)
        {
            KeyDownEvent(this, new KeyEventArgs(keyCode));
        }
    }
    
    protected void OnKeyUpEvent(int keyCode)
    {
        if(KeyUpEvent != null)
        {
            KeyUpEvent(this, new KeyEventArgs(keyCode));
        }
    }
    
    public void Stop()
    {
        Dispose();
    }
    
    public void Dispose()
    {
        if(hookPtr != IntPtr.Zero)
        {
            UnhookWindowsHookEx(hookPtr);
            keyHookProc = null;
            hookPtr = IntPtr.Zero;
        }
    }
    
    public class KeyEventArgs : EventArgs
    {
        public int KeyCode { get; private set; }
        public KeyEventArgs(int keyCode)
        {
            KeyCode = keyCode;
        }
    }
}
