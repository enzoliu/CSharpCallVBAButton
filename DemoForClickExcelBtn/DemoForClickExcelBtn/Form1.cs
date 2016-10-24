using System;
using System.Windows.Forms;
using System.Threading.Tasks;
using System.Threading;
using System.Runtime.InteropServices;
using System.Windows.Automation;
using System.Collections.Generic;
using System.Text;

namespace DemoForClickExcelBtn
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Task t = new Task(() =>
            {
                Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel.Workbook wb = xlApp.Workbooks.Open("Demo.xlsm");
                ThreadPool.QueueUserWorkItem(delegate { xlApp.Run("btn1_Click"); });

                Thread.Sleep(3000);

                IntPtr winProcess = new IntPtr();
                
                // ThunderDFrame不要動，那是指透過VBA產生的Frame
                // 把UserForm1替換成被按鈕呼叫出來的form的名稱
                winProcess = FindWindow("ThunderDFrame", "UserForm1");

                // 以下將搜尋UserForm1內的button，因為我只有放一個按鈕，這邊再自己看情況調整
                // 組件參考的部分要記得新增
                AutomationElement element = AutomationElement.FromHandle(winProcess);
                AutomationElementCollection elements = element.FindAll(TreeScope.Descendants, Condition.TrueCondition);
                foreach (AutomationElement elementNode in elements)
                {
                    // 把按鈕的顯示名稱，替換掉Test
                    if (elementNode.Current.Name == "Test")
                    {
                        this.AppendTxt("element process id : " + elementNode.Current.ProcessId);
                        // 下面兩行是用來喚醒click行為用的
                        var invokePattern = elementNode.GetCurrentPattern(InvokePattern.Pattern) as InvokePattern;
                        invokePattern.Invoke();
                    }
                }
            });
            t.Start();
        }

        private void AppendTxt(String text)
        {
            this.Invoke((MethodInvoker)delegate
            {
                this.textBox1.AppendText(text + "\r\n");
            });
        }

        #region 外部參考
        public const uint WM_SETTEXT = 0x0c;
        public const uint WM_KEYDOWN = 0x100;
        public const uint WM_KEYUP = 0x0101;
        public const String KEY_ENTER = "0xD";

        [DllImport("User32.dll", EntryPoint = "FindWindow")]
        private static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

        [DllImport("user32.dll")]
        public static extern long SendMessage(IntPtr hWnd, uint msg, uint wparam, string text);

        [DllImport("user32.dll")]
        public static extern IntPtr PostMessage(IntPtr hWnd, uint Msg, int wParam, uint lParam);

        #endregion
        
        private void button2_Click(object sender, EventArgs e)
        {
            Task t = new Task(() =>
            {
                Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel.Workbook wb = xlApp.Workbooks.Open(Environment.CurrentDirectory + "\\Demo3.xlsm");
                ThreadPool.QueueUserWorkItem(delegate { xlApp.Run("btn1_Click"); });

                Thread.Sleep(3000);

                IntPtr winProcess = new IntPtr();
                IntPtr editHwnd = new IntPtr();

                // #32770 = Dialog class
                winProcess = FindWindow("#32770", "開啟舊檔");

                // Edit control handler (從winProcess的子視窗中尋找)
                editHwnd = this.FindChildHandler(winProcess, "Edit");

                if (editHwnd != IntPtr.Zero)
                {
                    // Send string to edit handler
                    SendMessage(editHwnd, WM_SETTEXT, 0, "kc.txt");

                    // Raise Enter event in edit handler
                    PostMessage(editHwnd, WM_KEYDOWN, Convert.ToInt32(KEY_ENTER, 16), 0);
                    PostMessage(editHwnd, WM_KEYUP, Convert.ToInt32(KEY_ENTER, 16), 0);
                }
            });
            t.Start();
        }
        
        private IntPtr FindChildHandler(IntPtr MainProcess, String ClassName)
        {
            var allChildWindows = new EnumerateWindowChild(MainProcess).GetAllChildHandles();
            foreach (IntPtr child in allChildWindows)
            {
                StringBuilder sb = new StringBuilder(256);
                int cRef = GetClassName(child, sb, sb.Capacity);
                if (cRef != 0)
                {
                    if (sb.ToString().Equals(ClassName))
                    {
                        return child;
                    }
                }
            }
            return IntPtr.Zero;
        }
    }
}
