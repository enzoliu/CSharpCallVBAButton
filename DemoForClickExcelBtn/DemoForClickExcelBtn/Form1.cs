using System;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using System.Threading.Tasks;
using System.Threading;
using System.Runtime.InteropServices;
using System.Windows.Automation;

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
                Workbook wb = xlApp.Workbooks.Open("Demo.xlsm");
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
        [DllImport("User32.dll", EntryPoint = "FindWindow")]
        private static extern IntPtr FindWindow(string lpClassName, string lpWindowName);
        #endregion
    }
}
