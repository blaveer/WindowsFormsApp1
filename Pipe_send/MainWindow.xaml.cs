using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Pipes;
using System.Linq;
using System.Runtime.InteropServices;
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

namespace Pipe_send
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            NamedPipeServerStream pipeServer = new NamedPipeServerStream("testpipe", PipeDirection.Out);
            pipeServer.WaitForConnection();
            try
            {
                StreamWriter sw = new StreamWriter(pipeServer);
                sw.AutoFlush = true;
                sw.WriteLine(text.Text);
                sw.Close();

            }
            catch (IOException e1)
            {
                MessageBox.Show(e1.Message);
            }
        }

        private void Text_TextChanged(object sender, TextChangedEventArgs e)
        {

        }
        //Win32 API函数
        [DllImport("User32.dll", EntryPoint = "SendMessage")]
        private static extern int SendMessage(int hWnd, int Msg, int wParam, ref COPYDATASTRUCT lParam);

        [DllImport("User32.dll", EntryPoint = "FindWindow")]
        private static extern int FindWindow(string lpClassName, string lpWindowName);

        const int WM_COPYDATA = 0x004A;
        private int num = 1;
        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            int hWnd = FindWindow(null, "ReceptForm");
            if (hWnd == 0)
            {
                MessageBox.Show("未找到消息接受者！");
            }
            else
            {
                byte[] sarr = System.Text.Encoding.Default.GetBytes(text.Text);
                int len = sarr.Length;
                COPYDATASTRUCT cds;
                cds.dwData = (IntPtr)Convert.ToInt16(num++);//可以是任意值
                cds.cbData = len + 1;//指定lpData内存区域的字节数
                cds.lpData = text.Text;//发送给目标窗口所在进程的数据
                SendMessage(hWnd, WM_COPYDATA, 0, ref cds);
            }
        }
    }
    public struct COPYDATASTRUCT
    {
        public IntPtr dwData;
        public int cbData;
        [MarshalAs(UnmanagedType.LPStr)]
        public string lpData;
    }
}
