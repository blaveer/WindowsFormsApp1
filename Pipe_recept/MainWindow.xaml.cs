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
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace Pipe_recept
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            this.Title = "ReceptForm";
        }



        private void Button_Click(object sender, RoutedEventArgs e)
        {
            NamedPipeClientStream pipeClient = new NamedPipeClientStream(".", "testpipe", PipeDirection.In);
            StreamReader sr = null;
            try
            {
                pipeClient.Connect();
                sr = new StreamReader(pipeClient);

                text.Text = sr.ReadToEnd();
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show(ex.Message);
                this.Close();
            }
        }
        protected override void OnSourceInitialized(EventArgs e)

        {

            base.OnSourceInitialized(e);

            HwndSource hwndSource = PresentationSource.FromVisual(this) as HwndSource;

            if (hwndSource != null)

            {

                IntPtr handle = hwndSource.Handle;

                hwndSource.AddHook(new HwndSourceHook(WndProc));

            }

        }
        const int WM_COPYDATA = 0x004A;
        IntPtr WndProc(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, ref bool handled)

        {

            if (msg == WM_COPYDATA)

            {

                COPYDATASTRUCT cds = (COPYDATASTRUCT)Marshal.PtrToStructure(lParam, typeof(COPYDATASTRUCT)); // 接收封装的消息

                string rece = cds.lpData; // 获取消息内容

                // 自定义行为

                text.Text += rece + "\r\n";

            }
            return hwnd;
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
