using System;
using System.Collections.Generic;
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

namespace Event_handing
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
        // 定义事件发起函数
        void ExtinguishFire(object sender, FireEventArgs fe)
        {

            String notice = String.Format("The ExtinguishFire function was called by {0}.", sender.ToString());
            fire_box.Items.Add(notice);

            if (fe.ferocity < 2)
            {
                String respond = String.Format("发生在{0} 的火情不大，快打点水", fe.room);
                fire_box.Items.Add(respond);
            }
            else if (fe.ferocity < 5)
            {
                String respond = String.Format("我在用灭火器扑灭{0}的火.", fe.room);
                fire_box.Items.Add(respond);
            }
            else
            {
                String respond = String.Format("在 {0} 的火情无法控制，已经打了119", fe.room);
                fire_box.Items.Add(respond);
            }

        }
        private void Button_Click(object sender, RoutedEventArgs e)
        {
            FireAlarm fireAlarm = new FireAlarm();
            fireAlarm.FireEvent += ExtinguishFire;
            fireAlarm.ActivateFireAlarm("客厅", 3);
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            FireAlarm fireAlarm = new FireAlarm();
            fireAlarm.FireEvent += ExtinguishFire;
            fireAlarm.ActivateFireAlarm("厨房", 10);
        }
    }
}
