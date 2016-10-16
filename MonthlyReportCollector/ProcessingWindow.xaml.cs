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
using System.Windows.Shapes;
using System.Windows.Threading;

namespace MonthlyReportCollector
{
    /// <summary>
    /// ProcessingWindow.xaml 的交互逻辑
    /// </summary>
    public partial class ProcessingWindow : Window
    {
        public ProcessingWindow()
        {

            InitializeComponent();
            this.Topmost = true;
            this.Loaded += new RoutedEventHandler(MainWin_Loaded);
        }
        private DateTime StartTime;
        private void MainWin_Loaded(object sender, RoutedEventArgs e)
        {

            //设置定时器
           DispatcherTimer  timer = new System.Windows.Threading.DispatcherTimer();
            timer.Interval = new TimeSpan(0, 0, 0,0,250);   //间隔1秒
            timer.Tick += new EventHandler(timer_Tick);
            timer.Start();
            StartTime = DateTime.Now;

        }
        private void timer_Tick(object sender, EventArgs e)
        {
           
            if((DateTime.Now.Ticks-StartTime.Ticks)>30000000)
            {
                processText.Text = "处理中...\n如程序提示报错，请点击隐藏此窗口，并根据报错提示修复错误。";
                hide.Visibility = Visibility.Visible;
            }

        }

        private void hide_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
    }
}
