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

namespace MonthlyReportCollector
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

        /// <summary>
        /// 检查文件夹下所有文件是否符合命名规则，并分类处理
        /// </summary>
        public void CheckAndMove()
        {
            //遍历文件夹的文件名

            //将符合命名规则的文件放置到新建的文件夹中，不符合命名规则的文件记录并放到另一个文件夹中

            //提示操作的文件列表和统计信息
        }
        /// <summary>
        /// 对于单个文件的行移动操作
        /// </summary>
        public void MigrateToOne()
        {

            //创建月度总表
            //校验表格合法性，不合法抛异常
            //获取数据行，作为对象加入List
            //加到总表内
            //异常
            //记录异常信息到LOG List
        }
        public void DoAllMigrate()
        {
            //清空LOG List
            //检查总表是否存在
            //存在即删除重建，不存在就建立
            //Foreach 遍历文件
            //对每一个文件执行Migrate

            //显示统计和错误信息
        }
        public void PostDeadLineSubmit( )
        {
            //导入上次总表信息，及补交表格信息
            //MigrateToOne()
        }
        /// <summary>
        /// 检查本月漏交或新增名单，需要上月总表格做对比，导出漏交名单
        /// </summary>
        public void CheckMissingOrNewList()
        {
            //


        }
        /// <summary>
        /// 提醒漏交一个月的同学
        /// </summary>
        public void RemindAllForget()
        {
            
        }
        /// <summary>
        /// 提醒连续两个月漏交报告的同学续交报告
        /// </summary>
        public void RemindDanger()
        {

        }


    }
}
