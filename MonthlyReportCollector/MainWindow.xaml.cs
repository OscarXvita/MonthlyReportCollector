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

using Ookii.Dialogs.Wpf;
using System.IO;
using System.Text.RegularExpressions;

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
            var files = Directory.GetFiles(this.txt_Path.Text, "*.xls", SearchOption.TopDirectoryOnly);

            List<string> okFiles = new List<string>();
            List<string> errFiles = new List<string>();
            foreach (string filename in files)
            {

                //if (!Directory.Exists(this.txt_Path.Text + "\\整理好的报表"))//如果不存在就创建文件夹
                //{
                //    Directory.CreateDirectory(this.txt_Path.Text + "\\整理好的报表");
                //}

                var name = System.IO.Path.GetFileName(filename);
                if (Regex.IsMatch(name, @"^MSP(0?[1-9]|1[0-2])月月报-[0-9]{4}"))//@".xls|.png|.gif$"))
                {
                    okFiles.Add(name);
                    //try
                    //{
                    //    FileInfo fi = new FileInfo(filename);
                    //    string tmp = this.txt_Path.Text + "\\整理好的报表\\" + fi.Name;
                    //    if (File.Exists(tmp))
                    //    {
                    //        File.Delete(tmp);
                    //    }
                    //    fi.MoveTo(tmp);

                    //}
                    //catch (Exception)
                    //{
                    //    throw;
                    //}
                }
                else //不符合命名规则
                {
                    errFiles.Add(name);
                    if (this.txt_Path.Text.Contains("问题报表"))
                    {

                    }
                    else
                    {
                        if (!Directory.Exists(this.txt_Path.Text + "\\问题报表"))//如果不存在就创建文件夹
                        {
                            Directory.CreateDirectory(this.txt_Path.Text + "\\问题报表");
                        }
                        try
                        {

                            FileInfo fi = new FileInfo(filename);
                            string tmp = this.txt_Path.Text + "\\问题报表\\" + fi.Name;
                            if (File.Exists(tmp))
                            {
                                File.Delete(tmp);
                            }
                            fi.MoveTo(tmp);

                        }
                        catch (Exception)
                        {
                            throw;
                        }
                    }

                }
            }

            if (errFiles.Count != 0)
                MessageBox.Show("初步整理中发现了一些问题。\n为确保接下来操作的数据完整性，可能存在问题的报表放置在\"问题报表\" 文件夹下。请进行手动处理，处理后请移动到主文件夹并再次进行一次整理。\n共处理 " + (errFiles.Count + okFiles.Count) +
                    " 个月报。\n其中：\n合格文件数量 " + okFiles.Count + " 个。\n" + "可能存在问题文件数量为 " + errFiles.Count + " 个。\n", "新版报表格式检查", MessageBoxButton.OK, MessageBoxImage.Error);
            else
            {
                if (okFiles.Count == 0)
                {
                    MessageBox.Show("没有找到任何Excel文件", "Error", MessageBoxButton.OK, MessageBoxImage.Error);

                }
                else
                    MessageBox.Show("初步整理完毕！\n共处理 " + (errFiles.Count + okFiles.Count) + " 个月报。\n所有文件均通过新版报表文件名格式检查。", "新版报表格式检查", MessageBoxButton.OK, MessageBoxImage.Asterisk);

            }
            //将符合命名规则的文件放置到新建的文件夹中，不符合命名规则的文件记录并放到另一个文件夹中

            //提示操作的文件列表和统计信息
        }

        /// <summary>
        /// 对于单个文件的行移动操作
        /// </summary>
        public void MigrateToOne()
        {
            if (!Directory.Exists(this.txt_Path.Text + "\\月度总报表"))//如果不存在就创建文件夹
            {
                Directory.CreateDirectory(this.txt_Path.Text + "\\月度总报表");
            }
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
        public void PostDeadLineSubmit()
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

       

        private void btn_SelPath_Click(object sender, RoutedEventArgs e)
        {

            VistaFolderBrowserDialog dialog = new VistaFolderBrowserDialog();
            dialog.Description = "Please select a folder.";
            dialog.UseDescriptionForTitle = true; // This applies to the Vista style dialog only, not the old dialog.


            if ((bool)dialog.ShowDialog(this))
            {
                this.txt_Path.Text = dialog.SelectedPath;
                CheckAndMove();
            }
        }

        private void btn_Migrate_Click(object sender, RoutedEventArgs e)
        {
            //  CheckAndMove();
        }

        private void txt_Path_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (txt_Path.Text != "")
            {
                this.btn_Migrate.IsEnabled = true;
            }
            else
            {
                this.btn_Migrate.IsEnabled = false;
            }
        }
    }
}
