/*******************************************************************************
 * You may amend and distribute as you like, but don't remove this header!
 * 
 * All rights reserved.
 * 
 * EPPlus is an Open Source project provided under the 
 * GNU General Public License (GPL) as published by the 
 * Free Software Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA 02111-1307 USA
 * 
 * EPPlus provides server-side generation of Excel 2007 spreadsheets.
 * See http://www.codeplex.com/EPPlus for details.
 *
 *
 * 
 * The GNU General Public License can be viewed at http://www.opensource.org/licenses/gpl-license.php
 * If you unfamiliar with this license or have questions about it, here is an http://www.gnu.org/licenses/gpl-faq.html
 * 
 * The code for this project may be used and redistributed by any means PROVIDING it is 
 * not sold for profit without the author's written consent, and providing that this notice 
 * and the author's name and all copyright notices remain intact.
 * 
 * All code and executables are provided "as is" with no warranty either express or implied. 
 * The author accepts no liability for any damage or loss of business that this product may cause.
 *
 *
 * Code change notes:
 *  Author							Change						Date
 *******************************************************************************
 * Oscar		                     Added		             16-OCT-2016
 * 
 *******************************************************************************/

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
using System.Runtime.InteropServices;
using OfficeOpenXml;
using System.Diagnostics;

namespace MonthlyReportCollector
{

    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        [DllImport("kernel32.dll")]
        public static extern IntPtr _lopen(string lpPathName, int iReadWrite);

        [DllImport("kernel32.dll")]
        public static extern bool CloseHandle(IntPtr hObject);

        public const int OF_READWRITE = 2;
        public const int OF_SHARE_DENY_NONE = 0x40;
        public readonly IntPtr HFILE_ERROR = new IntPtr(-1);

        public MainWindow()
        {
            InitializeComponent();
        }

        ProcessingWindow windows = new ProcessingWindow();

        /// <summary>
        /// 检查文件夹下所有文件是否符合命名规则，并分类处理
        /// </summary>
        public void CheckAndMove()
        {

            ProcessingWindow windows = new ProcessingWindow();
            windows.Show();
            //遍历文件夹的文件名
            var files = Directory.GetFiles(this.txt_Path.Text, "*.xls", SearchOption.TopDirectoryOnly);
            List<string> notSupportedFiles = new List<string>();
            List<string> okFiles = new List<string>();
            List<string> errFiles = new List<string>();
            foreach (string filename in files)
            {

                //if (!Directory.Exists(this.txt_Path.Text + "\\整理好的报表"))//如果不存在就创建文件夹
                //{
                //    Directory.CreateDirectory(this.txt_Path.Text + "\\整理好的报表");
                //}

                var name = System.IO.Path.GetFileName(filename);
                var extension = System.IO.Path.GetExtension(filename);

                if (extension == ".xls")
                {
                    notSupportedFiles.Add(name);
                    MoveProblemFiles(filename);
                }
                else if (Regex.IsMatch(name, @"^MSP(0?[1-9]|1[0-2])月月报-[0-9]{4}")) //@".xls|.png|.gif$"))
                {

                    okFiles.Add(name);
                    //此处处理为放置在原位置不动。

                }
                else //不符合命名规则
                {
                    errFiles.Add(name);
                    MoveProblemFiles(filename);
                }

            }
            windows.Close();
            if (errFiles.Count != 0 || notSupportedFiles.Count != 0)
                MessageBox.Show(
                    "初步整理中发现了一些问题。\n为确保接下来操作的数据完整性，可能存在问题的报表放置在该目录下创建的\"问题报表\" 子文件夹下。请进行手动处理，处理后请移动到主文件夹并再次进行一次整理。\n共处理 "
                    + (files.Length) + " 个月报。\n其中：\n合格文件数量 " + okFiles.Count + " 个。\n" +
                    "不支持的老版本 .xls 文件数量 " + notSupportedFiles.Count + " 个。\n" + "文件名可能有问题的文件数量为 " + errFiles.Count +
                    " 个。\n", "存在不合格的文件", MessageBoxButton.OK, MessageBoxImage.Error);
            else
            {
                if (okFiles.Count == 0)
                {
                    MessageBox.Show("没有找到任何Excel文件", "Error", MessageBoxButton.OK, MessageBoxImage.Error);

                }

                else
                    MessageBox.Show("初步整理完毕！\n共处理 " + (files.Length) + " 个月报。\n所有文件均通过新版报表文件名格式检查。", "新版报表格式检查",
                        MessageBoxButton.OK, MessageBoxImage.Asterisk);

            }
            //将符合命名规则的文件放置到新建的文件夹中，不符合命名规则的文件记录并放到另一个文件夹中

            //提示操作的文件列表和统计信息
        }

        private void MoveProblemFiles(string filename)
        {
            if (!Directory.Exists(this.txt_Path.Text + "\\问题报表")) //如果不存在就创建文件夹
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
            catch (Exception ex)
            {
                FileInfo fo = new FileInfo(filename);

                IntPtr vHandle = _lopen(filename, OF_READWRITE | OF_SHARE_DENY_NONE);
                if (vHandle == HFILE_ERROR)
                {
                    MessageBox.Show("移动文件出现问题，名为" + fo.Name + "的文件被占用！\n请确保关闭所有打开的Excel窗口并重新执行整理。");
                    windows.Close();
                    return;
                }
                CloseHandle(vHandle);
                MessageBox.Show("文件操作中出现错误！错误信息为：" + ex.Message);


            }
        }

        List<MonthlyReport> reportList = new List<MonthlyReport>();
        /// <summary>
        /// 对于单个文件的行移动操作
        /// </summary>
        private void MigrateToOne(string file) //string file)
        {
            FileInfo report = new FileInfo(file);
            using (ExcelPackage package = new ExcelPackage(report))
            {

                // get the first worksheet in the workbook
                ExcelWorksheet ws = package.Workbook.Worksheets[1];
                int row = 2; //The item description
                             // output the data in column 2
                try
                {
                    MonthlyReport rp = new MonthlyReport
                    {
                        Id = (string)ws.Cells[row, 1].Value,
                        Name = (string)ws.Cells[row, 2].Value,
                        Sex = (string)ws.Cells[row, 3].Value,
                        Team = (string)ws.Cells[row, 4].Value,
                        Phone = (string)ws.Cells[row, 5].Value,
                        Email = (string)ws.Cells[row, 6].Value,
                        City = (string)ws.Cells[row, 7].Value,
                        School = (string)ws.Cells[row, 8].Value,
                        Major = (string)ws.Cells[row, 9].Value,
                        Grade = (string)ws.Cells[row, 10].Value,
                        BlogNum = (int)ws.Cells[row, 11].Value,
                        BlogLink = (string)ws.Cells[row, 12].Value,
                        SocialNum = (int)ws.Cells[row, 13].Value,
                        SocialLink = (string)ws.Cells[row, 14].Value,
                        Retweets = (int)ws.Cells[row, 15].Value,
                        RtLink = (string)ws.Cells[row, 16].Value,
                        PostAccepted = (int)ws.Cells[row, 17].Value,
                        PostLink = (string)ws.Cells[row, 18].Value,
                        WindowsApps = (int)ws.Cells[row, 19].Value,
                        WaLink = (string)ws.Cells[row, 20].Value,
                        ActivityJoinNum = (int)ws.Cells[row, 21].Value,
                        AhLink = (string)ws.Cells[row, 22].Value,
                        ActivityHeldNum = (int)ws.Cells[row, 23].Value,
                        AjNum = (string)ws.Cells[row, 24].Value
                    };
                    reportList.Add(rp);
                }
                catch (Exception ex)
                {   //异常
                    //记录异常信息到LOG List
                    MessageBox.Show("处理文件：" + new FileInfo(file).Name + "发生问题！，请检查文件！");
                    this.WriteLine(ex.Message);
                }
                
            }//end of using




        }




        public string ErrorLog;
        private void WriteLine(string something)
        {
            ErrorLog = ErrorLog + something + "\n";
        }


        public void DoAllMigrate()
        {

           
            if (!Directory.Exists(this.txt_Path.Text + "\\月度总报表"))//如果不存在就创建文件夹
            {

                try
                {
                    Directory.CreateDirectory(this.txt_Path.Text + "\\月度总报表");
                }
                catch (Exception ex)
                {
                    MessageBox.Show("无法写入新文件夹！请检查目录是否存在或有写入权限！" + Environment.NewLine + ex.Message);
                }
            }
            //创建月度总表
            if (File.Exists(this.txt_Path.Text + "\\月度总报表\\" + "MSP" + DateTime.Now.Month.ToString() + "月总报表"))
            //存在就删除重新生成咯
            {
                File.Delete(this.txt_Path.Text + "\\月度总报表\\" + "MSP" + DateTime.Now.Month.ToString() + "月总报表");

            }
            var filename = "MSP新版月报模板.xlsx";
            Stream str = System.Reflection.Assembly.GetExecutingAssembly().GetManifestResourceStream("MonthlyReportCollector.Assets." + filename);
            using (Stream output = File.Create(this.txt_Path.Text + "\\月度总报表\\" + "MSP" + (DateTime.Now.AddMonths(-1).Month).ToString() + "月总报表.xlsx"))
            {
                str.CopyTo(output);
            }
            //
            
            ErrorLog = "";
            var files = Directory.GetFiles(this.txt_Path.Text, "*.xlsx", SearchOption.TopDirectoryOnly);

            foreach (var file in files)
            {
                MigrateToOne(file);
               
            }
            List<MonthlyReport> sortedList = reportList.OrderBy(o => o.Id).ToList();
            FileInfo mainReport = new FileInfo(this.txt_Path.Text + "\\月度总报表\\" + "MSP" + (DateTime.Now.AddMonths(-1).Month).ToString() + "月总报表.xlsx");
            using (ExcelPackage mainExcel = new ExcelPackage(mainReport))
            {
                ExcelWorksheet ws = mainExcel.Workbook.Worksheets[1];
                int row = 2; //The item description
                             // output the data in column 2
              
                foreach (var report in sortedList )
                {
                    ws.Cells[row, 1].Value = report.Id;
                    ws.Cells[row, 2].Value = report.Name;
                    ws.Cells[row, 3].Value = report.Sex;
                    ws.Cells[row, 4].Value = report.Team;
                    ws.Cells[row, 5].Value = report.Phone;
                    ws.Cells[row, 6].Value = report.Email;
                    ws.Cells[row, 7].Value = report.City;
                    ws.Cells[row, 8].Value = report.School;
                    ws.Cells[row, 9].Value = report.Major;
                    ws.Cells[row, 10].Value = report.Grade;
                    ws.Cells[row, 11].Value = report.BlogNum;
                    ws.Cells[row, 12].Value = report.BlogLink;
                    ws.Cells[row, 13].Value = report.SocialNum;
                    ws.Cells[row, 14].Value = report.SocialLink;
                    ws.Cells[row, 15].Value = report.Retweets;
                    ws.Cells[row, 16].Value = report.RtLink;
                    ws.Cells[row, 17].Value = report.PostAccepted;
                    ws.Cells[row, 18].Value = report.PostLink;
                    ws.Cells[row, 19].Value = report.WindowsApps;
                    ws.Cells[row, 20].Value = report.WaLink;
                    ws.Cells[row, 21].Value = report.ActivityHeldNum;
                    ws.Cells[row, 22].Value = report.AhLink;
                    ws.Cells[row, 23].Value = report.ActivityJoinNum;
                    ws.Cells[row, 24].Value = report.ActivityJoinNum;
                    row++;
                }
            }
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

            }
        }

        private void btn_Migrate_Click(object sender, RoutedEventArgs e)
        {
            DoAllMigrate();
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

        private void btn_Verify_Click(object sender, RoutedEventArgs e)
        {
            if (Directory.Exists(txt_Path.Text))
                CheckAndMove();
            else
            {
                MessageBox.Show("请选择合法的路径！", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }
}
