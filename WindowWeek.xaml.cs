using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace AutoUpdate
{
    /// <summary>
    /// WindowWeek.xaml 的交互逻辑
    /// </summary>
    public partial class WindowWeek : Window
    {
        private const string weekUrl = @"\\YL-HQFS1\Share Files\厦门银鹭食品集团\信息中心\应用开发部\对内\00工作计划及安排\总结ppt\2021总结";
        private string newDateRange = "";
        string constituteFile = "";
        private string targetUrl = "";
        public WindowWeek()
        {
            InitializeComponent();
            tBoxWork1.Focus();
        }

        private void BtnSend_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(constituteFile))
            {
                MessageBox.Show("文件名不能为空!");
                return;
            }
            if (string.IsNullOrWhiteSpace(newDateRange))
            {
                MessageBox.Show("日期范围不能为空!");
                return;
            }
            string templatePath = Environment.CurrentDirectory + "\\" + "周报模板.xlsx";
            string outPath = targetUrl;
            if (string.IsNullOrWhiteSpace(targetUrl))
            {
                MessageBox.Show("目标路径不能为空!");
                return;
            }
            IList<WeekModel> list = new List<WeekModel>();
            list.Add(new WeekModel { workContent = tBoxWork1.Text, workTarget = tBoxTarget1.Text, completion = combobox1.Text });
            list.Add(new WeekModel { workContent = tBoxWork2.Text, workTarget = tBoxTarget2.Text, completion = combobox2.Text });
            list.Add(new WeekModel { workContent = tBoxWork3.Text, workTarget = tBoxTarget3.Text, completion = combobox3.Text });
            list.Add(new WeekModel { workContent = tBoxWork4.Text, workTarget = tBoxTarget4.Text, completion = combobox4.Text });
            list.Add(new WeekModel { workContent = tBoxWork5.Text, workTarget = tBoxTarget5.Text, completion = combobox5.Text });
            list.Add(new WeekModel { workContent = tBoxWork6.Text, workTarget = tBoxTarget6.Text, completion = combobox6.Text });
            var weekList = list.Where(t => !t.workContent.Trim().Equals("")).ToList();
            
            if (weekList.Count == 0)
            {
                MessageBox.Show("工作内容不能为空!");
                return;
            }
            string outTargetPath = outPath + @"\" + constituteFile + ".xlsx";
            WriteToExcel.SaveWrokExcel(templatePath, outTargetPath, newDateRange, weekList);
            MessageBox.Show("发送成功!");
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            var file = GetAllFileInfos(weekUrl).Where(t => t.Name.StartsWith("W")).OrderByDescending(o => o.CreationTime).FirstOrDefault();
            if (file == null)
            {
                return;
            }
            
            int diffHours = DateDiff(DateTime.Now, file.CreationTime);
            string dateRange = Regex.Match(file.Name, @"W(?<num>\d+)\((?<date>[\s\S]*?)\)").Groups["date"].Value;
            string weekNum = Regex.Match(file.Name, @"W(?<num>\d+)\((?<date>[\s\S]*?)\)").Groups["num"].Value;
            newDateRange = dateRange;
            if (diffHours < 36)
            {
                targetUrl = file.FullName;
                constituteFile = "应用开发部W" + weekNum + "周报_周振平(" + dateRange + ")";
                fileName.Text = constituteFile;
            }
        }
        public static FileSystemInfo[] GetAllFileInfos(string url)
        {
            DirectoryInfo theFolder = new DirectoryInfo(url);
            DirectoryInfo dir = theFolder;
            FileSystemInfo[] files = dir.GetFileSystemInfos();
            return files;
        }
        /// <summary>
        /// 计算两个日期的时间间隔,返回的是时间间隔的日期差的绝对值.
        /// </summary>
        /// <param name="DateTime1">第一个日期和时间</param>
        /// <param name="DateTime2">第二个日期和时间</param>
         /// <returns></returns>
        public static int DateDiff(DateTime DateTime1, DateTime DateTime2)
        {
            int dateDiff = 0;
            try
            {
                TimeSpan ts1 = new TimeSpan(DateTime1.Ticks);
                TimeSpan ts2 = new TimeSpan(DateTime2.Ticks);
                TimeSpan ts = ts1.Subtract(ts2).Duration();
                dateDiff = ts.Days * 24 + ts.Hours;
            }
            catch
            {

            }
            return dateDiff;
        }

        private void Window_KeyDown(object sender, KeyEventArgs e)
        {
            string key = e.Key.ToString().ToUpperInvariant();
            if (key.Equals("ESCAPE"))
            {
                this.Close();
            }
        }
    }
}
