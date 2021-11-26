using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
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
using System.Threading;
using System.Windows.Forms;
using AutoUpdate.Win32Service.Win32;
using MessageBox = System.Windows.MessageBox;

namespace AutoUpdate
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        private string websiteUrl = @"C:\Program Files (x86)\Kingdee\K3Cloud\WebSite\bin\";//dll更新的目标位置
        private string shellUrl = @"C:\Program Files (x86)\Kingdee\K3Cloud\DeskClient\K3CloudClient\";
        private string shellName = "Kingdee.BOS.XPF.App";//客户端
        private string bosName = "Kingdee.BOS.IDE";//集成开发平台BOS
        private string cmdCommand = "iisreset";
        System.Timers.Timer timer;
        System.Timers.Timer timer1;
        System.Timers.Timer timer2;
        System.Timers.Timer bosTimer;

        System.Timers.Timer remoteTimerNew;
        System.Timers.Timer remoteTimerOld;
        System.Timers.Timer remoteTimer130;
        System.Timers.Timer additionProcTimer;
        public MainWindow()
        {
            InitializeComponent();
        }
        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            timer = new System.Timers.Timer();
            timer.Interval = 3000;
            timer.Elapsed += Timer_Elapsed;

            timer1 = new System.Timers.Timer();
            timer1.Interval = 50000;
            timer1.Elapsed += Timer1_Elapsed;

            bosTimer = new System.Timers.Timer();
            bosTimer.Interval = 5000;
            bosTimer.Elapsed += BosTimer_Elapsed;

            additionProcTimer = new System.Timers.Timer();
            additionProcTimer.Interval = 1500;
            additionProcTimer.Elapsed += AdditionProcTimer_Elapsed;

        }

        private void BosTimer_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
            bosTimer.Stop();
            SendKeys.SendWait("{TAB 4}");
            Thread.Sleep(500);
            SendKeys.SendWait("{1}");
            Thread.Sleep(50);
            SendKeys.SendWait("{2}");
            Thread.Sleep(50);
            SendKeys.SendWait("{3}");
            Thread.Sleep(50);
            SendKeys.SendWait("{4}");
            Thread.Sleep(50);
            SendKeys.SendWait("{5}");
            Thread.Sleep(50);
            SendKeys.SendWait("{6}");
            Thread.Sleep(500);
            User.keybd_event((byte)Keys.Enter, 0, 0, 0);
            //Thread.Sleep(500);
            //User.keybd_event((byte)Keys.Enter, 0, 0, 0);
        }

        private void Window_Unloaded(object sender, RoutedEventArgs e)
        {
            if (timer.Enabled)
            {
                timer.Stop();
            }
            if (timer1.Enabled)
            {
                timer1.Stop();
            }
            if (timer2.Enabled)
            {
                timer2.Stop();
            }
            if (bosTimer.Enabled)
            {
                bosTimer.Stop();
            }
            if (remoteTimerNew.Enabled)
            {
                remoteTimerNew.Stop();
            }
            if (remoteTimerOld.Enabled)
            {
                remoteTimerOld.Stop();
            }
            if (remoteTimer130.Enabled)
            {
                remoteTimerOld.Stop();
            }
            if (additionProcTimer.Enabled)
            {
                additionProcTimer.Stop();
            }
            Environment.Exit(0);
        }

        private void ReplaceFiles()
        {
            int index = comboModule.SelectedIndex;
            SysT<Readconfig>.GetInstance.ReadConfig(string.Format("{0}\\Config\\config.ini", Environment.CurrentDirectory));
            int confingCount = SysT<Readconfig>.GetInstance.Count;
            SysT<Config>.GetInstance.configList = new List<string>();
            for (int i = 1; i <= confingCount; i++)
            {
                SysT<Config>.GetInstance.configList.Add(SysT<Readconfig>.GetInstance["Module" + i]);
            }
            string sourceFile = SysT<Config>.GetInstance.configList[index].Split(',')[0] + SysT<Config>.GetInstance.configList[index].Split(',')[1];
            string destFile = websiteUrl + SysT<Config>.GetInstance.configList[index].Split(',')[1];
            if (File.Exists(destFile))
            {
                File.Delete(destFile);
            }
            File.Copy(sourceFile, destFile);
        }
        private void BtnExecute_Click(object sender, RoutedEventArgs e)
        {
            BtnExecute.IsEnabled = false;
            RunExecute();
        }
        /// <summary>
        /// 执行
        /// </summary>
        private void RunExecute()
        {
            listbox.Items.Clear();
            ReplaceFiles();
            listbox.Items.Add("DLL替换完成-----\r\n");
            string cmdResult = WriteToCmd.StartCmd(cmdCommand);
            listbox.Items.Add(cmdResult.TrimEnd() + "\r\n");
            listbox.Items.Add("iis重启完成-----\r\n");
            OpenExe(shellName, shellUrl);
            listbox.Items.Add("K3启动登录中-----\r\n");
            listbox.Items.Add(DateTime.Now.ToString());
            listbox.Items.Add("-----------------------");
            timer.Start();
            BtnExecute.IsEnabled = true;
        }
        private void Timer_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
            Process[] myPs;
            myPs = Process.GetProcesses();
            var ps = myPs.OrderBy(o => o.ProcessName);
            foreach (Process p in myPs)
            {
                if (p.ProcessName == shellName)
                {
                    timer.Stop();
                    timer1.Start();
                }
            }
        }
        private void Timer1_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
            timer1.Stop();
            if (timer2 != null && timer2.Enabled)
            {
                timer2.Stop();
            }
            User.keybd_event((byte)Keys.Enter, 0, 0, 0);
            Thread.Sleep(50);
            User.keybd_event((byte)Keys.Enter, 0, 0, 0);
            Thread.Sleep(50);
            User.keybd_event((byte)Keys.D1, 0, 0, 0);
            Thread.Sleep(50);
            User.keybd_event((byte)Keys.D2, 0, 0, 0);
            Thread.Sleep(50);
            User.keybd_event((byte)Keys.D3, 0, 0, 0);
            Thread.Sleep(50);
            User.keybd_event((byte)Keys.D4, 0, 0, 0);
            Thread.Sleep(50);
            User.keybd_event((byte)Keys.D5, 0, 0, 0);
            Thread.Sleep(50);
            User.keybd_event((byte)Keys.D6, 0, 0, 0);
            Thread.Sleep(100);
            User.keybd_event((byte)Keys.Enter, 0, 0, 0);
            Thread.Sleep(1000);
            User.keybd_event((byte)Keys.Enter, 0, 0, 0);
            Thread.Sleep(2000);
            User.keybd_event((byte)Keys.Enter, 0, 0, 0);
        }
        private void OpenExe(string processName, string fileUrl,bool closeOld=true)
        {
            try
            {
                if (!string.IsNullOrEmpty(fileUrl))
                {
                    if (closeOld)
                    {
                        CloseExe(processName);
                    }
                    //Delaytimer.Stop();
                    //开启一个新process
                    ProcessStartInfo p = null;
                    Process proc = null;

                    p = new ProcessStartInfo(processName + ".exe");
                    p.WorkingDirectory = fileUrl;//设置此外部程序所在windows目录
                    p.UseShellExecute = true;
                    proc = Process.Start(p);//调用外部程序
                    proc.Dispose();
                    proc.Close();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void CloseExe(string processName)
        {
            Process[] myPs;
            myPs = Process.GetProcesses();
            var ps = myPs.OrderBy(o => o.ProcessName);
            foreach (Process p in myPs)
            {
                if (p.ProcessName == processName)
                {
                    p.Kill();
                    p.WaitForExit();
                    p.Close();
                    break;
                }
            }
        }

        private void BtnIntegration_Click(object sender, RoutedEventArgs e)
        {
            RunInteration();
        }
        /// <summary>
        /// 打开集成开发平台
        /// </summary>
        private void RunInteration()
        {
            OpenExe(bosName, shellUrl);
            bosTimer.Start();
            this.WindowState = WindowState.Minimized;
        }
        private void BtnRemote_Click(object sender, RoutedEventArgs e)
        {
            BtnRemote.IsEnabled = false;
            RunNewRemote();
        }
        /// <summary>
        /// 打开新K3远程桌面
        /// </summary>
        private void RunNewRemote()
        {
            OpenExe("mstsc", @"C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Accessories");
            remoteTimerNew = new System.Timers.Timer();
            remoteTimerNew.Interval = 800;
            remoteTimerNew.Elapsed += Timer_Tick;
            remoteTimerNew.Start();
            this.WindowState = WindowState.Minimized;
            BtnRemote.IsEnabled = true;
        }
        private void BtnRemoteOld_Click(object sender, RoutedEventArgs e)
        {
            BtnRemoteOld.IsEnabled = false;
            RunOldRemote();
        }
        private void RunOldRemote()
        {
            OpenExe("mstsc", @"C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Accessories");
            remoteTimerOld = new System.Timers.Timer();
            remoteTimerOld.Interval = 1000;
            remoteTimerOld.Elapsed += TimerOld_Tick;
            remoteTimerOld.Start();
            this.WindowState = WindowState.Minimized;
            BtnRemoteOld.IsEnabled = true;
        }
        private void Timer_Tick(object sender, EventArgs e)
        {
            remoteTimerNew.Stop();
            SameKeys1();
            SendKeys.SendWait("{1}");
            Thread.Sleep(50);
            SendKeys.SendWait("{8}");
            Thread.Sleep(50);
            SendKeys.SendWait("{2}");
            Thread.Sleep(50);

            SendKeys.SendWait("{:}");
            Thread.Sleep(50);

            SendKeys.SendWait("{8}");
            Thread.Sleep(50);
            SendKeys.SendWait("{9}");
            Thread.Sleep(50);
            SendKeys.SendWait("{1}");
            Thread.Sleep(50);
            SendKeys.SendWait("{7}");
            Thread.Sleep(200);
            SendKeys.SendWait("{ENTER}");
            SameKeys2();
        }
        private void Run130Remote()
        {
            OpenExe("mstsc", @"C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Accessories");
            remoteTimer130 = new System.Timers.Timer();
            remoteTimer130.Interval = 1000;
            remoteTimer130.Elapsed += RemoteTimer130_Elapsed;
            remoteTimer130.Start();
            this.WindowState = WindowState.Minimized;
            BtnRemote130.IsEnabled = true;
        }

        private void RemoteTimer130_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
            remoteTimer130.Stop();
            SendKeys.SendWait("+");
            Thread.Sleep(500);
            SendKeys.SendWait("{1}");
            Thread.Sleep(50);
            SendKeys.SendWait("{0}");
            Thread.Sleep(50);
            SendKeys.SendWait("{.}");
            Thread.Sleep(50);

            SendKeys.SendWait("{1}");
            Thread.Sleep(50);
            SendKeys.SendWait("{0}");
            Thread.Sleep(50);
            SendKeys.SendWait("{.}");
            Thread.Sleep(50);

            SendKeys.SendWait("{2}");
            Thread.Sleep(50);
            SendKeys.SendWait("{0}");
            Thread.Sleep(50);
            SendKeys.SendWait("{1}");
            Thread.Sleep(50);
            SendKeys.SendWait("{.}");
            Thread.Sleep(50);
            SendKeys.SendWait("{1}");
            Thread.Sleep(50);
            SendKeys.SendWait("{3}");
            Thread.Sleep(50);
            SendKeys.SendWait("{0}");
            Thread.Sleep(50);

            SendKeys.SendWait("{:}");
            Thread.Sleep(50);

            SendKeys.SendWait("{8}");
            Thread.Sleep(50);
            SendKeys.SendWait("{9}");
            Thread.Sleep(50);
            SendKeys.SendWait("{1}");
            Thread.Sleep(50);
            SendKeys.SendWait("{7}");
            Thread.Sleep(50);
            SendKeys.SendWait("{ENTER}");
            Thread.Sleep(500);
            SameKeys2();
            //Thread.Sleep(1000);
            //SendKeys.SendWait("{ENTER}");
        }

        private void TimerOld_Tick(object sender, EventArgs e)
        {
            remoteTimerOld.Stop();
            SameKeys1();
            SendKeys.SendWait("{2}");
            Thread.Sleep(50);
            SendKeys.SendWait("{0}");
            Thread.Sleep(50);
            SendKeys.SendWait("{5}");
            Thread.Sleep(50);

            SendKeys.SendWait("{:}");
            Thread.Sleep(50);

            SendKeys.SendWait("{8}");
            Thread.Sleep(50);
            SendKeys.SendWait("{9}");
            Thread.Sleep(50);
            SendKeys.SendWait("{1}");
            Thread.Sleep(50);
            SendKeys.SendWait("{7}");
            Thread.Sleep(50);
            SendKeys.SendWait("{ENTER}");
            Thread.Sleep(500);
            SameKeys2();
            Thread.Sleep(1000);
            SendKeys.SendWait("{ENTER}");
        }
        private void SameKeys1()
        {
            SendKeys.SendWait("+");
            Thread.Sleep(500);
            SendKeys.SendWait("{1}");
            Thread.Sleep(50);
            SendKeys.SendWait("{0}");
            Thread.Sleep(50);
            SendKeys.SendWait("{.}");
            Thread.Sleep(50);

            SendKeys.SendWait("{1}");
            Thread.Sleep(50);
            SendKeys.SendWait("{0}");
            Thread.Sleep(50);
            SendKeys.SendWait("{.}");
            Thread.Sleep(50);

            SendKeys.SendWait("{2}");
            Thread.Sleep(50);
            SendKeys.SendWait("{0}");
            Thread.Sleep(50);
            SendKeys.SendWait("{2}");
            Thread.Sleep(50);
            SendKeys.SendWait("{.}");
            Thread.Sleep(50);
        }
        private void SameKeys2()
        {
            Thread.Sleep(300);
            SendKeys.SendWait("+^");
            Thread.Sleep(1000);
            SendKeys.SendWait("{Z}");
            Thread.Sleep(50);
            SendKeys.SendWait("{z}");
            Thread.Sleep(50);
            SendKeys.SendWait("{p}");
            Thread.Sleep(50);
            SendKeys.SendWait("{4}");
            Thread.Sleep(50);
            SendKeys.SendWait("{2}");
            Thread.Sleep(50);
            SendKeys.SendWait("{9}");
            Thread.Sleep(50);
            SendKeys.SendWait("{0}");
            Thread.Sleep(50);
            SendKeys.SendWait("{5}");
            Thread.Sleep(50);
            SendKeys.SendWait("{6}");
            Thread.Sleep(50);
            SendKeys.SendWait("{0}");
            Thread.Sleep(50);
            SendKeys.SendWait("{1}");
            Thread.Sleep(50);
            SendKeys.SendWait("{5}");
            Thread.Sleep(100);
            SendKeys.SendWait("{ENTER}");
        }

        private void BtnAdditionProc_Click(object sender, RoutedEventArgs e)
        {
            additionProcTimer.Start();
        }
        private void AdditionProcTimer_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
            additionProcTimer.Stop();
           this.Dispatcher.Invoke(new Action(()=> this.WindowState = WindowState.Minimized)) ;
            Thread.Sleep(50);
            SendKeys.SendWait("%d");
            Thread.Sleep(500);
            SendKeys.SendWait("p");
            Thread.Sleep(2000);
            SendKeys.SendWait("{w}");
            Thread.Sleep(2000);
            SendKeys.SendWait("{ENTER}");
        }

        private void BtnWeekNewsPaper_Click(object sender, RoutedEventArgs e)
        {
            WindowWeek windowWeek = new WindowWeek();
            windowWeek.ShowDialog();
        }

        private void Window_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            string key = e.Key.ToString().ToUpperInvariant();
            switch (key)
            {
                case "F1":
                    RunInteration();
                    break;
                case "F2":
                    RunNewRemote();
                    break;
                case "F3":
                    RunOldRemote();
                    break;
                case "F4":
                    additionProcTimer.Start();
                    break;
                case "F5":
                    WindowWeek windowWeek = new WindowWeek();
                    windowWeek.ShowDialog();
                    break;
                case "F6":
                    StartNewK3();
                    break;
                case "F7":
                    CloseExe(shellName);
                    break;
                case "F8":
                    OpenWeekReport();
                    break;
                case "F9":
                    OpenJobInfo();
                    break;
                case "SYSTEM":
                    OpenVS();
                    break;
                case "F11":
                    Run130Remote();
                    break;
                case "RETURN":
                    RunExecute();
                    break;
            }
        }

        private void BtnSearch_Click(object sender, RoutedEventArgs e)
        {
            SendKeys.SendWait("{TAB}");
            Thread.Sleep(500);
            string text = comboBill.Text;
            this.WindowState = WindowState.Minimized;
        }

        private void BtnStartNewK3_Click(object sender, RoutedEventArgs e)
        {
            StartNewK3();
        }
        private void StartNewK3()
        {
            OpenExe(shellName, shellUrl);
            timer2 = new System.Timers.Timer();
            timer2.Interval = 6000;
            timer2.Elapsed += Timer1_Elapsed;
            timer2.Start();
            this.WindowState = WindowState.Minimized;  
        }

        private void BtnWeekReport_Click(object sender, RoutedEventArgs e)
        {
            OpenWeekReport();
        }
        private void OpenWeekReport()
        {
            string v_OpenFolderPath = @"\\YL-HQFS1\Share Files\厦门银鹭食品集团\信息中心\应用开发部\对内\00工作计划及安排\总结ppt\2021总结";
            Process.Start("explorer.exe", v_OpenFolderPath);
        }
        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            Environment.Exit(0);
        }

        private void BtnJobInfo_Click(object sender, RoutedEventArgs e)
        {
            OpenJobInfo();
        }
        private void OpenJobInfo()
        {
            string openFolderPath = @"D:\工作资料";
            Process.Start("explorer.exe", openFolderPath);
        }

        private void BtnCloseNewK3_Click(object sender, RoutedEventArgs e)
        {
            CloseExe(shellName);
        }

        private void BtnOpenVS_Click(object sender, RoutedEventArgs e)
        {
            OpenVS();
        }
        private void OpenVS()
        {
            OpenExe("devenv", @"C:\Program Files (x86)\Microsoft Visual Studio\2019\Enterprise\Common7\IDE", false);
            this.WindowState = WindowState.Minimized;
        }

        private void BtnRemote130_Click(object sender, RoutedEventArgs e)
        {
            BtnRemote130.IsEnabled = true;
            Run130Remote();
        }
    }
}
