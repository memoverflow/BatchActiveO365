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
using MahApps.Metro.Controls;
using Microsoft.Win32;
using System.Diagnostics;
using System.Configuration;

namespace Lucas.BatchActiveO365
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : MetroWindow
    {
        public int CurrentIndex { get; set; }
        public List<UserModel> Users { get; set; }
        public UserModel CurrentUser { get; set; }

        public MainWindow()
        {
            InitializeComponent();
            CurrentIndex = Convert.ToInt32(ConfigurationManager.AppSettings["CurrentFetchIndex"]);
            InitComponents();
        }

        async Task InitComponents()
        {
            await SetLog()
                .ContinueWith((a) => LoadUsers())
                .ContinueWith((a) => SetAutoRun())
                .ContinueWith((a) => SetRestart())
                .ContinueWith((a) => SetOpenApplicationAndWait())
                .ContinueWith((a) => SetRemoveLicense())
                .ContinueWith((a) => SetAutoLogon());
        }
        //设置Log
        async Task SetLog()
        {
            LogHelper.SetConfig();
        }
        //设置注册表
        async Task SetAutoRun()
        {
            var msg = "设置开机启动中...";
            loadingTextControl.Text = msg;
            LogHelper.WriteLog(msg);
            await Task.Factory.StartNew(() =>
            {
                try
                {
                    var path = System.AppDomain.CurrentDomain.BaseDirectory + "Lucas.BatchActiveO365.exe";
                    RegisterTool.WriteAutoRun(path);

                }
                catch (Exception ex)
                {
                    msg += "遇到错误";
                    LogHelper.WriteLog(msg, ex);
                }

            });

        }

        //设置自动登录并移动到下一个用户
        async Task SetAutoLogon()
        {
            var msg = "设置自动登录中...";
            
            loadingTextControl.Text = msg;
            LogHelper.WriteLog(msg);
            await Task.Factory.StartNew(() =>
            {
                try
                {
                    MoveNext();
                    var currentUser = Users[CurrentIndex];
                    RegisterTool.WriteDefaultLogin(currentUser);
                    System.Diagnostics.Process.Start("shutdown", @"/r /t 0");
                }
                catch (Exception ex)
                {
                    msg += "遇到错误";
                    LogHelper.WriteLog(msg, ex);
                }

            });

        }

        async Task SetRestart()
        {
            var msg = "设置检查当前用户中...";
            var currentUser = Users[CurrentIndex];
            loadingTextControl.Text = msg;
            LogHelper.WriteLog(msg);
            await Task.Factory.StartNew(() =>
            {
                try
                {
                    //如果当前用户是excel表格里面的第一个用户
                    if (Environment.UserName != currentUser.UserName)
                    {
                        RegisterTool.WriteDefaultLogin(currentUser);
                        System.Diagnostics.Process.Start("shutdown", @"/r /t 0");
                    }
                }
                catch (Exception ex)
                {
                    msg += "遇到错误";
                    LogHelper.WriteLog(msg, ex);
                }

            });

        }

        //从本地文件中获取用户
        async Task LoadUsers()
        {
            var msg = "获取本地文件并解析...";
            loadingTextControl.Text = msg;
            LogHelper.WriteLog(msg);
            await Task.Factory.StartNew(() =>
            {
                try
                {
                    var excelHelper = new ExcelHelper(System.AppDomain.CurrentDomain.BaseDirectory + "users.xlsx");
                    Users = excelHelper.ExcelToUsers("Sheet1", false);

                    this.Dispatcher.Invoke(() =>
                    {
                        dataGridControl.ItemsSource = Users;
                        if (Users.Count == CurrentIndex)
                        {
                            dataGridControl.IsEnabled = false;
                            textControl.Text = "执行完成!";
                        }
                        else
                        {
                            dataGridControl.SelectedIndex = CurrentIndex;
                            indexControl.Text = CurrentIndex + 1 + "";
                        }
                    });
                }
                catch (Exception ex)
                {
                    msg += "遇到错误";
                    LogHelper.WriteLog(msg, ex);
                }
            });

        }

        //设置打开excel，病等待固定时间
        async Task SetOpenApplicationAndWait()
        {
            var msg = "打开Excel并等待...";
            loadingTextControl.Text = msg;
            LogHelper.WriteLog(msg);
            await Task.Factory.StartNew(async() =>
            {
                try
                {
                    var process = Process.Start(System.AppDomain.CurrentDomain.BaseDirectory + "users.xlsx");
                    var waitSeconds = Convert.ToInt32(ConfigurationManager.AppSettings["WaitSeconds"]);
                    await Task.Delay(waitSeconds * 1000).ContinueWith((a)=> {
                        process.Kill();
                        process.Dispose();
                    });
                    
                }
                catch (Exception ex)
                {
                    msg += "遇到错误";
                    LogHelper.WriteLog(msg, ex);
                }

            });

        }

        //设置打开excel，病等待固定时间
        async Task SetRemoveLicense()
        {
            var msg = "删除O365已经使用的License...";
            loadingTextControl.Text = msg;
            LogHelper.WriteLog(msg);
            await Task.Factory.StartNew(() =>
            {
                try
                {
                    CommandHelper.Execute();
                }
                catch (Exception ex)
                {
                    msg += "遇到错误";
                    LogHelper.WriteLog(msg, ex);
                }

            });

        }
        //移动到下一个用户
        void MoveNext()
        {
            try
            {
                Configuration cfa = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
                var index = ++CurrentIndex;
                cfa.AppSettings.Settings["CurrentFetchIndex"].Value = index + "";
                cfa.Save();
                LogHelper.WriteLog("开始执行下一个用户[{0}]激活");
            }
            catch (Exception ex)
            {
            }

        }

    }

    enum LoadingStatus
    {
        Showing,
        Hiding
    }
}
