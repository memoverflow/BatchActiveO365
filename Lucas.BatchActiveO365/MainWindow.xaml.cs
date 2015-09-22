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
using System.Threading;

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
            await SetLog();
            var isCompleted = await LoadUsers();
            if (isCompleted) return;
            await SetAutoRun();
            var result = await SetRestart();
            if (result)
            {
                await SetOpenApplicationAndWait();
                await SetRemoveLicense();
                await SetComputerName();
                await SetAutoLogon();
            }
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

                    RegisterTool.WriteDefaultLogin(currentUser, object.Equals(ConfigurationManager.AppSettings["IsEnableDomain"], "true"));
                    System.Diagnostics.Process.Start("shutdown", @"/r /t 0");

                }
                catch (Exception ex)
                {
                    msg += "遇到错误";
                    LogHelper.WriteLog(msg, ex);
                }

            });

        }

        async Task SetComputerName()
        {
            var msg = "设置计算机名称中...";

            loadingTextControl.Text = msg;
            LogHelper.WriteLog(msg);
            await Task.Factory.StartNew(() =>
            {
                try
                {
                    if (CurrentIndex % 5 == 0)
                    {
                        var currentUser = Users[CurrentIndex];
                        RegisterTool.SetMachineName(currentUser.UserName + "PC");
                    }
                }
                catch (Exception ex)
                {
                    msg += "遇到错误";
                    LogHelper.WriteLog(msg, ex);
                }

            });
        }

        async Task<bool> SetRestart()
        {
            var msg = "设置检查当前用户中...";
            var currentUser = Users[CurrentIndex];
            loadingTextControl.Text = msg;
            LogHelper.WriteLog(msg);
            return await Task<bool>.Factory.StartNew(() =>
            {
                try
                {
                    //如果当前用户是excel表格里面的第一个用户
                    if (Environment.UserName != currentUser.UserName)
                    {
                        
                        return this.Dispatcher.Invoke(() =>
                        {
                            var result = MessageBox.Show("当前登录用户和要执行用户不是同一用户，是否重启？", "", MessageBoxButton.YesNo);
                            if (result == MessageBoxResult.Yes)
                            {

                                RegisterTool.WriteDefaultLogin(currentUser, object.Equals(ConfigurationManager.AppSettings["IsEnableDomain"], "true"));
                                System.Diagnostics.Process.Start("shutdown", @"/r /t 0");
                                return true;
                            }
                            else
                            {
                                return false;
                            }

                            
                        });

                    }
                    return true;
                }
                catch (Exception ex)
                {
                    msg += "遇到错误";
                    LogHelper.WriteLog(msg, ex);
                    return false;
                }
            });

        }

        //从本地文件中获取用户
        async Task<bool> LoadUsers()
        {
            var msg = "获取本地文件并解析...";
            loadingTextControl.Text = msg;
            LogHelper.WriteLog(msg);
            return await Task.Factory.StartNew(() =>
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
                            return true;
                        }
                        else
                        {
                            dataGridControl.SelectedIndex = CurrentIndex;
                            indexControl.Text = CurrentIndex + 1 + "";
                        }

                        return false;
                    });
                }
                catch (Exception ex)
                {
                    msg += "遇到错误";
                    LogHelper.WriteLog(msg, ex);
                }
                return false;
            });

        }

        //设置打开excel，病等待固定时间
        async Task SetOpenApplicationAndWait()
        {
            var msg = "打开Excel并等待...";
            loadingTextControl.Text = msg;
            LogHelper.WriteLog(msg);
            await Task.Factory.StartNew(() =>
            {
                try
                {
                    var waitSeconds = Convert.ToInt32(ConfigurationManager.AppSettings["WaitSeconds"]);
                    var currentUser = Users[CurrentIndex];
                    var fileName = System.AppDomain.CurrentDomain.BaseDirectory + "users.xlsx";
                    //var excelHelper = new ExcelHelper(fileName);
                    //excelHelper.WriteToExcel("Sheet1", CurrentIndex, "正在处理" + currentUser.UserName);
                    var process = Process.Start(fileName);
                    var tokenSource = new CancellationTokenSource();
                    var token = tokenSource.Token;
                    Task taskDelay = Task.Delay(waitSeconds * 1000, token);
                    taskDelay.Wait();
                    process.Kill();
                    process.Dispose();
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
