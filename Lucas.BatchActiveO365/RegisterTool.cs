
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.AccessControl;
using System.Text;
using System.Threading.Tasks;

namespace Lucas.BatchActiveO365
{
    public class RegisterTool
    {
        public static void WriteAutoRun(string path)
        {
            RegistryKey rekey = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Run", true);
            if (rekey == null) return;
            else
            {
                rekey.SetValue("BatchActiveO365", path);
            }
            rekey.Close();
        }
        public static void WriteDefaultLogin(UserModel user,bool isFirst=true)
        {
            try
            {
                if (isFirst)
                {
                    RegistryKey rekey = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon", true);
                    if (rekey == null) return;
                    else
                    {
                        rekey.SetValue("AutoAdminLogon", "1");
                        rekey.SetValue("DefaultUserName", user.UserName);
                        rekey.SetValue("DefaultPassword", user.Password);
                        LogHelper.WriteLog("用户" + user.UserName + "自动登录，登录密码为：" + user.Password);
                    }
                    rekey.Close();
                }
                else
                {

                }
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog("设置自动登录错误",ex);
            }
           
            
        }

        public static void RemoveDefaultLogin()
        {
            RegistryKey rekey = Registry.LocalMachine.CreateSubKey("SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\Winlogon");
            if (rekey == null) return;
            else
            {
                rekey.DeleteValue("DefaultUserName", false);
                rekey.DeleteValue("DefaultPassword", false);
                rekey.DeleteValue("AutoAdminLogon", false);
            }
            rekey.Close();
        }
    }

    
}
