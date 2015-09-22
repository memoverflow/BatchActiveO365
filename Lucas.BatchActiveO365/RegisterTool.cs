
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Management;
using System.Runtime.InteropServices;
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
        public static void WriteDefaultLogin(UserModel user,bool isEnableDomain,bool isFirst=true)
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
                        if (isEnableDomain)
                            rekey.SetValue("DefaultDomainName", user.Domain);
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

        public static bool SetMachineName(string Name)
        {
            String RegLocComputerName = @"SYSTEM\CurrentControlSet\Control\ComputerName\ComputerName";
            try
            {
                string compPath = "Win32_ComputerSystem.Name='" + System.Environment.MachineName + "'";
                using (ManagementObject mo = new ManagementObject(new ManagementPath(compPath)))
                {
                    ManagementBaseObject inputArgs = mo.GetMethodParameters("Rename");
                    inputArgs["Name"] = Name;
                    ManagementBaseObject output = mo.InvokeMethod("Rename", inputArgs, null);
                    uint retValue = (uint)Convert.ChangeType(output.Properties["ReturnValue"].Value, typeof(uint));
                    if (retValue != 0)
                    {
                        LogHelper.WriteLog("Computer could not be changed due to unknown reason.");
                    }
                }

                RegistryKey ComputerName = Registry.LocalMachine.OpenSubKey(RegLocComputerName);
                if (ComputerName == null)
                {
                    LogHelper.WriteLog("Registry location '" + RegLocComputerName + "' is not readable.");
                }
                if (((String)ComputerName.GetValue("ComputerName")) != Name)
                {
                    LogHelper.WriteLog("The computer name was set by WMI but was not updated in the registry location: '" + RegLocComputerName + "'");
                }
                LogHelper.WriteLog("更改计算机名："+Name);
                ComputerName.Close();
                ComputerName.Dispose();
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog("出现错误", ex);
                return false;
            }
            return true;
        }
    }

    
}
