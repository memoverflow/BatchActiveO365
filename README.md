# BatchActiveO365

本应用使用WPF创建，构建与.NET Framework 4.5基础上，为了解决批量激活O365客户端的问题，使用ADFS进行授权并激活

使用须知:

1. 请使用X64的操作系统，推荐使用Windows7+。
2. 使用过程中可能涉及到多次重启。
3. 安装.NET Framework 4.5。
4. 本应用并没有在X32上测试过，不建议使用。
5. Excel文件请复制到根目录，并命名为users.xlsx，格式请参考下载文件中的格式 username | password | domain。
6. Lucas.BatchActiveO365.exe 是程序运行文件。
7. Lucas.BatchActiveO365.exe.config 是程序配置文件，如果测试本地用户请将后面键值对设为false <add key="IsEnableDomain" value ="true"/>。
8. 请切记关闭系统UAC，应为程序以管理员的身份运行，不关闭UAC需要每次重启电脑后点击一次。

更改计算机名请下载 https://github.com/xuren87/BatchActiveO365/blob/master/Lucas.BatchActiveO365/BatchActive.zip?raw=true

不计算机名请下载 https://github.com/xuren87/BatchActiveO365/blob/master/Lucas.BatchActiveO365/BatchActiveWithoutChangeComputer.zip?raw=true

