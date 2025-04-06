# TaskTimerSharp
TaskTimer 是一个基于 Windows 服务的定时任务调度系统，支持按时间间隔或周期性地调用 API 接口、执行数据库任务，并可在执行成功后发送邮件通知。配置简单，功能灵活，适用于企业内部接口调度、定时任务管理等场景。

🔧 功能特性
⏱ 支持固定时间间隔和周期性执行（如每日、每周）

📆 支持配置开始时间和结束时间

🌐 可通过 HTTP 调用外部 API 接口，运行存储过程

📬 支持执行成功后自动发送邮件通知

🛡 可配置安全验证机制（如 API 密钥）

🧩 配置通过 XML 文件驱动，易于部署与修改

💻 基于 Windows 服务运行，稳定、后台执行

### 配置说明
#### 系统设置
默认再System文件夹中。多给此XML添加一个节点,系统会自动创建相应的文件夹[NameEN]。命名规则：TimeProjectJob.xml
```xml
<TimeProject>
  <TimeProjectJob>
    <NameCN>任务1</NameCN><!--job 中文名称-->
    <NameEN>Jobone</NameEN><!--job 英文名称 必填 创建文件夹只用这个名称 避免用特殊字符 可能会造成创建文件夹不成功-->
    <Status>True</Status><!--是否执行   执行:True  停止执行：False-->
    <Type>0</Type><!--执行状态  0：调用API  1: 顺序执行（必须有FUN_Sequence.xml 配合才能使用）  2：执行存储过程（暂时还没有）-->
  </TimeProjectJob>
  <TimeProjectJob>
    <NameCN>任务2</NameCN>
    <NameEN>Jobtwo</NameEN>
    <Status>false</Status>
    <Type>0</Type>
  </TimeProjectJob>
  <TimeProjectJob>
    <NameCN>任务3</NameCN>
    <NameEN>Jobthree</NameEN>
    <Status>false</Status>
    <Type>0</Type>
  </TimeProjectJob>
</TimeProject>
```

#### 单个执行的配置文件
在这个文件夹中加入这个XML文件,按规则设置好定时时间。 命名规则：NameEN_Setup.xml [文件夹的名字_Setup.xml]
```xml
<TimeProject>
  <TimeProjectJob>
    <StartTime>2020-01-14 16:47</StartTime> <!-- 可选，任务开始时间 -->
    <EndTime>2020-01-14 16:48</EndTime>     <!-- 可选，任务结束时间 -->
    <ExecutionStatus>0</ExecutionStatus>    <!-- 执行模式：0 为按间隔，1 为按周期 -->
    <IntervalsTime>30</IntervalsTime>       <!-- 执行间隔时间，单位：秒 -->
    <CycltType>EveryWeek</CycltType>        <!-- 周期类型，可选：EveryDay / EveryWeek -->
    <DayOfWeek>Monday</DayOfWeek>           <!-- 每周任务执行的具体星期几 -->
    <SpecificTime>05:00</SpecificTime>      <!-- 执行时间点（24小时制） -->
    <ApiUrl>Url</ApiUrl>                    <!-- 请求接口地址Get -->
    <Remark>获取任务1</Remark>               <!-- 描述任务用途 -->
    <MailTo>xxxx@gmail.com</MailTo>         <!-- 执行成功后通知的收件人 -->
    <SendMail>True</SendMail>               <!-- 是否发送通知邮件 -->
    <ConnStr></ConnStr>                     <!-- 可选数据库连接字符串 -->
    <Verification>False</Verification>      <!-- 是否启用 API 密钥验证 -->
    <AuthenticationKey></AuthenticationKey><!-- 安全验证密钥 -->
  </TimeProjectJob>
</TimeProject>
```

### 按顺序执行的配置文件
需要在"系统设置"将指定任务Type改为1
在文件加中再加如一个新的XML文件。命名规则：NameEN_Sequence.xml [文件夹的名字_Sequence.xml]
```xml
<TimeProject>
  <TimeProjectJob>
    <Project>API地址</Project>
    <Info>https://xxxx.com/api/datasync/shop</Info>
  </TimeProjectJob>
  <TimeProjectJob>
    <Project>API地址</Project>
    <Info>https://xxxx.com/api/datasync/synccategoriesandimg</Info>
  </TimeProjectJob>
  <TimeProjectJob>
    <Project>存储过程</Project>
    <Info>BackChangeProductTimeByItemMaster</Info>
  </TimeProjectJob>
</TimeProject>
```

### 邮件设置
在TimerProjectByWindowsService.exe.config 中配置
```xml
<?xml version="1.0" encoding="utf-8"?>
<configuration>
<startup><supportedRuntime version="v4.0" sku=".NETFramework,Version=v4.5"/></startup>
  <appSettings>
    <!--指定发送邮件的服务器地址或IP，如smtp.163.com-->
    <add key="MailHost" value="smtp.office365.com"/>
    <!--指定发送邮件端口 此值必须为数字，否则程序会出错-->
    <add key="Port" value="587"/>
    <!--发件人邮箱地址-->
    <add key="MailAddress" value="xxxx@163.com"/>
    <!--发件人邮箱用户名-->
    <add key="MailDisplayName" value="Test定時程序"/>
    <!--发件人邮箱密码-->
    <add key="MailPassWord" value="123445"/>
    <!--发件人是否要SSL加密-->
    <add key="SSL" value="true"/>
    <!--程序名称,表明当前是哪只定时程序-->
    <add key="ProgramName" value="Test定時程序BySiYang"/>
    <!--ACTION通知会加入以下邮件地址和程序配置地址一起发送-->
    <add key="SendTo" value="bbbb@163.com"/>
    <!--Info通知,是否加入SendTo配置地址一起发送 true:发送 false:不发送-->
    <add key="InfoMessage" value="true"/>
  </appSettings>
</configuration>
```

### 安装说明
所有操作，请在管理员模式运行，否则可能不成功

打开 .NET Framework

cd C:\Windows\Microsoft.NET\Framework64\v4.0.30319\

安装 运行 你的服务 EXE的路径
installutil.exe E:\test\TimerProjectByWindowsService.exe

卸载 运行 你的服务 EXE的路径
installutil.exe /u E:\test\TimerProjectByWindowsService.exe

启动服务
net start TimeProject 

关闭服务
net stop TimeProject 

将服务设为自动启动
sc config TimeProject start= auto

将服务设为手动启动
sc config TimeProject start= demand

打开服务窗体
CMD >  services.msc

##### 基本语法：
启动服务：net start 服务名  
停止服务：net stop 服务名  
暂停服务：net pause 服务名  
恢复被暂停的服务：net continue 服务名  
禁用服务：sc config 服务名 start=disabled  
将服务设为自动启动：sc config 服务名 start= auto  
将服务设为手动启动：sc config 服务名 start= demand  





