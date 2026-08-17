# TaskTimerSharp
TaskTimer 是一个基于 Windows 服务的定时任务调度系统，支持按时间间隔或周期性地调用 API 接口、执行数据库存储过程，并可在执行成功后发送邮件通知。配置简单，功能灵活，适用于企业内部接口调度、定时任务管理等场景。

### 功能特性
- 支持固定时间间隔和周期性执行（如每日、每周）
- 支持配置开始时间和结束时间
- 可通过 HTTP(GET) 调用外部 API 接口、运行存储过程（支持 API+存储过程 的顺序执行）
- 支持执行成功后自动发送邮件通知，执行异常发送 [ACTION] 告警邮件（同类错误自动节流，避免邮件风暴）
- 可配置安全验证机制（时间戳 + SHA1 签名）
- 支持失败自动重试（可配置重试次数与间隔）
- 执行历史持久化（Log/History 目录），服务重启后周期任务不会当天重复执行
- 配置通过 XML 文件驱动，修改任务配置无需重启服务（每秒自动加载）
- 基于 .NET Framework 4.8，Windows 服务后台运行

### 任务类型（Type）
| Type | 说明 |
|------|------|
| 0 | 调用单个 API（GET），接口地址配置在 `NameEN_Setup.xml` 的 `ApiUrl` |
| 1 | 顺序执行，步骤配置在 `NameEN_Sequence.xml`（API 与存储过程可混合，任一步失败立即中断后续步骤） |
| 2 | 独立执行存储过程，存储过程名配置在 `NameEN_Setup.xml` 的 `StoredProcedure`，并需配置 `ConnStr` |

### 配置说明
#### 系统设置
默认在 System 文件夹中。每给此 XML 添加一个节点，系统会自动创建相应的文件夹[NameEN]。命名规则：TimeProjectJob.xml
```xml
<TimeProject>
  <TimeProjectJob>
    <NameCN>任务1</NameCN><!--job 中文名称-->
    <NameEN>Jobone</NameEN><!--job 英文名称 必填 创建文件夹只用这个名称 避免用特殊字符 可能会造成创建文件夹不成功-->
    <Status>True</Status><!--是否执行   执行:True  停止执行：False-->
    <Type>0</Type><!--任务类型  0：调用API  1：顺序执行（必须有 NameEN_Sequence.xml 配合）  2：独立执行存储过程-->
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
在任务文件夹中放入此 XML 文件，按规则设置好定时时间。命名规则：NameEN_Setup.xml [文件夹的名字_Setup.xml]
```xml
<TimeProject>
  <TimeProjectJob>
    <StartTime>2020-01-14 16:47</StartTime> <!-- 可选，任务开始时间，格式 yyyy-MM-dd HH:mm -->
    <EndTime>2020-01-14 16:48</EndTime>     <!-- 可选，任务结束时间，格式 yyyy-MM-dd HH:mm -->
    <ExecutionStatus>0</ExecutionStatus>    <!-- 执行模式：0 为按间隔，1 为按周期 -->
    <IntervalsTime>30</IntervalsTime>       <!-- 执行间隔时间，单位：秒（执行模式0使用，从上次执行完成开始计时） -->
    <CycltType>EveryWeek</CycltType>        <!-- 周期类型，可选：EveryDay / EveryWeek（执行模式1使用） -->
    <DayOfWeek>Monday</DayOfWeek>           <!-- 每周任务执行的具体星期几，英文 Monday~Sunday（执行模式1使用） -->
    <SpecificTime>05:00</SpecificTime>      <!-- 执行时间点，24小时制 HH:mm（执行模式1使用） -->
    <ApiUrl>Url</ApiUrl>                    <!-- 请求接口地址Get（Type=0使用） -->
    <StoredProcedure></StoredProcedure>     <!-- 存储过程名称（Type=2使用，需同时配置ConnStr） -->
    <Remark>获取任务1</Remark>               <!-- 描述任务用途 -->
    <MailTo>xxxx@gmail.com</MailTo>         <!-- 执行成功后通知的收件人，多个用;分隔 -->
    <SendMail>True</SendMail>               <!-- 是否发送通知邮件 -->
    <ConnStr></ConnStr>                     <!-- 数据库连接字符串（顺序执行含存储过程步骤、或Type=2时必填） -->
    <Verification>False</Verification>      <!-- 是否启用 API 密钥验证 -->
    <AuthenticationKey></AuthenticationKey><!-- 安全验证密钥 -->
    <RetryCount>0</RetryCount>              <!-- 可选，失败重试次数，默认0不重试，最大10。注意：存储过程超时后重试可能导致重复执行，请确保存储过程可重入或幂等 -->
    <RetryInterval>60</RetryInterval>       <!-- 可选，重试间隔秒数，默认60 -->
    <HttpTimeout>600</HttpTimeout>          <!-- 可选，HTTP请求超时秒数，默认600 -->
  </TimeProjectJob>
</TimeProject>
```

#### 按顺序执行的配置文件
需要在"系统设置"将指定任务 Type 改为 1。在任务文件夹中加入一个新的 XML 文件，命名规则：NameEN_Sequence.xml [文件夹的名字_Sequence.xml]。
按配置顺序依次执行，任一步失败会立即中断后续步骤并发送告警邮件。
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
在 TimerProjectByWindowsService.exe.config 中配置。
> 安全提示：`MailPassWord` 不要提交真实密码到代码库。仓库中的 app.config 只保留占位符，部署时在服务器安装目录的 .exe.config 中填写真实密码。
```xml
<?xml version="1.0" encoding="utf-8"?>
<configuration>
<startup><supportedRuntime version="v4.0" sku=".NETFramework,Version=v4.8"/></startup>
  <appSettings>
    <!--指定发送邮件的服务器地址或IP，如smtp.163.com-->
    <add key="MailHost" value="smtp.office365.com"/>
    <!--指定发送邮件端口 此值必须为数字，否则程序会出错-->
    <add key="Port" value="587"/>
    <!--发件人邮箱地址-->
    <add key="MailAddress" value="xxxx@163.com"/>
    <!--发件人邮箱用户名-->
    <add key="MailDisplayName" value="Test定時程序"/>
    <!--发件人邮箱密码（部署时填写，勿提交到代码库）-->
    <add key="MailPassWord" value=""/>
    <!--发件人是否要SSL加密-->
    <add key="SSL" value="true"/>
    <!--程序名称,表明当前是哪只定时程序-->
    <add key="ProgramName" value="Test定時程序BySiYang"/>
    <!--ACTION通知会加入以下邮件地址和程序配置地址一起发送-->
    <add key="SendTo" value="bbbb@163.com"/>
    <!--Info通知,是否加入SendTo配置地址一起发送 true:发送 false:不发送-->
    <add key="InfoMessage" value="true"/>
    <!--可选：日志队列容量(默认20000)、节流窗口分钟(默认30)、单文件上限MB(默认50)-->
    <!--
    <add key="LogQueueSize" value="20000"/>
    <add key="LogThrottleMinutes" value="30"/>
    <add key="LogMaxMB" value="50"/>
    -->
  </appSettings>
</configuration>
```

### 运行日志与执行历史
- 运行日志：`安装目录\Log\任务名\yyyyMM\yyyyMMdd.log`
  - 由专用写入线程异步写入（有界队列，不阻塞调度线程），关机时自动排空，保证最后一条日志落盘
  - 重复消息自动节流（同任务同消息 30 分钟内只写一条，窗口过后补写被省略的次数）
  - 单文件超过 50MB 自动轮转（保留最近一个旧文件）
  - 写入失败（磁盘满/权限不足等）自动写入 Windows 事件查看器，不再静默丢失
  - 可选 appSettings 配置项（缺省值即可直接使用）：`LogQueueSize`（队列容量，默认 20000）、`LogThrottleMinutes`（节流窗口，默认 30 分钟）、`LogMaxMB`（单文件上限，默认 50MB）
- 执行历史：`安装目录\Log\History\任务名\yyyyMM.csv`，每次执行记录：时间、触发方式、耗时、状态（成功/失败）、结果摘要。服务重启后，周期任务会读取当天历史，避免同一天重复执行。配置错误导致的执行失败也会写入历史。

### 配置管理工具（TimerProjectConfigTool）
<img width="1100" height="720" alt="image" src="https://github.com/user-attachments/assets/b7a00ac9-27e7-429d-8d40-735d903970d4" />
图形化配置管理工具（WinForms，.NET Framework 4.8），与服务项目在同一解决方案中独立编译，互不引用。

**部署方式**：把编译产物 `TimerProjectConfigTool.exe` 和 `Languages` 文件夹拷贝到服务安装根目录（与 TimerProjectByWindowsService.exe 同级），双击运行即可。工具默认以自身所在目录为根目录，也可在界面上切换，切换后会记住该目录，下次启动自动沿用。

**功能**：
- 任务管理：新建/编辑/启用/停用任务（不提供删除，只能停用），自动创建任务文件夹与 XML 配置；保存后服务每秒自动加载，无需重启。**新建任务默认停用**，请在列表中手动点[启用]
- 一致性检查：任务列表会标出"已注册，缺任务文件夹"/"有任务文件夹，未注册"等异常（鼠标悬浮有说明），打开编辑并保存一次即可自动修复
- 编辑表单按选项联动显隐字段，提供 [测试连接]（验证数据库连接串）
- 日志浏览：树状列出运行日志（含 System 服务级日志）与执行历史，单击预览、双击用系统默认程序打开
- 邮件设置：修改 exe.config 的 appSettings（密码掩码显示），提供 [测试发送]；保存后需重启服务才生效，界面会提示并支持一键重启
- 多语言：外置 XML 语言包（内置简体中文/English），新增语言只需在 Languages 目录添加 XML 文件，无需重新编译

**注意**：重启服务需要管理员权限。若非管理员运行，工具会提示"以管理员身份重新运行本工具"或手动执行 `net stop TimeProject && net start TimeProject`。服务未安装的机器上重启按钮自动禁用。

### 安装说明
所有操作，请在管理员模式CMD运行，否则可能不成功

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
