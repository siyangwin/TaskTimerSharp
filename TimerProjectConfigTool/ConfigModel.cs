using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Xml;

namespace TimerProjectConfigTool
{
    /// <summary>System/TimeProjectJob.xml 中的一条任务注册记录</summary>
    public class JobRegistryEntry
    {
        public string NameCN = "";
        public string NameEN = "";
        public bool Enabled;
        public string Type = "0";
    }

    /// <summary>&lt;NameEN&gt;_Setup.xml 全字段（与服务端 GetXmlInfo 兼容）</summary>
    public class JobSetup
    {
        public string StartTime = "";
        public string EndTime = "";
        public string ExecutionStatus = "0";
        public string IntervalsTime = "60";
        public string CycltType = "EveryDay";
        public string DayOfWeek = "Monday";
        public string SpecificTime = "08:00";
        public string ApiUrl = "";
        public string StoredProcedure = "";
        public string Remark = "";
        public string MailTo = "";
        public string SendMail = "False";
        public string ConnStr = "";
        public string Verification = "False";
        public string AuthenticationKey = "";
        public string RetryCount = "0";
        public string RetryInterval = "60";
        public string HttpTimeout = "600";
    }

    /// <summary>&lt;NameEN&gt;_Sequence.xml 中的一个顺序步骤（Project 只能是 API地址/存储过程）</summary>
    public class SequenceStep
    {
        public string Project = "API地址";
        public string Info = "";
    }

    /// <summary>主窗体任务列表行：注册表 ∪ 实际文件夹</summary>
    public class JobOverviewRow
    {
        public string NameCN = "";
        public string NameEN = "";
        public string Type = "";
        public bool InRegistry;
        public bool Enabled;
        public bool FolderExists;
        public bool SetupExists;
        public bool SequenceExists;
        public string IssueKey = "";
    }

    public static class ConfigModel
    {
        public const string ServiceConfigFileName = "TimerProjectByWindowsService.exe.config";
        public const string SystemXmlFileName = "TimeProjectJob.xml";
        private static readonly Encoding Utf8NoBom = new UTF8Encoding(false);

        // ---------- 路径 ----------
        public static string ConfigDir(string root) { return Path.Combine(root, "Config"); }
        public static string SystemDir(string root) { return Path.Combine(ConfigDir(root), "System"); }
        public static string SystemXmlPath(string root) { return Path.Combine(SystemDir(root), SystemXmlFileName); }
        public static string JobDir(string root, string nameEN) { return Path.Combine(ConfigDir(root), nameEN); }
        public static string SetupPath(string root, string nameEN) { return Path.Combine(JobDir(root, nameEN), nameEN + "_Setup.xml"); }
        public static string SequencePath(string root, string nameEN) { return Path.Combine(JobDir(root, nameEN), nameEN + "_Sequence.xml"); }
        public static string LogDir(string root) { return Path.Combine(root, "Log"); }
        public static string HistoryDir(string root) { return Path.Combine(LogDir(root), "History"); }
        public static string ServiceConfigPath(string root) { return Path.Combine(root, ServiceConfigFileName); }

        // ---------- 注册表 System.xml ----------
        public static List<JobRegistryEntry> LoadRegistry(string root)
        {
            var list = new List<JobRegistryEntry>();
            string path = SystemXmlPath(root);
            if (!File.Exists(path)) return list;

            XmlDocument doc = new XmlDocument();
            doc.Load(path);
            XmlNode rootNode = doc.SelectSingleNode("TimeProject");
            if (rootNode == null) return list;

            foreach (XmlNode node in rootNode.ChildNodes)
            {
                if (node.NodeType != XmlNodeType.Element) continue;
                var e = new JobRegistryEntry
                {
                    NameCN = ChildText(node, "NameCN"),
                    NameEN = ChildText(node, "NameEN"),
                    Enabled = string.Equals(ChildText(node, "Status"), "true", StringComparison.OrdinalIgnoreCase),
                    Type = ChildText(node, "Type")
                };
                list.Add(e);
            }
            return list;
        }

        public static void SaveRegistry(string root, List<JobRegistryEntry> entries)
        {
            XmlDocument doc = NewDoc();
            XmlElement rootEl = doc.CreateElement("TimeProject");
            foreach (var e in entries)
            {
                XmlElement job = doc.CreateElement("TimeProjectJob");
                AppendField(doc, job, "NameCN", e.NameCN);
                AppendField(doc, job, "NameEN", e.NameEN);
                AppendField(doc, job, "Status", e.Enabled ? "True" : "False");
                AppendField(doc, job, "Type", e.Type);
                rootEl.AppendChild(job);
            }
            doc.AppendChild(rootEl);
            AtomicWriteXml(SystemXmlPath(root), doc);
        }

        // ---------- Setup.xml ----------
        public static JobSetup LoadSetup(string root, string nameEN)
        {
            string path = SetupPath(root, nameEN);
            if (!File.Exists(path)) return new JobSetup();

            XmlDocument doc = new XmlDocument();
            doc.Load(path);
            XmlNode job = FirstJobNode(doc);
            if (job == null) return new JobSetup();

            var s = new JobSetup
            {
                StartTime = ChildText(job, "StartTime"),
                EndTime = ChildText(job, "EndTime"),
                ExecutionStatus = ChildText(job, "ExecutionStatus"),
                IntervalsTime = ChildText(job, "IntervalsTime"),
                CycltType = ChildText(job, "CycltType"),
                DayOfWeek = ChildText(job, "DayOfWeek"),
                SpecificTime = ChildText(job, "SpecificTime"),
                ApiUrl = ChildText(job, "ApiUrl"),
                StoredProcedure = ChildText(job, "StoredProcedure"),
                Remark = ChildText(job, "Remark"),
                MailTo = ChildText(job, "MailTo"),
                SendMail = ChildText(job, "SendMail"),
                ConnStr = ChildText(job, "ConnStr"),
                Verification = ChildText(job, "Verification"),
                AuthenticationKey = ChildText(job, "AuthenticationKey"),
                RetryCount = ChildText(job, "RetryCount"),
                RetryInterval = ChildText(job, "RetryInterval"),
                HttpTimeout = ChildText(job, "HttpTimeout")
            };
            return s;
        }

        public static void SaveSetup(string root, string nameEN, JobSetup s)
        {
            XmlDocument doc = NewDoc();
            XmlElement rootEl = doc.CreateElement("TimeProject");
            XmlElement job = doc.CreateElement("TimeProjectJob");
            AppendField(doc, job, "StartTime", s.StartTime);
            AppendField(doc, job, "EndTime", s.EndTime);
            AppendField(doc, job, "ExecutionStatus", s.ExecutionStatus);
            AppendField(doc, job, "IntervalsTime", s.IntervalsTime);
            AppendField(doc, job, "CycltType", s.CycltType);
            AppendField(doc, job, "DayOfWeek", s.DayOfWeek);
            AppendField(doc, job, "SpecificTime", s.SpecificTime);
            AppendField(doc, job, "ApiUrl", s.ApiUrl);
            AppendField(doc, job, "StoredProcedure", s.StoredProcedure);
            AppendField(doc, job, "Remark", s.Remark);
            AppendField(doc, job, "MailTo", s.MailTo);
            AppendField(doc, job, "SendMail", s.SendMail);
            AppendField(doc, job, "ConnStr", s.ConnStr);
            AppendField(doc, job, "Verification", s.Verification);
            AppendField(doc, job, "AuthenticationKey", s.AuthenticationKey);
            AppendField(doc, job, "RetryCount", s.RetryCount);
            AppendField(doc, job, "RetryInterval", s.RetryInterval);
            AppendField(doc, job, "HttpTimeout", s.HttpTimeout);
            rootEl.AppendChild(job);
            doc.AppendChild(rootEl);
            AtomicWriteXml(SetupPath(root, nameEN), doc);
        }

        // ---------- Sequence.xml ----------
        public static List<SequenceStep> LoadSequence(string root, string nameEN)
        {
            var list = new List<SequenceStep>();
            string path = SequencePath(root, nameEN);
            if (!File.Exists(path)) return list;

            XmlDocument doc = new XmlDocument();
            doc.Load(path);
            XmlNode rootNode = doc.SelectSingleNode("TimeProject");
            if (rootNode == null) return list;

            foreach (XmlNode node in rootNode.ChildNodes)
            {
                if (node.NodeType != XmlNodeType.Element) continue;
                list.Add(new SequenceStep
                {
                    Project = ChildText(node, "Project"),
                    Info = ChildText(node, "Info")
                });
            }
            return list;
        }

        public static void SaveSequence(string root, string nameEN, List<SequenceStep> steps)
        {
            XmlDocument doc = NewDoc();
            XmlElement rootEl = doc.CreateElement("TimeProject");
            foreach (var step in steps)
            {
                XmlElement job = doc.CreateElement("TimeProjectJob");
                AppendField(doc, job, "Project", step.Project);
                AppendField(doc, job, "Info", step.Info);
                rootEl.AppendChild(job);
            }
            doc.AppendChild(rootEl);
            AtomicWriteXml(SequencePath(root, nameEN), doc);
        }

        // ---------- 任务总览（注册表 ∪ 文件夹） ----------
        public static List<JobOverviewRow> LoadOverview(string root)
        {
            var rows = new List<JobOverviewRow>();
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (var e in LoadRegistry(root))
            {
                if (string.IsNullOrWhiteSpace(e.NameEN)) continue;
                var row = new JobOverviewRow
                {
                    NameCN = e.NameCN,
                    NameEN = e.NameEN,
                    Type = e.Type,
                    InRegistry = true,
                    Enabled = e.Enabled,
                    FolderExists = Directory.Exists(JobDir(root, e.NameEN)),
                    SetupExists = File.Exists(SetupPath(root, e.NameEN)),
                    SequenceExists = File.Exists(SequencePath(root, e.NameEN))
                };
                if (!row.FolderExists) row.IssueKey = "main.issueNoFolder";
                rows.Add(row);
                seen.Add(e.NameEN);
            }

            try
            {
                if (Directory.Exists(ConfigDir(root)))
                {
                    foreach (string dir in Directory.GetDirectories(ConfigDir(root)))
                    {
                        string name = Path.GetFileName(dir);
                        if (string.Equals(name, "System", StringComparison.OrdinalIgnoreCase)) continue;
                        if (seen.Contains(name)) continue;
                        rows.Add(new JobOverviewRow
                        {
                            NameEN = name,
                            InRegistry = false,
                            FolderExists = true,
                            SetupExists = File.Exists(SetupPath(root, name)),
                            SequenceExists = File.Exists(SequencePath(root, name)),
                            IssueKey = "main.issueNotRegistered"
                        });
                    }
                }
            }
            catch
            {
            }
            return rows;
        }

        // ---------- 新建任务 ----------
        public static bool IsValidNameEN(string name)
        {
            return !string.IsNullOrWhiteSpace(name) && Regex.IsMatch(name, @"^[A-Za-z0-9_]+$");
        }

        public static bool NameENExists(string root, string nameEN)
        {
            if (Directory.Exists(JobDir(root, nameEN))) return true;
            foreach (var e in LoadRegistry(root))
            {
                if (string.Equals(e.NameEN, nameEN, StringComparison.OrdinalIgnoreCase)) return true;
            }
            return false;
        }

        public static void CreateJob(string root, string nameCN, string nameEN, string type, bool enabled)
        {
            Directory.CreateDirectory(JobDir(root, nameEN));
            SaveSetup(root, nameEN, new JobSetup());
            if (type == "1")
            {
                SaveSequence(root, nameEN, new List<SequenceStep> { new SequenceStep() });
            }
            var registry = LoadRegistry(root);
            registry.Add(new JobRegistryEntry { NameCN = nameCN, NameEN = nameEN, Enabled = enabled, Type = type });
            SaveRegistry(root, registry);
        }

        // ---------- 原子写入 ----------
        /// <summary>先写临时文件再替换，避免服务每秒读取时读到半个文件；IOException 重试 3 次</summary>
        public static void AtomicWrite(string path, string content)
        {
            string dir = Path.GetDirectoryName(path);
            if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir)) Directory.CreateDirectory(dir);

            string tmp = path + ".tmp_" + Guid.NewGuid().ToString("N");
            try
            {
                File.WriteAllText(tmp, content, Utf8NoBom);
                AtomicReplace(tmp, path);
            }
            finally
            {
                try { if (File.Exists(tmp)) File.Delete(tmp); } catch { }
            }
        }

        /// <summary>把已写好的临时文件替换到目标路径（带重试）</summary>
        public static void AtomicReplace(string tmpPath, string path)
        {
            string dir = Path.GetDirectoryName(path);
            if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir)) Directory.CreateDirectory(dir);

            for (int attempt = 1; ; attempt++)
            {
                try
                {
                    if (File.Exists(path)) File.Replace(tmpPath, path, null);
                    else File.Move(tmpPath, path);
                    return;
                }
                catch (IOException)
                {
                    if (attempt >= 3) throw;
                    Thread.Sleep(100 * attempt);
                }
                catch (UnauthorizedAccessException)
                {
                    if (attempt >= 3) throw;
                    Thread.Sleep(100 * attempt);
                }
            }
        }

        private static void AtomicWriteXml(string path, XmlDocument doc)
        {
            string dir = Path.GetDirectoryName(path);
            if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir)) Directory.CreateDirectory(dir);

            // 必须写到字节流而不是 StringBuilder：StringBuilder 上 XmlWriter 的声明编码会被写成 utf-16，
            // 而文件字节其实是 UTF-8，解析方会因声明与字节不符而报错。
            string tmp = path + ".tmp_" + Guid.NewGuid().ToString("N");
            try
            {
                var settings = new XmlWriterSettings { Indent = true, Encoding = Utf8NoBom };
                using (var fs = new FileStream(tmp, FileMode.Create, FileAccess.Write, FileShare.None))
                using (var xw = XmlWriter.Create(fs, settings))
                {
                    doc.Save(xw);
                }
                AtomicReplace(tmp, path);
            }
            finally
            {
                try { if (File.Exists(tmp)) File.Delete(tmp); } catch { }
            }
        }

        // ---------- XML 辅助 ----------
        private static XmlDocument NewDoc()
        {
            XmlDocument doc = new XmlDocument();
            doc.AppendChild(doc.CreateXmlDeclaration("1.0", "utf-8", null));
            return doc;
        }

        private static XmlNode FirstJobNode(XmlDocument doc)
        {
            XmlNode rootNode = doc.SelectSingleNode("TimeProject");
            if (rootNode == null) return null;
            foreach (XmlNode n in rootNode.ChildNodes)
            {
                if (n.NodeType == XmlNodeType.Element) return n;
            }
            return null;
        }

        private static string ChildText(XmlNode parent, string name)
        {
            XmlNode n = parent.SelectSingleNode(name);
            return n == null ? "" : (n.InnerText ?? "").Trim();
        }

        private static void AppendField(XmlDocument doc, XmlElement parent, string name, string value)
        {
            XmlElement el = doc.CreateElement(name);
            el.InnerText = value ?? "";
            parent.AppendChild(el);
        }
    }
}
