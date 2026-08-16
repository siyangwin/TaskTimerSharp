using System;
using System.Collections.Generic;
using System.IO;
using System.Xml;

namespace TimerProjectConfigTool
{
    public class LangInfo
    {
        public string Code;
        public string DisplayName;

        public override string ToString()
        {
            return DisplayName;
        }
    }

    /// <summary>
    /// 外置 XML 语言包：Languages/&lt;代码&gt;.xml。缺失键回退 zh-CN，再缺失返回键名本身。
    /// </summary>
    public static class Lang
    {
        private static readonly string LangDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Languages");
        private static Dictionary<string, string> _strings = new Dictionary<string, string>(StringComparer.Ordinal);
        private static Dictionary<string, string> _fallback = new Dictionary<string, string>(StringComparer.Ordinal);
        private static string _code = "zh-CN";

        public static string CurrentCode
        {
            get { return _code; }
        }

        public static event EventHandler LanguageChanged;

        public static List<LangInfo> GetAvailableLanguages()
        {
            var list = new List<LangInfo>();
            try
            {
                if (Directory.Exists(LangDir))
                {
                    foreach (string file in Directory.GetFiles(LangDir, "*.xml"))
                    {
                        string code = Path.GetFileNameWithoutExtension(file);
                        string name = code;
                        try
                        {
                            XmlDocument doc = new XmlDocument();
                            doc.Load(file);
                            if (doc.DocumentElement != null)
                            {
                                string attr = doc.DocumentElement.GetAttribute("name");
                                if (!string.IsNullOrEmpty(attr)) name = attr;
                            }
                        }
                        catch
                        {
                        }
                        list.Add(new LangInfo { Code = code, DisplayName = name });
                    }
                }
            }
            catch
            {
            }
            if (list.Count == 0) list.Add(new LangInfo { Code = "zh-CN", DisplayName = "简体中文" });
            return list;
        }

        public static void Load(string code)
        {
            if (string.IsNullOrEmpty(code)) code = "zh-CN";
            _strings = LoadFile(Path.Combine(LangDir, code + ".xml"));
            _fallback = code == "zh-CN" ? _strings : LoadFile(Path.Combine(LangDir, "zh-CN.xml"));
            _code = code;
            if (LanguageChanged != null) LanguageChanged(null, EventArgs.Empty);
        }

        private static Dictionary<string, string> LoadFile(string path)
        {
            var dict = new Dictionary<string, string>(StringComparer.Ordinal);
            try
            {
                if (File.Exists(path))
                {
                    XmlDocument doc = new XmlDocument();
                    doc.Load(path);
                    XmlNodeList nodes = doc.SelectNodes("Language/String");
                    if (nodes != null)
                    {
                        foreach (XmlNode n in nodes)
                        {
                            if (n.Attributes == null) continue;
                            XmlAttribute key = n.Attributes["key"];
                            if (key != null && !string.IsNullOrEmpty(key.Value))
                            {
                                dict[key.Value] = n.InnerText;
                            }
                        }
                    }
                }
            }
            catch
            {
            }
            return dict;
        }

        public static string Get(string key)
        {
            string v;
            if (_strings.TryGetValue(key, out v) && v.Length > 0) return v;
            if (_fallback != null && _fallback.TryGetValue(key, out v) && v.Length > 0) return v;
            return key;
        }

        public static string Get(string key, params object[] args)
        {
            string template = Get(key);
            try
            {
                return string.Format(template, args);
            }
            catch
            {
                return template;
            }
        }
    }

    /// <summary>
    /// 工具自身设置（语言选择、根目录），保存在工具目录 TimerProjectConfigTool.settings.xml
    /// </summary>
    public static class SettingsStore
    {
        private static readonly string SettingsPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "TimerProjectConfigTool.settings.xml");

        public static string LoadLanguage()
        {
            string v = LoadSetting("Language");
            return v.Length > 0 ? v : "zh-CN";
        }

        public static void SaveLanguage(string code)
        {
            SaveSetting("Language", code);
        }

        public static string LoadRootDir()
        {
            return LoadSetting("RootDir");
        }

        public static void SaveRootDir(string path)
        {
            SaveSetting("RootDir", path);
        }

        private static string LoadSetting(string elementName)
        {
            try
            {
                if (File.Exists(SettingsPath))
                {
                    XmlDocument doc = new XmlDocument();
                    doc.Load(SettingsPath);
                    XmlNode n = doc.SelectSingleNode("Settings/" + elementName);
                    if (n != null && !string.IsNullOrWhiteSpace(n.InnerText)) return n.InnerText.Trim();
                }
            }
            catch
            {
            }
            return "";
        }

        private static void SaveSetting(string elementName, string value)
        {
            try
            {
                XmlDocument doc = new XmlDocument();
                if (File.Exists(SettingsPath))
                {
                    doc.Load(SettingsPath);
                }
                else
                {
                    doc.AppendChild(doc.CreateXmlDeclaration("1.0", "utf-8", null));
                }
                XmlNode root = doc.SelectSingleNode("Settings");
                if (root == null)
                {
                    root = doc.CreateElement("Settings");
                    doc.AppendChild(root);
                }
                XmlNode el = root.SelectSingleNode(elementName);
                if (el == null)
                {
                    el = doc.CreateElement(elementName);
                    root.AppendChild(el);
                }
                el.InnerText = value ?? "";
                doc.Save(SettingsPath);
            }
            catch
            {
            }
        }
    }
}
