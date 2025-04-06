using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace TimerProjectByWindowsService.Common
{
    /// <summary>
    /// 设置本地目录
    /// </summary>
    public class LocalFile
    {

        //设置日志本地根目录
        //string path = Path.Combine(Directory.GetCurrentDirectory(), @"Log");  //winfrom可以用这个
        string path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"Log");  //服务用这个

        //设置配置文件本地根目录
        string PathConfig = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"Config");  //服务用这个

        #region Log文件夹操作
        /// <summary>
        /// 获取Log写入路径
        /// </summary>
        /// <returns></returns>
        public string LocalLog()
        {
            string LogPath = path;
            //判断文件夹是否存在
            if (!Directory.Exists(LogPath))
            {
                //不存在则添加
                Directory.CreateDirectory(LogPath);//不存在就创建文件夹 
            }
            return LogPath;
        }

        /// <summary>
        /// 获取Config写入路径
        /// </summary>
        /// <returns></returns>
        public string LocalConfig()
        {
            string ConfigPath = PathConfig;
            //判断文件夹是否存在
            if (!Directory.Exists(ConfigPath))
            {
                //不存在则添加
                Directory.CreateDirectory(ConfigPath);//不存在就创建文件夹 
            }
            return ConfigPath;
        }

        /// <summary>
        /// 检测指定路径是否存在
        /// </summary>
        /// <param name="path">路径</param>
        /// <param name="isDirectory">是否是目录</param>
        /// <returns></returns>
        public bool IsExist(string path, bool isDirectory)
        {
            return isDirectory ? Directory.Exists(path) : File.Exists(path);
        }


        /// <summary>
        /// 检测指定文件夹是否存在 不存在则创建
        /// </summary>
        /// <param name="path">路径</param>
        /// <returns></returns>
        public bool DirectoryIsExist(string path)
        {
            bool res = false;
            //判断文件是否存在
            if (!Directory.Exists(path))
            {
                //不存在则添加
                Directory.CreateDirectory(path);//不存在就创建文件夹 
                res = true;
            }
            else
            {
                res = true;
            }
            return res;
        }


        /// <summary> 
        /// 检测指定文件是否存在  不存在则创建
        /// </summary>
        /// <param name="path">路径</param>
        /// <returns></returns>
        public bool FileIsExist(string path)
        {
            bool res = false;
            //判断文件是否存在
            if (!File.Exists(path))
            {
                //由于文件被占用不能读写，所以报错“另一个程序正在使用此文件进程无法访问”
                //解决方法是在创建文件后立即Dispose掉
                File.Create(path).Dispose();
                res = true;
            }
            else
            {
                res = true;
            }
            return res;
        }
        #endregion

        #region  XML操作类

        /// <summary>
        /// 读取最大的Xml  判断有多少需要执行的JOB
        /// </summary>
        /// <param name="FilePath">文件的保存路径</param>
        /// <param name="FileName">文件名称</param>
        /// <returns></returns>
        public virtual DataTable GetXmlInfo(string FilePath, string FileName)
        {
            if (File.Exists(FilePath + "\\" + FileName))
            {
                DataTable DT = new DataTable();
                XmlDocument xml = new XmlDocument();
                xml.Load(FilePath + "\\" + FileName);
                XmlNodeList xnl = xml.SelectSingleNode("TimeProject").ChildNodes;
                if (xnl.Count > 0)
                {
                    DT = CreateDataTable(xnl[0]);
                }
                foreach (XmlNode xns in xnl)
                {
                    DataRow DR = DT.NewRow();
                    for (int i = 0; i < xns.ChildNodes.Count; i++)
                    {
                        if (xns.ChildNodes[i].NodeType != XmlNodeType.Comment) // 判断不等于注释时
                        {
                            DR[xns.ChildNodes[i].Name] = xns.ChildNodes[i].InnerText;
                        }
                    }
                    DT.Rows.Add(DR);
                }
                return DT;
            }
            return new DataTable();
        }


        /// <summary>
        /// 保存mail
        /// </summary>
        /// <param name="DT">Mail的信息集合</param>
        /// <returns></returns>
        public virtual void SetXmlInfo(DataTable DT, string FilePath, string FileName)
        {
            //判断路径是否存在
            if (!DirectoryIsExist(FilePath))
            {
                return;
            }

            //如果数据为空,则删除文件
            if (DT == null || DT.Rows.Count == 0)
            {
                if (File.Exists(FilePath + "\\" + FileName))
                {
                    File.Delete(FilePath + "\\" + FileName);
                }
                return;
            }

            if (File.Exists(FilePath + "\\" + FileName))
            {
                File.Delete(FilePath + "\\" + FileName);
            }



            XmlDocument xml = new XmlDocument();
            XmlElement topXmlElement = xml.CreateElement("Mail");//创建最顶端的节点
            for (int i = 0; i < DT.Rows.Count; i++)
            {
                XmlElement X1 = xml.CreateElement("UserMail");
                foreach (DataColumn DC in DT.Columns)
                {
                    XmlElement xx = xml.CreateElement(DC.ColumnName);
                    xx.InnerText = DT.Rows[i][DC].ToString();
                    X1.AppendChild(xx);
                }
                topXmlElement.AppendChild(X1);
            }
            xml.AppendChild(topXmlElement);
            xml.Save(FilePath + "\\" + FileName);
        }


        /// <summary>
        /// 将XML转换为DataTable
        /// </summary>
        /// <param name="xn"></param>
        /// <returns></returns>
        private DataTable CreateDataTable(XmlNode xn)
        {
            System.Data.DataTable DT = new System.Data.DataTable();
            for (int i = 0; i < xn.ChildNodes.Count; i++)
            {
                if (xn.ChildNodes[i].NodeType != XmlNodeType.Comment) // 判断不等于注释时
                {
                    DT.Columns.Add(xn.ChildNodes[i].Name);
                }
            }
            return DT;
        }
        #endregion

        #region 安全验证必须
        /// <summary>
        /// 获取时间戳  十位(秒)
        /// </summary>
        /// <returns></returns>
        public static string GetTimeStampBySeconds()
        {
            TimeSpan ts = DateTime.Now - new DateTime(1970, 1, 1, 0, 0, 0, 0);
            return Convert.ToInt64(ts.TotalSeconds).ToString();
        }

        /// <summary>
        /// 返回SHA1
        /// </summary>
        /// <param name="key"></param>
        /// <returns></returns>
        public static string GetSha1(string key)
        {
            //SHA1加密方法
            var sha1 = new SHA1CryptoServiceProvider();
            byte[] str01 = Encoding.Default.GetBytes(key);
            byte[] str02 = sha1.ComputeHash(str01);
            var result = BitConverter.ToString(str02).Replace("-", "");
            return result;
        }



        /// <summary>
        /// 获取时间戳  十三位(毫秒)
        /// </summary>
        /// <returns></returns>
        public static string GetTimeStampByMilliseconds()
        {
            TimeSpan ts = DateTime.UtcNow - new DateTime(1970, 1, 1, 0, 0, 0, 0);
            return Convert.ToInt64(ts.TotalMilliseconds).ToString();
        }

        #endregion

    }
}
