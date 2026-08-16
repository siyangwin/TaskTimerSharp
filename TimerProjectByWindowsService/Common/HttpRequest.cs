using System;
using System.IO;
using System.Net;
using System.Text;

namespace TimerProjectByWindowsService.Common
{
    public class HttpRequest
    {
        private static readonly object ProtocolSync = new object();
        private static bool _protocolSet;

        //确保启用TLS1.2（老版本.NET默认可能不含），只设置一次
        private static void EnsureSecurityProtocol()
        {
            if (_protocolSet) return;
            lock (ProtocolSync)
            {
                if (_protocolSet) return;
                ServicePointManager.SecurityProtocol |= SecurityProtocolType.Tls | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls12;
                _protocolSet = true;
            }
        }

        #region 发送post请求
        /// <summary>
        /// Post
        /// </summary>
        /// <param name="Url">请求接口</param>
        /// <param name="str">请求参数</param>
        /// <returns></returns>
        public static string Post(string Url, string str)
        {
            EnsureSecurityProtocol();

            HttpWebRequest req = (HttpWebRequest)WebRequest.Create(Url);
            req.Method = "POST";
            req.ContentType = "application/x-www-form-urlencoded";

            byte[] data = Encoding.UTF8.GetBytes(str);//把字符串转换为字节
            req.ContentLength = data.Length; //请求长度

            using (Stream reqStream = req.GetRequestStream()) //获取
            {
                reqStream.Write(data, 0, data.Length);//向当前流中写入字节
            }

            using (HttpWebResponse resp = (HttpWebResponse)req.GetResponse()) //响应结果
            {
                return ReadResponseBody(resp);
            }
        }
        #endregion

        #region 发送GET请求
        /// <summary>
        /// GET请求与获取结果。非2xx状态码会抛出带状态码和响应内容的异常。
        /// </summary>
        /// <param name="Url">请求地址</param>
        /// <param name="postDataStr">参数</param>
        /// <param name="Timeout">超时时间  默认请传0 单位毫秒</param>
        /// <returns></returns>
        public static string HttpGet(string Url, string postDataStr, int Timeout)
        {
            EnsureSecurityProtocol();

            string fullUrl = Url + (string.IsNullOrEmpty(postDataStr) ? "" : (Url.Contains("?") ? "&" : "?") + postDataStr);
            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(fullUrl);
            request.Timeout = Timeout > 0 ? Timeout : 100 * 1000;
            request.ReadWriteTimeout = request.Timeout;
            request.Method = "GET"; //设置请求方式
            request.ContentType = "text/html;charset=UTF-8"; //设置内容类型

            try
            {
                using (HttpWebResponse response = (HttpWebResponse)request.GetResponse()) //返回响应
                {
                    int code = (int)response.StatusCode;
                    string body = ReadResponseBody(response);
                    if (code >= 400)
                    {
                        throw new Exception("HTTP请求返回错误状态码 " + code + "：" + LocalFile.Truncate(body, 500));
                    }
                    return body;
                }
            }
            catch (WebException ex)
            {
                //4xx/5xx时GetResponse直接抛WebException，把响应体读出来方便排查
                using (HttpWebResponse resp = ex.Response as HttpWebResponse)
                {
                    if (resp != null)
                    {
                        string body = "";
                        try { body = ReadResponseBody(resp); } catch { }
                        throw new Exception("HTTP请求失败，状态码 " + (int)resp.StatusCode + "：" + LocalFile.Truncate(body, 500), ex);
                    }
                }
                throw;
            }
        }
        #endregion

        /// <summary>
        /// 按响应声明的编码读取响应体，无法识别时按UTF8
        /// </summary>
        private static string ReadResponseBody(HttpWebResponse response)
        {
            Encoding encoding = Encoding.UTF8;
            if (!string.IsNullOrEmpty(response.CharacterSet))
            {
                try { encoding = Encoding.GetEncoding(response.CharacterSet); }
                catch { encoding = Encoding.UTF8; }
            }

            using (Stream stream = response.GetResponseStream())
            using (StreamReader reader = new StreamReader(stream, encoding))
            {
                return reader.ReadToEnd();
            }
        }
    }
}
