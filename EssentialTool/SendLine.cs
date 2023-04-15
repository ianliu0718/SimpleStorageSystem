using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Reflection.Emit;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace 簡易倉儲系統.EssentialTool
{
    internal class SendLine
    {
        List<stLineToken> ListLineToken = null;
        class stLineToken
        {
            public String Send_Group { get; set; }
            public String LineToken { get; set; }
        }
        public SendLine()
        {
            if (!Directory.Exists("ExcelJpg"))
                Directory.CreateDirectory("ExcelJpg");
            if (!Directory.Exists("log"))
                Directory.CreateDirectory("log");
            ListLineToken.Add(new stLineToken() { Send_Group = "Ian", LineToken = "PkOjQVn809ZiLtwkmnZqGPy8WmZYnnCsxDfdLLCptlc" });
        }

        public static string SendLineMessage(String LineToken, String Message)
        {
            String result;
            try
            {
                string Url = "https://notify-api.line.me/api/notify";
                HttpWebRequest request = (HttpWebRequest)WebRequest.Create(Url);
                request.Method = "POST";
                request.KeepAlive = true; //是否保持連線
                request.ContentType = "application/x-www-form-urlencoded";
                request.Headers.Set("Authorization", "Bearer " + LineToken);
                string content = "";
                content += "message=" + Message.Replace("%", "").Replace("&", "");//發送的文字訊息內容

                byte[] bytes = Encoding.UTF8.GetBytes(content);
                using (var stream = request.GetRequestStream())
                {
                    stream.Write(bytes, 0, bytes.Length);
                }

                var response = (HttpWebResponse)request.GetResponse();
                using (var streamReader = new StreamReader(response.GetResponseStream()))
                {
                    result = streamReader.ReadToEnd();
                    if (result != "{\"status\":200,\"message\":\"ok\"}")
                        throw new Exception("LINE 伺服器回傳為正確收到訊息，訊息內容為：\r\n" + result);
                }
                return result;
            }
            catch (Exception)
            {
                throw;
            }

        }

        /// <summary>
        /// 發送有附件檔的LINE通知訊息
        /// </summary>
        /// <param name="LineToken">LINE TOKEN</param>
        /// <param name="Message">通知訊息</param>
        /// <param name="FileFullPath">附件檔，僅能圖片檔(jpg,jpeg,png)</param>
        /// <returns>LINE伺服器回傳訊息</returns>
        public string SendLineMessage(String LineToken, String Message, String FileFullPath)
        {
            String ApiURL = "https://notify-api.line.me/api/notify";
            String FileParameter = "imageFile";
            string strToken = LineToken;
            string _ApiURL, ReturnValue = "";
            HttpWebRequest req;
            StringBuilder LogMessage = new StringBuilder();
            WebResponse wr;
            StreamReader myStreamReader;
            string boundary = "---------------------------" + DateTime.Now.Ticks.ToString("x");
            byte[] boundarybytes = System.Text.Encoding.UTF8.GetBytes("\r\n--" + boundary + "\r\n");
            string formdataTemplate = "Content-Disposition: form-data; name=\"{0}\"\r\n\r\n{1}";
            string headerTemplate = "Content-Disposition: form-data; name=\"{0}\"; filename=\"{1}\"\r\nContent-Type: {2}\r\n\r\n";
            string header;
            byte[] headerbytes;
            string contentType;
            FileInfo getFileExtension;
            Stream rs;
            FileStream fileStream;
            _ApiURL = ApiURL;

            try
            {
                req = (HttpWebRequest)WebRequest.Create(_ApiURL);//宣告並配置網站連線，設定網址
                req.Method = "POST";
                //req.KeepAlive = false;
                req.Timeout = -1;
                req.ContentType = "multipart/form-data; boundary=" + boundary;
                req.Headers.Set("Authorization", "Bearer " + strToken);
                req.Credentials = System.Net.CredentialCache.DefaultCredentials;
                req.KeepAlive = true;

                //req.Proxy = null;
                if (String.IsNullOrEmpty(Message))
                {
                    throw new Exception("Message為空");
                }
                rs = req.GetRequestStream();
                //寫入API參數

                rs.Write(boundarybytes, 0, boundarybytes.Length);

                string formitem = string.Format(formdataTemplate, "message", Message);
                byte[] formitembytes = System.Text.Encoding.UTF8.GetBytes(formitem);
                rs.Write(formitembytes, 0, formitembytes.Length);
                rs.Write(boundarybytes, 0, boundarybytes.Length);


                //上傳檔案
                if (!File.Exists(FileFullPath))
                    throw new FileNotFoundException("找不到檔案" + FileFullPath[0]);

                getFileExtension = new FileInfo(FileFullPath);

                switch (getFileExtension.Extension.ToLower())
                {
                    case ".jpg":
                        contentType = "image/jpeg";
                        break;
                    case ".jpeg":
                        contentType = "image/jpeg";
                        break;
                    case ".png":
                        contentType = "image/png";
                        break;
                    default:
                        contentType = null;
                        break;
                }
                if (string.IsNullOrEmpty(contentType))
                {
                    throw new Exception("不支援的檔案格式");
                }
                header = string.Format(headerTemplate, FileParameter, FileFullPath, contentType);
                headerbytes = System.Text.Encoding.UTF8.GetBytes(header);
                rs.Write(headerbytes, 0, headerbytes.Length);

                fileStream = new FileStream(FileFullPath, FileMode.Open, FileAccess.Read);

                byte[] buffer = new byte[4096];
                int bytesRead = 0;
                while ((bytesRead = fileStream.Read(buffer, 0, buffer.Length)) != 0)
                {
                    rs.Write(buffer, 0, bytesRead);
                }
                fileStream.Close();

                byte[] trailer = System.Text.Encoding.UTF8.GetBytes("\r\n--" + boundary + "--\r\n");
                rs.Write(trailer, 0, trailer.Length);

                rs.Close();
                using (wr = req.GetResponse())
                {
                    using (myStreamReader = new StreamReader(wr.GetResponseStream(), System.Text.Encoding.Default))
                    {
                        ReturnValue = myStreamReader.ReadToEnd();
                    }
                }

                LogMessage.Append("URL:\r\n" + _ApiURL + "\r\n");
                LogMessage.Append("API 參數:\r\n" + Message + "\r\n");
                LogMessage.Append("API 回傳:\r\n" + ReturnValue + "\r\n");
                Console.WriteLine(LogMessage.ToString());

                if (req != null)
                    req.Abort();
                req = null;

                return ReturnValue;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
                Image img = Image.FromFile(FileFullPath);
                FileInfo fi = new FileInfo(FileFullPath);
                throw new Exception($"附件圖檔大小: 寬{img.Width} 高{img.Height} 檔案大小{fi.Length}\r\n" + ex.ToString(), ex);
            }
        }

        /// <summary>
        /// 重試機制
        /// </summary>
        /// <param name="operation">重試的funtion</param>
        /// <param name="retryNumber">重試次數</param>
        /// <param name="millisecondsTimeout">重試間格毫秒數</param>
        /// <returns></returns>
        private Boolean Retry(Action operation, int retryNumber = 3, int millisecondsTimeout = 1000)
        {
            for (int i = 0; i < retryNumber; i++)
            {
                try
                {
                    operation();
                    return false;
                }
                catch { Task.Delay(TimeSpan.FromMilliseconds(millisecondsTimeout)).Wait(); }
            }
            return true;
        }

        /// <summary>
        /// Funtion() = True => return true
        /// Funtion() = False => ReTry retryNumber次 => return false
        /// </summary>
        /// <param name="action">回傳 Boolean的Funtion</param>
        /// <param name="retryNumber">重試次數</param>
        /// <param name="millisecondsTimeout">重試間格毫秒數</param>
        /// <returns></returns>
        //https://stackoverflow.com/questions/1563191/cleanest-way-to-write-retry-logic
        private static Boolean RetryTrue(Func<Boolean> action, int retryNumber = 3, int millisecondsTimeout = 1000)
        {
            for (int i = 0; i < retryNumber; i++)
            {
                try
                {
                    if (action())
                    {
                        return true;
                    }
                    Task.Delay(TimeSpan.FromMilliseconds(millisecondsTimeout)).Wait();
                }
                catch { Task.Delay(TimeSpan.FromMilliseconds(millisecondsTimeout)).Wait(); }
            }
            return false;
        }
    }
}
