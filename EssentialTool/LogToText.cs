using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using 簡易倉儲系統.Properties;

namespace 簡易倉儲系統.EssentialTool
{
    internal class LogToText
    {
        public enum enumLogType { Trace, Debug, Info, Warn, Error, Fatal };

        //SendMail AutoReport;

        String LogPath;
        private String errMsg;

        public LogToText(String LogPath)
        {
            this.LogPath = LogPath;
            //AutoReport = new SendMail();
        }

        public String ErrroMessage
        {
            get { return errMsg; }
        }

        /// <summary>
        /// 記錄log
        /// </summary>
        /// <param name="LogMessage">要log的東西</param>
        /// <param name="LogType">log類型(切開檔案用)</param>
        /// <param name="SendMail">是否發送通知mail</param>
        /// <returns></returns>
        public Boolean LogMessage(String LogMessage, enumLogType LogType)
        {
            if (!Directory.Exists(LogPath))
            {
                Directory.CreateDirectory(LogPath);
            }
            DirectoryInfo dirinfo = new DirectoryInfo(LogPath);
            FileInfo[] FileList = dirinfo.GetFiles();
            String FileName;
            FileInfo tmpfi;
            switch (LogType)
            {
                case enumLogType.Trace:
                    FileName = LogPath + @"\trace\" + DateTime.Now.ToString("yyyy-MM-dd") + "Trace" + "_Message.txt";
                    break;
                case enumLogType.Debug:
                    FileName = LogPath + @"\debug\" + DateTime.Now.ToString("yyyy-MM-dd") + "Debug" + "_Message.txt";
                    break;
                case enumLogType.Info:
                    FileName = LogPath + @"\info\" + DateTime.Now.ToString("yyyy-MM-dd") + "Info" + "_Message.txt";
                    break;
                case enumLogType.Warn:
                    FileName = LogPath + @"\warn\" + DateTime.Now.ToString("yyyy-MM-dd") + "Warn" + "_Message.txt";
                    break;
                case enumLogType.Error:
                    FileName = LogPath + @"\error\" + DateTime.Now.ToString("yyyy-MM-dd") + "Error" + "_Message.txt";
                    break;
                case enumLogType.Fatal:
                    FileName = LogPath + @"\fatal\" + DateTime.Now.ToString("yyyy-MM-dd") + "Fatal" + "_Message.txt";
                    break;
                default:
                    return false;
            }

            try
            {
                tmpfi = new FileInfo(FileName);
                if (!Directory.Exists(tmpfi.DirectoryName))
                {
                    Directory.CreateDirectory(tmpfi.DirectoryName);
                }
                //if (SendMail)
                //    if (AutoReport.Send(Settings.AlarmEmails, System.Net.Dns.GetHostName() + "系統錯誤", LogMessage))
                //    {
                //        String MailFileName = LogPath + @"\SendMail\" + DateTime.Now.ToString("yyyy-MM-dd") + "_SendMailError.txt";
                //        using (StreamWriter sw = File.AppendText(MailFileName))
                //        {
                //            sw.WriteLine("訊息時間:" + DateTime.Now.ToString());
                //            sw.WriteLine(AutoReport.ErrorMessage);
                //            sw.WriteLine("====================");
                //        }
                //    }

                using (StreamWriter sw = File.AppendText(FileName))
                {
                    sw.WriteLine("訊息時間:" + DateTime.Now.ToString());
                    sw.WriteLine(LogMessage);
                    sw.WriteLine("====================");
                }

                foreach (FileInfo item in FileList)
                {
                    if (item.Extension == ".txt")
                    {
                        if (item.Name.IndexOf("_logMessage") != -1 || item.Name.IndexOf("_SendMailError") != -1)
                        {
                            //保存錯誤訊息6個月
                            if (Convert.ToDateTime(item.Name.Substring(0, 10)) < DateTime.Now.AddMonths(-6))
                            {
                                item.Delete();
                            }
                        }
                    }
                }
                return false;
            }
            catch (Exception ex)
            {
                errMsg = ex.ToString();
                return true;
            }
        }
    }
}
