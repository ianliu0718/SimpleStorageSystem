using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace 簡易倉儲系統.EssentialTool
{
    internal class Settings
    {
        static String SettingsPath;
        static SystemConfig Config;
        /// <summary>
        /// 設定檔初始設定
        /// </summary>
        public static void StartUp(string _SettingsPath)
        {
            try
            {
                SettingsPath = _SettingsPath + @"\Setting.xml";
                Config = new SystemConfig(SettingsPath);
                if (!Directory.Exists(_SettingsPath))
                {
                    Directory.CreateDirectory(_SettingsPath);
                }
                if (!File.Exists(SettingsPath))
                {
                    XmlWriterSettings settings = new XmlWriterSettings();
                    settings.Indent = true;
                    settings.NewLineOnAttributes = true;
                    XmlWriter writer = XmlWriter.Create(SettingsPath, settings);
                    writer.WriteStartElement("Setup");
                    writer.WriteElementString("每日檢查", _每日檢查);
                    writer.WriteElementString("主機序號", _主機序號);
                    writer.WriteElementString("序號", _序號);
                    writer.WriteElementString("資料庫路徑", _資料庫路徑);
                    writer.WriteElementString("Excel路徑", _Excel路徑);
                    writer.WriteElementString("印表機名稱", _印表機名稱);
                    writer.WriteElementString("販售地區1", _販售地區1);
                    writer.WriteElementString("類型1", _類型1);
                    writer.WriteElementString("販售地區2", _販售地區2);
                    writer.WriteElementString("類型2", _類型2);
                    writer.WriteElementString("販售地區3", _販售地區3);
                    writer.WriteElementString("類型3", _類型3);
                    writer.WriteElementString("SQL語法", _SQL語法);
                    writer.WriteElementString("廠商標題1", _廠商標題1);
                    writer.WriteElementString("廠商標題2", _廠商標題2);
                    writer.WriteElementString("廠商標題3", _廠商標題3);
                    writer.WriteEndElement();
                    writer.Flush();
                    writer.Close();
                }
            }
            catch// (Exception ex)
            {
                throw;
            }
        }

        private static string _廠商標題3 = "";
        /// <summary>
        /// 廠商標題3
        /// </summary>
        public static String 廠商標題3
        {
            get
            {
                String tmpValue = null;
                try
                {
                    Config.getConfigValue("廠商標題3", out tmpValue);
                }
                catch (Exception ee)
                {
                    if (ee.Message == "查無【廠商標題3】此參數設定")
                    {
                        Config.setConfigValue("廠商標題3", _廠商標題3);
                        tmpValue = _廠商標題3;
                    }
                }

                return tmpValue;
            }
            set
            {
                try
                { Config.setConfigValue("廠商標題3", value); }
                catch { }
            }
        }

        private static string _廠商標題2 = "";
        /// <summary>
        /// 廠商標題2
        /// </summary>
        public static String 廠商標題2
        {
            get
            {
                String tmpValue = null;
                try
                {
                    Config.getConfigValue("廠商標題2", out tmpValue);
                }
                catch (Exception ee)
                {
                    if (ee.Message == "查無【廠商標題2】此參數設定")
                    {
                        Config.setConfigValue("廠商標題2", _廠商標題2);
                        tmpValue = _廠商標題2;
                    }
                }

                return tmpValue;
            }
            set
            {
                try
                { Config.setConfigValue("廠商標題2", value); }
                catch { }
            }
        }

        private static string _廠商標題1 = "";
        /// <summary>
        /// 廠商標題1
        /// </summary>
        public static String 廠商標題1
        {
            get
            {
                String tmpValue = null;
                try
                {
                    Config.getConfigValue("廠商標題1", out tmpValue);
                }
                catch (Exception ee)
                {
                    if (ee.Message == "查無【廠商標題1】此參數設定")
                    {
                        Config.setConfigValue("廠商標題1", _廠商標題1);
                        tmpValue = _廠商標題1;
                    }
                }

                return tmpValue;
            }
            set
            {
                try
                { Config.setConfigValue("廠商標題1", value); }
                catch { }
            }
        }

        private static string _SQL語法 = "";
        /// <summary>
        /// SQL語法
        /// </summary>
        public static String SQL語法
        {
            get
            {
                String tmpValue = null;
                try
                {
                    Config.getConfigValue("SQL語法", out tmpValue);
                }
                catch (Exception ee)
                {
                    if (ee.Message == "查無【SQL語法】此參數設定")
                    {
                        Config.setConfigValue("SQL語法", _SQL語法);
                        tmpValue = SQL語法;
                    }
                }

                return tmpValue;
            }
            set
            {
                try
                { Config.setConfigValue("SQL語法", value); }
                catch { }
            }
        }

        private static string _類型3 = "12粒/15粒/18粒/20粒/24粒/28粒";
        /// <summary>
        /// 類型3
        /// </summary>
        public static String 類型3
        {
            get
            {
                String tmpValue = null;
                try
                {
                    Config.getConfigValue("類型3", out tmpValue);
                }
                catch (Exception ee)
                {
                    if (ee.Message == "查無【類型3】此參數設定")
                    {
                        Config.setConfigValue("類型3", _類型3);
                        tmpValue = 類型3;
                    }
                }

                return tmpValue;
            }
            set
            {
                try
                { Config.setConfigValue("類型3", value); }
                catch { }
            }
        }

        private static string _販售地區3 = "超市/台斤";
        /// <summary>
        /// 販售地區3
        /// </summary>
        public static String 販售地區3
        {
            get
            {
                String tmpValue = null;
                try
                {
                    Config.getConfigValue("販售地區3", out tmpValue);
                }
                catch (Exception ee)
                {
                    if (ee.Message == "查無【販售地區3】此參數設定")
                    {
                        Config.setConfigValue("販售地區3", _販售地區3);
                        tmpValue = _販售地區3;
                    }
                }

                return tmpValue;
            }
            set
            {
                try
                { Config.setConfigValue("販售地區3", value); }
                catch { }
            }
        }

        private static string _類型2 = "9粒/12粒/14粒/16粒/20粒/24粒/小";
        /// <summary>
        /// 類型2
        /// </summary>
        public static String 類型2
        {
            get
            {
                String tmpValue = null;
                try
                {
                    Config.getConfigValue("類型2", out tmpValue);
                }
                catch (Exception ee)
                {
                    if (ee.Message == "查無【類型2】此參數設定")
                    {
                        Config.setConfigValue("類型2", _類型2);
                        tmpValue = 類型2;
                    }
                }

                return tmpValue;
            }
            set
            {
                try
                { Config.setConfigValue("類型2", value); }
                catch { }
            }
        }

        private static string _販售地區2 = "外銷日本/公斤";
        /// <summary>
        /// 販售地區2
        /// </summary>
        public static String 販售地區2
        {
            get
            {
                String tmpValue = null;
                try
                {
                    Config.getConfigValue("販售地區2", out tmpValue);
                }
                catch (Exception ee)
                {
                    if (ee.Message == "查無【販售地區2】此參數設定")
                    {
                        Config.setConfigValue("販售地區2", _販售地區2);
                        tmpValue = _販售地區2;
                    }
                }

                return tmpValue;
            }
            set
            {
                try
                { Config.setConfigValue("販售地區2", value); }
                catch { }
            }
        }

        private static string _類型1 = "9粒/12粒/14粒/16粒/20粒/24粒/小";
        /// <summary>
        /// 類型1
        /// </summary>
        public static String 類型1
        {
            get
            {
                String tmpValue = null;
                try
                {
                    Config.getConfigValue("類型1", out tmpValue);
                }
                catch (Exception ee)
                {
                    if (ee.Message == "查無【類型1】此參數設定")
                    {
                        Config.setConfigValue("類型1", _類型1);
                        tmpValue = 類型1;
                    }
                }

                return tmpValue;
            }
            set
            {
                try
                { Config.setConfigValue("類型1", value); }
                catch { }
            }
        }

        private static string _販售地區1 = "外銷韓國/公斤";
        /// <summary>
        /// 販售地區1
        /// </summary>
        public static String 販售地區1
        {
            get
            {
                String tmpValue = null;
                try
                {
                    Config.getConfigValue("販售地區1", out tmpValue);
                }
                catch (Exception ee)
                {
                    if (ee.Message == "查無【販售地區1】此參數設定")
                    {
                        Config.setConfigValue("販售地區1", _販售地區1);
                        tmpValue = _販售地區1;
                    }
                }

                return tmpValue;
            }
            set
            {
                try
                { Config.setConfigValue("販售地區1", value); }
                catch { }
            }
        }

        private static string _主機序號 = "";
        /// <summary>
        /// 主機序號
        /// </summary>
        public static String 主機序號
        {
            get
            {
                String tmpValue = null;
                try
                {
                    Config.getConfigValue("主機序號", out tmpValue);
                }
                catch (Exception ee)
                {
                    if (ee.Message == "查無【主機序號】此參數設定")
                    {
                        Config.setConfigValue("主機序號", _主機序號);
                        tmpValue = _主機序號;
                    }
                }

                return tmpValue;
            }
            set
            {
                try
                { Config.setConfigValue("主機序號", value); }
                catch { }
            }
        }

        private static string _印表機名稱 = "Canon GM2000 series";
        /// <summary>
        /// 印表機名稱
        /// </summary>
        public static String 印表機名稱
        {
            get
            {
                String tmpValue = null;
                try
                {
                    Config.getConfigValue("印表機名稱", out tmpValue);
                }
                catch { }

                return tmpValue;
            }
            set
            {
                try
                { Config.setConfigValue("印表機名稱", value); }
                catch { }
            }
        }

        private static string _Excel路徑 = @".\";
        /// <summary>
        /// Excel路徑
        /// </summary>
        public static String Excel路徑
        {
            get
            {
                String tmpValue = null;
                try
                {
                    Config.getConfigValue("Excel路徑", out tmpValue);
                }
                catch
                {
                    try
                    { 
                        Config.setConfigValue("Excel路徑", _Excel路徑);
                        tmpValue = _Excel路徑;
                    }
                    catch { }
                }

                return tmpValue;
            }
            set
            {
                try
                { Config.setConfigValue("Excel路徑", value); }
                catch { }
            }
        }

        private static string _每日檢查 = "";
        /// <summary>
        /// 每日檢查
        /// </summary>
        public static String 每日檢查
        {
            get
            {
                String tmpValue = null;
                try
                {
                    Config.getConfigValue("每日檢查", out tmpValue);
                }
                catch { }

                return tmpValue;
            }
            set
            {
                try
                { Config.setConfigValue("每日檢查", value); }
                catch { }
            }
        }

        private static string _序號 = "";
        /// <summary>
        /// 序號
        /// </summary>
        public static String 序號
        {
            get
            {
                String tmpValue = null;
                try
                {
                    Config.getConfigValue("序號", out tmpValue);
                }
                catch { }

                return tmpValue;
            }
            set
            {
                try
                { Config.setConfigValue("序號", value); }
                catch { }
            }
        }

        private static string _資料庫路徑 = @".\";
        /// <summary>
        /// 資料庫路徑
        /// </summary>
        public static String 資料庫路徑
        {
            get
            {
                String tmpValue = null;
                try
                {
                    Config.getConfigValue("資料庫路徑", out tmpValue);
                }
                catch { }

                return tmpValue;
            }
            set
            {
                try
                { Config.setConfigValue("資料庫路徑", value); }
                catch { }
            }
        }
    }
}
