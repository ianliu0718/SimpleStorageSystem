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
                    writer.WriteElementString("序號", _序號);
                    writer.WriteElementString("資料庫路徑", _資料庫路徑);
                    writer.WriteElementString("Excel路徑", _Excel路徑);
                    writer.WriteElementString("印表機名稱", _印表機名稱);
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
