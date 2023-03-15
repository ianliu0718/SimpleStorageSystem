using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace 簡易倉儲系統.EssentialTool
{
    internal class SystemConfig
    {
        String FilePath;
        private String errMsg;
        public String getErrorMessage { get { return errMsg; } }

        public SystemConfig(String XMLFilePath)
        {
            FilePath = XMLFilePath;
        }

        private static bool IsFileLocked(string file)
        {
            try
            {
                using (File.Open(file, FileMode.Open, FileAccess.Write, FileShare.None))
                {
                    return false;
                }
            }
            catch (IOException exception)
            {
                var errorCode = Marshal.GetHRForException(exception) & 65535;
                return errorCode == 32 || errorCode == 33;
            }
            catch (Exception)
            {
                return false;
            }
        }

        public void setConfigValue(String ConfigName, Object ConfigValue)
        {
            errMsg = null;
            try
            {
                if (!File.Exists(FilePath))
                {
                    errMsg = "無此檔案或目錄：" + FilePath;
                    throw new FileNotFoundException("無此檔案或目錄：" + FilePath);
                }

                while (IsFileLocked(FilePath))
                    Thread.Sleep(100);

                XElement root = XElement.Load(FilePath);
                IEnumerable<XElement> SearchValue = from el in root.Elements()
                                                    where el.Name.LocalName == ConfigName
                                                    select el;
                if (SearchValue.Count() != 0)
                    root.SetElementValue(ConfigName, ConfigValue);
                else
                    root.Add(new XElement(ConfigName, ConfigValue));
                root.Save(FilePath);
            }
            catch (Exception)
            { throw; }
        }

        public void getConfigValue(String ConfigName, out String ConfigValue)
        {
            ConfigValue = null;
            try
            {
                XElement root = XElement.Load(FilePath);
                IEnumerable<XElement> SearchValue = from el in root.Elements()
                                                    where el.Name.LocalName == ConfigName
                                                    select el;
                if (SearchValue.Count() != 0)
                    ConfigValue = SearchValue.First().Value;
                else
                {
                    throw new KeyNotFoundException($"查無{ConfigName}此參數設定");
                }
            }
            catch (Exception)
            { throw; }
        }

        public void RemoveConfigValue(String ConfigName)
        {
            errMsg = null;
            try
            {
                if (!File.Exists(FilePath))
                {
                    errMsg = "無此檔案或目錄：" + FilePath;
                    throw new FileNotFoundException("無此檔案或目錄：" + FilePath);
                }

                while (IsFileLocked(FilePath))
                    Thread.Sleep(100);

                XElement root = XElement.Load(FilePath);
                IEnumerable<XElement> SearchValue = from el in root.Elements()
                                                    where el.Name.LocalName == ConfigName
                                                    select el;

                if (SearchValue.Count() != 0)
                {
                    SearchValue.Remove();
                    root.Save(FilePath);
                }
            }
            catch (Exception)
            { throw; }
        }
    }
}
