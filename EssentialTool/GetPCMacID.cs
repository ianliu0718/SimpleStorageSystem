using System;
using System.Collections.Generic;
using System.Linq;
using System.Management;
using System.Text;
using System.Threading.Tasks;

namespace 簡易倉儲系統.EssentialTool
{
    internal class GetPCMacID
    {
        /// <summary>
        /// 0.獲取主板序列號代碼
        /// </summary>
        /// <returns></returns>
        public static string GetBaseboardID()
        {
            try
            {
                string BaseboardInfo = "";//主板序列號
                ManagementObjectSearcher mos = new ManagementObjectSearcher("select * from Win32_baseboard");
                foreach (ManagementObject mo in mos.Get())
                {
                    BaseboardInfo = mo["SerialNumber"].ToString();
                    break;
                    //Response.Write("主板製造商:" + mo["Manufacturer"]);
                    //Response.Write("型號:" + mo["Product"]);
                    //Response.Write("序列號:" + mo["SerialNumber"].ToString());
                }
                mos = null;
                return BaseboardInfo;
            }
            catch
            {
                return "unknow";
            }
        }
        /// <summary>
        /// 1.獲取CPU序列號代碼
        /// </summary>
        /// <returns></returns>
        public static string GetCpuID()
        {
            try
            {
                string cpuInfo = "";//cpu序列號
                ManagementClass mc = new ManagementClass("Win32_Processor");
                ManagementObjectCollection moc = mc.GetInstances();
                foreach (ManagementObject mo in moc)
                {
                    cpuInfo = mo.Properties["ProcessorId"].Value.ToString();
                }
                moc = null;
                mc = null;
                return cpuInfo;
            }
            catch
            {
                return "unknow";
            }
        }

        /// <summary>
        /// 2.獲取網卡硬件地址
        /// </summary>
        /// <returns></returns>
        public static string GetMacAddress()
        {
            try
            {
                string mac = "";
                ManagementClass mc = new ManagementClass("Win32_NetworkAdapterConfiguration");
                ManagementObjectCollection moc = mc.GetInstances();
                foreach (ManagementObject mo in moc)
                {
                    if ((bool)mo["IPEnabled"] == true)
                    {
                        mac = mo["MacAddress"].ToString();
                        break;
                    }
                }
                moc = null;
                mc = null;
                return mac;
            }
            catch
            {
                return "unknow";
            }
        }

        /// <summary>
        /// 3.獲取硬盤ID
        /// </summary>
        /// <returns></returns>
        string GetDiskID()
        {
            try
            {
                String HDid = "";
                ManagementClass mc = new ManagementClass("Win32_DiskDrive");
                ManagementObjectCollection moc = mc.GetInstances();
                foreach (ManagementObject mo in moc)
                {
                    HDid = (string)mo.Properties["Model"].Value;
                }
                moc = null;
                mc = null;
                return HDid;
            }
            catch
            {
                return "unknow";
            }
        }

        /// <summary>
        /// 4.獲取IP地址
        /// </summary>
        /// <returns></returns>
        public static string GetIPAddress()
        {
            try
            {
                string st = "";
                ManagementClass mc = new ManagementClass("Win32_NetworkAdapterConfiguration");
                ManagementObjectCollection moc = mc.GetInstances();
                foreach (ManagementObject mo in moc)
                {
                    if ((bool)mo["IPEnabled"] == true)
                    {
                        //st=mo["IpAddress"].ToString();
                        System.Array ar;
                        ar = (System.Array)(mo.Properties["IpAddress"].Value);
                        st = ar.GetValue(0).ToString();
                        break;
                    }
                }
                moc = null;
                mc = null;
                return st;
            }
            catch
            {
                return "unknow";
            }
        }

        /// <summary>
        /// 5.操作系統的登錄用戶名
        /// </summary>
        /// <returns></returns>
        public static string GetUserName()
        {
            try
            {
                return Environment.UserName;
            }
            catch
            {
                return "unknow";
            }
        }

        /// <summary>
        /// 6.獲取計算機名
        /// </summary>
        /// <returns></returns>
        public static string GetComputerName()
        {
            try
            {
                return System.Environment.MachineName;
            }
            catch
            {
                return "unknow";
            }
        }

        /// <summary>
        /// 7. PC類型
        /// </summary>
        /// <returns></returns>
        public static string GetSystemType()
        {
            try
            {
                string st = "";
                ManagementClass mc = new ManagementClass("Win32_ComputerSystem");
                ManagementObjectCollection moc = mc.GetInstances();
                foreach (ManagementObject mo in moc)
                {
                    st = mo["SystemType"].ToString();
                }
                moc = null;
                mc = null;
                return st;
            }
            catch
            {
                return "unknow";
            }
        }

        /// <summary>
        /// 8.物理內存
        /// </summary>
        /// <returns></returns>
        string GetTotalPhysicalMemory()
        {
            try
            {
                string st = "";
                ManagementClass mc = new ManagementClass("Win32_ComputerSystem");
                ManagementObjectCollection moc = mc.GetInstances();
                foreach (ManagementObject mo in moc)
                {
                    st = mo["TotalPhysicalMemory"].ToString();
                }
                moc = null;
                mc = null;
                return st;
            }
            catch
            {
                return "unknow";
            }
        }

        private void GetInfo()
        {
            string cpuInfo = "";//cpu序列號
            ManagementClass cimobject = new ManagementClass("Win32_Processor");
            ManagementObjectCollection moc = cimobject.GetInstances();
            foreach (ManagementObject mo in moc)
            {
                cpuInfo = mo.Properties["ProcessorId"].Value.ToString();
                //Response.Write("cpu序列號：" + cpuInfo.ToString());
            }
            //獲取硬盤ID
            String HDid;
            ManagementClass cimobject1 = new ManagementClass("Win32_DiskDrive");
            ManagementObjectCollection moc1 = cimobject1.GetInstances();
            foreach (ManagementObject mo in moc1)
            {
                HDid = (string)mo.Properties["Model"].Value;
                //Response.Write("硬盤序列號：" + HDid.ToString());
            }

            //獲取網卡硬件地址 

            ManagementClass mc = new ManagementClass("Win32_NetworkAdapterConfiguration");
            ManagementObjectCollection moc2 = mc.GetInstances();
            foreach (ManagementObject mo in moc2)
            {
                //if ((bool)mo["IPEnabled"] == true)
                //    Response.Write("MAC address\t{0}" + mo["MacAddress"].ToString());
                mo.Dispose();
            }
            //主板
            //string strbNumber = string.Empty;
            ManagementObjectSearcher mos = new ManagementObjectSearcher("select * from Win32_baseboard");
            foreach (ManagementObject mo in mos.Get())
            {
                ////strbNumber = mo["SerialNumber"].ToString();
                ////break;
                //Response.Write("主板製造商:" + mo["Manufacturer"]);
                //Response.Write("型號:" + mo["Product"]);
                //Response.Write("序列號:" + mo["SerialNumber"].ToString());
            }
            // 物理內存 
            ManagementClass mc4 = new ManagementClass("Win32_ComputerSystem");
            ManagementObjectCollection moc4 = mc4.GetInstances();
            foreach (ManagementObject mo4 in moc4)
            {
                //Response.Write("內存:" + Convert.ToString(Convert.ToInt64(mo4["TotalPhysicalMemory"]) / 1024) + "K");
            }
        }
    }
}
