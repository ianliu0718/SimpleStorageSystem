using OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Information;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
using OfficeOpenXml.Style;
using Spire.Pdf.Exporting.XPS.Schema;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.Entity.Core.Common.CommandTrees.ExpressionBuilder;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Printing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Reflection.Emit;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;
using 簡易倉儲系統.DB;
using 簡易倉儲系統.EssentialTool;
using 簡易倉儲系統.EssentialTool.Excel;
using static 簡易倉儲系統.EssentialTool.LogToText;

namespace 簡易倉儲系統
{
    public partial class UserView : Form
    {
        LogToText log = new LogToText(@".\Log"); 
        DB_SQLite dB_SQLite = new DB_SQLite();

        /// <summary>
        /// 設定檔路徑
        /// </summary>
        public static string Setting_Path = @".\";
        /// <summary>
        /// 程式版號
        /// </summary>
        public static string VersionNumber = "";
        /// <summary>
        /// 資料庫路徑
        /// </summary>
        public static string DB_Path = "";
        /// <summary>
        /// 資料表名稱
        /// </summary>
        public static string TableName = "";
        /// <summary>
        /// 類型
        /// </summary>
        public static string type = "";
        /// <summary>
        /// 販售地區
        /// </summary>
        public static string salesArea = "";
        /// <summary>
        /// 重量單位
        /// </summary>
        public static string unit = "";
        /// <summary>
        /// 單價
        /// </summary>
        public static string unitPrice = "";

        public UserView()
        {
            InitializeComponent();
        }

        private void UserView_Load(object sender, EventArgs e)
        {
            Settings.StartUp(Setting_Path);
            VersionNumber = Application.ProductVersion;
            this.Text += $"v.{VersionNumber} Bulid{File.GetLastWriteTime(Application.ExecutablePath).ToString("yyyyMMdd")}";
            DB_Path = Settings.資料庫路徑 + @"data.db";
            label5.Text = "";
            comboBox2.SelectedIndex = 0;
            comboBox2.Location = new Point(radioButton9.Location.X + radioButton9.Size.Width + 10, comboBox2.Location.Y);

            //定時偵測序號
            timer_detection_Tick(sender, e);

            #region 取得類型設定參數
            try
            {
                log.LogMessage("取得類型設定參數 開始", enumLogType.Trace);
                radioButton8.Text = Settings.販售地區1.Split('/')[0] + "(F1)";
                label1.Text = Settings.販售地區1.Split('/')[1];
                string _Type = Settings.類型1;
                int _TypeCount = _Type.Split('/').Count();
                int _Space = 10;
                #region 8個類型的位置調整
                if (_TypeCount >= 1)
                {
                    radioButton1.Text = _Type.Split('/')[0] + "(F4)";
                    radioButton1.Enabled = true;
                    radioButton1.Visible = true;
                }
                else
                {
                    radioButton1.Enabled = false;
                    radioButton1.Visible = false;
                }
                if (_TypeCount >= 2)
                {
                    radioButton2.Text = _Type.Split('/')[1] + "(F5)";
                    radioButton2.Location = new Point(radioButton1.Location.X + radioButton1.Size.Width + _Space, radioButton2.Location.Y);
                    radioButton2.Enabled = true;
                    radioButton2.Visible = true;
                }
                else
                {
                    radioButton2.Enabled = false;
                    radioButton2.Visible = false;
                }
                if (_TypeCount >= 3)
                {
                    radioButton3.Text = _Type.Split('/')[2] + "(F6)";
                    radioButton3.Location = new Point(radioButton2.Location.X + radioButton2.Size.Width + _Space, radioButton2.Location.Y);
                    radioButton3.Enabled = true;
                    radioButton3.Visible = true;
                }
                else
                {
                    radioButton3.Enabled = false;
                    radioButton3.Visible = false;
                }
                if (_TypeCount >= 4)
                {
                    radioButton4.Text = _Type.Split('/')[3] + "(F7)";
                    radioButton4.Location = new Point(radioButton3.Location.X + radioButton3.Size.Width + _Space, radioButton3.Location.Y);
                    radioButton4.Enabled = true;
                    radioButton4.Visible = true;
                }
                else
                {
                    radioButton4.Enabled = false;
                    radioButton4.Visible = false;
                }
                if (_TypeCount >= 5)
                {
                    radioButton5.Text = _Type.Split('/')[4] + "(F8)";
                    radioButton5.Location = new Point(radioButton4.Location.X + radioButton4.Size.Width + _Space, radioButton4.Location.Y);
                    radioButton5.Enabled = true;
                    radioButton5.Visible = true;
                }
                else
                {
                    radioButton5.Enabled = false;
                    radioButton5.Visible = false;
                }
                if (_TypeCount >= 6)
                {
                    radioButton6.Text = _Type.Split('/')[5] + "(F9)";
                    radioButton6.Location = new Point(radioButton5.Location.X + radioButton5.Size.Width + _Space, radioButton5.Location.Y);
                    radioButton6.Enabled = true;
                    radioButton6.Visible = true;
                }
                else
                {
                    radioButton6.Enabled = false;
                    radioButton6.Visible = false;
                }
                if (_TypeCount >= 7)
                {
                    radioButton7.Text = _Type.Split('/')[6] + "(F10)";
                    radioButton7.Location = new Point(radioButton6.Location.X + radioButton6.Size.Width + _Space, radioButton6.Location.Y);
                    radioButton7.Enabled = true;
                    radioButton7.Visible = true;
                }
                else
                {
                    radioButton7.Enabled = false;
                    radioButton7.Visible = false;
                }
                if (_TypeCount >= 8)
                {
                    radioButton11.Text = _Type.Split('/')[7] + "(F11)";
                    radioButton11.Location = new Point(radioButton7.Location.X + radioButton7.Size.Width + _Space, radioButton7.Location.Y);
                    radioButton11.Enabled = true;
                    radioButton11.Visible = true;
                }
                else
                {
                    radioButton11.Enabled = false;
                    radioButton11.Visible = false;
                }
                #endregion
                radioButton9.Text = Settings.販售地區2.Split('/')[0] + "(F2)";
                radioButton10.Text = Settings.販售地區3.Split('/')[0] + "(F3)";
                log.LogMessage("取得類型設定參數 成功\r\n" + radioButton8.Text.Split('(')[0] + $@"：{_Type}", enumLogType.Info);
                log.LogMessage("取得類型設定參數 成功\r\n" + radioButton8.Text.Split('(')[0] + $@"：{_Type}", enumLogType.Trace);
            }
            catch (Exception ee)
            {
                log.LogMessage("取得類型設定參數 失敗\r\n" + ee.Message, enumLogType.Error);
                MessageBox.Show("取得類型設定參數 失敗\r\n" + ee.Message);
                return;
            }
            #endregion

            try
            {
                comboBox1.Items.Clear();
                // 讀取資料
                foreach (DataRow item in dB_SQLite.GetDataTable(DB_Path, $@"SELECT CustomerID, CustomerName FROM CustomerProfile;").Rows)
                {
                    comboBox1.Items.Add(item[0].ToString() + "_" + item[1].ToString());
                }
                log.LogMessage("讀取資料庫 成功。", enumLogType.Trace);
            }
            catch (Exception ee)
            {
                log.LogMessage("連線資料庫 失敗\r\n" + ee.Message, enumLogType.Error);
                MessageBox.Show("連線資料庫 失敗\r\n" + ee.Message);
                return;
            }
            textBox1.Focus();
            log.LogMessage("使用者介面啓動", enumLogType.Info);
        }

        private void timer_detection_Tick(object sender, EventArgs e)
        {
            #region 檢查時間為最新
            try
            {
                log.LogMessage("檢查時間 開始", enumLogType.Trace);

                if (!String.IsNullOrEmpty(Settings.每日檢查))
                {
                    string _TimeText = EncryptionDecryption.desDecryptBase64(Settings.每日檢查);
                    DateTime dateTime = DateTime.Parse(_TimeText);
                    if (dateTime > DateTime.Now)
                    {
                        log.LogMessage("檢查時間_無效 失敗", enumLogType.Error);
                        MessageBox.Show("檢查時間_無效 失敗");
                        Application.Exit();
                        return;
                    }
                }
                Settings.每日檢查 = EncryptionDecryption.desEncryptBase64(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
                log.LogMessage("檢查時間 成功", enumLogType.Info);
                log.LogMessage("檢查時間 成功", enumLogType.Trace);
            }
            catch (Exception ee)
            {
                log.LogMessage("檢查時間 失敗" + ee.Message, enumLogType.Error);
                MessageBox.Show("檢查時間 失敗");
                Application.Exit();
                return;
            }
            #endregion

            #region 檢查程式是否符合效期內
            try
            {
                log.LogMessage("檢查序號 開始", enumLogType.Trace);
                if (String.IsNullOrEmpty(Settings.序號))
                {
                    log.LogMessage("請於設定檔內輸入序號", enumLogType.Info);
                    MessageBox.Show("請於設定檔內輸入序號");
                    Application.Exit();
                    return;
                }
                string _SerialNumber = EncryptionDecryption.desDecryptBase64(Settings.序號);
                if (DateTime.Now < DateTime.Parse(_SerialNumber.Split('/')[1])
                    || DateTime.Now > DateTime.Parse(_SerialNumber.Split('/')[2]))
                {
                    //表示此程式非有效期
                    log.LogMessage("此序號已失效，請聯絡相關廠商", enumLogType.Error);
                    SendLine.SendLineMessage("PkOjQVn809ZiLtwkmnZqGPy8WmZYnnCsxDfdLLCptlc",
                        "此序號已失效\r\n主機板ID：" + GetPCMacID.GetBaseboardID() +
                        "\r\nCPUID：" + GetPCMacID.GetCpuID() +
                        "\r\n網卡硬件地址：" + GetPCMacID.GetMacAddress() +
                        "\r\nIP地址：" + GetPCMacID.GetIPAddress() +
                        "\r\n操作系統的登錄用戶名：" + GetPCMacID.GetUserName() +
                        "\r\n計算機名：" + GetPCMacID.GetComputerName() +
                        "\r\nPC類型：" + GetPCMacID.GetSystemType()
                        );
                    MessageBox.Show("此序號已失效，請聯絡相關廠商");
                    Application.Exit();
                    return;
                }
                log.LogMessage("檢查序號 成功", enumLogType.Info);
                log.LogMessage("檢查序號 成功", enumLogType.Trace);
            }
            catch (Exception ee)
            {
                log.LogMessage("序號啟動 失敗" + ee.Message, enumLogType.Error);
                MessageBox.Show("序號啟動 失敗");
                Application.Exit();
                return;
            }
            #endregion

            #region 檢查程式是否有重複開啟
            Process[] proc = Process.GetProcessesByName(Process.GetCurrentProcess().ProcessName);
            if (proc.Length > 1)
            {
                //表示此程式已被開啟
                Application.Exit();
                return;
            }
            log.LogMessage("系統啓動", enumLogType.Trace);
            #endregion

            #region 比對 CPU ID 是否吻合
            try
            {
                log.LogMessage("比對 CPU ID 是否吻合 開始", enumLogType.Trace);
                if (String.IsNullOrEmpty(Settings.主機序號))
                {
                    log.LogMessage("未綁定主機，請聯絡相關資訊人員。", enumLogType.Info);
                    MessageBox.Show("未綁定主機，請聯絡相關資訊人員。");
                    Application.Exit();
                    return;
                }
                else if (Settings.主機序號 == GetPCMacID.GetCpuID())
                {
                    Settings.主機序號 = EncryptionDecryption.desEncryptBase64(GetPCMacID.GetCpuID() + GetPCMacID.GetBaseboardID());
                    SendLine.SendLineMessage("PkOjQVn809ZiLtwkmnZqGPy8WmZYnnCsxDfdLLCptlc",
                        "主機綁定成功\r\n主機板ID：" + GetPCMacID.GetBaseboardID() +
                        "\r\nCPUID：" + GetPCMacID.GetCpuID() +
                        "\r\n網卡硬件地址：" + GetPCMacID.GetMacAddress() +
                        "\r\nIP地址：" + GetPCMacID.GetIPAddress() +
                        "\r\n操作系統的登錄用戶名：" + GetPCMacID.GetUserName() +
                        "\r\n計算機名：" + GetPCMacID.GetComputerName() +
                        "\r\nPC類型：" + GetPCMacID.GetSystemType()
                        );
                    MessageBox.Show("綁定成功");
                    log.LogMessage("比對 CPU ID 綁定 成功", enumLogType.Info);
                    log.LogMessage("比對 CPU ID 綁定 成功", enumLogType.Trace);
                }
                else if (EncryptionDecryption.desDecryptBase64(Settings.主機序號) == GetPCMacID.GetCpuID())
                {
                    Settings.主機序號 = EncryptionDecryption.desEncryptBase64(GetPCMacID.GetCpuID() + GetPCMacID.GetBaseboardID());
                    SendLine.SendLineMessage("PkOjQVn809ZiLtwkmnZqGPy8WmZYnnCsxDfdLLCptlc",
                        "原先為綁定CPUID改為主機板ID成功\r\n主機板ID：" + GetPCMacID.GetBaseboardID() +
                        "\r\nCPUID：" + GetPCMacID.GetCpuID() +
                        "\r\n網卡硬件地址：" + GetPCMacID.GetMacAddress() +
                        "\r\nIP地址：" + GetPCMacID.GetIPAddress() +
                        "\r\n操作系統的登錄用戶名：" + GetPCMacID.GetUserName() +
                        "\r\n計算機名：" + GetPCMacID.GetComputerName() +
                        "\r\nPC類型：" + GetPCMacID.GetSystemType()
                        );
                }
                else if (EncryptionDecryption.desDecryptBase64(Settings.主機序號) != (GetPCMacID.GetCpuID() + GetPCMacID.GetBaseboardID()))
                {
                    log.LogMessage("程式已綁定，無法在此電腦執行！", enumLogType.Info);
                    MessageBox.Show("程式已綁定，無法在此電腦執行！");
                    Application.Exit();
                    return;
                }
                log.LogMessage("比對 CPU ID 是否吻合 成功", enumLogType.Info);
                log.LogMessage("比對 CPU ID 是否吻合 成功", enumLogType.Trace);
            }
            catch (Exception ee)
            {
                log.LogMessage("比對 CPU ID 失敗" + ee.Message, enumLogType.Error);
                MessageBox.Show("比對 CPU ID 失敗");
                Application.Exit();
                return;
            }
            #endregion
        }

        //類型設定，單價設定
        private void radioButton_type_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                if (((RadioButton)sender).Checked)
                {
                    log.LogMessage("類型設定，單價設定 開始", enumLogType.Trace);

                    ////radioButton6 選擇後會被不知名原因卡住，所以多按一下"ESC"解除
                    //SendKeys.SendWait("{ESC}");
                    type = ((ButtonBase)sender).Text.Split('(')[0];
                    ((GroupBox)((RadioButton)sender).Parent).BackColor = SystemColors.Control;
                    ((RadioButton)sender).BackColor = Color.GreenYellow;

                    if (!string.IsNullOrEmpty(TableName))
                    {
                        DataTable DT = new DataTable();
                        try
                        {
                            if (File.Exists(DB_Path))
                                DT = dB_SQLite.GetDataTable(DB_Path, $@"SELECT * FROM {TableName} WHERE Date = '{DateTime.Now.ToString("yyyy-MM-dd")}'");
                        }
                        catch (Exception ee)
                        {
                            log.LogMessage("類型設定，單價設定：無法連線資料庫\r\n" + ee.Message, enumLogType.Info);
                            MessageBox.Show("無法連線資料庫\r\n" + ee.Message);
                            return;
                        }
                        if (DT.Rows.Count <= 0)
                        {
                            panel1.BackColor = Color.IndianRed;
                            log.LogMessage("管理者尚未輸入當日價格", enumLogType.Info);
                            MessageBox.Show("管理者尚未輸入當日價格");
                            return;
                        }
                        unitPrice = DT.Rows[0][Int32.Parse(((RadioButton)sender).Tag.ToString())].ToString();
                        label3.Text = unitPrice;
                        panel1.BackColor = SystemColors.Control;
                    }

                    log.LogMessage("類型設定，單價設定 成功：" + unitPrice, enumLogType.Trace);
                }
                else
                {
                    ((RadioButton)sender).BackColor = SystemColors.Control;
                }
                //使用者視窗_Tab其他移除，保留輸入框與列印
                ((RadioButton)sender).TabStop = false;
            }
            catch (Exception ee)
            {
                log.LogMessage("類型設定，單價設定 失敗：\r\n" + ee.Message, enumLogType.Error);
            }
        }

        //販售地點設定，資料表名稱設定，重量單位設定，針對不同地區客製化功能
        int salesArea_Checked_OK = -3;  //-4重複點選 //-3第一次設定 //-2設定中 //-1未設定 //0設定為No //1設定為Yes
        private void radioButton_salesArea_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                log.LogMessage("販售地點設定，資料表名稱設定，重量單位設定，針對不同地區客製化功能 開始", enumLogType.Trace);
                if (salesArea_Checked_OK == -1)
                {
                    if (DialogResult.No == MessageBox.Show("是否需要更改販售地區", "更改販售地區", MessageBoxButtons.YesNo))
                    {
                        salesArea_Checked_OK = -2;
                        ((RadioButton)sender).Checked = !((RadioButton)sender).Checked;
                        salesArea_Checked_OK = 0;
                        return;
                    }
                    salesArea_Checked_OK = 1;
                }
                else if (salesArea_Checked_OK == -2)
                {
                    return;
                }
                else if (salesArea_Checked_OK == -3)
                {
                    salesArea_Checked_OK = -1;
                }
                else if (salesArea_Checked_OK == 0)
                {
                    salesArea_Checked_OK = -2;
                    ((RadioButton)sender).Checked = false;
                    salesArea_Checked_OK = -1;
                }
                else if (salesArea_Checked_OK == 1)
                {
                    salesArea_Checked_OK = -1;
                }

                string _Text = "";
                if (((RadioButton)sender).Checked)
                {
                    dataGridView1.Rows.Clear();

                    string _Type = "";
                    //針對不同地區客製化功能
                    if (((ButtonBase)sender).Text.Contains("F1"))
                    {
                        TableName = "ExportKoreaUnitPrice";
                        var item = dB_SQLite.GetDataTable(DB_Path, @"SELECT * FROM Setting WHERE SettingName = 'ShowMoney_ExportKorea'").Rows;
                        if (item.Count > 0)
                        {
                            panel1.Enabled = Boolean.Parse(item[0][1].ToString());
                            panel1.Visible = Boolean.Parse(item[0][1].ToString());
                            dataGridView1.Columns[5].Visible = Boolean.Parse(item[0][1].ToString());
                        }
                        _Text = Settings.販售地區1.Split('/')[0];
                        unit = Settings.販售地區1.Split('/')[1];
                        _Type = Settings.類型1;
                    }
                    else if (((ButtonBase)sender).Text.Contains("F2"))
                    {
                        TableName = "ExportJapanUnitPrice";
                        var item = dB_SQLite.GetDataTable(DB_Path, @"SELECT * FROM Setting WHERE SettingName = 'ShowMoney_ExportJapan'").Rows;
                        if (item.Count > 0)
                        {
                            panel1.Enabled = Boolean.Parse(item[0][1].ToString());
                            panel1.Visible = Boolean.Parse(item[0][1].ToString());
                            dataGridView1.Columns[5].Visible = Boolean.Parse(item[0][1].ToString());
                        }
                        _Text = Settings.販售地區2.Split('/')[0];
                        unit = Settings.販售地區2.Split('/')[1];
                        _Type = Settings.類型2;
                    }
                    else if (((ButtonBase)sender).Text.Contains("F3"))
                    {
                        TableName = "ExportSupermarketUnitPrice";
                        var item = dB_SQLite.GetDataTable(DB_Path, @"SELECT * FROM Setting WHERE SettingName = 'ShowMoney_ExportSupermarket'").Rows;
                        if (item.Count > 0)
                        {
                            panel1.Enabled = Boolean.Parse(item[0][1].ToString());
                            panel1.Visible = Boolean.Parse(item[0][1].ToString());
                            dataGridView1.Columns[5].Visible = Boolean.Parse(item[0][1].ToString());
                        }
                        _Text = Settings.販售地區3.Split('/')[0];
                        unit = Settings.販售地區3.Split('/')[1];
                        _Type = Settings.類型3;
                    }
                    label1.Text = unit;
                    int _TypeCount = _Type.Split('/').Count();
                    int _Space = 10;
                    #region 8個類型的位置調整
                    if (_TypeCount >= 1)
                    {
                        radioButton1.Text = _Type.Split('/')[0] + "(F4)";
                        radioButton1.Enabled = true;
                        radioButton1.Visible = true;
                    }
                    else
                    {
                        radioButton1.Enabled = false;
                        radioButton1.Visible = false;
                    }
                    if (_TypeCount >= 2)
                    {
                        radioButton2.Text = _Type.Split('/')[1] + "(F5)";
                        radioButton2.Location = new Point(radioButton1.Location.X + radioButton1.Size.Width + _Space, radioButton2.Location.Y);
                        radioButton2.Enabled = true;
                        radioButton2.Visible = true;
                    }
                    else
                    {
                        radioButton2.Enabled = false;
                        radioButton2.Visible = false;
                    }
                    if (_TypeCount >= 3)
                    {
                        radioButton3.Text = _Type.Split('/')[2] + "(F6)";
                        radioButton3.Location = new Point(radioButton2.Location.X + radioButton2.Size.Width + _Space, radioButton2.Location.Y);
                        radioButton3.Enabled = true;
                        radioButton3.Visible = true;
                    }
                    else
                    {
                        radioButton3.Enabled = false;
                        radioButton3.Visible = false;
                    }
                    if (_TypeCount >= 4)
                    {
                        radioButton4.Text = _Type.Split('/')[3] + "(F7)";
                        radioButton4.Location = new Point(radioButton3.Location.X + radioButton3.Size.Width + _Space, radioButton3.Location.Y);
                        radioButton4.Enabled = true;
                        radioButton4.Visible = true;
                    }
                    else
                    {
                        radioButton4.Enabled = false;
                        radioButton4.Visible = false;
                    }
                    if (_TypeCount >= 5)
                    {
                        radioButton5.Text = _Type.Split('/')[4] + "(F8)";
                        radioButton5.Location = new Point(radioButton4.Location.X + radioButton4.Size.Width + _Space, radioButton4.Location.Y);
                        radioButton5.Enabled = true;
                        radioButton5.Visible = true;
                    }
                    else
                    {
                        radioButton5.Enabled = false;
                        radioButton5.Visible = false;
                    }
                    if (_TypeCount >= 6)
                    {
                        radioButton6.Text = _Type.Split('/')[5] + "(F9)";
                        radioButton6.Location = new Point(radioButton5.Location.X + radioButton5.Size.Width + _Space, radioButton5.Location.Y);
                        radioButton6.Enabled = true;
                        radioButton6.Visible = true;
                    }
                    else
                    {
                        radioButton6.Enabled = false;
                        radioButton6.Visible = false;
                    }
                    if (_TypeCount >= 7)
                    {
                        radioButton7.Text = _Type.Split('/')[6] + "(F10)";
                        radioButton7.Location = new Point(radioButton6.Location.X + radioButton6.Size.Width + _Space, radioButton6.Location.Y);
                        radioButton7.Enabled = true;
                        radioButton7.Visible = true;
                    }
                    else
                    {
                        radioButton7.Enabled = false;
                        radioButton7.Visible = false;
                    }
                    if (_TypeCount >= 8)
                    {
                        radioButton11.Text = _Type.Split('/')[7] + "(F11)";
                        radioButton11.Location = new Point(radioButton7.Location.X + radioButton7.Size.Width + _Space, radioButton7.Location.Y);
                        radioButton11.Enabled = true;
                        radioButton11.Visible = true;
                    }
                    else
                    {
                        radioButton11.Enabled = false;
                        radioButton11.Visible = false;
                    }
                    #endregion

                    unitPrice = "";
                    type = "";
                    salesArea = _Text;
                    label1.Text = unit;
                    ((GroupBox)((RadioButton)sender).Parent).BackColor = SystemColors.Control;
                    for (int i = 0; i < ((RadioButton)sender).Parent.Controls.Count; i++)
                    {
                        if (((RadioButton)sender).Parent.Controls[i].GetType().Name == "ComboBox")
                            continue;
                        if (((RadioButton)((RadioButton)sender).Parent.Controls[i]).Checked)
                            ((RadioButton)((RadioButton)sender).Parent.Controls[i]).BackColor = Color.GreenYellow;
                        else
                            ((RadioButton)((RadioButton)sender).Parent.Controls[i]).BackColor = SystemColors.Control;
                    }
                    for (int i = 0; i < groupBox1.Controls.Count; i++)
                    {
                        //清空類型內選項
                        if (((RadioButton)(groupBox1.Controls[i])).Checked)
                        {
                            ((RadioButton)(groupBox1.Controls[i])).Checked = false;
                            ((RadioButton)(groupBox1.Controls[i])).BackColor = SystemColors.Control;
                        }
                        label3.Text = "0";
                    }

                    log.LogMessage("販售地點設定，資料表名稱設定，重量單位設定，針對不同地區客製化功能 成功", enumLogType.Trace);
                }
                //使用者視窗_Tab其他移除，保留輸入框與列印
                ((RadioButton)sender).TabStop = false;
            }
            catch (Exception ee)
            {
                log.LogMessage("販售地點設定，資料表名稱設定，重量單位設定，針對不同地區客製化功能 失敗：\r\n" + ee.Message, enumLogType.Error);
            }
        }
        private void radioButton_Click(object sender, EventArgs e)
        {
            try
            {
                if (salesArea_Checked_OK == -1)
                {
                    log.LogMessage("更新客戶姓名資料庫 開始", enumLogType.Trace);
                    comboBox1.Items.Clear();
                    // 讀取資料
                    foreach (DataRow item in dB_SQLite.GetDataTable(DB_Path, $@"SELECT CustomerID, CustomerName FROM CustomerProfile;").Rows)
                    {
                        comboBox1.Items.Add(item[0].ToString() + "_" + item[1].ToString());
                    }
                    log.LogMessage("更新客戶姓名資料庫 成功", enumLogType.Info);
                    log.LogMessage("更新客戶姓名資料庫 成功", enumLogType.Trace);
                }
            }
            catch (Exception ee)
            {
                log.LogMessage("更新客戶姓名 失敗\r\n" + ee.Message, enumLogType.Error);
                MessageBox.Show("更新客戶姓名 失敗\r\n" + ee.Message);
                return;
            }
        }

        //確認後寫入DB
        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            Boolean _OK = true;
            if (string.IsNullOrEmpty(type))
            {
                groupBox1.BackColor = Color.IndianRed;
                _OK = false;
            }
            if (string.IsNullOrEmpty(salesArea))
            {
                groupBox2.BackColor = Color.IndianRed;
                _OK = false;
            }
            if (string.IsNullOrEmpty(unitPrice) || unitPrice == "0")
            {
                if (!panel1.Visible)    //隱藏單價情況就跳出通知
                    MessageBox.Show("請等待管理者輸入單價！");
                panel1.BackColor = Color.IndianRed;
                _OK = false;
            }
            if (string.IsNullOrEmpty(comboBox1.Text))
            {
                panel2.BackColor = Color.IndianRed;
                _OK = false;
            }
            if (!string.IsNullOrEmpty(label5.Text))
            {
                dataGridView1.Rows.Clear();
            }
            if (string.IsNullOrEmpty(textBox1.Text))
            {
                _OK = false;
            }

            if (_OK && e.KeyChar == ((char)Keys.Enter))
            {
                log.LogMessage("確認 開始", enumLogType.Trace);

                label5.Text = "";
                DataGridViewRow row = new DataGridViewRow();
                DateTime now = DateTime.Now;

                #region DataGridView修改

                //重量
                Double weight = Convert.ToDouble(((TextBox)sender).Text);
                if (Int32.TryParse(textBox3.Text, out Int32 countBuff))
                {
                    if (Double.TryParse(textBox2.Text, out Double weightBuff))
                    {
                        weight = weight - (countBuff * weightBuff);
                    }
                }
                //重繪才能讀取到別的使用者登錄的資料
                row.CreateCells(dataGridView1);
                row.SetValues(new string[] { "", now.ToString("yyyy-MM-dd HH:mm:ss"), comboBox1.Text, type
                    , (Convert.ToDouble((int)(weight * 100)) / 100).ToString(), unitPrice, unit, salesArea
                    , countBuff.ToString() });
                dataGridView1.Rows.Insert(0, row);
                dataGridView1.Rows[0].Selected = true;
                dataGridView1.CurrentCell = dataGridView1.Rows[0].Cells[0];

                textBox1.Text = "";
                textBox1.Focus();

                log.LogMessage("確認_DB新增 成功", enumLogType.Info);
                #endregion

                log.LogMessage("確認後寫入DataGridView 成功", enumLogType.Trace);
            }
        }
        //快捷鍵指向
        private void textBox1_KeyDown(object sender, KeyEventArgs e)
        {
            Boolean isCheck = false;
            switch (e.KeyValue)
            {
                case ((char)Keys.Up):
                    textBox1.Focus();
                    return;
                case ((char)Keys.Down):
                    textBox3.Focus();
                    return;
                case ((char)Keys.Left):
                    for (int i = 0; i < groupBox1.Controls.Count; i++)
                    {
                        if (groupBox1.Controls[i].GetType().Name == "RadioButton")
                        {
                            if (((RadioButton)groupBox1.Controls[i]).Checked)
                            {
                                isCheck = true;
                                if ((i + 1) >= 0 && (i + 1) < groupBox1.Controls.Count)
                                    ((RadioButton)groupBox1.Controls[i + 1]).Checked = true;
                                break;
                            }
                        }
                    }
                    if (!isCheck) ((RadioButton)groupBox1.Controls[groupBox1.Controls.Count - 1]).Checked = true;
                    break;
                case ((char)Keys.Right):
                    for (int i = 0; i < groupBox1.Controls.Count; i++)
                    {
                        if (groupBox1.Controls[i].GetType().Name == "RadioButton")
                        {
                            if (((RadioButton)groupBox1.Controls[i]).Checked)
                            {
                                isCheck = true;
                                if ((i - 1) >= 0 && (i - 1) < groupBox1.Controls.Count)
                                    ((RadioButton)groupBox1.Controls[i - 1]).Checked = true;
                                break;
                            }
                        }
                    }
                    if (!isCheck) ((RadioButton)groupBox1.Controls[groupBox1.Controls.Count - 1]).Checked = true;
                    break;
                case ((char)Keys.F1):
                    if (radioButton8.Visible)
                        radioButton8.Checked = true;
                    break;
                case ((char)Keys.F2):
                    if (radioButton9.Visible)
                        radioButton9.Checked = true;
                    break;
                case ((char)Keys.F3):
                    if (radioButton10.Visible)
                        radioButton10.Checked = true;
                    break;
                case ((char)Keys.F4):
                    if (radioButton1.Visible)
                        radioButton1.Checked = true;
                    break;
                case ((char)Keys.F5):
                    if (radioButton2.Visible)
                        radioButton2.Checked = true;
                    break;
                case ((char)Keys.F6):
                    if (radioButton3.Visible)
                        radioButton3.Checked = true;
                    break;
                case ((char)Keys.F7):
                    if (radioButton4.Visible)
                        radioButton4.Checked = true;
                    break;
                case ((char)Keys.F8):
                    if (radioButton5.Visible)
                        radioButton5.Checked = true;
                    break;
                case ((char)Keys.F9):
                    if (radioButton6.Visible)
                        radioButton6.Checked = true;
                    break;
                case ((char)Keys.F10):
                    if (radioButton7.Visible)
                        radioButton7.Checked = true;
                    //F10會呼叫出控制列視窗，需要按兩次ESC做取消
                    SendKeys.SendWait("{ESC}");
                    SendKeys.SendWait("{ESC}");
                    SendKeys.SendWait("F10");
                    timer_ESC.Start();
                    break;
                case ((char)Keys.F11):
                    if (radioButton11.Visible)
                        radioButton11.Checked = true;
                    break;
                case ((char)Keys.F12):
                    if (comboBox1.Visible)
                    {
                        comboBox1.Focus();
                        panel2.BackColor = Color.DodgerBlue;
                        return;
                    }
                    break;
                case ((char)Keys.Space):
                    break;
                default:
                    break;
            }
            textBox1.Focus();
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            if (textBox1.Text == "")
                return;
            if (textBox1.Text == "0")
                return;
            if (!Double.TryParse(textBox1.Text, out Double _buffDouble))
            {
                MessageBox.Show("請輸入數字");
                textBox1.Text = "";
            }
            if (!Int32.TryParse((_buffDouble * 100).ToString(), out Int32 _buffInt))
            {
                MessageBox.Show("小數點最多兩位數");
                textBox1.Text = (Convert.ToDouble((int)(_buffDouble * 100)) / 100).ToString();
            }
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            if (textBox2.Text == "")
                return;
            if (!Double.TryParse(textBox2.Text, out Double _buffDouble))
            {
                MessageBox.Show("請輸入數字");
                textBox2.Text = "";
            }
            if (!Int32.TryParse((_buffDouble * 10).ToString(), out Int32 _buffInt))
            {
                MessageBox.Show("小數點最多一位數");
                textBox2.Text = (Convert.ToDouble((int)(_buffDouble * 10)) / 10).ToString();
            }
        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {
            if (textBox3.Text == "")
                return;
            if (!Double.TryParse(textBox3.Text, out Double _buffDouble))
            {
                MessageBox.Show("請輸入數字");
                textBox3.Text = "";
            }
            if (!Int32.TryParse((_buffDouble).ToString(), out Int32 _buffInt))
            {
                MessageBox.Show("只能輸入整數");
                textBox3.Text = (Convert.ToDouble((int)(_buffDouble))).ToString();
            }
        }

        private void dataGridView1_UserDeletedRow(object sender, DataGridViewRowEventArgs e)
        {
            //刪除失敗則還原
            try
            {
                //log.LogMessage("刪除失敗則還原 開始", enumLogType.Trace);
                //dB_SQLite.Manipulate(DB_Path, $@"DELETE FROM SalesRecord WHERE No = '{e.Row.Cells[0].Value.ToString()}';");
                //log.LogMessage($@"刪除失敗則還原 成功：No = '{e.Row.Cells[0].Value.ToString()}'", enumLogType.Trace);
                //log.LogMessage($@"刪除失敗則還原 成功：No = '{e.Row.Cells[0].Value.ToString()}'", enumLogType.Info);
                textBox1.Focus();
            }
            catch (Exception ee)
            {
                DB_SQLite.DatatableToDatagridview(dB_SQLite.GetDataTable(DB_Path, $@"SELECT No, Time, Name, Type, Count, UnitPrice, Unit, SalesArea FROM SalesRecord WHERE Date > '{DateTime.Now.ToString("yyyy-MM-dd")}';"), dataGridView1);
                log.LogMessage("刪除失敗則還原 失敗：" + ee.Message, enumLogType.Trace);
                MessageBox.Show("刪除失敗：" + ee.Message);
                return;
            }
        }

        //未付款列印
        private void button1_Click(object sender, EventArgs e)
        {
            textBox1.Focus();

            if (label5.Text != "")
            {
                if (DialogResult.Yes == MessageBox.Show("是否重複列印", "重複列印", MessageBoxButtons.YesNo))
                {
                    _Page = 1;
                    //宣告一個印表機
                    PrintDocument printDocument = new PrintDocument();
                    //設定印表機邊界
                    Margins margin = new Margins(0, 0, 0, 0);
                    printDocument.DefaultPageSettings.Margins = margin;
                    //印表機事件設定
                    printDocument.PrintPage += PrintDocument_PrintPage;
                    printDocument.PrinterSettings.PrinterName = Settings.印表機名稱;
                    printDocument.Print();   //列印

                    log.LogMessage("重複列印 成功", enumLogType.Trace);
                    log.LogMessage("重複列印 成功", enumLogType.Info);
                }
                return;
            }

            button1.Enabled = false;
            button3.Enabled = false;
            DateTime _Now = DateTime.Now; //時間

            InsertSalesRecord(_Now, out string _No, out string _Name, out string _Unit, out string _SalesArea);

            ExcelProcess excel = new ExcelProcess(log);
            if (excel.ExcelExportImage(dataGridView1, $@"{Settings.Excel路徑}{_No}_{_Name}.xlsx", _Now, _No, _Name, _Unit, _SalesArea, panel1.Visible))
            {
                try
                {
                    #region 列印
                    _Page = 1;
                    //宣告一個印表機
                    PrintDocument printDocument = new PrintDocument();
                    //設定印表機邊界
                    Margins margin = new Margins(0, 0, 0, 0);
                    printDocument.DefaultPageSettings.Margins = margin;
                    //印表機事件設定
                    printDocument.PrintPage += PrintDocument_PrintPage;
                    printDocument.PrinterSettings.PrinterName = Settings.印表機名稱;
                    //printDocument.DefaultPageSettings.Landscape = true;           //此处更改页面为横向打印 
                    printDocument.Print();   //列印
                    #endregion
                }
                catch (Exception ee)
                {
                    log.LogMessage("列印 失敗：\r\n" + ee.Message, enumLogType.Error);
                    button1.Enabled = true;
                    button3.Enabled = true;
                    dB_SQLite.Manipulate(DB_Path, $@"DELETE FROM SalesRecord WHERE No = '{_No}';");
                    _No = "";
                }
            }
            else
            {
                button1.Enabled = true;
                button3.Enabled = true;
                dB_SQLite.Manipulate(DB_Path, $@"DELETE FROM SalesRecord WHERE No = '{_No}';");
                _No = "";
            }

            comboBox1.SelectedIndex = -1;
            label5.Text = _No;
        }
        //已付款列印
        private void button3_Click(object sender, EventArgs e)
        {
            textBox1.Focus();

            if (label5.Text != "")
            {
                if (DialogResult.Yes == MessageBox.Show("是否重複已付款列印", "重複已付款列印", MessageBoxButtons.YesNo))
                {
                    UpdateNoPaid(label5.Text);

                    _Page = 1;
                    //宣告一個印表機
                    PrintDocument printDocument = new PrintDocument();
                    //設定印表機邊界
                    Margins margin = new Margins(0, 0, 0, 0);
                    printDocument.DefaultPageSettings.Margins = margin;
                    //印表機事件設定
                    printDocument.PrintPage += PrintDocument_PrintPage;
                    printDocument.PrinterSettings.PrinterName = Settings.印表機名稱;
                    printDocument.Print();   //列印

                    log.LogMessage("已付款重複列印 成功", enumLogType.Trace);
                    log.LogMessage("已付款重複列印 成功", enumLogType.Info);
                    return;
                }
                return;
            }

            button1.Enabled = false;
            button3.Enabled = false;
            DateTime _Now = DateTime.Now; //時間

            InsertSalesRecord(_Now, out string _No, out string _Name, out string _Unit, out string _SalesArea);

            ExcelProcess excel = new ExcelProcess(log);
            if (excel.ExcelExportImage(dataGridView1, $@"{Settings.Excel路徑}{_No}_{_Name}.xlsx", _Now, _No, _Name, _Unit, _SalesArea, panel1.Visible))
            {
                try
                {
                    #region 列印
                    _Page = 1;
                    //宣告一個印表機
                    PrintDocument printDocument = new PrintDocument();
                    //設定印表機邊界
                    Margins margin = new Margins(0, 0, 0, 0);
                    printDocument.DefaultPageSettings.Margins = margin;
                    //印表機事件設定
                    printDocument.PrintPage += PrintDocument_PrintPage;
                    printDocument.PrinterSettings.PrinterName = Settings.印表機名稱;
                    //printDocument.DefaultPageSettings.Landscape = true;           //此处更改页面为横向打印 
                    printDocument.Print();   //列印
                    #endregion
                }
                catch (Exception ee)
                {
                    log.LogMessage("列印 失敗：\r\n" + ee.Message, enumLogType.Error);
                    button1.Enabled = true;
                    button3.Enabled = true;
                    dB_SQLite.Manipulate(DB_Path, $@"DELETE FROM SalesRecord WHERE No = '{_No}';");
                    _No = "";
                }
            }
            else
            {
                button1.Enabled = true;
                button3.Enabled = true;
                dB_SQLite.Manipulate(DB_Path, $@"DELETE FROM SalesRecord WHERE No = '{_No}';");
                _No = "";
            }

            comboBox1.SelectedIndex = -1;
            UpdateNoPaid(_No);
            label5.Text = _No;
        }
        private void InsertSalesRecord(DateTime _Now, out string _No, out string _Name, out string _Unit, out string _SalesArea)
        {
            _No = ""; _Name = ""; _Unit = ""; _SalesArea = "";
            try
            {
                ///取單號
                DataTable dataTable = dB_SQLite.GetDataTable(DB_Path, $@"
                    SELECT CASE WHEN MAX(No) ISNULL THEN '{_Now.ToString("yyyyMMdd") + "001"}' ELSE MAX(No)+1 END No
                    FROM SalesRecord WHERE Time > '{_Now.ToString("yyyy-MM-dd")}';");
                _No = dataTable.Rows[0][0].ToString();

                string insertstring = $@"INSERT INTO SalesRecord (No, Time, Name, Type, Count, UnitPrice, Unit, salesArea, BasketCount) VALUES";
                /// 插入資料
                /// 可自動抓取新單號新增
                foreach (DataGridViewRow row in dataGridView1.Rows)
                {
                    //暫存資料，用於匯出Excel時統合
                    if (_Name == "")
                        _Name = row.Cells[2].Value.ToString();
                    else if (!row.Cells[2].Value.ToString().Split('，').Contains(row.Cells[2].Value.ToString()))
                        _Name += row.Cells[2].Value.ToString();
                    if (_Unit == "")
                        _Unit = row.Cells[6].Value.ToString();
                    else if (!row.Cells[6].Value.ToString().Split('，').Contains(row.Cells[6].Value.ToString()))
                        _Unit += row.Cells[6].Value.ToString();
                    if (_SalesArea == "")
                        _SalesArea = row.Cells[7].Value.ToString();
                    else if (!row.Cells[7].Value.ToString().Split('，').Contains(row.Cells[7].Value.ToString()))
                        _SalesArea += row.Cells[7].Value.ToString();

                    insertstring += $@" ('{_No}', '{Convert.ToDateTime(row.Cells[1].Value).ToString("yyyy-MM-dd HH:mm:ss")}', '{row.Cells[2].Value.ToString()}', 
                            '{row.Cells[3].Value.ToString()}', '{row.Cells[4].Value.ToString()}', '{row.Cells[5].Value.ToString()}', 
                            '{row.Cells[6].Value.ToString()}', '{row.Cells[7].Value.ToString()}', '{row.Cells[8].Value.ToString()}') ,";
                }
                insertstring = insertstring.Remove(insertstring.Length - 1, 1);
                dB_SQLite.Manipulate(DB_Path, insertstring);

                DB_SQLite.DatatableToDatagridview(dB_SQLite.GetDataTable(DB_Path, $@"SELECT No, Time, Name, Type, 
                    Count, UnitPrice, Unit, salesArea, BasketCount FROM SalesRecord 
                    WHERE Time > '{DateTime.Now.ToString("yyyy-MM-dd")}' AND No = '{_No}';"), dataGridView1);

                log.LogMessage("確認_DB新增 成功路徑：" + DB_Path + "\r\n語法：" + insertstring, enumLogType.Trace);
                log.LogMessage("確認_DB新增 成功單號：" + _No, enumLogType.Info);
            }
            catch (Exception ee)
            {
                MessageBox.Show("DB新增 失敗：\r\n" + ee.Message);
                log.LogMessage("確認_DB新增 失敗：\r\n" + ee.Message, enumLogType.Error);
                button1.Enabled = true;
                return;
            }
        }
        int _Page = 1;
        int _PageHeight = 0;
        private void PrintDocument_PrintPage(object sender, PrintPageEventArgs e)
        {
            #region 
            GC.Collect();
            using (Bitmap bitmap = new Bitmap(@".\ianimage.png"))
            {
                int _Y = 0; //輸出圖片時要從Y軸哪個點開始
                int _1PageSize = 1098;
                int _2PageSize = 1085;
                Rectangle newarea = new Rectangle();
                newarea.X = 0;
                newarea.Y = 0;
                newarea.Width = bitmap.Width;
                newarea.Height = bitmap.Height - 30;
                //第一頁
                if (_Page == 1)
                {
                    _PageHeight = newarea.Height;
                    if (_PageHeight > _1PageSize)
                    {
                        e.HasMorePages = true; //此处打开多页打印属性}
                        _PageHeight = _PageHeight - _1PageSize;
                        newarea.Height = _1PageSize;
                    }
                    else
                    {
                        newarea.Height = _PageHeight;
                    }
                }
                else if (_Page >= 2)
                {
                    //_Y-2是為了格子上方那條線
                    _Y = bitmap.Height - _PageHeight - 2;

                    newarea.Y = 30;
                    if (_PageHeight > _2PageSize)
                    {
                        e.HasMorePages = true; //此处打开多页打印属性}
                        _PageHeight = _PageHeight - _2PageSize + 2;
                        newarea.Height = _2PageSize;
                    }
                    else
                    {
                        newarea.Height = _PageHeight;
                    }
                }


                int _width = newarea.Width + 60;
                newarea.Width = newarea.Width + 60;
                if (_width > 810)
                {
                    _width = 810;
                    newarea.Width = 810;
                }
                e.Graphics.DrawImage(bitmap, newarea, 0 - 60, _Y, _width, newarea.Height, GraphicsUnit.Pixel);
                _Page++;
                if (!e.HasMorePages)
                {
                    button1.Enabled = true;
                    button3.Enabled = true;
                }
            }
            #endregion
        }
        //單號已付
        private void UpdateNoPaid(string No)
        {
            if (String.IsNullOrEmpty(No))
                return;
            string _SQL = $@"SELECT SUM(Count * UnitPrice)AS SumUnpaid FROM SalesRecord 
                WHERE No = '{No}';";
            var _Buff = dB_SQLite.GetDataTable(DB_Path, _SQL).Rows;
            if (_Buff.Count > 0)
            {
                var Paid = "";
                Paid = dB_SQLite.GetDataTable(DB_Path, _SQL).Rows[0][0].ToString();
                UpdateNoPaid(No, Paid);
            }
        }
        private void UpdateNoPaid(string No, string Paid)
        {
            try
            {
                DateTime _now = DateTime.Now;
                string _UpdateSQL = $@"UPDATE SalesRecord SET PaidTime = '{_now.ToString("yyyy-MM-dd HH:mm:ss")}'
                            , Paid = '{(int)Math.Round(Convert.ToDouble(Paid), 0, MidpointRounding.AwayFromZero)}'
                            WHERE No = '{No}'";

                dB_SQLite.Manipulate(DB_Path, _UpdateSQL);
                log.LogMessage("已付修改 成功路徑：" + DB_Path + "\r\n語法：" + _UpdateSQL, enumLogType.Trace);
            }
            catch (Exception ee)
            {
                throw ee;
            }
        }

        string _comboBoxSelectText = "";
        Boolean _comboBoxKeyPressSet = false;
        private void timer_ComboBoxSelect_Tick(object sender, EventArgs e)
        {
            //textBox1.Text = _comboBoxSelectText = "";
            //timer_ComboBoxSelect.Stop();
            //panel2.BackColor = SystemColors.Control;
            //textBox1.Focus();
        }


        private void comboBox1_KeyDown(object sender, KeyEventArgs e)
        {
            //不知道為什麼按一次按鍵，會執行兩次，所以多做判斷
            _comboBoxKeyPressSet = true;
        }
        private void comboBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            //不知道為什麼按一次按鍵，會執行兩次，所以多做判斷
            if (_comboBoxKeyPressSet)
            {
                if (e.KeyChar == (char)Keys.Enter)
                {
                    textBox1.Text = _comboBoxSelectText = "";
                    panel2.BackColor = SystemColors.Control;
                    textBox1.Focus();
                    //timer_ComboBoxSelect.Stop();
                    //timer_ComboBoxSelect_Tick(sender, e);
                    return;
                }
                int _key = -1;
                //if (!timer_ComboBoxSelect.Enabled)
                //{
                //    timer_ComboBoxSelect.Start();
                //}
                if (Int32.TryParse(e.KeyChar.ToString(), out _key))
                {
                    _comboBoxSelectText += _key;
                    textBox1.Text = _comboBoxSelectText;
                    foreach (var item in comboBox1.Items)
                    {
                        string itemText = item.ToString();
                        if (itemText.Substring(1).StartsWith(_comboBoxSelectText))
                        {
                            comboBox1.Text = itemText;
                            break;
                        }
                    }
                }
            }
            _comboBoxKeyPressSet = false;
        }

        private void timer_FocusTextBox1_Tick(object sender, EventArgs e)
        {
            if (comboBox1.Focused)
                return;
            else if (button1.Focused)
                return;
            else if (dataGridView1.Focused)
                return;
            else if (comboBox2.Focused)
                return;
            //if (!textBox1.Focused) 
            //    textBox1.Focus();
        }

        private void button_Enter(object sender, EventArgs e)
        {
            ((Button)sender).BackColor = SystemColors.ActiveCaption;
        }
        private void button_Leave(object sender, EventArgs e)
        {
            if (((Button)sender).Name == "button2")
                ((Button)sender).BackColor = SystemColors.ActiveBorder;
            else
                ((Button)sender).BackColor = SystemColors.Control;
        }
        private void textBox_Enter(object sender, EventArgs e)
        {
            ((TextBox)sender).Parent.BackColor = SystemColors.ActiveCaption;
            if (((TextBox)sender).Name == "textBox3")
            {
                ((TextBox)sender).Text = "";
            }
        }
        private void textBox_Leave(object sender, EventArgs e)
        {
            ((TextBox)sender).Parent.BackColor = SystemColors.Control;
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox2.SelectedIndex == 0)
                salesArea = radioButton9.Text.Split('(')[0];
            else if (salesArea.Split('_')[0] == radioButton9.Text.Split('(')[0])
                salesArea = radioButton9.Text.Split('(')[0] + "_" + comboBox2.Text;
            textBox1.Focus();
        }

        //刪除
        private void button2_Click(object sender, EventArgs e)
        {
            if (dataGridView1.Rows.Count > 0)
            {
                if (DialogResult.Yes == MessageBox.Show("是否刪除首筆資料！", "刪除", MessageBoxButtons.YesNo))
                {
                    dataGridView1.Rows.RemoveAt(0);
                }
            }
        }

        private void timer_ESC_Tick(object sender, EventArgs e)
        {
            SendKeys.SendWait("F10");
            SendKeys.SendWait("{ESC}");
            //_ESC = true;
            timer_ESC.Stop();
        }
    }
}
