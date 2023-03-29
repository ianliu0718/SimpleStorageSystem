using OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Information;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;
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
        /// 販售地點
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
                    Settings.主機序號 = EncryptionDecryption.desEncryptBase64(Settings.主機序號);
                    MessageBox.Show("綁定成功");
                    log.LogMessage("比對 CPU ID 綁定 成功", enumLogType.Info);
                    log.LogMessage("比對 CPU ID 綁定 成功", enumLogType.Trace);
                }
                else if (EncryptionDecryption.desDecryptBase64(Settings.主機序號) != GetPCMacID.GetCpuID())
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
            log.LogMessage("使用者介面啓動", enumLogType.Info);
        }

        //類型設定，單價設定
        private void radioButton_type_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                if (((RadioButton)sender).Checked)
                {
                    log.LogMessage("類型設定，單價設定 開始", enumLogType.Trace);

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
            }
            catch (Exception ee)
            {
                log.LogMessage("類型設定，單價設定 失敗：\r\n" + ee.Message, enumLogType.Error);
            }
        }

        //販售地點設定，資料表名稱設定，重量單位設定，針對不同地區客製化功能
        int salesArea_Checked_OK = -3;  //-3第一次設定 //-2設定中 //-1未設定 //0設定為No //1設定為Yes
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

                string _Text = ((ButtonBase)sender).Text;
                if (((RadioButton)sender).Checked)
                {
                    dataGridView1.Rows.Clear();

                    //針對不同地區客製化功能
                    if (_Text.Contains("外銷韓國"))
                    {
                        _Text = "外銷韓國";
                        TableName = "ExportKoreaUnitPrice";
                        var item = dB_SQLite.GetDataTable(DB_Path, @"SELECT * FROM Setting WHERE SettingName = 'ShowMoney_ExportKorea'").Rows;
                        if (item.Count > 0)
                        {
                            panel1.Enabled = Boolean.Parse(item[0][1].ToString());
                            panel1.Visible = Boolean.Parse(item[0][1].ToString());
                            dataGridView1.Columns[5].Visible = Boolean.Parse(item[0][1].ToString());
                        }
                        unit = "公斤";
                        radioButton1.Text = "9粒(F5)";
                        radioButton2.Text = "12粒(F6)";
                        radioButton3.Text = "14粒(F7)";
                        radioButton4.Text = "16粒(F8)";
                        radioButton5.Text = "20粒(F9)";
                        radioButton6.Text = "24粒(F10)";
                        radioButton7.Text = "小(F11)";
                        radioButton7.Enabled = true;
                        radioButton7.Visible = true;
                    }
                    else if (_Text.Contains("外銷日本"))
                    {
                        _Text = "外銷日本";
                        TableName = "ExportJapanUnitPrice";
                        var item = dB_SQLite.GetDataTable(DB_Path, @"SELECT * FROM Setting WHERE SettingName = 'ShowMoney_ExportJapan'").Rows;
                        if (item.Count > 0)
                        {
                            panel1.Enabled = Boolean.Parse(item[0][1].ToString());
                            panel1.Visible = Boolean.Parse(item[0][1].ToString());
                            dataGridView1.Columns[5].Visible = Boolean.Parse(item[0][1].ToString());
                        }
                        unit = "公斤";
                        radioButton1.Text = "9粒(F5)";
                        radioButton2.Text = "12粒(F6)";
                        radioButton3.Text = "14粒(F7)";
                        radioButton4.Text = "16粒(F8)";
                        radioButton5.Text = "20粒(F9)";
                        radioButton6.Text = "24粒(F10)";
                        radioButton7.Text = "小(F11)";
                        radioButton7.Enabled = true;
                        radioButton7.Visible = true;
                    }
                    else if (_Text.Contains("超市"))
                    {
                        _Text = "超市";
                        TableName = "ExportSupermarketUnitPrice";
                        var item = dB_SQLite.GetDataTable(DB_Path, @"SELECT * FROM Setting WHERE SettingName = 'ShowMoney_ExportSupermarket'").Rows;
                        if (item.Count > 0)
                        {
                            panel1.Enabled = Boolean.Parse(item[0][1].ToString());
                            panel1.Visible = Boolean.Parse(item[0][1].ToString());
                            dataGridView1.Columns[5].Visible = Boolean.Parse(item[0][1].ToString());
                        }
                        unit = "台斤";
                        radioButton1.Text = "12粒(F5)";
                        radioButton2.Text = "15粒(F6)";
                        radioButton3.Text = "18粒(F7)";
                        radioButton4.Text = "20粒(F8)";
                        radioButton5.Text = "24粒(F9)";
                        radioButton6.Text = "28粒(F10)";
                        radioButton7.Text = "小(F11)";
                        radioButton7.Enabled = false;
                        radioButton7.Visible = false;
                    }

                    unitPrice = "";
                    type = "";
                    salesArea = _Text;
                    label1.Text = unit;
                    ((GroupBox)((RadioButton)sender).Parent).BackColor = SystemColors.Control;
                    for (int i = 0; i < ((RadioButton)sender).Parent.Controls.Count; i++)
                    {
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
            }
            catch (Exception ee)
            {
                log.LogMessage("販售地點設定，資料表名稱設定，重量單位設定，針對不同地區客製化功能 失敗：\r\n" + ee.Message, enumLogType.Error);
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

            if (_OK && e.KeyChar == ((char)Keys.Enter))
            {
                log.LogMessage("確認 開始", enumLogType.Trace);

                label5.Text = "";
                DataGridViewRow row = new DataGridViewRow();
                DateTime now = DateTime.Now;

                #region DataGridView修改

                //重繪才能讀取到別的使用者登錄的資料
                row.CreateCells(dataGridView1);
                row.SetValues(new string[] { "", now.ToString("yyyy-MM-dd HH:mm:ss"), comboBox1.Text, type, ((TextBox)sender).Text, unitPrice, unit, salesArea });
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
            switch (e.KeyValue)
            {
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
                    break;
                case ((char)Keys.F5):
                    if (radioButton1.Visible)
                        radioButton1.Checked = true;
                    break;
                case ((char)Keys.F6):
                    if (radioButton2.Visible)
                        radioButton2.Checked = true;
                    break;
                case ((char)Keys.F7):
                    if (radioButton3.Visible)
                        radioButton3.Checked = true;
                    break;
                case ((char)Keys.F8):
                    if (radioButton4.Visible)
                        radioButton4.Checked = true;
                    break;
                case ((char)Keys.F9):
                    if (radioButton5.Visible)
                        radioButton5.Checked = true;
                    break;
                case ((char)Keys.F10):
                    if (radioButton6.Visible)
                        radioButton6.Checked = true;
                    break;
                case ((char)Keys.F11):
                    if (radioButton7.Visible)
                        radioButton7.Checked = true;
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

        private void dataGridView1_UserDeletedRow(object sender, DataGridViewRowEventArgs e)
        {
            log.LogMessage("刪除失敗則還原 開始", enumLogType.Trace);

            //刪除失敗則還原
            try
            {
                dB_SQLite.Manipulate(DB_Path, $@"DELETE FROM SalesRecord WHERE No = '{e.Row.Cells[0].Value.ToString()}';");
                log.LogMessage("刪除失敗則還原 成功", enumLogType.Trace);
            }
            catch (Exception ee)
            {
                DB_SQLite.DatatableToDatagridview(dB_SQLite.GetDataTable(DB_Path, $@"SELECT No, Date, Name, Type, Count, UnitPrice, Unit, SalesArea FROM SalesRecord WHERE Date > '{DateTime.Now.ToString("yyyy-MM-dd")}';"), dataGridView1);
                log.LogMessage("刪除失敗則還原 失敗：" + ee.Message, enumLogType.Trace);
                MessageBox.Show("刪除失敗：" + ee.Message);
                return;
            }
        }

        private void comboBox1_TextChanged(object sender, EventArgs e)
        {
            //panel2.BackColor = SystemColors.Control;
        }

        //列印
        private void button1_Click(object sender, EventArgs e)
        {
            textBox1.Focus();

            Boolean RepeatPrinting = false;
            if (label5.Text != "")
            {
                if (DialogResult.No == MessageBox.Show("是否重複列印", "重複列印", MessageBoxButtons.YesNo))
                {
                    return;
                }
                RepeatPrinting = true;
            }

            button1.Enabled = false;
            string _No = ""; //單號
            DateTime _Now = DateTime.Now; //時間
            string _Name = ""; //姓名
            string _Unit = ""; //單位
            string _SalesArea = ""; //販售地區

            #region DB新增
            try
            {
                ///取單號
                if (RepeatPrinting)
                {
                    _No = label5.Text;
                }
                else
                {
                    DataTable dataTable = dB_SQLite.GetDataTable(DB_Path, $@"
                    SELECT CASE WHEN MAX(No) ISNULL THEN '{_Now.ToString("yyyyMMdd") + "001"}' ELSE MAX(No)+1 END No
                    FROM SalesRecord WHERE Date > '{_Now.ToString("yyyy-MM-dd")}';");
                    _No = dataTable.Rows[0][0].ToString();
                }

                string insertstring = $@"INSERT INTO SalesRecord (No, Date, Name, Type, Count, UnitPrice, Unit, salesArea) VALUES";
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

                    if (!RepeatPrinting)
                        insertstring += $@" ('{_No}', '{Convert.ToDateTime(row.Cells[1].Value).ToString("yyyy-MM-dd HH:mm:ss")}', '{row.Cells[2].Value.ToString()}', 
                            '{row.Cells[3].Value.ToString()}', '{row.Cells[4].Value.ToString()}', '{row.Cells[5].Value.ToString()}', 
                            '{row.Cells[6].Value.ToString()}', '{row.Cells[7].Value.ToString()}') ,";
                }
                if (!RepeatPrinting) insertstring = insertstring.Remove(insertstring.Length - 1, 1);
                if (!RepeatPrinting) dB_SQLite.Manipulate(DB_Path, insertstring);

                DB_SQLite.DatatableToDatagridview(dB_SQLite.GetDataTable(DB_Path, $@"SELECT No, Date, Name, Type, 
                    Count, UnitPrice, Unit, salesArea FROM SalesRecord 
                    WHERE Date > '{DateTime.Now.ToString("yyyy-MM-dd")}' AND No = '{_No}';"), dataGridView1);

                if (RepeatPrinting)
                    insertstring = "重複列印單號：" + _No;
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
            #endregion

            #region 列印
            try
            {
                EPPlus ePPlus = new EPPlus();
                List<List<MExcelCell>> excelCells = new List<List<MExcelCell>>();
                List<MExcelCell> excelCell = new List<MExcelCell>();
                DataGridView view = dataGridView1;

                //標頭
                excelCell.Add(new MExcelCell() { Content = "單號" });
                excelCell.Add(new MExcelCell() { Content = _No });
                excelCells.Add(excelCell);
                excelCell = new List<MExcelCell>();
                excelCell.Add(new MExcelCell() { Content = "日期" });
                excelCell.Add(new MExcelCell() { Content = _Now.ToString("yyyy-MM-dd") });
                excelCells.Add(excelCell);
                excelCell = new List<MExcelCell>();
                excelCell.Add(new MExcelCell() { Content = "姓名" });
                excelCell.Add(new MExcelCell() { Content = _Name });
                excelCells.Add(excelCell);
                excelCell = new List<MExcelCell>();
                excelCell.Add(new MExcelCell() { Content = "單位" });
                excelCell.Add(new MExcelCell() { Content = _Unit });
                excelCells.Add(excelCell);
                excelCell = new List<MExcelCell>();
                excelCell.Add(new MExcelCell() { Content = "販售地區" });
                excelCell.Add(new MExcelCell() { Content = _SalesArea });
                excelCells.Add(excelCell);
                //空一行
                excelCells.Add(new List<MExcelCell>());


                //頁首
                List<string> _HideHeader = new List<string>() { "單號", "時間", "姓名", "單位", "販售地區" };
                excelCell = new List<MExcelCell>();
                foreach (DataGridViewColumn col in view.Columns)
                {
                    //隱藏
                    if (_HideHeader.Contains(col.HeaderText))
                    {
                        continue;
                    }
                    //列印隱藏單價
                    if (col.HeaderText == "單價" && !panel1.Visible)
                    {
                        continue;
                    }
                    excelCell.Add(new MExcelCell()
                    {
                        Content = col.HeaderText
                    });
                }
                //列印顯示價格
                if (panel1.Visible)
                {
                    excelCell.Add(new MExcelCell()
                    {
                        Content = "價格"
                    });
                }
                excelCells.Add(excelCell);

                //內容
                Double _ALLPrice = 0;
                foreach (DataGridViewRow row in view.Rows)
                {
                    Double _unitPrice = 0;
                    Double _count = 1;
                    excelCell = new List<MExcelCell>();
                    foreach (DataGridViewCell cell in row.Cells)
                    {
                        //隱藏
                        if (_HideHeader.Contains(view.Columns[cell.ColumnIndex].HeaderText))
                        {
                            continue;
                        }
                        //列印隱藏單價/保存單價價格
                        if (view.Columns[cell.ColumnIndex].HeaderText == "單價")
                        {
                            if (!panel1.Visible)
                                continue;
                            _unitPrice = Convert.ToDouble(cell.Value);
                        }
                        //保存數量
                        else if (view.Columns[cell.ColumnIndex].HeaderText == "數量")
                        {
                            _count = Convert.ToDouble(cell.Value);
                        }
                        excelCell.Add(new MExcelCell()
                        {
                            Content = cell.Value
                        });
                    }
                    //價格加總
                    if (panel1.Visible)
                    {
                        _ALLPrice = _ALLPrice + Convert.ToDouble(_unitPrice * _count);
                        excelCell.Add(new MExcelCell()
                        {
                            Content = _unitPrice * _count
                        });
                    }
                    excelCells.Add(excelCell);
                }
                //總價
                if (panel1.Visible)
                {
                    //空一行
                    excelCells.Add(new List<MExcelCell>());
                    excelCell = new List<MExcelCell>();
                    for (int i = 0; i < view.Columns.Count; i++)
                    {
                        //隱藏
                        if (_HideHeader.Contains(view.Columns[i].HeaderText))
                        {
                            continue;
                        }
                        excelCell.Add(new MExcelCell());
                    }
                    //只要扣掉一列就好，因為有一列是加總出來的價格，不會在列表裡
                    excelCell.Remove(excelCell[excelCell.Count - 1]);
                    excelCell.Add(new MExcelCell() { Content = "總價" });
                    excelCell.Add(new MExcelCell() { Content = _ALLPrice });
                    excelCells.Add(excelCell);
                }

                //匯出成檔案
                string _Path = $@"{Settings.Excel路徑}{_No}_{comboBox1.Text}.xlsx";
                ePPlus.AddSheet(excelCells, _No);
                ePPlus.Export(_Path);
                ePPlus.ChangeExcel2Image(_Path, @".\ianimage.png");  //利用Spire將excel轉換成圖片

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

                log.LogMessage("確認_列印 成功路徑：" + _Path, enumLogType.Trace);
                log.LogMessage("確認_列印 成功", enumLogType.Info);
            }
            catch (Exception ee)
            {
                MessageBox.Show("列印 失敗：\r\n" + ee.Message);
                log.LogMessage("確認_列印 失敗：\r\n" + ee.Message, enumLogType.Error);
                button1.Enabled = true;
                return;
            }
            #endregion

            comboBox1.SelectedIndex = -1;
            label5.Text = _No;
        }
        int _Page = 1;
        int _PageHeight = 0;
        private void PrintDocument_PrintPage(object sender, PrintPageEventArgs e)
        {
            #region 
            GC.Collect();
            //e.HasMorePages = true; //此处打开多页打印属性
            //imagepath是指 excel轉成的圖片的路徑
            using (Bitmap bitmap = new Bitmap(@".\ianimage.png"))
            {
                int _Y = 0; //輸出圖片時要從Y軸哪個點開始
                int _1PageSize = 1098;
                int _2PageSize = 1085;
                Rectangle newarea = new Rectangle();
                newarea.X = 0;
                newarea.Y = 0;
                newarea.Width = bitmap.Width;
                newarea.Height = bitmap.Height;
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
                

                int _width = newarea.Width;
                if (newarea.Width > 810)
                {
                    _width = 810;
                    newarea.Width = 810;
                }
                e.Graphics.DrawImage(bitmap, newarea, 0, _Y, _width, newarea.Height, GraphicsUnit.Pixel);
                _Page++;
                if (!e.HasMorePages)
                    button1.Enabled = true;
            }
            #endregion
        }

        private void UserView_KeyDown(object sender, KeyEventArgs e)
        {
            textBox1.Focus();
        }

        string _comboBoxSelectText = "";
        Boolean _comboBoxKeyPressSet = false;
        private void timer_ComboBoxSelect_Tick(object sender, EventArgs e)
        {
            textBox1.Text = _comboBoxSelectText = "";
            timer_ComboBoxSelect.Stop();
            panel2.BackColor = SystemColors.Control;
            textBox1.Focus();
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
                    timer_ComboBoxSelect.Stop();
                    timer_ComboBoxSelect_Tick(sender, e);
                    return;
                }
                int _key = -1;
                if (!timer_ComboBoxSelect.Enabled)
                {
                    timer_ComboBoxSelect.Start();
                }
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
    }
}
