using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.Entity.Core.Common.CommandTrees.ExpressionBuilder;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;
using 簡易倉儲系統.DB;
using 簡易倉儲系統.EssentialTool;
using static 簡易倉儲系統.EssentialTool.LogToText;

namespace 簡易倉儲系統
{
    public partial class UserView : Form
    {
        LogToText log = new LogToText(@".\Log");

        public static string Setting_Path = @".\";          //設定檔路徑
        public static string VersionNumber = "";            //程式版號
        public static string DB_Path = "";                  //資料庫路徑
        public static string TableName = "";                //資料表名稱
        public static string type = "";                     //類型
        public static string salesArea = "";                //販售地點
        public static string unit = "";                     //重量單位
        public static string unitPrice = "";                //單價

        public UserView()
        {
            InitializeComponent();
        }

        private void UserView_Load(object sender, EventArgs e)
        {
            Settings.StartUp(Setting_Path);
            VersionNumber = Application.ProductVersion;
            this.Text += $"v.{VersionNumber} Bulid{File.GetLastWriteTime(Application.ExecutablePath).ToString("yyyyMMdd")}";

            //檢查時間為最新
            try
            {
                log.LogMessage("檢查時間 開始", enumLogType.Info);

                if (@String.IsNullOrEmpty(Settings.每日檢查))
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
            }
            catch (Exception ee)
            {
                log.LogMessage("檢查時間 失敗" + ee.Message, enumLogType.Error);
                MessageBox.Show("檢查時間 失敗");
                Application.Exit();
                return;
            }

            //檢查程式是否符合效期內
            try
            {
                log.LogMessage("檢查序號 開始", enumLogType.Info);
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
            }
            catch (Exception ee)
            {
                log.LogMessage("序號啟動 失敗" + ee.Message, enumLogType.Error);
                MessageBox.Show("序號啟動 失敗");
                Application.Exit();
                return;
            }

            //檢查程式是否有重複開啟
            Process[] proc = Process.GetProcessesByName(Process.GetCurrentProcess().ProcessName);
            if (proc.Length > 1)
            {
                //表示此程式已被開啟
                Application.Exit();
                return;
            }
            log.LogMessage("系統啓動", enumLogType.Trace);
            log.LogMessage("使用者介面啓動", enumLogType.Info);

            try
            {
                // 讀取資料
                //DB_Path = Properties.Settings.Default.資料庫路徑 + @"/data.db";
                DB_Path = Settings.資料庫路徑 + @"data.db";
                DB_SQLite dB_SQLite = new DB_SQLite();
                DB_SQLite.DatatableToDatagridview(dB_SQLite.GetDataTable(DB_Path, $@"SELECT * FROM SalesRecord WHERE Date > '{DateTime.Now.ToString("yyyy-MM-dd")}';"), dataGridView1);
                log.LogMessage("讀取資料庫 成功。", enumLogType.Trace);
            }
            catch (Exception ee)
            {
                log.LogMessage("連線資料庫 失敗\r\n" + ee.Message, enumLogType.Error);
                MessageBox.Show("連線資料庫 失敗\r\n" + ee.Message);
                return;
            }
        }

        //類型設定，單價設定
        private void radioButton_type_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                if (((RadioButton)sender).Checked)
                {
                    log.LogMessage("類型設定，單價設定 開始", enumLogType.Trace);

                    type = ((ButtonBase)sender).Text;
                    ((GroupBox)((RadioButton)sender).Parent).BackColor = SystemColors.Control;

                    if (!string.IsNullOrEmpty(TableName))
                    {
                        DB_SQLite dB_SQLite = new DB_SQLite();
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
                            log.LogMessage("管理者尚未輸入當日價格", enumLogType.Info);
                            MessageBox.Show("管理者尚未輸入當日價格");
                            return;
                        }
                        unitPrice = DT.Rows[0][Int32.Parse(((RadioButton)sender).Tag.ToString())].ToString();
                        label3.Text = unitPrice;
                        panel1.BackColor = SystemColors.Control;
                    }

                    for (int i = 0; i < ((RadioButton)sender).Parent.Controls.Count; i++)
                    {
                        if (((RadioButton)((RadioButton)sender).Parent.Controls[i]).Checked)
                            ((RadioButton)((RadioButton)sender).Parent.Controls[i]).BackColor = Color.GreenYellow;
                        else
                            ((RadioButton)((RadioButton)sender).Parent.Controls[i]).BackColor = SystemColors.Control;
                    }

                    log.LogMessage("類型設定，單價設定 成功", enumLogType.Trace);
                }
            }
            catch (Exception ee)
            {
                log.LogMessage("類型設定，單價設定 失敗：\r\n" + ee.Message, enumLogType.Error);
            }
        }

        //販售地點設定，資料表名稱設定，重量單位設定，針對不同地區客製化功能
        private void radioButton_salesArea_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                log.LogMessage("販售地點設定，資料表名稱設定，重量單位設定，針對不同地區客製化功能 開始", enumLogType.Trace);

                string _Text = ((ButtonBase)sender).Text;
                if (((RadioButton)sender).Checked)
                {
                    switch (Int32.Parse(((RadioButton)sender).Tag.ToString()) - 1)
                    {
                        case 0: //外銷韓國
                            TableName = "ExportKoreaUnitPrice";
                            break;
                        case 1: //外銷日本 
                            TableName = "ExportJapanUnitPrice";
                            break;
                        case 2: //超市 
                            TableName = "ExportSupermarketUnitPrice";
                            break;
                        default:
                            return;
                    }

                    salesArea = _Text;
                    if (_Text.Contains("公斤"))
                        unit = "公斤";
                    else if (_Text.Contains("台斤"))
                        unit = "台斤";

                    label1.Text = unit;
                    ((GroupBox)((RadioButton)sender).Parent).BackColor = SystemColors.Control;
                    for (int i = 0; i < ((RadioButton)sender).Parent.Controls.Count; i++)
                    {
                        if (((RadioButton)((RadioButton)sender).Parent.Controls[i]).Checked)
                            ((RadioButton)((RadioButton)sender).Parent.Controls[i]).BackColor = Color.GreenYellow;
                        else
                            ((RadioButton)((RadioButton)sender).Parent.Controls[i]).BackColor = SystemColors.Control;
                    }
                }
                for (int i = 0; i < groupBox1.Controls.Count; i++)
                {
                    //針對不同地區客製化功能
                    if (_Text.Contains("外銷韓國") || _Text.Contains("外銷日本"))
                    {
                        radioButton3.Text = "14粒";
                        radioButton4.Text = "16粒";
                        radioButton7.Enabled = true;
                        radioButton7.Visible = true;
                    }
                    else if (_Text.Contains("超市"))
                    {
                        radioButton3.Text = "15粒";
                        radioButton4.Text = "18粒";
                        radioButton7.Enabled = false;
                        radioButton7.Visible = false;
                    }

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
                panel1.BackColor = Color.IndianRed;
                _OK = false;
            }

            if (_OK && e.KeyChar == ((char)Keys.Enter))
            {
                log.LogMessage("確認 開始", enumLogType.Trace);

                DataGridViewRow row = new DataGridViewRow();
                DateTime now = DateTime.Now;

                #region DB新增

                // 插入資料
                DB_SQLite dB_SQLite = new DB_SQLite();
                string insertstring = $@"INSERT INTO SalesRecord (No, Date, Type, Count, UnitPrice, Unit, salesArea) 
                         SELECT CASE WHEN MAX(No) ISNULL THEN '{now.ToString("yyyyMMdd") + "001"}' ELSE MAX(No)+1 END No, 
                        '{now.ToString("yyyy-MM-dd HH:mm:ss")}' Date , '{type}' Type, '{((TextBox)sender).Text}' Count, 
                        '{unitPrice}' UnitPrice, '{unit}' Unit, '{salesArea}' salesArea 
                        FROM SalesRecord WHERE Date > '{now.ToString("yyyy-MM-dd")}';";
                dB_SQLite.Manipulate(DB_Path, insertstring);


                log.LogMessage("確認_DB新增 成功路徑：" + DB_Path + "\r\n語法：" + insertstring, enumLogType.Trace);
                log.LogMessage("確認_DB新增 成功", enumLogType.Info);
                #endregion

                #region DataGridView修改

                //重繪才能讀取到別的使用者登錄的資料
                DB_SQLite.DatatableToDatagridview(dB_SQLite.GetDataTable(DB_Path, $@"SELECT * FROM SalesRecord WHERE Date > '{DateTime.Now.ToString("yyyy-MM-dd")}';"), dataGridView1);
                //row.CreateCells(dataGridView1);
                //row.SetValues(new string[] { now.ToString("yyyyMMdd") + "001", now.ToString("yyyy-MM-dd HH:mm:ss"), type, ((TextBox)sender).Text, unitPrice, unit, salesArea });
                //dataGridView1.Rows.Insert(0, row);
                //dataGridView1.Rows[0].Selected = true;
                //dataGridView1.CurrentCell = dataGridView1.Rows[0].Cells[0];

                textBox1.Text = "";
                textBox1.Focus();

                log.LogMessage("確認_DB新增 成功", enumLogType.Info);
                #endregion

                log.LogMessage("確認後寫入DB 成功", enumLogType.Trace);
            }
        }

        private void dataGridView1_UserDeletedRow(object sender, DataGridViewRowEventArgs e)
        {
            log.LogMessage("刪除失敗則還原 開始", enumLogType.Trace);

            //刪除失敗則還原
            DB_SQLite dB_SQLite = new DB_SQLite();
            try
            {
                dB_SQLite.Manipulate(DB_Path, $@"DELETE FROM SalesRecord WHERE No = '{e.Row.Cells[0].Value.ToString()}';");
                log.LogMessage("刪除失敗則還原 成功", enumLogType.Trace);
            }
            catch (Exception ee)
            {
                DB_SQLite.DatatableToDatagridview(dB_SQLite.GetDataTable(DB_Path, $@"SELECT * FROM SalesRecord WHERE Date > '{DateTime.Now.ToString("yyyy-MM-dd")}';"), dataGridView1);
                log.LogMessage("刪除失敗則還原 失敗：" + ee.Message, enumLogType.Trace);
                MessageBox.Show("刪除失敗：" + ee.Message);
            }
        }
    }
}
