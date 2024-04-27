using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Logical;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
using Spire.Pdf.Exporting.XPS.Schema;
using Spire.Pdf.Graphics;
using Spire.Xls.Core;
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
using System.Reflection.Emit;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;
using 簡易倉儲系統.DB;
using 簡易倉儲系統.EssentialTool;
using 簡易倉儲系統.EssentialTool.Excel;
using static System.Windows.Forms.AxHost;
using static 簡易倉儲系統.EssentialTool.LogToText;
using Label = System.Windows.Forms.Label;

namespace 簡易倉儲系統
{
    public partial class ManagerView : Form
    {
        LogToText log = new LogToText(@".\Log");
        DB_SQLite dB_SQLite = new DB_SQLite();

        public static DataTable _SelectDT = new DataTable();
        public static string _SelectType = "";
        public static List<Control> _SelectControl;
        public static Size _dataGridView4Size;
        public static Point _dataGridView4Point;
        public static string IUDCustomerProfile = "";
        public static string Inquire = "";
        public static string Setting_Path = @".\";          //設定檔路徑
        public static string VersionNumber = "";            //程式版號
        public static string DB_Path = "";    //DB路徑
        public static string[][] type = { new string[] { "", "", "", "", "", "", "", "", "", "" }
                                        , new string[] { "", "", "", "", "", "", "", "", "", "" }
                                        , new string[] { "", "", "", "", "", "", "", "", "", "" } };

        public ManagerView()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            Settings.StartUp(Setting_Path);
            VersionNumber = Application.ProductVersion;
            this.Text += $"v.{VersionNumber} Bulid{File.GetLastWriteTime(Application.ExecutablePath).ToString("yyyyMMdd")}";
            textBox21.Text = "";
            label23.Text = "";
            label25.Text = "";
            label28.Text = "";
            checkedListBox1.Items.Clear();
            checkedListBox2.Items.Clear();
            DB_Path = Settings.資料庫路徑 + @"data.db";

            if (!string.IsNullOrEmpty(Settings.SQL語法))
            {
                try
                {
                    log.LogMessage("執行SQL語法 開始", enumLogType.Trace);
                    dB_SQLite.Manipulate(DB_Path, $@"{Settings.SQL語法}");
                    Settings.SQL語法 = "";
                    log.LogMessage($@"執行SQL語法 成功：{Settings.SQL語法}", enumLogType.Info);
                    log.LogMessage($@"執行SQL語法 成功：{Settings.SQL語法}", enumLogType.Trace);
                    MessageBox.Show($@"執行SQL語法 成功：{Settings.SQL語法}");
                }
                catch (Exception ee)
                {
                    log.LogMessage("執行SQL語法 失敗" + ee.Message, enumLogType.Error);
                    MessageBox.Show("執行SQL語法 失敗" + ee.Message);
                }
                Application.Exit();
            }

            //各種偵測，每24小時偵測一次
            timer_detection_Tick(sender, e);

            #region 取得類型設定參數
            try
            {
                log.LogMessage("取得類型設定參數 開始", enumLogType.Trace);

                string _Type = "";
                string LogMessage = "";
                List<Control[]> _LabelList = new List<Control[]>();

                tabPage1.Text = Settings.販售地區1.Split('/')[0];
                _Type = Settings.類型1;
                _LabelList = new List<Control[]>() {new Control[] { panel1, label1 }
                    , new Control[] { panel2, label2 }, new Control[] { panel3, label3 }
                    , new Control[] { panel4, label4 }, new Control[] { panel5, label5 }
                    , new Control[] { panel6, label6 }, new Control[] { panel7, label7 }
                    , new Control[] { panel22, label27 }, new Control[] { panel29, label35 }
                    , new Control[] { panel31, label37 } };
                if (ConvectTypeText(_Type, _LabelList, dataGridView1))
                    LogMessage += tabPage1.Text + $@"：{_Type}";

                tabPage2.Text = Settings.販售地區2.Split('/')[0];
                _Type = Settings.類型2;
                _LabelList = new List<Control[]>() {new Control[] { panel8, label8 }
                    , new Control[] { panel9, label9 }, new Control[] { panel10, label10 }
                    , new Control[] { panel11, label11 }, new Control[] { panel12, label12 }
                    , new Control[] { panel13, label13 }, new Control[] { panel14, label14 }
                    , new Control[] { panel23, label29 }, new Control[] { panel28, label34 }
                    , new Control[] { panel30, label36 } };
                if (ConvectTypeText(_Type, _LabelList, dataGridView2))
                    LogMessage += " / " + tabPage2.Text + $@"：{_Type}";

                tabPage3.Text = Settings.販售地區3.Split('/')[0];
                _Type = Settings.類型3;
                _LabelList = new List<Control[]>() {new Control[] { panel15, label15 }
                    , new Control[] { panel16, label16 }, new Control[] { panel17, label17 }
                    , new Control[] { panel18, label18 }, new Control[] { panel19, label19 }
                    , new Control[] { panel20, label20 }, new Control[] { panel24, label30 }
                    , new Control[] { panel25, label31 }, new Control[] { panel26, label32 }
                    , new Control[] { panel27, label33 } };
                if (ConvectTypeText(_Type, _LabelList, dataGridView3))
                    LogMessage += " / " + tabPage3.Text + $@"：{_Type}";

                log.LogMessage("取得類型設定參數 成功\r\n" + tabPage1.Text + " / " +
                    tabPage2.Text + " / " + tabPage3.Text, enumLogType.Info);
                log.LogMessage("取得類型設定參數 成功\r\n" + LogMessage, enumLogType.Trace);
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
                if (!File.Exists(DB_Path))
                {
                    log.LogMessage("偵測到無資料庫，準備開始建立。", enumLogType.Debug);
                    var createtablestring = "";

                    // 建立 SQLite 資料庫
                    dB_SQLite.CreateDatabase(DB_Path);

                    // 建立資料表 設定客戶資料 CustomerProfile
                    createtablestring = @"CREATE TABLE CustomerProfile (ID Integer NOT NULL, CustomerID TEXT, CustomerName TEXT, PRIMARY KEY(ID AUTOINCREMENT));";
                    dB_SQLite.CreateTable(DB_Path, createtablestring);

                    // 建立資料表 設定參數值 Setting
                    createtablestring = @"CREATE TABLE Setting (SettingName TEXT, SettingValue TEXT);";
                    dB_SQLite.CreateTable(DB_Path, createtablestring);
                    dB_SQLite.Manipulate(DB_Path, $@"
                        INSERT INTO Setting (SettingName, SettingValue) VALUES ('ShowMoney_ExportKorea', 'True');
                        INSERT INTO Setting (SettingName, SettingValue) VALUES ('ShowMoney_ExportJapan', 'True');
                        INSERT INTO Setting (SettingName, SettingValue) VALUES ('ShowMoney_ExportSupermarket', 'True');
                    ");

                    // 建立資料表 販售紀錄 SalesRecord
                    createtablestring = @"CREATE TABLE SalesRecord (No Integer, Time DateTime, Name TEXT, Type TEXT, Count double
                    , UnitPrice double, Unit TEXT, SalesArea TEXT, Paid Integer, PaidTime DateTime);";
                    dB_SQLite.CreateTable(DB_Path, createtablestring);

                    // 建立資料表 外銷韓國 ExportKoreaUnitPrice
                    createtablestring = @"CREATE TABLE ExportKoreaUnitPrice (Date DateTime, Type1 double, Type2 double
                    , Type3 double, Type4 double, Type5 double, Type6 double, Type7 double, Type8 double, Type9 double
                    , Type10 double);";
                    dB_SQLite.CreateTable(DB_Path, createtablestring);

                    // 建立資料表 外銷日本 ExportJapanUnitPrice
                    createtablestring = @"CREATE TABLE ExportJapanUnitPrice (Date DateTime, Type1 double, Type2 double
                    , Type3 double, Type4 double, Type5 double, Type6 double, Type7 double, Type8 double, Type9 double
                    , Type10 double);";
                    dB_SQLite.CreateTable(DB_Path, createtablestring);

                    // 建立資料表 超市 ExportSupermarketUnitPrice
                    createtablestring = @"CREATE TABLE ExportSupermarketUnitPrice (Date DateTime, Type1 double, Type2 double
                    , Type3 double, Type4 double, Type5 double, Type6 double, Type7 double, Type8 double, Type9 double
                    , Type10 double);";
                    dB_SQLite.CreateTable(DB_Path, createtablestring);

                    log.LogMessage("建立資料庫 成功。", enumLogType.Debug);
                }

                // 讀取資料
                foreach (DataRow item in dB_SQLite.GetDataTable(DB_Path, @"SELECT * FROM Setting").Rows)
                {
                    switch (item[0].ToString())
                    {
                        case "ShowMoney_ExportKorea":
                            checkBox1.Checked = Boolean.Parse(item[1].ToString());
                            break;
                        case "ShowMoney_ExportJapan":
                            checkBox2.Checked = Boolean.Parse(item[1].ToString());
                            break;
                        case "ShowMoney_ExportSupermarket":
                            checkBox3.Checked = Boolean.Parse(item[1].ToString());
                            break;
                    }
                }
                DatatableToDatagridview(dB_SQLite.GetDataTable(DB_Path, @"SELECT * FROM ExportKoreaUnitPrice"), dataGridView1);
                DatatableToDatagridview(dB_SQLite.GetDataTable(DB_Path, @"SELECT * FROM ExportJapanUnitPrice"), dataGridView2);
                DatatableToDatagridview(dB_SQLite.GetDataTable(DB_Path, @"SELECT * FROM ExportSupermarketUnitPrice"), dataGridView3);
                DB_SQLite.DatatableToDatagridview(dB_SQLite.GetDataTable(DB_Path, "SELECT * FROM CustomerProfile"), dataGridView5);
                log.LogMessage("讀取資料庫 成功。", enumLogType.Trace);
            }
            catch (Exception ee)
            {
                log.LogMessage("連線資料庫 失敗\r\n" + ee.Message, enumLogType.Error);
                MessageBox.Show("連線資料庫 失敗\r\n" + ee.Message);
                return;
            }
            log.LogMessage("管理者介面啓動", enumLogType.Info);
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
                    log.LogMessage("請於設定檔內輸入序號：" + Setting_Path + @"\Setting.xml", enumLogType.Info);
                    MessageBox.Show("請於設定檔內輸入序號：" + Setting_Path + @"\Setting.xml");
                    Application.Exit();
                    return;
                }
                string _ianNo = EncryptionDecryption.desEncryptBase64("ian/2023-04-17/2023-04-28/ian");
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

        private bool ConvectTypeText(string TypeText, List<Control[]> labelList, DataGridView dataGridView)
        {
            int _TypeCount = TypeText.Split('/').Count();
            #region 10個類型的位置調整
            for (int i = 0; i < 10; i++)
            {
                if (_TypeCount > i)
                {
                    ((Panel)labelList[i][0]).Enabled = true;
                    ((Panel)labelList[i][0]).Visible = true;
                    ((Label)labelList[i][1]).Text = TypeText.Split('/')[i];
                    dataGridView.Columns[i + 1].HeaderText = TypeText.Split('/')[i];
                    dataGridView.Columns[i + 1].Visible = true;
                }
                else
                {
                    ((Panel)labelList[i][0]).Enabled = false;
                    ((Panel)labelList[i][0]).Visible = false;
                    dataGridView.Columns[i + 1].Visible = false;
                }
            }
            #endregion
            return true;
        }

        /// <summary>
        /// Datatable轉出Datagridview
        /// 全部刪除重繪
        /// </summary>
        /// <param name="DT"></param>
        /// <param name="DGV"></param>
        /// <returns></returns>
        private Boolean DatatableToDatagridview(DataTable DT, DataGridView DGV)
        {
            try
            {
                log.LogMessage("Datatable轉出Datagridview 開始", enumLogType.Trace);
                DGV.Rows.Clear();
                for (int i = 0; i < DT.Rows.Count; i++)
                {
                    DataGridViewRow _data = new DataGridViewRow();
                    _data.CreateCells(DGV);
                    List<string> strings = new List<string>();
                    for (int j = 0; j < DT.Columns.Count; j++)
                    {
                        string _value = DT.Rows[i][j].ToString();
                        if (DT.Columns[j].ColumnName == "Date")
                            _value = DateTime.Parse(_value).ToString("yyyy-MM-dd");
                        if (DT.Columns[j].ColumnName == "Unpaid")
                            _value = ((int)Math.Round(Convert.ToDouble(_value), 0, MidpointRounding.AwayFromZero)).ToString();
                        strings.Insert(j, _value);
                    }
                    _data.SetValues(strings.ToArray());
                    DGV.Rows.Insert(0, _data);
                    DGV.Rows[0].Selected = true;
                    DGV.CurrentCell = DGV.Rows[0].Cells[0];
                }
                log.LogMessage("Datatable轉出Datagridview 成功", enumLogType.Trace);
                return true;
            }
            catch (Exception ee)
            {
                log.LogMessage("Datatable轉出Datagridview 失敗：" + ee.Message, enumLogType.Error);
                MessageBox.Show("Datatable轉出Datagridview 失敗：" + ee.Message);
                return false;
            }
        }

        //換分頁時清空資料
        private void tabControl_Click(object sender, EventArgs e)
        {
            log.LogMessage("換分頁時清空資料 開始", enumLogType.Trace);

            //清空搜尋頁
            dataGridView4.Rows.Clear();
            checkedListBox1.Items.Clear();
            checkedListBox2.Items.Clear();
            _SelectDT = new DataTable();

            //清空暫存
            type = new string[][] { new string[] { "", "", "", "", "", "", "", "", "", "" }
                                  , new string[] { "", "", "", "", "", "", "", "", "", "" }
                                  , new string[] { "", "", "", "", "", "", "", "", "", "" } };

            //要清空的TextBox元件
            System.Windows.Forms.TextBox[] _textBoxes = { textBox1, textBox2, textBox3, textBox4, textBox5
                    , textBox6, textBox7, textBox8, textBox9, textBox10, textBox11, textBox12, textBox13
                    , textBox14, textBox15, textBox16, textBox17, textBox18, textBox19, textBox20, textBox21
                    , textBox22, textBox23, textBox24, textBox25, textBox26, textBox30, textBox32, textBox29
                    , textBox31, textBox27, textBox28};
            foreach (var _textBox in _textBoxes)
            {
                _textBox.Text = "";
            }

            log.LogMessage("換分頁時清空資料 成功", enumLogType.Trace);
        }

        //文字變更時，判斷顏色與暫存資料
        private void textBox_TextChanged(object sender, EventArgs e)
        {
            log.LogMessage("文字變更時，判斷顏色與暫存資料 開始", enumLogType.Trace);

            Double _UnitPrice = 0;
            if (((System.Windows.Forms.TextBox)sender).Text == "")
            {
                //int _Index = int.Parse(((TabControl)((Control)sender).Parent.Parent.Parent.Parent).SelectedIndex.ToString());
                int _Index = int.Parse(((Control)sender).Parent.Parent.Parent.Name.Replace("tabPage", "")) - 1;
                int _No = int.Parse(((Control)sender).Parent.Tag.ToString());
                type[_Index][_No - 1] = "";
                ((Control)sender).Parent.BackColor = Color.Transparent;

                log.LogMessage("文字變更為空 成功", enumLogType.Trace);
            }
            else if (!Double.TryParse(((System.Windows.Forms.TextBox)sender).Text, out _UnitPrice))
            {
                ((Control)sender).Parent.BackColor = Color.IndianRed;

                log.LogMessage("文字變更非數字 失敗", enumLogType.Error);
            }
            else
            {
                //int _Index = int.Parse(((TabControl)((Control)sender).Parent.Parent.Parent.Parent).SelectedIndex.ToString());
                int _Index = int.Parse(((Control)sender).Parent.Parent.Parent.Name.Replace("tabPage", "")) - 1;
                int _No = int.Parse(((Control)sender).Parent.Tag.ToString());
                type[_Index][_No - 1] = _UnitPrice.ToString();
                ((Control)sender).Parent.BackColor = Color.GreenYellow;

                log.LogMessage($@"文字變更{_UnitPrice.ToString()}： 成功", enumLogType.Trace);
            }
        }

        //新增/修改
        private void button_Click(object sender, EventArgs e)
        {
            try
            {
                log.LogMessage("新增/修改 開始", enumLogType.Info);

                DateTime now = DateTime.Now;
                string _state = "I";
                DataGridViewRow _data = new DataGridViewRow();
                DataGridView _view = new DataGridView();

                foreach (var item in ((Control)sender).Parent.Parent.Controls)
                {
                    if (item.GetType().Name == "DataGridView")
                    {
                        _view = (DataGridView)item;
                        break;
                    }
                }

                foreach (DataGridViewRow _row in _view.Rows)
                {
                    if (_row.Cells[0].Value != null && _row.Cells[0].Value.ToString().Contains(now.ToString("yyyy-MM-dd")))
                    {
                        //修改
                        _state = "U";
                        _data = _row;
                        break;
                    }
                    else
                    {
                        //新增
                        _state = "I";
                        break;
                    }
                }

                int _Index = int.Parse(((Control)sender).Parent.Parent.Name.Replace("tabPage", "")) - 1;
                List<string> _typeList = type[_Index].ToList();
                _typeList.Insert(0, now.ToString("yyyy-MM-dd"));

                //DB前置設定
                string _TableName = "";
                switch (_Index)
                {
                    case 0: //外銷韓國
                        _TableName = "ExportKoreaUnitPrice";
                        break;
                    case 1: //外銷日本 
                        _TableName = "ExportJapanUnitPrice";
                        break;
                    case 2: //超市 
                        _TableName = "ExportSupermarketUnitPrice";
                        break;
                    default:
                        return;
                }

                if (_state == "U")
                {
                    #region DB修改

                    // 插入資料
                    string insertstring = $@"UPDATE {_TableName} SET ";
                    for (int i = 0; i < _typeList.Count; i++)
                    {
                        if (i == 0)
                            insertstring += $@"Date = '{_typeList[i]}'";
                        else
                            insertstring += $@", Type{i} = '{_typeList[i]}'";
                    }
                    insertstring += $@" WHERE Date = '{now.ToString("yyyy-MM-dd")}';";
                    dB_SQLite.Manipulate(DB_Path, insertstring);

                    log.LogMessage("新增/修改_DB修改 成功路徑：" + DB_Path + "\r\n語法：" + insertstring, enumLogType.Trace);
                    log.LogMessage("新增/修改_DB修改 成功路徑：" + DB_Path + "\r\n語法：" + insertstring, enumLogType.Info);
                    #endregion

                    #region DataGridView修改

                    //修改
                    _data.SetValues(_typeList.ToArray());

                    log.LogMessage("新增/修改_DataGridView修改 成功資料：[" + string.Join(", ", _typeList.ToArray()) + "]", enumLogType.Trace);
                    log.LogMessage("新增/修改_DataGridView修改 成功資料：[" + string.Join(", ", _typeList.ToArray()) + "]", enumLogType.Info);
                    #endregion
                }
                else if (_state == "I")
                {
                    #region DB新增

                    // 插入資料
                    string insertstring = $@"INSERT INTO {_TableName} (";
                    for (int i = 0; i < _typeList.Count; i++)
                    {
                        if (i == 0)
                            insertstring += "Date";
                        else
                            insertstring += $@", Type{i}";
                    }
                    insertstring += $@") VALUES ('{String.Join("','", _typeList)}');";
                    dB_SQLite.Manipulate(DB_Path, insertstring);

                    log.LogMessage("新增/修改_DB新增 成功路徑：" + DB_Path + "\r\n語法：" + insertstring, enumLogType.Trace);
                    log.LogMessage("新增/修改_DB新增 成功路徑：" + DB_Path + "\r\n語法：" + insertstring, enumLogType.Info);
                    #endregion
                    //// 讀取資料_20230313_Ian_先採用各別新增加入的方式，防止未來資料過大需要重繪DataGridView時，造成吃掉大量記憶體
                    //var dataTable = dB_SQLite.GetDataTable(DB_Path, $@"SELECT * FROM {_TableName}");


                    #region DataGridView新增
                    log.LogMessage("新增/修改_DataGridView新增 開始", enumLogType.Info);

                    //新增
                    _data.CreateCells(_view);
                    _data.SetValues(_typeList.ToArray());

                    _view.Rows.Insert(0, _data);
                    _view.Rows[0].Selected = true;
                    _view.CurrentCell = _view.Rows[0].Cells[0];
                    ((Control)sender).Parent.Controls[0].Controls[1].Focus();

                    log.LogMessage("新增/修改_DataGridView新增 成功資料：[" + string.Join(", ", _typeList.ToArray()) + "]", enumLogType.Trace);
                    log.LogMessage("新增/修改_DataGridView新增 成功資料：[" + string.Join(", ", _typeList.ToArray()) + "]", enumLogType.Info);
                    #endregion

                }
            }
            catch (Exception ee)
            {
                log.LogMessage("新增修改 失敗：" + ee.Message, enumLogType.Error);
                MessageBox.Show("新增修改 失敗：" + ee.Message);
            }
        }

        //dataGridView轉出單價
        private void dataGridView_Click(object sender, EventArgs e)
        {
            try
            {
                log.LogMessage("dataGridView轉出 開始", enumLogType.Trace);

                Control control = new Control();
                //取得目前頁面上的 單價設定 所有欄位 
                foreach (Control cont in tabControl1.SelectedTab.Controls)
                {
                    if (cont.Text == "單價設定")
                    {
                        control = cont;
                        break;
                    }
                }

                if (((DataGridView)sender).Rows.Count <= 0)
                    return;
                if (((DataGridView)sender).Rows[((DataGridView)sender).CurrentRow.Index].Cells[0].Value == null)
                    return;
                for (int i = 1; i < ((DataGridView)sender).Rows[((DataGridView)sender).CurrentRow.Index].Cells.Count; i++)
                {
                    if (((DataGridView)sender).Rows[((DataGridView)sender).CurrentRow.Index].Cells[i].Value != null)
                    {
                        control.Controls[i - 1].Controls[1].Text =
                            ((DataGridView)sender).Rows[((DataGridView)sender).CurrentRow.Index].Cells[i].Value.ToString();
                    }
                }

                log.LogMessage("dataGridView轉出 成功", enumLogType.Trace);
            }
            catch (Exception ee)
            {
                log.LogMessage("dataGridView轉出 失敗：" + ee.Message, enumLogType.Error);
                MessageBox.Show("dataGridView轉出 失敗：" + ee.Message);
            }
        }

        //顯示金額
        private void checkBox_CheckedChanged(object sender, EventArgs e)
        {
            string _SettingName = ((System.Windows.Forms.Control)sender).Tag.ToString();
            string _SettingValue = ((System.Windows.Forms.CheckBox)sender).Checked.ToString();
            dB_SQLite.Manipulate(DB_Path, $@"UPDATE Setting SET SettingValue = '{_SettingValue}' WHERE SettingName = '{_SettingName}';");
        }

        //單號或姓名查詢
        private void radioButton_查詢_CheckedChanged(object sender, EventArgs e)
        {
            string _Text = ((ButtonBase)sender).Text;
            if (((RadioButton)sender).Checked)
            {
                groupBox8.Text = _Text.Replace("查詢", "");
                if (_Text == "單號查詢")
                {
                    Inquire = "單號";
                    checkBox4.Checked = false;
                    checkBox4.Enabled = false;
                    checkBox4.Visible = false;
                    panel21.Enabled = false;
                    panel21.Visible = false;
                    button7.Enabled = true;
                    button7.Visible = true;
                    button8.Enabled = true;
                    button8.Visible = true;
                    this.Column10.HeaderText = "時間";
                    //if (!dataGridView4.Columns[3].Visible)
                        dataGridView4.Rows.Clear();
                    checkedListBox1.Items.Clear();
                    checkedListBox2.Items.Clear();
                    dataGridView4.Columns[3].Visible = true;
                    dataGridView4.Columns[4].Visible = true;
                    dataGridView4.Columns[5].Visible = true;
                    dataGridView4.Columns[6].Visible = true;
                    dataGridView4.Columns[7].Visible = true;
                    dataGridView4.Columns[10].Visible = true;
                }
                else if (_Text == "姓名查詢")
                {
                    Inquire = "姓名";
                    checkBox4.Checked = false;
                    checkBox4.Enabled = true;
                    checkBox4.Visible = true;
                    button7.Enabled = false;
                    button7.Visible = false;
                    button8.Enabled = false;
                    button8.Visible = false;
                    this.Column10.HeaderText = "時間";
                    //if (!dataGridView4.Columns[3].Visible)
                        dataGridView4.Rows.Clear();
                    checkedListBox1.Items.Clear();
                    checkedListBox2.Items.Clear();
                    dataGridView4.Columns[3].Visible = true;
                    dataGridView4.Columns[4].Visible = true;
                    dataGridView4.Columns[5].Visible = true;
                    dataGridView4.Columns[6].Visible = true;
                    dataGridView4.Columns[7].Visible = true;
                    dataGridView4.Columns[10].Visible = true;
                }
                else if (_Text == "整合查詢")
                {
                    Inquire = "整合";
                    checkBox4.Checked = true;
                    checkBox4.Enabled = true;
                    checkBox4.Visible = true;
                    button7.Enabled = true;
                    button7.Visible = true;
                    button8.Enabled = true;
                    button8.Visible = true;
                    this.Column10.HeaderText = "日期";
                    dataGridView4.Rows.Clear();
                    checkedListBox1.Items.Clear();
                    checkedListBox2.Items.Clear();
                    label28.Text = "";
                    label25.Text = "";
                    label23.Text = "";
                    dataGridView4.Columns[3].Visible = false;
                    dataGridView4.Columns[4].Visible = false;
                    dataGridView4.Columns[5].Visible = false;
                    dataGridView4.Columns[6].Visible = false;
                    dataGridView4.Columns[7].Visible = true;
                    dataGridView4.Columns[10].Visible = false;
                }

                ((GroupBox)((RadioButton)sender).Parent).BackColor = Color.Transparent;
                for (int i = 0; i < ((RadioButton)sender).Parent.Controls.Count; i++)
                {
                    if (((RadioButton)((RadioButton)sender).Parent.Controls[i]).Checked)
                        ((RadioButton)((RadioButton)sender).Parent.Controls[i]).BackColor = Color.GreenYellow;
                    else
                        ((RadioButton)((RadioButton)sender).Parent.Controls[i]).BackColor = Color.Transparent;
                }
                textBox21.Text = "";
            }
        }

        //確認搜尋
        private void textBox21_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                if (Inquire == "單號" && textBox21.Text == "")
                {
                    return;
                }
                if (Inquire == "姓名" && textBox21.Text == "" && !checkBox4.Checked)
                {
                    return;
                }
                if (Inquire == "整合" && textBox21.Text == "" && !checkBox4.Checked)
                {
                    return;
                }
                if (Inquire == "")
                {
                    groupBox7.BackColor = Color.IndianRed;
                    return;
                }

                try
                {
                    log.LogMessage("確認搜尋 開始", enumLogType.Trace);
                    Int32 _ALLUnitPrice = 0;
                    string _SQL = $@"SELECT No, Time, Name, Type, Count, UnitPrice, Unit, 
                        SalesArea, Paid, (Count * UnitPrice)AS Unpaid, PaidTime FROM SalesRecord ";
                    if (Inquire == "單號")
                    {
                        if (textBox21.Text.Length <= 11)
                            _SQL += $@" WHERE substr(No,1,11) = '{textBox21.Text}' ";
                        else
                            _SQL += $@" WHERE No Like '{textBox21.Text}%' ";
                    }
                    else if (Inquire == "姓名")
                    {
                        _SQL += $@" WHERE 1 = 1 ";
                        if (textBox21.Text != "")
                        {
                            _SQL += $@" AND Name LIKE '%{textBox21.Text}%' ";
                        }
                        if (checkBox4.Checked)
                        {
                            _SQL += $@" AND Time between '{dateTimePicker1.Value.ToString("yyyy-MM-dd")}' AND '{dateTimePicker2.Value.AddDays(1).ToString("yyyy-MM-dd")}' ";
                        }
                    }
                    else if (Inquire == "整合")
                    {
                        label25.Text = "";
                        label23.Text = "";
                        label28.Text = "";
                        dataGridView4.Rows.Clear();
                        checkedListBox1.Items.Clear();
                        checkedListBox1.Items.Add("已付款", true);
                        checkedListBox1.Items.Add("未付款", true);
                        _SQL = $@"SELECT No, MAX(Time)AS'Date', MAX(Name)AS'Name', ''AS'Buff1', ''AS'Buff2', ''AS'Buff3', 
                            ''AS'Buff4', SalesArea, MAX(Paid)AS'Paid', SUM(Round(Count * UnitPrice))AS Unpaid, ''AS'Buff6'
                            FROM SalesRecord WHERE 1 = 1 ";
                        if (textBox21.Text != "")
                        {
                            _SQL += $@" AND Name LIKE '%{textBox21.Text}%' ";
                        }
                        if (checkBox4.Checked)
                        {
                            _SQL += $@" AND Time between '{dateTimePicker1.Value.ToString("yyyy-MM-dd")}' AND '{dateTimePicker2.Value.AddDays(1).ToString("yyyy-MM-dd")}' ";
                        }
                        _SQL += $@"GROUP BY No;";
                        _SelectDT = dB_SQLite.GetDataTable(DB_Path, _SQL);
                        foreach (DataRow row in _SelectDT.Rows)
                        {
                            //單價金額加總
                            _ALLUnitPrice += (int)Math.Round(row.Field<Double>("Unpaid"), 0, MidpointRounding.AwayFromZero);
                        }
                        label23.Text = _ALLUnitPrice.ToString();
                        DatatableToDatagridview(_SelectDT, dataGridView4);
                        log.LogMessage("確認整合搜尋 成功 語法：" + _SQL, enumLogType.Info);
                        log.LogMessage("確認整合搜尋 成功 語法：" + _SQL, enumLogType.Trace);
                        return;
                    }
                    _SelectDT = dB_SQLite.GetDataTable(DB_Path, _SQL);
                    DatatableToDatagridview(_SelectDT, dataGridView4);

                    //總計算   //類型選單建立
                    Double _ALLCount = 0;
                    label28.Text = "";
                    checkedListBox1.Items.Clear();
                    checkedListBox1.Items.Add("已付款", true);
                    checkedListBox1.Items.Add("未付款", true);
                    checkedListBox2.Items.Clear();

                    List<ALLTypeModel> typeModels = new List<ALLTypeModel>();
                    List<String> salesAreaModels = new List<string>();
                    foreach (DataRow row in _SelectDT.Rows)
                    {
                        //類型匯入
                        string _Type = row.Field<String>("Type");
                        var typeModel = typeModels.Find(f => f.Type == _Type);
                        if (typeModel == null)
                        {
                            typeModels.Add(new ALLTypeModel() { Type = _Type });
                        }
                        //販售地區匯入
                        string _SalesArea = row.Field<String>("SalesArea");
                        if (!checkedListBox2.Items.Contains(_SalesArea))
                        {
                            checkedListBox2.Items.Add(_SalesArea, true);
                            salesAreaModels.Add(_SalesArea);
                        }
                        //單價金額加總
                        _ALLUnitPrice += (int)Math.Round(row.Field<Double>("Unpaid"), 0, MidpointRounding.AwayFromZero);
                        //單筆重量加總
                        typeModels.Find(f => f.Type == _Type)._ALLCount += row.Field<Double>("Count");
                        //重量加總
                        _ALLCount += row.Field<Double>("Count");
                    }
                    if (typeModels.Count <= 0)
                        label28.Text = "";
                    foreach (string item in TypeGradation())
                    {
                        ALLTypeModel typeModel = typeModels.Find(f => f.Type == item);
                        if (typeModel != null)
                        {
                            label28.Text += "【" + typeModel.Type + "：" + typeModel._ALLCount + "】";
                            if (!checkedListBox1.Items.Contains(typeModel.Type))
                            {
                                checkedListBox1.Items.Add(typeModel.Type, true);
                            }
                            typeModels.Remove(typeModel);
                        }
                    }
                    foreach (ALLTypeModel item in typeModels)
                    {
                        if (!checkedListBox1.Items.Contains(item.Type))
                        {
                            checkedListBox1.Items.Add(item.Type, true);
                        }
                        label28.Text += "【" + item.Type + "：" + item._ALLCount + "】";
                    }
                    label23.Text = _ALLUnitPrice.ToString();
                    label25.Text = _ALLCount.ToString();

                    log.LogMessage("確認" + Inquire + "搜尋 成功 總金額：" + label23.Text + "\r\n語法：" + _SQL, enumLogType.Info);
                    log.LogMessage("確認" + Inquire + "搜尋 成功 總金額：" + label23.Text + "\r\n語法：" + _SQL, enumLogType.Trace);
                }
                catch (Exception ee)
                {
                    log.LogMessage("確認搜尋 失敗：" + ee.Message, enumLogType.Error);
                    MessageBox.Show("確認搜尋 失敗：" + ee.Message);
                }
            }
        }

        //類型匯入後，點選變更
        private void checkedListBox_Click(object sender, EventArgs e)
        {
            timer_SelectType.Start();
        }
        private void timer_SelectType_Tick(object sender, EventArgs e)
        {
            try
            {
                log.LogMessage("類型點選變更 開始", enumLogType.Trace);
                string strCollected = string.Empty;
                Double _ALLUnitPrice = 0;
                Double _ALLCount = 0;
                DataTable dt_Buff = _SelectDT.Clone();
                DataTable dt = _SelectDT.Clone();
                Boolean _已付款 = false;
                Boolean _未付款 = false;
                label28.Text = "";
                List<ALLTypeModel> typeModels_Buff = new List<ALLTypeModel>();
                List<ALLTypeModel> typeModels = new List<ALLTypeModel>();

                for (int i = 0; i < checkedListBox1.Items.Count; i++)
                {
                    if (checkedListBox1.GetItemChecked(i))
                    {
                        DataRow[] rows;
                        string _Text = checkedListBox1.GetItemText(checkedListBox1.Items[i]);
                        string _SelectText = "Type = '" + _Text + "' ";

                        if (_Text == "已付款")
                        {
                            _已付款 = true;
                            continue;
                        }
                        else if (_Text == "未付款")
                        {
                            _未付款 = true;
                            continue;
                        }

                        if (_已付款 && _未付款)
                            rows = _SelectDT.Select(_SelectText);
                        else if (_已付款)
                            rows = _SelectDT.Select(_SelectText + "AND (Paid is not null)");
                        else if (_未付款)
                            rows = _SelectDT.Select(_SelectText + "AND (Paid is null)");
                        else
                            break;

                        if (rows.Count() > 0)
                            typeModels_Buff.Add(new ALLTypeModel() { Type = _Text });
                        foreach (DataRow row in rows)
                        {
                            dt_Buff.ImportRow(row);
                        }
                    }
                }
                //販賣地區
                for (int i = 0; i < checkedListBox2.Items.Count; i++)
                {
                    if (checkedListBox2.GetItemChecked(i))
                    {
                        DataRow[] rows;
                        string _Text = checkedListBox2.GetItemText(checkedListBox2.Items[i]);
                        rows = dt_Buff.Select($@"SalesArea = '{_Text}'");
                        foreach (DataRow row in rows)
                        {
                            //單價加總
                            _ALLUnitPrice += (int)Math.Round(row.Field<Double>("Unpaid"), 0, MidpointRounding.AwayFromZero);
                            //重量加總
                            _ALLCount += row.Field<Double>("Count");
                            //類型分類重量
                            string typeText = row.Field<String>("Type");
                            ALLTypeModel typeModel = typeModels_Buff.Find(f => f.Type == typeText);
                            //單筆重量加總
                            typeModel._ALLCount += row.Field<Double>("Count");
                            if (typeModel != null
                                && typeModels.Find(f => f.Type == typeText) == null)
                            {
                                typeModels.Add(typeModel);
                            }
                            dt.ImportRow(row);
                        }
                    }
                }
                if (dt.Rows.Count <= 0)
                {
                    dt = _SelectDT.Clone();
                    typeModels = new List<ALLTypeModel>();
                    _ALLUnitPrice = 0;
                    _ALLCount = 0;
                }
                if (Inquire == "整合")
                {
                    DataRow[] rows;
                    if (_已付款 && _未付款)
                        rows = _SelectDT.Select();
                    else if (_已付款)
                        rows = _SelectDT.Select("(Paid is not null)");
                    else if (_未付款)
                        rows = _SelectDT.Select("(Paid is null)");
                    else
                        rows = new DataRow[] { };
                    foreach (DataRow row in rows)
                    {
                        //單價加總
                        _ALLUnitPrice += (int)Math.Round(row.Field<Double>("Unpaid"), 0, MidpointRounding.AwayFromZero);
                        dt.ImportRow(row);
                    }
                }

                foreach (ALLTypeModel item in typeModels)
                {
                    label28.Text += "【" + item.Type + "：" + item._ALLCount + "】";
                }
                label23.Text = _ALLUnitPrice.ToString();
                label25.Text = _ALLCount.ToString();
                DatatableToDatagridview(dt, dataGridView4);

                log.LogMessage("類型點選變更 成功 總金額：" + label23.Text + "\t總重量：" + label25.Text, enumLogType.Info);
                log.LogMessage("類型點選變更 成功 總金額：" + label23.Text + "\t總重量：" + label25.Text, enumLogType.Trace);
                timer_SelectType.Stop();
            }
            catch (Exception ee)
            {
                log.LogMessage("類型點選變更 失敗：" + ee.Message, enumLogType.Error);
                timer_SelectType.Stop();
                MessageBox.Show("類型點選變更 失敗：" + ee.Message);
            }
        }

        //日期是否顯示
        private void checkBox4_CheckedChanged(object sender, EventArgs e)
        {
            panel21.Enabled = ((CheckBox)sender).Checked;
            panel21.Visible = ((CheckBox)sender).Checked;
        }
        //時間最大最小相互影響
        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            if (dateTimePicker1.Value > dateTimePicker2.Value)
                dateTimePicker2.Value = dateTimePicker1.Value;
        }
        private void dateTimePicker2_ValueChanged(object sender, EventArgs e)
        {
            if (dateTimePicker1.Value > dateTimePicker2.Value)
                dateTimePicker1.Value = dateTimePicker2.Value;
        }

        //新增客戶
        private void textBox22_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                IU_CustomerProfile();
            }
        }

        /// <summary>
        /// 新刪修客戶資料
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void radioButton_CustomerProfile_CheckedChanged(object sender, EventArgs e)
        {
            string _Text = ((ButtonBase)sender).Text;
            if (((RadioButton)sender).Checked)
            {
                if (_Text == "新增")
                {
                    IUDCustomerProfile = "I";
                    groupBox9.Text = "新增客戶資料";
                    textBox22.Location = new Point(comboBox1.Location.X + comboBox1.Width + 50, textBox22.Location.Y);
                    button4.Enabled = false;
                    button4.Visible = false;
                    button6.Enabled = false;
                    button6.Visible = false;
                    label24.Enabled = false;
                    label24.Visible = false;
                    button5.Enabled = true;
                    button5.Visible = true;
                }
                else if (_Text == "修改 / 刪除")
                {
                    IUDCustomerProfile = "U";
                    groupBox9.Text = "(修改 / 刪除) 客戶資料";
                    int _X = textBox22.Location.X;
                    textBox22.Location = new Point(label24.Location.X + label24.Width + 50, textBox22.Location.Y);
                    button4.Enabled = true;
                    button4.Visible = true;
                    button6.Enabled = true;
                    button6.Visible = true;
                    label24.Enabled = true;
                    label24.Visible = true;
                    button5.Enabled = false;
                    button5.Visible = false;
                }

                ((GroupBox)((RadioButton)sender).Parent).BackColor = Color.Transparent;
                for (int i = 0; i < ((RadioButton)sender).Parent.Controls.Count; i++)
                {
                    if (((RadioButton)((RadioButton)sender).Parent.Controls[i]).Checked)
                        ((RadioButton)((RadioButton)sender).Parent.Controls[i]).BackColor = Color.GreenYellow;
                    else
                        ((RadioButton)((RadioButton)sender).Parent.Controls[i]).BackColor = Color.Transparent;
                }
                textBox22.Text = "";
            }
        }
        private void IU_CustomerProfile()
        {
            if (textBox22.Text == "")
                return;
            if (comboBox1.Text == "")
                return;
            try
            {
                string _SQL = "";
                if (IUDCustomerProfile == "I")
                {
                    log.LogMessage("新增客戶 開始", enumLogType.Trace);
                    Int32 _ID = Int32.Parse(dB_SQLite.GetDataTable(DB_Path, $@"SELECT 
                        CASE WHEN MAX(ID) ISNULL THEN '001' ELSE MAX(ID)+1 END ID FROM CustomerProfile;").Rows[0][0].ToString());

                    _SQL = $@"INSERT INTO CustomerProfile (ID, CustomerID, CustomerName) 
                        VALUES ('{_ID}', '{comboBox1.Text + _ID.ToString("D3")}', '{textBox22.Text}');"; ;
                    dB_SQLite.Manipulate(DB_Path, _SQL);

                    log.LogMessage("新增客戶 成功 語法：" + _SQL, enumLogType.Trace);
                }
                else if (IUDCustomerProfile == "U")
                {
                    log.LogMessage("修改客戶 開始", enumLogType.Trace);
                    _SQL = $@"UPDATE CustomerProfile SET CustomerID = '{comboBox1.Text + label24.Text.Substring(1)}', 
                            CustomerName = '{textBox22.Text}' WHERE ID = '{Int32.Parse(label24.Text.Substring(1))}';";
                    dB_SQLite.Manipulate(DB_Path, _SQL);

                    log.LogMessage("修改客戶 成功 語法：" + _SQL, enumLogType.Trace);
                }
                _SQL = "SELECT * FROM CustomerProfile";
                DB_SQLite.DatatableToDatagridview(dB_SQLite.GetDataTable(DB_Path, _SQL), dataGridView5);
                textBox22.Text = "";
            }
            catch (Exception ee)
            {
                log.LogMessage("新增/修改 客戶 失敗：" + ee.Message, enumLogType.Error);
                MessageBox.Show("新增/修改 客戶 失敗：" + ee.Message);
            }
        }
        //新增/修改 客戶資料
        private void button_IU_CustomerProfile_Click(object sender, EventArgs e)
        {
            IU_CustomerProfile();
        }
        //刪除客戶資料
        private void button4_Click(object sender, EventArgs e)
        {
            if (label24.Text == "")
                return;
            try
            {
                string _SQL = ""; ;
                log.LogMessage("刪除客戶 開始", enumLogType.Trace);

                _SQL = $@"DELETE FROM CustomerProfile WHERE ID = '{Int32.Parse(label24.Text.Substring(1))}';"; ;
                dB_SQLite.Manipulate(DB_Path, _SQL);

                _SQL = "SELECT * FROM CustomerProfile";
                DB_SQLite.DatatableToDatagridview(dB_SQLite.GetDataTable(DB_Path, _SQL), dataGridView5);
                textBox22.Text = "";

                log.LogMessage("刪除客戶 成功 語法：" + _SQL, enumLogType.Trace);
            }
            catch (Exception ee)
            {
                log.LogMessage("刪除 客戶 失敗：" + ee.Message, enumLogType.Error);
                MessageBox.Show("刪除 客戶 失敗：" + ee.Message);
            }
        }
        //更新人員前的代入
        private void dataGridView5_Click(object sender, EventArgs e)
        {
            if (IUDCustomerProfile == "U")
            {
                comboBox1.Text = ((DataGridView)sender).Rows[((DataGridView)sender).CurrentRow.Index].Cells[1].Value.ToString().Substring(0, 1);
                label24.Text = ((DataGridView)sender).Rows[((DataGridView)sender).CurrentRow.Index].Cells[1].Value.ToString();
                textBox22.Text = ((DataGridView)sender).Rows[((DataGridView)sender).CurrentRow.Index].Cells[2].Value.ToString();
            }
        }

        //已付
        private void button7_Click(object sender, EventArgs e)
        {
            try
            {
                log.LogMessage("已付修改 開始", enumLogType.Trace);

                string _No = "";
                if (Inquire == "單號")
                {
                    _No = dataGridView4.Rows[0].Cells[0].Value.ToString();
                    UpdateNoPaid(_No, label23.Text);

                    string _SelectSQL = $@"SELECT No, Time, Name, Type, Count, UnitPrice, Unit, 
                        SalesArea, Paid, (Count * UnitPrice)AS Unpaid FROM SalesRecord  
                        WHERE No = '{_No}';";
                    DB_SQLite.DatatableToDatagridview(dB_SQLite.GetDataTable(DB_Path, _SelectSQL), dataGridView4);
                }
                else if (Inquire == "整合")
                {
                    DataGridViewRow row = new DataGridViewRow();
                    foreach (DataGridViewRow _row in dataGridView4.Rows)
                    {
                        UpdateNoPaid(_row.Cells[0].Value.ToString(), _row.Cells[9].Value.ToString());
                        _row.Cells[8].Value = _row.Cells[9].Value.ToString();
                    }
                }

                log.LogMessage("已付修改 成功路徑：" + DB_Path, enumLogType.Trace);
                log.LogMessage("已付修改 成功路徑：" + DB_Path, enumLogType.Info);
            }
            catch (Exception ee)
            {
                log.LogMessage("已付修改 失敗：" + ee.Message, enumLogType.Error);
                MessageBox.Show("已付修改 失敗：" + ee.Message);
            }
        }
        //單號已付
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

        //匯出Excel
        private void button8_Click(object sender, EventArgs e)
        {
            #region 匯出Excel
            try
            {
                log.LogMessage("匯出Excel 開始", enumLogType.Trace);
                button8.Enabled = false;
                //選取指定的資料夾
                FolderBrowserDialog folder = new FolderBrowserDialog();
                if (folder.ShowDialog() != DialogResult.OK)
                {
                    button8.Enabled = true;
                    return;
                }

                DataGridView view = dataGridView4;
                if (view.Rows.Count <= 0)
                    return;

                EPPlus ePPlus = new EPPlus();
                List<List<MExcelCell>> excelCells = new List<List<MExcelCell>>();
                List<MExcelCell> excelCell = new List<MExcelCell>();
                string _Path = "";
                string _No = "";
                DateTime _Now = new DateTime();
                string _Name = "";
                string _Unit = "";
                string _SalesArea = "";

                if (Inquire == "單號")
                {
                    _No = view.Rows[0].Cells[0].Value.ToString();
                    _Now = Convert.ToDateTime(view.Rows[0].Cells[1].Value.ToString());
                    _Name = view.Rows[0].Cells[2].Value.ToString();
                    _Unit = view.Rows[0].Cells[6].Value.ToString();
                    _SalesArea = view.Rows[0].Cells[7].Value.ToString();

                    ExcelProcess excel = new ExcelProcess(log);
                    _Path = folder.SelectedPath + $@"\{_No}_{_Name}.xlsx";
                    excel.ExcelExportImage_3(view, _Path, _Now, _No, _Name, _Unit, _SalesArea, true);
                }
                else if (Inquire == "整合")
                {
                    //廠商標題
                    excelCell.Add(new MExcelCell() { Content = Settings.廠商標題1 });
                    excelCells.Add(excelCell);
                    //空一行
                    excelCells.Add(new List<MExcelCell>());
                    excelCell = new List<MExcelCell>();
                    excelCell.Add(new MExcelCell() { Content = Settings.廠商標題2 });
                    excelCells.Add(excelCell);
                    excelCell = new List<MExcelCell>();
                    excelCell.Add(new MExcelCell() { Content = Settings.廠商標題3 });
                    excelCells.Add(excelCell);
                    //空一行
                    excelCells.Add(new List<MExcelCell>());

                    //頁首
                    List<string> _HideHeader = new List<string>() { "類型", "重量", "單價", "單位", "已付時間" };
                    excelCell = new List<MExcelCell>();
                    foreach (DataGridViewColumn col in view.Columns)
                    {
                        //隱藏
                        if (_HideHeader.Contains(col.HeaderText))
                        {
                            continue;
                        }
                        excelCell.Add(new MExcelCell()
                        {
                            Content = col.HeaderText
                        });
                    }
                    excelCells.Add(excelCell);

                    //內容
                    Int32 _ALLPrice = 0;
                    foreach (DataGridViewRow row in view.Rows)
                    {
                        excelCell = new List<MExcelCell>();
                        foreach (DataGridViewCell cell in row.Cells)
                        {
                            //隱藏
                            if (_HideHeader.Contains(view.Columns[cell.ColumnIndex].HeaderText))
                            {
                                continue;
                            }
                            excelCell.Add(new MExcelCell()
                            {
                                Content = cell.Value
                            });
                        }
                        excelCells.Add(excelCell);
                    }
                    //空一行
                    excelCells.Add(new List<MExcelCell>());
                    excelCell = new List<MExcelCell>();
                    excelCell.Add(new MExcelCell() { Content = "民智自動化有限公司" });
                    //總價
                    for (int i = 0; i < view.Columns.Count; i++)
                    {
                        //隱藏
                        if (_HideHeader.Contains(view.Columns[i].HeaderText))
                        {
                            continue;
                        }
                        excelCell.Add(new MExcelCell() { Content = "" });
                    }
                    excelCell.Remove(excelCell[excelCell.Count - 1]);
                    excelCell.Remove(excelCell[excelCell.Count - 1]);
                    excelCell.Remove(excelCell[excelCell.Count - 1]);
                    if (!string.IsNullOrEmpty(label23.Text))
                        _ALLPrice = Convert.ToInt32(label23.Text);
                    else
                        _ALLPrice = 0;
                    excelCell.Add(new MExcelCell() { Content = "總價" });
                    excelCell.Add(new MExcelCell() { Content = _ALLPrice });
                    excelCells.Add(excelCell);
                    //匯出成檔案
                    ePPlus.AddSheet(excelCells, "整合", 15, 14);
                    ePPlus.MergeColumn(1, 1, 2, 6);
                    ePPlus.FontSize(1, 1, 36, true, OfficeOpenXml.Style.ExcelBorderStyle.None);
                    ePPlus.ExcelCenterCell(1, 1, OfficeOpenXml.Style.ExcelHorizontalAlignment.CenterContinuous);
                    ePPlus.MergeColumn(3, 1, 3, 6);
                    ePPlus.FontSize(3, 1, 14, false, OfficeOpenXml.Style.ExcelBorderStyle.None);
                    ePPlus.ExcelCenterCell(3, 1, OfficeOpenXml.Style.ExcelHorizontalAlignment.CenterContinuous);
                    ePPlus.MergeColumn(4, 1, 4, 6);
                    ePPlus.FontSize(4, 1, 14, false, OfficeOpenXml.Style.ExcelBorderStyle.None);
                    ePPlus.ExcelCenterCell(4, 1, OfficeOpenXml.Style.ExcelHorizontalAlignment.CenterContinuous);
                    //for (int i = 0; i <= view.Rows.Count; i++)
                    //{
                    //    ePPlus.MergeColumn(i + 6, 3, i + 6, 4);
                    //}
                    ePPlus.MergeColumn(ePPlus.EndCell, 1, ePPlus.EndCell, 3);
                    ePPlus.FontSize(ePPlus.EndCell, 1, 14, false, OfficeOpenXml.Style.ExcelBorderStyle.None);
                    _Path = folder.SelectedPath + $@"\整合_{DateTime.Now.ToString("yyyyMMddhhmmss")}.xlsx";
                    ePPlus.Export(_Path);
                }

                button8.Enabled = true;
                log.LogMessage("匯出Excel 成功路徑：" + _Path, enumLogType.Trace);
                log.LogMessage("匯出Excel 成功", enumLogType.Info);
            }
            catch (Exception ee)
            {
                MessageBox.Show("匯出Excel 失敗：\r\n" + ee.Message);
                log.LogMessage("匯出Excel 失敗：\r\n" + ee.Message, enumLogType.Error);
                button8.Enabled = true;
                return;
            }
            #endregion
        }

        //類型順序調整
        public static List<string> TypeGradation()
        {
            #region 類型順序調整
            List<string> _TypeGradation = new List<string>() { };
            string[] _TypeGradation1 = Settings.類型1.Split('/');
            string[] _TypeGradation2 = Settings.類型2.Split('/');
            string[] _TypeGradation3 = Settings.類型3.Split('/');
            for (int i = 0; i < 8; i++)
            {
                if (i < _TypeGradation1.Length)
                {
                    if (!_TypeGradation.Contains(_TypeGradation1[i]))
                    {
                        _TypeGradation.Add(_TypeGradation1[i]);
                    }
                }
                if (i < _TypeGradation2.Length)
                {
                    if (!_TypeGradation.Contains(_TypeGradation2[i]))
                    {
                        _TypeGradation.Add(_TypeGradation2[i]);
                    }
                }
                if (i < _TypeGradation3.Length)
                {
                    if (!_TypeGradation.Contains(_TypeGradation3[i]))
                    {
                        _TypeGradation.Add(_TypeGradation3[i]);
                    }
                }
            }
            return _TypeGradation;
            #endregion
        }
    }
    public class ALLTypeModel
    {
        /// <summary>各類型名稱</summary>
        public string Type { get; set; } = string.Empty;

        /// <summary>各類型總重量</summary>
        public Double _ALLCount { get; set; } = 0;

        /// <summary>各類型總金額</summary>
        public Int32 _ALLMoney { get; set; } = 0;

        /// <summary>各類型單價</summary>
        public Int32 _UnitPrice { get; set; } = 0;
    }
    public class Integrate
    {
        /// <summary>序號</summary>
        public string No { get; set; } = string.Empty;

        /// <summary>日期</summary>
        public string Date { get; set; } = string.Empty;

        /// <summary>客戶名稱</summary>
        public string Name { get; set; } = string.Empty;

        /// <summary>已付金額</summary>
        public string Paid { get; set; } = string.Empty;

        /// <summary>未付金額</summary>
        public Int32 Unpaid { get; set; } = 0;
    }
}
