using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.Entity.Core.Common.CommandTrees.ExpressionBuilder;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using 簡易倉儲系統.DB;
using 簡易倉儲系統.EssentialTool;
using static System.Data.Entity.Infrastructure.Design.Executor;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Button;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.ProgressBar;
using static 簡易倉儲系統.EssentialTool.LogToText;

namespace 簡易倉儲系統
{
    public partial class ManagerView : Form
    {
        LogToText log = new LogToText(@".\Log");

        public static string Setting_Path = @".\";          //設定檔路徑
        public static string VersionNumber = "";            //程式版號
        public static string DB_Path = "";    //DB路徑
        public static string[][] type = { new string[] { "", "", "", "", "", "", "" }
                                        , new string[] { "", "", "", "", "", "", "" }
                                        , new string[] { "", "", "", "", "", "" } };

        public ManagerView()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
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
            }
            log.LogMessage("系統啓動", enumLogType.Trace);
            log.LogMessage("管理者介面啓動", enumLogType.Info);

            try
            {
                DB_Path = Settings.資料庫路徑 + @"data.db";
                DB_SQLite dB_SQLite = new DB_SQLite();

                if (!File.Exists(DB_Path))
                {
                    log.LogMessage("偵測到無資料庫，準備開始建立。", enumLogType.Debug);
                    var createtablestring = "";

                    // 建立 SQLite 資料庫
                    dB_SQLite.CreateDatabase(DB_Path);

                    // 建立資料表 販售紀錄 SalesRecord
                    createtablestring = @"CREATE TABLE SalesRecord (No Integer, Date DateTime, Type TEXT, Count double
                    , UnitPrice double, Unit TEXT, salesArea TEXT);";
                    dB_SQLite.CreateTable(DB_Path, createtablestring);

                    // 建立資料表 外銷韓國 ExportKoreaUnitPrice
                    createtablestring = @"CREATE TABLE ExportKoreaUnitPrice (Date DateTime, Type1 double, Type2 double
                    , Type3 double, Type4 double, Type5 double, Type6 double, Type7 double);";
                    dB_SQLite.CreateTable(DB_Path, createtablestring);

                    // 建立資料表 外銷日本 ExportJapanUnitPrice
                    createtablestring = @"CREATE TABLE ExportJapanUnitPrice (Date DateTime, Type1 double, Type2 double
                    , Type3 double, Type4 double, Type5 double, Type6 double, Type7 double);";
                    dB_SQLite.CreateTable(DB_Path, createtablestring);

                    // 建立資料表 超市 ExportSupermarketUnitPrice
                    createtablestring = @"CREATE TABLE ExportSupermarketUnitPrice (Date DateTime, Type1 double, Type2 double
                    , Type3 double, Type4 double, Type5 double, Type6 double);";
                    dB_SQLite.CreateTable(DB_Path, createtablestring);

                    log.LogMessage("建立資料庫 成功。", enumLogType.Debug);
                }
                //ianTest
                //var insertstring = @"
                //    INSERT INTO ExportKoreaUnitPrice (Date, Type1, Type2, Type3, Type4, Type5, Type6, Type7) VALUES ('2023-03-11', '100', '20', '200', '100', '20', '200', '3');
                //    INSERT INTO ExportKoreaUnitPrice (Date, Type1, Type2, Type3, Type4, Type5, Type6, Type7) VALUES ('2023-03-12', '100', '20', '200', '100', '20', '200', '3');
                //    INSERT INTO ExportJapanUnitPrice (Date, Type1, Type2, Type3, Type4, Type5, Type6, Type7) VALUES ('2023-03-11', '100', '20', '200', '100', '20', '200', '3');
                //    INSERT INTO ExportJapanUnitPrice (Date, Type1, Type2, Type3, Type4, Type5, Type6, Type7) VALUES ('2023-03-12', '100', '20', '200', '100', '20', '200', '3');
                //    INSERT INTO ExportSupermarketUnitPrice (Date, Type1, Type2, Type3, Type4, Type5, Type6) VALUES ('2023-03-11', '100', '20', '200', '100', '20', '200');
                //    INSERT INTO ExportSupermarketUnitPrice (Date, Type1, Type2, Type3, Type4, Type5, Type6) VALUES ('2023-03-12', '100', '20', '200', '100', '20', '200');
                //";
                //dB_SQLite.Manipulate(DB_Path, insertstring);

                // 讀取資料
                DatatableToDatagridview(dB_SQLite.GetDataTable(DB_Path, @"SELECT * FROM ExportKoreaUnitPrice"), dataGridView1);
                DatatableToDatagridview(dB_SQLite.GetDataTable(DB_Path, @"SELECT * FROM ExportJapanUnitPrice"), dataGridView2);
                DatatableToDatagridview(dB_SQLite.GetDataTable(DB_Path, @"SELECT * FROM ExportSupermarketUnitPrice"), dataGridView3);
                log.LogMessage("讀取資料庫 成功。", enumLogType.Trace);
            }
            catch (Exception ee)
            {
                log.LogMessage("連線資料庫 失敗\r\n" + ee.Message, enumLogType.Error);
                MessageBox.Show("連線資料庫 失敗\r\n" + ee.Message);
                return;
            }
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
                        if (j == 0)
                            _value = DateTime.Parse(_value).ToString("yyyy-MM-dd");
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
                log.LogMessage("Datatable轉出Datagridview 失敗：\r\n" + ee.Message, enumLogType.Error);
                return false;
            }
        }

        //換分頁時清空資料
        private void tabControl_Click(object sender, EventArgs e)
        {
            log.LogMessage("換分頁時清空資料 開始", enumLogType.Trace);

            //清空暫存
            type = new string[][] { new string[] { "", "", "", "", "", "", "" }
                                  , new string[] { "", "", "", "", "", "", "" }
                                  , new string[] { "", "", "", "", "", "" } };

            //要清空的TextBox元件
            System.Windows.Forms.TextBox[] _textBoxes = { textBox1, textBox2, textBox3, textBox4, textBox5, textBox6, textBox7
                    , textBox8, textBox9, textBox10, textBox11, textBox12, textBox13, textBox14
                    , textBox15, textBox16, textBox17, textBox18, textBox19, textBox20};
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
                DataGridView _view = (DataGridView)((Control)sender).Parent.Parent.Controls[0];

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
                DB_SQLite dB_SQLite = new DB_SQLite();
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
                    log.LogMessage("新增/修改_DB修改 成功", enumLogType.Info);
                    #endregion

                    #region DataGridView修改

                    //修改
                    _data.SetValues(_typeList.ToArray());

                    log.LogMessage("新增/修改_DataGridView修改 成功資料：[" + string.Join(", ", _typeList.ToArray()) + "]", enumLogType.Trace);
                    log.LogMessage("新增/修改_DataGridView修改 成功", enumLogType.Info);
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
                    log.LogMessage("新增/修改_DB新增 成功", enumLogType.Info);
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
                    log.LogMessage("新增/修改_DataGridView新增 成功", enumLogType.Info);
                    #endregion

                }
            }
            catch (Exception ee)
            {
                MessageBox.Show("新增修改失敗\r\n" + ee.Message);
            }
        }

        private void dataGridView_Click(object sender, EventArgs e)
        {
            try
            {
                log.LogMessage("dataGridView轉出 開始", enumLogType.Trace);

                if (((DataGridView)sender).Rows[((DataGridView)sender).CurrentRow.Index].Cells[0].Value == null)
                    return;
                for (int i = 1; i < ((DataGridView)sender).Rows[((DataGridView)sender).CurrentRow.Index].Cells.Count; i++)
                {
                    ((Control)sender).Parent.Controls[1].Controls[i - 1].Controls[1].Text =
                        ((DataGridView)sender).Rows[((DataGridView)sender).CurrentRow.Index].Cells[i].Value.ToString();
                }
                
                log.LogMessage("dataGridView轉出 成功", enumLogType.Trace);
            }
            catch (Exception ee)
            {
                log.LogMessage("dataGridView轉出 失敗：" + ee.Message, enumLogType.Error);
            }
        }
    }
}
