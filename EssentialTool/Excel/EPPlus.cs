using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Printing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using Spire.Xls;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;
using static 簡易倉儲系統.EssentialTool.LogToText;
//using Microsoft.Office.Tools.Excel;

namespace 簡易倉儲系統.EssentialTool.Excel
{
    internal class EPPlus
    {
        /// <summary>起始列</summary>
        public int BeginCell { get; set; }
        /// <summary>終點列</summary>
        public int EndCell { get; set; }
        /// <summary>建立Excel</summary>
        private ExcelPackage Epackage = new ExcelPackage();
        /// <summary>建立第一個Sheet，後方為定義Sheet的名稱</summary>
        private ExcelWorksheet Sheet;

        /// <summary>加入工作表</summary>
        public void AddSheet(List<List<MExcelCell>> Table, string SheetName, int ColumnWidth = 0, float Size = 11)
        {
            //建立一個Sheet，後方為定義Sheet的名稱
            Sheet = Epackage.Workbook.Worksheets.Add(SheetName);

            //內文
            for (int col = 0; col < Table.Count; col++)
            {
                for (int row = 0; row < Table[col].Count; row++)
                {
                    int x = col + 1;
                    int y = row + 1;
                    if (!string.IsNullOrEmpty(Table[col][row].Formula))
                    {
                        Sheet.Cells[x, y].Formula = Table[col][row].Formula;
                    }
                    else if (string.IsNullOrEmpty(Table[col][row].Content.ToString()))
                    {
                        continue;
                    }
                    else
                        Sheet.Cells[x, y].Value = Table[col][row].Content;
                    if (!string.IsNullOrEmpty(Table[col][row].NumberFormat))
                        Sheet.Cells[x, y].Style.Numberformat.Format = Table[col][row].NumberFormat;
                    if (Table[col][row].PatternType != ExcelFillStyle.None)
                        Sheet.Cells[x, y].Style.Fill.PatternType = Table[col][row].PatternType;
                    if (Table[col][row].BackgroundColor != System.Drawing.Color.White)
                        Sheet.Cells[x, y].Style.Fill.BackgroundColor.SetColor(Table[col][row].BackgroundColor);
                    Sheet.Cells[x, y].Style.Font.Size = Size;
                    Sheet.Cells[x, y].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                }
            }

            //自動調整欄寬
            for (int i = 1; i <= Table[1].Count; i++)
            {
                Sheet.Column(i).AutoFit();
                Sheet.Column(i).Width += 2;
            }
            if (ColumnWidth > 0)
            {
                if (Table.Count >= 6)
                {
                    for (int i = 1; i <= Table[5].Count; i++)
                    {
                        Sheet.Column(i).Width = ColumnWidth;
                    }
                }
            }
            //高度設定
            for (int i = 1; i <= Table.Count; i++)
            {
                Sheet.Row(i).Height = 20;
            }

            //調整邊距
            decimal inch = 1M / 2.54M;
            Sheet.PrinterSettings.TopMargin = inch;//因為EPPlus單位都是英吋
            Sheet.PrinterSettings.LeftMargin = inch;
            Sheet.PrinterSettings.RightMargin = inch;
            Sheet.PrinterSettings.BottomMargin = inch;
            
            //取得起始列
            GetBeginCell();
            //取得終點列
            GetEndCell();
        }

        /// <summary>儲存格置中</summary>
        public void ExcelCenterCell(int row, int col, OfficeOpenXml.Style.ExcelHorizontalAlignment excelHorizontalAlignment, Boolean Vertical = true)
        {
            Sheet.Cells[row, col].Style.HorizontalAlignment = excelHorizontalAlignment;
            if (Vertical)
                Sheet.Cells[row, col].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
        }

        /// <summary>儲存格字體大小</summary>
        public void FontSize(int row, int col, float size, Boolean isBold, ExcelBorderStyle excelBorderStyle)
        {
            Sheet.Cells[row, col].Style.Font.Size = size;
            Sheet.Cells[row, col].Style.Font.Bold = isBold;
            Sheet.Cells[row, col].Style.Border.BorderAround(excelBorderStyle);
        }
        /// <summary>設定欄位上方框框</summary>
        public void FontSizeTop(int row, int col, ExcelBorderStyle excelBorderStyle)
        {
            if (row > 1)
                Sheet.Cells[row - 1, col].Style.Border.Bottom.Style = excelBorderStyle;
            Sheet.Cells[row, col].Style.Border.Top.Style = excelBorderStyle;
        }

        /// <summary>合併儲存格</summary>
        public void MergeColumn(int row1, int col1, int row2, int col2)
        {
            Sheet.Cells[row1, col1, row2, col2].Merge = true;
        }

        /// <summary>設定寬度</summary>
        public void ColumnWidth(int col, int value)
        {
            Sheet.Column(col).Width = value;
        }

        /// <summary>取得起始列</summary>
        public void GetBeginCell()
        {
            string address = Sheet.Dimension.Address;
            string[] cells = address.Split(new char[] { ':' });
            BeginCell = Int32.Parse(Regex.Replace(cells[0], "[^0-9]", ""));
        }

        /// <summary>取得終點列</summary>
        public void GetEndCell()
        {
            string address = Sheet.Dimension.Address;
            string[] cells = address.Split(new char[] { ':' });
            EndCell = Int32.Parse(Regex.Replace(cells[1], "[^0-9]", ""));
        }

        /// <summary>指定列文字對齊(目前測試只有套用到2列以上才會生效)</summary>
        public void Alignment(string range, OfficeOpenXml.Style.ExcelHorizontalAlignment HorAlign = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center, OfficeOpenXml.Style.ExcelVerticalAlignment VerAlign = OfficeOpenXml.Style.ExcelVerticalAlignment.Center)
        {
            Sheet.Cells[range].Style.HorizontalAlignment = HorAlign;
            Sheet.Cells[range].Style.VerticalAlignment = VerAlign;
        }

        /// <summary>匯出物件資料</summary>
        public MemoryStream Export()
        {
            //因為ep.Stream是加密過的串流，故要透過SaveAs將資料寫到MemoryStream，
            //在將MemoryStream使用FileStreamResult回傳到前端。
            MemoryStream fileStream = new MemoryStream();
            Epackage.SaveAs(fileStream);

            Epackage.Dispose();
            //如果這邊不下Dispose，建議此ep要用using包起來，
            //但是要記得先將資料寫進MemoryStream在Dispose。
            fileStream.Position = 0;
            //不重新將位置設為0，excel開啟後會出現錯誤
            //經銷商審核資料OR店務資料
            return fileStream;
        }
        public void Export(string fileName)
        {
            FileInfo file = new FileInfo(fileName);
            Epackage.SaveAs(file);
            Epackage.Dispose();
        }
        public void ChangeExcel2Image(string filename, string ImageName)
        {
            using (Workbook workbook = new Workbook())
            {
                workbook.LoadFromFile(filename);
                using (Worksheet sheet = workbook.Worksheets[0])
                {
                    sheet.SaveToImage(ImageName); //圖片尾碼.bmp ,imagepath自己設置
                    //sheet.ToImage(1, 1, BeginCell, EndCell);  //直接轉出Image
                }
            }
        }
    }

    /// <summary>Excel欄位資料</summary>
    public class MExcelCell
    {
        /// <summary>欄位內容</summary>
        public object Content { get; set; } = string.Empty;
        /// <summary>背景顏色</summary>
        public Color BackgroundColor { get; set; } = Color.White;
        /// <summary>樣式</summary>
        public ExcelFillStyle PatternType { get; set; } = ExcelFillStyle.None;
        /// <summary>數字格式</summary>
        public string NumberFormat { get; set; } = string.Empty;
        /// <summary>公式</summary>
        public string Formula { get; set; } = string.Empty;
        /// <summary>文字水平位置</summary>
        public OfficeOpenXml.Style.ExcelHorizontalAlignment HorAlign { get; set; } = OfficeOpenXml.Style.ExcelHorizontalAlignment.General;
    }
}
