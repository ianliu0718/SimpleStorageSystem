﻿using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;
using System;
using System.Collections.Generic;
using System.Drawing.Printing;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static 簡易倉儲系統.EssentialTool.LogToText;
using System.Windows.Forms;
using System.Reflection.Emit;

namespace 簡易倉儲系統.EssentialTool.Excel
{
    internal class ExcelProcess
    {
        public static LogToText log;
        public ExcelProcess(LogToText _log)
        {
            log = _log;
        }

        public Boolean ExcelExportImage_3(DataGridView view, string Path, DateTime Now, string No, string Name, string Unit, string SalesArea, Boolean UnitPriceShow)
        {
            try
            {
                EPPlus ePPlus = new EPPlus();
                List<List<MExcelCell>> excelCells = new List<List<MExcelCell>>();
                List<MExcelCell> excelCell = new List<MExcelCell>();

                #region 收款聯
                //廠商標題
                excelCell.Add(new MExcelCell() { Content = Settings.廠商標題1 });
                excelCells.Add(excelCell);
                excelCells.Add(new List<MExcelCell>()); //空一行
                excelCell = new List<MExcelCell>();
                excelCell.Add(new MExcelCell() { Content = Settings.廠商標題2 });
                excelCells.Add(excelCell);
                excelCell = new List<MExcelCell>();
                excelCell.Add(new MExcelCell() { Content = Settings.廠商標題3 });
                excelCells.Add(excelCell);

                //標頭
                excelCell = new List<MExcelCell>();
                excelCell.Add(new MExcelCell() { Content = "單號" });
                excelCell.Add(new MExcelCell() { Content = No });
                excelCell.Add(new MExcelCell() { Content = " " });
                excelCell.Add(new MExcelCell() { Content = "姓名" });
                excelCell.Add(new MExcelCell() { Content = Name });
                excelCell.Add(new MExcelCell() { Content = " " });
                excelCell.Add(new MExcelCell() { Content = " " });
                excelCell.Add(new MExcelCell() { Content = " " });
                excelCell.Add(new MExcelCell() { Content = " " });
                excelCells.Add(excelCell);
                excelCell = new List<MExcelCell>();
                excelCell.Add(new MExcelCell() { Content = " " });
                excelCell.Add(new MExcelCell() { Content = " " });
                excelCell.Add(new MExcelCell() { Content = " " });
                excelCell.Add(new MExcelCell() { Content = " " });
                excelCell.Add(new MExcelCell() { Content = " " });
                excelCell.Add(new MExcelCell() { Content = " " });
                excelCell.Add(new MExcelCell() { Content = " " });
                excelCell.Add(new MExcelCell() { Content = " " });
                excelCell.Add(new MExcelCell() { Content = " " });
                excelCells.Add(excelCell);
                excelCell = new List<MExcelCell>();
                excelCell.Add(new MExcelCell() { Content = "日期" });
                excelCell.Add(new MExcelCell() { Content = Now.ToString("yyyy-MM-dd") });
                excelCell.Add(new MExcelCell() { Content = " " });
                excelCell.Add(new MExcelCell() { Content = "單位" });
                excelCell.Add(new MExcelCell() { Content = Unit });
                excelCell.Add(new MExcelCell() { Content = " " });
                excelCell.Add(new MExcelCell() { Content = "地區" });
                excelCell.Add(new MExcelCell() { Content = SalesArea });
                excelCell.Add(new MExcelCell() { Content = " " });
                excelCells.Add(excelCell);
                excelCell = new List<MExcelCell>();
                excelCell.Add(new MExcelCell() { Content = " " });
                excelCell.Add(new MExcelCell() { Content = " " });
                excelCell.Add(new MExcelCell() { Content = " " });
                excelCell.Add(new MExcelCell() { Content = " " });
                excelCell.Add(new MExcelCell() { Content = " " });
                excelCell.Add(new MExcelCell() { Content = " " });
                excelCell.Add(new MExcelCell() { Content = " " });
                excelCell.Add(new MExcelCell() { Content = " " });
                excelCell.Add(new MExcelCell() { Content = " " });
                excelCells.Add(excelCell);

                List<ALLTypeModel> typeModelsBuff = new List<ALLTypeModel>();
                List<ALLTypeModel> typeModels = new List<ALLTypeModel>();

                int _TypeIndex = -1;
                int _CountIndex = -1;
                int _MoneyIndex = -1;
                foreach (DataGridViewColumn column in view.Columns)
                {
                    if (column.HeaderText == "類型")
                        _TypeIndex = column.Index;
                    if (column.HeaderText == "重量")
                        _CountIndex = column.Index;
                    if (column.HeaderText == "單價")
                        _MoneyIndex = column.Index;
                    if (_TypeIndex != -1 && _CountIndex != -1 && _MoneyIndex != -1)
                        break;
                }
                foreach (DataGridViewRow row in view.Rows)
                {
                    //類型匯入
                    string _Type = row.Cells[_TypeIndex].Value.ToString();
                    ALLTypeModel typeModel = typeModelsBuff.Find(f => f.Type == _Type);
                    if (typeModel == null)
                    {
                        typeModel = new ALLTypeModel() { Type = _Type };
                        typeModelsBuff.Add(typeModel);
                    }
                    //單筆重量加總
                    typeModel._ALLCount += Convert.ToDouble(row.Cells[_CountIndex].Value.ToString());
                    //單筆金額加總
                    if (UnitPriceShow)
                    {
                        if (typeModel._UnitPrice == 0)
                            typeModel._UnitPrice = Convert.ToInt32(row.Cells[_MoneyIndex].Value.ToString());
                        Double _money = Convert.ToDouble(Convert.ToDouble(row.Cells[_CountIndex].Value.ToString()) * Convert.ToInt32(row.Cells[_MoneyIndex].Value.ToString()));
                        typeModel._ALLMoney += (int)Math.Round(_money, 0, MidpointRounding.AwayFromZero);
                    }
                }
                //排序
                foreach (string item in ManagerView.TypeGradation())
                {
                    ALLTypeModel typeModel = typeModelsBuff.Find(f => f.Type == item);
                    if (typeModel != null)
                    {
                        typeModels.Add(typeModel);
                    }
                }
                for (int i = 0; i < typeModels.Count; i = i + 2)
                {
                    excelCell = new List<MExcelCell>();
                    List<MExcelCell> excelCell2 = new List<MExcelCell>();
                    excelCell.Add(new MExcelCell() { Content = typeModels[i].Type.ToString() });
                    excelCell.Add(new MExcelCell() { Content = typeModels[i]._ALLCount.ToString() });
                    excelCell.Add(new MExcelCell() { Content = "金額：" + typeModels[i]._ALLMoney.ToString() });
                    excelCell.Add(new MExcelCell() { Content = " " });
                    excelCell2.Add(new MExcelCell() { Content = "(" + typeModels[i]._UnitPrice.ToString() + ")" });
                    excelCell2.Add(new MExcelCell() { Content = Unit });
                    excelCell2.Add(new MExcelCell() { Content = " " });
                    excelCell2.Add(new MExcelCell() { Content = " " });
                    if (i + 1 < typeModels.Count)
                    {
                        excelCell.Add(new MExcelCell() { Content = typeModels[i + 1].Type.ToString() });
                        excelCell.Add(new MExcelCell() { Content = typeModels[i + 1]._ALLCount.ToString() });
                        excelCell.Add(new MExcelCell() { Content = "金額：" + typeModels[i + 1]._ALLMoney.ToString() });
                        excelCell.Add(new MExcelCell() { Content = " " });
                        excelCell2.Add(new MExcelCell() { Content = "(" + typeModels[i + 1]._UnitPrice.ToString() + ")" });
                        excelCell2.Add(new MExcelCell() { Content = Unit });
                        excelCell2.Add(new MExcelCell() { Content = " " });
                        excelCell2.Add(new MExcelCell() { Content = " " });
                    }
                    excelCells.Add(excelCell);
                    excelCells.Add(excelCell2);
                }
                ////空一行
                //excelCells.Add(new List<MExcelCell>());
                excelCell = new List<MExcelCell>();
                excelCell.Add(new MExcelCell() { Content = "" });
                excelCell.Add(new MExcelCell() { Content = "" });
                excelCell.Add(new MExcelCell() { Content = "" });
                excelCell.Add(new MExcelCell() { Content = "" });
                excelCell.Add(new MExcelCell() { Content = "" });
                excelCell.Add(new MExcelCell() { Content = "" });
                excelCell.Add(new MExcelCell() { Content = "" });
                excelCell.Add(new MExcelCell() { Content = "總重" });
                excelCell.Add(new MExcelCell() { Content = typeModels.Sum(s => s._ALLCount) });
                excelCells.Add(excelCell);

                excelCell = new List<MExcelCell>();
                excelCell.Add(new MExcelCell() { Content = "民智自動化有限公司" });
                excelCell.Add(new MExcelCell() { Content = "" });
                excelCell.Add(new MExcelCell() { Content = "" });
                excelCell.Add(new MExcelCell() { Content = "收款聯" });
                excelCell.Add(new MExcelCell() { Content = "" });
                excelCell.Add(new MExcelCell() { Content = "" });
                //總價
                if (UnitPriceShow)
                {
                    excelCell.Add(new MExcelCell() { Content = "" });
                    int _ALLPrice = 0;
                    foreach (var typeModel in typeModels)
                    {
                        _ALLPrice += typeModel._ALLMoney;
                    }
                    excelCell.Add(new MExcelCell() { Content = "總價" });
                    excelCell.Add(new MExcelCell() { Content = _ALLPrice });
                }
                excelCells.Add(excelCell);
                #endregion

                //深層複製至Buff，深層複製才可以修改內容
                List<List<MExcelCell>> excelCellsBuff = new List<List<MExcelCell>>();
                foreach (var _excelCell in excelCells)
                {
                    excelCell = new List<MExcelCell>();
                    foreach (var _Cell in _excelCell)
                    {
                        excelCell.Add(new MExcelCell() { Content = _Cell.Content });
                    }
                    excelCellsBuff.Add(excelCell);
                }
                //拉出上下中間的間隔
                int _X = 21;
                for (int i = excelCells.Count; i < _X; i++)
                {
                    excelCells.Add(new List<MExcelCell>());
                }

                #region 客戶聯

                excelCellsBuff[excelCellsBuff.Count - 1][3].Content = "客戶聯";
                excelCells.InsertRange(excelCells.Count, excelCellsBuff);

                #endregion


                //匯出成檔案
                ePPlus.AddSheet(excelCells, No);
                ePPlus.MergeColumn(1, 1, 2, 9);
                ePPlus.FontSize(1, 1, 36, true, OfficeOpenXml.Style.ExcelBorderStyle.None);
                ePPlus.ExcelCenterCell(1, 1, OfficeOpenXml.Style.ExcelHorizontalAlignment.Center);
                ePPlus.MergeColumn(3, 1, 3, 9);
                ePPlus.FontSize(3, 1, 14, false, OfficeOpenXml.Style.ExcelBorderStyle.None);
                ePPlus.ExcelCenterCell(3, 1, OfficeOpenXml.Style.ExcelHorizontalAlignment.Center);
                ePPlus.MergeColumn(4, 1, 4, 9);
                ePPlus.FontSize(4, 1, 14, false, OfficeOpenXml.Style.ExcelBorderStyle.None);
                ePPlus.ExcelCenterCell(4, 1, OfficeOpenXml.Style.ExcelHorizontalAlignment.Center);
                ePPlus.MergeColumn(5, 1, 6, 1);
                ePPlus.FontSize(5, 1, 13, false, OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                ePPlus.ExcelCenterCell(5, 1, OfficeOpenXml.Style.ExcelHorizontalAlignment.Center);
                ePPlus.MergeColumn(5, 2, 6, 3);
                ePPlus.FontSize(5, 2, 13, false, OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                ePPlus.ExcelCenterCell(5, 2, OfficeOpenXml.Style.ExcelHorizontalAlignment.Center);
                ePPlus.MergeColumn(5, 4, 6, 4);
                ePPlus.FontSize(5, 4, 13, false, OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                ePPlus.ExcelCenterCell(5, 4, OfficeOpenXml.Style.ExcelHorizontalAlignment.Center);
                ePPlus.MergeColumn(5, 5, 6, 9);
                ePPlus.FontSize(5, 5, 13, false, OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                ePPlus.ExcelCenterCell(5, 5, OfficeOpenXml.Style.ExcelHorizontalAlignment.Center);
                ePPlus.MergeColumn(7, 1, 8, 1);
                ePPlus.FontSize(7, 1, 13, false, OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                ePPlus.ExcelCenterCell(7, 1, OfficeOpenXml.Style.ExcelHorizontalAlignment.Center);
                ePPlus.MergeColumn(7, 2, 8, 3);
                ePPlus.FontSize(7, 2, 13, false, OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                ePPlus.ExcelCenterCell(7, 2, OfficeOpenXml.Style.ExcelHorizontalAlignment.Center);
                ePPlus.MergeColumn(7, 4, 8, 4);
                ePPlus.FontSize(7, 4, 13, false, OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                ePPlus.ExcelCenterCell(7, 4, OfficeOpenXml.Style.ExcelHorizontalAlignment.Center);
                ePPlus.MergeColumn(7, 5, 8, 6);
                ePPlus.FontSize(7, 5, 13, false, OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                ePPlus.ExcelCenterCell(7, 5, OfficeOpenXml.Style.ExcelHorizontalAlignment.Center);
                ePPlus.MergeColumn(7, 7, 8, 7);
                ePPlus.FontSize(7, 7, 13, false, OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                ePPlus.ExcelCenterCell(7, 7, OfficeOpenXml.Style.ExcelHorizontalAlignment.Center);
                ePPlus.MergeColumn(7, 8, 8, 9);
                ePPlus.FontSize(7, 8, 13, false, OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                ePPlus.ExcelCenterCell(7, 8, OfficeOpenXml.Style.ExcelHorizontalAlignment.Center);
                for (int i = 0; i < typeModels.Count; i = i + 2)
                {
                    //ePPlus.MergeColumn(9 + ((i + i) / 2), 1, 10 + ((i + i) / 2), 1);
                    ePPlus.FontSize(9 + ((i + i) / 2), 1, 13, false, OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                    ePPlus.ExcelCenterCell(9 + ((i + i) / 2), 1, OfficeOpenXml.Style.ExcelHorizontalAlignment.Center);
                    ePPlus.FontSize(10 + ((i + i) / 2), 1, 13, false, OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                    ePPlus.FontSizeTop(10 + ((i + i) / 2), 1, OfficeOpenXml.Style.ExcelBorderStyle.None);
                    ePPlus.ExcelCenterCell(10 + ((i + i) / 2), 1, OfficeOpenXml.Style.ExcelHorizontalAlignment.Center);
                    //ePPlus.MergeColumn(9 + ((i + i) / 2), 2, 10 + ((i + i) / 2), 2);
                    ePPlus.FontSize(9 + ((i + i) / 2), 2, 13, false, OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                    ePPlus.ExcelCenterCell(9 + ((i + i) / 2), 2, OfficeOpenXml.Style.ExcelHorizontalAlignment.Center);
                    ePPlus.FontSize(10 + ((i + i) / 2), 2, 13, false, OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                    ePPlus.FontSizeTop(10 + ((i + i) / 2), 2, OfficeOpenXml.Style.ExcelBorderStyle.None);
                    ePPlus.ExcelCenterCell(10 + ((i + i) / 2), 2, OfficeOpenXml.Style.ExcelHorizontalAlignment.Center);
                    ePPlus.MergeColumn(9 + ((i + i) / 2), 3, 10 + ((i + i) / 2), 4);
                    ePPlus.FontSize(9 + ((i + i) / 2), 3, 13, false, OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                    ePPlus.ExcelCenterCell(9 + ((i + i) / 2), 3, OfficeOpenXml.Style.ExcelHorizontalAlignment.Center);
                    if (i + 1 < typeModels.Count)
                    {
                        //ePPlus.MergeColumn(9 + ((i + i) / 2), 5, 10 + ((i + i) / 2), 5);
                        ePPlus.FontSize(9 + ((i + i) / 2), 5, 13, false, OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                        ePPlus.ExcelCenterCell(9 + ((i + i) / 2), 5, OfficeOpenXml.Style.ExcelHorizontalAlignment.Center);
                        ePPlus.FontSize(10 + ((i + i) / 2), 5, 13, false, OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                        ePPlus.FontSizeTop(10 + ((i + i) / 2), 5, OfficeOpenXml.Style.ExcelBorderStyle.None);
                        ePPlus.ExcelCenterCell(10 + ((i + i) / 2), 5, OfficeOpenXml.Style.ExcelHorizontalAlignment.Center);
                        //ePPlus.MergeColumn(9 + ((i + i) / 2), 6, 10 + ((i + i) / 2), 6);
                        ePPlus.FontSize(9 + ((i + i) / 2), 6, 13, false, OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                        ePPlus.ExcelCenterCell(9 + ((i + i) / 2), 6, OfficeOpenXml.Style.ExcelHorizontalAlignment.Center);
                        ePPlus.FontSize(10 + ((i + i) / 2), 6, 13, false, OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                        ePPlus.FontSizeTop(10 + ((i + i) / 2), 6, OfficeOpenXml.Style.ExcelBorderStyle.None);
                        ePPlus.ExcelCenterCell(10 + ((i + i) / 2), 6, OfficeOpenXml.Style.ExcelHorizontalAlignment.Center);
                        ePPlus.MergeColumn(9 + ((i + i) / 2), 7, 10 + ((i + i) / 2), 8);
                        ePPlus.FontSize(9 + ((i + i) / 2), 7, 13, false, OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                        ePPlus.ExcelCenterCell(9 + ((i + i) / 2), 7, OfficeOpenXml.Style.ExcelHorizontalAlignment.Center);
                    }
                }

                ePPlus.MergeColumn(1 + _X, 1, 2 + _X, 9);
                ePPlus.FontSize(1 + _X, 1, 36, true, OfficeOpenXml.Style.ExcelBorderStyle.None);
                ePPlus.ExcelCenterCell(1 + _X, 1, OfficeOpenXml.Style.ExcelHorizontalAlignment.Center);
                ePPlus.MergeColumn(3 + _X, 1, 3 + _X, 9);
                ePPlus.FontSize(3 + _X, 1, 14, false, OfficeOpenXml.Style.ExcelBorderStyle.None);
                ePPlus.ExcelCenterCell(3 + _X, 1, OfficeOpenXml.Style.ExcelHorizontalAlignment.Center);
                ePPlus.MergeColumn(4 + _X, 1, 4 + _X, 9);
                ePPlus.FontSize(4 + _X, 1, 14, false, OfficeOpenXml.Style.ExcelBorderStyle.None);
                ePPlus.ExcelCenterCell(4 + _X, 1, OfficeOpenXml.Style.ExcelHorizontalAlignment.Center);
                ePPlus.MergeColumn(5 + _X, 1, 6 + _X, 1);
                ePPlus.FontSize(5 + _X, 1, 13, false, OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                ePPlus.ExcelCenterCell(5 + _X, 1, OfficeOpenXml.Style.ExcelHorizontalAlignment.Center);
                ePPlus.MergeColumn(5 + _X, 2, 6 + _X, 3);
                ePPlus.FontSize(5 + _X, 2, 13, false, OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                ePPlus.ExcelCenterCell(5 + _X, 2, OfficeOpenXml.Style.ExcelHorizontalAlignment.Center);
                ePPlus.MergeColumn(5 + _X, 4, 6 + _X, 4);
                ePPlus.FontSize(5 + _X, 4, 13, false, OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                ePPlus.ExcelCenterCell(5 + _X, 4, OfficeOpenXml.Style.ExcelHorizontalAlignment.Center);
                ePPlus.MergeColumn(5 + _X, 5, 6 + _X, 9);
                ePPlus.FontSize(5 + _X, 5, 13, false, OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                ePPlus.ExcelCenterCell(5 + _X, 5, OfficeOpenXml.Style.ExcelHorizontalAlignment.Center);
                ePPlus.MergeColumn(7 + _X, 1, 8 + _X, 1);
                ePPlus.FontSize(7 + _X, 1, 13, false, OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                ePPlus.ExcelCenterCell(7 + _X, 1, OfficeOpenXml.Style.ExcelHorizontalAlignment.Center);
                ePPlus.MergeColumn(7 + _X, 2, 8 + _X, 3);
                ePPlus.FontSize(7 + _X, 2, 13, false, OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                ePPlus.ExcelCenterCell(7 + _X, 2, OfficeOpenXml.Style.ExcelHorizontalAlignment.Center);
                ePPlus.MergeColumn(7 + _X, 4, 8 + _X, 4);
                ePPlus.FontSize(7 + _X, 4, 13, false, OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                ePPlus.ExcelCenterCell(7 + _X, 4, OfficeOpenXml.Style.ExcelHorizontalAlignment.Center);
                ePPlus.MergeColumn(7 + _X, 5, 8 + _X, 6);
                ePPlus.FontSize(7 + _X, 5, 13, false, OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                ePPlus.ExcelCenterCell(7 + _X, 5, OfficeOpenXml.Style.ExcelHorizontalAlignment.Center);
                ePPlus.MergeColumn(7 + _X, 7, 8 + _X, 7);
                ePPlus.FontSize(7 + _X, 7, 13, false, OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                ePPlus.ExcelCenterCell(7 + _X, 7, OfficeOpenXml.Style.ExcelHorizontalAlignment.Center);
                ePPlus.MergeColumn(7 + _X, 8, 8 + _X, 9);
                ePPlus.FontSize(7 + _X, 8, 13, false, OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                ePPlus.ExcelCenterCell(7 + _X, 8, OfficeOpenXml.Style.ExcelHorizontalAlignment.Center);
                for (int i = 0; i < typeModels.Count; i = i + 2)
                {
                    //ePPlus.MergeColumn(9 + _X + ((i + i) / 2), 1, 10 + _X + ((i + i) / 2), 1);
                    ePPlus.FontSize(9 + _X + ((i + i) / 2), 1, 13, false, OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                    ePPlus.ExcelCenterCell(9 + _X + ((i + i) / 2), 1, OfficeOpenXml.Style.ExcelHorizontalAlignment.Center);
                    ePPlus.FontSize(10 + _X + ((i + i) / 2), 1, 13, false, OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                    ePPlus.FontSizeTop(10 + _X + ((i + i) / 2), 1, OfficeOpenXml.Style.ExcelBorderStyle.None);
                    ePPlus.ExcelCenterCell(10 + _X + ((i + i) / 2), 1, OfficeOpenXml.Style.ExcelHorizontalAlignment.Center);
                    //ePPlus.MergeColumn(9 + _X + ((i + i) / 2), 2, 10 + _X + ((i + i) / 2), 2);
                    ePPlus.FontSize(9 + _X + ((i + i) / 2), 2, 13, false, OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                    ePPlus.ExcelCenterCell(9 + _X + ((i + i) / 2), 2, OfficeOpenXml.Style.ExcelHorizontalAlignment.Center);
                    ePPlus.FontSize(10 + _X + ((i + i) / 2), 2, 13, false, OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                    ePPlus.FontSizeTop(10 + _X + ((i + i) / 2), 2, OfficeOpenXml.Style.ExcelBorderStyle.None);
                    ePPlus.ExcelCenterCell(10 + _X + ((i + i) / 2), 2, OfficeOpenXml.Style.ExcelHorizontalAlignment.Center);
                    ePPlus.MergeColumn(9 + _X + ((i + i) / 2), 3, 10 + _X + ((i + i) / 2), 4);
                    ePPlus.FontSize(9 + _X + ((i + i) / 2), 3, 13, false, OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                    ePPlus.ExcelCenterCell(9 + _X + ((i + i) / 2), 3, OfficeOpenXml.Style.ExcelHorizontalAlignment.Center);
                    if (i + 1 < typeModels.Count)
                    {
                        //ePPlus.MergeColumn(9 + _X + ((i + i) / 2), 5, 10 + _X + ((i + i) / 2), 5);
                        ePPlus.FontSize(9 + _X + ((i + i) / 2), 5, 13, false, OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                        ePPlus.ExcelCenterCell(9 + _X + ((i + i) / 2), 5, OfficeOpenXml.Style.ExcelHorizontalAlignment.Center);
                        ePPlus.FontSize(10 + _X + ((i + i) / 2), 5, 13, false, OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                        ePPlus.FontSizeTop(10 + _X + ((i + i) / 2), 5, OfficeOpenXml.Style.ExcelBorderStyle.None);
                        ePPlus.ExcelCenterCell(10 + _X + ((i + i) / 2), 5, OfficeOpenXml.Style.ExcelHorizontalAlignment.Center);
                        //ePPlus.MergeColumn(9 + _X + ((i + i) / 2), 6, 10 + _X + ((i + i) / 2), 6);
                        ePPlus.FontSize(9 + _X + ((i + i) / 2), 6, 13, false, OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                        ePPlus.ExcelCenterCell(9 + _X + ((i + i) / 2), 6, OfficeOpenXml.Style.ExcelHorizontalAlignment.Center);
                        ePPlus.FontSize(10 + _X + ((i + i) / 2), 6, 13, false, OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                        ePPlus.FontSizeTop(10 + _X + ((i + i) / 2), 6, OfficeOpenXml.Style.ExcelBorderStyle.None);
                        ePPlus.ExcelCenterCell(10 + _X + ((i + i) / 2), 6, OfficeOpenXml.Style.ExcelHorizontalAlignment.Center);
                        ePPlus.MergeColumn(9 + _X + ((i + i) / 2), 7, 10 + _X + ((i + i) / 2), 8);
                        ePPlus.FontSize(9 + _X + ((i + i) / 2), 7, 13, false, OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                        ePPlus.ExcelCenterCell(9 + _X + ((i + i) / 2), 7, OfficeOpenXml.Style.ExcelHorizontalAlignment.Center);
                    }
                }
                ePPlus.FontSize(excelCellsBuff.Count - 1, 8, 13, false, OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                ePPlus.FontSize(excelCellsBuff.Count - 1, 9, 13, false, OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                ePPlus.FontSize(ePPlus.EndCell - 1, 8, 13, false, OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                ePPlus.FontSize(ePPlus.EndCell - 1, 9, 13, false, OfficeOpenXml.Style.ExcelBorderStyle.Thin);

                ePPlus.MergeColumn(excelCellsBuff.Count, 1, excelCellsBuff.Count, 3);
                ePPlus.FontSize(excelCellsBuff.Count, 1, 11, false, OfficeOpenXml.Style.ExcelBorderStyle.None);
                ePPlus.MergeColumn(excelCellsBuff.Count, 4, excelCellsBuff.Count, 6);
                ePPlus.FontSize(excelCellsBuff.Count, 4, 13, false, OfficeOpenXml.Style.ExcelBorderStyle.None);
                ePPlus.ExcelCenterCell(excelCellsBuff.Count, 4, OfficeOpenXml.Style.ExcelHorizontalAlignment.Center, false);
                ePPlus.FontSize(excelCellsBuff.Count, 8, 13, false, OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                ePPlus.FontSize(excelCellsBuff.Count, 9, 13, false, OfficeOpenXml.Style.ExcelBorderStyle.Thin);

                ePPlus.MergeColumn(ePPlus.EndCell, 1, ePPlus.EndCell, 3);
                ePPlus.FontSize(ePPlus.EndCell, 1, 11, false, OfficeOpenXml.Style.ExcelBorderStyle.None);
                ePPlus.MergeColumn(ePPlus.EndCell, 4, ePPlus.EndCell, 6);
                ePPlus.FontSize(ePPlus.EndCell, 4, 13, false, OfficeOpenXml.Style.ExcelBorderStyle.None);
                ePPlus.ExcelCenterCell(ePPlus.EndCell, 4, OfficeOpenXml.Style.ExcelHorizontalAlignment.Center, false);
                ePPlus.FontSize(ePPlus.EndCell, 8, 13, false, OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                ePPlus.FontSize(ePPlus.EndCell, 9, 13, false, OfficeOpenXml.Style.ExcelBorderStyle.Thin);

                ePPlus.Export(Path);
                ePPlus.ChangeExcel2Image(Path, @".\ianimage.png");  //利用Spire將excel轉換成圖片

                log.LogMessage("匯出圖片 成功", enumLogType.Trace);
                log.LogMessage("匯出圖片 成功", enumLogType.Info);
                return true;
            }
            catch (Exception ee)
            {
                MessageBox.Show("匯出圖片 失敗：\r\n" + ee.Message);
                log.LogMessage("匯出圖片 失敗：\r\n" + ee.Message, enumLogType.Error);
                return false;
            }
        }

        public Boolean ExcelExportImage_2(DataGridView view, string Path, DateTime Now, string No, string Name, string Unit, string SalesArea, Boolean UnitPriceShow)
        {
            try
            {
                EPPlus ePPlus = new EPPlus();
                List<List<MExcelCell>> excelCells = new List<List<MExcelCell>>();
                List<MExcelCell> excelCell = new List<MExcelCell>();

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

                //標頭
                excelCell = new List<MExcelCell>();
                excelCell.Add(new MExcelCell() { Content = "單號" });
                excelCell.Add(new MExcelCell() { Content = No });
                excelCell.Add(new MExcelCell() { Content = " " });
                excelCell.Add(new MExcelCell() { Content = "姓名" });
                excelCell.Add(new MExcelCell() { Content = Name });
                excelCell.Add(new MExcelCell() { Content = " " });
                excelCell.Add(new MExcelCell() { Content = " " });
                excelCell.Add(new MExcelCell() { Content = " " });
                excelCell.Add(new MExcelCell() { Content = " " });
                excelCells.Add(excelCell);
                excelCell = new List<MExcelCell>();
                excelCell.Add(new MExcelCell() { Content = "日期" });
                excelCell.Add(new MExcelCell() { Content = Now.ToString("yyyy-MM-dd") });
                excelCell.Add(new MExcelCell() { Content = " " });
                excelCell.Add(new MExcelCell() { Content = "單位" });
                excelCell.Add(new MExcelCell() { Content = Unit });
                excelCell.Add(new MExcelCell() { Content = " " });
                excelCell.Add(new MExcelCell() { Content = "地區" });
                excelCell.Add(new MExcelCell() { Content = SalesArea });
                excelCell.Add(new MExcelCell() { Content = " " });
                excelCells.Add(excelCell);

                List<ALLTypeModel> typeModelsBuff = new List<ALLTypeModel>();
                List<ALLTypeModel> typeModels = new List<ALLTypeModel>();

                int _TypeIndex = -1;
                int _CountIndex = -1;
                foreach (DataGridViewColumn column in view.Columns)
                {
                    if (column.HeaderText == "類型")
                        _TypeIndex = column.Index;
                    if (column.HeaderText == "重量")
                        _CountIndex = column.Index;
                    if (_TypeIndex != -1 && _CountIndex != -1)
                        break;
                }
                foreach (DataGridViewRow row in view.Rows)
                {
                    //類型匯入
                    string _Type = row.Cells[_TypeIndex].Value.ToString();
                    ALLTypeModel typeModel = typeModelsBuff.Find(f => f.Type == _Type);
                    if (typeModel == null)
                    {
                        typeModel = new ALLTypeModel() { Type = _Type };
                        typeModelsBuff.Add(typeModel);
                    }
                    //單筆重量加總
                    typeModel._ALLCount += Convert.ToDouble(row.Cells[_CountIndex].Value.ToString());
                }
                //排序
                foreach (string item in ManagerView.TypeGradation())
                {
                    ALLTypeModel typeModel = typeModelsBuff.Find(f => f.Type == item);
                    if (typeModel != null)
                    {
                        typeModels.Add(typeModel);
                    }
                }
                for (int i = 0; i < typeModels.Count; i = i + 3)
                {
                    excelCell = new List<MExcelCell>();
                    excelCell.Add(new MExcelCell() { Content = typeModels[i].Type.ToString() });
                    excelCell.Add(new MExcelCell() { Content = typeModels[i]._ALLCount.ToString() });
                    excelCell.Add(new MExcelCell() { Content = " " });
                    if (i + 1 < typeModels.Count)
                    {
                        excelCell.Add(new MExcelCell() { Content = typeModels[i + 1].Type.ToString() });
                        excelCell.Add(new MExcelCell() { Content = typeModels[i + 1]._ALLCount.ToString() });
                        excelCell.Add(new MExcelCell() { Content = " " });
                    }
                    if (i + 2 < typeModels.Count)
                    {
                        excelCell.Add(new MExcelCell() { Content = typeModels[i + 2].Type.ToString() });
                        excelCell.Add(new MExcelCell() { Content = typeModels[i + 2]._ALLCount.ToString() });
                        excelCell.Add(new MExcelCell() { Content = " " });
                    }
                    excelCells.Add(excelCell);
                }
                //空一行
                excelCells.Add(new List<MExcelCell>());


                //頁首
                List<string> _HideHeader = new List<string>() { "單號", "時間", "姓名", "單位", "地區", "已付款金額", "未付款金額", "已付時間" };
                excelCell = new List<MExcelCell>();
                foreach (DataGridViewColumn col in view.Columns)
                {
                    //隱藏
                    if (_HideHeader.Contains(col.HeaderText))
                    {
                        continue;
                    }
                    //列印隱藏單價
                    if (col.HeaderText == "單價" && !UnitPriceShow)
                    {
                        excelCell.Add(new MExcelCell());
                        continue;
                    }
                    excelCell.Add(new MExcelCell()
                    {
                        Content = col.HeaderText
                    });
                }
                //列印顯示價格
                if (UnitPriceShow)
                {
                    excelCell.Add(new MExcelCell()
                    {
                        Content = "價格"
                    });
                }
                else
                    excelCell.Add(new MExcelCell());
                excelCell.Add(new MExcelCell());
                foreach (DataGridViewColumn col in view.Columns)
                {
                    //隱藏
                    if (_HideHeader.Contains(col.HeaderText))
                    {
                        continue;
                    }
                    //列印隱藏單價
                    if (col.HeaderText == "單價" && !UnitPriceShow)
                    {
                        excelCell.Add(new MExcelCell());
                        continue;
                    }
                    excelCell.Add(new MExcelCell()
                    {
                        Content = col.HeaderText
                    });
                }
                //列印顯示價格
                if (UnitPriceShow)
                {
                    excelCell.Add(new MExcelCell()
                    {
                        Content = "價格"
                    });
                }
                else
                    excelCell.Add(new MExcelCell());
                excelCells.Add(excelCell);

                //內容
                Int32 _ALLPrice = 0;
                for (int i = 0; i < view.Rows.Count; i = i + 2)
                {
                    Double _unitPrice = 0;
                    Double _count = 1;
                    excelCell = new List<MExcelCell>();
                    foreach (DataGridViewCell cell in view.Rows[i].Cells)
                    {
                        //隱藏
                        if (_HideHeader.Contains(view.Columns[cell.ColumnIndex].HeaderText))
                        {
                            continue;
                        }
                        //列印隱藏單價/保存單價價格
                        if (view.Columns[cell.ColumnIndex].HeaderText == "單價")
                        {
                            if (!UnitPriceShow)
                            {
                                excelCell.Add(new MExcelCell());
                                continue;
                            }
                            _unitPrice = Convert.ToDouble(cell.Value);
                        }
                        //保存重量
                        else if (view.Columns[cell.ColumnIndex].HeaderText == "重量")
                        {
                            _count = Convert.ToDouble(cell.Value);
                        }
                        excelCell.Add(new MExcelCell()
                        {
                            Content = cell.Value
                        });
                    }
                    //價格加總
                    if (UnitPriceShow)
                    {
                        _ALLPrice += (int)Math.Round(Convert.ToDouble(_unitPrice * _count), 0, MidpointRounding.AwayFromZero);
                        excelCell.Add(new MExcelCell()
                        {
                            Content = (int)Math.Round(Convert.ToDouble(_unitPrice * _count), 0, MidpointRounding.AwayFromZero)
                        });
                    }
                    else
                        excelCell.Add(new MExcelCell());

                    excelCell.Add(new MExcelCell());
                    if (i + 1 < view.Rows.Count)
                    {
                        foreach (DataGridViewCell cell in view.Rows[i + 1].Cells)
                        {
                            //隱藏
                            if (_HideHeader.Contains(view.Columns[cell.ColumnIndex].HeaderText))
                            {
                                continue;
                            }
                            //列印隱藏單價/保存單價價格
                            if (view.Columns[cell.ColumnIndex].HeaderText == "單價")
                            {
                                if (!UnitPriceShow)
                                {
                                    excelCell.Add(new MExcelCell());
                                    continue;
                                }
                                _unitPrice = Convert.ToDouble(cell.Value);
                            }
                            //保存重量
                            else if (view.Columns[cell.ColumnIndex].HeaderText == "重量")
                            {
                                _count = Convert.ToDouble(cell.Value);
                            }
                            excelCell.Add(new MExcelCell()
                            {
                                Content = cell.Value
                            });
                        }

                        //價格加總
                        if (UnitPriceShow)
                        {
                            _ALLPrice += (int)Math.Round(Convert.ToDouble(_unitPrice * _count), 0, MidpointRounding.AwayFromZero);
                            excelCell.Add(new MExcelCell()
                            {
                                Content = (int)Math.Round(Convert.ToDouble(_unitPrice * _count), 0, MidpointRounding.AwayFromZero)
                            });
                        }
                        else
                            excelCell.Add(new MExcelCell());
                    }
                    excelCells.Add(excelCell);
                }
                //空一行
                excelCells.Add(new List<MExcelCell>());
                excelCell = new List<MExcelCell>();
                excelCell.Add(new MExcelCell() { Content = "民智自動化有限公司" });
                //總價
                if (UnitPriceShow)
                {
                    for (int i = 0; i < view.Columns.Count; i++)
                    {
                        //隱藏
                        if (_HideHeader.Contains(view.Columns[i].HeaderText))
                        {
                            continue;
                        }
                        excelCell.Add(new MExcelCell() { Content = "" });
                        excelCell.Add(new MExcelCell() { Content = "" });
                    }
                    //excelCell.Remove(excelCell[excelCell.Count - 1]);
                    excelCell.Add(new MExcelCell() { Content = "總價" });
                    excelCell.Add(new MExcelCell() { Content = _ALLPrice });
                }
                excelCells.Add(excelCell);

                //匯出成檔案
                ePPlus.AddSheet(excelCells, No, 0, 14);
                ePPlus.MergeColumn(1, 1, 2, 9);
                ePPlus.FontSize(1, 1, 36, true, OfficeOpenXml.Style.ExcelBorderStyle.None);
                ePPlus.ExcelCenterCell(1, 1, OfficeOpenXml.Style.ExcelHorizontalAlignment.CenterContinuous);
                ePPlus.MergeColumn(3, 1, 3, 9);
                ePPlus.FontSize(3, 1, 14, false, OfficeOpenXml.Style.ExcelBorderStyle.None);
                ePPlus.ExcelCenterCell(3, 1, OfficeOpenXml.Style.ExcelHorizontalAlignment.CenterContinuous);
                ePPlus.MergeColumn(4, 1, 4, 9);
                ePPlus.FontSize(4, 1, 14, false, OfficeOpenXml.Style.ExcelBorderStyle.None);
                ePPlus.ExcelCenterCell(4, 1, OfficeOpenXml.Style.ExcelHorizontalAlignment.CenterContinuous);
                ePPlus.MergeColumn(6, 2, 6, 3);
                ePPlus.MergeColumn(6, 5, 6, 9);
                ePPlus.MergeColumn(7, 2, 7, 3);
                ePPlus.MergeColumn(7, 5, 7, 6);
                ePPlus.MergeColumn(7, 8, 7, 9);
                for (int i = 0; i < typeModels.Count; i = i + 3)
                {
                    ePPlus.MergeColumn(8 + (i / 3), 2, 8 + (i / 3), 3);
                    if (i + 1 < typeModels.Count)
                    {
                        ePPlus.MergeColumn(8 + (i / 3), 5, 8 + (i / 3), 6);
                    }
                    if (i + 2 < typeModels.Count)
                    {
                        ePPlus.MergeColumn(8 + (i / 3), 8, 8 + (i / 3), 9);
                    }
                }
                ePPlus.MergeColumn(ePPlus.EndCell, 1, ePPlus.EndCell, 3);
                ePPlus.FontSize(ePPlus.EndCell, 1, 11, false, OfficeOpenXml.Style.ExcelBorderStyle.None);
                ePPlus.Export(Path);
                ePPlus.ChangeExcel2Image(Path, @".\ianimage.png");  //利用Spire將excel轉換成圖片

                log.LogMessage("匯出圖片 成功", enumLogType.Trace);
                log.LogMessage("匯出圖片 成功", enumLogType.Info);
                return true;
            }
            catch (Exception ee)
            {
                MessageBox.Show("匯出圖片 失敗：\r\n" + ee.Message);
                log.LogMessage("匯出圖片 失敗：\r\n" + ee.Message, enumLogType.Error);
                return false;
            }
        }
    }
}
