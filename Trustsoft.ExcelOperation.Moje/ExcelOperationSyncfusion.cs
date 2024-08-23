using Markdig.Extensions.Bootstrap;
using Microsoft.CodeAnalysis.VisualBasic.Syntax;
using NPOI.HSSF.Record.CF;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using SixLabors.ImageSharp;
using Soneta.Business;
using Soneta.Kasa.BialaListaApi.Model;
using Soneta.Kasa.Extensions;
using Soneta.Types;
using Syncfusion.Drawing;
using Syncfusion.XlsIO;
using System;
using System.Drawing;
using Trustsoft.ExcelOperation.Moje.SyncfusionException;
using static Soneta.Kalend.DefinicjaZestawieniaCzasu;
using IDataValidation = Syncfusion.XlsIO.IDataValidation;
using IName = Syncfusion.XlsIO.IName;

namespace Trustsoft.ExcelOperation.Moje
{
    public class ExcelOperationSyncfusion : IExcelOperationService, IDisposable
    {
        private ExcelEngine _excelEngine;
        private IApplication _application;
        private Syncfusion.XlsIO.IWorkbook _workbook;

        ///<summary>
        /// Default constructor. DefaultVersion = Xlsx.
        /// </summary>
        public ExcelOperationSyncfusion()
        {
            _excelEngine = new ExcelEngine();
            _application = _excelEngine.Excel;
            _application.DefaultVersion = ExcelVersion.Xlsx;
        }

        ///<summary>
        ///Constructor passing the excel version.
        /// </summary>
        public ExcelOperationSyncfusion(ExcelVersion excelVersion)
        {
            _excelEngine = new ExcelEngine();
            _application = _excelEngine.Excel;
            _application.DefaultVersion = excelVersion;
        }

        public void Dispose()
        {
            _workbook.Close();
            _excelEngine.Dispose();
        }

        public object CreateWorkbook()
        {
            if (_excelEngine == null && _application == null)
            {
                throw new SyncfusionNullApplicationException();
            }
            _workbook = _application.Workbooks.Create(1);
            _workbook.Worksheets[0].Name = "Sheet1";
            return _workbook;
        }
        public object CreateWorkbook(string sheetName)
        {
            if (_excelEngine == null && _application == null)
            {
                throw new SyncfusionNullApplicationException();
            }
            _workbook = _application.Workbooks.Create(1);
            _workbook.Worksheets[0].Name = sheetName;
            return _workbook;
        }

        public List<Sheet> GetNameSheet()
        {
            List<Sheet> sheets = new List<Sheet>();
            int sheetIndex = 0;

            foreach (var sheet in _workbook.Worksheets)
            {
                sheets.Add(new Sheet(sheetIndex, sheet.Name));
                sheetIndex++;
            }
            return sheets;
        }

        public int AddWorksheet(string sheetName)
        {
            var sheet = _workbook.Worksheets.Create(sheetName);
            return sheet.Index;
        }

        public void ChangeNameWorksheet(string sheetName, string newName)
        {
            _workbook.Worksheets[sheetName].Name = newName;
        }

        public void ChangeNameWorksheet(int indexsheet, string newName)
        {
            _workbook.Worksheets[indexsheet].Name = newName;
        }
        
        public void DeleteWorksheet(string sheetName)
        {
            _workbook.Worksheets.Remove(sheetName);
        }

        public void DeleteWorksheet(int sheetIndex)
        {
            _workbook.Worksheets.Remove(sheetIndex);
        }

        public void AddRow(int rowIndex, string sheetName)
        {
            _workbook.Worksheets[sheetName].InsertRow(rowIndex + 1);
        }

        public void AddRow(int rowIndex, string sheetName, int rowCount)
        {
            _workbook.Worksheets[sheetName].InsertRow(rowIndex + 1, rowCount);
        }

        public void AddRow(string sheetName, int rowIndex, int rowCount)
        {
            _workbook.Worksheets[sheetName].InsertRow(rowIndex + 1, rowCount);
        }

        public void AddRow(int sheetIndex, int rowIndex, int rowCount)
        {
            _workbook.Worksheets[sheetIndex].InsertRow(rowIndex + 1, rowCount);
        }

        public void AddRow(string sheetName, int rowIndex)
        {
            _workbook.Worksheets[sheetName].InsertRow(rowIndex + 1);
        }

        public void AddColumn(int columnIndex, string sheetName)
        {
            _workbook.Worksheets[sheetName].InsertColumn(columnIndex + 1);
        }

        public void AddColumn(string sheetName, int columnIndex)
        {
            _workbook.Worksheets[sheetName].InsertColumn(columnIndex + 1);
        }

        public void AddRow(int sheetIndex, int rowIndex)
        {
            _workbook.Worksheets[sheetIndex].InsertRow(rowIndex + 1);
        }

        public void AddColumn(int sheetIndex, int columnIndex)
        {
            _workbook.Worksheets[sheetIndex].InsertColumn(columnIndex + 1);
        }

        public void AddColumn(int columnIndex, string sheetName, int columnCount)
        {
            _workbook.Worksheets[sheetName].InsertColumn(columnIndex + 1, columnCount);
        }

        public void AddColumn(string sheetName, int columnIndex, int columnCount)
        {
            _workbook.Worksheets[sheetName].InsertColumn(columnIndex + 1, columnCount);
        }

        public void AddColumn(int sheetIndex, int columnIndex, int columnCount)
        {
            _workbook.Worksheets[sheetIndex].InsertColumn(columnIndex + 1, columnCount);
        }

        public void DeleteRow(string sheetName, int rowIndex)
        {
            _workbook.Worksheets[sheetName].DeleteRow(rowIndex + 1);
        }

        public void DeleteRow(int rowIndex, string sheetName)
        {
            _workbook.Worksheets[sheetName].DeleteRow(rowIndex + 1);
        }

        public void DeleteRow(int sheetIndex, int rowIndex)
        {
            _workbook.Worksheets[sheetIndex].DeleteRow(rowIndex + 1);
        }

        public void DeleteRow(string sheetName, int rowIndex, int rowCount)
        {
            _workbook.Worksheets[sheetName].DeleteRow(rowIndex + 1, rowCount);
        }
        
        public void DeleteRow(int rowIndex, string sheetName, int rowCount)
        {
            _workbook.Worksheets[sheetName].DeleteRow(rowIndex + 1, rowCount);
        }
        
        public void DeleteRow(int sheetIndex, int rowIndex, int rowCount)
        {
            _workbook.Worksheets[sheetIndex].DeleteRow(rowIndex + 1, rowCount);
        }

        public void DeleteColumn(string sheetName, int columnIndex)
        {
            _workbook.Worksheets[sheetName].DeleteColumn(columnIndex + 1);
        }

        public void DeleteColumn(int columnIndex, string sheetName)
        {
            _workbook.Worksheets[sheetName].DeleteColumn(columnIndex + 1);
        }

        public void DeleteColumn(int sheetIndex, int columnIndex)
        {
            _workbook.Worksheets[sheetIndex].DeleteColumn(columnIndex + 1);
        }

        public void DeleteColumn(string sheetName, int columnIndex, int columnCount)
        {
            _workbook.Worksheets[sheetName].DeleteColumn(columnIndex + 1, columnCount);
        }
        
        public void DeleteColumn(int columnIndex, string sheetName, int columnCount)
        {
            _workbook.Worksheets[sheetName].DeleteColumn(columnIndex + 1, columnCount);
        }
        
        public void DeleteColumn(int sheetIndex, int columnIndex, int columnCount)
        {
            _workbook.Worksheets[sheetIndex].DeleteColumn(columnIndex + 1, columnCount);
        }

        public void AddCellValueText(int sheetIndex, int rowIndex, int columnIndex, string text)
        {
            _workbook.Worksheets[sheetIndex][rowIndex + 1, columnIndex + 1].Text = text;
            _workbook.Worksheets[sheetIndex][rowIndex + 1, columnIndex + 1].NumberFormat = "@";
        }

        public void AddCellValueText(string sheetName, int rowIndex, int columnIndex, string text)
        {
            _workbook.Worksheets[sheetName][rowIndex + 1, columnIndex + 1].Text = text;
            _workbook.Worksheets[sheetName][rowIndex + 1, columnIndex + 1].NumberFormat = "@";
        }

        public void AddCellFormula(int sheetIndex, int rowIndex, int columnIndex, string formula)
        {
            _workbook.Worksheets[sheetIndex][rowIndex + 1, columnIndex + 1].Formula = $"={formula}";
        }

        public void AddCellFormula(string sheetName, int rowIndex, int columnIndex, string formula)
        {
            _workbook.Worksheets[sheetName][rowIndex + 1, columnIndex + 1].Formula = $"={formula}";
        }

        public void AddCellFormula(int sheetIndex, int firstRowIndex, int firstColumnIndex, int lastRowIndex, int lastCoulmnIndex, string formula)
        {
            _workbook.Worksheets[sheetIndex][firstRowIndex + 1, firstColumnIndex + 1, lastRowIndex + 1, lastCoulmnIndex + 1].Formula = $"={formula}";
        }

        public void AddCellFormula(string sheetName, int firstRowIndex, int firstColumnIndex, int lastRowIndex, int lastCoulmnIndex, string formula)
        {
            _workbook.Worksheets[sheetName][firstRowIndex + 1, firstColumnIndex + 1, lastRowIndex + 1, lastCoulmnIndex + 1].Formula = $"={formula}";
        }

        public void AddCellValueInt(int sheetIndex, int rowIndex, int columnIndex, int value)
        {
            _workbook.Worksheets[sheetIndex][rowIndex + 1, columnIndex + 1].Number = value;
        }

        public void AddCellValueInt(string sheetName, int rowIndex, int columnIndex, int value)
        {
            _workbook.Worksheets[sheetName][rowIndex + 1, columnIndex + 1].Number = value;
        }

        public void AddCellValueInt(int sheetIndex, int rowIndex, int columnIndex, int value, string format)
        {
            _workbook.Worksheets[sheetIndex][rowIndex + 1, columnIndex + 1].Number = value;
            _workbook.Worksheets[sheetIndex][rowIndex + 1, columnIndex + 1].NumberFormat = format;
        }
        public void AddCellValueInt(string sheetName, int rowIndex, int columnIndex, int value, string format)
        {
            _workbook.Worksheets[sheetName][rowIndex + 1, columnIndex + 1].Number = value;
            _workbook.Worksheets[sheetName][rowIndex + 1, columnIndex + 1].NumberFormat = format;
        }

        public void AddCellValueDecimal(int sheetIndex, int rowIndex, int columnIndex, decimal value)
        {
            _workbook.Worksheets[sheetIndex][rowIndex + 1, columnIndex + 1].Number = (double)value;
        }

        public void AddCellValueDecimal(string sheetName, int rowIndex, int columnIndex, decimal value)
        {
            _workbook.Worksheets[sheetName][rowIndex + 1, columnIndex + 1].Number = (double)value;
        }

        public void AddCellValueDecimal(int sheetIndex, int rowIndex, int columnIndex, decimal value, string format)
        {
            _workbook.Worksheets[sheetIndex][rowIndex + 1, columnIndex + 1].Number = (double)value;
            _workbook.Worksheets[sheetIndex][rowIndex + 1, columnIndex + 1].NumberFormat = format;
        }
        public void AddCellValueDecimal(string sheetName, int rowIndex, int columnIndex, decimal value, string format)
        {
            _workbook.Worksheets[sheetName][rowIndex + 1, columnIndex + 1].Number = (double)value;
            _workbook.Worksheets[sheetName][rowIndex + 1, columnIndex + 1].NumberFormat = format;
        }

        public void AddCellValueDouble(int sheetIndex, int rowIndex, int columnIndex, double value)
        {
            _workbook.Worksheets[sheetIndex][rowIndex + 1, columnIndex + 1].Number = value;
        }

        public void AddCellValueDouble(string sheetName, int rowIndex, int columnIndex, double value)
        {
            _workbook.Worksheets[sheetName][rowIndex + 1, columnIndex + 1].Number = value;
        }

        public void AddCellValueDouble(int sheetIndex, int rowIndex, int columnIndex, double value, string format)
        {
            _workbook.Worksheets[sheetIndex][rowIndex + 1, columnIndex + 1].Number = value;
            _workbook.Worksheets[sheetIndex][rowIndex + 1, columnIndex + 1].NumberFormat = format;
        }

        public void AddCellValueDouble(string sheetName, int rowIndex, int columnIndex, double value, string format)
        {
            _workbook.Worksheets[sheetName][rowIndex + 1, columnIndex + 1].Number = value;
            _workbook.Worksheets[sheetName][rowIndex + 1, columnIndex + 1].NumberFormat = format;
        }

        public void HeightRow(int sheetIndex, int rowIndex, double height)
        {
            _workbook.Worksheets[sheetIndex].SetRowHeight(rowIndex + 1, height);
        }
        public void HeightRow(int sheetIndex, int rowIndex, int rowCount, double height)
        {
            for (int i = 0; i < rowCount; i++)
            {
                _workbook.Worksheets[sheetIndex].Rows[rowIndex + i].RowHeight = height;
            }
        }
        public void HeightRow(int sheetIndex, int[] rowIndices, double height)
        {
            foreach (var index in rowIndices)
            {
                _workbook.Worksheets[sheetIndex].Rows[index].RowHeight = height;
            }
        }

        public void HeightRow(string sheetName, int rowIndex, int rowCount, double height)
        {
            for (int i = 0; i < rowCount; i++)
            {
                _workbook.Worksheets[sheetName].Rows[rowIndex + i].RowHeight = height;
            }
        }

        public void HeightRow(string sheetName, int[] rowIndices, double height)
        {
            foreach (var index in rowIndices)
            {
                _workbook.Worksheets[sheetName].Rows[index].RowHeight = height;
            }
        }

        public void HeightRow(string sheetName, int rowIndex, double height)
        {
            _workbook.Worksheets[sheetName].SetRowHeight(rowIndex + 1, height);
        }

        public void WidthColumn(int sheetIndex, int columnIndex, double width)
        {
            _workbook.Worksheets[sheetIndex].SetColumnWidth(columnIndex + 1, width);
        }

        public void WidthColumn(int sheetIndex, int columnIndex, int columnCount, double width)
        {
            for (int i = 0; i < columnCount; i++)
            {
                _workbook.Worksheets[sheetIndex].Columns[columnIndex + i].ColumnWidth = width;
            }
        }

        public void WidthColumn(int sheetIndex, int[] columnIndices, double width)
        {
            foreach (var index in columnIndices)
            {
                _workbook.Worksheets[sheetIndex].Columns[index].ColumnWidth = width;
            }
        }

        public void WidthColumn(string sheetName, int columnIndex, int columnCount, double width)
        {
            for (int i = 0; i < columnCount; i++)
            {
                _workbook.Worksheets[sheetName].Columns[columnIndex + i].ColumnWidth = width;
            }
        }

        public void WidthColumn(string sheetName, int[] columnIndices, double width)
        {
            foreach (var index in columnIndices)
            {
                _workbook.Worksheets[sheetName].Columns[index].ColumnWidth = width;
            }
        }

        public void WidthColumn(string sheetName, int columnIndex, double width)
        {
            _workbook.Worksheets[sheetName].SetColumnWidth(columnIndex + 1, width);
        }

        public void AddCellValueDate(int sheetIndex, int rowIndex, int columnIndex, DateTime date)
        {
            _workbook.Worksheets[sheetIndex][rowIndex + 1, columnIndex + 1].DateTime = date;
            _workbook.Worksheets[sheetIndex][rowIndex + 1, columnIndex + 1].NumberFormat = "dd.mm.yyyy";
        }

        public void AddCellValueDate(string sheetName, int rowIndex, int columnIndex, DateTime date)
        {
            _workbook.Worksheets[sheetName][rowIndex + 1, columnIndex + 1].DateTime = date;
            _workbook.Worksheets[sheetName][rowIndex + 1, columnIndex + 1].NumberFormat = "dd.mm.yyyy";
        }

        public void AddCellValueDate(int sheetIndex, int rowIndex, int columnIndex, DateTime date, string format)
        {
            _workbook.Worksheets[sheetIndex][rowIndex + 1, columnIndex + 1].DateTime = date;
            _workbook.Worksheets[sheetIndex][rowIndex + 1, columnIndex + 1].NumberFormat = format;
        }

        public void AddCellValueDate(string sheetName, int rowIndex, int columnIndex, DateTime date, string format)
        {
            _workbook.Worksheets[sheetName][rowIndex + 1, columnIndex + 1].DateTime = date;
            _workbook.Worksheets[sheetName][rowIndex + 1, columnIndex + 1].NumberFormat = format;
        }

        public void AddCellValueCurrency(int sheetIndex, int rowIndex, int columnIndex, Currency currency) 
        {
            _workbook.Worksheets[sheetIndex][rowIndex + 1, columnIndex + 1].Text = currency.ToString();
        }

        public void AddCellValueCurrency(string sheetName, int rowIndex, int columnIndex, Currency currency)
        {
            _workbook.Worksheets[sheetName][rowIndex + 1, columnIndex + 1].Text = currency.ToString();
        }

        public void AddCellValuePercent(int sheetIndex, int rowIndex, int columnIndex, Percent percent)
        {
            _workbook.Worksheets[sheetIndex][rowIndex + 1, columnIndex + 1].Number = (double)percent;
            _workbook.Worksheets[sheetIndex][rowIndex + 1, columnIndex + 1].NumberFormat = "0.00%";
        }

        public void AddCellValuePercent(string sheetName, int rowIndex, int columnIndex, Percent percent)
        {
            _workbook.Worksheets[sheetName][rowIndex + 1, columnIndex + 1].Number = (double)percent;
            _workbook.Worksheets[sheetName][rowIndex + 1, columnIndex + 1].NumberFormat = "0.00%";
        }

        public void AddCellValuePercent(int sheetIndex, int rowIndex, int columnIndex, Percent percent, string format)
        {
            _workbook.Worksheets[sheetIndex][rowIndex + 1, columnIndex + 1].Number = (double)percent;
            _workbook.Worksheets[sheetIndex][rowIndex + 1, columnIndex + 1].NumberFormat = format;
        }

        public void AddCellValuePercent(string sheetName, int rowIndex, int columnIndex, Percent percent, string format)
        {
            _workbook.Worksheets[sheetName][rowIndex + 1, columnIndex + 1].Number = (double)percent;
            _workbook.Worksheets[sheetName][rowIndex + 1, columnIndex + 1].NumberFormat = format;
        }

        public void AddCellValueTime(int sheetIndex, int rowIndex, int columnIndex, Time time)
        {
            _workbook.Worksheets[sheetIndex][rowIndex + 1, columnIndex + 1].Text = time.ToString();
        }

        public void AddCellValueTime(string sheetName, int rowIndex, int columnIndex, Time time)
        {
            _workbook.Worksheets[sheetName][rowIndex + 1, columnIndex + 1].Text = time.ToString();
        }

        public void AddCellValueFraction(int sheetIndex, int rowIndex, int columnIndex, Fraction fraction)
        {
            _workbook.Worksheets[sheetIndex][rowIndex + 1, columnIndex + 1].Number = fraction;
        }

        public void AddCellValueFraction(string sheetName, int rowIndex, int columnIndex, Fraction fraction)
        {
            _workbook.Worksheets[sheetName][rowIndex + 1, columnIndex + 1].Number = fraction;
        }

        public void AddCellValueFraction(int sheetIndex, int rowIndex, int columnIndex, Fraction fraction, string format)
        {
            _workbook.Worksheets[sheetIndex][rowIndex + 1, columnIndex + 1].Number = fraction;
            _workbook.Worksheets[sheetIndex][rowIndex + 1, columnIndex + 1].NumberFormat = format;
        }

        public void AddCellValueFraction(string sheetName, int rowIndex, int columnIndex, Fraction fraction, string format)
        {
            _workbook.Worksheets[sheetName][rowIndex + 1, columnIndex + 1].Number = fraction;
            _workbook.Worksheets[sheetName][rowIndex + 1, columnIndex + 1].NumberFormat = format;
        }

        public void CellColor(int sheetIndex, int rowIndex, int columnIndex, int a, int r, int g, int b)
        {
            IRange cell = _workbook.Worksheets[sheetIndex][rowIndex + 1, columnIndex + 1];
            string nameSheet = _workbook.Worksheets[sheetIndex].Name;
            IStyle style = cell.CellStyle ?? _workbook.Styles.Add($"style{nameSheet}{rowIndex + 1}{columnIndex + 1}");

            Syncfusion.Drawing.Color c = Syncfusion.Drawing.Color.FromArgb(a, r, g, b);
            style.Color = c;

            cell.CellStyle = style;
        }

        public void CellColor(string sheetName, int rowIndex, int columnIndex, int a, int r, int g, int b)
        {
            IRange cell = _workbook.Worksheets[sheetName][rowIndex + 1, columnIndex + 1];
            IStyle style = cell.CellStyle ?? _workbook.Styles.Add($"style{sheetName}{rowIndex + 1}{columnIndex + 1}");

            Syncfusion.Drawing.Color c = Syncfusion.Drawing.Color.FromArgb(a, r, g, b);
            style.Color = c;

            cell.CellStyle = style;
        }

        public void CellColor(int sheetIndex, int firstRowIndex, int firstColumnIndex, int lastRowIndex, int lastColumnIndex, int a, int r, int g, int b)
        {
            for (int row = firstRowIndex + 1; row <= lastRowIndex + 1; row++)
            {
                for (int col = firstColumnIndex + 1; col <= lastColumnIndex + 1; col++)
                {
                    IRange cell = _workbook.Worksheets[sheetIndex][row, col];
                    string nameSheet = _workbook.Worksheets[sheetIndex].Name;
                    IStyle style = cell.CellStyle ?? _workbook.Styles.Add($"style{nameSheet}{row}{col}");

                    Syncfusion.Drawing.Color c = Syncfusion.Drawing.Color.FromArgb(a, r, g, b);
                    style.Color = c;

                    cell.CellStyle = style;
                }
            }
        }

        public void CellColor(string sheetName, int firstRowIndex, int firstColumnIndex, int lastRowIndex, int lastColumnIndex, int a, int r, int g, int b)
        {
            for (int row = firstRowIndex + 1; row <= lastRowIndex + 1; row++)
            {
                for (int col = firstColumnIndex + 1; col <= lastColumnIndex + 1; col++)
                {
                    IRange cell = _workbook.Worksheets[sheetName][row, col];
                    IStyle style = cell.CellStyle ?? _workbook.Styles.Add($"style{sheetName}{row}{col}");

                    Syncfusion.Drawing.Color c = Syncfusion.Drawing.Color.FromArgb(a, r, g, b);
                    style.Color = c;

                    cell.CellStyle = style;
                }
            }
        }

        public void SetBorder(BorderIndex[] borderIndex, LinesIndex lineIndex, string nameSheet, int rowIndex, int columnIndex)
        {
            IRange cell = _workbook.Worksheets[nameSheet][rowIndex + 1, columnIndex + 1];
            IStyle style = cell.CellStyle ?? _workbook.Styles.Add($"style{nameSheet}{rowIndex + 1}{columnIndex + 1}");

            foreach (var itemBorder in borderIndex)
            {
                var syncBorderIndex = SyncfusionHelper.ConvertFromBordexIndexSyncfusion(itemBorder, out bool isEmpty);
                if (!isEmpty)
                {
                    var syncLinesIndex = SyncfusionHelper.ConvertFromLineStyleSyncfusion(lineIndex, out bool isEmptyLine);
                    if (!isEmptyLine)
                    {
                        foreach (ExcelBordersIndex excelBorderIndex in syncBorderIndex)
                        {
                            style.Borders[excelBorderIndex].ColorRGB = Syncfusion.Drawing.Color.FromArgb(255, 0, 0, 0);
                            style.Borders[excelBorderIndex].LineStyle = syncLinesIndex;
                        }
                    }
                }
            }

            cell.CellStyle = style;
        }

        public void SetBorder(BorderIndex[] borderIndex, LinesIndex lineIndex, int sheetIndex, int rowIndex, int columnIndex)
        {
            IRange cell = _workbook.Worksheets[sheetIndex][rowIndex + 1, columnIndex + 1];
            string nameSheet = _workbook.Worksheets[sheetIndex].Name;
            IStyle style = cell.CellStyle ?? _workbook.Styles.Add($"style{nameSheet}{rowIndex + 1}{columnIndex + 1}");

            foreach (var itemBorder in borderIndex)
            {
                var syncBorderIndex = SyncfusionHelper.ConvertFromBordexIndexSyncfusion(itemBorder, out bool isEmpty);
                if (!isEmpty)
                {
                    var syncLinesIndex = SyncfusionHelper.ConvertFromLineStyleSyncfusion(lineIndex, out bool isEmptyLine);
                    if (!isEmptyLine)
                    {
                        foreach (ExcelBordersIndex excelBorderIndex in syncBorderIndex)
                        {
                            style.Borders[excelBorderIndex].ColorRGB = Syncfusion.Drawing.Color.FromArgb(255, 0, 0, 0);
                            style.Borders[excelBorderIndex].LineStyle = syncLinesIndex;
                        }
                    }
                }
            }

            cell.CellStyle = style;
        }


        //Borders default black

        public void SetBorder(BorderIndex[] borderIndex, LinesIndex lineIndex, string nameSheet, int firstRowIndex, int firstColumnIndex, int lastRowIndex, int lastColumnIndex, bool allRange)
        {
            if (!allRange)
            {
                for (int row = firstRowIndex + 1; row <= lastRowIndex + 1; row++)
                {
                    for (int col = firstColumnIndex + 1; col <= lastColumnIndex + 1; col++)
                    {
                        IRange cell = _workbook.Worksheets[nameSheet][row, col];
                        IStyle style = cell.CellStyle ?? _workbook.Styles.Add($"style{nameSheet}{row}{col}");
                        foreach (var itemBoreder in borderIndex)
                        {
                            var syncBorderIndex = SyncfusionHelper.ConvertFromBordexIndexSyncfusion(itemBoreder, out bool isEmpty);
                            if (!isEmpty)
                            {
                                var syncLinesIndex = SyncfusionHelper.ConvertFromLineStyleSyncfusion(lineIndex, out bool isEmptyLine);
                                if (!isEmptyLine)
                                {
                                    foreach (ExcelBordersIndex ExcelborderIndex in syncBorderIndex)
                                    {
                                        if (row == firstRowIndex + 1 && ExcelborderIndex == ExcelBordersIndex.EdgeTop)
                                        {
                                            style.Borders[ExcelborderIndex].ColorRGB = Syncfusion.Drawing.Color.FromArgb(255, 0, 0, 0);
                                            style.Borders[ExcelborderIndex].LineStyle = syncLinesIndex;
                                        }
                                        if (row == lastRowIndex + 1 && ExcelborderIndex == ExcelBordersIndex.EdgeBottom)
                                        {
                                            style.Borders[ExcelborderIndex].ColorRGB = Syncfusion.Drawing.Color.FromArgb(255, 0, 0, 0);
                                            style.Borders[ExcelborderIndex].LineStyle = syncLinesIndex;
                                        }
                                        if (col == firstColumnIndex + 1 && ExcelborderIndex == ExcelBordersIndex.EdgeLeft)
                                        {
                                            style.Borders[ExcelborderIndex].ColorRGB = Syncfusion.Drawing.Color.FromArgb(255, 0, 0, 0);
                                            style.Borders[ExcelborderIndex].LineStyle = syncLinesIndex;
                                        }
                                        if (col == lastColumnIndex + 1 && ExcelborderIndex == ExcelBordersIndex.EdgeRight)
                                        {
                                            style.Borders[ExcelborderIndex].ColorRGB = Syncfusion.Drawing.Color.FromArgb(255, 0, 0, 0);
                                            style.Borders[ExcelborderIndex].LineStyle = syncLinesIndex;
                                        }

                                    }
                                }

                            }
                            
                        }
                        cell.CellStyle = style;
                    }
                }
            }
            else
            {
                for (int row = firstRowIndex + 1; row <= lastRowIndex + 1; row++)
                {
                    for (int col = firstColumnIndex +1; col <= lastColumnIndex + 1; col++)
                    {
                        IRange cell = _workbook.Worksheets[nameSheet][row, col];
                        IStyle style = cell.CellStyle ?? _workbook.Styles.Add($"style{nameSheet}{row}{col}");
                        foreach (var itemBoreder in borderIndex)
                        {
                            
                            var syncBorderIndex = SyncfusionHelper.ConvertFromBordexIndexSyncfusion(itemBoreder, out bool isEmpty);
                            if (!isEmpty)
                            {
                                var syncLinesIndex = SyncfusionHelper.ConvertFromLineStyleSyncfusion(lineIndex, out bool isEmptyLine);
                                if (!isEmptyLine)
                                {
                                    foreach (ExcelBordersIndex ExcelborderIndex in syncBorderIndex)
                                    {
                                        style.Borders[ExcelborderIndex].ColorRGB = Syncfusion.Drawing.Color.FromArgb(255, 0, 0, 0);
                                        style.Borders[ExcelborderIndex].LineStyle = syncLinesIndex;
                                    }
                                }

                            }
                            cell.CellStyle = style;
                        }
                    }
                }
            }
            
        }

        public void SetBorder(BorderIndex[] borderIndex, LinesIndex lineIndex, int sheetIndex, int firstRowIndex, int firstColumnIndex, int lastRowIndex, int lastColumnIndex, bool allRange)
        {
            if (!allRange)
            {
                for (int row = firstRowIndex + 1; row <= lastRowIndex + 1; row++)
                {
                    for (int col = firstColumnIndex + 1; col <= lastColumnIndex + 1; col++)
                    {
                        IRange cell = _workbook.Worksheets[sheetIndex][row, col];
                        string nameSheet = _workbook.Worksheets[sheetIndex].Name;
                        IStyle style = cell.CellStyle ?? _workbook.Styles.Add($"style{nameSheet}{row}{col}");
                        foreach (var itemBoreder in borderIndex)
                        {
                            var syncBorderIndex = SyncfusionHelper.ConvertFromBordexIndexSyncfusion(itemBoreder, out bool isEmpty);
                            if (!isEmpty)
                            {
                                var syncLinesIndex = SyncfusionHelper.ConvertFromLineStyleSyncfusion(lineIndex, out bool isEmptyLine);
                                if (!isEmptyLine)
                                {
                                    foreach (ExcelBordersIndex ExcelborderIndex in syncBorderIndex)
                                    {
                                        if (row == firstRowIndex + 1 && ExcelborderIndex == ExcelBordersIndex.EdgeTop)
                                        {
                                            style.Borders[ExcelborderIndex].ColorRGB = Syncfusion.Drawing.Color.FromArgb(255, 0, 0, 0);
                                            style.Borders[ExcelborderIndex].LineStyle = syncLinesIndex;
                                        }
                                        if (row == lastRowIndex + 1 && ExcelborderIndex == ExcelBordersIndex.EdgeBottom)
                                        {
                                            style.Borders[ExcelborderIndex].ColorRGB = Syncfusion.Drawing.Color.FromArgb(255, 0, 0, 0);
                                            style.Borders[ExcelborderIndex].LineStyle = syncLinesIndex;
                                        }
                                        if (col == firstColumnIndex + 1 && ExcelborderIndex == ExcelBordersIndex.EdgeLeft)
                                        {
                                            style.Borders[ExcelborderIndex].ColorRGB = Syncfusion.Drawing.Color.FromArgb(255, 0, 0, 0);
                                            style.Borders[ExcelborderIndex].LineStyle = syncLinesIndex;
                                        }
                                        if (col == lastColumnIndex + 1 && ExcelborderIndex == ExcelBordersIndex.EdgeRight)
                                        {
                                            style.Borders[ExcelborderIndex].ColorRGB = Syncfusion.Drawing.Color.FromArgb(255, 0, 0, 0);
                                            style.Borders[ExcelborderIndex].LineStyle = syncLinesIndex;
                                        }

                                    }
                                }

                            }

                        }
                        cell.CellStyle = style;
                    }
                }
            }
            else
            {
                for (int row = firstRowIndex + 1; row <= lastRowIndex + 1; row++)
                {
                    for (int col = firstColumnIndex + 1; col <= lastColumnIndex + 1; col++)
                    {
                        IRange cell = _workbook.Worksheets[sheetIndex][row, col];
                        string nameSheet = _workbook.Worksheets[sheetIndex].Name;
                        IStyle style = cell.CellStyle ?? _workbook.Styles.Add($"style{nameSheet}{row}{col}");
                        foreach (var itemBoreder in borderIndex)
                        {

                            var syncBorderIndex = SyncfusionHelper.ConvertFromBordexIndexSyncfusion(itemBoreder, out bool isEmpty);
                            if (!isEmpty)
                            {
                                var syncLinesIndex = SyncfusionHelper.ConvertFromLineStyleSyncfusion(lineIndex, out bool isEmptyLine);
                                if (!isEmptyLine)
                                {
                                    foreach (ExcelBordersIndex ExcelborderIndex in syncBorderIndex)
                                    {
                                        style.Borders[ExcelborderIndex].ColorRGB = Syncfusion.Drawing.Color.FromArgb(255, 0, 0, 0);
                                        style.Borders[ExcelborderIndex].LineStyle = syncLinesIndex;
                                    }
                                }

                            }
                            cell.CellStyle = style;
                        }
                    }
                }
            }
        }

        public void SetBorder(BorderIndex[] borderIndex, LinesIndex lineIndex, string nameSheet, int rowIndex, int columnIndex, int a, int r, int g, int b)
        {
            IRange cell = _workbook.Worksheets[nameSheet][rowIndex + 1, columnIndex + 1];
            IStyle style = cell.CellStyle ?? _workbook.Styles.Add($"style{nameSheet}{rowIndex + 1}{columnIndex + 1}");
            foreach (var itemBorder in borderIndex)
            {
                var syncBorderIndex = SyncfusionHelper.ConvertFromBordexIndexSyncfusion(itemBorder, out bool isEmpty);
                if (!isEmpty)
                {
                    var syncLinesIndex = SyncfusionHelper.ConvertFromLineStyleSyncfusion(lineIndex, out bool isEmptyLine);
                    if (!isEmptyLine)
                    {
                        
                        foreach (ExcelBordersIndex ExcelborderIndex in syncBorderIndex)
                        {
                            style.Borders[ExcelborderIndex].ColorRGB = Syncfusion.Drawing.Color.FromArgb(a, r, g, b);
                            style.Borders[ExcelborderIndex].LineStyle = syncLinesIndex;
                        }
                    }

                }
                _workbook.Worksheets[nameSheet][rowIndex + 1, columnIndex + 1].CellStyle = style;
            }
        }

        public void SetBorder(BorderIndex[] borderIndex, LinesIndex lineIndex, int sheetIndex, int rowIndex, int columnIndex, int a, int r, int g, int b)
        {
            IRange cell = _workbook.Worksheets[sheetIndex][rowIndex + 1, columnIndex + 1];
            string nameSheet = _workbook.Worksheets[sheetIndex].Name;
            IStyle style = cell.CellStyle ?? _workbook.Styles.Add($"style{nameSheet}{rowIndex + 1}{columnIndex + 1}");
            foreach (var itemBorder in borderIndex)
            {
                var syncBorderIndex = SyncfusionHelper.ConvertFromBordexIndexSyncfusion(itemBorder, out bool isEmpty);
                if (!isEmpty)
                {
                    var syncLinesIndex = SyncfusionHelper.ConvertFromLineStyleSyncfusion(lineIndex, out bool isEmptyLine);
                    if (!isEmptyLine)
                    {

                        foreach (ExcelBordersIndex ExcelborderIndex in syncBorderIndex)
                        {
                            style.Borders[ExcelborderIndex].ColorRGB = Syncfusion.Drawing.Color.FromArgb(a, r, g, b);
                            style.Borders[ExcelborderIndex].LineStyle = syncLinesIndex;
                        }
                    }

                }
                _workbook.Worksheets[nameSheet][rowIndex + 1, columnIndex + 1].CellStyle = style;
            }
        }

        public void SetBorder(BorderIndex[] borderIndex, LinesIndex lineIndex, string nameSheet, int firstRowIndex, int firstColumnIndex, int lastRowIndex, int lastColumnIndex, int a, int r, int g, int b, bool allRange)
        {
            if (!allRange)
            {
                for (int row = firstRowIndex + 1; row <= lastRowIndex + 1; row++)
                {
                    for (int col = firstColumnIndex + 1; col <= lastColumnIndex + 1; col++)
                    {
                        IRange cell = _workbook.Worksheets[nameSheet][row, col];
                        IStyle style = cell.CellStyle ?? _workbook.Styles.Add($"style{nameSheet}{row}{col}");
                        foreach (var itemBoreder in borderIndex)
                        {
                            var syncBorderIndex = SyncfusionHelper.ConvertFromBordexIndexSyncfusion(itemBoreder, out bool isEmpty);
                            if (!isEmpty)
                            {
                                var syncLinesIndex = SyncfusionHelper.ConvertFromLineStyleSyncfusion(lineIndex, out bool isEmptyLine);
                                if (!isEmptyLine)
                                {
                                    foreach (ExcelBordersIndex ExcelborderIndex in syncBorderIndex)
                                    {
                                        if (row == firstRowIndex + 1 && ExcelborderIndex == ExcelBordersIndex.EdgeTop)
                                        {
                                            style.Borders[ExcelborderIndex].ColorRGB = Syncfusion.Drawing.Color.FromArgb(a, r, g, b);
                                            style.Borders[ExcelborderIndex].LineStyle = syncLinesIndex;
                                        }
                                        if (row == lastRowIndex + 1 && ExcelborderIndex == ExcelBordersIndex.EdgeBottom)
                                        {
                                            style.Borders[ExcelborderIndex].ColorRGB = Syncfusion.Drawing.Color.FromArgb(a, r, g, b);
                                            style.Borders[ExcelborderIndex].LineStyle = syncLinesIndex;
                                        }
                                        if (col == firstColumnIndex + 1 && ExcelborderIndex == ExcelBordersIndex.EdgeLeft)
                                        {
                                            style.Borders[ExcelborderIndex].ColorRGB = Syncfusion.Drawing.Color.FromArgb(a, r, g, b);
                                            style.Borders[ExcelborderIndex].LineStyle = syncLinesIndex;
                                        }
                                        if (col == lastColumnIndex + 1 && ExcelborderIndex == ExcelBordersIndex.EdgeRight)
                                        {
                                            style.Borders[ExcelborderIndex].ColorRGB = Syncfusion.Drawing.Color.FromArgb(a, r, g, b);
                                            style.Borders[ExcelborderIndex].LineStyle = syncLinesIndex;
                                        }

                                    }
                                }

                            }

                        }
                        cell.CellStyle = style;
                    }
                }
            }
            else
            {
                for (int row = firstRowIndex + 1; row <= lastRowIndex + 1; row++)
                {
                    for (int col = firstColumnIndex + 1; col <= lastColumnIndex + 1; col++)
                    {
                        IRange cell = _workbook.Worksheets[nameSheet][row, col];
                        IStyle style = cell.CellStyle ?? _workbook.Styles.Add($"style{nameSheet}{row}{col}");
                        foreach (var itemBoreder in borderIndex)
                        {

                            var syncBorderIndex = SyncfusionHelper.ConvertFromBordexIndexSyncfusion(itemBoreder, out bool isEmpty);
                            if (!isEmpty)
                            {
                                var syncLinesIndex = SyncfusionHelper.ConvertFromLineStyleSyncfusion(lineIndex, out bool isEmptyLine);
                                if (!isEmptyLine)
                                {
                                    foreach (ExcelBordersIndex ExcelborderIndex in syncBorderIndex)
                                    {
                                        style.Borders[ExcelborderIndex].ColorRGB = Syncfusion.Drawing.Color.FromArgb(a, r, g, b);
                                        style.Borders[ExcelborderIndex].LineStyle = syncLinesIndex;
                                    }
                                }
                            }
                            cell.CellStyle = style;
                        }
                    }
                }
            }
        }

        public void SetBorder(BorderIndex[] borderIndex, LinesIndex lineIndex, int sheetIndex, int firstRowIndex, int firstColumnIndex, int lastRowIndex, int lastColumnIndex, int a, int r, int g, int b, bool allRange)
        {
            if (!allRange)
            {
                for (int row = firstRowIndex + 1; row <= lastRowIndex + 1; row++)
                {
                    for (int col = firstColumnIndex + 1; col <= lastColumnIndex + 1; col++)
                    {
                        IRange cell = _workbook.Worksheets[sheetIndex][row, col];
                        string nameSheet = _workbook.Worksheets[sheetIndex].Name;
                        IStyle style = cell.CellStyle ?? _workbook.Styles.Add($"style{nameSheet}{row}{col}");
                        foreach (var itemBoreder in borderIndex)
                        {
                            var syncBorderIndex = SyncfusionHelper.ConvertFromBordexIndexSyncfusion(itemBoreder, out bool isEmpty);
                            if (!isEmpty)
                            {
                                var syncLinesIndex = SyncfusionHelper.ConvertFromLineStyleSyncfusion(lineIndex, out bool isEmptyLine);
                                if (!isEmptyLine)
                                {
                                    foreach (ExcelBordersIndex ExcelborderIndex in syncBorderIndex)
                                    {
                                        if (row == firstRowIndex + 1 && ExcelborderIndex == ExcelBordersIndex.EdgeTop)
                                        {
                                            style.Borders[ExcelborderIndex].ColorRGB = Syncfusion.Drawing.Color.FromArgb(a, r, g, b);
                                            style.Borders[ExcelborderIndex].LineStyle = syncLinesIndex;
                                        }
                                        if (row == lastRowIndex + 1 && ExcelborderIndex == ExcelBordersIndex.EdgeBottom)
                                        {
                                            style.Borders[ExcelborderIndex].ColorRGB = Syncfusion.Drawing.Color.FromArgb(a, r, g, b);
                                            style.Borders[ExcelborderIndex].LineStyle = syncLinesIndex;
                                        }
                                        if (col == firstColumnIndex + 1 && ExcelborderIndex == ExcelBordersIndex.EdgeLeft)
                                        {
                                            style.Borders[ExcelborderIndex].ColorRGB = Syncfusion.Drawing.Color.FromArgb(a, r, g, b);
                                            style.Borders[ExcelborderIndex].LineStyle = syncLinesIndex;
                                        }
                                        if (col == lastColumnIndex + 1 && ExcelborderIndex == ExcelBordersIndex.EdgeRight)
                                        {
                                            style.Borders[ExcelborderIndex].ColorRGB = Syncfusion.Drawing.Color.FromArgb(a, r, g, b);
                                            style.Borders[ExcelborderIndex].LineStyle = syncLinesIndex;
                                        }

                                    }
                                }

                            }

                        }
                        cell.CellStyle = style;
                    }
                }
            }
            else
            {
                for (int row = firstRowIndex + 1; row <= lastRowIndex + 1; row++)
                {
                    for (int col = firstColumnIndex + 1; col <= lastColumnIndex + 1; col++)
                    {
                        IRange cell = _workbook.Worksheets[sheetIndex][row, col];
                        string nameSheet = _workbook.Worksheets[sheetIndex].Name;
                        IStyle style = cell.CellStyle ?? _workbook.Styles.Add($"style{nameSheet}{row}{col}");
                        foreach (var itemBoreder in borderIndex)
                        {

                            var syncBorderIndex = SyncfusionHelper.ConvertFromBordexIndexSyncfusion(itemBoreder, out bool isEmpty);
                            if (!isEmpty)
                            {
                                var syncLinesIndex = SyncfusionHelper.ConvertFromLineStyleSyncfusion(lineIndex, out bool isEmptyLine);
                                if (!isEmptyLine)
                                {
                                    foreach (ExcelBordersIndex ExcelborderIndex in syncBorderIndex)
                                    {
                                        style.Borders[ExcelborderIndex].ColorRGB = Syncfusion.Drawing.Color.FromArgb(a, r, g, b);
                                        style.Borders[ExcelborderIndex].LineStyle = syncLinesIndex;
                                    }
                                }
                            }
                            cell.CellStyle = style;
                        }
                    }
                }
            }
        }

        public void SetFont(FontSettings fontSettings, int sheetIndex, int rowIndex, int columnIndex)
        {
            IRange cell = _workbook.Worksheets[sheetIndex][rowIndex + 1, columnIndex + 1];
            string nameSheet = _workbook.Worksheets[sheetIndex].Name;
            IStyle style = cell.CellStyle ?? _workbook.Styles.Add($"style{nameSheet}{rowIndex + 1}{columnIndex + 1}");

            if (fontSettings.Bold.HasValue)
            {
                style.Font.Bold = fontSettings.Bold.Value;
            }
            if (fontSettings.Italics.HasValue)
            {
                style.Font.Italic = fontSettings.Italics.Value;
            }
            if (fontSettings.DoubleUnderline.HasValue && fontSettings.DoubleUnderline.Value)
            {
                style.Font.Underline = ExcelUnderline.Double;
            }
            else if (fontSettings.Underline.HasValue && fontSettings.Underline.Value)
            {
                style.Font.Underline = ExcelUnderline.Single;
            }
            if (fontSettings.TextCrossed.HasValue)
            {
                style.Font.Strikethrough = fontSettings.TextCrossed.Value;
            }
            if (fontSettings.TextWrapping.HasValue)
            {
                style.WrapText = fontSettings.TextWrapping.Value;
            }
            if (!string.IsNullOrEmpty(fontSettings.FontName))
            {
                style.Font.FontName = fontSettings.FontName;
            }
            if (fontSettings.FontSize.HasValue)
            {
                style.Font.Size = fontSettings.FontSize.Value;
            }
            if (fontSettings.A.HasValue && fontSettings.R.HasValue && fontSettings.G.HasValue && fontSettings.B.HasValue)
            {
                Syncfusion.Drawing.Color c = Syncfusion.Drawing.Color.FromArgb(fontSettings.A.Value, fontSettings.R.Value, fontSettings.G.Value, fontSettings.B.Value);
                style.Font.RGBColor = c;
            }
            cell.CellStyle = style;
        }

        public void SetFont(FontSettings fontSettings, string sheetName, int rowIndex, int columnIndex)
        {
            IRange cell = _workbook.Worksheets[sheetName][rowIndex + 1, columnIndex + 1];
            IStyle style = cell.CellStyle ?? _workbook.Styles.Add($"style{sheetName}{rowIndex + 1}{columnIndex + 1}");

            if (fontSettings.Bold.HasValue)
            {
                style.Font.Bold = fontSettings.Bold.Value;
            }
            if (fontSettings.Italics.HasValue)
            {
                style.Font.Italic = fontSettings.Italics.Value;
            }
            if (fontSettings.DoubleUnderline.HasValue && fontSettings.DoubleUnderline.Value)
            {
                style.Font.Underline = ExcelUnderline.Double;
            }
            else if (fontSettings.Underline.HasValue && fontSettings.Underline.Value)
            {
                style.Font.Underline = ExcelUnderline.Single;
            }
            if (fontSettings.TextCrossed.HasValue)
            {
                style.Font.Strikethrough = fontSettings.TextCrossed.Value;
            }
            if (fontSettings.TextWrapping.HasValue)
            {
                style.WrapText = fontSettings.TextWrapping.Value;
            }
            if (!string.IsNullOrEmpty(fontSettings.FontName))
            {
                style.Font.FontName = fontSettings.FontName;
            }
            if (fontSettings.FontSize.HasValue)
            {
                style.Font.Size = fontSettings.FontSize.Value;
            }
            if (fontSettings.A.HasValue && fontSettings.R.HasValue && fontSettings.G.HasValue && fontSettings.B.HasValue)
            {
                Syncfusion.Drawing.Color c = Syncfusion.Drawing.Color.FromArgb(fontSettings.A.Value, fontSettings.R.Value, fontSettings.G.Value, fontSettings.B.Value);
                style.Font.RGBColor = c;
            }
            cell.CellStyle = style;
        }

        public void SetFont(FontSettings fontSettings, int sheetIndex, int firstRowIndex, int firstColumnIndex, int lastRowIndex, int lastColumnIndex)
        {
            for (int row = firstRowIndex + 1; row <= lastRowIndex + 1; row++)
            {
                for (int col = firstColumnIndex + 1; col <= lastColumnIndex + 1; col++)
                {
                    IRange cell = _workbook.Worksheets[sheetIndex][row, col];
                    string nameSheet = _workbook.Worksheets[sheetIndex].Name;
                    IStyle style = cell.CellStyle ?? _workbook.Styles.Add($"style{nameSheet}{row}{col}");

                    if (fontSettings.Bold.HasValue)
                    {
                        style.Font.Bold = fontSettings.Bold.Value;
                    }
                    if (fontSettings.Italics.HasValue)
                    {
                        style.Font.Italic = fontSettings.Italics.Value;
                    }
                    if (fontSettings.DoubleUnderline.HasValue && fontSettings.DoubleUnderline.Value)
                    {
                        style.Font.Underline = ExcelUnderline.Double;
                    }
                    else if (fontSettings.Underline.HasValue && fontSettings.Underline.Value)
                    {
                        style.Font.Underline = ExcelUnderline.Single;
                    }
                    if (fontSettings.TextCrossed.HasValue)
                    {
                        style.Font.Strikethrough = fontSettings.TextCrossed.Value;
                    }
                    if (fontSettings.TextWrapping.HasValue)
                    {
                        style.WrapText = fontSettings.TextWrapping.Value;
                    }
                    if (!string.IsNullOrEmpty(fontSettings.FontName))
                    {
                        style.Font.FontName = fontSettings.FontName;
                    }
                    if (fontSettings.FontSize.HasValue)
                    {
                        style.Font.Size = fontSettings.FontSize.Value;
                    }
                    if (fontSettings.A.HasValue && fontSettings.R.HasValue && fontSettings.G.HasValue && fontSettings.B.HasValue)
                    {
                        Syncfusion.Drawing.Color c = Syncfusion.Drawing.Color.FromArgb(fontSettings.A.Value, fontSettings.R.Value, fontSettings.G.Value, fontSettings.B.Value);
                        style.Font.RGBColor = c;
                    }
                    cell.CellStyle = style;
                }
            }
        }

        public void SetFont(FontSettings fontSettings, string sheetName, int firstRowIndex, int firstColumnIndex, int lastRowIndex, int lastColumnIndex)
        {
            for (int row = firstRowIndex + 1; row <= lastRowIndex + 1; row++)
            {
                for (int col = firstColumnIndex + 1; col <= lastColumnIndex + 1; col++)
                {
                    IRange cell = _workbook.Worksheets[sheetName][row, col];
                    IStyle style = cell.CellStyle ?? _workbook.Styles.Add($"style{sheetName}{row}{col}");

                    if (fontSettings.Bold.HasValue)
                    {
                        style.Font.Bold = fontSettings.Bold.Value;
                    }
                    if (fontSettings.Italics.HasValue)
                    {
                        style.Font.Italic = fontSettings.Italics.Value;
                    }
                    if (fontSettings.DoubleUnderline.HasValue && fontSettings.DoubleUnderline.Value)
                    {
                        style.Font.Underline = ExcelUnderline.Double;
                    }
                    else if (fontSettings.Underline.HasValue && fontSettings.Underline.Value)
                    {
                        style.Font.Underline = ExcelUnderline.Single;
                    }
                    if (fontSettings.TextCrossed.HasValue)
                    {
                        style.Font.Strikethrough = fontSettings.TextCrossed.Value;
                    }
                    if (fontSettings.TextWrapping.HasValue)
                    {
                        style.WrapText = fontSettings.TextWrapping.Value;
                    }
                    if (!string.IsNullOrEmpty(fontSettings.FontName))
                    {
                        style.Font.FontName = fontSettings.FontName;
                    }
                    if (fontSettings.FontSize.HasValue)
                    {
                        style.Font.Size = fontSettings.FontSize.Value;
                    }
                    if (fontSettings.A.HasValue && fontSettings.R.HasValue && fontSettings.G.HasValue && fontSettings.B.HasValue)
                    {
                        Syncfusion.Drawing.Color c = Syncfusion.Drawing.Color.FromArgb(fontSettings.A.Value, fontSettings.R.Value, fontSettings.G.Value, fontSettings.B.Value);
                        style.Font.RGBColor = c;
                    }
                    cell.CellStyle = style;
                }
            }
        }

        public void SetHorizontalAlignment(HorizontalAlignmentIndex horizontalAlignmentIndex, int sheetIndex, int rowIndex, int columnIndex)
        {
            IRange cell = _workbook.Worksheets[sheetIndex][rowIndex + 1, columnIndex + 1];
            string nameSheet = _workbook.Worksheets[sheetIndex].Name;
            IStyle style = cell.CellStyle ?? _workbook.Styles.Add($"style{nameSheet}{rowIndex + 1}{columnIndex + 1}");
            var hAlign = SyncfusionHelper.ConvertFromHAlign(horizontalAlignmentIndex, out bool isEmpty);
            if (!isEmpty)
            {
                style.HorizontalAlignment = hAlign;
            }
            cell.CellStyle = style;
        }

        public void SetHorizontalAlignment(HorizontalAlignmentIndex horizontalAlignmentIndex, string sheetName, int rowIndex, int columnIndex)
        {
            IRange cell = _workbook.Worksheets[sheetName][rowIndex + 1, columnIndex + 1];
            IStyle style = cell.CellStyle ?? _workbook.Styles.Add($"style{sheetName}{rowIndex + 1}{columnIndex + 1}");
            var hAlign = SyncfusionHelper.ConvertFromHAlign(horizontalAlignmentIndex, out bool isEmpty);
            if (!isEmpty)
            {
                style.HorizontalAlignment = hAlign;
            }
            cell.CellStyle = style;
        }

        public void SetHorizontalAlignment(HorizontalAlignmentIndex horizontalAlignmentIndex, int sheetIndex, int firstRowIndex, int firstColumnIndex, int lastRowIndex, int lastColumnIndex)
        {
            for (int row = firstRowIndex + 1; row <= lastRowIndex + 1; row++)
            {
                for (int col = firstColumnIndex + 1; col <= lastColumnIndex + 1; col++)
                {
                    IRange cell = _workbook.Worksheets[sheetIndex][row, col];
                    string nameSheet = _workbook.Worksheets[sheetIndex].Name;
                    IStyle style = cell.CellStyle ?? _workbook.Styles.Add($"style{nameSheet}{row}{col}");
                    var hAlign = SyncfusionHelper.ConvertFromHAlign(horizontalAlignmentIndex, out bool isEmpty);
                    if (!isEmpty)
                    {
                        style.HorizontalAlignment = hAlign;
                    }
                    cell.CellStyle = style;
                }
            }
        }

        public void SetHorizontalAlignment(HorizontalAlignmentIndex horizontalAlignmentIndex, string sheetName, int firstRowIndex, int firstColumnIndex, int lastRowIndex, int lastColumnIndex)
        {
            for (int row = firstRowIndex + 1; row <= lastRowIndex + 1; row++)
            {
                for (int col = firstColumnIndex + 1; col <= lastColumnIndex + 1; col++)
                {
                    IRange cell = _workbook.Worksheets[sheetName][row, col];
                    IStyle style = cell.CellStyle ?? _workbook.Styles.Add($"style{sheetName}{row}{col}");
                    var hAlign = SyncfusionHelper.ConvertFromHAlign(horizontalAlignmentIndex, out bool isEmpty);
                    if (!isEmpty)
                    {
                        style.HorizontalAlignment = hAlign;
                    }
                    cell.CellStyle = style;
                }
            }
        }

        public void SetVerticalAlignment(VerticalAlignmentIndex verticalAlignmentIndex, int sheetIndex, int rowIndex, int columnIndex)
        {
            IRange cell = _workbook.Worksheets[sheetIndex][rowIndex + 1, columnIndex + 1];
            string nameSheet = _workbook.Worksheets[sheetIndex].Name;
            IStyle style = cell.CellStyle ?? _workbook.Styles.Add($"style{nameSheet}{rowIndex + 1}{columnIndex + 1}");
            var vAlign = SyncfusionHelper.ConvertFromVAlign(verticalAlignmentIndex, out bool isEmpty);
            if (!isEmpty)
            {
                style.VerticalAlignment = vAlign;
            }
            cell.CellStyle = style;
        }

        public void SetVerticalAlignment(VerticalAlignmentIndex verticalAlignmentIndex, string sheetName, int rowIndex, int columnIndex)
        {
            IRange cell = _workbook.Worksheets[sheetName][rowIndex + 1, columnIndex + 1];
            IStyle style = cell.CellStyle ?? _workbook.Styles.Add($"style{sheetName}{rowIndex + 1}{columnIndex + 1}");
            var vAlign = SyncfusionHelper.ConvertFromVAlign(verticalAlignmentIndex, out bool isEmpty);
            if (!isEmpty)
            {
                style.VerticalAlignment = vAlign;
            }
            cell.CellStyle = style;
        }

        public void SetVerticalAlignment(VerticalAlignmentIndex verticalAlignmentIndex, int sheetIndex, int firstRowIndex, int firstColumnIndex, int lastRowIndex, int lastColumnIndex)
        {
            for (int row = firstRowIndex + 1; row <= lastRowIndex + 1; row++)
            {
                for (int col = firstColumnIndex + 1; col <= lastColumnIndex + 1; col++)
                {
                    IRange cell = _workbook.Worksheets[sheetIndex][row, col];
                    string nameSheet = _workbook.Worksheets[sheetIndex].Name;
                    IStyle style = cell.CellStyle ?? _workbook.Styles.Add($"style{nameSheet}{row}{col}");
                    var vAlign = SyncfusionHelper.ConvertFromVAlign(verticalAlignmentIndex, out bool isEmpty);
                    if (!isEmpty)
                    {
                        style.VerticalAlignment = vAlign;
                    }
                    cell.CellStyle = style;
                }
            }
        }

        public void SetVerticalAlignment(VerticalAlignmentIndex verticalAlignmentIndex, string sheetName, int firstRowIndex, int firstColumnIndex, int lastRowIndex, int lastColumnIndex)
        {
            for (int row = firstRowIndex + 1; row <= lastRowIndex + 1; row++)
            {
                for (int col = firstColumnIndex + 1; col <= lastColumnIndex + 1; col++)
                {
                    IRange cell = _workbook.Worksheets[sheetName][row, col];
                    IStyle style = cell.CellStyle ?? _workbook.Styles.Add($"style{sheetName}{row}{col}");
                    var vAlign = SyncfusionHelper.ConvertFromVAlign(verticalAlignmentIndex, out bool isEmpty);
                    if (!isEmpty)
                    {
                        style.VerticalAlignment = vAlign;
                    }
                    cell.CellStyle = style;
                }
            }
        }

        public void MergeCells(int sheetIndex, int firstRowIndex, int firstColumnIndex, int lastRowIndex, int lastColumnIndex)
        {
            _workbook.Worksheets[sheetIndex].Range[firstRowIndex + 1, firstColumnIndex + 1, lastRowIndex + 1, lastColumnIndex + 1].Merge();
        }

        public void MergeCells(string sheetName, int firstRowIndex, int lastRowIndex, int firstColumnIndex, int lastColumnIndex)
        {
            _workbook.Worksheets[sheetName].Range[firstRowIndex + 1, firstColumnIndex + 1, lastRowIndex + 1, lastColumnIndex + 1].Merge();
        }

        // orientation

        public void ValueOrientation(int sheetIndex, int rowIndex, int columnIndex, short orientation)
        {
            IRange cell = _workbook.Worksheets[sheetIndex][rowIndex + 1, columnIndex + 1];
            string nameSheet = _workbook.Worksheets[sheetIndex].Name;
            IStyle style = cell.CellStyle ?? _workbook.Styles.Add($"style{nameSheet}{rowIndex + 1}{columnIndex + 1}");
            style.Rotation = orientation;
            cell.CellStyle = style;
        }

        public void ValueOrientation(string sheetName, int rowIndex, int columnIndex, short orientation)
        {
            IRange cell = _workbook.Worksheets[sheetName][rowIndex + 1, columnIndex + 1];
            IStyle style = cell.CellStyle ?? _workbook.Styles.Add($"style{sheetName}{rowIndex + 1}{columnIndex + 1}");
            style.Rotation = orientation;
            cell.CellStyle = style;
        }

        public void ValueOrientation(int sheetIndex, int firstRowIndex, int firstColumnIndex, int lastRowIndex, int lastColumnIndex, short orientation)
        {
            for (int row = firstRowIndex + 1;  row <= lastRowIndex + 1; row++)
            {
                for (int col = firstColumnIndex + 1; col <= lastColumnIndex + 1; col++)
                {
                    IRange cell = _workbook.Worksheets[sheetIndex][row, col];
                    string nameSheet = _workbook.Worksheets[sheetIndex].Name;
                    IStyle style = cell.CellStyle ?? _workbook.Styles.Add($"style{nameSheet}{row}{col}");
                    style.Rotation = orientation;
                    cell.CellStyle = style;
                }
            }
        }

        public void ValueOrientation(string sheetName, int firstRowIndex, int firstColumnIndex, int lastRowIndex, int lastColumnIndex, short orientation)
        {
            for (int row = firstRowIndex + 1; row <= lastRowIndex + 1; row++)
            {
                for (int col = firstColumnIndex + 1; col <= lastColumnIndex + 1; col++)
                {
                    IRange cell = _workbook.Worksheets[sheetName][row, col];
                    IStyle style = cell.CellStyle ?? _workbook.Styles.Add($"style{sheetName}{row}{col}");
                    style.Rotation = orientation;
                    cell.CellStyle = style;
                }
            }
        }

        public void SetProtectSheet(int sheetIndex, string password)
        {
            _workbook.Worksheets[sheetIndex].Protect(password);
        }

        public void SetProtectSheet(string sheetName, string password)
        {
            _workbook.Worksheets[sheetName].Protect(password);
        }

        public void DropDownList(int sheetIndex, int dataSheetIndex, int firstRowIndex, int firstColumnIndex, int lastRowIndex, int lastColumnIndex, string rangeDataToList)
        {
            IWorksheet sheet = _workbook.Worksheets[sheetIndex];
            IWorksheet dataSheet = _workbook.Worksheets[dataSheetIndex];

            IRange range = dataSheet.Range[rangeDataToList];

            int index = 0;
            string[] data;
            if (range.Rows.Count() > 1 && range.Columns.Count() == 1)
            {
                data = new string[range.Rows.Count()];

                foreach (var item in range)
                {
                    data[index++] = item.Text;
                }
            }
            else
            {
                data = new string[range.Columns.Count()];

                foreach (var item in range)
                {
                    data[index++] = item.Text;
                }
            }
            
            IRange cellRange = sheet.Range[firstRowIndex + 1, firstColumnIndex + 1, lastRowIndex + 1, lastColumnIndex + 1];
            IDataValidation dataValidation = cellRange.DataValidation;
            dataValidation.ListOfValues = data;

            dataValidation.IsSuppressDropDownArrow = false;
            dataValidation.ShowErrorBox = true;
            cellRange.CellStyle.NumberFormat = range.CellStyle.NumberFormat;
        }

        public void DropDownList(string sheetName, int dataSheetIndex, int firstRowIndex, int firstColumnIndex, int lastRowIndex, int lastColumnIndex, string rangeDataToList)
        {
            IWorksheet sheet = _workbook.Worksheets[sheetName];
            IWorksheet dataSheet = _workbook.Worksheets[dataSheetIndex];

            IRange range = dataSheet.Range[rangeDataToList];

            int index = 0;
            string[] data;
            if (range.Rows.Count() > 1 && range.Columns.Count() == 1)
            {
                data = new string[range.Rows.Count()];

                foreach (var item in range)
                {
                    data[index++] = item.Text;
                }
            }
            else
            {
                data = new string[range.Columns.Count()];

                foreach (var item in range)
                {
                    data[index++] = item.Text;
                }
            }
            IRange cellRange = sheet.Range[firstRowIndex + 1, firstColumnIndex + 1, lastRowIndex + 1, lastColumnIndex + 1];
            IDataValidation dataValidation = cellRange.DataValidation;
            dataValidation.ListOfValues = data;

            dataValidation.IsSuppressDropDownArrow = false;
            dataValidation.ShowErrorBox = true;
            cellRange.CellStyle.NumberFormat = range.CellStyle.NumberFormat;
        }

        public void DropDownList(int sheetIndex, string dataSheetName, int firstRowIndex, int firstColumnIndex, int lastRowIndex, int lastColumnIndex, string rangeDataToList)
        {
            IWorksheet sheet = _workbook.Worksheets[sheetIndex];
            IWorksheet dataSheet = _workbook.Worksheets[dataSheetName];

            IRange range = dataSheet.Range[rangeDataToList];

            int index = 0;
            string[] data;
            if (range.Rows.Count() > 1 && range.Columns.Count() == 1)
            {
                data = new string[range.Rows.Count()];

                foreach (var item in range)
                {
                    data[index++] = item.Text;
                }
            }
            else
            {
                data = new string[range.Columns.Count()];

                foreach (var item in range)
                {
                    data[index++] = item.Text;
                }
            }
            IRange cellRange = sheet.Range[firstRowIndex + 1, firstColumnIndex + 1, lastRowIndex + 1, lastColumnIndex + 1];
            IDataValidation dataValidation = cellRange.DataValidation;
            dataValidation.ListOfValues = data;

            dataValidation.IsSuppressDropDownArrow = false;
            dataValidation.ShowErrorBox = true;
            cellRange.CellStyle.NumberFormat = range.CellStyle.NumberFormat;
        }

        public void DropDownList(string sheetName, string dataSheetName, int firstRowIndex, int firstColumnIndex, int lastRowIndex, int lastColumnIndex, string rangeDataToList)
        {
            IWorksheet sheet = _workbook.Worksheets[sheetName];
            IWorksheet dataSheet = _workbook.Worksheets[dataSheetName];

            IRange range = dataSheet.Range[rangeDataToList];

            int index = 0;
            string[] data;
            if (range.Rows.Count() > 1 && range.Columns.Count() == 1)
            {
                data = new string[range.Rows.Count()];

                foreach (var item in range)
                {
                    data[index++] = item.Text;
                }
            }
            else
            {
                data = new string[range.Columns.Count()];

                foreach (var item in range)
                {
                    data[index++] = item.Text;
                }
            }
            IRange cellRange = sheet.Range[firstRowIndex + 1, firstColumnIndex + 1, lastRowIndex + 1, lastColumnIndex + 1];
            IDataValidation dataValidation = cellRange.DataValidation;
            dataValidation.ListOfValues = data;

            dataValidation.IsSuppressDropDownArrow = false;
            dataValidation.ShowErrorBox = true;
            cellRange.CellStyle.NumberFormat = range.CellStyle.NumberFormat;
        }

        public void SetAutoWidth(int sheetIndex, int columnIndex)
        {
            _workbook.Worksheets[sheetIndex].AutofitColumn(columnIndex + 1);
        }

        public void SetAutoWidth(string sheetName, int columnIndex)
        {
            _workbook.Worksheets[sheetName].AutofitColumn(columnIndex + 1);
        }

        public void SetAutoWidth(int sheetIndex, int firstColumnIndex, int lastColumnIndex)
        {
            for (int col = firstColumnIndex + 1; col <= lastColumnIndex + 1; col++)
            {
                _workbook.Worksheets[sheetIndex].AutofitColumn(col);
            }
        }

        public void SetAutoWidth(string sheetName, int firstColumnIndex, int lastColumnIndex)
        {
            for (int col = firstColumnIndex + 1; col <= lastColumnIndex + 1; col++)
            {
                _workbook.Worksheets[sheetName].AutofitColumn(col);
            }
        }

        public void SetAutoWidth(int sheetIndex)
        {
            int allUsedColumns = _workbook.Worksheets[sheetIndex].UsedRange.LastColumn;
            for (int col = 1; col <= allUsedColumns; col++)
            {
                _workbook.Worksheets[sheetIndex].AutofitColumn(col);
            }
        }

        public void SetAutoWidth(string sheetName)
        {
            int allUsedColumns = _workbook.Worksheets[sheetName].UsedRange.LastColumn;
            for (int col = 1; col <= allUsedColumns; col++)
            {
                _workbook.Worksheets[sheetName].AutofitColumn(col);
            }
        }

        public void ConditionalFormatting(int sheetIndex, int firstRowIndex, int firstColumnIndex, int lastRowIndex, int lastColumnIndex, ConditionAndFormatting[] conditionAndFormattings)
        {
            IConditionalFormats conditionalFormats = _workbook.Worksheets[sheetIndex][firstRowIndex + 1, firstColumnIndex + 1, lastRowIndex + 1, lastColumnIndex + 1].ConditionalFormats;

            foreach (var conditionalFormatting in conditionAndFormattings)
            {
                var comparisonOperatorSyncfusion = SyncfusionHelper.ConvertFromComparisonOperatorSyncfusion(conditionalFormatting.ComparisonOperatorIndex, out bool isEmpty);
                if (!isEmpty)
                {
                    IConditionalFormat format = conditionalFormats.AddCondition();
                    format.FormatType = ExcelCFType.CellValue;
                    format.Operator = comparisonOperatorSyncfusion;
                    format.FirstFormula = conditionalFormatting.Condition;
                    if (conditionalFormatting.Bold.HasValue)
                    {
                        format.IsBold = conditionalFormatting.Bold.Value;
                    }
                    if (conditionalFormatting.Italics.HasValue)
                    {
                        format.IsItalic = conditionalFormatting.Italics.Value;
                    }
                    if (conditionalFormatting.DoubleUnderline.HasValue && conditionalFormatting.DoubleUnderline.Value)
                    {
                        format.Underline = ExcelUnderline.Double;
                    }
                    else if (conditionalFormatting.Underline.HasValue && conditionalFormatting.Underline.Value)
                    {
                        format.Underline = ExcelUnderline.Single;
                    }
                    if (conditionalFormatting.BackgroundColorA.HasValue && conditionalFormatting.BackgroundColorR.HasValue && conditionalFormatting.BackgroundColorG.HasValue && conditionalFormatting.BackgroundColorB.HasValue)
                    {
                        format.BackColorRGB = Syncfusion.Drawing.Color.FromArgb(conditionalFormatting.BackgroundColorA.Value, conditionalFormatting.BackgroundColorR.Value, conditionalFormatting.BackgroundColorG.Value, conditionalFormatting.BackgroundColorB.Value);
                    }
                    if (conditionalFormatting.TextColorA.HasValue && conditionalFormatting.TextColorR.HasValue && conditionalFormatting.TextColorG.HasValue && conditionalFormatting.TextColorB.HasValue)
                    {
                        format.FontColorRGB = Syncfusion.Drawing.Color.FromArgb(conditionalFormatting.TextColorA.Value, conditionalFormatting.TextColorR.Value, conditionalFormatting.TextColorG.Value, conditionalFormatting.TextColorB.Value);
                    }
                }
            }
            
        }

        public void ConditionalFormatting(string sheetName, int firstRowIndex, int firstColumnIndex, int lastRowIndex, int lastColumnIndex, ConditionAndFormatting[] conditionAndFormattings)
        {
            IConditionalFormats conditionalFormats = _workbook.Worksheets[sheetName][firstRowIndex + 1, firstColumnIndex + 1, lastRowIndex + 1, lastColumnIndex + 1].ConditionalFormats;

            foreach (var conditionalFormatting in conditionAndFormattings)
            {
                var comparisonOperatorSyncfusion = SyncfusionHelper.ConvertFromComparisonOperatorSyncfusion(conditionalFormatting.ComparisonOperatorIndex, out bool isEmpty);
                if (!isEmpty)
                {
                    IConditionalFormat format = conditionalFormats.AddCondition();
                    format.FormatType = ExcelCFType.CellValue;
                    format.Operator = comparisonOperatorSyncfusion;
                    format.FirstFormula = conditionalFormatting.Condition;
                    if (conditionalFormatting.Bold.HasValue)
                    {
                        format.IsBold = conditionalFormatting.Bold.Value;
                    }
                    if (conditionalFormatting.Italics.HasValue)
                    {
                        format.IsItalic = conditionalFormatting.Italics.Value;
                    }
                    if (conditionalFormatting.DoubleUnderline.HasValue && conditionalFormatting.DoubleUnderline.Value)
                    {
                        format.Underline = ExcelUnderline.Double;
                    }
                    else if (conditionalFormatting.Underline.HasValue && conditionalFormatting.Underline.Value)
                    {
                        format.Underline = ExcelUnderline.Single;
                    }
                    if (conditionalFormatting.BackgroundColorA.HasValue && conditionalFormatting.BackgroundColorR.HasValue && conditionalFormatting.BackgroundColorG.HasValue && conditionalFormatting.BackgroundColorB.HasValue)
                    {
                        format.BackColorRGB = Syncfusion.Drawing.Color.FromArgb(conditionalFormatting.BackgroundColorA.Value, conditionalFormatting.BackgroundColorR.Value, conditionalFormatting.BackgroundColorG.Value, conditionalFormatting.BackgroundColorB.Value);
                    }
                    if (conditionalFormatting.TextColorA.HasValue && conditionalFormatting.TextColorR.HasValue && conditionalFormatting.TextColorG.HasValue && conditionalFormatting.TextColorB.HasValue)
                    {
                        format.FontColorRGB = Syncfusion.Drawing.Color.FromArgb(conditionalFormatting.TextColorA.Value, conditionalFormatting.TextColorR.Value, conditionalFormatting.TextColorG.Value, conditionalFormatting.TextColorB.Value);
                    }
                }
            }

        }

        public void ConditionalFormatting(int sheetIndex, int rowIndex, int columnIndex, ConditionAndFormatting[] conditionAndFormattings)
        {
            IConditionalFormats conditionalFormats = _workbook.Worksheets[sheetIndex][rowIndex + 1, columnIndex + 1, rowIndex + 1, columnIndex + 1].ConditionalFormats;

            foreach (var conditionalFormatting in conditionAndFormattings)
            {
                var comparisonOperatorSyncfusion = SyncfusionHelper.ConvertFromComparisonOperatorSyncfusion(conditionalFormatting.ComparisonOperatorIndex, out bool isEmpty);
                if (!isEmpty)
                {
                    IConditionalFormat format = conditionalFormats.AddCondition();
                    format.FormatType = ExcelCFType.CellValue;
                    format.Operator = comparisonOperatorSyncfusion;
                    format.FirstFormula = conditionalFormatting.Condition;
                    if (conditionalFormatting.Bold.HasValue)
                    {
                        format.IsBold = conditionalFormatting.Bold.Value;
                    }
                    if (conditionalFormatting.Italics.HasValue)
                    {
                        format.IsItalic = conditionalFormatting.Italics.Value;
                    }
                    if (conditionalFormatting.DoubleUnderline.HasValue && conditionalFormatting.DoubleUnderline.Value)
                    {
                        format.Underline = ExcelUnderline.Double;
                    }
                    else if (conditionalFormatting.Underline.HasValue && conditionalFormatting.Underline.Value)
                    {
                        format.Underline = ExcelUnderline.Single;
                    }
                    if (conditionalFormatting.BackgroundColorA.HasValue && conditionalFormatting.BackgroundColorR.HasValue && conditionalFormatting.BackgroundColorG.HasValue && conditionalFormatting.BackgroundColorB.HasValue)
                    {
                        format.BackColorRGB = Syncfusion.Drawing.Color.FromArgb(conditionalFormatting.BackgroundColorA.Value, conditionalFormatting.BackgroundColorR.Value, conditionalFormatting.BackgroundColorG.Value, conditionalFormatting.BackgroundColorB.Value);
                    }
                    if (conditionalFormatting.TextColorA.HasValue && conditionalFormatting.TextColorR.HasValue && conditionalFormatting.TextColorG.HasValue && conditionalFormatting.TextColorB.HasValue)
                    {
                        format.FontColorRGB = Syncfusion.Drawing.Color.FromArgb(conditionalFormatting.TextColorA.Value, conditionalFormatting.TextColorR.Value, conditionalFormatting.TextColorG.Value, conditionalFormatting.TextColorB.Value);
                    }
                }
            }

        }

        public void ConditionalFormatting(string sheetName, int rowIndex, int columnIndex, ConditionAndFormatting[] conditionAndFormattings)
        {
            IConditionalFormats conditionalFormats = _workbook.Worksheets[sheetName][rowIndex + 1, columnIndex + 1, rowIndex + 1, columnIndex + 1].ConditionalFormats;

            foreach (var conditionalFormatting in conditionAndFormattings)
            {
                var comparisonOperatorSyncfusion = SyncfusionHelper.ConvertFromComparisonOperatorSyncfusion(conditionalFormatting.ComparisonOperatorIndex, out bool isEmpty);
                if (!isEmpty)
                {
                    IConditionalFormat format = conditionalFormats.AddCondition();
                    format.FormatType = ExcelCFType.CellValue;
                    format.Operator = comparisonOperatorSyncfusion;
                    format.FirstFormula = conditionalFormatting.Condition;
                    if (conditionalFormatting.Bold.HasValue)
                    {
                        format.IsBold = conditionalFormatting.Bold.Value;
                    }
                    if (conditionalFormatting.Italics.HasValue)
                    {
                        format.IsItalic = conditionalFormatting.Italics.Value;
                    }
                    if (conditionalFormatting.DoubleUnderline.HasValue && conditionalFormatting.DoubleUnderline.Value)
                    {
                        format.Underline = ExcelUnderline.Double;
                    }
                    else if (conditionalFormatting.Underline.HasValue && conditionalFormatting.Underline.Value)
                    {
                        format.Underline = ExcelUnderline.Single;
                    }
                    if (conditionalFormatting.BackgroundColorA.HasValue && conditionalFormatting.BackgroundColorR.HasValue && conditionalFormatting.BackgroundColorG.HasValue && conditionalFormatting.BackgroundColorB.HasValue)
                    {
                        format.BackColorRGB = Syncfusion.Drawing.Color.FromArgb(conditionalFormatting.BackgroundColorA.Value, conditionalFormatting.BackgroundColorR.Value, conditionalFormatting.BackgroundColorG.Value, conditionalFormatting.BackgroundColorB.Value);
                    }
                    if (conditionalFormatting.TextColorA.HasValue && conditionalFormatting.TextColorR.HasValue && conditionalFormatting.TextColorG.HasValue && conditionalFormatting.TextColorB.HasValue)
                    {
                        format.FontColorRGB = Syncfusion.Drawing.Color.FromArgb(conditionalFormatting.TextColorA.Value, conditionalFormatting.TextColorR.Value, conditionalFormatting.TextColorG.Value, conditionalFormatting.TextColorB.Value);
                    }
                }
            }

        }

        public double GetCellValueNumber(int sheetIndex, int rowIndex, int columnIndex)
        {
            return _workbook.Worksheets[sheetIndex][rowIndex + 1, columnIndex + 1].Number;
        }

        public double GetCellValueNumber(string sheetName, int rowIndex, int columnIndex)
        {
            return _workbook.Worksheets[sheetName][rowIndex + 1, columnIndex + 1].Number;
        }

        public string GetCellValueText(int sheetIndex, int rowIndex, int columnIndex)
        {
            return _workbook.Worksheets[sheetIndex][rowIndex + 1, columnIndex + 1].Text;
        }

        public string GetCellValueText(string sheetName, int rowIndex, int columnIndex)
        {
            return _workbook.Worksheets[sheetName][rowIndex + 1, columnIndex + 1].Text;
        }

        public DateTime GetCellValueDate(int sheetIndex, int rowIndex, int columnIndex)
        {
            return _workbook.Worksheets[sheetIndex][rowIndex + 1, columnIndex + 1].DateTime;
        }

        public DateTime GetCellValueDate(string sheetName, int rowIndex, int columnIndex)
        {
            return _workbook.Worksheets[sheetName][rowIndex + 1, columnIndex + 1].DateTime;
        }
        public void MetaData(string author, string subject, string title)
        {
            _workbook.Author = author;
            _workbook.BuiltInDocumentProperties.Subject = subject;
            _workbook.BuiltInDocumentProperties.Title = title;
        }

        //COMING SOON NEXT UPDATE
    }
}
