using NPOI.HSSF.Util;
using NPOI.OOXML.XSSF.UserModel;
using NPOI.SS.Format;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using NPOI.XWPF.UserModel;
using Soneta.Data.QueryDefinition;
using Soneta.Types;
using Syncfusion.XlsIO;
using Syncfusion.XlsIO.Implementation;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static Soneta.Place.WypElementNadgodziny;
using ICell = NPOI.SS.UserModel.ICell;
using IDataValidation = NPOI.SS.UserModel.IDataValidation;
using IFont = NPOI.SS.UserModel.IFont;
using IName = NPOI.SS.UserModel.IName;



namespace Trustsoft.ExcelOperation.Moje
{
    public class ExcelOperationNPOI : IExcelOperationService, IDisposable
    {
        private NPOI.SS.UserModel.IWorkbook _workbook;

        public ExcelOperationNPOI() { }

        public void AddCellFormula(int sheetIndex, int rowIndex, int columnIndex, string formula)
        {
            ISheet sheet = _workbook.GetSheetAt(sheetIndex);
            IRow row = sheet.GetRow(rowIndex) ?? sheet.CreateRow(rowIndex);
            ICell cell = row.GetCell(columnIndex) ?? row.CreateCell(columnIndex);

            cell.SetCellType(CellType.Formula);
            cell.SetCellFormula(formula);
        }

        public void AddCellFormula(string sheetName, int rowIndex, int columnIndex, string formula)
        {
            ISheet sheet = _workbook.GetSheet(sheetName);
            IRow row = sheet.GetRow(rowIndex) ?? sheet.CreateRow(rowIndex);
            ICell cell = row.GetCell(columnIndex) ?? row.CreateCell(columnIndex);

            cell.SetCellType(CellType.Formula);
            cell.SetCellFormula(formula);
        }

        public void AddCellFormula(int sheetIndex, int firstRowIndex, int firstColumnIndex, int lastRowIndex, int lastCoulmnIndex, string formula)
        {
            ISheet sheet = _workbook.GetSheetAt(sheetIndex);

            for (int r = firstRowIndex; r <= lastRowIndex; r++)
            {
                IRow row = sheet.GetRow(r) ?? sheet.CreateRow(r);
                for (int c = firstColumnIndex; c <= lastCoulmnIndex; c++)
                {
                    ICell cell = row.GetCell(c) ?? row.CreateCell(c);
                    cell.SetCellType(CellType.Formula);
                    cell.SetCellFormula(formula);
                }
            }
        }

        public void AddCellFormula(string sheetName, int firstRowIndex, int firstColumnIndex, int lastRowIndex, int lastCoulmnIndex, string formula)
        {
            ISheet sheet = _workbook.GetSheet(sheetName);

            for (int r = firstRowIndex; r <= lastRowIndex; r++)
            {
                IRow row = sheet.GetRow(r) ?? sheet.CreateRow(r);
                for (int c = firstColumnIndex; c <= lastCoulmnIndex; c++)
                {
                    ICell cell = row.GetCell(c) ?? row.CreateCell(c);
                    cell.SetCellType(CellType.Formula);
                    cell.SetCellFormula(formula);
                }
            }
        }

        public void AddCellValueCurrency(int sheetIndex, int rowIndex, int columnIndex, Currency currency)
        {
            ISheet sheet = _workbook.GetSheetAt(sheetIndex);
            IRow row = sheet.GetRow(rowIndex) ?? sheet.CreateRow(rowIndex);
            ICell cell = row.GetCell(columnIndex) ?? row.CreateCell(columnIndex);

            cell.SetCellType(CellType.String);
            cell.SetCellValue(currency.ToString());
        }

        public void AddCellValueCurrency(string sheetName, int rowIndex, int columnIndex, Currency currency)
        {
            ISheet sheet = _workbook.GetSheet(sheetName);
            IRow row = sheet.GetRow(rowIndex) ?? sheet.CreateRow(rowIndex);
            ICell cell = row.GetCell(columnIndex) ?? row.CreateCell(columnIndex);

            cell.SetCellType(CellType.String);
            cell.SetCellValue(currency.ToString());
        }

        public void AddCellValueDate(int sheetIndex, int rowIndex, int columnIndex, DateTime date)
        {
            ISheet sheet = _workbook.GetSheetAt(sheetIndex);
            IRow row = sheet.GetRow(rowIndex) ?? sheet.CreateRow(rowIndex);
            ICell cell = row.GetCell(columnIndex) ?? row.CreateCell(columnIndex);

            cell.SetCellValue(date);

            ICellStyle oldCellStyle = cell.CellStyle;
            ICellStyle newCellStyle = _workbook.CreateCellStyle();
            newCellStyle.CloneStyleFrom(oldCellStyle);
            newCellStyle.DataFormat = _workbook.CreateDataFormat().GetFormat("dd.mm.yyyy");

            cell.CellStyle = newCellStyle;
        }

        public void AddCellValueDate(string sheetName, int rowIndex, int columnIndex, DateTime date)
        {
            ISheet sheet = _workbook.GetSheet(sheetName);
            IRow row = sheet.GetRow(rowIndex) ?? sheet.CreateRow(rowIndex);
            ICell cell = row.GetCell(columnIndex) ?? row.CreateCell(columnIndex);

            cell.SetCellValue(date);

            ICellStyle oldCellStyle = cell.CellStyle;
            ICellStyle newCellStyle = _workbook.CreateCellStyle();
            newCellStyle.CloneStyleFrom(oldCellStyle);
            newCellStyle.DataFormat = _workbook.CreateDataFormat().GetFormat("dd.mm.yyyy");

            cell.CellStyle = newCellStyle;
        }

        public void AddCellValueDecimal(int sheetIndex, int rowIndex, int columnIndex, decimal value)
        {
            ISheet sheet = _workbook.GetSheetAt(sheetIndex);
            IRow row = sheet.GetRow(rowIndex) ?? sheet.CreateRow(rowIndex);
            ICell cell = row.GetCell(columnIndex) ?? row.CreateCell(columnIndex);

            cell.SetCellType(CellType.Numeric);
            cell.SetCellValue((double)value);
        }

        public void AddCellValueDecimal(string sheetName, int rowIndex, int columnIndex, decimal value)
        {
            ISheet sheet = _workbook.GetSheet(sheetName);
            IRow row = sheet.GetRow(rowIndex) ?? sheet.CreateRow(rowIndex);
            ICell cell = row.GetCell(columnIndex) ?? row.CreateCell(columnIndex);

            cell.SetCellType(CellType.Numeric);
            cell.SetCellValue((double)value);
        }

        public void AddCellValueDouble(int sheetIndex, int rowIndex, int columnIndex, double value)
        {
            ISheet sheet = _workbook.GetSheetAt(sheetIndex);
            IRow row = sheet.GetRow(rowIndex) ?? sheet.CreateRow(rowIndex);
            ICell cell = row.GetCell(columnIndex) ?? row.CreateCell(columnIndex);

            cell.SetCellType(CellType.Numeric);
            cell.SetCellValue(value);
        }

        public void AddCellValueDouble(string sheetName, int rowIndex, int columnIndex, double value)
        {
            ISheet sheet = _workbook.GetSheet(sheetName);
            IRow row = sheet.GetRow(rowIndex) ?? sheet.CreateRow(rowIndex);
            ICell cell = row.GetCell(columnIndex) ?? row.CreateCell(columnIndex);

            cell.SetCellType(CellType.Numeric);
            cell.SetCellValue(value);
        }

        public void AddCellValueFraction(int sheetIndex, int rowIndex, int columnIndex, Fraction fraction)
        {
            ISheet sheet = _workbook.GetSheetAt(sheetIndex);
            IRow row = sheet.GetRow(rowIndex) ?? sheet.CreateRow(rowIndex);
            ICell cell = row.GetCell(columnIndex) ?? row.CreateCell(columnIndex);

            cell.SetCellType(CellType.Numeric);
            cell.SetCellValue(fraction);
        }

        public void AddCellValueFraction(string sheetName, int rowIndex, int columnIndex, Fraction fraction)
        {
            ISheet sheet = _workbook.GetSheet(sheetName);
            IRow row = sheet.GetRow(rowIndex) ?? sheet.CreateRow(rowIndex);
            ICell cell = row.GetCell(columnIndex) ?? row.CreateCell(columnIndex);

            cell.SetCellType(CellType.Numeric);
            cell.SetCellValue(fraction);
        }

        public void AddCellValueInt(int sheetIndex, int rowIndex, int columnIndex, int value)
        {
            ISheet sheet = _workbook.GetSheetAt(sheetIndex);
            IRow row = sheet.GetRow(rowIndex) ?? sheet.CreateRow(rowIndex);
            ICell cell = row.GetCell(columnIndex) ?? row.CreateCell(columnIndex);

            cell.SetCellType(CellType.Numeric);
            cell.SetCellValue(value);
        }

        public void AddCellValueInt(string sheetName, int rowIndex, int columnIndex, int value)
        {
            ISheet sheet = _workbook.GetSheet(sheetName);
            IRow row = sheet.GetRow(rowIndex) ?? sheet.CreateRow(rowIndex);
            ICell cell = row.GetCell(columnIndex) ?? row.CreateCell(columnIndex);

            cell.SetCellType(CellType.Numeric);
            cell.SetCellValue(value);
        }

        public void AddCellValuePercent(int sheetIndex, int rowIndex, int columnIndex, Percent percent)
        {
            ISheet sheet = _workbook.GetSheetAt(sheetIndex);
            IRow row = sheet.GetRow(rowIndex) ?? sheet.CreateRow(rowIndex);
            ICell cell = row.GetCell(columnIndex) ?? row.CreateCell(columnIndex);

            cell.SetCellValue(((double)percent));

            ICellStyle oldCellStyle = cell.CellStyle;
            ICellStyle newCellStyle = _workbook.CreateCellStyle();
            newCellStyle.CloneStyleFrom(oldCellStyle);
            newCellStyle.DataFormat = _workbook.CreateDataFormat().GetFormat("0.00%");

            cell.CellStyle = newCellStyle;
        }

        public void AddCellValuePercent(string sheetName, int rowIndex, int columnIndex, Percent percent)
        {
            ISheet sheet = _workbook.GetSheet(sheetName);
            IRow row = sheet.GetRow(rowIndex) ?? sheet.CreateRow(rowIndex);
            ICell cell = row.GetCell(columnIndex) ?? row.CreateCell(columnIndex);

            cell.SetCellValue(((double)percent));

            ICellStyle oldCellStyle = cell.CellStyle;
            ICellStyle newCellStyle = _workbook.CreateCellStyle();
            newCellStyle.CloneStyleFrom(oldCellStyle);
            newCellStyle.DataFormat = _workbook.CreateDataFormat().GetFormat("0.00%");

            cell.CellStyle = newCellStyle;
        }

        public void AddCellValueText(int sheetIndex, int rowIndex, int columnIndex, string text)
        {
            ISheet sheet = _workbook.GetSheetAt(sheetIndex);
            IRow row = sheet.GetRow(rowIndex) ?? sheet.CreateRow(rowIndex);
            ICell cell = row.GetCell(columnIndex) ?? row.CreateCell(columnIndex);

            cell.SetCellType(CellType.String);
            cell.SetCellValue(text);
        }

        public void AddCellValueText(string sheetName, int rowIndex, int columnIndex, string text)
        {
            ISheet sheet = _workbook.GetSheet(sheetName);
            IRow row = sheet.GetRow(rowIndex) ?? sheet.CreateRow(rowIndex);
            ICell cell = row.GetCell(columnIndex) ?? row.CreateCell(columnIndex);

            cell.SetCellType(CellType.String);
            cell.SetCellValue(text);
        }

        public void AddCellValueTime(int sheetIndex, int rowIndex, int columnIndex, Time time)
        {
            ISheet sheet = _workbook.GetSheetAt(sheetIndex);
            IRow row = sheet.GetRow(rowIndex) ?? sheet.CreateRow(rowIndex);
            ICell cell = row.GetCell(columnIndex) ?? row.CreateCell(columnIndex);

            cell.SetCellType(CellType.String);
            cell.SetCellValue(time.ToString());
        }

        public void AddCellValueTime(string sheetName, int rowIndex, int columnIndex, Time time)
        {
            ISheet sheet = _workbook.GetSheet(sheetName);
            IRow row = sheet.GetRow(rowIndex) ?? sheet.CreateRow(rowIndex);
            ICell cell = row.GetCell(columnIndex) ?? row.CreateCell(columnIndex);

            cell.SetCellType(CellType.String);
            cell.SetCellValue(time.ToString());
        }

        public void AddColumn(int columnIndex, string sheetName)
        {
            ISheet sheet = _workbook.GetSheet(sheetName);
            foreach (IRow row in sheet)
            {
                if (row == null)
                {
                    continue;
                }
                for (int i = row.LastCellNum - 1; i >= columnIndex; i--)
                {
                    ICell oldCell = row.GetCell(i);
                    ICell newCell = row.CreateCell(i + 1);

                    if (oldCell != null)
                    {
                        newCell.CellStyle = oldCell.CellStyle;
                        newCell.SetCellType(oldCell.CellType);

                        switch (oldCell.CellType)
                        {
                            case CellType.Blank:
                                newCell.SetBlank();
                                break;
                            case CellType.String:
                                newCell.SetCellValue(oldCell.StringCellValue);
                                break;
                            case CellType.Numeric:
                                newCell.SetCellValue(oldCell.NumericCellValue);
                                break;
                            case CellType.Formula:
                                newCell.SetCellFormula(oldCell.CellFormula);
                                break;
                            case CellType.Boolean:
                                newCell.SetCellValue(oldCell.BooleanCellValue);
                                break;
                            case CellType.Error:
                                newCell.SetCellErrorValue(oldCell.ErrorCellValue);
                                break;
                            case CellType.Unknown:
                            default:
                                newCell.SetCellValue(oldCell.StringCellValue.ToString());
                                break;
                        }
                        row.RemoveCell(oldCell);
                    }
                }
                row.CreateCell(columnIndex);
            }
        }

        public void AddColumn(string sheetName, int columnIndex)
        {
            ISheet sheet = _workbook.GetSheet(sheetName);
            foreach (IRow row in sheet)
            {
                if (row == null)
                {
                    continue;
                }
                for (int i = row.LastCellNum - 1; i >= columnIndex; i--)
                {
                    ICell oldCell = row.GetCell(i);
                    ICell newCell = row.CreateCell(i + 1);

                    if (oldCell != null)
                    {
                        newCell.CellStyle = oldCell.CellStyle;
                        newCell.SetCellType(oldCell.CellType);

                        switch (oldCell.CellType)
                        {
                            case CellType.Blank:
                                newCell.SetBlank();
                                break;
                            case CellType.String:
                                newCell.SetCellValue(oldCell.StringCellValue);
                                break;
                            case CellType.Numeric:
                                newCell.SetCellValue(oldCell.NumericCellValue);
                                break;
                            case CellType.Formula:
                                newCell.SetCellFormula(oldCell.CellFormula);
                                break;
                            case CellType.Boolean:
                                newCell.SetCellValue(oldCell.BooleanCellValue);
                                break;
                            case CellType.Error:
                                newCell.SetCellErrorValue(oldCell.ErrorCellValue);
                                break;
                            case CellType.Unknown:
                            default:
                                newCell.SetCellValue(oldCell.StringCellValue.ToString());
                                break;

                        }
                        row.RemoveCell(oldCell);
                    }
                }
                row.CreateCell(columnIndex);
            }
        }

        public void AddColumn(int sheetIndex, int columnIndex)
        {
            ISheet sheet = _workbook.GetSheetAt(sheetIndex);
            foreach (IRow row in sheet)
            {
                if (row == null)
                {
                    continue;
                }
                for (int i = row.LastCellNum - 1; i >= columnIndex; i--)
                {
                    ICell oldCell = row.GetCell(i);
                    ICell newCell = row.CreateCell(i + 1);

                    if (oldCell != null)
                    {
                        newCell.CellStyle = oldCell.CellStyle;
                        newCell.SetCellType(oldCell.CellType);

                        switch (oldCell.CellType)
                        {
                            case CellType.Blank:
                                newCell.SetBlank();
                                break;
                            case CellType.String:
                                newCell.SetCellValue(oldCell.StringCellValue);
                                break;
                            case CellType.Numeric:
                                newCell.SetCellValue(oldCell.NumericCellValue);
                                break;
                            case CellType.Formula:
                                newCell.SetCellFormula(oldCell.CellFormula);
                                break;
                            case CellType.Boolean:
                                newCell.SetCellValue(oldCell.BooleanCellValue);
                                break;
                            case CellType.Error:
                                newCell.SetCellErrorValue(oldCell.ErrorCellValue);
                                break;
                            case CellType.Unknown:
                            default:
                                newCell.SetCellValue(oldCell.StringCellValue.ToString());
                                break;

                        }
                        row.RemoveCell(oldCell);
                    }
                }
                row.CreateCell(columnIndex);
            }
        }

        public void AddColumn(int columnIndex, string sheetName, int columnCount)
        {
            ISheet sheet = _workbook.GetSheet(sheetName);
            foreach (IRow row in sheet)
            {
                if (row == null)
                {
                    continue;
                }
                for (int i = row.LastCellNum - 1; i >= columnIndex; i--)
                {
                    ICell oldCell = row.GetCell(i);
                    ICell newCell = row.CreateCell(i + columnCount);
                    
                    if (oldCell != null)
                    {
                        newCell.CellStyle = oldCell.CellStyle;
                        newCell.SetCellType(oldCell.CellType);

                        switch(oldCell.CellType)
                        {
                            case CellType.Blank:
                                newCell.SetBlank();
                                break;
                            case CellType.String:
                                newCell.SetCellValue(oldCell.StringCellValue);
                                break;
                            case CellType.Numeric:
                                newCell.SetCellValue(oldCell.NumericCellValue);
                                break;
                            case CellType.Formula:
                                newCell.SetCellFormula(oldCell.CellFormula);
                                break;
                            case CellType.Boolean:
                                newCell.SetCellValue(oldCell.BooleanCellValue);
                                break;
                            case CellType.Error:
                                newCell.SetCellErrorValue(oldCell.ErrorCellValue);
                                break;
                            case CellType.Unknown:
                            default:
                                newCell.SetCellValue(oldCell.StringCellValue.ToString());
                                break;
                        }
                        row.RemoveCell(oldCell);
                    }
                }
                row.CreateCell(columnIndex);
            }
        }

        public void AddColumn(string sheetName, int columnIndex, int columnCount)
        {
            ISheet sheet = _workbook.GetSheet(sheetName);
            foreach (IRow row in sheet)
            {
                if (row == null)
                {
                    continue;
                }
                for (int i = row.LastCellNum - 1; i >= columnIndex; i--)
                {
                    ICell oldCell = row.GetCell(i);
                    ICell newCell = row.CreateCell(i + columnCount);

                    if (oldCell != null)
                    {
                        newCell.CellStyle = oldCell.CellStyle;
                        newCell.SetCellType(oldCell.CellType);

                        switch (oldCell.CellType)
                        {
                            case CellType.Blank:
                                newCell.SetBlank();
                                break;
                            case CellType.String:
                                newCell.SetCellValue(oldCell.StringCellValue);
                                break;
                            case CellType.Numeric:
                                newCell.SetCellValue(oldCell.NumericCellValue);
                                break;
                            case CellType.Formula:
                                newCell.SetCellFormula(oldCell.CellFormula);
                                break;
                            case CellType.Boolean:
                                newCell.SetCellValue(oldCell.BooleanCellValue);
                                break;
                            case CellType.Error:
                                newCell.SetCellErrorValue(oldCell.ErrorCellValue);
                                break;
                            case CellType.Unknown:
                            default:
                                newCell.SetCellValue(oldCell.StringCellValue.ToString());
                                break;
                        }
                        row.RemoveCell(oldCell);
                    }
                }
                row.CreateCell(columnIndex);
            }
        }

        public void AddColumn(int sheetIndex, int columnIndex, int columnCount)
        {
            ISheet sheet = _workbook.GetSheetAt(sheetIndex);
            foreach (IRow row in sheet)
            {
                if (row == null)
                {
                    continue;
                }
                for (int i = row.LastCellNum - 1; i >= columnIndex; i--)
                {
                    ICell oldCell = row.GetCell(i);
                    ICell newCell = row.CreateCell(i + columnCount);

                    if (oldCell != null)
                    {
                        newCell.CellStyle = oldCell.CellStyle;
                        newCell.SetCellType(oldCell.CellType);

                        switch (oldCell.CellType)
                        {
                            case CellType.Blank:
                                newCell.SetBlank();
                                break;
                            case CellType.String:
                                newCell.SetCellValue(oldCell.StringCellValue);
                                break;
                            case CellType.Numeric:
                                newCell.SetCellValue(oldCell.NumericCellValue);
                                break;
                            case CellType.Formula:
                                newCell.SetCellFormula(oldCell.CellFormula);
                                break;
                            case CellType.Boolean:
                                newCell.SetCellValue(oldCell.BooleanCellValue);
                                break;
                            case CellType.Error:
                                newCell.SetCellErrorValue(oldCell.ErrorCellValue);
                                break;
                            case CellType.Unknown:
                            default:
                                newCell.SetCellValue(oldCell.StringCellValue.ToString());
                                break;
                        }
                        row.RemoveCell(oldCell);
                    }
                }
                row.CreateCell(columnIndex);
            }
        }

        public void AddRow(int rowIndex, string sheetName)
        {
            ISheet sheet = _workbook.GetSheet(sheetName);
            sheet.ShiftRows(rowIndex, sheet.LastRowNum, 1);
            IRow newRow = sheet.CreateRow(rowIndex);
        }

        public void AddRow(string sheetName, int rowIndex)
        {
            ISheet sheet = _workbook.GetSheet(sheetName);
            sheet.ShiftRows(rowIndex, sheet.LastRowNum, 1);
            IRow newRow = sheet.CreateRow(rowIndex);
        }

        public void AddRow(int sheetIndex, int rowIndex)
        {
            ISheet sheet = _workbook.GetSheetAt(sheetIndex);
            sheet.ShiftRows(rowIndex, sheet.LastRowNum, 1);
            IRow newRow = sheet.CreateRow(rowIndex);
        }

        public void AddRow(int rowIndex, string sheetName, int rowCount)
        {
            ISheet sheet = _workbook.GetSheet(sheetName);
            sheet.ShiftRows(rowIndex, sheet.LastRowNum, rowCount);

            for (int i = 0; i < rowCount; i++)
            {
                IRow newRow = sheet.CreateRow(rowIndex + 1);
            }
        }

        public void AddRow(string sheetName, int rowIndex, int rowCount)
        {
            ISheet sheet = _workbook.GetSheet(sheetName);
            sheet.ShiftRows(rowIndex, sheet.LastRowNum, rowCount);

            for (int i = 0; i < rowCount; i++)
            {
                IRow newRow = sheet.CreateRow(rowIndex + 1);
            }
        }

        public void AddRow(int sheetIndex, int rowIndex, int rowCount)
        {
            ISheet sheet = _workbook.GetSheetAt(sheetIndex);
            sheet.ShiftRows(rowIndex, sheet.LastRowNum, rowCount);

            for (int i = 0; i < rowCount; i++)
            {
                IRow newRow = sheet.CreateRow(rowIndex + 1);
            }
        }

        public int AddWorksheet(string sheetName)
        {
            var sheet = _workbook.CreateSheet(sheetName);
            return _workbook.GetSheetIndex(sheet);
        }

        public void CellColor(int sheetIndex, int rowIndex, int columnIndex, int a, int r, int g, int b)
        {
            ISheet sheet = _workbook.GetSheetAt(sheetIndex);
            IRow row = sheet.GetRow(rowIndex) ?? sheet.CreateRow(rowIndex);
            ICell cell = row.GetCell(columnIndex) ?? row.CreateCell(columnIndex);

            byte[] bytesARGB = new byte[] {(byte)a,  (byte)r, (byte)g, (byte)b};
                

            XSSFColor xSSFColor = new XSSFColor(bytesARGB);
            ICellStyle oldCellStyle = cell.CellStyle;
            ICellStyle newCellStyle = _workbook.CreateCellStyle();
            newCellStyle.CloneStyleFrom(oldCellStyle);
            ((XSSFCellStyle)newCellStyle).SetFillForegroundColor(xSSFColor);
            newCellStyle.FillPattern = FillPattern.SolidForeground;

            cell.CellStyle = newCellStyle;
        }

        public void CellColor(string sheetName, int rowIndex, int columnIndex, int a, int r, int g, int b)
        {
            ISheet sheet = _workbook.GetSheet(sheetName);
            IRow row = sheet.GetRow(rowIndex) ?? sheet.CreateRow(rowIndex);
            ICell cell = row.GetCell(columnIndex) ?? row.CreateCell(columnIndex);

            byte[] bytesARGB = new byte[] { (byte)a, (byte)r, (byte)g, (byte)b };


            XSSFColor xSSFColor = new XSSFColor(bytesARGB);
            ICellStyle oldCellStyle = cell.CellStyle;
            ICellStyle newCellStyle = _workbook.CreateCellStyle();
            newCellStyle.CloneStyleFrom(oldCellStyle);
            ((XSSFCellStyle)newCellStyle).SetFillForegroundColor(xSSFColor);
            newCellStyle.FillPattern = FillPattern.SolidForeground;

            cell.CellStyle = newCellStyle;
        }

        public void CellColor(int sheetIndex, int firstRowIndex, int firstColumnIndex, int lastRowIndex, int lastColumnIndex, int a, int r, int g, int b)
        {
            ISheet sheet = _workbook.GetSheetAt(sheetIndex);

            byte[] bytesARGB = new byte[] { (byte)a, (byte)r, (byte)g, (byte)b };
            XSSFColor xSSFColor = new XSSFColor(bytesARGB);

            for (int i = firstRowIndex; i <= lastRowIndex; i++)
            {
                IRow row = sheet.GetRow(i) ?? sheet.CreateRow(i);
                for (int j = firstColumnIndex; j <= lastColumnIndex; j++)
                {
                    ICell cell = row.GetCell(j) ?? row.CreateCell(j);

                    ICellStyle oldCellStyle = cell.CellStyle;
                    ICellStyle newCellStyle = _workbook.CreateCellStyle();
                    newCellStyle.CloneStyleFrom(oldCellStyle);
                    ((XSSFCellStyle)newCellStyle).SetFillForegroundColor(xSSFColor);
                    newCellStyle.FillPattern = FillPattern.SolidForeground;

                    cell.CellStyle = newCellStyle;
                }
            }
  
        }

        public void CellColor(string sheetName, int firstRowIndex, int firstColumnIndex, int lastRowIndex, int lastColumnIndex, int a, int r, int g, int b)
        {
            ISheet sheet = _workbook.GetSheet(sheetName);

            byte[] bytesARGB = new byte[] { (byte)a, (byte)r, (byte)g, (byte)b };
            XSSFColor xSSFColor = new XSSFColor(bytesARGB);

            for (int i = firstRowIndex; i <= lastRowIndex; i++)
            {
                IRow row = sheet.GetRow(i) ?? sheet.CreateRow(i);
                for (int j = firstColumnIndex; j <= lastColumnIndex; j++)
                {
                    ICell cell = row.GetCell(j) ?? row.CreateCell(j);

                    ICellStyle oldCellStyle = cell.CellStyle;
                    ICellStyle newCellStyle = _workbook.CreateCellStyle();
                    newCellStyle.CloneStyleFrom(oldCellStyle);
                    ((XSSFCellStyle)newCellStyle).SetFillForegroundColor(xSSFColor);
                    newCellStyle.FillPattern = FillPattern.SolidForeground;

                    cell.CellStyle = newCellStyle;
                }
            }
        }

        public void ChangeNameWorksheet(string sheetName, string newName)
        {
            _workbook.SetSheetName(_workbook.GetSheetIndex(sheetName), newName);
        }

        public void ChangeNameWorksheet(int indexsheet, string newName)
        {
            _workbook.SetSheetName(indexsheet, newName);
        }

        public object CreateWorkbook()
        {
            _workbook = new XSSFWorkbook();
            _workbook.CreateSheet("Sheet1");
            return _workbook;
        }

        public object CreateWorkbook(string sheetName)
        {
            _workbook = new XSSFWorkbook();
            _workbook.CreateSheet(sheetName);
            return _workbook;
        }

        public List<Sheet> GetNameSheet()
        {
            List<Sheet> sheets = new List<Sheet>();
            int sheetIndex = 0;

            foreach (ISheet sheet in _workbook)
            {
                sheets.Add(new Sheet(sheetIndex, sheet.SheetName)); 
                sheetIndex++;
            }

            return sheets;
        }

        public void Dispose()
        {
            _workbook.Close();
        }

        public void DeleteColumn(string sheetName, int columnIndex)
        {
            ISheet sheet = _workbook.GetSheet(sheetName);
            int lastColumn = 0;
            foreach (IRow row in sheet)
            {
                if (row.LastCellNum > lastColumn)
                {
                    lastColumn = row.LastCellNum;
                }
            }

            foreach (IRow row in sheet)
            {
                for (int i = columnIndex; i < lastColumn; i++)
                {
                    ICell oldCell = row.GetCell(i + 1);
                    ICell newCell = row.GetCell(i);

                    if (oldCell != null)
                    {
                        if (newCell == null)
                        {
                            newCell = row.CreateCell(i);
                        }

                        newCell.CellStyle = oldCell.CellStyle;
                        newCell.SetCellType(oldCell.CellType);

                        switch (oldCell.CellType)
                        {
                            case CellType.Blank:
                                newCell.SetBlank();
                                break;
                            case CellType.Boolean:
                                newCell.SetCellValue(oldCell.BooleanCellValue);
                                break;
                            case CellType.String:
                                newCell.SetCellValue(oldCell.StringCellValue);
                                break;
                            case CellType.Numeric:
                                newCell.SetCellValue(oldCell.NumericCellValue);
                                break;
                            case CellType.Formula:
                                newCell.SetCellFormula(oldCell.CellFormula);
                                break;
                            case CellType.Error:
                                newCell.SetCellErrorValue(oldCell.ErrorCellValue);
                                break;
                            case CellType.Unknown:
                            default:
                                newCell.SetCellValue(oldCell.StringCellValue.ToString());
                                break;
                        }
                    }
                    else if (newCell != null)
                    {
                        row.RemoveCell(newCell);
                    }
                    ICell lastCell = row.GetCell(lastColumn);
                    if (lastCell != null)
                    {
                        row.RemoveCell(lastCell);
                    }
                }
            }
        }

        public void DeleteColumn(int columnIndex, string sheetName)
        {
            ISheet sheet = _workbook.GetSheet(sheetName);
            int lastColumn = 0;
            foreach (IRow row in sheet)
            {
                if (row.LastCellNum > lastColumn)
                {
                    lastColumn = row.LastCellNum;
                }
            }

            foreach (IRow row in sheet)
            {
                for (int i = columnIndex; i < lastColumn; i++)
                {
                    ICell oldCell = row.GetCell(i + 1);
                    ICell newCell = row.GetCell(i);

                    if (oldCell != null)
                    {
                        if (newCell == null)
                        {
                            newCell = row.CreateCell(i);
                        }

                        newCell.CellStyle = oldCell.CellStyle;
                        newCell.SetCellType(oldCell.CellType);

                        switch (oldCell.CellType)
                        {
                            case CellType.Blank:
                                newCell.SetBlank();
                                break;
                            case CellType.Boolean:
                                newCell.SetCellValue(oldCell.BooleanCellValue);
                                break;
                            case CellType.String:
                                newCell.SetCellValue(oldCell.StringCellValue);
                                break;
                            case CellType.Numeric:
                                newCell.SetCellValue(oldCell.NumericCellValue);
                                break;
                            case CellType.Formula:
                                newCell.SetCellFormula(oldCell.CellFormula);
                                break;
                            case CellType.Error:
                                newCell.SetCellErrorValue(oldCell.ErrorCellValue);
                                break;
                            case CellType.Unknown:
                            default:
                                newCell.SetCellValue(oldCell.StringCellValue.ToString());
                                break;
                        }
                    }
                    else if (newCell != null)
                    {
                        row.RemoveCell(newCell);
                    }
                    ICell lastCell = row.GetCell(lastColumn);
                    if (lastCell != null)
                    {
                        row.RemoveCell(lastCell);
                    }
                }
            }
        }

        public void DeleteColumn(int sheetIndex, int columnIndex)
        {
            ISheet sheet = _workbook.GetSheetAt(sheetIndex);
            int lastColumn = 0;
            foreach (IRow row in sheet)
            {
                if (row.LastCellNum > lastColumn)
                {
                    lastColumn = row.LastCellNum;
                }
            }

            foreach (IRow row in sheet)
            {
                for (int i = columnIndex; i < lastColumn; i++)
                {
                    ICell oldCell = row.GetCell(i + 1);
                    ICell newCell = row.GetCell(i);

                    if (oldCell != null)
                    {
                        if (newCell == null)
                        {
                            newCell = row.CreateCell(i);
                        }

                        newCell.CellStyle = oldCell.CellStyle;
                        newCell.SetCellType(oldCell.CellType);

                        switch (oldCell.CellType)
                        {
                            case CellType.Blank:
                                newCell.SetBlank();
                                break;
                            case CellType.Boolean:
                                newCell.SetCellValue(oldCell.BooleanCellValue);
                                break;
                            case CellType.String:
                                newCell.SetCellValue(oldCell.StringCellValue);
                                break;
                            case CellType.Numeric:
                                newCell.SetCellValue(oldCell.NumericCellValue);
                                break;
                            case CellType.Formula:
                                newCell.SetCellFormula(oldCell.CellFormula);
                                break;
                            case CellType.Error:
                                newCell.SetCellErrorValue(oldCell.ErrorCellValue);
                                break;
                            case CellType.Unknown:
                            default:
                                newCell.SetCellValue(oldCell.StringCellValue.ToString());
                                break;
                        }
                    }
                    else if (newCell != null)
                    {
                        row.RemoveCell(newCell);
                    }
                    ICell lastCell = row.GetCell(lastColumn);
                    if (lastCell != null)
                    {
                        row.RemoveCell(lastCell);
                    }
                }
            }
        }

        public void DeleteColumn(string sheetName, int columnIndex, int columnCount)
        {
            ISheet sheet = _workbook.GetSheet(sheetName);
            int lastColumn = 0;
            foreach (IRow row in sheet)
            {
                if (row.LastCellNum > lastColumn)
                {
                    lastColumn = row.LastCellNum;
                }
            }

            foreach (IRow row in sheet)
            {
                for (int i = columnIndex; i < lastColumn - columnCount; i++)
                {
                    ICell oldCell = row.GetCell(i + columnCount);
                    ICell newCell = row.GetCell(i) ?? row.CreateCell(i);

                    if (oldCell != null)
                    {
                        if (newCell == null)
                        {
                            newCell = row.CreateCell(i);
                        }

                        newCell.CellStyle = oldCell.CellStyle;
                        newCell.SetCellType(oldCell.CellType);

                        switch (oldCell.CellType)
                        {
                            case CellType.Blank:
                                newCell.SetBlank();
                                break;
                            case CellType.Boolean:
                                newCell.SetCellValue(oldCell.BooleanCellValue);
                                break;
                            case CellType.String:
                                newCell.SetCellValue(oldCell.StringCellValue);
                                break;
                            case CellType.Numeric:
                                newCell.SetCellValue(oldCell.NumericCellValue);
                                break;
                            case CellType.Formula:
                                newCell.SetCellFormula(oldCell.CellFormula);
                                break;
                            case CellType.Error:
                                newCell.SetCellErrorValue(oldCell.ErrorCellValue);
                                break;
                            case CellType.Unknown:
                            default:
                                newCell.SetCellValue(oldCell.StringCellValue.ToString());
                                break;
                        }
                    }
                    else if (newCell != null)
                    {
                        row.RemoveCell(newCell);
                    }
  
                }
                for (int i = lastColumn - columnCount; i < lastColumn; i++)
                {
                    ICell deleteCell = row.GetCell(i);
                    if (deleteCell != null)
                    {
                        row.RemoveCell(deleteCell);
                    }
                }
            }
        }

        public void DeleteColumn(int columnIndex, string sheetName, int columnCount)
        {
            ISheet sheet = _workbook.GetSheet(sheetName);
            int lastColumn = 0;
            foreach (IRow row in sheet)
            {
                if (row.LastCellNum > lastColumn)
                {
                    lastColumn = row.LastCellNum;
                }
            }

            foreach (IRow row in sheet)
            {
                for (int i = columnIndex; i < lastColumn - columnCount; i++)
                {
                    ICell oldCell = row.GetCell(i + columnCount);
                    ICell newCell = row.GetCell(i) ?? row.CreateCell(i);

                    if (oldCell != null)
                    {
                        if (newCell == null)
                        {
                            newCell = row.CreateCell(i);
                        }

                        newCell.CellStyle = oldCell.CellStyle;
                        newCell.SetCellType(oldCell.CellType);

                        switch (oldCell.CellType)
                        {
                            case CellType.Blank:
                                newCell.SetBlank();
                                break;
                            case CellType.Boolean:
                                newCell.SetCellValue(oldCell.BooleanCellValue);
                                break;
                            case CellType.String:
                                newCell.SetCellValue(oldCell.StringCellValue);
                                break;
                            case CellType.Numeric:
                                newCell.SetCellValue(oldCell.NumericCellValue);
                                break;
                            case CellType.Formula:
                                newCell.SetCellFormula(oldCell.CellFormula);
                                break;
                            case CellType.Error:
                                newCell.SetCellErrorValue(oldCell.ErrorCellValue);
                                break;
                            case CellType.Unknown:
                            default:
                                newCell.SetCellValue(oldCell.StringCellValue.ToString());
                                break;
                        }
                    }
                    else if (newCell != null)
                    {
                        row.RemoveCell(newCell);
                    }

                }
                for (int i = lastColumn - columnCount; i < lastColumn; i++)
                {
                    ICell deleteCell = row.GetCell(i);
                    if (deleteCell != null)
                    {
                        row.RemoveCell(deleteCell);
                    }
                }
            }
        }

        public void DeleteColumn(int sheetIndex, int columnIndex, int columnCount)
        {
            ISheet sheet = _workbook.GetSheetAt(sheetIndex);
            int lastColumn = 0;
            foreach (IRow row in sheet)
            {
                if (row.LastCellNum > lastColumn)
                {
                    lastColumn = row.LastCellNum;
                }
            }

            foreach (IRow row in sheet)
            {
                for (int i = columnIndex; i < lastColumn - columnCount; i++)
                {
                    ICell oldCell = row.GetCell(i + columnCount);
                    ICell newCell = row.GetCell(i) ?? row.CreateCell(i);

                    if (oldCell != null)
                    {
                        if (newCell == null)
                        {
                            newCell = row.CreateCell(i);
                        }

                        newCell.CellStyle = oldCell.CellStyle;
                        newCell.SetCellType(oldCell.CellType);

                        switch (oldCell.CellType)
                        {
                            case CellType.Blank:
                                newCell.SetBlank();
                                break;
                            case CellType.Boolean:
                                newCell.SetCellValue(oldCell.BooleanCellValue);
                                break;
                            case CellType.String:
                                newCell.SetCellValue(oldCell.StringCellValue);
                                break;
                            case CellType.Numeric:
                                newCell.SetCellValue(oldCell.NumericCellValue);
                                break;
                            case CellType.Formula:
                                newCell.SetCellFormula(oldCell.CellFormula);
                                break;
                            case CellType.Error:
                                newCell.SetCellErrorValue(oldCell.ErrorCellValue);
                                break;
                            case CellType.Unknown:
                            default:
                                newCell.SetCellValue(oldCell.StringCellValue.ToString());
                                break;
                        }
                    }
                    else if (newCell != null)
                    {
                        row.RemoveCell(newCell);
                    }

                }
                for (int i = lastColumn - columnCount; i < lastColumn; i++)
                {
                    ICell deleteCell = row.GetCell(i);
                    if (deleteCell != null)
                    {
                        row.RemoveCell(deleteCell);
                    }
                }
            }
        }

        public void DeleteRow(string sheetName, int rowIndex)
        {
            ISheet sheet = _workbook.GetSheet(sheetName);
            int lastRowNum = sheet.LastRowNum;

            if (rowIndex >= 0 && rowIndex < lastRowNum)
            {
                sheet.ShiftRows(rowIndex + 1, lastRowNum, -1);
            }

            IRow deleteRow = sheet.GetRow(lastRowNum);
            if (deleteRow != null)
            {
                sheet.RemoveRow(deleteRow);
            }

        }

        public void DeleteRow(int rowIndex, string sheetName)
        {
            ISheet sheet = _workbook.GetSheet(sheetName);
            int lastRowNum = sheet.LastRowNum;

            if (rowIndex >= 0 && rowIndex < lastRowNum)
            {
                sheet.ShiftRows(rowIndex + 1, lastRowNum, -1);
            }

            IRow deleteRow = sheet.GetRow(lastRowNum);
            if (deleteRow != null)
            {
                sheet.RemoveRow(deleteRow);
            }

        }

        public void DeleteRow(int sheetIndex, int rowIndex)
        {
            ISheet sheet = _workbook.GetSheetAt(sheetIndex);
            int lastRowNum = sheet.LastRowNum;

            if (rowIndex >= 0 && rowIndex < lastRowNum)
            {
                sheet.ShiftRows(rowIndex + 1, lastRowNum, -1);
            }

            IRow deleteRow = sheet.GetRow(lastRowNum);
            if (deleteRow != null)
            {
                sheet.RemoveRow(deleteRow);
            }
        }

        public void DeleteRow(string sheetName, int rowIndex, int rowCount)
        {
            ISheet sheet = _workbook.GetSheet(sheetName);
            int lastRowNum = sheet.LastRowNum;

            if (rowIndex >= 0 && rowIndex < lastRowNum)
            {
                sheet.ShiftRows(rowIndex + rowCount, lastRowNum, -rowCount);
            }
            for (int i = lastRowNum; i > lastRowNum - rowCount; i--)
            {
                IRow deleteRow = sheet.GetRow(i);
                if (deleteRow != null)
                {
                    sheet.RemoveRow(deleteRow);
                }
            }
                
        }

        public void DeleteRow(int rowIndex, string sheetName, int rowCount)
        {
            ISheet sheet = _workbook.GetSheet(sheetName);
            int lastRowNum = sheet.LastRowNum;

            if (rowIndex >= 0 && rowIndex < lastRowNum)
            {
                sheet.ShiftRows(rowIndex + rowCount, lastRowNum, -rowCount);
            }
            for (int i = lastRowNum; i > lastRowNum - rowCount; i--)
            {
                IRow deleteRow = sheet.GetRow(i);
                if (deleteRow != null)
                {
                    sheet.RemoveRow(deleteRow);
                }
            }
        }

        public void DeleteRow(int sheetIndex, int rowIndex, int rowCount)
        {
            ISheet sheet = _workbook.GetSheetAt(sheetIndex);
            int lastRowNum = sheet.LastRowNum;

            if (rowIndex >= 0 && rowIndex < lastRowNum)
            {
                sheet.ShiftRows(rowIndex + rowCount, lastRowNum, -rowCount);
            }
            for (int i = lastRowNum; i > lastRowNum - rowCount; i--)
            {
                IRow deleteRow = sheet.GetRow(i);
                if (deleteRow != null)
                {
                    sheet.RemoveRow(deleteRow);
                }
            }
        }

        public void DeleteWorksheet(string sheetName)
        {
            _workbook.RemoveSheetAt(_workbook.GetSheetIndex(sheetName));
        }

        public void DeleteWorksheet(int sheetIndex)
        {
            _workbook.RemoveSheetAt(sheetIndex);
        }

        

        

        public void HeightRow(int sheetIndex, int rowIndex, double height)
        {
            ISheet sheet = _workbook.GetSheetAt(sheetIndex);
            IRow row = sheet.GetRow(rowIndex) ?? sheet.CreateRow(rowIndex);
            row.HeightInPoints = (float)height;
        }

        public void HeightRow(int sheetIndex, int rowIndex, int rowCount, double height)
        {
            ISheet sheet = _workbook.GetSheetAt(sheetIndex);
            for (int i = 0; i < rowCount; i++)
            {
                IRow row = sheet.GetRow(rowIndex + i) ?? sheet.CreateRow(rowIndex + i);
                row.HeightInPoints = (float)height;
            }
        }
        public void HeightRow(int sheetIndex, int[] rowIndices, double height)
        {
            ISheet sheet = _workbook.GetSheetAt(sheetIndex);
            foreach (var index in rowIndices)
            {
                IRow row = sheet.GetRow(index) ?? sheet.CreateRow(index);
                row.HeightInPoints = (float)height;
            }
        }

        public void HeightRow(string sheetName, int rowIndex, int rowCount, double height)
        {
            ISheet sheet = _workbook.GetSheet(sheetName);
            for (int i = 0; i < rowCount; i++)
            {
                IRow row = sheet.GetRow(rowIndex + i) ?? sheet.CreateRow(rowIndex + i);
                row.HeightInPoints = (float)height;
            }
        }

        public void HeightRow(string sheetName, int[] rowIndices, double height)
        {
            ISheet sheet = _workbook.GetSheet(sheetName);
            foreach (var index in rowIndices)
            {
                IRow row = sheet.GetRow(index) ?? sheet.CreateRow(index);
                row.HeightInPoints = (float)height;
            }
        }

        public void HeightRow(string sheetName, int rowIndex, double height)
        {
            ISheet sheet = _workbook.GetSheet(sheetName);
            IRow row = sheet.GetRow(rowIndex) ?? sheet.CreateRow(rowIndex);
            row.HeightInPoints = (float)height;
        }

        public void MergeCells(int sheetIndex, int firstRowIndex, int firstColumnIndex, int lastRowIndex, int lastColumnIndex)
        {
            ISheet sheet = _workbook.GetSheetAt(sheetIndex);

            CellRangeAddress cellRangeAddress = new CellRangeAddress(firstRowIndex, lastRowIndex, firstColumnIndex, lastColumnIndex);
            sheet.AddMergedRegion(cellRangeAddress);
        }

        public void MergeCells(string sheetName, int firstRowIndex, int lastRowIndex, int firstColumnIndex, int lastColumnIndex)
        {
            ISheet sheet = _workbook.GetSheet(sheetName);

            CellRangeAddress cellRangeAddress = new CellRangeAddress(firstRowIndex, lastRowIndex, firstColumnIndex, lastColumnIndex);
            sheet.AddMergedRegion(cellRangeAddress);
        }

        public void SetBorder(BorderIndex[] borderIndex, LinesIndex lineIndex, string nameSheet, int rowIndex, int columnIndex)
        {
            ISheet sheet = _workbook.GetSheet(nameSheet);
            IRow row = sheet.GetRow(rowIndex) ?? sheet.CreateRow(rowIndex);
            ICell cell = row.GetCell(columnIndex) ?? row.CreateCell(columnIndex);

            byte[] bytesARGB = new byte[] { (byte)255, (byte)0, (byte)0, (byte)0 };
            XSSFColor xSSFColor = new XSSFColor(bytesARGB);

            ICellStyle oldCellStyle = cell.CellStyle;
            ICellStyle newCellStyle = _workbook.CreateCellStyle();
            newCellStyle.CloneStyleFrom(oldCellStyle);

            foreach (var itemBorder in borderIndex)
            {
                var npoiLinesIndex = NpoiHelper.ConvertFromLineStyleNpoi(lineIndex, out bool isEmpty);
                if (!isEmpty)
                {
                   
                    switch (itemBorder)
                    {
                        case BorderIndex.Bottom:
                            newCellStyle.BorderBottom = npoiLinesIndex;
                            ((XSSFCellStyle)newCellStyle).SetBottomBorderColor(xSSFColor);
                            break;
                        case BorderIndex.Top:
                            newCellStyle.BorderTop = npoiLinesIndex;
                            ((XSSFCellStyle)newCellStyle).SetTopBorderColor(xSSFColor);
                            break;
                        case BorderIndex.Left:
                            newCellStyle.BorderLeft = npoiLinesIndex;
                            ((XSSFCellStyle)newCellStyle).SetLeftBorderColor(xSSFColor);
                            break;
                        case BorderIndex.Right:
                            newCellStyle.BorderRight = npoiLinesIndex;
                            ((XSSFCellStyle)newCellStyle).SetRightBorderColor(xSSFColor);
                            break;
                        case BorderIndex.All:
                            newCellStyle.BorderBottom = npoiLinesIndex;
                            newCellStyle.BorderTop = npoiLinesIndex;
                            newCellStyle.BorderLeft = npoiLinesIndex;
                            newCellStyle.BorderRight = npoiLinesIndex;
                            ((XSSFCellStyle)newCellStyle).SetBottomBorderColor(xSSFColor);
                            ((XSSFCellStyle)newCellStyle).SetTopBorderColor(xSSFColor);
                            ((XSSFCellStyle)newCellStyle).SetLeftBorderColor(xSSFColor);
                            ((XSSFCellStyle)newCellStyle).SetRightBorderColor(xSSFColor);
                            break;
                        case BorderIndex.None:
                            break;
                        default:
                            throw new NotImplementedException();

                    }
                  
                }
                cell.CellStyle = newCellStyle;
            }
        }

        public void SetBorder(BorderIndex[] borderIndex, LinesIndex lineIndex, int sheetIndex, int rowIndex, int columnIndex)
        {
            ISheet sheet = _workbook.GetSheetAt(sheetIndex);
            IRow row = sheet.GetRow(rowIndex) ?? sheet.CreateRow(rowIndex);
            ICell cell = row.GetCell(columnIndex) ?? row.CreateCell(columnIndex);

            byte[] bytesARGB = new byte[] { (byte)255, (byte)0, (byte)0, (byte)0 };
            XSSFColor xSSFColor = new XSSFColor(bytesARGB);

            ICellStyle oldCellStyle = cell.CellStyle;
            ICellStyle newCellStyle = _workbook.CreateCellStyle();
            newCellStyle.CloneStyleFrom(oldCellStyle);

            foreach (var itemBorder in borderIndex)
            {
                var npoiLinesIndex = NpoiHelper.ConvertFromLineStyleNpoi(lineIndex, out bool isEmpty);
                if (!isEmpty)
                {

                    switch (itemBorder)
                    {
                        case BorderIndex.Bottom:
                            newCellStyle.BorderBottom = npoiLinesIndex;
                            ((XSSFCellStyle)newCellStyle).SetBottomBorderColor(xSSFColor);
                            break;
                        case BorderIndex.Top:
                            newCellStyle.BorderTop = npoiLinesIndex;
                            ((XSSFCellStyle)newCellStyle).SetTopBorderColor(xSSFColor);
                            break;
                        case BorderIndex.Left:
                            newCellStyle.BorderLeft = npoiLinesIndex;
                            ((XSSFCellStyle)newCellStyle).SetLeftBorderColor(xSSFColor);
                            break;
                        case BorderIndex.Right:
                            newCellStyle.BorderRight = npoiLinesIndex;
                            ((XSSFCellStyle)newCellStyle).SetRightBorderColor(xSSFColor);
                            break;
                        case BorderIndex.All:
                            newCellStyle.BorderBottom = npoiLinesIndex;
                            newCellStyle.BorderTop = npoiLinesIndex;
                            newCellStyle.BorderLeft = npoiLinesIndex;
                            newCellStyle.BorderRight = npoiLinesIndex;
                            ((XSSFCellStyle)newCellStyle).SetBottomBorderColor(xSSFColor);
                            ((XSSFCellStyle)newCellStyle).SetTopBorderColor(xSSFColor);
                            ((XSSFCellStyle)newCellStyle).SetLeftBorderColor(xSSFColor);
                            ((XSSFCellStyle)newCellStyle).SetRightBorderColor(xSSFColor);
                            break;
                        case BorderIndex.None:
                            break;
                        default:
                            throw new NotImplementedException();

                    }

                }
                cell.CellStyle = newCellStyle;
            }
        }

        public void SetBorder(BorderIndex[] borderIndex, LinesIndex lineIndex, string nameSheet, int firstRowIndex, int firstColumnIndex, int lastRowIndex, int lastColumnIndex, bool allRange)
        {
            ISheet sheet = _workbook.GetSheet(nameSheet);

            if (allRange)
            {
                for (int i = firstRowIndex; i <= lastRowIndex; i++)
                {
                    IRow row = sheet.GetRow(i) ?? sheet.CreateRow(i);
                    for (int j = firstColumnIndex; j <= lastColumnIndex; j++)
                    {
                        ICell cell = row.GetCell(j) ?? row.CreateCell(j);

                        byte[] bytesARGB = new byte[] { (byte)255, (byte)0, (byte)0, (byte)0 };
                        XSSFColor xSSFColor = new XSSFColor(bytesARGB);

                        ICellStyle oldCellStyle = cell.CellStyle;
                        ICellStyle newCellStyle = _workbook.CreateCellStyle();
                        newCellStyle.CloneStyleFrom(oldCellStyle);

                        foreach (var itemBorder in borderIndex)
                        {
                            var npoiLinesIndex = NpoiHelper.ConvertFromLineStyleNpoi(lineIndex, out bool isEmpty);
                            if (!isEmpty)
                            {
                                switch (itemBorder)
                                {
                                    case BorderIndex.Bottom:
                                        newCellStyle.BorderBottom = npoiLinesIndex;
                                        ((XSSFCellStyle)newCellStyle).SetBottomBorderColor(xSSFColor);
                                        break;
                                    case BorderIndex.Top:
                                        newCellStyle.BorderTop = npoiLinesIndex;
                                        ((XSSFCellStyle)newCellStyle).SetTopBorderColor(xSSFColor);
                                        break;
                                    case BorderIndex.Left:
                                        newCellStyle.BorderLeft = npoiLinesIndex;
                                        ((XSSFCellStyle)newCellStyle).SetLeftBorderColor(xSSFColor);
                                        break;
                                    case BorderIndex.Right:
                                        newCellStyle.BorderRight = npoiLinesIndex;
                                        ((XSSFCellStyle)newCellStyle).SetRightBorderColor(xSSFColor);
                                        break;
                                    case BorderIndex.All:
                                        newCellStyle.BorderBottom = npoiLinesIndex;
                                        newCellStyle.BorderTop = npoiLinesIndex;
                                        newCellStyle.BorderLeft = npoiLinesIndex;
                                        newCellStyle.BorderRight = npoiLinesIndex;
                                        ((XSSFCellStyle)newCellStyle).SetBottomBorderColor(xSSFColor);
                                        ((XSSFCellStyle)newCellStyle).SetTopBorderColor(xSSFColor);
                                        ((XSSFCellStyle)newCellStyle).SetLeftBorderColor(xSSFColor);
                                        ((XSSFCellStyle)newCellStyle).SetRightBorderColor(xSSFColor);
                                        break;
                                    case BorderIndex.None:
                                        break;
                                    default:
                                        throw new NotImplementedException();

                                }
                              
                            }
                            cell.CellStyle = newCellStyle;
                        }
                    }

                }
            }
            else
            {
                for (int i = firstRowIndex; i <= lastRowIndex; i++)
                {
                    IRow row = sheet.GetRow(i) ?? sheet.CreateRow(i);
                    for (int j = firstColumnIndex; j <= lastColumnIndex; j++)
                    {
                        ICell cell = row.GetCell(j) ?? row.CreateCell(j);

                        byte[] bytesARGB = new byte[] { (byte)255, (byte)0, (byte)0, (byte)0 };
                        XSSFColor xSSFColor = new XSSFColor(bytesARGB);

                        ICellStyle oldCellStyle = cell.CellStyle;
                        ICellStyle newCellStyle = _workbook.CreateCellStyle();
                        newCellStyle.CloneStyleFrom(oldCellStyle);

                        foreach (var itemBorder in borderIndex)
                        {
                            var npoiLinesIndex = NpoiHelper.ConvertFromLineStyleNpoi(lineIndex, out bool isEmpty);
                            if (!isEmpty)
                            {
                                if (i == firstRowIndex && (itemBorder == BorderIndex.Top || itemBorder == BorderIndex.All))
                                {
                                    newCellStyle.BorderTop = npoiLinesIndex;
                                    ((XSSFCellStyle)newCellStyle).SetTopBorderColor(xSSFColor);
                                }
                                if (i == lastRowIndex && (itemBorder == BorderIndex.Bottom || itemBorder == BorderIndex.All))
                                {
                                    newCellStyle.BorderBottom = npoiLinesIndex;
                                    ((XSSFCellStyle)newCellStyle).SetBottomBorderColor(xSSFColor);
                                }
                                if (j == firstColumnIndex && (itemBorder == BorderIndex.Left || itemBorder == BorderIndex.All))
                                {
                                    newCellStyle.BorderLeft = npoiLinesIndex;
                                    ((XSSFCellStyle)newCellStyle).SetLeftBorderColor(xSSFColor);
                                }
                                if (j == lastColumnIndex && (itemBorder == BorderIndex.Right || itemBorder == BorderIndex.All))
                                {
                                    newCellStyle.BorderRight = npoiLinesIndex;
                                    ((XSSFCellStyle)newCellStyle).SetRightBorderColor(xSSFColor);
                                }
                            
                            }
                            cell.CellStyle = newCellStyle;
                        }
                    }

                }
            }
            
        }

        public void SetBorder(BorderIndex[] borderIndex, LinesIndex lineIndex, int sheetIndex, int firstRowIndex, int firstColumnIndex, int lastRowIndex, int lastColumnIndex, bool allRange)
        {
            ISheet sheet = _workbook.GetSheetAt(sheetIndex);

            if (allRange)
            {
                for (int i = firstRowIndex; i <= lastRowIndex; i++)
                {
                    IRow row = sheet.GetRow(i) ?? sheet.CreateRow(i);
                    for (int j = firstColumnIndex; j <= lastColumnIndex; j++)
                    {
                        ICell cell = row.GetCell(j) ?? row.CreateCell(j);

                        byte[] bytesARGB = new byte[] { (byte)255, (byte)0, (byte)0, (byte)0 };
                        XSSFColor xSSFColor = new XSSFColor(bytesARGB);

                        ICellStyle oldCellStyle = cell.CellStyle;
                        ICellStyle newCellStyle = _workbook.CreateCellStyle();
                        newCellStyle.CloneStyleFrom(oldCellStyle);

                        foreach (var itemBorder in borderIndex)
                        {
                            var npoiLinesIndex = NpoiHelper.ConvertFromLineStyleNpoi(lineIndex, out bool isEmpty);
                            if (!isEmpty)
                            {
                                switch (itemBorder)
                                {
                                    case BorderIndex.Bottom:
                                        newCellStyle.BorderBottom = npoiLinesIndex;
                                        ((XSSFCellStyle)newCellStyle).SetBottomBorderColor(xSSFColor);
                                        break;
                                    case BorderIndex.Top:
                                        newCellStyle.BorderTop = npoiLinesIndex;
                                        ((XSSFCellStyle)newCellStyle).SetTopBorderColor(xSSFColor);
                                        break;
                                    case BorderIndex.Left:
                                        newCellStyle.BorderLeft = npoiLinesIndex;
                                        ((XSSFCellStyle)newCellStyle).SetLeftBorderColor(xSSFColor);
                                        break;
                                    case BorderIndex.Right:
                                        newCellStyle.BorderRight = npoiLinesIndex;
                                        ((XSSFCellStyle)newCellStyle).SetRightBorderColor(xSSFColor);
                                        break;
                                    case BorderIndex.All:
                                        newCellStyle.BorderBottom = npoiLinesIndex;
                                        newCellStyle.BorderTop = npoiLinesIndex;
                                        newCellStyle.BorderLeft = npoiLinesIndex;
                                        newCellStyle.BorderRight = npoiLinesIndex;
                                        ((XSSFCellStyle)newCellStyle).SetBottomBorderColor(xSSFColor);
                                        ((XSSFCellStyle)newCellStyle).SetTopBorderColor(xSSFColor);
                                        ((XSSFCellStyle)newCellStyle).SetLeftBorderColor(xSSFColor);
                                        ((XSSFCellStyle)newCellStyle).SetRightBorderColor(xSSFColor);
                                        break;
                                    case BorderIndex.None:
                                        break;
                                    default:
                                        throw new NotImplementedException();

                                }

                            }
                            cell.CellStyle = newCellStyle;
                        }
                    }

                }
            }
            else
            {
                for (int i = firstRowIndex; i <= lastRowIndex; i++)
                {
                    IRow row = sheet.GetRow(i) ?? sheet.CreateRow(i);
                    for (int j = firstColumnIndex; j <= lastColumnIndex; j++)
                    {
                        ICell cell = row.GetCell(j) ?? row.CreateCell(j);

                        byte[] bytesARGB = new byte[] { (byte)255, (byte)0, (byte)0, (byte)0 };
                        XSSFColor xSSFColor = new XSSFColor(bytesARGB);

                        ICellStyle oldCellStyle = cell.CellStyle;
                        ICellStyle newCellStyle = _workbook.CreateCellStyle();
                        newCellStyle.CloneStyleFrom(oldCellStyle);

                        foreach (var itemBorder in borderIndex)
                        {
                            var npoiLinesIndex = NpoiHelper.ConvertFromLineStyleNpoi(lineIndex, out bool isEmpty);
                            if (!isEmpty)
                            {
                                if (i == firstRowIndex && (itemBorder == BorderIndex.Top || itemBorder == BorderIndex.All))
                                {
                                    newCellStyle.BorderTop = npoiLinesIndex;
                                    ((XSSFCellStyle)newCellStyle).SetTopBorderColor(xSSFColor);
                                }
                                if (i == lastRowIndex && (itemBorder == BorderIndex.Bottom || itemBorder == BorderIndex.All))
                                {
                                    newCellStyle.BorderBottom = npoiLinesIndex;
                                    ((XSSFCellStyle)newCellStyle).SetBottomBorderColor(xSSFColor);
                                }
                                if (j == firstColumnIndex && (itemBorder == BorderIndex.Left || itemBorder == BorderIndex.All))
                                {
                                    newCellStyle.BorderLeft = npoiLinesIndex;
                                    ((XSSFCellStyle)newCellStyle).SetLeftBorderColor(xSSFColor);
                                }
                                if (j == lastColumnIndex && (itemBorder == BorderIndex.Right || itemBorder == BorderIndex.All))
                                {
                                    newCellStyle.BorderRight = npoiLinesIndex;
                                    ((XSSFCellStyle)newCellStyle).SetRightBorderColor(xSSFColor);
                                }

                            }
                            cell.CellStyle = newCellStyle;
                        }
                    }

                }
            }
        }

        public void SetBorder(BorderIndex[] borderIndex, LinesIndex lineIndex, string nameSheet, int rowIndex, int columnIndex, int a, int r, int g, int b)
        {
            ISheet sheet = _workbook.GetSheet(nameSheet);
            IRow row = sheet.GetRow(rowIndex) ?? sheet.CreateRow(rowIndex);
            ICell cell = row.GetCell(columnIndex) ?? row.CreateCell(columnIndex);

            ICellStyle oldCellStyle = cell.CellStyle;
            ICellStyle newCellStyle = _workbook.CreateCellStyle();
            newCellStyle.CloneStyleFrom(oldCellStyle);

            byte[] bytesARGB = new byte[] { (byte)a, (byte)r, (byte)g, (byte)b };
            XSSFColor xSSFColor = new XSSFColor(bytesARGB);

            foreach (var itemBorder in borderIndex)
            {
                var npoiLinesIndex = NpoiHelper.ConvertFromLineStyleNpoi(lineIndex, out bool isEmpty);
                if (!isEmpty)
                {

                    switch (itemBorder)
                    {
                        case BorderIndex.Bottom:
                            newCellStyle.BorderBottom = npoiLinesIndex;
                            ((XSSFCellStyle)newCellStyle).SetBottomBorderColor(xSSFColor);
                            break;
                        case BorderIndex.Top:
                            newCellStyle.BorderTop = npoiLinesIndex;
                            ((XSSFCellStyle)newCellStyle).SetTopBorderColor(xSSFColor);
                            break;
                        case BorderIndex.Left:
                            newCellStyle.BorderLeft = npoiLinesIndex;
                            ((XSSFCellStyle)newCellStyle).SetLeftBorderColor(xSSFColor);
                            break;
                        case BorderIndex.Right:
                            newCellStyle.BorderRight = npoiLinesIndex;
                            ((XSSFCellStyle)newCellStyle).SetRightBorderColor(xSSFColor);
                            break;
                        case BorderIndex.All:
                            newCellStyle.BorderBottom = npoiLinesIndex;
                            newCellStyle.BorderTop = npoiLinesIndex;
                            newCellStyle.BorderLeft = npoiLinesIndex;
                            newCellStyle.BorderRight = npoiLinesIndex;
                            ((XSSFCellStyle)newCellStyle).SetBottomBorderColor(xSSFColor);
                            ((XSSFCellStyle)newCellStyle).SetTopBorderColor(xSSFColor);
                            ((XSSFCellStyle)newCellStyle).SetLeftBorderColor(xSSFColor);
                            ((XSSFCellStyle)newCellStyle).SetRightBorderColor(xSSFColor);
                            break;
                        case BorderIndex.None:
                            break;
                        default:
                            throw new NotImplementedException();

                    }
                }
                cell.CellStyle = newCellStyle;
            }
        }

        public void SetBorder(BorderIndex[] borderIndex, LinesIndex lineIndex, int sheetIndex, int rowIndex, int columnIndex, int a, int r, int g, int b)
        {
            ISheet sheet = _workbook.GetSheetAt(sheetIndex);
            IRow row = sheet.GetRow(rowIndex) ?? sheet.CreateRow(rowIndex);
            ICell cell = row.GetCell(columnIndex) ?? row.CreateCell(columnIndex);

            ICellStyle oldCellStyle = cell.CellStyle;
            ICellStyle newCellStyle = _workbook.CreateCellStyle();
            newCellStyle.CloneStyleFrom(oldCellStyle);

            byte[] bytesARGB = new byte[] { (byte)a, (byte)r, (byte)g, (byte)b };
            XSSFColor xSSFColor = new XSSFColor(bytesARGB);

            foreach (var itemBorder in borderIndex)
            {
                var npoiLinesIndex = NpoiHelper.ConvertFromLineStyleNpoi(lineIndex, out bool isEmpty);
                if (!isEmpty)
                {

                    switch (itemBorder)
                    {
                        case BorderIndex.Bottom:
                            newCellStyle.BorderBottom = npoiLinesIndex;
                            ((XSSFCellStyle)newCellStyle).SetBottomBorderColor(xSSFColor);
                            break;
                        case BorderIndex.Top:
                            newCellStyle.BorderTop = npoiLinesIndex;
                            ((XSSFCellStyle)newCellStyle).SetTopBorderColor(xSSFColor);
                            break;
                        case BorderIndex.Left:
                            newCellStyle.BorderLeft = npoiLinesIndex;
                            ((XSSFCellStyle)newCellStyle).SetLeftBorderColor(xSSFColor);
                            break;
                        case BorderIndex.Right:
                            newCellStyle.BorderRight = npoiLinesIndex;
                            ((XSSFCellStyle)newCellStyle).SetRightBorderColor(xSSFColor);
                            break;
                        case BorderIndex.All:
                            newCellStyle.BorderBottom = npoiLinesIndex;
                            newCellStyle.BorderTop = npoiLinesIndex;
                            newCellStyle.BorderLeft = npoiLinesIndex;
                            newCellStyle.BorderRight = npoiLinesIndex;
                            ((XSSFCellStyle)newCellStyle).SetBottomBorderColor(xSSFColor);
                            ((XSSFCellStyle)newCellStyle).SetTopBorderColor(xSSFColor);
                            ((XSSFCellStyle)newCellStyle).SetLeftBorderColor(xSSFColor);
                            ((XSSFCellStyle)newCellStyle).SetRightBorderColor(xSSFColor);
                            break;
                        case BorderIndex.None:
                            break;
                        default:
                            throw new NotImplementedException();

                    }
                }
                cell.CellStyle = newCellStyle;
            }
        }

        public void SetBorder(BorderIndex[] borderIndex, LinesIndex lineIndex, string nameSheet, int firstRowIndex, int firstColumnIndex, int lastRowIndex, int lastColumnIndex, int a, int r, int g, int b, bool allRange)
        {
            ISheet sheet = _workbook.GetSheet(nameSheet);

            if (allRange)
            {
                for (int i = firstRowIndex; i <= lastRowIndex; i++)
                {
                    IRow row = sheet.GetRow(i) ?? sheet.CreateRow(i);
                    for (int j = firstColumnIndex; j <= lastColumnIndex; j++)
                    {
                        ICell cell = row.GetCell(j) ?? row.CreateCell(j);

                        byte[] bytesARGB = new byte[] { (byte)a, (byte)r, (byte)g, (byte)b };
                        XSSFColor xSSFColor = new XSSFColor(bytesARGB);

                        ICellStyle oldCellStyle = cell.CellStyle;
                        ICellStyle newCellStyle = _workbook.CreateCellStyle();
                        newCellStyle.CloneStyleFrom(oldCellStyle);

                        foreach (var itemBorder in borderIndex)
                        {
                            var npoiLinesIndex = NpoiHelper.ConvertFromLineStyleNpoi(lineIndex, out bool isEmpty);
                            if (!isEmpty)
                            {
                                switch (itemBorder)
                                {
                                    case BorderIndex.Bottom:
                                        newCellStyle.BorderBottom = npoiLinesIndex;
                                        ((XSSFCellStyle)newCellStyle).SetBottomBorderColor(xSSFColor);
                                        break;
                                    case BorderIndex.Top:
                                        newCellStyle.BorderTop = npoiLinesIndex;
                                        ((XSSFCellStyle)newCellStyle).SetTopBorderColor(xSSFColor);
                                        break;
                                    case BorderIndex.Left:
                                        newCellStyle.BorderLeft = npoiLinesIndex;
                                        ((XSSFCellStyle)newCellStyle).SetLeftBorderColor(xSSFColor);
                                        break;
                                    case BorderIndex.Right:
                                        newCellStyle.BorderRight = npoiLinesIndex;
                                        ((XSSFCellStyle)newCellStyle).SetRightBorderColor(xSSFColor);
                                        break;
                                    case BorderIndex.All:
                                        newCellStyle.BorderBottom = npoiLinesIndex;
                                        newCellStyle.BorderTop = npoiLinesIndex;
                                        newCellStyle.BorderLeft = npoiLinesIndex;
                                        newCellStyle.BorderRight = npoiLinesIndex;
                                        ((XSSFCellStyle)newCellStyle).SetBottomBorderColor(xSSFColor);
                                        ((XSSFCellStyle)newCellStyle).SetTopBorderColor(xSSFColor);
                                        ((XSSFCellStyle)newCellStyle).SetLeftBorderColor(xSSFColor);
                                        ((XSSFCellStyle)newCellStyle).SetRightBorderColor(xSSFColor);
                                        break;
                                    case BorderIndex.None:
                                        break;
                                    default:
                                        throw new NotImplementedException();

                                }
                            }
                            cell.CellStyle = newCellStyle;
                        }
                    }

                }
            }
            else
            {
                for (int i = firstRowIndex; i <= lastRowIndex; i++)
                {
                    IRow row = sheet.GetRow(i) ?? sheet.CreateRow(i);
                    for (int j = firstColumnIndex; j <= lastColumnIndex; j++)
                    {
                        ICell cell = row.GetCell(j) ?? row.CreateCell(j);

                        byte[] bytesARGB = new byte[] { (byte)a, (byte)r, (byte)g, (byte)b };
                        XSSFColor xSSFColor = new XSSFColor(bytesARGB);

                        ICellStyle oldCellStyle = cell.CellStyle;
                        ICellStyle newCellStyle = _workbook.CreateCellStyle();
                        newCellStyle.CloneStyleFrom(oldCellStyle);

                        foreach (var itemBorder in borderIndex)
                        {
                            var npoiLinesIndex = NpoiHelper.ConvertFromLineStyleNpoi(lineIndex, out bool isEmpty);
                            if (!isEmpty)
                            {
                                if (i == firstRowIndex && (itemBorder == BorderIndex.Top || itemBorder == BorderIndex.All))
                                {
                                    newCellStyle.BorderTop = npoiLinesIndex;
                                    ((XSSFCellStyle)newCellStyle).SetTopBorderColor(xSSFColor);
                                }
                                if (i == lastRowIndex && (itemBorder == BorderIndex.Bottom || itemBorder == BorderIndex.All))
                                {
                                    newCellStyle.BorderBottom = npoiLinesIndex;
                                    ((XSSFCellStyle)newCellStyle).SetBottomBorderColor(xSSFColor);
                                }
                                if (j == firstColumnIndex && (itemBorder == BorderIndex.Left || itemBorder == BorderIndex.All))
                                {
                                    newCellStyle.BorderLeft = npoiLinesIndex;
                                    ((XSSFCellStyle)newCellStyle).SetLeftBorderColor(xSSFColor);
                                }
                                if (j == lastColumnIndex && (itemBorder == BorderIndex.Right || itemBorder == BorderIndex.All))
                                {
                                    newCellStyle.BorderRight = npoiLinesIndex;
                                    ((XSSFCellStyle)newCellStyle).SetRightBorderColor(xSSFColor);
                                }
                            }
                            cell.CellStyle = newCellStyle;
                        }
                    }

                }
            }
        }

        public void SetBorder(BorderIndex[] borderIndex, LinesIndex lineIndex, int sheetIndex, int firstRowIndex, int firstColumnIndex, int lastRowIndex, int lastColumnIndex, int a, int r, int g, int b, bool allRange)
        {
            ISheet sheet = _workbook.GetSheetAt(sheetIndex);

            if (allRange)
            {
                for (int i = firstRowIndex; i <= lastRowIndex; i++)
                {
                    IRow row = sheet.GetRow(i) ?? sheet.CreateRow(i);
                    for (int j = firstColumnIndex; j <= lastColumnIndex; j++)
                    {
                        ICell cell = row.GetCell(j) ?? row.CreateCell(j);

                        byte[] bytesARGB = new byte[] { (byte)a, (byte)r, (byte)g, (byte)b };
                        XSSFColor xSSFColor = new XSSFColor(bytesARGB);

                        ICellStyle oldCellStyle = cell.CellStyle;
                        ICellStyle newCellStyle = _workbook.CreateCellStyle();
                        newCellStyle.CloneStyleFrom(oldCellStyle);

                        foreach (var itemBorder in borderIndex)
                        {
                            var npoiLinesIndex = NpoiHelper.ConvertFromLineStyleNpoi(lineIndex, out bool isEmpty);
                            if (!isEmpty)
                            {
                                switch (itemBorder)
                                {
                                    case BorderIndex.Bottom:
                                        newCellStyle.BorderBottom = npoiLinesIndex;
                                        ((XSSFCellStyle)newCellStyle).SetBottomBorderColor(xSSFColor);
                                        break;
                                    case BorderIndex.Top:
                                        newCellStyle.BorderTop = npoiLinesIndex;
                                        ((XSSFCellStyle)newCellStyle).SetTopBorderColor(xSSFColor);
                                        break;
                                    case BorderIndex.Left:
                                        newCellStyle.BorderLeft = npoiLinesIndex;
                                        ((XSSFCellStyle)newCellStyle).SetLeftBorderColor(xSSFColor);
                                        break;
                                    case BorderIndex.Right:
                                        newCellStyle.BorderRight = npoiLinesIndex;
                                        ((XSSFCellStyle)newCellStyle).SetRightBorderColor(xSSFColor);
                                        break;
                                    case BorderIndex.All:
                                        newCellStyle.BorderBottom = npoiLinesIndex;
                                        newCellStyle.BorderTop = npoiLinesIndex;
                                        newCellStyle.BorderLeft = npoiLinesIndex;
                                        newCellStyle.BorderRight = npoiLinesIndex;
                                        ((XSSFCellStyle)newCellStyle).SetBottomBorderColor(xSSFColor);
                                        ((XSSFCellStyle)newCellStyle).SetTopBorderColor(xSSFColor);
                                        ((XSSFCellStyle)newCellStyle).SetLeftBorderColor(xSSFColor);
                                        ((XSSFCellStyle)newCellStyle).SetRightBorderColor(xSSFColor);
                                        break;
                                    case BorderIndex.None:
                                        break;
                                    default:
                                        throw new NotImplementedException();

                                }
                            }
                            cell.CellStyle = newCellStyle;
                        }
                    }

                }
            }
            else
            {
                for (int i = firstRowIndex; i <= lastRowIndex; i++)
                {
                    IRow row = sheet.GetRow(i) ?? sheet.CreateRow(i);
                    for (int j = firstColumnIndex; j <= lastColumnIndex; j++)
                    {
                        ICell cell = row.GetCell(j) ?? row.CreateCell(j);

                        byte[] bytesARGB = new byte[] { (byte)a, (byte)r, (byte)g, (byte)b };
                        XSSFColor xSSFColor = new XSSFColor(bytesARGB);

                        ICellStyle oldCellStyle = cell.CellStyle;
                        ICellStyle newCellStyle = _workbook.CreateCellStyle();
                        newCellStyle.CloneStyleFrom(oldCellStyle);

                        foreach (var itemBorder in borderIndex)
                        {
                            var npoiLinesIndex = NpoiHelper.ConvertFromLineStyleNpoi(lineIndex, out bool isEmpty);
                            if (!isEmpty)
                            {
                                if (i == firstRowIndex && (itemBorder == BorderIndex.Top || itemBorder == BorderIndex.All))
                                {
                                    newCellStyle.BorderTop = npoiLinesIndex;
                                    ((XSSFCellStyle)newCellStyle).SetTopBorderColor(xSSFColor);
                                }
                                if (i == lastRowIndex && (itemBorder == BorderIndex.Bottom || itemBorder == BorderIndex.All))
                                {
                                    newCellStyle.BorderBottom = npoiLinesIndex;
                                    ((XSSFCellStyle)newCellStyle).SetBottomBorderColor(xSSFColor);
                                }
                                if (j == firstColumnIndex && (itemBorder == BorderIndex.Left || itemBorder == BorderIndex.All))
                                {
                                    newCellStyle.BorderLeft = npoiLinesIndex;
                                    ((XSSFCellStyle)newCellStyle).SetLeftBorderColor(xSSFColor);
                                }
                                if (j == lastColumnIndex && (itemBorder == BorderIndex.Right || itemBorder == BorderIndex.All))
                                {
                                    newCellStyle.BorderRight = npoiLinesIndex;
                                    ((XSSFCellStyle)newCellStyle).SetRightBorderColor(xSSFColor);
                                }
                            }
                            cell.CellStyle = newCellStyle;
                        }
                    }

                }
            }
        }

        public void SetFont(FontSettings fontSettings, int sheetIndex, int rowIndex, int columnIndex)
        {
            ISheet sheet = _workbook.GetSheetAt(sheetIndex);
            IRow row = sheet.GetRow(rowIndex) ?? sheet.CreateRow(rowIndex);
            ICell cell = row.GetCell(columnIndex) ?? row.CreateCell(columnIndex);

            ICellStyle oldCellStyle = cell.CellStyle;
            ICellStyle newCellStyle = _workbook.CreateCellStyle();
            newCellStyle.CloneStyleFrom(oldCellStyle);
            IFont font = _workbook.CreateFont();

            if (fontSettings.Bold.HasValue)
            {
                font.IsBold = fontSettings.Bold.Value;
            }
            if (fontSettings.Italics.HasValue)
            {
                font.IsItalic = fontSettings.Italics.Value;
            }
            if (fontSettings.DoubleUnderline.HasValue && fontSettings.DoubleUnderline.Value)
            {
                font.Underline = FontUnderlineType.Double;
            }
            else if (fontSettings.Underline.HasValue && fontSettings.Underline.Value)
            {
                font.Underline = FontUnderlineType.Single;
            }
            if (fontSettings.TextCrossed.HasValue)
            {
                font.IsStrikeout = fontSettings.TextCrossed.Value;
            }
            if (fontSettings.TextWrapping.HasValue)
            {
                newCellStyle.WrapText = fontSettings.TextWrapping.Value;
            }
            if (!string.IsNullOrEmpty(fontSettings.FontName))
            {
                font.FontName = fontSettings.FontName;
            }
            if (fontSettings.FontSize.HasValue)
            {
                font.FontHeightInPoints = fontSettings.FontSize.Value;
            }
            if (fontSettings.A.HasValue && fontSettings.R.HasValue && fontSettings.G.HasValue && fontSettings.B.HasValue)
            {
                byte[] TextColorARGB = new byte[] {(byte)fontSettings.A.Value, (byte)fontSettings.R.Value, (byte)fontSettings.G.Value, (byte)fontSettings.B.Value};
                XSSFColor xSSFColor = new XSSFColor(TextColorARGB);
                ((XSSFFont)font).SetColor(xSSFColor);
            }

            newCellStyle.SetFont(font);
            cell.CellStyle = newCellStyle;
        }

        public void SetFont(FontSettings fontSettings, string sheetName, int rowIndex, int columnIndex)
        {
            ISheet sheet = _workbook.GetSheet(sheetName);
            IRow row = sheet.GetRow(rowIndex) ?? sheet.CreateRow(rowIndex);
            ICell cell = row.GetCell(columnIndex) ?? row.CreateCell(columnIndex);

            ICellStyle oldCellStyle = cell.CellStyle;
            ICellStyle newCellStyle = _workbook.CreateCellStyle();
            newCellStyle.CloneStyleFrom(oldCellStyle);
            IFont font = _workbook.CreateFont();

            if (fontSettings.Bold.HasValue)
            {
                font.IsBold = fontSettings.Bold.Value;
            }
            if (fontSettings.Italics.HasValue)
            {
                font.IsItalic = fontSettings.Italics.Value;
            }
            if (fontSettings.DoubleUnderline.HasValue && fontSettings.DoubleUnderline.Value)
            {
                font.Underline = FontUnderlineType.Double;
            }
            else if (fontSettings.Underline.HasValue && fontSettings.Underline.Value)
            {
                font.Underline = FontUnderlineType.Single;
            }
            if (fontSettings.TextCrossed.HasValue)
            {
                font.IsStrikeout = fontSettings.TextCrossed.Value;
            }
            if (fontSettings.TextWrapping.HasValue)
            {
                newCellStyle.WrapText = fontSettings.TextWrapping.Value;
            }
            if (!string.IsNullOrEmpty(fontSettings.FontName))
            {
                font.FontName = fontSettings.FontName;
            }
            if (fontSettings.FontSize.HasValue)
            {
                font.FontHeightInPoints = fontSettings.FontSize.Value;
            }
            if (fontSettings.A.HasValue && fontSettings.R.HasValue && fontSettings.G.HasValue && fontSettings.B.HasValue)
            {
                byte[] TextColorARGB = new byte[] { (byte)fontSettings.A.Value, (byte)fontSettings.R.Value, (byte)fontSettings.G.Value, (byte)fontSettings.B.Value };
                XSSFColor xSSFColor = new XSSFColor(TextColorARGB);
                ((XSSFFont)font).SetColor(xSSFColor);
            }

            newCellStyle.SetFont(font);
            cell.CellStyle = newCellStyle;
        }

        public void SetFont(FontSettings fontSettings, int sheetIndex, int firstRowIndex, int firstColumnIndex, int lastRowIndex, int lastColumnIndex)
        {
            ISheet sheet = _workbook.GetSheetAt(sheetIndex);
            for (int i = firstRowIndex; i <= lastRowIndex; i++)
            {
                for (int j = firstColumnIndex; j <= lastColumnIndex; j++)
                {
                    IRow row = sheet.GetRow(i) ?? sheet.CreateRow(i);
                    ICell cell = row.GetCell(j) ?? row.CreateCell(j);

                    ICellStyle oldCellStyle = cell.CellStyle;
                    ICellStyle newCellStyle = _workbook.CreateCellStyle();
                    newCellStyle.CloneStyleFrom(oldCellStyle);
                    IFont font = _workbook.CreateFont();

                    if (fontSettings.Bold.HasValue)
                    {
                        font.IsBold = fontSettings.Bold.Value;
                    }
                    if (fontSettings.Italics.HasValue)
                    {
                        font.IsItalic = fontSettings.Italics.Value;
                    }
                    if (fontSettings.DoubleUnderline.HasValue && fontSettings.DoubleUnderline.Value)
                    {
                        font.Underline = FontUnderlineType.Double;
                    }
                    else if (fontSettings.Underline.HasValue && fontSettings.Underline.Value)
                    {
                        font.Underline = FontUnderlineType.Single;
                    }
                    if (fontSettings.TextCrossed.HasValue)
                    {
                        font.IsStrikeout = fontSettings.TextCrossed.Value;
                    }
                    if (fontSettings.TextWrapping.HasValue)
                    {
                        newCellStyle.WrapText = fontSettings.TextWrapping.Value;
                    }
                    if (!string.IsNullOrEmpty(fontSettings.FontName))
                    {
                        font.FontName = fontSettings.FontName;
                    }
                    if (fontSettings.FontSize.HasValue)
                    {
                        font.FontHeightInPoints = fontSettings.FontSize.Value;
                    }
                    if (fontSettings.A.HasValue && fontSettings.R.HasValue && fontSettings.G.HasValue && fontSettings.B.HasValue)
                    {
                        byte[] TextColorARGB = new byte[] { (byte)fontSettings.A.Value, (byte)fontSettings.R.Value, (byte)fontSettings.G.Value, (byte)fontSettings.B.Value };
                        XSSFColor xSSFColor = new XSSFColor(TextColorARGB);
                        ((XSSFFont)font).SetColor(xSSFColor);
                    }

                    newCellStyle.SetFont(font);
                    cell.CellStyle = newCellStyle;
                }
            }
        }

        public void SetFont(FontSettings fontSettings, string sheetName, int firstRowIndex, int firstColumnIndex, int lastRowIndex, int lastColumnIndex)
        {
            ISheet sheet = _workbook.GetSheet(sheetName);
            for (int i = firstRowIndex; i <= lastRowIndex; i++)
            {
                for (int j = firstColumnIndex; j <= lastColumnIndex; j++)
                {
                    IRow row = sheet.GetRow(i) ?? sheet.CreateRow(i);
                    ICell cell = row.GetCell(j) ?? row.CreateCell(j);

                    ICellStyle oldCellStyle = cell.CellStyle;
                    ICellStyle newCellStyle = _workbook.CreateCellStyle();
                    newCellStyle.CloneStyleFrom(oldCellStyle);
                    IFont font = _workbook.CreateFont();

                    if (fontSettings.Bold.HasValue)
                    {
                        font.IsBold = fontSettings.Bold.Value;
                    }
                    if (fontSettings.Italics.HasValue)
                    {
                        font.IsItalic = fontSettings.Italics.Value;
                    }
                    if (fontSettings.DoubleUnderline.HasValue && fontSettings.DoubleUnderline.Value)
                    {
                        font.Underline = FontUnderlineType.Double;
                    }
                    else if (fontSettings.Underline.HasValue && fontSettings.Underline.Value)
                    {
                        font.Underline = FontUnderlineType.Single;
                    }
                    if (fontSettings.TextCrossed.HasValue)
                    {
                        font.IsStrikeout = fontSettings.TextCrossed.Value;
                    }
                    if (fontSettings.TextWrapping.HasValue)
                    {
                        newCellStyle.WrapText = fontSettings.TextWrapping.Value;
                    }
                    if (!string.IsNullOrEmpty(fontSettings.FontName))
                    {
                        font.FontName = fontSettings.FontName;
                    }
                    if (fontSettings.FontSize.HasValue)
                    {
                        font.FontHeightInPoints = fontSettings.FontSize.Value;
                    }
                    if (fontSettings.A.HasValue && fontSettings.R.HasValue && fontSettings.G.HasValue && fontSettings.B.HasValue)
                    {
                        byte[] TextColorARGB = new byte[] { (byte)fontSettings.A.Value, (byte)fontSettings.R.Value, (byte)fontSettings.G.Value, (byte)fontSettings.B.Value };
                        XSSFColor xSSFColor = new XSSFColor(TextColorARGB);
                        ((XSSFFont)font).SetColor(xSSFColor);
                    }

                    newCellStyle.SetFont(font);
                    cell.CellStyle = newCellStyle;
                }
            }
        }

        public void SetHorizontalAlignment(HorizontalAlignmentIndex horizontalAlignmentIndex, int sheetIndex,  int rowIndex, int columnIndex)
        {
            ISheet sheet = _workbook.GetSheetAt(sheetIndex);
            IRow row = sheet.GetRow(rowIndex) ?? sheet.CreateRow(rowIndex);
            ICell cell = row.GetCell(columnIndex) ?? row.CreateCell(columnIndex);

            ICellStyle oldCellStyle = cell.CellStyle;
            ICellStyle newCellStyle = _workbook.CreateCellStyle();
            newCellStyle.CloneStyleFrom(oldCellStyle);

            var hAlignment = NpoiHelper.ConverFromHorizontalAlignmentNpoi(horizontalAlignmentIndex, out bool isEmpty);
            if (!isEmpty)
            {
                newCellStyle.Alignment = hAlignment;
                cell.CellStyle = newCellStyle;
            }
        }

        public void SetHorizontalAlignment(HorizontalAlignmentIndex horizontalAlignmentIndex, string sheetName, int rowIndex, int columnIndex)
        {
            ISheet sheet = _workbook.GetSheet(sheetName);
            IRow row = sheet.GetRow(rowIndex) ?? sheet.CreateRow(rowIndex);
            ICell cell = row.GetCell(columnIndex) ?? row.CreateCell(columnIndex);

            ICellStyle oldCellStyle = cell.CellStyle;
            ICellStyle newCellStyle = _workbook.CreateCellStyle();
            newCellStyle.CloneStyleFrom(oldCellStyle);

            var hAlignment = NpoiHelper.ConverFromHorizontalAlignmentNpoi(horizontalAlignmentIndex, out bool isEmpty);
            if (!isEmpty)
            {
                newCellStyle.Alignment = hAlignment;
                cell.CellStyle = newCellStyle;
            }
        }

        public void SetHorizontalAlignment(HorizontalAlignmentIndex horizontalAlignmentIndex, int sheetIndex, int firstRowIndex, int firstColumnIndex, int lastRowIndex, int lastColumnIndex)
        {
            ISheet sheet = _workbook.GetSheetAt(sheetIndex);
            for (int r = firstRowIndex; r <= lastRowIndex; r++)
            {
                IRow row = sheet.GetRow(r) ?? sheet.CreateRow(r);
                for (int c = firstColumnIndex; c <= lastColumnIndex; c++)
                {
                    ICell cell = row.GetCell(c) ?? row.CreateCell(c);

                    ICellStyle oldCellStyle = cell.CellStyle;
                    ICellStyle newCellStyle = _workbook.CreateCellStyle();
                    newCellStyle.CloneStyleFrom(oldCellStyle);

                    var hAlignment = NpoiHelper.ConverFromHorizontalAlignmentNpoi(horizontalAlignmentIndex, out bool isEmpty);
                    if (!isEmpty)
                    {
                        newCellStyle.Alignment = hAlignment;
                        cell.CellStyle = newCellStyle;
                    }
                }
            }
        }

        public void SetHorizontalAlignment(HorizontalAlignmentIndex horizontalAlignmentIndex, string sheetName, int firstRowIndex, int firstColumnIndex, int lastRowIndex, int lastColumnIndex)
        {
            ISheet sheet = _workbook.GetSheet(sheetName);
            for (int r = firstRowIndex; r <= lastRowIndex; r++)
            {
                IRow row = sheet.GetRow(r) ?? sheet.CreateRow(r);
                for (int c = firstColumnIndex; c <= lastColumnIndex; c++)
                {
                    ICell cell = row.GetCell(c) ?? row.CreateCell(c);

                    ICellStyle oldCellStyle = cell.CellStyle;
                    ICellStyle newCellStyle = _workbook.CreateCellStyle();
                    newCellStyle.CloneStyleFrom(oldCellStyle);

                    var hAlignment = NpoiHelper.ConverFromHorizontalAlignmentNpoi(horizontalAlignmentIndex, out bool isEmpty);
                    if (!isEmpty)
                    {
                        newCellStyle.Alignment = hAlignment;
                        cell.CellStyle = newCellStyle;
                    }
                }
            }
        }

        public void SetVerticalAlignment(VerticalAlignmentIndex verticalAlignmentIndex, int sheetIndex, int rowIndex, int columnIndex)
        {
            ISheet sheet = _workbook.GetSheetAt(sheetIndex);
            IRow row = sheet.GetRow(rowIndex) ?? sheet.CreateRow(rowIndex);
            ICell cell = row.GetCell(columnIndex) ?? row.CreateCell(columnIndex);

            ICellStyle oldCellStyle = cell.CellStyle;
            ICellStyle newCellStyle = _workbook.CreateCellStyle();
            newCellStyle.CloneStyleFrom(oldCellStyle);

            var VAligment = NpoiHelper.ConverFromVerticalAligmentNpoi(verticalAlignmentIndex, out bool isEmpty);
            if (!isEmpty)
            {
                newCellStyle.VerticalAlignment = VAligment;
                cell.CellStyle = newCellStyle;
            }
        }

        public void SetVerticalAlignment(VerticalAlignmentIndex verticalAlignmentIndex, string sheetName, int rowIndex, int columnIndex)
        {
            ISheet sheet = _workbook.GetSheet(sheetName);
            IRow row = sheet.GetRow(rowIndex) ?? sheet.CreateRow(rowIndex);
            ICell cell = row.GetCell(columnIndex) ?? row.CreateCell(columnIndex);

            ICellStyle oldCellStyle = cell.CellStyle;
            ICellStyle newCellStyle = _workbook.CreateCellStyle();
            newCellStyle.CloneStyleFrom(oldCellStyle);

            var VAligment = NpoiHelper.ConverFromVerticalAligmentNpoi(verticalAlignmentIndex, out bool isEmpty);
            if (!isEmpty)
            {
                newCellStyle.VerticalAlignment = VAligment;
                cell.CellStyle = newCellStyle;
            }
        }

        public void SetVerticalAlignment(VerticalAlignmentIndex verticalAlignmentIndex, int sheetIndex, int firstRowIndex, int firstColumnIndex, int lastRowIndex, int lastColumnIndex)
        {
            ISheet sheet = _workbook.GetSheetAt(sheetIndex);

            for (int r = firstRowIndex; r <= lastRowIndex; r++)
            {
                IRow row = sheet.GetRow(r) ?? sheet.CreateRow(r);
                for (int c = firstColumnIndex; c <= lastColumnIndex; c++)
                {
                    ICell cell = row.GetCell(c) ?? row.CreateCell(c);

                    ICellStyle oldCellStyle = cell.CellStyle;
                    ICellStyle newCellStyle = _workbook.CreateCellStyle();
                    newCellStyle.CloneStyleFrom(oldCellStyle);

                    var VAligment = NpoiHelper.ConverFromVerticalAligmentNpoi(verticalAlignmentIndex, out bool isEmpty);
                    if (!isEmpty)
                    {
                        newCellStyle.VerticalAlignment = VAligment;
                        cell.CellStyle = newCellStyle;
                    }
                }
            }
        }

        public void SetVerticalAlignment(VerticalAlignmentIndex verticalAlignmentIndex, string sheetName, int firstRowIndex, int firstColumnIndex, int lastRowIndex, int lastColumnIndex)
        {
            ISheet sheet = _workbook.GetSheet(sheetName);

            for (int r = firstRowIndex; r <= lastRowIndex; r++)
            {
                IRow row = sheet.GetRow(r) ?? sheet.CreateRow(r);
                for (int c = firstColumnIndex; c <= lastColumnIndex; c++)
                {
                    ICell cell = row.GetCell(c) ?? row.CreateCell(c);

                    ICellStyle oldCellStyle = cell.CellStyle;
                    ICellStyle newCellStyle = _workbook.CreateCellStyle();
                    newCellStyle.CloneStyleFrom(oldCellStyle);

                    var VAligment = NpoiHelper.ConverFromVerticalAligmentNpoi(verticalAlignmentIndex, out bool isEmpty);
                    if (!isEmpty)
                    {
                        newCellStyle.VerticalAlignment = VAligment;
                        cell.CellStyle = newCellStyle;
                    }
                }
            }
        }

        public void ValueOrientation(int sheetIndex, int rowIndex, int columnIndex, short orientation)
        {
            ISheet sheet = _workbook.GetSheetAt(sheetIndex);

            IRow row = sheet.GetRow(rowIndex) ?? sheet.CreateRow(rowIndex);
            ICell cell = row.GetCell(columnIndex) ?? row.GetCell(columnIndex);

            ICellStyle oldStyle = cell.CellStyle;
            ICellStyle newStyle = _workbook.CreateCellStyle();
            newStyle.CloneStyleFrom(oldStyle);
            newStyle.Rotation = orientation;
            cell.CellStyle = newStyle;
        }

        public void ValueOrientation(string sheetName, int rowIndex, int columnIndex, short orientation)
        {
            ISheet sheet = _workbook.GetSheet(sheetName);

            IRow row = sheet.GetRow(rowIndex) ?? sheet.CreateRow(rowIndex);
            ICell cell = row.GetCell(columnIndex) ?? row.GetCell(columnIndex);

            ICellStyle oldStyle = cell.CellStyle;
            ICellStyle newStyle = _workbook.CreateCellStyle();
            newStyle.CloneStyleFrom(oldStyle);
            newStyle.Rotation = orientation;
            cell.CellStyle = newStyle;
        }

        public void ValueOrientation(int sheetIndex, int firstRowIndex, int firstColumnIndex, int lastRowIndex, int lastColumnIndex, short orientation)
        {
            ISheet sheet = _workbook.GetSheetAt(sheetIndex);

            for (int r = firstRowIndex; r <= lastRowIndex; r++)
            {
                IRow row = sheet.GetRow(r) ?? sheet.CreateRow(r);
                for (int c = firstColumnIndex; c <= lastColumnIndex; c++)
                {
                    ICell cell = row.GetCell(c) ?? row.GetCell(c);

                    ICellStyle oldStyle = cell.CellStyle;
                    ICellStyle newStyle = _workbook.CreateCellStyle();
                    newStyle.CloneStyleFrom(oldStyle);
                    newStyle.Rotation = orientation;
                    cell.CellStyle = newStyle;
                }
            }
        }

        public void ValueOrientation(string sheetName, int firstRowIndex, int firstColumnIndex, int lastRowIndex, int lastColumnIndex, short orientation)
        {
            ISheet sheet = _workbook.GetSheet(sheetName);

            for (int r = firstRowIndex; r <= lastRowIndex; r++)
            {
                IRow row = sheet.GetRow(r) ?? sheet.CreateRow(r);
                for (int c = firstColumnIndex; c <= lastColumnIndex; c++)
                {
                    ICell cell = row.GetCell(c) ?? row.GetCell(c);

                    ICellStyle oldStyle = cell.CellStyle;
                    ICellStyle newStyle = _workbook.CreateCellStyle();
                    newStyle.CloneStyleFrom(oldStyle);
                    newStyle.Rotation = orientation;
                    cell.CellStyle = newStyle;
                }
            }
        }

        public void WidthColumn(int sheetIndex, int columnIndex, double width)
        {
            ISheet sheet = _workbook.GetSheetAt(sheetIndex);
            sheet.SetColumnWidth(columnIndex, width * 256);
        }

        public void WidthColumn(int sheetIndex, int columnIndex, int columnCount, double width)
        {
            ISheet sheet = _workbook.GetSheetAt(sheetIndex);
            for (int i = 0; i < columnCount; i++)
            {
                sheet.SetColumnWidth(columnIndex + i, width * 256);
            }
        }
        public void WidthColumn(int sheetIndex, int[] columnIndices, double width)
        {
            ISheet sheet = _workbook.GetSheetAt(sheetIndex);
            
            foreach (var index in columnIndices)
            {
                sheet.SetColumnWidth(index, width * 256);
            }
        }

        public void WidthColumn(string sheetName, int columnIndex, int columnCount, double width)
        {
            ISheet sheet = _workbook.GetSheet(sheetName);
            for (int i = 0; i < columnCount; i++)
            {
                sheet.SetColumnWidth(columnIndex + i, width * 256);
            }
        }

        public void WidthColumn(string sheetName, int[] columnIndices, double width)
        {
            ISheet sheet = _workbook.GetSheet(sheetName);

            foreach (var index in columnIndices)
            {
                sheet.SetColumnWidth(index, width * 256);
            }
        }

        public void WidthColumn(string sheetName, int columnIndex, double width)
        {
            ISheet sheet = _workbook.GetSheet(sheetName);
            sheet.SetColumnWidth(columnIndex, width * 256);
        }

        public void SetProtectSheet(int sheetIndex, string password)
        {
            ISheet sheet = _workbook.GetSheetAt(sheetIndex);
            sheet.ProtectSheet(password);
        }

        public void SetProtectSheet(string sheetName, string password)
        {
            ISheet sheet = _workbook.GetSheet(sheetName);
            sheet.ProtectSheet(password);
        }

        public void DropDownList(int sheetIndex, int dataSheetIndex,  int firstRowIndex, int firstColumnIndex, int lastRowIndex, int lastColumnIndex, string rangeDataToList)
        {
            ISheet sheet = _workbook.GetSheetAt(sheetIndex);
            ISheet dataSheet = _workbook.GetSheetAt(dataSheetIndex);
            IName nameRange = _workbook.CreateName();
            nameRange.RefersToFormula = $"{dataSheet.SheetName}!{rangeDataToList}";
            nameRange.NameName = "DropDownValues";

            IDataValidationHelper dataValidationHelper = sheet.GetDataValidationHelper();
            CellRangeAddressList cellRangeAddressList = new CellRangeAddressList(firstRowIndex, lastRowIndex, firstColumnIndex, lastColumnIndex);
            IDataValidationConstraint dataValidationConstraint = dataValidationHelper.CreateFormulaListConstraint("DropDownValues");
            IDataValidation dataValidation = dataValidationHelper.CreateValidation(dataValidationConstraint, cellRangeAddressList);
            dataValidation.SuppressDropDownArrow = true;
            dataValidation.ShowErrorBox = true;
            sheet.AddValidationData(dataValidation);
        }

        public void DropDownList(string sheetName, int dataSheetIndex, int firstRowIndex, int firstColumnIndex, int lastRowIndex, int lastColumnIndex, string rangeDataToList)
        {
            ISheet sheet = _workbook.GetSheet(sheetName);
            ISheet dataSheet = _workbook.GetSheetAt(dataSheetIndex);
            IName nameRange = _workbook.CreateName();
            nameRange.RefersToFormula = $"{dataSheet.SheetName}!{rangeDataToList}";
            nameRange.NameName = "DropDownValues";

            IDataValidationHelper dataValidationHelper = sheet.GetDataValidationHelper();
            CellRangeAddressList cellRangeAddressList = new CellRangeAddressList(firstRowIndex, lastRowIndex, firstColumnIndex, lastColumnIndex);
            IDataValidationConstraint dataValidationConstraint = dataValidationHelper.CreateFormulaListConstraint("DropDownValues");
            IDataValidation dataValidation = dataValidationHelper.CreateValidation(dataValidationConstraint, cellRangeAddressList);
            dataValidation.SuppressDropDownArrow = true;
            dataValidation.ShowErrorBox = true;
            sheet.AddValidationData(dataValidation);
        }

        public void DropDownList(int sheetIndex, string dataSheetName, int firstRowIndex, int firstColumnIndex, int lastRowIndex, int lastColumnIndex, string rangeDataToList)
        {
            ISheet sheet = _workbook.GetSheetAt(sheetIndex);
            ISheet dataSheet = _workbook.GetSheet(dataSheetName);
            IName nameRange = _workbook.CreateName();
            nameRange.RefersToFormula = $"{dataSheet.SheetName}!{rangeDataToList}";
            nameRange.NameName = "DropDownValues";

            IDataValidationHelper dataValidationHelper = sheet.GetDataValidationHelper();
            CellRangeAddressList cellRangeAddressList = new CellRangeAddressList(firstRowIndex, lastRowIndex, firstColumnIndex, lastColumnIndex);
            IDataValidationConstraint dataValidationConstraint = dataValidationHelper.CreateFormulaListConstraint("DropDownValues");
            IDataValidation dataValidation = dataValidationHelper.CreateValidation(dataValidationConstraint, cellRangeAddressList);
            dataValidation.SuppressDropDownArrow = true;
            dataValidation.ShowErrorBox = true;
            sheet.AddValidationData(dataValidation);
        }

        public void DropDownList(string sheetName, string dataSheetName, int firstRowIndex, int firstColumnIndex, int lastRowIndex, int lastColumnIndex, string rangeDataToList)
        {
            ISheet sheet = _workbook.GetSheet(sheetName);
            ISheet dataSheet = _workbook.GetSheet(dataSheetName);
            IName nameRange = _workbook.CreateName();
            nameRange.RefersToFormula = $"{dataSheet.SheetName}!{rangeDataToList}";
            nameRange.NameName = "DropDownValues";

            IDataValidationHelper dataValidationHelper = sheet.GetDataValidationHelper();
            CellRangeAddressList cellRangeAddressList = new CellRangeAddressList(firstRowIndex, lastRowIndex, firstColumnIndex, lastColumnIndex);
            IDataValidationConstraint dataValidationConstraint = dataValidationHelper.CreateFormulaListConstraint("DropDownValues");
            IDataValidation dataValidation = dataValidationHelper.CreateValidation(dataValidationConstraint, cellRangeAddressList);
            dataValidation.SuppressDropDownArrow = true;
            dataValidation.ShowErrorBox = true;
            sheet.AddValidationData(dataValidation);
        }
    }
}
