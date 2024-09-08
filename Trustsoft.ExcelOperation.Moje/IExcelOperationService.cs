using NPOI.OpenXmlFormats.Spreadsheet;
using NPOI.SS.UserModel;
using Soneta.Types;
using Syncfusion.XlsIO;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Trustsoft.ExcelOperation.Moje
{
    internal interface IExcelOperationService
    {
        /// <summary>
        /// Creates a workbook according to the conditions passed to the class.
        /// </summary>
        /// <returns>Returns the object.</returns>
        public object CreateWorkbook();

        /// <summary>
        /// Creates a workbook according to the conditions passed to the class and with the sheet name passed in the parameter.
        /// </summary>
        /// <param name="sheetName">Sheet name string.</param>
        /// <returns>Returns the object.</returns>
        public object CreateWorkbook(string sheetName);

        /// <summary>
        /// Changes the name of a specified worksheet in the workbook.
        /// </summary>
        /// <param name="sheetName">The current name of the worksheet to be renamed.</param>
        /// <param name="newName">The new name to assign to the worksheet</param>
        public void ChangeNameWorksheet(string sheetName, string newName);

        /// <summary>
        /// Changes the name of a specified worksheet in the workbook.
        /// </summary>
        /// <param name="indexsheet">The current index of the worksheet to be renamed.</param>
        /// <param name="newName">The new name to assign to the worksheet</param>
        public void ChangeNameWorksheet(int indexsheet, string newName);

        /// <summary>
        /// Creates a new sheet with the specified name in the workbook.
        /// </summary>
        /// <param name="sheetName">New sheet name.</param>
        /// <returns> Returns the sheet index.</returns>
        public int AddWorksheet(string sheetName);

        /// <summary>
        /// Creates a list of sheet containing the index and sheet name.
        /// </summary>
        /// <returns>Returns a list of sheets.</returns>
        public List<Sheet> GetNameSheet();

        /// <summary>
        /// Deletes a sheet with the specified name in the workbook.
        /// </summary>
        /// <param name="sheetName">The name of the sheet to be deleted.</param>
        public void DeleteWorksheet(string sheetName);

        /// <summary>
        /// Deletes a sheet with the specified name in the workbook.
        /// </summary>
        /// <param name="sheetIndex">The index of the sheet to be deleted.</param>
        public void DeleteWorksheet(int sheetIndex);

        /// <summary>
        /// Creates a row in the worksheet by index (line number). A sheet is identified by its name.
        /// </summary>
        /// <param name="rowIndex">Line number where to insert a new row.</param>
        /// <param name="sheetName">The name of the sheet in which to insert a new row.</param>
        public void AddRow(int rowIndex, string sheetName);

        /// <summary>
        /// Creates a row in the worksheet by index (line number). A sheet is identified by its name.
        /// </summary>
        /// <param name="sheetName">The name of the sheet in which to insert a new row.</param>
        /// <param name="rowIndex">Line number where to insert a new row.</param>
        public void AddRow(string sheetName, int rowIndex);

        /// <summary>
        /// Creates a rows in the worksheet by index (line number) and quantity. A sheet is identified by its name.
        /// </summary>
        /// <param name="rowIndex">The row number at which to start inserting a new rows.</param>
        /// <param name="sheetName">The name of the sheet in which to insert a new row.</param>
        /// <param name="rowCount">Number of rows to insert.</param>
        public void AddRow(int rowIndex, string sheetName, int rowCount);

        /// <summary>
        /// Creates a rows in the worksheet by index (line number) and quantity. A sheet is identified by its name.
        /// </summary>
        /// <param name="sheetName">The name of the sheet in which to insert a new rows.</param>
        /// <param name="rowIndex">The row number at which to start inserting a new row.</param>
        /// <param name="rowCount">Number of rows to insert.</param>
        public void AddRow(string sheetName, int rowIndex, int rowCount);

        /// <summary>
        /// Creates a rows in the worksheet by index (line number) and quantity. A sheet is identified by its index.
        /// </summary>
        /// <param name="sheetIndex">The index of the sheet in which to insert a new rows.</param>
        /// <param name="rowIndex">The row number at which to start inserting a new row.</param>
        /// <param name="rowCount">Number of rows to insert.</param>
        public void AddRow(int sheetIndex, int rowIndex, int rowCount);

        /// <summary>
        /// Creates a row in the worksheet by index (line number). A sheet is identified by its index.
        /// </summary>
        /// <param name="sheetIndex">Sheet index number in which to insert the row.</param>
        /// <param name="rowIndex">Line number where to insert a new row.</param>
        public void AddRow(int sheetIndex, int rowIndex);

        /// <summary>
        /// Creates a column in the worksheet by index (column number). A sheet is identified by its name.
        /// </summary>
        /// <param name="coulmnIndex">Column number where to insert a new column.</param>
        /// <param name="sheetName">The name of the sheet in which to insert a new column.</param>
        public void AddColumn(int columnIndex, string sheetName);

        /// <summary>
        /// Creates a column in the worksheet by index (column number). A sheet is identified by its name.
        /// </summary>
        /// <param name="sheetName">The name of the sheet in which to insert a new column.</param>
        /// <param name="columnIndex">Column number where to insert a new column.</param>
        public void AddColumn(string sheetName, int columnIndex);

        /// <summary>
        /// Creates a columns in the worksheet by index (column number) and quantity. A sheet is identified by its index.
        /// </summary>
        /// <param name="sheetIndex">The index of the sheet in which to insert a new columns.</param>
        /// <param name="columnIndex">The column number at which to start inserting a new column.</param>
        /// <param name="columnCount">Number of columns to insert.</param>
        public void AddColumn(int sheetIndex, int columnIndex, int columnCount);

        /// <summary>
        /// Creates a columns in the worksheet by index (column number) and quantity. A sheet is identified by its name.
        /// </summary>
        /// <param name="columnIndex">The column number at which to start inserting a new column.</param>
        /// <param name="sheetName">The name of the sheet in which to insert a new columns.</param>
        /// <param name="columnCount">Number of columns to insert.</param>
        public void AddColumn(int columnIndex, string sheetName, int columnCount);

        /// <summary>
        /// Creates a columns in the worksheet by index (column number) and quantity. A sheet is identified by its name.
        /// </summary>
        /// <param name="sheetName">The name of the sheet in which to insert a new columns.</param>
        /// <param name="columnIndex">The column number at which to start inserting a new column.</param>
        /// <param name="columnCount">Number of columns to insert.</param>
        public void AddColumn(string sheetName, int columnIndex, int columnCount);

        /// <summary>
        /// Creates a column in the worksheet by index (column number). A sheet is identified by its index.
        /// </summary>
        /// <param name="sheetIndex">The index of the sheet in which to insert a new column.</param>
        /// <param name="columnIndex">Column number where to insert a new column.</param>
        public void AddColumn(int sheetIndex, int columnIndex);

        /// <summary>
        /// Inserts a text value into a cell in a worksheet. Sheet is identified by its index. The cell is identyfied by its row number and column number.
        /// </summary>
        /// <param name="sheetIndex">The index of the sheet in which to insert text value.</param>
        /// <param name="rowIndex">The row number where the value is to be set.</param>
        /// <param name="columnIndex">The column number where the value is to be set.</param>
        /// <param name="text">A text value.</param>
        public void AddCellValueText(int sheetIndex, int rowIndex, int columnIndex, string text);

        /// <summary>
        /// Inserts a text value into a cell in a worksheet. Sheet is identified by its name. The cell is identyfied by its row number and column number.
        /// </summary>
        /// <param name="sheetName">The name of the sheet in which to insert text value.</param>
        /// <param name="rowIndex">The row number where the value is to be set.</param>
        /// <param name="columnIndex">The column number where the value is to be set.</param>
        /// <param name="text">A text value.</param>
        public void AddCellValueText(string sheetName, int rowIndex, int columnIndex, string text);

        /// <summary>
        /// Inserts a formula into a cell in a worksheet. Sheet is identified by its index. The cell is identyfied by its row number and column number.
        /// </summary>
        /// <param name="sheetIndex">The index of the sheet in which to insert formula.</param>
        /// <param name="rowIndex">The row number where the formula is to be set.</param>
        /// <param name="columnIndex">The column number where the formula is to be set.</param>
        /// <param name="formula">Formula to be set.</param>
        public void AddCellFormula(int sheetIndex, int rowIndex, int columnIndex, string formula);

        /// <summary>
        /// Inserts a formula into a cell in a worksheet. Sheet is identified by its name. The cell is identyfied by its row number and column number.
        /// </summary>
        /// <param name="sheetName">The name of the sheet in which to insert formula.</param>
        /// <param name="rowIndex">The row number where the formula is to be set.</param>
        /// <param name="columnIndex">The column number where the formula is to be set.</param>
        /// <param name="formula">Formula to be set.</param>
        public void AddCellFormula(string sheetName, int rowIndex, int columnIndex, string formula);

        /// <summary>
        /// Inserts a formula into a range in a worksheet. The sheet is identified by its index. 
        /// The range is identified by its firstRow number, firstColumn number, lastRow number and lastColumn number.
        /// </summary>
        /// <param name="sheetIndex">The index of the sheet in which to insert formula.</param>
        /// <param name="firstRowIndex">The index of the first row in the range.</param>
        /// <param name="firstColumnIndex">The index of the first column in the range.</param>
        /// <param name="lastRowIndex">The index of the last row in the range.</param>
        /// <param name="lastCoulmnIndex">The index of the last column in the range.</param>
        /// <param name="formula">Formula to be set.</param>
        public void AddCellFormula(int sheetIndex, int firstRowIndex, int firstColumnIndex, int lastRowIndex, int lastCoulmnIndex, string formula);

        /// <summary>
        /// Inserts a formula into a range in a worksheet. The sheet is identified by its name. 
        /// The range is identified by its firstRow number, firstColumn number, lastRow number and lastColumn number.
        /// </summary>
        /// <param name="sheetName">The name of the sheet in which to insert formula.</param>
        /// <param name="firstRowIndex">The index of the first row in the range.</param>
        /// <param name="firstColumnIndex">The index of the first column in the range.</param>
        /// <param name="lastRowIndex">The index of the last row in the range.</param>
        /// <param name="lastCoulmnIndex">The index of the last column in the range.</param>
        /// <param name="formula">Formula to be set.</param>
        public void AddCellFormula(string sheetName, int firstRowIndex, int firstColumnIndex, int lastRowIndex, int lastCoulmnIndex, string formula);

        /// <summary>
        /// Deletes the row (identified by its line number) from the sheet. The sheet is identified by name. 
        /// </summary>
        /// <param name="sheetName">The name of the sheet which to delete row.</param>
        /// <param name="rowIndex">The row number which to delete.</param>
        public void DeleteRow(string sheetName, int rowIndex);

        /// <summary>
        /// Deletes the row (identified by its line number) from the sheet. The sheet is identified by name. 
        /// </summary>
        /// <param name="rowIndex">The row number which to delete.</param>
        /// <param name="sheetName">The name of the sheet which to delete row.</param>
        public void DeleteRow(int rowIndex, string sheetName);

        /// <summary>
        /// Deletes the row (identified by its line number) from the sheet. The sheet is identified by index.
        /// </summary>
        /// <param name="sheetIndex">The index of the sheet in which to delete row.</param>
        /// <param name="rowIndex">The row number which to delete.</param>
        public void DeleteRow(int sheetIndex, int rowIndex);

        /// <summary>
        /// Deletes the rows (identified by its line number) and quantity from the sheet. The sheet is identified by name.
        /// </summary>
        /// <param name="sheetName">The name of the sheet which to delete rows.</param>
        /// <param name="rowIndex">The row number at which to start deleting a row.</param>
        /// <param name="rowCount">Number of rows to delete.</param>
        public void DeleteRow(string sheetName, int rowIndex, int rowCount);

        /// <summary>
        /// Deletes the rows (identified by its line number) and quantity from the sheet. The sheet is identified by name.
        /// </summary>
        /// <param name="rowIndex">The row number at which to start deleting a row.</param>
        /// <param name="sheetName">The name of the sheet which to delete rows.</param>
        /// <param name="rowCount">Number of rows to delete.</param>
        public void DeleteRow(int rowIndex, string sheetName, int rowCount);

        /// <summary>
        /// Deletes the rows (identified by its line number) and quantity from the sheet. The sheet is identified by index.
        /// </summary>
        /// <param name="sheetIndex">The index of the sheet in which to delete rows.</param>
        /// <param name="rowIndex">The row number at which to start deleting a row.</param>
        /// <param name="rowCount">Number of rows to delete.</param>
        public void DeleteRow(int sheetIndex, int rowIndex, int rowCount);

        /// <summary>
        /// Deletes the column (identified by its column number) from the sheet. The sheet is identified by name.
        /// </summary>
        /// <param name="sheetName">The name of the sheet which to delete column.</param>
        /// <param name="columnIndex">The column number which to delete.</param>
        public void DeleteColumn(string sheetName, int columnIndex);

        /// <summary>
        /// Deletes the column (identified by its column number) from the sheet. The sheet is identified by name.
        /// </summary>
        /// <param name="columnIndex">The column number which to delete.</param>
        /// <param name="sheetName">The name of the sheet which to delete column.</param>
        public void DeleteColumn(int columnIndex, string sheetName);

        /// <summary>
        /// Deletes the column (identified by its column number) from the sheet. The sheet is identified by index.
        /// </summary>
        /// <param name="sheetIndex">The index of the sheet which to delete column.</param>
        /// <param name="columnIndex">The column number which to delete.</param>
        public void DeleteColumn(int sheetIndex, int columnIndex);

        /// <summary>
        /// Deletes the columns (identified by its column number) and quantity from the sheet. The sheet is identified by name.
        /// </summary>
        /// <param name="sheetName">The name of the sheet which to delete columns.</param>
        /// <param name="columnIndex">The column number at which to start deleting a column.</param>
        /// <param name="columnCount">Number of columns to delete.</param>
        public void DeleteColumn(string sheetName, int columnIndex, int columnCount);

        /// <summary>
        /// Deletes the columns (identified by its column number) and quantity from the sheet. The sheet is identified by name.
        /// </summary>
        /// <param name="columnIndex">The column number at which to start deleting a column.</param>
        /// <param name="sheetName">The name of the sheet which to delete columns.</param>
        /// <param name="columnCount">Number of columns to delete.</param>
        public void DeleteColumn(int columnIndex, string sheetName, int columnCount);

        /// <summary>
        /// Deletes the columns (identified by its column number) and quantity from the sheet. The sheet is identified by index.
        /// </summary>
        /// <param name="sheetIndex">The index of the sheet which to delete column.</param>
        /// <param name="columnIndex">The column number at which to start deleting a column.</param>
        /// <param name="columnCount">Number of columns to delete.</param>
        public void DeleteColumn(int sheetIndex, int columnIndex, int columnCount);

        /// <summary>
        /// Inserts a integer value into a cell in a worksheet. Sheet is identified by its index. The cell is identyfied by its row number and column number.
        /// </summary>
        /// <param name="sheetIndex">The index of the sheet in which to insert integer value.</param>
        /// <param name="rowIndex">The row number where the value is to be set.</param>
        /// <param name="columnIndex">The column number where the value is to be set.</param>
        /// <param name="value">A integer value.</param>
        public void AddCellValueInt(int sheetIndex, int rowIndex, int columnIndex, int value);

        /// <summary>
        /// Inserts a integer value into a cell and set cell format in a worksheet. Sheet is identified by its index. The cell is identyfied by its row number and column number.
        /// </summary>
        /// <param name="sheetIndex">The index of the sheet in which to insert integer value.</param>
        /// <param name="rowIndex">The row number where the value is to be set.</param>
        /// <param name="columnIndex">The column number where the value is to be set.</param>
        /// <param name="value">A integer value.</param>
        /// <param name="format">Cell format to be set.</param>
        public void AddCellValueInt(int sheetIndex, int rowIndex, int columnIndex, int value, string format);

        /// <summary>
        /// Inserts a integer value into a cell in a worksheet. Sheet is identified by its name. The cell is identyfied by its row number and column number.
        /// </summary>
        /// <param name="sheetName">The name of the sheet in which to insert text value.</param>
        /// <param name="rowIndex">The row number where the value is to be set.</param>
        /// <param name="columnIndex">The column number where the value is to be set.</param>
        /// <param name="value">A integer value.</param>
        public void AddCellValueInt(string sheetName, int rowIndex, int columnIndex, int value);

        /// <summary>
        /// Inserts a integer value into a cell and set cell format in a worksheet. Sheet is identified by its name. The cell is identyfied by its row number and column number.
        /// </summary>
        /// <param name="sheetName">The name of the sheet in which to insert text value.</param>
        /// <param name="rowIndex">The row number where the value is to be set.</param>
        /// <param name="columnIndex">The column number where the value is to be set.</param>
        /// <param name="value">A integer value.</param>
        /// <param name="format">Cell format to be set.</param>
        public void AddCellValueInt(string sheetName, int rowIndex, int columnIndex, int value, string format);

        /// <summary>
        /// Inserts a decimal value into a cell in a worksheet. Sheet is identified by its index. The cell is identyfied by its row number and column number.
        /// </summary>
        /// <param name="sheetIndex">The index of the sheet in which to insert decimal value.</param>
        /// <param name="rowIndex">The row number where the value is to be set.</param>
        /// <param name="columnIndex">The column number where the value is to be set.</param>
        /// <param name="value">A decimal value.</param>
        public void AddCellValueDecimal(int sheetIndex, int rowIndex, int columnIndex, decimal value);

        /// <summary>
        /// Inserts a decimal value into a cell and set cell format in a worksheet. Sheet is identified by its index. The cell is identyfied by its row number and column number.
        /// </summary>
        /// <param name="sheetIndex">The index of the sheet in which to insert decimal value.</param>
        /// <param name="rowIndex">The row number where the value is to be set.</param>
        /// <param name="columnIndex">The column number where the value is to be set.</param>
        /// <param name="value">A decimal value.</param>
        /// <param name="format">Cell format to be set.</param>
        public void AddCellValueDecimal(int sheetIndex, int rowIndex, int columnIndex, decimal value, string format);

        /// <summary>
        /// Inserts a decimal value into a cell in a worksheet. Sheet is identified by its name. The cell is identyfied by its row number and column number.
        /// </summary>
        /// <param name="sheetName">The name of the sheet in which to insert decimal value.</param>
        /// <param name="rowIndex">The row number where the value is to be set.</param>
        /// <param name="columnIndex">The column number where the value is to be set.</param>
        /// <param name="value">A decimal value.</param>
        public void AddCellValueDecimal(string sheetName, int rowIndex, int columnIndex, decimal value);

        /// <summary>
        /// Inserts a decimal value into a cell and set cell format in a worksheet. Sheet is identified by its name. The cell is identyfied by its row number and column number.
        /// </summary>
        /// <param name="sheetName">The name of the sheet in which to insert decimal value.</param>
        /// <param name="rowIndex">The row number where the value is to be set.</param>
        /// <param name="columnIndex">The column number where the value is to be set.</param>
        /// <param name="value">A decimal value.</param>
        /// <param name="format">Cell format to be set.</param>
        public void AddCellValueDecimal(string sheetName, int rowIndex, int columnIndex, decimal value, string format);

        /// <summary>
        /// Inserts a double precision value into a cell in a worksheet. Sheet is identified by its index. The cell is identyfied by its row number and column number.
        /// </summary>
        /// <param name="sheetIndex">The index of the sheet in which to insert double precision value.</param>
        /// <param name="rowIndex">The row number where the value is to be set.</param>
        /// <param name="columnIndex">The column number where the value is to be set.</param>
        /// <param name="value">A double precision value.</param>
        public void AddCellValueDouble(int sheetIndex, int rowIndex, int columnIndex,  double value);

        /// <summary>
        /// Inserts a double precision value into a cell and set cell format in a worksheet. Sheet is identified by its index. The cell is identyfied by its row number and column number.
        /// </summary>
        /// <param name="sheetIndex">The index of the sheet in which to insert double precision value.</param>
        /// <param name="rowIndex">The row number where the value is to be set.</param>
        /// <param name="columnIndex">The column number where the value is to be set.</param>
        /// <param name="value">A double precision value.</param>
        /// <param name="format">Cell format to be set.</param>
        public void AddCellValueDouble(int sheetIndex, int rowIndex, int columnIndex, double value, string format);

        /// <summary>
        /// Inserts a double precision value into a cell in a worksheet. Sheet is identified by its name. The cell is identyfied by its row number and column number.
        /// </summary>
        /// <param name="sheetName">The name of the sheet in which to insert double precision value.</param>
        /// <param name="rowIndex">The row number where the value is to be set.</param>
        /// <param name="columnIndex">The column number where the value is to be set.</param>
        /// <param name="value">A double precision value.</param>
        public void AddCellValueDouble(string sheetName, int rowIndex, int columnIndex, double value);

        /// <summary>
        /// Inserts a double precision value into a cell and set cell format in a worksheet. Sheet is identified by its name. The cell is identyfied by its row number and column number.
        /// </summary>
        /// <param name="sheetName">The name of the sheet in which to insert double precision value.</param>
        /// <param name="rowIndex">The row number where the value is to be set.</param>
        /// <param name="columnIndex">The column number where the value is to be set.</param>
        /// <param name="value">A double precision value.</param>
        /// <param name="format">Cell format to be set.</param>
        public void AddCellValueDouble(string sheetName, int rowIndex, int columnIndex, double value, string format);

        /// <summary>
        /// Sets the height of rows in a sheet. The sheet is identified by index. Rows are identified by row number and quantity (next rows are changed after rowIndex , including rowIndex).
        /// </summary>
        /// <param name="sheetIndex">The index of the sheet in which to change height in row.</param>
        /// <param name="rowIndex">The row number where the height is to be change.</param>
        /// <param name="rowCount">Number of rows to change height.</param>
        /// <param name="height">Row height.</param>
        public void HeightRow(int sheetIndex, int rowIndex, int rowCount, double height);

        /// <summary>
        /// Sets the height of rows in a sheet. The sheet is identified by index. Rows are identified by row number from table rowIndices.
        /// </summary>
        /// <param name="sheetIndex">The index of the sheet in which to change height in rows.</param>
        /// <param name="rowIndices">Array of rows where the height is to be change.</param>
        /// <param name="height">Row height.</param>
        public void HeightRow(int sheetIndex, int[] rowIndices, double height);

        /// <summary>
        /// Sets the height of row in a sheet. The sheet is identified by index. The row is identified by row number.
        /// </summary>
        /// <param name="sheetIndex">The index of the sheet in which to change height in row.</param>
        /// <param name="rowIndex">The row number where the height is to be change.</param>
        /// <param name="height">Row height.</param>
        public void HeightRow(int sheetIndex, int rowIndex, double height);

        /// <summary>
        /// Sets the height of rows in a sheet. The sheet is identified by name. Rows are identified by row number and quantity (next rows are changed after rowIndex , including rowIndex).
        /// </summary>
        /// <param name="sheetName">The name of the sheet in which to change height in rows.</param>
        /// <param name="rowIndex">The row number where the height is to be change.</param>
        /// <param name="rowCount">Number of rows to change height.</param>
        /// <param name="height">Row height.</param>
        public void HeightRow(string sheetName, int rowIndex, int rowCount, double height);

        /// <summary>
        /// Sets the height of rows in a sheet. The sheet is identified by name. Rows are identified by row number from table rowIndices.
        /// </summary>
        /// <param name="sheetName">The name of the sheet in which to change height in rows.</param>
        /// <param name="rowIndices">Array of rows where the height is to be change.</param>
        /// <param name="height">Row height.</param>
        public void HeightRow(string sheetName, int[] rowIndices, double height);

        /// <summary>
        /// Sets the height of row in a sheet. The sheet is identified by name. The row is identified by row number.
        /// </summary>
        /// <param name="sheetName">The name of the sheet in which to change height in row.</param>
        /// <param name="rowIndex">The row number where the height is to be change.</param>
        /// <param name="height">Row height.</param>
        public void HeightRow(string sheetName, int rowIndex, double height);

        /// <summary>
        /// Sets the width of columns in a sheet. The sheet is identified by index. Columnss are identified by column number and quantity (next columns are changed after columnIndex , including columnIndex).
        /// </summary>
        /// <param name="sheetIndex">The index of the sheet in which to change width in columns.</param>
        /// <param name="columnIndex">The column number where the width is to be change.</param>
        /// <param name="columnCount">Number of columns to change height.</param>
        /// <param name="width">Column width.</param>
        public void WidthColumn(int sheetIndex, int columnIndex, int columnCount, double width);

        /// <summary>
        /// Sets the width of columns in a sheet. The sheet is identified by index. Columns are identified by column number from table columnIndices.
        /// </summary>
        /// <param name="sheetIndex">The index of the sheet in which to change width in columns.</param>
        /// <param name="columnIndices">Array of columns where the height is to be change.</param>
        /// <param name="width">Column width.</param>
        public void WidthColumn(int sheetIndex, int[] columnIndices, double width);

        /// <summary>
        /// Sets the width of column in a sheet. The sheet is identified by index. The column is identified by column number.
        /// </summary>
        /// <param name="sheetIndex">The index of the sheet in which to change width in column.</param>
        /// <param name="columnIndex">The column number where the width is to be change.</param>
        /// <param name="width">Column width.</param>
        public void WidthColumn(int sheetIndex, int columnIndex, double width);

        /// <summary>
        /// Sets the width of columns in a sheet. The sheet is identified by name. Columns are identified by column number and quantity (next columns are changed after columnIndex , including columnIndex).
        /// </summary>
        /// <param name="sheetName">The name of the sheet in which to change width in columns.</param>
        /// <param name="columnIndex">The column number where the width is to be change.</param>
        /// <param name="columnCount">Number of columns to change width.</param>
        /// <param name="width">Column width.</param>
        public void WidthColumn(string sheetName, int columnIndex, int columnCount, double width);

        /// <summary>
        /// Sets the width of columns in a sheet. The sheet is identified by name. Columns are identified by column number from table columnIndices.
        /// </summary>
        /// <param name="sheetName">The name of the sheet in which to change width in columns.</param>
        /// <param name="columnIndices">Array of columns where the width is to be change.</param>
        /// <param name="width">Column width.</param>
        public void WidthColumn(string sheetName, int[] columnIndices, double width);

        /// <summary>
        /// Sets the width of column in a sheet. The sheet is identified by name. The column is identified by column number.
        /// </summary>
        /// <param name="sheetName">The name of the sheet in which to change width in column.</param>
        /// <param name="columnIndex">The column number where the width is to be change.</param>
        /// <param name="width">Column width.</param>
        public void WidthColumn(string sheetName, int columnIndex, double width);

        /// <summary>
        /// Inserts a datetime value into a cell in a worksheet. Sheet is identified by its index. The cell is identyfied by its row number and column number.
        /// </summary>
        /// <param name="sheetIndex">The index of the sheet in which to insert datetime value.</param>
        /// <param name="rowIndex">The row number where the value is to be set.</param>
        /// <param name="columnIndex">The column number where the value is to be set.</param>
        /// <param name="date">A datetime value.</param>
        public void AddCellValueDate(int sheetIndex, int rowIndex, int columnIndex, DateTime date);

        /// <summary>
        /// Inserts a datetime value into a cell and set cell format in a worksheet. Sheet is identified by its index. The cell is identyfied by its row number and column number.
        /// </summary>
        /// <param name="sheetIndex">The index of the sheet in which to insert datetime value.</param>
        /// <param name="rowIndex">The row number where the value is to be set.</param>
        /// <param name="columnIndex">The column number where the value is to be set.</param>
        /// <param name="date">A datetime value.</param>
        /// <param name="format">Cell format to be set.</param>
        public void AddCellValueDate(int sheetIndex, int rowIndex, int columnIndex, DateTime date, string format);

        /// <summary>
        /// Inserts a datetime value into a cell in a worksheet. Sheet is identified by its name. The cell is identyfied by its row number and column number.
        /// </summary>
        /// <param name="sheetName">The name of the sheet in which to insert datetime value.</param>
        /// <param name="rowIndex">The row number where the value is to be set.</param>
        /// <param name="columnIndex">The column number where the value is to be set.</param>
        /// <param name="date">A datetime value.</param>
        public void AddCellValueDate(string sheetName, int rowIndex, int columnIndex, DateTime date);

        /// <summary>
        /// Inserts a datetime value into a cell and set cell format in a worksheet. Sheet is identified by its name. The cell is identyfied by its row number and column number.
        /// </summary>
        /// <param name="sheetName">The name of the sheet in which to insert datetime value.</param>
        /// <param name="rowIndex">The row number where the value is to be set.</param>
        /// <param name="columnIndex">The column number where the value is to be set.</param>
        /// <param name="date">A datetime value.</param>
        /// <param name="format">Cell format to be set.</param>
        public void AddCellValueDate(string sheetName, int rowIndex, int columnIndex, DateTime date, string format);

        /// <summary>
        /// Inserts a text value into a cell in a worksheet. Sheet is identified by its index. The cell is identyfied by its row number and column number.
        /// </summary>
        /// <param name="sheetIndex">The index of the sheet in which to insert text value.</param>
        /// <param name="rowIndex">The row number where the value is to be set.</param>
        /// <param name="columnIndex">The column number where the value is to be set.</param>
        /// <param name="currency">A currency vakue.</param>
        public void AddCellValueCurrency(int sheetIndex, int rowIndex, int columnIndex, Currency currency);

        /// <summary>
        /// Inserts a text value into a cell in a worksheet. Sheet is identified by its name. The cell is identyfied by its row number and column number.
        /// </summary>
        /// <param name="sheetName">The name of the sheet in which to insert text value.</param>
        /// <param name="rowIndex">The row number where the value is to be set.</param>
        /// <param name="columnIndex">The column number where the value is to be set.</param>
        /// <param name="currency">A currency vakue.</param>
        public void AddCellValueCurrency(string sheetName, int rowIndex, int columnIndex, Currency currency);

        /// <summary>
        /// Inserts a percent value into a cell in a worksheet. Sheet is identified by its index. The cell is identyfied by its row number and column number.
        /// </summary>
        /// <param name="sheetIndex">The index of the sheet in which to insert percent value.</param>
        /// <param name="rowIndex">The row number where the value is to be set.</param>
        /// <param name="columnIndex">The column number where the value is to be set.</param>
        /// <param name="percent">A percent value.</param>
        public void AddCellValuePercent(int sheetIndex, int rowIndex, int columnIndex, Percent percent);

        /// <summary>
        /// Inserts a percent value into a cell and set cell format in a worksheet. Sheet is identified by its index. The cell is identyfied by its row number and column number.
        /// </summary>
        /// <param name="sheetIndex">The index of the sheet in which to insert percent value.</param>
        /// <param name="rowIndex">The row number where the value is to be set.</param>
        /// <param name="columnIndex">The column number where the value is to be set.</param>
        /// <param name="percent">A percent value.</param>
        /// <param name="format">Cell format to be set.</param>
        public void AddCellValuePercent(int sheetIndex, int rowIndex, int columnIndex, Percent percent, string format);

        /// <summary>
        /// Inserts a percent value into a cell in a worksheet. Sheet is identified by its name. The cell is identyfied by its row number and column number.
        /// </summary>
        /// <param name="sheetName">The name of the sheet in which to insert percent value.</param>
        /// <param name="rowIndex">The row number where the value is to be set.</param>
        /// <param name="columnIndex">The column number where the value is to be set.</param>
        /// <param name="percent">A percent value.</param>
        public void AddCellValuePercent(string sheetName, int rowIndex, int columnIndex, Percent percent);

        /// <summary>
        /// Inserts a percent value into a cell and set cell format in a worksheet. Sheet is identified by its name. The cell is identyfied by its row number and column number.
        /// </summary>
        /// <param name="sheetName">The name of the sheet in which to insert percent value.</param>
        /// <param name="rowIndex">The row number where the value is to be set.</param>
        /// <param name="columnIndex">The column number where the value is to be set.</param>
        /// <param name="percent">A percent value.</param>
        /// <param name="format">Cell format to be set.</param>
        public void AddCellValuePercent(string sheetName, int rowIndex, int columnIndex, Percent percent, string format);

        /// <summary>
        /// Inserts a text value into a cell in a worksheet. Sheet is identified by its index. The cell is identyfied by its row number and column number.
        /// </summary>
        /// <param name="sheetIndex">The index of the sheet in which to insert text value.</param>
        /// <param name="rowIndex">The row number where the value is to be set.</param>
        /// <param name="columnIndex">The column number where the value is to be set.</param>
        /// <param name="time">A time value.</param>
        public void AddCellValueTime(int sheetIndex, int rowIndex, int columnIndex, Time time);

        /// <summary>
        /// Inserts a text value into a cell in a worksheet. Sheet is identified by its name. The cell is identyfied by its row number and column number.
        /// </summary>
        /// <param name="sheetName">The name of the sheet in which to insert text value.</param>
        /// <param name="rowIndex">The row number where the value is to be set.</param>
        /// <param name="columnIndex">The column number where the value is to be set.</param>
        /// <param name="time">A time value.</param>
        public void AddCellValueTime(string sheetName, int rowIndex, int columnIndex, Time time);

        /// <summary>
        /// Inserts a fraction value into a cell in a worksheet. Sheet is identified by its index. The cell is identyfied by its row number and column number.
        /// </summary>
        /// <param name="sheetIndex">The index of the sheet in which to insert fraction value.</param>
        /// <param name="rowIndex">The row number where the value is to be set.</param>
        /// <param name="columnIndex">The column number where the value is to be set.</param>
        /// <param name="fraction">A fraction value.</param>
        public void AddCellValueFraction(int sheetIndex, int rowIndex, int columnIndex, Fraction fraction);

        /// <summary>
        /// Inserts a fraction value into a cell and set cell format in a worksheet. Sheet is identified by its index. The cell is identyfied by its row number and column number.
        /// </summary>
        /// <param name="sheetIndex">The index of the sheet in which to insert fraction value.</param>
        /// <param name="rowIndex">The row number where the value is to be set.</param>
        /// <param name="columnIndex">The column number where the value is to be set.</param>
        /// <param name="fraction">A fraction value.</param>
        /// <param name="format">Cell format to be set.</param>
        public void AddCellValueFraction(int sheetIndex, int rowIndex, int columnIndex, Fraction fraction, string format);

        /// <summary>
        /// Inserts a fraction value into a cell in a worksheet. Sheet is identified by its name. The cell is identyfied by its row number and column number.
        /// </summary>
        /// <param name="sheetName">The name of the sheet in which to insert fraction value.</param>
        /// <param name="rowIndex">The row number where the value is to be set.</param>
        /// <param name="columnIndex">The column number where the value is to be set.</param>
        /// <param name="fraction">A fraction value.</param>
        public void AddCellValueFraction(string sheetName, int rowIndex, int columnIndex, Fraction fraction);

        /// <summary>
        /// Inserts a fraction value into a cell and set cell format in a worksheet. Sheet is identified by its name. The cell is identyfied by its row number and column number.
        /// </summary>
        /// <param name="sheetName">The name of the sheet in which to insert fraction value.</param>
        /// <param name="rowIndex">The row number where the value is to be set.</param>
        /// <param name="columnIndex">The column number where the value is to be set.</param>
        /// <param name="fraction">A fraction value.</param>
        /// <param name="format">Cell format to be set.</param>
        public void AddCellValueFraction(string sheetName, int rowIndex, int columnIndex, Fraction fraction, string format);

        /// <summary>
        /// Sets the default black border for a cell. Borders are set from the borderIndex enum, and the line style is set from the LinesIndex. 
        /// The sheet is identified by its name. The cell is identified by its row number and column number.
        /// </summary>
        /// <param name="borderIndex">An array of edges to set.</param>
        /// <param name="lineIndex">Line style to set.</param>
        /// <param name="nameSheet">The name of the sheet where the border is to be set.</param>
        /// <param name="rowIndex">The row number where the border is to be set.</param>
        /// <param name="columnIndex">The column number where the border is to be set.</param>
        public void SetBorder(BorderIndex[] borderIndex, LinesIndex lineIndex, string nameSheet, int rowIndex, int columnIndex);

        /// <summary>
        /// Sets the default black border for a cell. Borders are set from the borderIndex enum, and the line style is set from the LinesIndex. 
        /// The sheet is identified by its index. The cell is identified by its row number and column number.
        /// </summary>
        /// <param name="borderIndex">An array of edges to set.</param>
        /// <param name="lineIndex">Line style to set.</param>
        /// <param name="sheetIndex">The index of the sheet where the border is to be set.</param>
        /// <param name="rowIndex">The row number where the border is to be set.</param>
        /// <param name="columnIndex">The column number where the border is to be set.</param>
        public void SetBorder(BorderIndex[] borderIndex, LinesIndex lineIndex, int sheetIndex, int rowIndex, int columnIndex);

        /// <summary>
        /// Sets the color border for a cell. Borders are set from the borderIndex enum, and the line style is set from the LinesIndex. 
        /// The sheet is identified by its name. The cell is identified by its row number and column number.
        /// </summary>
        /// <param name="borderIndex">An array of edges to set.</param>
        /// <param name="lineIndex">Line style to set.</param>
        /// <param name="nameSheet">The name of the sheet where the border is to be set.</param>
        /// <param name="rowIndex">The row number where the border is to be set.</param>
        /// <param name="columnIndex">The column number where the border is to be set.</param>
        /// <param name="a">The alpha component of the color (0-255).</param>
        /// <param name="r">The red component of the color (0-255).</param>
        /// <param name="g">The green component of the color (0-255).</param>
        /// <param name="b">The blue component of the color (0-255).</param>
        public void SetBorder(BorderIndex[] borderIndex, LinesIndex lineIndex, string nameSheet, int rowIndex, int columnIndex, int a, int r, int g, int b);

        /// <summary>
        /// Sets the color border for a cell. Borders are set from the borderIndex enum, and the line style is set from the LinesIndex. 
        /// The sheet is identified by its index. The cell is identified by its row number and column number.
        /// </summary>
        /// <param name="borderIndex">An array of edges to set.</param>
        /// <param name="lineIndex">Line style to set.</param>
        /// <param name="sheetIndex">The index of the sheet where the border is to be set.</param>
        /// <param name="rowIndex">The row number where the border is to be set.</param>
        /// <param name="columnIndex">The column number where the border is to be set.</param>
        /// <param name="a">The alpha component of the color (0-255).</param>
        /// <param name="r">The red component of the color (0-255).</param>
        /// <param name="g">The green component of the color (0-255).</param>
        /// <param name="b">The blue component of the color (0-255).</param>
        public void SetBorder(BorderIndex[] borderIndex, LinesIndex lineIndex, int sheetIndex, int rowIndex, int columnIndex, int a, int r, int g, int b);

        /// <summary>
        /// Sets the default black border for a range. Borders are set from the borderIndex enum, and the line style is set from the LinesIndex. 
        /// The sheet is identified by its name. The range is identified by its firstRow number, firstColumn number, lastRow number and lastColumn number.
        /// </summary>
        /// <param name="borderIndex">An array of edges to set.</param>
        /// <param name="lineIndex">Line style to set.</param>
        /// <param name="nameSheet">The name of the sheet where the border is to be set.</param>
        /// <param name="firstRowIndex">The index of the first row in the range.</param>
        /// <param name="firstColumnIndex">The index of the first column in the range.</param>
        /// <param name="lastRowIndex">The index of the last row in the range.</param>
        /// <param name="lastColumnIndex">The index of the last column in the range.</param>
        /// <param name="allRange">A logical value indicating whether the operation should be applied to the entire range or only to the border. True = allRange, False = !allRange</param>
        public void SetBorder(BorderIndex[] borderIndex, LinesIndex lineIndex, string nameSheet, int firstRowIndex, int firstColumnIndex, int lastRowIndex, int lastColumnIndex, bool allRange);

        /// <summary>
        /// Sets the default black border for a range. Borders are set from the borderIndex enum, and the line style is set from the LinesIndex. 
        /// The sheet is identified by its index. The range is identified by its firstRow number, firstColumn number, lastRow number and lastColumn number.
        /// </summary>
        /// <param name="borderIndex">An array of edges to set.</param>
        /// <param name="lineIndex">Line style to set.</param>
        /// <param name="sheetIndex">The index of the sheet where the border is to be set.</param>
        /// <param name="firstRowIndex">The index of the first row in the range.</param>
        /// <param name="firstColumnIndex">The index of the first column in the range.</param>
        /// <param name="lastRowIndex">The index of the last row in the range.</param>
        /// <param name="lastColumnIndex">The index of the last column in the range.</param>
        /// <param name="allRange">A logical value indicating whether the operation should be applied to the entire range or only to the border. True = allRange, False = !allRange</param>
        public void SetBorder(BorderIndex[] borderIndex, LinesIndex lineIndex, int sheetIndex, int firstRowIndex, int firstColumnIndex, int lastRowIndex, int lastColumnIndex, bool allRange);

        /// <summary>
        /// Sets the color border for a range. Borders are set from the borderIndex enum, and the line style is set from the LinesIndex. 
        /// The sheet is identified by its name. The range is identified by its firstRow number, firstColumn number, lastRow number and lastColumn number.
        /// </summary>
        /// <param name="borderIndex">An array of edges to set.</param>
        /// <param name="lineIndex">Line style to set.</param>
        /// <param name="nameSheet">The name of the sheet where the border is to be set.</param>
        /// <param name="firstRowIndex">The index of the first row in the range.</param>
        /// <param name="firstColumnIndex">The index of the first column in the range.</param>
        /// <param name="lastRowIndex">The index of the last row in the range.</param>
        /// <param name="lastColumnIndex">The index of the last column in the range.</param>
        /// <param name="a">The alpha component of the color (0-255).</param>
        /// <param name="r">The red component of the color (0-255).</param>
        /// <param name="g">The green component of the color (0-255).</param>
        /// <param name="b">The blue component of the color (0-255).</param>
        /// <param name="allRange">A logical value indicating whether the operation should be applied to the entire range or only to the border. True = allRange, False = !allRange</param>
        public void SetBorder(BorderIndex[] borderIndex, LinesIndex lineIndex, string nameSheet, int firstRowIndex, int firstColumnIndex, int lastRowIndex, int lastColumnIndex, int a, int r, int g, int b, bool allRange);

        /// <summary>
        /// Sets the color border for a range. Borders are set from the borderIndex enum, and the line style is set from the LinesIndex. 
        /// The sheet is identified by its index. The range is identified by its firstRow number, firstColumn number, lastRow number and lastColumn number.
        /// </summary>
        /// <param name="borderIndex">An array of edges to set.</param>
        /// <param name="lineIndex">Line style to set.</param>
        /// <param name="sheetIndex">The index of the sheet where the border is to be set.</param>
        /// <param name="firstRowIndex">The index of the first row in the range.</param>
        /// <param name="firstColumnIndex">The index of the first column in the range.</param>
        /// <param name="lastRowIndex">The index of the last row in the range.</param>
        /// <param name="lastColumnIndex">The index of the last column in the range.</param>
        /// <param name="a">The alpha component of the color (0-255).</param>
        /// <param name="r">The red component of the color (0-255).</param>
        /// <param name="g">The green component of the color (0-255).</param>
        /// <param name="b">The blue component of the color (0-255).</param>
        /// <param name="allRange">A logical value indicating whether the operation should be applied to the entire range or only to the border. True = allRange, False = !allRange</param>
        public void SetBorder(BorderIndex[] borderIndex, LinesIndex lineIndex, int sheetIndex, int firstRowIndex, int firstColumnIndex, int lastRowIndex, int lastColumnIndex, int a, int r, int g, int b, bool allRange);

        /// <summary>
        /// Sets the cell color. The sheet is identified by its index. The cell is identified by its row number and column number.
        /// </summary>
        /// <param name="sheetIndex">The index of the sheet where the cell color is to be set.</param>
        /// <param name="rowIndex">The row number where the cell color is to be set.</param>
        /// <param name="columnIndex">The column number where the cell color is to be set.</param>
        /// <param name="a">The alpha component of the color (0-255).</param>
        /// <param name="r">The red component of the color (0-255).</param>
        /// <param name="g">The green component of the color (0-255).</param>
        /// <param name="b">The blue component of the color (0-255).</param>
        public void CellColor(int sheetIndex, int rowIndex, int columnIndex, int a, int r, int g, int b);

        /// <summary>
        /// Sets the cell color. The sheet is identified by its name. The cell is identified by its row number and column number.
        /// </summary>
        /// <param name="sheetName">The name of the sheet where the cell color is to be set.</param>
        /// <param name="rowIndex">The row number where the cell color is to be set.</param>
        /// <param name="columnIndex">The column number where the cell color is to be set.</param>
        /// <param name="a">The alpha component of the color (0-255).</param>
        /// <param name="r">The red component of the color (0-255).</param>
        /// <param name="g">The green component of the color (0-255).</param>
        /// <param name="b">The blue component of the color (0-255).</param>
        public void CellColor(string sheetName, int rowIndex, int columnIndex, int a, int r, int g, int b);

        /// <summary>
        /// Sets the cell color for a range. The sheet is identified by its index. 
        /// The range is identified by its firstRow number, firstColumn number, lastRow number and lastColumn number.
        /// </summary>
        /// <param name="sheetIndex">The index of the sheet where the cell color is to be set.</param>
        /// <param name="firstRowIndex">The index of the first row in the range.</param>
        /// <param name="firstColumnIndex">The index of the first column in the range.</param>
        /// <param name="lastRowIndex">The index of the last row in the range.</param>
        /// <param name="lastColumnIndex">The index of the last column in the range.</param>
        /// <param name="a">The alpha component of the color (0-255).</param>
        /// <param name="r">The red component of the color (0-255).</param>
        /// <param name="g">The green component of the color (0-255).</param>
        /// <param name="b">The blue component of the color (0-255).</param>
        public void CellColor(int sheetIndex, int firstRowIndex, int firstColumnIndex, int lastRowIndex, int lastColumnIndex, int a, int r, int g, int b);

        /// <summary>
        /// Sets the cell color for a range. The sheet is identified by its index. 
        /// The range is identified by its firstRow number, firstColumn number, lastRow number and lastColumn number.
        /// </summary>
        /// <param name="sheetName">The name of the sheet where the cell color is to be set.</param>
        /// <param name="firstRowIndex">The index of the first row in the range.</param>
        /// <param name="firstColumnIndex">The index of the first column in the range.</param>
        /// <param name="lastRowIndex">The index of the last row in the range.</param>
        /// <param name="lastColumnIndex">The index of the last column in the range.</param>
        /// <param name="a">The alpha component of the color (0-255).</param>
        /// <param name="r">The red component of the color (0-255).</param>
        /// <param name="g">The green component of the color (0-255).</param>
        /// <param name="b">The blue component of the color (0-255).</param>
        public void CellColor(string sheetName, int firstRowIndex, int firstColumnIndex, int lastRowIndex, int lastColumnIndex, int a, int r, int g, int b);

        // text alignment
        /// <summary>
        /// Sets the vertical alignment for a cell. The sheet is identified by its index. The cell is identified by its row number and column number.
        /// </summary>
        /// <param name="verticalAlignmentIndex">Vertical alignment to set.</param>
        /// <param name="sheetIndex">The index of the sheet where the cell color is to be set.</param>
        /// <param name="rowIndex">The row number where the vertical alignment is to be set.</param>
        /// <param name="columnIndex">The column number where the vertical alignment is to be set.</param>
        public void SetVerticalAlignment(VerticalAlignmentIndex verticalAlignmentIndex, int sheetIndex, int rowIndex, int columnIndex);

        /// <summary>
        /// Sets the vertical alignment for a cell. The sheet is identified by its name. The cell is identified by its row number and column number.
        /// </summary>
        /// <param name="verticalAlignmentIndex">Vertical alignment to set.</param>
        /// <param name="sheetName">The name of the sheet where the cell color is to be set.</param>
        /// <param name="rowIndex">The row number where the vertical alignment is to be set.</param>
        /// <param name="columnIndex">The column number where the vertical alignment is to be set.</param>
        public void SetVerticalAlignment(VerticalAlignmentIndex verticalAlignmentIndex, string sheetName, int rowIndex, int columnIndex);

        /// <summary>
        /// Sets the vertical alignment for a range. The sheet is identified by its index. 
        /// The range is identified by its firstRow number, firstColumn number, lastRow number and lastColumn number.
        /// </summary>
        /// <param name="verticalAlignmentIndex">Vertical alignment to set.</param>
        /// <param name="sheetIndex">The index of the sheet where the cell color is to be set.</param>
        /// <param name="firstRowIndex">The index of the first row in the range.</param>
        /// <param name="firstColumnIndex">The index of the first column in the range.</param>
        /// <param name="lastRowIndex">The index of the last row in the range.</param>
        /// <param name="lastColumnIndex">The index of the last column in the range.</param>
        public void SetVerticalAlignment(VerticalAlignmentIndex verticalAlignmentIndex, int sheetIndex, int firstRowIndex, int firstColumnIndex, int lastRowIndex, int lastColumnIndex);

        /// <summary>
        /// Sets the vertical alignment for a range. The sheet is identified by its name. 
        /// The range is identified by its firstRow number, firstColumn number, lastRow number and lastColumn number.
        /// </summary>
        /// <param name="verticalAlignmentIndex">Vertical alignment to set.</param>
        /// <param name="sheetName">The name of the sheet where the cell color is to be set.</param>
        /// <param name="firstRowIndex">The index of the first row in the range.</param>
        /// <param name="firstColumnIndex">The index of the first column in the range.</param>
        /// <param name="lastRowIndex">The index of the last row in the range.</param>
        /// <param name="lastColumnIndex">The index of the last column in the range.</param>
        public void SetVerticalAlignment(VerticalAlignmentIndex verticalAlignmentIndex, string sheetName, int firstRowIndex, int firstColumnIndex, int lastRowIndex, int lastColumnIndex);

        /// <summary>
        /// Sets the horizontal alignment for a cell. The sheet is identified by its index. The cell is identified by its row number and column number.
        /// </summary>
        /// <param name="horizontalAlignmentIndex">Horizontal alignment to set.</param>
        /// <param name="sheetIndex">The index of the sheet where the cell color is to be set.</param>
        /// <param name="rowIndex">The row number where the horizontal alignment is to be set.</param>
        /// <param name="columnIndex">The column number where the horizontal alignment is to be set.</param>
        public void SetHorizontalAlignment(HorizontalAlignmentIndex horizontalAlignmentIndex, int sheetIndex, int rowIndex, int columnIndex);

        /// <summary>
        /// Sets the horizontal alignment for a cell. The sheet is identified by its name. The cell is identified by its row number and column number.
        /// </summary>
        /// <param name="horizontalAlignmentIndex">Horizontal alignment to set.</param>
        /// <param name="sheetName">The name of the sheet where the cell color is to be set.</param>
        /// <param name="rowIndex">The row number where the horizontal alignment is to be set.</param>
        /// <param name="columnIndex">The column number where the horizontal alignment is to be set.</param>
        public void SetHorizontalAlignment(HorizontalAlignmentIndex horizontalAlignmentIndex, string sheetName, int rowIndex, int columnIndex);

        /// <summary>
        /// Sets the horizontal alignment for a range. The sheet is identified by its index. 
        /// The range is identified by its firstRow number, firstColumn number, lastRow number and lastColumn number.
        /// </summary>
        /// <param name="horizontalAlignmentIndex">Horizontal alignment to set.</param>
        /// <param name="sheetIndex">The index of the sheet where the cell color is to be set.</param>
        /// <param name="firstRowIndex">The index of the first row in the range.</param>
        /// <param name="firstColumnIndex">The index of the first column in the range.</param>
        /// <param name="lastRowIndex">The index of the last row in the range.</param>
        /// <param name="lastColumnIndex">The index of the last column in the range.</param>
        public void SetHorizontalAlignment(HorizontalAlignmentIndex horizontalAlignmentIndex, int sheetIndex, int firstRowIndex, int firstColumnIndex, int lastRowIndex, int lastColumnIndex);

        /// <summary>
        /// Sets the horizontal alignment for a range. The sheet is identified by its name. 
        /// The range is identified by its firstRow number, firstColumn number, lastRow number and lastColumn number.
        /// </summary>
        /// <param name="horizontalAlignmentIndex">Horizontal alignment to set.</param>
        /// <param name="sheetName">The name of the sheet where the cell color is to be set.</param>
        /// <param name="firstRowIndex">The index of the first row in the range.</param>
        /// <param name="firstColumnIndex">The index of the first column in the range.</param>
        /// <param name="lastRowIndex">The index of the last row in the range.</param>
        /// <param name="lastColumnIndex">The index of the last column in the range.</param>
        public void SetHorizontalAlignment(HorizontalAlignmentIndex horizontalAlignmentIndex, string sheetName, int firstRowIndex, int firstColumnIndex, int lastRowIndex, int lastColumnIndex);

        /// <summary>
        /// Sets bold, italic, underline, double underline, font and font size, text strikethrough, text wrapping for a cell.
        /// The sheet is identified by its index. The cell is identified by its row number and column number.
        /// </summary>
        /// <param name="fontSettings">Configures a font using the Fluent API.</param>
        /// <example>
        /// Here is an example of how to use the method:
        /// <code>
        /// excelOperation.SetFont(new FontSettings()
        ///     .SetBold(true)
        ///     .SetTextCrossed(true)
        ///     .SetUnderline(true), sheetIndex, rowIndex, columnIndex);
        /// </code>
        /// </example>
        /// <param name="sheetIndex">The index of the sheet where the font is to be set."</param>
        /// <param name="rowIndex">The row number where the font is to be set.</param>
        /// <param name="columnIndex">The column number where the font is to be set.</param>
        public void SetFont(FontSettings fontSettings, int sheetIndex, int rowIndex, int columnIndex);

        /// <summary>
        /// Sets bold, italic, underline, double underline, font and font size, text strikethrough, text wrapping for a cell.
        /// The sheet is identified by its name. The cell is identified by its row number and column number.
        /// </summary>
        /// <param name="fontSettings">Configures a font using the Fluent API.</param>
        /// <example>
        /// Here is an example of how to use the method:
        /// <code>
        /// excelOperation.SetFont(new FontSettings()
        ///     .SetBold(true)
        ///     .SetTextCrossed(true)
        ///     .SetUnderline(true), sheetIndex, rowIndex, columnIndex);
        /// </code>
        /// </example>
        /// <param name="sheetName">The name of the sheet where the font is to be set."</param>
        /// <param name="rowIndex">The row number where the font is to be set.</param>
        /// <param name="columnIndex">The column number where the font is to be set.</param>
        public void SetFont(FontSettings fontSettings, string sheetName, int rowIndex, int columnIndex);

        /// <summary>
        /// Sets bold, italic, underline, double underline, font and font size, text strikethrough, text wrapping for a range.
        /// The sheet is identified by its index. The range is identified by its firstRow number, firstColumn number, lastRow number and lastColumn number.
        /// </summary>
        /// <param name="fontSettings">Configures a font using the Fluent API.</param>
        /// <example>
        /// Here is an example of how to use the method:
        /// <code>
        /// excelOperation.SetFont(new FontSettings()
        ///     .SetBold(true)
        ///     .SetTextCrossed(true)
        ///     .SetUnderline(true), sheetIndex, rowIndex, columnIndex);
        /// </code>
        /// </example>
        /// <param name="sheetIndex">The index of the sheet where the font is to be set."</param>
        /// <param name="firstRowIndex">The index of the first row in the range.</param>
        /// <param name="firstColumnIndex">The index of the first column in the range.</param>
        /// <param name="lastRowIndex">The index of the last row in the range.</param>
        /// <param name="lastColumnIndex">The index of the last column in the range.</param>
        public void SetFont(FontSettings fontSettings, int sheetIndex, int firstRowIndex, int firstColumnIndex, int lastRowIndex, int lastColumnIndex);

        /// <summary>
        /// Sets bold, italic, underline, double underline, font and font size, text strikethrough, text wrapping for a range.
        /// The sheet is identified by its name. The range is identified by its firstRow number, firstColumn number, lastRow number and lastColumn number.
        /// </summary>
        /// <param name="fontSettings">Configures a font using the Fluent API.</param>
        /// <example>
        /// Here is an example of how to use the method:
        /// <code>
        /// excelOperation.SetFont(new FontSettings()
        ///     .SetBold(true)
        ///     .SetTextCrossed(true)
        ///     .SetUnderline(true), sheetIndex, rowIndex, columnIndex);
        /// </code>
        /// </example>
        /// <param name="sheetName">The name of the sheet where the font is to be set."</param>
        /// <param name="firstRowIndex">The index of the first row in the range.</param>
        /// <param name="firstColumnIndex">The index of the first column in the range.</param>
        /// <param name="lastRowIndex">The index of the last row in the range.</param>
        /// <param name="lastColumnIndex">The index of the last column in the range.</param>
        public void SetFont(FontSettings fontSettings, string sheetName, int firstRowIndex, int firstColumnIndex, int lastRowIndex, int lastColumnIndex);

        /// <summary>
        /// Merges cells. The sheet is identified by its index. The range is identified by its firstRow number, firstColumn number, lastRow number and lastColumn number.
        /// </summary>
        /// <param name="sheetIndex">The index of the sheet where the cells are to be merged.</param>
        /// <param name="firstRowIndex">The index of the first row in the range.</param>
        /// <param name="lastRowIndex">The index of the last row in the range.</param>
        /// <param name="firstColumnIndex">The index of the first column in the range.</param>
        /// <param name="lastColumnIndex">The index of the last column in the range.</param>
        public void MergeCells(int sheetIndex, int firstRowIndex, int lastRowIndex, int firstColumnIndex, int lastColumnIndex);

        /// <summary>
        /// Merges cells. The sheet is identified by its name. The range is identified by its firstRow number, firstColumn number, lastRow number and lastColumn number.
        /// </summary>
        /// <param name="sheetName">The name of the sheet where the cells are to be merged.</param>
        /// <param name="firstRowIndex">The index of the first row in the range.</param>
        /// <param name="lastRowIndex">The index of the last row in the range.</param>
        /// <param name="firstColumnIndex">The index of the first column in the range.</param>
        /// <param name="lastColumnIndex">The index of the last column in the range.</param>
        public void MergeCells(string sheetName, int firstRowIndex, int lastRowIndex, int firstColumnIndex, int lastColumnIndex);

        //orientation

        /// <summary>
        /// Sets the orientation of the values ​​in the cell. The sheet is identified by its index. The cell is identified by its row number and column number.
        /// </summary>
        /// <param name="sheetIndex">The index of the sheet where the cell orientation is to be set.</param>
        /// <param name="rowIndex">The row number in which the orientation of the cell content is to be set.</param>
        /// <param name="columnIndex">The column number in which the orientation of the cell content is to be set.</param>
        /// <param name="orientation">Orientation value in degrees. (0-180)</param>
        public void ValueOrientation(int sheetIndex, int rowIndex, int columnIndex, short orientation);

        /// <summary>
        /// Sets the orientation of the values ​​in the cell. The sheet is identified by its name. The cell is identified by its row number and column number.
        /// </summary>
        /// <param name="sheetName">The name of the sheet where the cell orientation is to be set.</param>
        /// <param name="rowIndex">The row number in which the orientation of the cell content is to be set.</param>
        /// <param name="columnIndex">The column number in which the orientation of the cell content is to be set.</param>
        /// <param name="orientation">Orientation value in degrees. (0-180)</param>
        public void ValueOrientation(string sheetName, int rowIndex, int columnIndex, short orientation);

        /// <summary>
        /// Sets the orientation of the values ​​for a range. The sheet is identified by its index. 
        /// The range is identified by its firstRow number, firstColumn number, lastRow number and lastColumn number.
        /// </summary>
        /// <param name="sheetIndex">The index of the sheet where the cell orientation is to be set.</param>
        /// <param name="firstRowIndex">The index of the first row in the range.</param>
        /// <param name="firstColumnIndex">The index of the first column in the range.</param>
        /// <param name="lastRowIndex">The index of the last row in the range.</param>
        /// <param name="lastColumnIndex">The index of the last column in the range.</param>
        /// <param name="orientation">Orientation value in degrees. (0-180)</param>
        public void ValueOrientation(int sheetIndex, int firstRowIndex, int firstColumnIndex, int lastRowIndex, int lastColumnIndex, short orientation);

        /// <summary>
        /// Sets the orientation of the values ​​for a range. The sheet is identified by its name. 
        /// The range is identified by its firstRow number, firstColumn number, lastRow number and lastColumn number.
        /// </summary>
        /// <param name="sheetName">The name of the sheet where the cell orientation is to be set.</param>
        /// <param name="firstRowIndex">The index of the first row in the range.</param>
        /// <param name="firstColumnIndex">The index of the first column in the range.</param>
        /// <param name="lastRowIndex">The index of the last row in the range.</param>
        /// <param name="lastColumnIndex">The index of the last column in the range.</param>
        /// <param name="orientation">Orientation value in degrees. (0-180)</param>
        public void ValueOrientation(string sheetName, int firstRowIndex, int firstColumnIndex, int lastRowIndex, int lastColumnIndex, short orientation);

        /// <summary>
        /// Sets sheet protection. The sheet is identified by its index.
        /// </summary>
        /// <param name="sheetIndex">The index of the sheet which the sheet is to be protected.</param>
        /// <param name="password">Password for sheet.</param>
        public void SetProtectSheet(int sheetIndex, string password);

        /// <summary>
        /// Sets sheet protection. The sheet is identified by its index.
        /// </summary>
        /// <param name="sheetName">The name of the sheet which the sheet is to be protected.</param>
        /// <param name="password">Password for sheet.</param>
        public void SetProtectSheet(string sheetName, string password);

        /// <summary>
        /// Sets the dropdown list in a range. The sheet is identified by its index (where is drop-down list). The range named is identify by its name.
        /// The range is identified by its firstRow number, firstColumn number, lastRow number and lastColumn number.
        /// </summary>
        /// <param name="sheetIndex">The index of the sheet in which the drop-down list is to be located.</param>
        /// <param name="namedRange">Range name set in the Name Manager.</param>
        /// <param name="firstRowIndex">The index of the first row in the range where the dropdown list is to be.</param>
        /// <param name="firstColumnIndex">The index of the first column in the range where the dropdown list is to be.</param>
        /// <param name="lastRowIndex">The index of the last row in the range where the dropdown list is to be.</param>
        /// <param name="lastColumnIndex">The index of the last column in the range where the dropdown list is to be.</param>
        public void DropDownList(int sheetIndex, string namedRange, int firstRowIndex, int firstColumnIndex, int lastRowIndex, int lastColumnIndex);

        /// <summary>
        /// Sets the dropdown list in a range. The sheet is identified by its name (where is drop-down list). The range name is identify by its name.
        /// The range is identified by its firstRow number, firstColumn number, lastRow number and lastColumn number.
        /// </summary>
        /// <param name="sheetName">The name of the sheet in which the drop-down list is to be located.</param>
        /// <param name="namedRange">Range name set in the Name Manager.</param>
        /// <param name="firstRowIndex">The index of the first row in the range where the dropdown list is to be.</param>
        /// <param name="firstColumnIndex">The index of the first column in the range where the dropdown list is to be.</param>
        /// <param name="lastRowIndex">The index of the last row in the range where the dropdown list is to be.</param>
        /// <param name="lastColumnIndex">The index of the last column in the range where the dropdown list is to be.</param>
        public void DropDownList(string sheetName, string namedRange, int firstRowIndex, int firstColumnIndex, int lastRowIndex, int lastColumnIndex);

        /// <summary>
        /// Automatically sets the column width to the width cell contents. The sheet is identified by its index. The column is identified by its column number.
        /// </summary>
        /// <param name="sheetIndex">The index of the sheet in which the column width is to be automatically set to the width cell contents.</param>
        /// <param name="columnIndex">Column number for which the width to be automatically set.</param>
        public void SetAutoWidth(int sheetIndex, int columnIndex);

        /// <summary>
        /// Automatically sets the column width to the width cell contents. The sheet is identified by its name. The column is identified by its column number.
        /// </summary>
        /// <param name="sheetName">The name of the sheet in which the column width is to be automatically set to the width cell contents.</param>
        /// <param name="columnIndex">Column number for which the width to be automatically set.</param>
        public void SetAutoWidth(string sheetName, int columnIndex);

        /// <summary>
        /// Automatically sets the columns width to the width cell contenst from firtstColumnIndex to lastColumnIndex. The sheet is identified by its index.
        /// The range is identified by its firstColumn number and lastColumn number.
        /// </summary>
        /// <param name="sheetIndex">The index of the sheet in which the column width is to be automatically set to the width cell contents.</param>
        /// <param name="firstColumnIndex">Index of the first column in the range in which to automatically set the column width to the width of the cell contents.</param>
        /// <param name="lastColumnIndex">Index of the last column in the range in which to automatically set the column width to the width of the cell contents.</param>
        public void SetAutoWidth(int sheetIndex, int firstColumnIndex, int lastColumnIndex);

        /// <summary>
        /// Automatically sets the columns width to the width cell contenst from firtstColumnIndex to lastColumnIndex. The sheet is identified by its name.
        /// The range is identified by its firstColumn number and lastColumn number.
        /// </summary>
        /// <param name="sheetName">The name of the sheet in which the column width is to be automatically set to the width cell contents.</param>
        /// <param name="firstColumnIndex">Index of the first column in the range in which to automatically set the column width to the width of the cell contents.</param>
        /// <param name="lastColumnIndex">Index of the last column in the range in which to automatically set the column width to the width of the cell contents.</param>
        public void SetAutoWidth(string sheetName, int firstColumnIndex, int lastColumnIndex);

        /// <summary>
        /// Automatically sets column widths to the width of the cell contents for the all worksheet. The sheet is identified by its index.
        /// </summary>
        /// <param name="sheetIndex">The index of the sheet in which the column width is to be automatically set to the width cell contents.</param>
        public void SetAutoWidth(int sheetIndex);

        /// <summary>
        /// Automatically sets column widths to the width of the cell contents for the all worksheet. The sheet is identified by its name.
        /// </summary>
        /// <param name="sheetName">The name of the sheet in which the column width is to be automatically set to the width cell contents.</param>
        public void SetAutoWidth(string sheetName);

        /// <summary>
        /// Sets condition, comparison operator, bold, italic, underline, double underline, font color and background color for a cell.
        /// Sets conditional formatting for a cell. The sheet is identified by its index. The cell is identified by its row number and column number.
        /// Cell condition and styles are identified by ConditionAndFormationg.
        /// If condition is string then string must be inside \".
        /// </summary>
        /// <param name="sheetIndex">The index of the sheet in which conditional formatting will be set.</param>
        /// <param name="rowIndex">The row number in which conditional formatting is to be set.</param>
        /// <param name="columnIndex">The column number on which conditional formatting is to be set.</param>
        /// <param name="conditionAndFormattings">Configures condition, comparison operator, font and cell style using API.</param>
        /// <example>
        /// Here is an example of how to use the method:
        /// <code>
        /// excelOperationSyncfusion.ConditionalFormatting(0, 0, 0, new ConditionAndFormatting[] 
        /// { 
        ///     new ConditionAndFormatting(ComparisonOperatorIndex.Equal, "0")
        ///     .SetBackgroundColor(255, 255, 0, 0)
        ///     .SetBold(true), 
        ///     new ConditionAndFormatting(ComparisonOperatorIndex.GreaterThan, "3")
        ///     .SetBackgroundColor(255, 0, 255, 0).SetTextColor(255, 0, 0, 255)
        ///     .SetItalics(true) 
        /// });
        /// </code>
        /// </example>
        public void ConditionalFormatting(int sheetIndex, int rowIndex, int columnIndex, ConditionAndFormatting[] conditionAndFormattings);

        /// <summary>
        /// Sets condition, comparison operator, bold, italic, underline, double underline, font color and background color for a cell.
        /// Sets conditional formatting for a cell. The sheet is identified by its name. The range is identified by its firstRow number, 
        /// firstColumn number, lastRow number and lastColumn number. Cell condition and styles are identified by ConditionAndFormationg.
        /// If condition is string then string must be inside \".
        /// </summary>
        /// <param name="sheetIndex">The index of the sheet in which conditional formatting will be set.</param>
        /// <param name="firstRowIndex">The index of the first row in the range for which conditional formatting is to be set.</param>
        /// <param name="firstColumnIndex">The index of the first column in the range for which conditional formatting is to be set.</param>
        /// <param name="lastRowIndex">The index of the last row in the range for which conditional formatting is to be set.</param>
        /// <param name="lastColumnIndex">The index of the last column in the range for which conditional formatting is to be set.</param>
        /// <param name="conditionAndFormattings">Configures condition, comparison operator, font and cell style using API.</param>
        /// <example>
        /// Here is an example of how to use the method:
        /// <code>
        /// excelOperationSyncfusion.ConditionalFormatting(0, 0, 0, 6, 0, new ConditionAndFormatting[] 
        /// { 
        ///     new ConditionAndFormatting(ComparisonOperatorIndex.Equal, "0")
        ///     .SetBackgroundColor(255, 255, 0, 0)
        ///     .SetBold(true), 
        ///     new ConditionAndFormatting(ComparisonOperatorIndex.GreaterThan, "3")
        ///     .SetBackgroundColor(255, 0, 255, 0).SetTextColor(255, 0, 0, 255)
        ///     .SetItalics(true) 
        /// });
        /// </code>
        /// </example>
        public void ConditionalFormatting(int sheetIndex, int firstRowIndex, int firstColumnIndex, int lastRowIndex, int lastColumnIndex, ConditionAndFormatting[] conditionAndFormattings);

        /// <summary>
        /// Sets condition, comparison operator, bold, italic, underline, double underline, font color and background color for a cell.
        /// Sets conditional formatting for a cell. The sheet is identified by its name. The cell is identified by its row number and column number.
        /// Cell condition and styles are identified by ConditionAndFormationg.
        /// If condition is string then string must be inside \".
        /// </summary>
        /// <param name="sheetName">The name of the sheet in which conditional formatting will be set.</param>
        /// <param name="rowIndex">The row number in which conditional formatting is to be set.</param>
        /// <param name="columnIndex">The column number on which conditional formatting is to be set.</param>
        /// <param name="conditionAndFormatting">Configures condition, comparison operator, font and cell style using API.</param>
        /// <example>
        /// Here is an example of how to use the method:
        /// <code>
        /// excelOperationSyncfusion.ConditionalFormatting(0, 0, 0, new ConditionAndFormatting[] 
        /// { 
        ///     new ConditionAndFormatting(ComparisonOperatorIndex.Equal, "0")
        ///     .SetBackgroundColor(255, 255, 0, 0)
        ///     .SetBold(true), 
        ///     new ConditionAndFormatting(ComparisonOperatorIndex.GreaterThan, "3")
        ///     .SetBackgroundColor(255, 0, 255, 0).SetTextColor(255, 0, 0, 255)
        ///     .SetItalics(true) 
        /// });
        /// </code>
        /// </example>
        public void ConditionalFormatting(string sheetName, int rowIndex, int columnIndex, ConditionAndFormatting[] conditionAndFormatting);

        /// <summary>
        /// Sets condition, comparison operator, bold, italic, underline, double underline, font color and background color for a cell.
        /// Sets conditional formatting for a cell. The sheet is identified by its name. The range is identified by its firstRow number, 
        /// firstColumn number, lastRow number and lastColumn number. Cell condition and styles are identified by ConditionAndFormationg.
        /// If condition is string then string must be inside \".
        /// </summary>
        /// <param name="sheetName">The name of the sheet in which conditional formatting will be set.</param>
        /// <param name="firstRowIndex">The index of the first row in the range for which conditional formatting is to be set.</param>
        /// <param name="firstColumnIndex">The index of the first column in the range for which conditional formatting is to be set.</param>
        /// <param name="lastRowIndex">The index of the last row in the range for which conditional formatting is to be set.</param>
        /// <param name="lastColumnIndex">The index of the last column in the range for which conditional formatting is to be set.</param>
        /// <param name="conditionAndFormatting">Configures condition, comparison operator, font and cell style using API.</param>
        /// <example>
        /// Here is an example of how to use the method:
        /// <code>
        /// excelOperationSyncfusion.ConditionalFormatting(0, 0, 0, 6, 0, new ConditionAndFormatting[] 
        /// { 
        ///     new ConditionAndFormatting(ComparisonOperatorIndex.Equal, "0")
        ///     .SetBackgroundColor(255, 255, 0, 0)
        ///     .SetBold(true), 
        ///     new ConditionAndFormatting(ComparisonOperatorIndex.GreaterThan, "3")
        ///     .SetBackgroundColor(255, 0, 255, 0).SetTextColor(255, 0, 0, 255)
        ///     .SetItalics(true) 
        /// });
        /// </code>
        /// </example>
        public void ConditionalFormatting(string sheetName, int firstRowIndex, int firstColumnIndex, int lastRowIndex, int lastColumnIndex, ConditionAndFormatting[] conditionAndFormatting);

        /// <summary>
        /// Returns the numeric contents of a cell from a worksheet. The sheet is identified by its index. The cell is identified by its row number and column number.
        /// </summary>
        /// <param name="sheetIndex">The index of the sheet from which the number is to be taken.</param>
        /// <param name="rowIndex">The row number from which the number is to be taken.</param>
        /// <param name="columnIndex">The column number from which the number is to be taken.</param>
        /// <returns>Returns the numeric contents of a cell from a worksheet.</returns>
        public double GetCellValueNumber(int sheetIndex, int rowIndex, int columnIndex);

        /// <summary>
        /// Returns the numeric contents of a cell from a worksheet. The sheet is identified by its name. The cell is identified by its row number and column number.
        /// </summary>
        /// <param name="sheetName">The name of the sheet from which the number is to be taken.</param>
        /// <param name="rowIndex">The row number from which the number is to be taken.</param>
        /// <param name="columnIndex">The column number from which the number is to be taken.</param>
        /// <returns>Returns the numeric contents of a cell from a worksheet.</returns>
        public double GetCellValueNumber(string sheetName, int rowIndex, int columnIndex);

        /// <summary>
        /// Returns the text contents of a cell from a worksheet. The sheet is identified by its index. The cell is identified by its row number and column number.
        /// </summary>
        /// <param name="sheetIndex">The index of the sheet from which the text is to be taken.</param>
        /// <param name="rowIndex">The row number from which the text is to be taken.</param>
        /// <param name="columnIndex">The column number from which the text is to be taken.</param>
        /// <returns>Returns the text contents of a cell from a worksheet.</returns>
        public string GetCellValueText(int sheetIndex, int rowIndex, int columnIndex);

        /// <summary>
        /// Returns the text contents of a cell from a worksheet. The sheet is identified by its name. The cell is identified by its row number and column number.
        /// </summary>
        /// <param name="sheetName">The name of the sheet from which the text is to be taken.</param>
        /// <param name="rowIndex">The row number from which the text is to be taken.</param>
        /// <param name="columnIndex">The column number from which the text is to be taken.</param>
        /// <returns>Returns the text contents of a cell from a worksheet.</returns>
        public string GetCellValueText(string sheetName, int rowIndex, int columnIndex);

        /// <summary>
        /// Returns the date contents of a cell from a worksheet. The sheet is identified by its index. The cell is identified by its row number and column number.
        /// </summary>
        /// <param name="sheetIndex">The index of the sheet from which the date is to be taken.</param>
        /// <param name="rowIndex">The row number from which the date is to be taken.</param>
        /// <param name="columnIndex">The column number from which the date is to be taken.</param>
        /// <returns>Returns the date contents of a cell from a worksheet.</returns>
        public DateTime GetCellValueDate(int sheetIndex, int rowIndex, int columnIndex);

        /// <summary>
        /// Returns the date contents of a cell from a worksheet. The sheet is identified by its name. The cell is identified by its row number and column number.
        /// </summary>
        /// <param name="sheetName">The name of the sheet from which the date is to be taken.</param>
        /// <param name="rowIndex">The row number from which the date is to be taken.</param>
        /// <param name="columnIndex">The column number from which the date is to be taken.</param>
        /// <returns>Returns the date contents of a cell from a worksheet.</returns>
        public DateTime GetCellValueDate(string sheetName, int rowIndex, int columnIndex);

        /// <summary>
        /// Sets the author, subject, and title of the spreadsheet.
        /// </summary>
        /// <param name="author">Author's name.</param>
        /// <param name="subject">File subject.</param>
        /// <param name="title">File title.</param>
        public void MetaData(string author, string subject, string title);

        /// <summary>
        /// Returns the number of rows from a sheet. The sheet is identified by its index.
        /// </summary>
        /// <param name="sheetIndex">The index of the sheet from which the number of rows is to be taken.</param>
        /// <returns>Returns the number of rows from a sheet.</returns>
        public int GetLastRow(int sheetIndex);

        /// <summary>
        /// Returns the number of rows from a sheet. The sheet is identified by its name.
        /// </summary>
        /// <param name="sheetName">The name of the sheet from which the number of rows is to be taken.</param>
        /// <returns>Returns the number of rows from a sheet.</returns>
        public int GetLastRow(string sheetName);

        /// <summary>
        /// Returns the number of columns from a sheet. The sheet is identified by its index.
        /// </summary>
        /// <param name="sheetIndex">The index of the sheet from which the number of columns is to be taken.</param>
        /// <returns>Returns the number of columns from a sheet.</returns>
        public int GetLastColumn(int sheetIndex);

        /// <summary>
        /// Returns the number of columns from a sheet. The sheet is identified by its name.
        /// </summary>
        /// <param name="sheetName">The name of the sheet from which the number of columns is to be taken.</param>
        /// <returns>Returns the number of columns from a sheet.</returns>
        public int GetLastColumn(string sheetName);

        /// <summary>
        /// Returns the object. Opens a file. The file is identified by its path.
        /// </summary>
        /// <param name="path">File path.</param>
        /// <returns>Returns the object.</returns>
        public object OpenSpreadsheet(FileStream path);

        /// <summary>
        /// Sets the sheet as hidden. The sheet is identified by its index. The sheet visibility is identified by SheetVisibilityIndex.
        /// </summary>
        /// <param name="sheetIndex">The index of the sheet to be hidden.</param>
        /// <param name="sheetVisibilityIndex">Visibility to be set.</param>
        public void HideSheet(int sheetIndex, SheetVisibilityIndex sheetVisibilityIndex);

        /// <summary>
        /// Sets the sheet as hidden. The sheet is identified by its name. The sheet visibility is identified by SheetVisibilityIndex.
        /// </summary>
        /// <param name="sheetName">The name of the sheet to be hidden.</param>
        /// <param name="sheetVisibilityIndex">Visibility to be set.</param>
        public void HideSheet(string sheetName, SheetVisibilityIndex sheetVisibilityIndex);

        /// <summary>
        /// Sets which sheet will be active by default when a file is opened.
        /// </summary>
        /// <param name="sheetIndex">The index of the sheet to be set as the default sheet when a file is opened.</param>
        public void ActiveSheet(int sheetIndex);

        /// <summary>
        /// Sets which sheet will be active by default when a file is opened.
        /// </summary>
        /// <param name="sheetName">The name of the sheet to be set as the default sheet when a file is opened.</param>
        public void ActiveSheet(string sheetName);

        /// <summary>
        /// Sets one sheet as hidden and the other as active (default). The hidden sheet is identified by index.
        /// The active sheet is identified by index. The sheet visibility is identified by SheetVisibilityIndex.
        /// </summary>
        /// <param name="hideSheetIndex">The index of the sheet to be hidden.</param>
        /// <param name="activeSheetIndex">The index of the sheet to be set as the default sheet when a file is opened.</param>
        /// <param name="sheetVisibilityIndex">Visibility to be set.</param>
        public void HideSheet(int hideSheetIndex, int activeSheetIndex, SheetVisibilityIndex sheetVisibilityIndex);

        /// <summary>
        /// Sets one sheet as hidden and the other as active (default). The hidden sheet is identified by name.
        /// The active sheet is identified by name. The sheet visibility is identified by SheetVisibilityIndex.
        /// </summary>
        /// <param name="hideSheetName">The name of the sheet to be hidden.</param>
        /// <param name="activeSheetName">The name of the sheet to be set as the default sheet when a file is opened.</param>
        /// <param name="sheetVisibilityIndex">Visibility to be set.</param>
        public void HideSheet(string hideSheetName, string activeSheetName, SheetVisibilityIndex sheetVisibilityIndex);

        /// <summary>
        /// Sets one sheet as hidden and the other as active (default). The hidden sheet is identified by index.
        /// The active sheet is identified by name. The sheet visibility is identified by SheetVisibilityIndex.
        /// </summary>
        /// <param name="hideSheetIndex">The index of the sheet to be hidden.</param>
        /// <param name="activeSheetName">The name of the sheet to be set as the default sheet when a file is opened.</param>
        /// <param name="sheetVisibilityIndex">Visibility to be set.</param>
        public void HideSheet(int hideSheetIndex, string activeSheetName, SheetVisibilityIndex sheetVisibilityIndex);

        /// <summary>
        /// Sets one sheet as hidden and the other as active (default). The hidden sheet is identified by name.
        /// The active sheet is identified by index. The sheet visibility is identified by SheetVisibilityIndex.
        /// </summary>
        /// <param name="hideSheetName">The name of the sheet to be hidden.</param>
        /// <param name="activeSheetIndex">The index of the sheet to be set as the default sheet when a file is opened.</param>
        /// <param name="sheetVisibilityIndex">Visibility to be set.</param>
        public void HideSheet(string hideSheetName, int activeSheetIndex, SheetVisibilityIndex sheetVisibilityIndex);

        /// <summary>
        /// Sets the row as hidden. The sheet is identified by its index. The row is identified by index.
        /// </summary>
        /// <param name="sheetIndex">The index of the sheet in which the row is to be hidden.</param>
        /// <param name="rowIndex">The index of the row to be hidden.</param>
        public void HideRow(int sheetIndex, int rowIndex);

        /// <summary>
        /// Sets the row as hidden. The sheet is identified by its name. The row is identified by index.
        /// </summary>
        /// <param name="sheetName">The name of the sheet in which the row is to be hidden.</param>
        /// <param name="rowIndex">The index of the row to be hidden.</param>
        public void HideRow(string sheetName, int rowIndex);

        /// <summary>
        /// Sets the rows as hidden. The sheet is identified by its index. 
        /// The range is identified by its firstRow number and lastRow number.
        /// </summary>
        /// <param name="sheetIndex">The index of the sheet in which the rows is to be hidden.</param>
        /// <param name="firstRowIndex">The index of the first row in the range to be hidden.</param>
        /// <param name="lastRowIndex">The index of the last row in the range to be hidden.</param>
        public void HideRow(int sheetIndex, int firstRowIndex, int lastRowIndex);

        /// <summary>
        /// Sets the rows as hidden. The sheet is identified by its name. 
        /// The range is identified by its firstRow number and lastRow number.
        /// </summary>
        /// <param name="sheetName">The name of the sheet in which the rows is to be hidden.</param>
        /// <param name="firstRowIndex">The index of the first row in the range to be hidden.</param>
        /// <param name="lastRowIndex">The index of the last row in the range to be hidden.</param>
        public void HideRow(string sheetName, int firstRowIndex, int lastRowIndex);

        /// <summary>
        /// Sets the rows as hidden. The sheet is identified by its index.
        /// Rows are identified by row number from table rowIndexes.
        /// </summary>
        /// <param name="sheetIndex">The index of the sheet in which the rows is to be hidden.</param>
        /// <param name="rowIndexes">Array of rows to be hidden.</param>
        public void HideRow(int sheetIndex, int[] rowIndexes);

        /// <summary>
        /// Sets the rows as hidden. The sheet is identified by its name.
        /// Rows are identified by row number from table rowIndexes.
        /// </summary>
        /// <param name="sheetName">The name of the sheet in which the rows is to be hidden.</param>
        /// <param name="rowIndexes">Array of rows to be hidden.</param>
        public void HideRow(string sheetName, int[] rowIndexes);

        /// <summary>
        /// Sets the column as hidden. The sheet is identified by its index. The column is identified by index.
        /// </summary>
        /// <param name="sheetIndex">The index of the sheet in which the column is to be hidden.</param>
        /// <param name="columnIndex">The index of the column to be hidden.</param>
        public void HideColumn(int sheetIndex, int columnIndex);

        /// <summary>
        /// Sets the column as hidden. The sheet is identified by its name. The column is identified by index.
        /// </summary>
        /// <param name="sheetName">The name of the sheet in which the column is to be hidden.</param>
        /// <param name="columnIndex">The index of the column to be hidden.</param>
        public void HideColumn(string sheetName, int columnIndex);

        /// <summary>
        /// Sets the columns as hidden. The sheet is identified by its index. 
        /// The range is identified by its firstColumn number and lastColumn number.
        /// </summary>
        /// <param name="sheetIndex">The index of the sheet in which the columns is to be hidden.</param>
        /// <param name="firstColumnIndex">The index of the first column in the range to be hidden.</param>
        /// <param name="lastColumnIndex">The index of the last column in the range to be hidden.</param>
        public void HideColumn(int sheetIndex, int firstColumnIndex, int lastColumnIndex);

        /// <summary>
        /// Sets the columns as hidden. The sheet is identified by its name. 
        /// The range is identified by its firstColumn number and lastColumn number.
        /// </summary>
        /// <param name="sheetName">The name of the sheet in which the columns is to be hidden.</param>
        /// <param name="firstColumnIndex">The index of the first column in the range to be hidden.</param>
        /// <param name="lastColumnIndex">The index of the last column in the range to be hidden.</param>
        public void HideColumn(string sheetName, int firstColumnIndex, int lastColumnIndex);

        /// <summary>
        /// Sets the columns as hidden. The sheet is identified by its index.
        /// Columns are identified by column number from table columnIndexes.
        /// </summary>
        /// <param name="sheetIndex">The index of the sheet in which the columns is to be hidden.</param>
        /// <param name="columnIndexes">Array of columns to be hidden.</param>
        public void HideColumn(int sheetIndex, int[] columnIndexes);

        /// <summary>
        /// Sets the columns as hidden. The sheet is identified by its name.
        /// Columns are identified by column number from table columnIndexes.
        /// </summary>
        /// <param name="sheetName">The name of the sheet in which the columns is to be hidden.</param>
        /// <param name="columnIndexes">Array of columns to be hidden.</param>
        public void HideColumn(string sheetName, int[] columnIndexes);

        /// <summary>
        /// Sets the rows and columns as hidden. The sheet is identified by its index.
        /// The range is identified by its firstRow number, firstColumn number, lastRow number and lastColumn number.
        /// </summary>
        /// <param name="sheetIndex">The index of the sheet in which the rows and columns is to be hidden.</param>
        /// <param name="firstRowIndex">The index of the first row in the range to be hidden.</param>
        /// <param name="firstColumnIndex">The index of the first column in the range to be hidden.</param>
        /// <param name="lastRowIndex">The index of the last row in the range to be hidden.</param>
        /// <param name="lastColumnIndex">The index of the last column in the range to be hidden.</param>
        public void HideRowAndColumn(int sheetIndex, int firstRowIndex, int firstColumnIndex, int lastRowIndex, int lastColumnIndex);

        /// <summary>
        /// Sets the rows and columns as hidden. The sheet is identified by its name.
        /// The range is identified by its firstRow number, firstColumn number, lastRow number and lastColumn number.
        /// </summary>
        /// <param name="sheetName">The name of the sheet in which the rows and columns is to be hidden.</param>
        /// <param name="firstRowIndex">The index of the first row in the range to be hidden.</param>
        /// <param name="firstColumnIndex">The index of the first column in the range to be hidden.</param>
        /// <param name="lastRowIndex">The index of the last row in the range to be hidden.</param>
        /// <param name="lastColumnIndex">The index of the last column in the range to be hidden.</param>
        public void HideRowAndColumn(string sheetName, int firstRowIndex, int firstColumnIndex, int lastRowIndex, int lastColumnIndex);

        /// <summary>
        /// Sets a name for the range. The sheet is identified by its index.
        /// </summary>
        /// <param name="sheetIndex">The index of the sheet containing the data for the range.</param>
        /// <param name="name">Range name.</param>
        /// <param name="range">Data range.</param>
        public void NameManager(int sheetIndex, string name, string range);

        /// <summary>
        /// Sets a name for the range. The sheet is identified by its index.
        /// </summary>
        /// <param name="sheetName">The name of the sheet containing the data for the range.</param>
        /// <param name="name">Range name.</param>
        /// <param name="range">Data range.</param>
        public void NameManager(string sheetName, string name, string range);

        /// <summary>
        /// Sets a name for the range. The sheet is identified by its index.
        /// </summary>
        /// <param name="sheetIndex">The index of the sheet containing the data for the range.</param>
        /// <param name="name">Range name.</param>
        /// <param name="range">Data range.</param>
        /// <param name="comment">Comment to the range.</param>
        public void NameManager(int sheetIndex, string name, string range, string comment);

        /// <summary>
        /// Sets a name for the range. The sheet is identified by its index.
        /// </summary>
        /// <param name="sheetName">The name of the sheet containing the data for the range.</param>
        /// <param name="name">Range name.</param>
        /// <param name="range">Data range.</param>
        /// <param name="comment">Comment to the range.</param>
        public void NameManager(string sheetName, string name, string range, string comment);

        /// <summary>
        /// Sets whether the cell to be protected. The sheet is identified by its index. The cell is identified by its row and column number.
        /// </summary>
        /// <param name="sheetIndex">The index of the sheet in which the cell is to be protected.</param>
        /// <param name="rowIndex">The index of the row in which the cell is to be protected.</param>
        /// <param name="columnIndex">The index of the row column in which the cell is to be protected.</param>
        /// <param name="block">Blocks or unblocks the cell. If block = true the cell is protected if block = false the cell is not protected.</param>
        public void ProtectCell(int sheetIndex, int rowIndex, int columnIndex, bool block);

        /// <summary>
        /// Sets whether the cell to be protected. The sheet is identified by its name. The cell is identified by its row and column number.
        /// </summary>
        /// <param name="sheetName">The name of the sheet in which the cell is to be protected.</param>
        /// <param name="rowIndex">The index of the row in which the cell is to be protected.</param>
        /// <param name="columnIndex">The index of the row column in which the cell is to be protected.</param>
        /// <param name="block">Blocks or unblocks the cell. If block = true the cell is protected if block = false the cell is not protected.</param>
        public void ProtectCell(string sheetName, int rowIndex, int columnIndex, bool block);

        /// <summary>
        /// Sets whether the cells to be protected. The sheet is identified by its index. The range is identified by its firstRow number, 
        /// firstColumn number, lastRow number and lastColumn number.
        /// </summary>
        /// <param name="sheetIndex">The index of the sheet in which the cells to be protected.</param>
        /// <param name="firstRowIndex">The index of the first row in the range to be protected.</param>
        /// <param name="firstColumnIndex">The index of the first column in the range to be protected.</param>
        /// <param name="lastRowIndex">The index of the last row in the range to be protected.</param>
        /// <param name="lastColumnIndex">The index of the last column in the range to be protected.</param>
        /// <param name="block">Blocks or unblocks the cells. If block = true the cells are protected if block = false the cells are not protected.</param>
        public void ProtectCell(int sheetIndex, int firstRowIndex, int firstColumnIndex, int lastRowIndex, int lastColumnIndex, bool block);

        /// <summary>
        /// Sets whether the cells to be protected. The sheet is identified by its name. The range is identified by its firstRow number, 
        /// firstColumn number, lastRow number and lastColumn number.
        /// </summary>
        /// <param name="sheetName">The name of the sheet in which the cells to be protected.</param>
        /// <param name="firstRowIndex">The index of the first row in the range to be protected.</param>
        /// <param name="firstColumnIndex">The index of the first column in the range to be protected.</param>
        /// <param name="lastRowIndex">The index of the last row in the range to be protected.</param>
        /// <param name="lastColumnIndex">The index of the last column in the range to be protected.</param>
        /// <param name="block">Blocks or unblocks the cells. If block = true the cells are protected if block = false the cells are not protected.</param>
        public void ProtectCell(string sheetName, int firstRowIndex, int firstColumnIndex, int lastRowIndex, int lastColumnIndex, bool block);

        /// <summary>
        /// Sets whether the row to be protected. The sheet is identified by its index. The row is identified by its row number.
        /// </summary>
        /// <param name="sheetIndex">The index of the sheet in which the row is to be protected.</param>
        /// <param name="rowIndex">The index of the row in which the row is to be protected.</param>
        /// <param name="block">Blocks or unblocks the row. If block = true the row is protected if block = false the row is not protected.</param>
        public void ProtectRow(int sheetIndex, int rowIndex, bool block);

        /// <summary>
        /// Sets whether the row to be protected. The sheet is identified by its name. The row is identified by its row number.
        /// </summary>
        /// <param name="sheetName">The name of the sheet in which the row is to be protected.</param>
        /// <param name="rowIndex">The index of the row in which the row is to be protected.</param>
        /// <param name="block">Blocks or unblocks the row. If block = true the row is protected if block = false the row is not protected.</param>
        public void ProtectRow(string sheetName, int rowIndex, bool block);

        /// <summary>
        /// Sets whether the rows to be protected. The sheet is identified by its index. The range is identified by its firstRow number and lastRow number.
        /// </summary>
        /// <param name="sheetIndex">The index of the sheet in which the rows to be protected.</param>
        /// <param name="firstRowIndex">The index of the first row in the range to be protected.</param>
        /// <param name="lastRowIndex">The index of the last row in the range to be protected.</param>
        /// <param name="block">Blocks or unblocks the rows. If block = true the rows are protected if block = false the rows are not protected.</param>
        public void ProtectRow(int sheetIndex, int firstRowIndex, int lastRowIndex, bool block);

        /// <summary>
        /// Sets whether the rows to be protected. The sheet is identified by its name. The range is identified by its firstRow number and lastRow number.
        /// </summary>
        /// <param name="sheetName">The name of the sheet in which the rows to be protected.</param>
        /// <param name="firstRowIndex">The index of the first row in the range to be protected.</param>
        /// <param name="lastRowIndex">The index of the last row in the range to be protected.</param>
        /// <param name="block">Blocks or unblocks the rows. If block = true the rows are protected if block = false the rows are not protected.</param>
        public void ProtectRow(string sheetName, int firstRowIndex, int lastRowIndex, bool block);

        /// <summary>
        /// Sets whether the rows to be protected. The sheet is identified by its index. Rows are identified by indexes from the rowIndexes array.
        /// </summary>
        /// <param name="sheetIndex">The index of the sheet in which the rows to be protected.</param>
        /// <param name="rowIndexes">The array of rows to be protected.</param>
        /// <param name="block">Blocks or unblocks the rows. If block = true the rows are protected if block = false the rows are not protected.</param>
        public void ProtectRow(int sheetIndex, int[] rowIndexes, bool block);

        /// <summary>
        /// Sets whether the rows to be protected. The sheet is identified by its name. Rows are identified by indexes from the rowIndexes array.
        /// </summary>
        /// <param name="sheetName">The name of the sheet in which the rows to be protected.</param>
        /// <param name="rowIndexes">The array of rows to be protected.</param>
        /// <param name="block">Blocks or unblocks the rows. If block = true the rows are protected if block = false the rows are not protected.</param>
        public void ProtectRow(string sheetName, int[] rowIndexes, bool block);

        /// <summary>
        /// Sets whether the column to be protected. The sheet is identified by its index. The column is identified by its column number.
        /// </summary>
        /// <param name="sheetIndex">The index of the sheet in which the column is to be protected.</param>
        /// <param name="columnIndex">The index of the column in which the column is to be protected.</param>
        /// <param name="block">Blocks or unblocks the column. If block = true the column is protected if block = false the column is not protected.</param>
        public void ProtectColumn(int sheetIndex, int columnIndex, bool block);

        /// <summary>
        /// Sets whether the column to be protected. The sheet is identified by its name. The column is identified by its column number.
        /// </summary>
        /// <param name="sheetName">The name of the sheet in which the column is to be protected.</param>
        /// <param name="columnIndex">The index of the column in which the column is to be protected.</param>
        /// <param name="block">Blocks or unblocks the column. If block = true the column is protected if block = false the column is not protected.</param>
        public void ProtectColumn(string sheetName, int columnIndex, bool block);

        /// <summary>
        /// Sets whether the columns to be protected. The sheet is identified by its index. The range is identified by its firstColumn number and lastColumn number.
        /// </summary>
        /// <param name="sheetIndex">The index of the sheet in which the columns to be protected.</param>
        /// <param name="firstColumnIndex">The index of the first column in the range to be protected.</param>
        /// <param name="lastColumnIndex">The index of the last column in the range to be protected.</param>
        /// <param name="block">Blocks or unblocks the columns. If block = true the columns are protected if block = false the columns are not protected.</param>
        public void ProtectColumn(int sheetIndex, int firstColumnIndex, int lastColumnIndex, bool block);

        /// <summary>
        /// Sets whether the columns to be protected. The sheet is identified by its name. The range is identified by its firstColumn number and lastColumn number.
        /// </summary>
        /// <param name="sheetName">The name of the sheet in which the columns to be protected.</param>
        /// <param name="firstColumnIndex">The index of the first column in the range to be protected.</param>
        /// <param name="lastColumnIndex">The index of the last column in the range to be protected.</param>
        /// <param name="block">Blocks or unblocks the columns. If block = true the columns are protected if block = false the columns are not protected.</param>
        public void ProtectColumn(string sheetName, int firstColumnIndex, int lastColumnIndex, bool block);

        /// <summary>
        /// Sets whether the columns to be protected. The sheet is identified by its index. Columns are identified by indexes from the columnIndexes array.
        /// </summary>
        /// <param name="sheetIndex">The index of the sheet in which the columns to be protected.</param>
        /// <param name="columnIndexes">The array of columns to be protected.</param>
        /// <param name="block">Blocks or unblocks the columns. If block = true the columns are protected if block = false the columns are not protected.</param>
        public void ProtectColumn(int sheetIndex, int[] columnIndexes, bool block);

        /// <summary>
        /// Sets whether the columns to be protected. The sheet is identified by its name. Columns are identified by indexes from the columnIndexes array.
        /// </summary>
        /// <param name="sheetName">The name of the sheet in which the columns to be protected.</param>
        /// <param name="columnIndexes">The array of columns to be protected.</param>
        /// <param name="block">Blocks or unblocks the columns. If block = true the columns are protected if block = false the columns are not protected.</param>
        public void ProtectColumn(string sheetName, int[] columnIndexes, bool block);

        /// <summary>
        /// Sets the cell format. The sheet is identified by its index. The cell is identified by its row and column number.
        /// </summary>
        /// <param name="sheetIndex">The index of the sheet in which the cell is to have the format set.</param>
        /// <param name="rowIndex">The index of the row at which the cell is to have the format set.</param>
        /// <param name="columnIndex">The index of the column at which the cell is to have the format set.</param>
        /// <param name="format">Cell format.</param>
        public void SetCellType(int sheetIndex, int rowIndex, int columnIndex, string format);

        /// <summary>
        /// Sets the cell format. The sheet is identified by its name. The cell is identified by its row and column number.
        /// </summary>
        /// <param name="sheetName">The name of the sheet in which the cell is to have the format set.</param>
        /// <param name="rowIndex">The index of the row at which the cell is to have the format set.</param>
        /// <param name="columnIndex">The index of the column at which the cell is to have the format set.</param>
        /// <param name="format">Cell format.</param>
        public void SetCellType(string sheetName, int rowIndex, int columnIndex, string format);

        /// <summary>
        /// Returns the index of the last row in column. The sheet is identified by its index. The column is identified by its column number.
        /// </summary>
        /// <param name="sheetIndex">The index of the sheet from which the last row of the column is to be taken.</param>
        /// <param name="columnIndex">The index of the column from which the last row is to be taken.</param>
        /// <returns>Returns the index of the last row in column.</returns>
        public int GetLastRowIndexInColumn(int sheetIndex, int columnIndex);

        /// <summary>
        /// Returns the index of the last row in column. The sheet is identified by its name. The column is identified by its column number.
        /// </summary>
        /// <param name="sheetName">The name of the sheet from which the last row of the column is to be taken.</param>
        /// <param name="columnIndex">The index of the column from which the last row is to be taken.</param>
        /// <returns>Returns the index of the last row in column.</returns>
        public int GetLastRowIndexInColumn(string sheetName, int columnIndex);

        /// <summary>
        /// Returns the index of the last column in row. The sheet is identified by its index. The row is identified by its row number.
        /// </summary>
        /// <param name="sheetIndex">The index of the sheet from which the last column of the row is to be taken.</param>
        /// <param name="rowIndex">The index of the row form which the last colum is to be taken.</param>
        /// <returns>Returns the index of the last column in row.</returns>
        public int GetLastColumnIndexInRow(int sheetIndex, int rowIndex);

        /// <summary>
        /// Returns the index of the last column in row. The sheet is identified by its name. The row is identified by its row number.
        /// </summary>
        /// <param name="sheetName">The name of the sheet from which the last column of the row is to be taken.</param>
        /// <param name="rowIndex">The index of the row form which the last colum is to be taken.</param>
        /// <returns>Returns the index of the last column in row.</returns>
        public int GetLastColumnIndexInRow(string sheetName, int rowIndex);
        //COMING SOON NEXT UPDATE
    }
}
