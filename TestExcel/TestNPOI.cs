﻿using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using Org.BouncyCastle.Crypto;
using Soneta.Business;
using Soneta.Core;
using Soneta.Handel;
using Soneta.Types;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TestExcel;
using Trustsoft.ExcelOperation.Moje;

[assembly: Worker(typeof(TestNPOI), typeof(DokEwidencja))]
namespace TestExcel
{
    public class TestNPOI
    {
        [Context]
        public Session Session { get; set; }

        private readonly FromTo fd = new FromTo();

        [Action("Test NPOI", Icon = ActionIcon.Test, Mode = ActionMode.Progress | ActionMode.SingleSession, Target = ActionTarget.Menu | ActionTarget.LocalMenu | ActionTarget.Divider | ActionTarget.ToolbarWithText)]
        public void Eksport()
        {
            ExcelOperationNPOI excelOperationNPOI = new ExcelOperationNPOI();
            var workbook = excelOperationNPOI.CreateWorkbook();
            IWorkbook workbook1 = workbook as IWorkbook;
            int index = excelOperationNPOI.AddWorksheet("Test1");
            int worksheet1 = excelOperationNPOI.AddWorksheet("Test2");
            int worksheetPusty = excelOperationNPOI.AddWorksheet("Pusty");
            List<Sheet> sheets = excelOperationNPOI.GetNameSheet();
            int i = 0;
            foreach (var item in sheets)
            {
                excelOperationNPOI.AddCellValueInt(index, i, 1, item.Index);
                excelOperationNPOI.AddCellValueText(index, i, 2, item.Name);
                i++;
            }
            Time time = new Time(12, 15);
            excelOperationNPOI.AddCellValueTime(0, 0, 1, time);
            excelOperationNPOI.GetCellValueText(0, 0, 1);
            excelOperationNPOI.MetaData("Jan", "Export", "Nowy");
            excelOperationNPOI.AddCellValuePercent(0, 1, 3, new Percent(0.8m));
            excelOperationNPOI.AddCellValueDate(0, 2, 3, new DateTime(2024, 8, 21));
            double procent = excelOperationNPOI.GetCellValueNumber(0, 1, 3);
            DateTime data = excelOperationNPOI.GetCellValueDate(0, 2, 3);
            /*excelOperationNPOI.AddCellValueInt(0, 0, 0, 0);
            excelOperationNPOI.AddCellValueInt(0, 1, 0, 1);
            excelOperationNPOI.AddCellValueInt(0, 2, 0, 2);
            excelOperationNPOI.AddCellValueInt(0, 3, 0, 3);
            excelOperationNPOI.AddCellValueInt(0, 4, 0, 4);
            excelOperationNPOI.AddCellValueInt(0, 5, 0, 5);
            excelOperationNPOI.AddCellValueInt(0, 6, 0, 6);*/
            excelOperationNPOI.AddCellValueText(0, 0, 0, "OK");
            excelOperationNPOI.AddCellValueText(0, 1, 0, "ERROR");
            excelOperationNPOI.AddCellValueText(0, 2, 0, "OK");
            excelOperationNPOI.AddCellValueInt(0, 3, 0, 1);
            excelOperationNPOI.AddCellValueInt(0, 4, 0, 1);
            excelOperationNPOI.AddCellValueInt(0, 5, 0, 3);
            excelOperationNPOI.AddCellFormula(0, 6, 0, "SUM(A4:A6)");
            excelOperationNPOI.ConditionalFormatting(0, 0, 0, 6, 0, new ConditionAndFormatting[]
            {
                    //new ConditionAndFormatting(ComparisonOperatorIndex.Equal, "OK").SetBackgroundColor(255, 198, 239, 206).SetTextColor(255, 0, 100, 0),
                    new ConditionAndFormatting(ComparisonOperatorIndex.Equal, "\"OK\"").SetBackgroundColor(255, 198, 239, 206).SetTextColor(255, 0, 100, 0),
                    //new ConditionAndFormatting(ComparisonOperatorIndex.Equal, "ERROR").SetBackgroundColor(255, 255, 199, 206).SetTextColor(255, 190, 1, 7),
                    new ConditionAndFormatting(ComparisonOperatorIndex.Equal, "\"ERROR\"").SetBackgroundColor(255, 255, 199, 206).SetTextColor(255, 190, 1, 7),
                    new ConditionAndFormatting(ComparisonOperatorIndex.Equal, "3").SetBackgroundColor(255,0,0,255)
            });

            /*excelOperationNPOI.AddCellValueText(0, 0, 1, "Bardzo długi test specjalinie do testów1");
            excelOperationNPOI.AddCellValueText(0, 1, 1, "Bardzo długi test specjalinie do testów11 hahahaha");
            excelOperationNPOI.AddCellValueText(0, 0, 2, "Bardzo długi test specjalinie do testów2");
            excelOperationNPOI.AddCellValueText(0, 0, 3, "Bardzo długi test specjalinie do testów3");
            excelOperationNPOI.SetAutoWidth(0,1,2);*/

            int IndexDane = excelOperationNPOI.AddWorksheet("Data");
            excelOperationNPOI.AddCellValueText(IndexDane, 0, 0, "Kod");
            excelOperationNPOI.AddCellValueText(IndexDane, 0, 1, "Imię i Nazwisko");
            excelOperationNPOI.AddCellValueText(IndexDane, 1, 0, "Code");
            excelOperationNPOI.AddCellValueText(IndexDane, 1, 1, "Name and Surname");

            excelOperationNPOI.AddCellValueText(IndexDane, 2, 0, "00002");
            excelOperationNPOI.AddCellValueText(IndexDane, 2, 1, "Jan Nowak_1");
            excelOperationNPOI.AddCellValueText(IndexDane, 3, 0, "723908");
            excelOperationNPOI.AddCellValueText(IndexDane, 3, 1, "Jan Nowak_2");
            excelOperationNPOI.AddCellValueText(IndexDane, 4, 0, "723914");
            excelOperationNPOI.AddCellValueText(IndexDane, 4, 1, "Jan Nowak_3");
            excelOperationNPOI.AddCellValueText(IndexDane, 5, 0, "723990");
            excelOperationNPOI.AddCellValueText(IndexDane, 5, 1, "Jan Nowak_4");
            excelOperationNPOI.AddCellValueText(IndexDane, 6, 0, "725578");
            excelOperationNPOI.AddCellValueText(IndexDane, 6, 1, "Jan Nowak_5");
            excelOperationNPOI.AddCellValueText(IndexDane, 7, 0, "725585");
            excelOperationNPOI.AddCellValueText(IndexDane, 7, 1, "Jan Nowak_6");
            excelOperationNPOI.AddCellValueText(IndexDane, 8, 0, "726479");
            excelOperationNPOI.AddCellValueText(IndexDane, 8, 1, "Jan Nowak_7");
            excelOperationNPOI.AddCellValueText(IndexDane, 9, 0, "726827");
            excelOperationNPOI.AddCellValueText(IndexDane, 9, 1, "Jan Nowak_8");
            excelOperationNPOI.AddCellValueText(IndexDane, 10, 0, "727019");
            excelOperationNPOI.AddCellValueText(IndexDane, 10, 1, "Jan Nowak_9");
            excelOperationNPOI.AddCellValueText(IndexDane, 11, 0, "727815");
            excelOperationNPOI.AddCellValueText(IndexDane, 11, 1, "Jan Nowak_10");
            excelOperationNPOI.AddCellValueText(IndexDane, 12, 0, "727816");
            excelOperationNPOI.AddCellValueText(IndexDane, 12, 1, "Jan Nowak_11");
            excelOperationNPOI.AddCellValueText(IndexDane, 13, 0, "727936");
            excelOperationNPOI.AddCellValueText(IndexDane, 13, 1, "Jan Nowak_12");
            excelOperationNPOI.AddCellValueText(IndexDane, 14, 0, "727937");
            excelOperationNPOI.AddCellValueText(IndexDane, 14, 1, "Jan Nowak_13");

            excelOperationNPOI.NameManager("Data", "pawel", "A3:A15");

            //excelOperationNPOI.DropDownList(0, IndexDane, 0, 0, 20, 0, "$A$3:$A$15");
            excelOperationNPOI.AddCellValueText(index, 0, 0, "Tekst11");
            excelOperationNPOI.AddCellValueText(index, 0, 1, "Tekst12");
            excelOperationNPOI.AddCellValueText(index, 0, 2, "Tekst13");
            excelOperationNPOI.AddCellValueText(index, 1, 0, "Tekst21");
            excelOperationNPOI.AddCellValueText(index, 1, 1, "Tekst22");
            excelOperationNPOI.AddCellValueText(index, 1, 2, "Tekst23");
            excelOperationNPOI.AddCellValueText(index, 2, 0, "Tekst31");
            excelOperationNPOI.AddCellValueText(index, 2, 1, "Tekst32");
            excelOperationNPOI.AddCellValueText(index, 2, 2, "Tekst33");

            excelOperationNPOI.AddCellFormula(index, 10, 1, "SUM(A6:A10)");

            var lastRow = excelOperationNPOI.GetLastRow("Test1");
            var lastColumn = excelOperationNPOI.GetLastColumn("Test1");
            excelOperationNPOI.HideSheet(3, SheetVisibilityIndex.Hidden);
            excelOperationNPOI.ActiveSheet(1);
            //excelOperationNPOI.HideRow("Test1", 0, 1);


            int indexManager = excelOperationNPOI.AddWorksheet("Manager");
            int indexNewManager = excelOperationNPOI.AddWorksheet("NewManager");
            excelOperationNPOI.AddCellValueText(indexManager, 0, 0, "Nazwisko");
            excelOperationNPOI.AddCellValueText(indexManager, 0, 1, "Imie");
            excelOperationNPOI.AddCellValueText(indexManager, 1, 0, "Kowalski");
            excelOperationNPOI.AddCellValueText(indexManager, 1, 1, "Jan");
            excelOperationNPOI.AddCellValueText(indexManager, 2, 0, "Kowalski");
            excelOperationNPOI.AddCellValueText(indexManager, 2, 1, "Anna");
            excelOperationNPOI.AddCellValueText(indexManager, 3, 0, "Nowak");
            excelOperationNPOI.AddCellValueText(indexManager, 3, 1, "Piotr");
            excelOperationNPOI.AddCellValueText(indexManager, 4, 0, "Nowal");
            excelOperationNPOI.AddCellValueText(indexManager, 4, 1, "Ewa");
            excelOperationNPOI.AddCellValueText(indexManager, 5, 0, "Wiśniewski");
            excelOperationNPOI.AddCellValueText(indexManager, 5, 1, "Marek");

            excelOperationNPOI.AddCellValueText(indexNewManager, 0, 4, "X");
            excelOperationNPOI.AddCellValueText(indexNewManager, 0, 5, "xx1");
            excelOperationNPOI.AddCellValueText(indexNewManager, 0, 6, "yy1");
            excelOperationNPOI.AddCellValueText(indexNewManager, 0, 7, "zz1");
            excelOperationNPOI.AddCellValueText(indexNewManager, 1, 4, "Y");
            excelOperationNPOI.AddCellValueText(indexNewManager, 1, 5, "xx2");
            excelOperationNPOI.AddCellValueText(indexNewManager, 1, 6, "yy2");
            excelOperationNPOI.AddCellValueText(indexNewManager, 1, 7, "zz2");
            excelOperationNPOI.AddCellValueText(indexNewManager, 2, 4, "Z");

            /*excelOperationNPOI.NameManager(indexNewManager, "XYZ", $"NewManager!$E$1:$E$3");
            excelOperationNPOI.NameManager(indexNewManager, "X", $"NewManager!$F$1:$F$2");
            excelOperationNPOI.NameManager(indexNewManager, "Y", $"NewManager!$G$1:$G$2");
            excelOperationNPOI.NameManager(indexNewManager, "Z", $"NewManager!$H$1:$H$2");*/
            excelOperationNPOI.NameManager(indexNewManager, "XYZ", "$E$1:$E$3");
            excelOperationNPOI.NameManager(indexNewManager, "X", "$F$1:$F$2");
            excelOperationNPOI.NameManager(indexNewManager, "Y", "$G$1:$G$2");
            excelOperationNPOI.NameManager(indexNewManager, "Z", "$H$1:$H$2");

            excelOperationNPOI.DropDownList(indexNewManager, "XYZ", 0, 0, 15, 0);
            excelOperationNPOI.DropDownList(indexNewManager, "INDIRECT($A1)", 0, 1, 15, 1);

            excelOperationNPOI.ProtectColumn("Test1", 0, false);
            excelOperationNPOI.ProtectRow("Test1", 0, false);
            excelOperationNPOI.ProtectCell("Test1", 0, 1, true);
            excelOperationNPOI.ProtectCell("Test1", 1, 0, true);
            excelOperationNPOI.SetProtectSheet("Test1", "XZ");

            bool jest = excelOperationNPOI.IsEmptyCell("Test1", 2, 0);
            bool niema = excelOperationNPOI.IsEmptyCell("Test1", 3, 0);
            excelOperationNPOI.SetCellType("Test1", 4, 0, "@");
            bool tylkoFormat = excelOperationNPOI.IsEmptyCell("Test1", 4, 0);
            { }
            /*int lastRowInColumnPusty = excelOperationNPOI.GetLastRowIndexInColumn("Pusty", 0);
            int lastColumnInRowPusty = excelOperationNPOI.GetLastColumnIndexInRow("Pusty", 0);
            int lastRowInColumnTest1 = excelOperationNPOI.GetLastRowIndexInColumn("Test1", 0);
            int lastRowInColumnTest12 = excelOperationNPOI.GetLastRowIndexInColumn("Test1", 1);
            int lastColumnInRowTest1 = excelOperationNPOI.GetLastColumnIndexInRow("Test1", 0);
            int lastColumnInRowTest12 = excelOperationNPOI.GetLastColumnIndexInRow("Test1", 10);
            { }*/

            /*excelOperationNPOI.NameManager(indexManager, "Kowalski", "B2:B3");
            excelOperationNPOI.NameManager(indexManager, "Nowak", "B4:B5");
            excelOperationNPOI.NameManager(indexManager, "Wiśniewski", "B6");

            excelOperationNPOI.DropDownList("Manager", "Manager", 1, 4, 5, 4, "A2:A6");

            excelOperationNPOI.DropDownList("Manager", "Manager", 1, 5, 1, 5, "INDIRECT(E2)");
            excelOperationNPOI.DropDownList("Manager", "Manager", 2, 5, 2, 5, "INDIRECT(E3)");
            excelOperationNPOI.DropDownList("Manager", "Manager", 3, 5, 3, 5, "INDIRECT(E4)");
            excelOperationNPOI.DropDownList("Manager", "Manager", 4, 5, 4, 5, "INDIRECT(E5)");
            excelOperationNPOI.DropDownList("Manager", "Manager", 5, 5, 5, 5, "INDIRECT(E6)");*/

            //excelOperationNPOI.HideRowAndColumn("Test1", 0, 0, 2, 2);
            //excelOperationNPOI.AddRow(index, 1);
            //excelOperationNPOI.AddColumn(index, 1);

            /*excelOperationNPOI.AddCellValueText(index, 1, 7, "Tekst32");

            excelOperationNPOI.AddCellValueText(index, 10, 7, "Tekst32");
            excelOperationNPOI.AddCellValueText(index, 1, 8, "Tekst32");
            excelOperationNPOI.AddCellValueText(index, 1, 10, "Tekst32");
            excelOperationNPOI.SetFont(new FontSettings().SetBold(true).SetTextCrossed(true).SetUnderline(true).SetTextColorARGB(255, 255, 0, 0).SetItalics(true).SetFontName("Arial Black").SetTextWrapping(true), index, 0, 4);
            excelOperationNPOI.SetFont(new FontSettings().SetBold(true).SetTextCrossed(true).SetUnderline(true).SetTextColorARGB(255, 255, 0, 0).SetItalics(true).SetFontName("Arial Black").SetTextWrapping(true), index, 0, 0, 0, 3);
            excelOperationNPOI.HeightRow(index, 0, 50);
            excelOperationNPOI.HeightRow(index, 1, 4, 100);
            excelOperationNPOI.HeightRow(index, new int[] { 5, 6 }, 50);

            excelOperationNPOI.WidthColumn(index, 0, 50);
            excelOperationNPOI.WidthColumn(index, 1, 4, 100);
            excelOperationNPOI.WidthColumn(index, new int[] { 5, 6 }, 50);
            excelOperationNPOI.CellColor(index, 0, 0, 255, 255, 255, 0);
            excelOperationNPOI.CellColor(index, 1,1,3,3,255,0,255,0);

            excelOperationNPOI.SetBorder(new BorderIndex[] { BorderIndex.All }, LinesIndex.Thick, "Test1", 1, 1, 3, 3, 255, 255, 0, 0, false);

            excelOperationNPOI.SetBorder(new BorderIndex[] { BorderIndex.All }, LinesIndex.Thick, "Test1", 2, 2, 5, 8, 255, 255, 0, 0, false);
            
            excelOperationNPOI.SetBorder(new BorderIndex[] { BorderIndex.All }, LinesIndex.Thick, "Test1", 7, 2, 10, 5, 255, 255, 0, 0, true);

            excelOperationNPOI.SetBorder(new BorderIndex[] { BorderIndex.Top }, LinesIndex.Hair, "Test1", 12, 2, 15, 6, 255, 255, 0, 0, true);
            
            excelOperationNPOI.SetHorizontalAlignment(HorizontalAlignmentIndex.Center, index, 0, 0, 0, 8);
            excelOperationNPOI.SetVerticalAlignment(VerticalAlignmentIndex.Center, index, 0, 0, 0, 8);
            excelOperationNPOI.ValueOrientation(index, 0, 0, 0, 8, 135);

            excelOperationNPOI.SetProtectSheet(index, "1234");*/

            //excelOperationNPOI.DeleteColumn(index, 0);
            //excelOperationNPOI.DeleteColumn("Test1", 7);
            //excelOperationNPOI.DeleteColumn("Test1", 6, 5);
            //excelOperationNPOI.DeleteRow(index, 0,4);
            //.SetFontSize(20).SetTextCrossed(true).SetUnderline(true).SetTextColorARGB(255, 255, 0, 0)


            /*excelOperationNPOI.SetHorizontalAlignment(HorizontalAlignmentIndex.General, index, 0, 0);
            excelOperationNPOI.SetHorizontalAlignment(HorizontalAlignmentIndex.Left, index, 0, 1);
            excelOperationNPOI.SetHorizontalAlignment(HorizontalAlignmentIndex.Center, index, 0, 2);
            excelOperationNPOI.SetHorizontalAlignment(HorizontalAlignmentIndex.Right, index, 0, 3);
            excelOperationNPOI.SetHorizontalAlignment(HorizontalAlignmentIndex.Fill, index, 0, 4);
            excelOperationNPOI.SetHorizontalAlignment(HorizontalAlignmentIndex.Justify, index, 0, 5);
            excelOperationNPOI.SetHorizontalAlignment(HorizontalAlignmentIndex.Distributed, index, 0, 6);

            excelOperationNPOI.SetVerticalAlignment(VerticalAlignmentIndex.Top, index, 0, 0);
            excelOperationNPOI.SetVerticalAlignment(VerticalAlignmentIndex.Center, index, 0, 1);
            excelOperationNPOI.SetVerticalAlignment(VerticalAlignmentIndex.Bottom, index, 0, 2);
            excelOperationNPOI.SetVerticalAlignment(VerticalAlignmentIndex.Justify, index, 0, 3);
            excelOperationNPOI.SetVerticalAlignment(VerticalAlignmentIndex.Distributed, index, 0, 4);


           

            


            //excelOperationNPOI.AddCellValueText(index, 3, 2, "Tekst4");
            //excelOperationNPOI.AddCellValueText(index, 4, 2, "Tekst5");
            //excelOperationNPOI.AddCellValueText(index, 5, 2, "Tekst6");

            excelOperationNPOI.AddCellValueCurrency(index, 3, 0, new Currency(99.99m, "PLN"));
            excelOperationNPOI.AddCellValueDate(index, 4, 0, new DateTime(2024, 07, 29));
            excelOperationNPOI.AddCellValueDecimal(index, 5, 0, 23.99m);
            excelOperationNPOI.AddCellValueDouble(index, 6, 0, 50.23);
            excelOperationNPOI.AddCellValueFraction(index, 7, 0, new Fraction(0.008m));
            excelOperationNPOI.AddCellValueInt(index, 8, 0, 2);
            excelOperationNPOI.AddCellValuePercent(index, 9, 0, new Percent(0.8m));
            excelOperationNPOI.AddCellValuePercent(index, 10, 0, new Percent(0.8m));
            excelOperationNPOI.AddCellValuePercent(index, 11, 0, new Percent(0.8m));
            excelOperationNPOI.AddCellValueTime(index, 12, 0, new Time(10, 30));
            excelOperationNPOI.AddCellValueCurrency(index, 13, 0, new Currency(99.99m, "PLN"));

            excelOperationNPOI.WidthColumn(index, 0, 100);
            excelOperationNPOI.WidthColumn(0, 1, 100);

            excelOperationNPOI.HeightRow(index, 0, 100);
            excelOperationNPOI.HeightRow(0, 1, 100);

            excelOperationNPOI.MergeCells(index, 16, 1, 18, 3);

            
            excelOperationNPOI.ValueOrientation(index, 4, 0, 45);
            
            excelOperationNPOI.CellColor(index, 4, 0, 255, 255, 0, 255);

            excelOperationNPOI.SetBorder(new BorderIndex[] { BorderIndex.All }, LinesIndex.Thick, "Test1", 4, 0, 255, 0, 0, 255);
            excelOperationNPOI.SetBorder(new BorderIndex[] { BorderIndex.All }, LinesIndex.Hair, "Test1", 6, 1, 255, 0, 0, 255);

            excelOperationNPOI.SetBorder(new BorderIndex[] { BorderIndex.Top, BorderIndex.Bottom }, LinesIndex.Thin, "Test1", 4, 4, 255, 0, 0, 255);
            excelOperationNPOI.SetBorder(new BorderIndex[] { BorderIndex.Left, BorderIndex.Right }, LinesIndex.Thick, "Test1", 4, 6, 255, 0, 0, 255);


            excelOperationNPOI.CellColor(0, 1, 1, 255, 255, 0, 0);
            excelOperationNPOI.CellColor(1, 1, 1, 255, 0, 0, 0);*/

            using (FileStream stream = new FileStream(@"C:\Users\pawel\Desktop\dane\TestExcelOperationNPOI.xlsx", FileMode.Create, FileAccess.ReadWrite))
            {
                workbook1.Write(stream);
            }
            excelOperationNPOI.Dispose();

            
            /*IWorkbook workbook = new XSSFWorkbook();
            ISheet sheet = workbook.CreateSheet("MySheet");
            NPOI.SS.UserModel.IRow row = sheet.CreateRow(0);
            ICell cell = row.CreateCell(0);
            cell.SetCellValue("Test");

            using(FileStream fileStream = new FileStream(@"C:\Users\pawel\Desktop\dane\TestExcelOperationNPOI.xlsx", FileMode.Create, FileAccess.ReadWrite))
            {
                workbook.Write(fileStream);
            }*/
        }
    }
}
