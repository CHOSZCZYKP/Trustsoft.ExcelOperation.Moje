using Soneta.Business;
using Soneta.Core;
using Soneta.Handel;
using Soneta.Kadry;
using Soneta.Types;
using Syncfusion.XlsIO;
using TestExcel;
using Trustsoft.ExcelOperation.Moje;


[assembly: Worker(typeof(TestSyncfusion), typeof(DokEwidencja))]
namespace TestExcel
{
    
    public class TestSyncfusion
    {
        [Context]
        public Session Session { get; set; }

        private readonly FromTo fd = new FromTo();

        [Action ("Test Syncfusion", Icon = ActionIcon.Test, Mode = ActionMode.Progress | ActionMode.SingleSession, Target = ActionTarget.Menu | ActionTarget.LocalMenu | ActionTarget.Divider | ActionTarget.ToolbarWithText)]
        public void Eksport()
        {
            
            using(ExcelOperationSyncfusion excelOperationSyncfusion = new ExcelOperationSyncfusion())
            {
                var workbook = excelOperationSyncfusion.CreateWorkbook();
                IWorkbook workbook1 = workbook as IWorkbook;
                //excelOperationSyncfusion.ChangeNameWorksheet("Sheet1", "Nowy");

                excelOperationSyncfusion.MetaData("Jan", "Export", "Nowy");

                
                int index = excelOperationSyncfusion.AddWorksheet("Test1");
                int worksheet1 = excelOperationSyncfusion.AddWorksheet("Test2");
                int workshetPusty = excelOperationSyncfusion.AddWorksheet("Pusty");
                //excelOperationSyncfusion.DeleteWorksheet("Test1");
                //excelOperationSyncfusion.ChangeNameWorksheet("Sheet1", "ZMIANA");
                //excelOperationSyncfusion.ChangeNameWorksheet(1, "ZMIANA1");
                List<Sheet> sheets = excelOperationSyncfusion.GetNameSheet();
                int i = 0;
                foreach (var item in sheets)
                {
                    excelOperationSyncfusion.AddCellValueInt(index, i, 1, item.Index);
                    excelOperationSyncfusion.AddCellValueText(index, i, 2, item.Name);
                    i++;
                }
                Time time = new Time(12, 15);
                excelOperationSyncfusion.AddCellValueTime(0, 0, 1, time);
                excelOperationSyncfusion.GetCellValueText(0, 0, 1);
                excelOperationSyncfusion.AddCellValuePercent(0, 1, 3, new Percent(0.8m));
                excelOperationSyncfusion.AddCellValueDate(0, 2, 3, new DateTime(2024, 8, 21));
                double procent = excelOperationSyncfusion.GetCellValueNumber(0, 1, 3);
                DateTime data = excelOperationSyncfusion.GetCellValueDate(0, 2, 3);
                /*excelOperationSyncfusion.AddCellValueText(0, 0, 1, "Bardzo długi test specjalinie do testów1");
                excelOperationSyncfusion.AddCellValueText(0, 1, 1, "Bardzo długi test specjalinie do testów11 hahahaha");
                excelOperationSyncfusion.AddCellValueText(0, 0, 2, "Bardzo długi test specjalinie do testów2");
                excelOperationSyncfusion.AddCellValueText(0, 0, 3, "Bardzo długi test specjalinie do testów3");
                excelOperationSyncfusion.SetAutoWidth(0, 1, 2);*/

                excelOperationSyncfusion.AddCellValueInt(0, 0, 0, 0);
                excelOperationSyncfusion.AddCellValueInt(0, 1, 0, 1);
                excelOperationSyncfusion.AddCellValueInt(0, 2, 0, 2);
                excelOperationSyncfusion.AddCellValueInt(0, 3, 0, 3);
                excelOperationSyncfusion.AddCellValueInt(0, 4, 0, 4);
                excelOperationSyncfusion.AddCellValueInt(0, 5, 0, 5);
                excelOperationSyncfusion.AddCellValueInt(0, 6, 0, 6);
                excelOperationSyncfusion.AddCellValueText(0, 0, 0, "OK");
                excelOperationSyncfusion.AddCellValueText(0, 1, 0, "ERROR");
                excelOperationSyncfusion.AddCellValueText(0, 2, 0, "OK");
                excelOperationSyncfusion.AddCellValueInt(0, 3, 0, 1);
                excelOperationSyncfusion.AddCellValueInt(0, 4, 0, 1);
                excelOperationSyncfusion.AddCellValueInt(0, 5, 0, 3);
                excelOperationSyncfusion.AddCellFormula(0, 6, 0, "SUM(A4:A6)");
                excelOperationSyncfusion.ConditionalFormatting(0, 0, 0, 6, 0, new ConditionAndFormatting[] 
                {
                    //new ConditionAndFormatting(ComparisonOperatorIndex.Equal, "OK").SetBackgroundColor(255, 198, 239, 206).SetTextColor(255, 0, 100, 0),
                    new ConditionAndFormatting(ComparisonOperatorIndex.Equal, "\"OK\"").SetBackgroundColor(255, 198, 239, 206).SetTextColor(255, 0, 100, 0),
                    //new ConditionAndFormatting(ComparisonOperatorIndex.Equal, "ERROR").SetBackgroundColor(255, 255, 199, 206).SetTextColor(255, 190, 1, 7),
                    new ConditionAndFormatting(ComparisonOperatorIndex.Equal, "\"ERROR\"").SetBackgroundColor(255, 255, 199, 206).SetTextColor(255, 190, 1, 7),
                    new ConditionAndFormatting(ComparisonOperatorIndex.Equal, "3").SetBackgroundColor(255,0,0,255)
                });

                

                int IndexDane = excelOperationSyncfusion.AddWorksheet("Data");
                excelOperationSyncfusion.AddCellValueText(IndexDane, 0, 0, "Kod");
                excelOperationSyncfusion.AddCellValueText(IndexDane, 0, 1, "Imię i Nazwisko");
                excelOperationSyncfusion.AddCellValueText(IndexDane, 1, 0, "Code");
                excelOperationSyncfusion.AddCellValueText(IndexDane, 1, 1, "Name and Surname");

                excelOperationSyncfusion.AddCellValueText(IndexDane, 2, 0, "0002");
                excelOperationSyncfusion.AddCellValueText(IndexDane, 2, 1, "Jan Nowak_1");
                excelOperationSyncfusion.AddCellValueText(IndexDane, 3, 0, "723908");
                excelOperationSyncfusion.AddCellValueText(IndexDane, 3, 1, "Jan Nowak_2");
                excelOperationSyncfusion.AddCellValueText(IndexDane, 4, 0, "723914");
                excelOperationSyncfusion.AddCellValueText(IndexDane, 4, 1, "Jan Nowak_3");
                excelOperationSyncfusion.AddCellValueText(IndexDane, 5, 0, "723990");
                excelOperationSyncfusion.AddCellValueText(IndexDane, 5, 1, "Jan Nowak_4");
                excelOperationSyncfusion.AddCellValueText(IndexDane, 6, 0, "725578");
                excelOperationSyncfusion.AddCellValueText(IndexDane, 6, 1, "Jan Nowak_5");
                excelOperationSyncfusion.AddCellValueText(IndexDane, 7, 0, "725585");
                excelOperationSyncfusion.AddCellValueText(IndexDane, 7, 1, "Jan Nowak_6");
                excelOperationSyncfusion.AddCellValueText(IndexDane, 8, 0, "726479");
                excelOperationSyncfusion.AddCellValueText(IndexDane, 8, 1, "Jan Nowak_7");
                excelOperationSyncfusion.AddCellValueText(IndexDane, 9, 0, "726827");
                excelOperationSyncfusion.AddCellValueText(IndexDane, 9, 1, "Jan Nowak_8");
                excelOperationSyncfusion.AddCellValueText(IndexDane, 10, 0, "727019");
                excelOperationSyncfusion.AddCellValueText(IndexDane, 10, 1, "Jan Nowak_9");
                excelOperationSyncfusion.AddCellValueText(IndexDane, 11, 0, "727815");
                excelOperationSyncfusion.AddCellValueText(IndexDane, 11, 1, "Jan Nowak_10");
                excelOperationSyncfusion.AddCellValueText(IndexDane, 12, 0, "727816");
                excelOperationSyncfusion.AddCellValueText(IndexDane, 12, 1, "Jan Nowak_11");
                excelOperationSyncfusion.AddCellValueText(IndexDane, 13, 0, "727936");
                excelOperationSyncfusion.AddCellValueText(IndexDane, 13, 1, "Jan Nowak_12");
                excelOperationSyncfusion.AddCellValueText(IndexDane, 14, 0, "727937");
                excelOperationSyncfusion.AddCellValueText(IndexDane, 14, 1, "Jan Nowak_13");

                //excelOperationSyncfusion.DropDownList(0, IndexDane, 0, 0, 20, 0, "$A$3:$A$15");

                excelOperationSyncfusion.NameManager("Data", "pawel", "A3:A15");

                excelOperationSyncfusion.AddCellValueText(index, 0, 0, "Tekst11");
                excelOperationSyncfusion.AddCellValueText(index, 0, 1, "Tekst12");
                excelOperationSyncfusion.AddCellValueText(index, 0, 2, "Tekst13");
                excelOperationSyncfusion.AddCellValueText(index, 1, 0, "Tekst21");
                excelOperationSyncfusion.AddCellValueText(index, 1, 1, "Tekst22");
                excelOperationSyncfusion.AddCellValueText(index, 1, 2, "Tekst23");
                excelOperationSyncfusion.AddCellValueText(index, 2, 0, "Tekst31");
                excelOperationSyncfusion.AddCellValueText(index, 2, 1, "Tekst32");
                excelOperationSyncfusion.AddCellValueText(index, 2, 2, "Tekst33");

                excelOperationSyncfusion.AddCellFormula(index, 10, 1, "SUM(A6:A10)");

                var lastRow = excelOperationSyncfusion.GetLastRow("Test1");
                var lastColumn = excelOperationSyncfusion.GetLastColumn("Test1");
                excelOperationSyncfusion.HideSheet(3, SheetVisibilityIndex.Hidden);
                excelOperationSyncfusion.ActiveSheet(1);
                //excelOperationSyncfusion.HideRow("Test1", 0, 1);

                int indexManager = excelOperationSyncfusion.AddWorksheet("Manager");
                int indexNewManager = excelOperationSyncfusion.AddWorksheet("NewManager");
                excelOperationSyncfusion.AddCellValueText(indexManager, 0, 0, "Nazwisko");
                excelOperationSyncfusion.AddCellValueText(indexManager, 0, 1, "Imie");
                excelOperationSyncfusion.AddCellValueText(indexManager, 1, 0, "Kowalski");
                excelOperationSyncfusion.AddCellValueText(indexManager, 1, 1, "Jan");
                excelOperationSyncfusion.AddCellValueText(indexManager, 2, 0, "Kowalski");
                excelOperationSyncfusion.AddCellValueText(indexManager, 2, 1, "Anna");
                excelOperationSyncfusion.AddCellValueText(indexManager, 3, 0, "Nowak");
                excelOperationSyncfusion.AddCellValueText(indexManager, 3, 1, "Piotr");
                excelOperationSyncfusion.AddCellValueText(indexManager, 4, 0, "Nowal");
                excelOperationSyncfusion.AddCellValueText(indexManager, 4, 1, "Ewa");
                excelOperationSyncfusion.AddCellValueText(indexManager, 5, 0, "Wiśniewski");
                excelOperationSyncfusion.AddCellValueText(indexManager, 5, 1, "Marek");

                excelOperationSyncfusion.AddCellValueText(indexNewManager, 0, 4, "X");
                excelOperationSyncfusion.AddCellValueText(indexNewManager, 0, 5, "xx1");
                excelOperationSyncfusion.AddCellValueText(indexNewManager, 0, 6, "yy1");
                excelOperationSyncfusion.AddCellValueText(indexNewManager, 0, 7, "zz1");
                excelOperationSyncfusion.AddCellValueText(indexNewManager, 1, 4, "Y");
                excelOperationSyncfusion.AddCellValueText(indexNewManager, 1, 5, "xx2");
                excelOperationSyncfusion.AddCellValueText(indexNewManager, 1, 6, "yy2");
                excelOperationSyncfusion.AddCellValueText(indexNewManager, 1, 7, "zz2");
                excelOperationSyncfusion.AddCellValueText(indexNewManager, 2, 4, "Z");

                /*excelOperationSyncfusion.NameManager(indexNewManager, "XYZ", $"NewManager!$E$1:$E$3");
                excelOperationSyncfusion.NameManager(indexNewManager, "X", $"NewManager!$F$1:$F$2");
                excelOperationSyncfusion.NameManager(indexNewManager, "Y", $"NewManager!$G$1:$G$2");
                excelOperationSyncfusion.NameManager(indexNewManager, "Z", $"NewManager!$H$1:$H$2");*/
                excelOperationSyncfusion.NameManager(indexNewManager, "XYZ", "$E$1:$E$3");
                excelOperationSyncfusion.NameManager(indexNewManager, "X", "$F$1:$F$2");
                excelOperationSyncfusion.NameManager(indexNewManager, "Y", "$G$1:$G$2");
                excelOperationSyncfusion.NameManager(indexNewManager, "Z", "$H$1:$H$2");

                excelOperationSyncfusion.DropDownList(indexNewManager, "XYZ", 0, 0, 15, 0);
                excelOperationSyncfusion.DropDownList(indexNewManager, "INDIRECT($A1)", 0, 1, 15, 1);


                /*excelOperationSyncfusion.DropDownList("NewManager", "Manager", 1, 0, 5, 0, "A2:A6");
                excelOperationSyncfusion.NameManager(indexManager, "Kowalski", "B2:B3");
                excelOperationSyncfusion.NameManager(indexManager, "Nowak", "B4:B5");
                excelOperationSyncfusion.NameManager(indexManager, "Wiśniewski", "B6");*/

                excelOperationSyncfusion.ProtectColumn("Test1", 0, false);
                excelOperationSyncfusion.ProtectRow("Test1", 0, false);
                excelOperationSyncfusion.ProtectCell("Test1", 0, 1, true);
                excelOperationSyncfusion.ProtectCell("Test1", 1, 0, true);
                excelOperationSyncfusion.SetProtectSheet("Test1", "XZ");


                bool jest = excelOperationSyncfusion.IsEmptyCell("Test1", 2, 0);
                bool niema = excelOperationSyncfusion.IsEmptyCell("Test1", 3, 0);
                excelOperationSyncfusion.SetCellType("Test1", 4, 0, "@");
                bool tylkoFormat = excelOperationSyncfusion.IsEmptyCell("Test1", 4, 0);
                { }
                /*int lastRowInColumnPusty = excelOperationSyncfusion.GetLastRowIndexInColumn("Pusty",1);
                int lastColumnInRowPusty = excelOperationSyncfusion.GetLastColumnIndexInRow("Pusty", 1);
                int lastRowInColumnTest1 = excelOperationSyncfusion.GetLastRowIndexInColumn("Test1", 0);
                int lastRowInColumnTest12 = excelOperationSyncfusion.GetLastRowIndexInColumn("Test1", 1);
                int lastColumnInRowTest1 = excelOperationSyncfusion.GetLastColumnIndexInRow("Test1", 0);
                int lastColumnInRowTest12 = excelOperationSyncfusion.GetLastColumnIndexInRow("Test1", 10);
                { }*/



                //excelOperationSyncfusion.HideRowAndColumn("Test1", 0, 0, 2, 2);
                //excelOperationSyncfusion.AddRow(index, 1);
                //excelOperationSyncfusion.AddColumn(index, 1);
                /*excelOperationSyncfusion.AddCellValueText(index, 1, 7, "Tekst32");

                excelOperationSyncfusion.AddCellValueText(index, 10, 7, "Tekst32");
                excelOperationSyncfusion.AddCellValueText(index, 1, 8, "Tekst32");
                excelOperationSyncfusion.AddCellValueText(index, 1, 10, "Tekst32");*/
                /*excelOperationSyncfusion.SetFont(new FontSettings().SetBold(true).SetTextCrossed(true).SetUnderline(true).SetTextColorARGB(255, 255, 0, 0).SetItalics(true).SetFontName("Arial Black").SetTextWrapping(true), index, 0, 4);
                excelOperationSyncfusion.SetFont(new FontSettings().SetBold(true).SetTextCrossed(true).SetUnderline(true).SetTextColorARGB(255, 255, 0, 0).SetItalics(true).SetFontName("Arial Black").SetTextWrapping(true), index, 0, 0, 0, 3);


                excelOperationSyncfusion.HeightRow(index, 0, 50);
                excelOperationSyncfusion.HeightRow(index, 1, 4, 100);
                excelOperationSyncfusion.HeightRow(index, new int[] { 5, 6 }, 50);

                excelOperationSyncfusion.WidthColumn(index, 0, 50);
                excelOperationSyncfusion.WidthColumn(index, 1, 4, 100);
                excelOperationSyncfusion.WidthColumn(index, new int[] { 5, 6 }, 50);
                excelOperationSyncfusion.CellColor(index, 1, 1, 3, 3, 255, 0, 255, 0);
                excelOperationSyncfusion.CellColor(index, 0, 0, 255, 255, 255, 0);
                //excelOperationSyncfusion.SetBorder(new BorderIndex[] { BorderIndex.All }, LinesIndex.Dotted, "Test1", 0, 0);

                excelOperationSyncfusion.SetBorder(new BorderIndex[] { BorderIndex.All }, LinesIndex.Thick, "Test1", 1, 1, 3, 3, 255, 255, 0, 0, false);
                
                excelOperationSyncfusion.SetBorder(new BorderIndex[] { BorderIndex.All }, LinesIndex.Thick, index, 2, 2, 5, 8, 255, 255, 0, 0, false);
                excelOperationSyncfusion.SetBorder(new BorderIndex[] { BorderIndex.All }, LinesIndex.Thick, index, 7, 2, 10, 5, 255, 255, 0, 0, true);

                excelOperationSyncfusion.SetBorder(new BorderIndex[] { BorderIndex.Top}, LinesIndex.Hair, "Test1", 12, 2, 15, 6, 255, 255, 0, 0, true);


                excelOperationSyncfusion.CellColor(index, 2, 10, 3, 11, 255, 0, 255, 255);
                excelOperationSyncfusion.SetHorizontalAlignment(HorizontalAlignmentIndex.Center, index, 0, 0, 0, 8);
                excelOperationSyncfusion.SetVerticalAlignment(VerticalAlignmentIndex.Center, index, 0, 0, 0, 8);
                excelOperationSyncfusion.ValueOrientation(index, 0, 0, 0, 8, 135);

                excelOperationSyncfusion.SetProtectSheet(index, "1234");*/

                //excelOperationNPOI.DeleteColumn(index, 0);
                //excelOperationNPOI.DeleteColumn("Test1", 7);
                //excelOperationSyncfusion.DeleteColumn("Test1", 6, 5);
                //excelOperationSyncfusion.DeleteRow(index, 0,4);

                /*excelOperationSyncfusion.SetHorizontalAlignment(HorizontalAlignmentIndex.General, index, 0, 0);
                excelOperationSyncfusion.SetHorizontalAlignment(HorizontalAlignmentIndex.Left, index, 0, 1);
                excelOperationSyncfusion.SetHorizontalAlignment(HorizontalAlignmentIndex.Center, index, 0, 2);
                excelOperationSyncfusion.SetHorizontalAlignment(HorizontalAlignmentIndex.Right, index, 0, 3);
                excelOperationSyncfusion.SetHorizontalAlignment(HorizontalAlignmentIndex.Fill, index, 0, 4);
                excelOperationSyncfusion.SetHorizontalAlignment(HorizontalAlignmentIndex.Justify, index, 0, 5);
                excelOperationSyncfusion.SetHorizontalAlignment(HorizontalAlignmentIndex.Distributed, index, 0, 6);
                

                excelOperationSyncfusion.SetVerticalAlignment(VerticalAlignmentIndex.Top, index, 0, 0);
                excelOperationSyncfusion.SetVerticalAlignment(VerticalAlignmentIndex.Center, index, 0, 1);
                excelOperationSyncfusion.SetVerticalAlignment(VerticalAlignmentIndex.Bottom, index, 0, 2);
                excelOperationSyncfusion.SetVerticalAlignment(VerticalAlignmentIndex.Justify, index, 0, 3);
                excelOperationSyncfusion.SetVerticalAlignment(VerticalAlignmentIndex.Distributed, index, 0, 4);

                //excelOperationSyncfusion.AddRow(index, 5, 15);
                
                //excelOperationSyncfusion.DeleteRow(index, 3);
                //excelOperationSyncfusion.DeleteRow("Test1", 2);
                //excelOperationSyncfusion.DeleteRow(1, "Test1");

                //excelOperationSyncfusion.DeleteColumn(index, 3);
                //excelOperationSyncfusion.DeleteColumn("Test1", 2);
                //excelOperationSyncfusion.DeleteColumn(1, "Test1");

                //excelOperationSyncfusion.AddCellValueInt(index, 1, 1, 1);
                excelOperationSyncfusion.AddCellValueDecimal(index, 2, 0, 1.05m);
                excelOperationSyncfusion.AddCellValueDouble(index, 2, 0, 3.33);
                excelOperationSyncfusion.AddCellValueFraction(index, 3, 0, new Fraction(1.008m));
                excelOperationSyncfusion.AddCellValueDate(index, 4, 0, new DateTime(2024, 7, 25));
                excelOperationSyncfusion.AddCellValueCurrency(index, 5, 0, new Currency(99.99, "PLN"));
                excelOperationSyncfusion.AddCellValuePercent(index, 6, 0, new Percent(0.83m));
                excelOperationSyncfusion.AddCellValueTime(index, 7, 0, new Time(10, 00));
                excelOperationSyncfusion.AddCellValueCurrency(index, 8, 0, new Currency(99.99, "PLN"));

                excelOperationSyncfusion.WidthColumn(index, 0, 100);
                excelOperationSyncfusion.WidthColumn(0, 1, 100);

                excelOperationSyncfusion.HeightRow(index, 0, 100);
                excelOperationSyncfusion.HeightRow(0, 1, 100);

                excelOperationSyncfusion.SetBorder(new BorderIndex[] {BorderIndex.All}, LinesIndex.Hair, "Test1", 19, 1);
                excelOperationSyncfusion.SetBorder(new BorderIndex[] { BorderIndex.Top, BorderIndex.Bottom }, LinesIndex.Double, "Test1", 19, 3);
                excelOperationSyncfusion.SetBorder(new BorderIndex[] { BorderIndex.Left, BorderIndex.Right }, LinesIndex.Thick, "Test1", 19, 5);

                BorderIndex[] borderIndex = new BorderIndex[] {BorderIndex.Top, BorderIndex.Bottom};
                
                excelOperationSyncfusion.SetBorder(borderIndex, LinesIndex.Medium, "Test1", 7, 7, 255, 255, 0, 0);
                excelOperationSyncfusion.SetBorder(new BorderIndex[] { BorderIndex.Top, BorderIndex.Bottom }, LinesIndex.Thick, "Test1", 7, 9, 255, 0, 255, 0);
                excelOperationSyncfusion.SetBorder(new BorderIndex[] {BorderIndex.Left, BorderIndex.Right}, LinesIndex.Dotted, "Test1", 7, 11, 255, 0, 0, 255);
                excelOperationSyncfusion.SetBorder(new BorderIndex[] {BorderIndex.Right}, LinesIndex.Dashed, "Test1", 7, 13, 255, 255, 255, 0);
                
                excelOperationSyncfusion.MergeCells(index, 15, 0, 17, 2);
                excelOperationSyncfusion.ValueOrientation(index, 0, 0, 45);
                
                excelOperationSyncfusion.CellColor(index, 0, 0, 255, 0, 0, 255);*/




                /*
                excelOperationSyncfusion.HeightRow(index, 1, 120.5);
                excelOperationSyncfusion.WidthColumn(index, 1, 200.38);
                

                excelOperationSyncfusion.BoldFont(index, 3, 1);
                excelOperationSyncfusion.Italics(index, 3, 1);
                excelOperationSyncfusion.Underline(index, 3, 1);
                excelOperationSyncfusion.DoubleUnderline(index, 4, 1);
                excelOperationSyncfusion. TextCrossed(index, 3, 1);

                excelOperationSyncfusion.CellColor(index, 1, 1, 255, 0, 0);
                excelOperationSyncfusion.TextColor(index, 1, 1, 0, 0, 255);
                excelOperationSyncfusion.FontAndSize(index, 1, 1, "Arial Black", 20.5);

                excelOperationSyncfusion.TopEdge(index, 2, 2, 255, 0, 255, 0);
                excelOperationSyncfusion.BottomEdge(index, 2, 2, 255, 0, 255, 0);
                excelOperationSyncfusion.RightEdge(index, 2, 2, 255, 0, 255, 0);
                excelOperationSyncfusion.LeftEdge(index, 2, 2, 255, 0, 255, 0);

                excelOperationSyncfusion.AllEdges(index, 5, 5, 255, 0, 255, 0);

                excelOperationSyncfusion.ClearTopEdges(index, 2, 2);
                excelOperationSyncfusion.ClearBottomEdges(index, 2, 2);
                excelOperationSyncfusion.ClearLeftEdges(index, 2, 2);
                excelOperationSyncfusion.ClearRightEdges(index, 2, 2);

                excelOperationSyncfusion.ClearEdges(index, 5, 5);

                excelOperationSyncfusion.DoubleBottomEdge(index, 2, 2, 255, 0, 255, 0);
                excelOperationSyncfusion.DoubleLeftEdge(index, 2, 2, 255, 0, 255, 0);
                excelOperationSyncfusion.DoubleRightEdge(index, 2, 2, 255, 0, 255, 0);
                excelOperationSyncfusion.DoubleTopEdge(index, 2, 2, 255, 0, 255, 0);

                excelOperationSyncfusion.AllDoubleEdges(index, 5, 5, 255, 0, 255, 0);

                excelOperationSyncfusion.ThickBottomEdge(index, 8, 8, 255, 0, 255, 0);
                excelOperationSyncfusion.ThickLeftEdge(index, 8, 8, 255, 0, 255, 0);
                excelOperationSyncfusion.ThickRightEdge(index, 8, 8, 255, 0, 255, 0);
                excelOperationSyncfusion.ThickTopEdge(index, 8, 8, 255, 0, 255, 0);

                excelOperationSyncfusion.AllThickEdges(index, 10, 10, 255, 0, 255, 0);

                excelOperationSyncfusion.TopAndBottomEdge(index, 3, 1, 255, 0, 255, 0);

                excelOperationSyncfusion.RightAndLeftEdge(index, 1, 3, 255, 0, 255, 0);

                excelOperationSyncfusion.TopEdgeAndThickBottomEdge(index, 10, 1, 255, 0, 255, 0);//

                excelOperationSyncfusion.TopEdgeAndDoubleBottomEdge(index, 20, 1, 255, 255, 255, 0);

                excelOperationSyncfusion.MediumBottomEdge(index, 20, 20, 255, 255, 0, 0);
                excelOperationSyncfusion.MediumLeftEdge(index, 20, 20, 255, 255, 0, 0);
                excelOperationSyncfusion.MediumRightEdge(index, 20, 20, 255, 255, 0, 0);
                excelOperationSyncfusion.MediumTopEdge(index, 20, 20, 255, 255, 0, 0);

                excelOperationSyncfusion.AllMediumEdge(index, 20, 10, 255, 0, 255, 255);

                excelOperationSyncfusion.TopAlign(index, 2, 2);
                excelOperationSyncfusion.BottomAlign(index, 2, 4);
                excelOperationSyncfusion.CenterVerticalAlign(index, 2, 5);
                excelOperationSyncfusion.CenterHorizontalAlign(index, 4, 5);
                excelOperationSyncfusion.RightAlign(index, 4, 2);
                excelOperationSyncfusion.LeftAlign(index, 4, 4);
                excelOperationSyncfusion.TextWrapping(index, 5, 2);

                excelOperationSyncfusion.MergeCells(index, "A14:D14");
                excelOperationSyncfusion.MergeCells(index, "A16:D18");

                excelOperationSyncfusion.MergeCells(index, 25, 1, 27, 4);
                excelOperationSyncfusion.MergeCells(index, 30, 1, 32, 1);

                excelOperationSyncfusion.ValueOrientation(index, 1, 1, 90);

                //excelOperationSyncfusion.AddRow(index, 3);
                //excelOperationSyncfusion.AddColumn(index, 1);

                //excelOperationSyncfusion.DeleteRow(index, 1);
                //excelOperationSyncfusion.DeleteRow("Test1", 1);
                //excelOperationSyncfusion.DeleteRow(1, "Test1");

                //excelOperationSyncfusion.DeleteColumn(index, 1);
                //excelOperationSyncfusion.DeleteColumn("Test1", 1);
                //excelOperationSyncfusion.DeleteColumn(1, "Test1");

                excelOperationSyncfusion.TopEdge(index, 2, 8, 255, 0, 0, 255);
                excelOperationSyncfusion.BottomEdge(index, 2, 8, 255, 0, 0, 255);
                excelOperationSyncfusion.RightEdge(index, 2, 8, 255, 0, 0, 255);
                excelOperationSyncfusion.LeftEdge(index, 2, 8, 255, 0, 0, 255);


                int index1 = excelOperationSyncfusion.AddWorksheet("Test2");
                excelOperationSyncfusion.SetBorder(BorderIndex.Top,"Test2", 2,3);
                excelOperationSyncfusion.TopEdge(index1, 2, 2);
                excelOperationSyncfusion.BottomEdge(index1, 2, 2);
                excelOperationSyncfusion.RightEdge(index1, 2, 2);
                excelOperationSyncfusion.LeftEdge(index1, 2, 2);

                excelOperationSyncfusion.TopEdge(index1, 2, 8, 255, 0, 0, 0);
                excelOperationSyncfusion.BottomEdge(index1, 2, 8, 255, 0, 0, 0);
                excelOperationSyncfusion.RightEdge(index1, 2, 8, 255, 0, 0, 0);
                excelOperationSyncfusion.LeftEdge(index1, 2, 8, 255, 0, 0, 0);

                excelOperationSyncfusion.AllEdges(index1, 2, 4);

                excelOperationSyncfusion.DoubleBottomEdge(index1, 4, 2);
                excelOperationSyncfusion.DoubleLeftEdge(index1, 4, 2);
                excelOperationSyncfusion.DoubleRightEdge(index1, 4, 2);
                excelOperationSyncfusion.DoubleTopEdge(index1, 4, 2);
                
                excelOperationSyncfusion.AllDoubleEdges(index1, 4, 4);

                excelOperationSyncfusion.ThickBottomEdge(index1, 6, 2);
                excelOperationSyncfusion.ThickLeftEdge(index1, 6, 2);
                excelOperationSyncfusion.ThickRightEdge(index1, 6, 2);
                excelOperationSyncfusion.ThickTopEdge(index1, 6, 2);

                excelOperationSyncfusion.AllThickEdges(index1, 6, 4);

                excelOperationSyncfusion.MediumBottomEdge(index1, 8, 2);
                excelOperationSyncfusion.MediumLeftEdge(index1, 8, 2);
                excelOperationSyncfusion.MediumRightEdge(index1, 8, 2);
                excelOperationSyncfusion.MediumTopEdge(index1, 8, 2);

                excelOperationSyncfusion.AllMediumEdge(index1, 8, 4);

                excelOperationSyncfusion.TopAndBottomEdge(index1, 10, 2);
                excelOperationSyncfusion.RightAndLeftEdge(index1, 10, 4);
                excelOperationSyncfusion.TopEdgeAndThickBottomEdge(index1, 10, 6);
                excelOperationSyncfusion.TopEdgeAndDoubleBottomEdge(index1, 10, 8);*/

                using (FileStream stream = new FileStream(@"C:\Users\pawel\Desktop\dane\TestExcelOperationSyncfusion.xlsx", FileMode.Create, FileAccess.ReadWrite))
                {
                    workbook1.SaveAs(stream);
                }
                excelOperationSyncfusion.Dispose();
            }
        }
    }
}
