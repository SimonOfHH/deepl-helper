using System.Text.RegularExpressions;
using DeepL;
using DeepL.Model;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace DeeplHelper;

public class DeeplHelperUtility
{
    public Translator translator;

    public DeeplHelperUtility(string authKey)
    {
        translator = new Translator(authKey);
    }
    public async Task<DeepL.Model.GlossaryInfo[]> GetGlossaries()
    {
        return await translator.ListGlossariesAsync();
    }
    public async Task<GlossaryEntries> GetGlossaryEntries(GlossaryInfo glossaryInfo)
    {
        return await translator.GetGlossaryEntriesAsync(glossaryInfo);
    }
    public async Task DeleteGlossary(GlossaryInfo glossaryInfo)
    {
        await translator.DeleteGlossaryAsync(glossaryInfo);
    }
    public async Task PrintExistingGlossaries(bool forDeletion = false, bool skipDetails = false)
    {
        var result = await GetGlossaries();
        if (result == null)
        {
            Console.WriteLine("No Glossaries exist at the moment");
            return;
        }
        if (result.Length == 0)
        {
            Console.WriteLine("No Glossaries exist at the moment");
            return;
        }
        await PrintExistingGlossaries(result, forDeletion, skipDetails);
    }
    public async Task PrintExistingGlossaries(DeepL.Model.GlossaryInfo[] glossaryInfos, bool forDeletion = false, bool skipDetails = false)
    {
        int counter = 0;
        foreach (var glossaryInfo in glossaryInfos)
        {
            counter++;
            Console.WriteLine(String.Format("[{0}]  {1} {2}", counter, glossaryInfo.GlossaryId, glossaryInfo.Name));
        }

        if (forDeletion)
            await SelectGlossaryForDelete(glossaryInfos);
        else
        {
            if (!skipDetails)
                await ShowGlossaryDetails(glossaryInfos);
        }
    }
    public async Task ShowGlossaryDetails(DeepL.Model.GlossaryInfo[] glossaryInfos)
    {
        Console.WriteLine("Show entries for (Default: 0): ");
        Console.WriteLine("");
        var input = Console.ReadLine();
        int selection = 0;
        int.TryParse(input, out selection);
        if (selection == 0)
            return;
        var selectedGlossary = glossaryInfos.ElementAt(selection - 1);
        var entries = await GetGlossaryEntries(selectedGlossary);

        foreach (var entry in entries.ToDictionary())
        {
            Console.WriteLine(String.Format("    {0} {1}", entry.Key.PadRight(30), entry.Value));
        }
    }
    public async Task SelectGlossaryForDelete(DeepL.Model.GlossaryInfo[] glossaryInfos)
    {
        Console.WriteLine("Select Glossary to delete (Default: 0): ");
        Console.WriteLine("");
        var input = Console.ReadLine();
        int selection = 0;
        int.TryParse(input, out selection);
        if (selection == 0)
            return;
        var selectedGlossary = glossaryInfos.ElementAt(selection - 1);
        await DeleteGlossary(selectedGlossary);
        Console.WriteLine(String.Format("Deleted glossary {0}", selectedGlossary.GlossaryId));
    }
    public async Task<GlossaryInfo> CreateGlossary(GlossaryEntries entries, string glossaryName = "Glossary", string sourceLanguage = "en", string targetLanguage = "de")
    {
        return await translator.CreateGlossaryAsync(glossaryName, sourceLanguage, targetLanguage, entries);
    }
    public Dictionary<string, string> ReadGlossaryFromExcel(string filename, bool skipHeader = false)
    {
        var entries = new Dictionary<string, string>();
        using (var spreadsheetDocument = SpreadsheetDocument.Open(filename, false))
        {
            var workbookPart = spreadsheetDocument.WorkbookPart;
            if (workbookPart == null || workbookPart.Workbook == null || workbookPart.Workbook.Sheets == null || workbookPart.SharedStringTablePart == null)
                return entries;
            WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();
            SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();
            foreach (Row r in sheetData.Elements<Row>().Skip((skipHeader ? 1 : 0)))
            {
                (int currentPositionleft, int currentPositionTop) = Console.GetCursorPosition();
                Console.Write(String.Format("Processing row {0} / {1}", r.RowIndex, sheetData.Elements<Row>().Count()));
                Console.SetCursorPosition(currentPositionleft, currentPositionTop);
                var cell1 = (Cell)r.ElementAt(0);
                var cell2 = (Cell)r.ElementAt(1);
                entries.Add(ExcelHelper.GetCellValue(workbookPart, cell1), ExcelHelper.GetCellValue(workbookPart, cell2));
            }
            Console.WriteLine();
        }
        return entries;
    }

    public async Task TranslateExcelFile(string filename, GlossaryInfo? glossaryInfo, bool skipHeader = false, string sourceLanguage = "en", string targetLanguage = "de", int columnToTranslate = 0, int resultColumn = 3)
    {
        var textTranslateOptions = new TextTranslateOptions();
        if (targetLanguage == "de") textTranslateOptions.Formality = Formality.More;
        if (glossaryInfo != null) textTranslateOptions.GlossaryId = glossaryInfo.GlossaryId;

        var entries = new Dictionary<string, string>();
        using (var spreadsheetDocument = SpreadsheetDocument.Open(filename, true))
        {
            var workbookPart = spreadsheetDocument.WorkbookPart;
            if (workbookPart == null || workbookPart.Workbook == null || workbookPart.Workbook.Sheets == null || workbookPart.SharedStringTablePart == null)
                return;
            WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();
            SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();
            foreach (Row r in sheetData.Elements<Row>().Skip((skipHeader ? 1 : 0)))
            {
                (int currentPositionleft, int currentPositionTop) = Console.GetCursorPosition();
                Console.Write(String.Format("Processing row {0} / {1}", r.RowIndex, sheetData.Elements<Row>().Count()));
                Console.SetCursorPosition(currentPositionleft, currentPositionTop);

                string valueToTranslate = ExcelHelper.GetCellValue(workbookPart, (Cell)r.ElementAt(columnToTranslate));
                DeepL.Model.TextResult result = await translator.TranslateTextAsync(valueToTranslate, sourceLanguage, targetLanguage, textTranslateOptions);
                int index = ExcelHelper.InsertSharedStringItem(spreadsheetDocument, result.Text);
                ExcelHelper.InsertCellInWorksheet("C", r.RowIndex, worksheetPart, index.ToString());
            }
        }
    }

    public GlossaryEntries DictionaryToGlossaryEntries(Dictionary<string, string> entries)
    {
        var glossaryEntries = new GlossaryEntries(entries);
        return glossaryEntries;
    }
}

public static class ExcelHelper
{
    // Given a document name and text, 
    // inserts a new worksheet and writes the text to cell "A1" of the new worksheet.
    /*
    public static void InsertText(string docName, string text)
    {
        
        // Open the document for editing.
        using (SpreadsheetDocument spreadSheet = SpreadsheetDocument.Open(docName, true))
        {
            // Get the SharedStringTablePart. If it does not exist, create a new one.
            SharedStringTablePart shareStringPart;
            if (spreadSheet.WorkbookPart.GetPartsOfType<SharedStringTablePart>().Count() > 0)
            {
                shareStringPart = spreadSheet.WorkbookPart.GetPartsOfType<SharedStringTablePart>().First();
            }
            else
            {
                shareStringPart = spreadSheet.WorkbookPart.AddNewPart<SharedStringTablePart>();
            }

            // Insert the text into the SharedStringTablePart.
            int index = InsertSharedStringItem(text, shareStringPart);

            // Insert a new worksheet.
            WorksheetPart worksheetPart = InsertWorksheet(spreadSheet.WorkbookPart);

            // Insert cell A1 into the new worksheet.
            Cell cell = InsertCellInWorksheet("A", 1, worksheetPart);

            // Set the value of cell A1.
            cell.CellValue = new CellValue(index.ToString());
            cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);

            // Save the new worksheet.
            worksheetPart.Worksheet.Save();
        }
    }
    */

    public static int InsertSharedStringItem(SpreadsheetDocument spreadSheet, string text)
    {
        SharedStringTablePart shareStringPart;
        if (spreadSheet.WorkbookPart.GetPartsOfType<SharedStringTablePart>().Count() > 0)
            shareStringPart = spreadSheet.WorkbookPart.GetPartsOfType<SharedStringTablePart>().First();
        else
            shareStringPart = spreadSheet.WorkbookPart.AddNewPart<SharedStringTablePart>();
        return InsertSharedStringItem(text, shareStringPart);
    }

    // Given text and a SharedStringTablePart, creates a SharedStringItem with the specified text 
    // and inserts it into the SharedStringTablePart. If the item already exists, returns its index.
    public static int InsertSharedStringItem(string text, SharedStringTablePart shareStringPart)
    {
        // If the part does not contain a SharedStringTable, create one.
        if (shareStringPart.SharedStringTable == null)
            shareStringPart.SharedStringTable = new SharedStringTable();

        int i = 0;

        // Iterate through all the items in the SharedStringTable. If the text already exists, return its index.
        foreach (SharedStringItem item in shareStringPart.SharedStringTable.Elements<SharedStringItem>())
        {
            if (item.InnerText == text)
            {
                return i;
            }

            i++;
        }

        // The text does not exist in the part. Create the SharedStringItem and return its index.
        shareStringPart.SharedStringTable.AppendChild(new SharedStringItem(new DocumentFormat.OpenXml.Spreadsheet.Text(text)));
        shareStringPart.SharedStringTable.Save();

        return i;
    }
    // Given a WorkbookPart, inserts a new worksheet.
    public static WorksheetPart InsertWorksheet(WorkbookPart workbookPart)
    {
        // Add a new worksheet part to the workbook.
        WorksheetPart newWorksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        newWorksheetPart.Worksheet = new Worksheet(new SheetData());
        newWorksheetPart.Worksheet.Save();

        Sheets sheets = workbookPart.Workbook.GetFirstChild<Sheets>();
        string relationshipId = workbookPart.GetIdOfPart(newWorksheetPart);

        // Get a unique ID for the new sheet.
        uint sheetId = 1;
        if (sheets.Elements<Sheet>().Count() > 0)
        {
            sheetId = sheets.Elements<Sheet>().Select(s => s.SheetId.Value).Max() + 1;
        }

        string sheetName = "Sheet" + sheetId;

        // Append the new worksheet and associate it with the workbook.
        Sheet sheet = new Sheet() { Id = relationshipId, SheetId = sheetId, Name = sheetName };
        sheets.Append(sheet);
        workbookPart.Workbook.Save();

        return newWorksheetPart;
    }
    // Given a column name, a row index, and a WorksheetPart, inserts a cell into the worksheet. 
    // If the cell already exists, returns it. 
    public static Cell InsertCellInWorksheet(string columnName, uint rowIndex, WorksheetPart worksheetPart, string sharedStringIndex)
    {
        Worksheet worksheet = worksheetPart.Worksheet;
        SheetData sheetData = worksheet.GetFirstChild<SheetData>();
        string cellReference = columnName + rowIndex;

        // If the worksheet does not contain a row with the specified row index, insert one.
        Row row;
        if (sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).Count() != 0)
            row = sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).First();
        else
        {
            row = new Row() { RowIndex = rowIndex };
            sheetData.Append(row);
        }

        // If there is not a cell with the specified column name, insert one.  
        if (row.Elements<Cell>().Where(c => c.CellReference.Value == columnName + rowIndex).Count() > 0)
        {
            return row.Elements<Cell>().Where(c => c.CellReference.Value == cellReference).First();
        }
        else
        {
            // Cells must be in sequential order according to CellReference. Determine where to insert the new cell.
            Cell refCell = null;
            foreach (Cell cell in row.Elements<Cell>())
            {
                if (cell.CellReference.Value.Length == cellReference.Length)
                {
                    if (string.Compare(cell.CellReference.Value, cellReference, true) > 0)
                    {
                        refCell = cell;
                        break;
                    }
                }
            }

            Cell newCell = new Cell() { CellReference = cellReference };
            row.InsertBefore(newCell, refCell);

            newCell.CellValue = new CellValue(sharedStringIndex);
            newCell.DataType = new EnumValue<CellValues>(CellValues.SharedString);

            worksheet.Save();
            return newCell;
        }
    }

    public static string GetCellValue(WorkbookPart workbookPart, Cell currentcell)
    {
        string currentcellvalue = string.Empty;
        if (currentcell.DataType != null)
        {
            if (currentcell.DataType == CellValues.SharedString)
            {
                int id = -1;

                if (Int32.TryParse(currentcell.InnerText, out id))
                {
                    SharedStringItem item = GetSharedStringItemById(workbookPart, id);

                    if (item.Text != null)
                    {
                        currentcellvalue = item.Text.Text;
                    }
                    else if (item.InnerText != null)
                    {
                        currentcellvalue = item.InnerText;
                    }
                    else if (item.InnerXml != null)
                    {
                        currentcellvalue = item.InnerXml;
                    }
                }
            }
        }
        return currentcellvalue;
    }
    private static SharedStringItem GetSharedStringItemById(WorkbookPart workbookPart, int id)
    {
        return workbookPart!.SharedStringTablePart!.SharedStringTable.Elements<SharedStringItem>().ElementAt(id);
    }
}