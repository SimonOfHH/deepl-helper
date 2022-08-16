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
    public async Task PrintExistingGlossaries(DeepL.Model.GlossaryInfo[] glossaryInfos, bool forDeletion = false)
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
            await ShowGlossaryDetails(glossaryInfos);
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
                var cell1 = (Cell)r.ElementAt(0);
                var cell2 = (Cell)r.ElementAt(1);
                entries.Add(GetCellValue(workbookPart, cell1), GetCellValue(workbookPart, cell2));
            }
        }
        return entries;
    }

    public GlossaryEntries DictionaryToGlossaryEntries(Dictionary<string, string> entries)
    {
        var glossaryEntries = new GlossaryEntries(entries);
        return glossaryEntries;
    }

    private string GetCellValue(WorkbookPart workbookPart, Cell currentcell)
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
    private SharedStringItem GetSharedStringItemById(WorkbookPart workbookPart, int id)
    {
        return workbookPart!.SharedStringTablePart!.SharedStringTable.Elements<SharedStringItem>().ElementAt(id);
    }
}