using DeepL.Model;
using DeeplHelper;
using System.IO;

string authKey = String.Empty;
string? input = String.Empty;
string? filename = String.Empty;
ProcessArgs(args, ref authKey, ref input, ref filename);
var utility = new DeeplHelperUtility(authKey);
while (true)
{
    PrintMenu();
    if (String.IsNullOrEmpty(input))
        input = Console.ReadLine();
    if (input == null) return;
    switch (input.ToUpper())
    {
        case "C":
            if (!GetFilename(ref filename))
                break;
            var entries = utility.ReadGlossaryFromExcel(filename, true);
            var glossary = await utility.CreateGlossary(utility.DictionaryToGlossaryEntries(entries));
            if (glossary != null)
                Console.WriteLine("Succesfully created Glossary.");
            break;
        case "S":
        case "D":
            var result = await GetGlossaries((input.ToUpper() == "D") ? true : false);
            break;
        case "T":
            if (!GetFilename(ref filename))
                break;
            Console.Write("Do you want to use an existing glossary for translation? (y/n) (Default: n): ");
            string? choice = Console.ReadLine();
            GlossaryInfo selectedGlossary = null;
            if (choice != null)
            {
                if (choice.ToUpper() == "Y")
                {
                    result = await GetGlossaries(false, true);
                    Console.Write("Select glossary (Default: 0): ");
                    var inputGlossary = Console.ReadLine();
                    int selection = 0;
                    int.TryParse(inputGlossary, out selection);
                    if (selection == 0)
                        return;
                    selectedGlossary = result.ElementAt(selection - 1);
                }
            }
            await utility.TranslateExcelFile(filename, selectedGlossary, true);
            break;
    }
    input = null;
}

void ProcessArgs(string[] args, ref string authKey, ref string input, ref string filename)
{
    if (args.Length < 1)
    {
        Console.WriteLine("You need to provide your DeepL API-Key");
        Console.ReadLine();
        return;
    }
    if ((args[0].Trim() == "<DeepL API Key>") || (String.IsNullOrEmpty(args[0])))
    {
        Console.WriteLine("You need to provide your DeepL API-Key");
        Console.ReadLine();
        return;
    }
    authKey = args[0];
    input = String.Empty;
    filename = String.Empty;
    if ((args.Length >= 2) && (args[1] != null) && (!String.IsNullOrEmpty(args[1].Trim())) && (args[1].Trim() != "<Input Selection[optional]>"))
    {
        input = args[1];
    }
    if ((args.Length >= 3) && (args[2] != null) && (!String.IsNullOrEmpty(args[2].Trim())) && (args[2].Trim() != "<Filename[optional]>"))
    {
        filename = args[2];
    }
}
void PrintMenu()
{
    Console.WriteLine("==================================");
    Console.WriteLine("Select Option: ");
    Console.WriteLine("  [S] Show Glossaries");
    Console.WriteLine("  [C] Create Glossary from file");
    Console.WriteLine("  [D] Delete Glossary");
    Console.WriteLine("  [T] Translate Excel");
    Console.WriteLine("==================================");
    Console.WriteLine("");
}
bool GetFilename(ref string? filename)
{
    if (String.IsNullOrEmpty(filename))
        filename = Console.ReadLine();
    if (String.IsNullOrEmpty(filename))
    {
        Console.WriteLine("No file specified.");
        return false;
    }
    if (!File.Exists(filename))
    {
        Console.WriteLine("File does not exist.");
        return false;
    }
    return true;
}
async Task<DeepL.Model.GlossaryInfo[]> GetGlossaries(bool forDeletion = false, bool skipDetails = false)
{
    var result = await utility.GetGlossaries();
    if (result == null)
    {
        Console.WriteLine("No Glossaries exist at the moment");
        return null;
    }
    if (result.Length == 0)
    {
        Console.WriteLine("No Glossaries exist at the moment");
        return result;
    }
    await utility.PrintExistingGlossaries(result, forDeletion, skipDetails);
    return result;
}