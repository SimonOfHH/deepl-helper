﻿using DeeplHelper;
using System.IO;

if (args.Length < 1)
{
    Console.WriteLine("You need to provide your DeepL API-Key");
    Console.ReadLine();
    return;
}
string? authKey = args[0];
string? input = String.Empty;
string? filename = String.Empty;
if ((args.Length >= 2) && (args[1] != null) && (!String.IsNullOrEmpty(args[1].Trim())))
{
    input = args[1];
}
if ((args.Length >= 3) && (args[2] != null) && (!String.IsNullOrEmpty(args[2].Trim())))
{
    filename = args[2];
}
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
            if (String.IsNullOrEmpty(filename))
                filename = Console.ReadLine();
            if (String.IsNullOrEmpty(filename))
            {
                Console.WriteLine("No file specified.");
                break;
            }
            if (!File.Exists(filename))
            {
                Console.WriteLine("File does not exist.");
                break;
            }
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
            if (String.IsNullOrEmpty(filename))
                filename = Console.ReadLine();            
            Console.Write("Do you want to use an existing glossary for translation? (y/n) (Default: n): ");
            string? choice = Console.ReadLine();
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
                    var selectedGlossary = result.ElementAt(selection - 1);
                }
            }
            break;
    }
    input = null;
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