using DeeplHelper;

if (args.Length < 1)
{
    Console.WriteLine("You need to provide your DeepL API-Key");
    Console.ReadLine();
    return;
}
string? authKey = args[0];
string input = String.Empty;
string filename = String.Empty;
if ((args.Length >= 2) && (args[1] != null))
{
    input = args[1];
}
if ((args.Length >= 3) && (args[2] != null))
{
    filename = args[2];
}
var utility = new DeeplHelperUtility(authKey);
while (true)
{
    Console.WriteLine("Select Option: ");
    Console.WriteLine("  [S] Show Glossaries");
    Console.WriteLine("  [G] Create Glossary from file");
    Console.WriteLine("  [D] Delete Glossary");
    Console.WriteLine("");
    if (String.IsNullOrEmpty(input))
        input = Console.ReadLine();
    if (input == null) return;
    switch (input.ToUpper())
    {
        case "G":
            //if (String.IsNullOrEmpty(filename))
            //    filename = Console.ReadLine();
            filename = "/host-home-folder/OneDrive - 4PS Group BV/Documenten/Data/GlossaryTest.xlsx";
            var entries = utility.ReadGlossaryFromExcel(filename, true);
            var glossary = await utility.CreateGlossary(utility.DictionaryToGlossaryEntries(entries));
            break;
        case "S":
        case "D":
            var result = await utility.GetGlossaries();
            if (result == null)
            {
                Console.WriteLine("No Glossaries exist at the moment");
                break;
            }
            if (result.Length == 0)
            {
                Console.WriteLine("No Glossaries exist at the moment");
                break;
            }
            await utility.PrintExistingGlossaries(result, (input.ToUpper() == "D") ? true : false);
            break;
    }
}