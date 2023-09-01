public class RenamingParser {
    public static Dictionary<string, string> ParseRename(string inputDirectory, string workbookName, string worksheetName) {
        SheetParser parser = new SheetParser(inputDirectory, workbookName, worksheetName);
        return parser.GetTable().ToDictionary(key => key[0], value => value[1]);
    }
}