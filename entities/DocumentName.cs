using System.ComponentModel;

public class DocumentName {
    
    public readonly string FileLocation;
    public readonly string SpreadsheetName;
    public readonly string WorksheetName;    
    
    public DocumentName(string fileLocation, string spreadsheetName, string worksheetName) {
        this.FileLocation = fileLocation;
        this.SpreadsheetName = spreadsheetName;
        this.WorksheetName = worksheetName;
    }
}