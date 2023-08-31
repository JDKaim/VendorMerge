using System.ComponentModel;

public class DocumentName {
    
    private string FileLocation;
    private string SpreadsheetName;
    private string WorksheetName;    
    
    public DocumentName(string fileLocation, string spreadsheetName, string worksheetName) {
        this.FileLocation = fileLocation;
        this.SpreadsheetName = spreadsheetName;
        this.WorksheetName = worksheetName;
    }

    public string getFileLocation() {
        return this.FileLocation;
    }

    public string getSpreadsheetName() {
        return this.SpreadsheetName;
    }

    public string getWorksheetName() {
        return this.WorksheetName;
    }
}