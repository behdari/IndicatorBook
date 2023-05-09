// See https://aka.ms/new-console-template for more information
public class RecivedLetter
{
    public int Id { get; set; }
    public int RowNumber { get; set; }
    public int PreviousRowNumber { get; set; }
    public string Date { get; set; }
    public string LetterOwners { get; set; }
    public string Description { get; set; }
    public bool HasAttachment { get; set; }
    public string RecivedLetterNumber { get; set; }
    public string RecivedLetterDate { get; set; }
    public byte[] ScanFile { get; set; }
    public string ScanFileExtension { get; set; }
}
