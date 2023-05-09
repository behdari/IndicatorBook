// See https://aka.ms/new-console-template for more information
public class SentLetter
{
    public int Id { get; set; }
    public string File { get; set; }
    public string Pamphleteer { get; set; }
    public int Number { get; set; }
    public string Title { get; set; }
    public string Description { get; set; }
    public int NextNumber { get; set; }
    public bool HasAttachment { get; set; }
    public string Date { get; set; }
    public byte[] WordFile { get; set; }
    public string WordFileExtension { get; set; }
}
