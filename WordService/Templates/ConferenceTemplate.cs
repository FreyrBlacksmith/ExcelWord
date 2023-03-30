namespace ExscelToWord.WordService.Templates;

#warning this only as template. need fo this more flexible
public class ConferenceTemplate : IWordDocumentTemplate
{
    public IEnumerable<IEnumerable<string>> Participants { get; init; }

    public string Label { get; init; }
    public string Description { get; init; }
    public string Location { get; init; }
    public int ColumnCounts => Participants.First().Count();
}