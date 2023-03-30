namespace ExscelToWord.Helpers;

public class TableColumnInfo
{
    public string? Name { get; init; }
    public bool IsSelected { get; set; }
    public int Index { get; init; }

    public TableColumnInfo()
    {
        IsSelected = false;
    }
}