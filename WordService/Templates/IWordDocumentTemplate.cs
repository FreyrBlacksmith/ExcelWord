namespace ExscelToWord.WordService.Templates;

/// <summary>
/// This document provide to logic for  all template for word
/// ToD0: Do auto create of word templates as possible with tags for create order
/// </summary>
public interface IWordDocumentTemplate
{
    public string Label { get; init; }
    public string Description { get; init; }
    public string Location { get; init; }
    public int ColumnCounts { get;}
}