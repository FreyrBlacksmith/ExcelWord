using ExscelToWord.WordService.Templates;
using NPOI.OpenXmlFormats.Wordprocessing;
using NPOI.XWPF.UserModel;

namespace ExscelToWord.WordService;
public class WordCreationResult
{
    public bool Status { get; set; }
    public string Message { get; set; }
}
public static class WordService
{
    public static WordCreationResult LoadAndFillWordDocument(string filePath, ConferenceTemplate template)
    {
        var result = new WordCreationResult();
        XWPFDocument document = null;
        try
        {
            using (var stream = new FileStream(filePath, FileMode.Open))
            {
                document = new XWPFDocument(stream);
            }
        }
        catch (Exception e)
        {
            result.Status = false;
            result.Message = e.Message;
        }

        if (!File.Exists(filePath))
        {
            CreateWordDocument(template, filePath);
        }
        else
        {
            try
            {
                FillTables(document, template.Participants);

                using (FileStream fs = new FileStream(filePath, FileMode.Create))
                {
                    document.Write(fs);
                    result.Status = true;
                    result.Message = "Word file created";
                }
            }
            catch (Exception e)
            {
                result.Status = false;
                result.Message = e.Message;
            }

        }

        return result;
    }

    /// <summary>
    /// This is a method for full document creation by template
    /// </summary>
    /// <param name="temaplete"></param>
    /// <param name="fileName"></param>
    /// <param name="imageFilePath"></param>
    private static void CreateWordDocument(ConferenceTemplate temaplete, string fileName)
    {
        XWPFDocument doc = new XWPFDocument();

        FillLablel(doc, temaplete.Label);
        FillDescription(doc, temaplete.Label);
       // SetupPicture(doc, imageFilePath);
        FillLocation(doc, temaplete.Location);
        FillTables(doc, temaplete.Participants);

        using (FileStream fs = new FileStream(fileName, FileMode.Create))
        {
            doc.Write(fs);
        }
    }


    private static string FillTables(XWPFDocument doc, IEnumerable<IEnumerable<string>> participants)
    {
        var counter = 0;
        var columnCount = participants.First().Count();
        XWPFTable table = doc.CreateTable(participants.Count(), columnCount + 1);
        var tablelLayout = table.GetCTTbl().tblPr.AddNewTblLayout();
        tablelLayout.type = ST_TblLayoutType.@fixed;


        CT_TblBorders tblBorders = new CT_TblBorders();
        tblBorders.top = new CT_Border { val = ST_Border.none };
        tblBorders.bottom = new CT_Border { val = ST_Border.none };
        tblBorders.left = new CT_Border { val = ST_Border.none };
        tblBorders.right = new CT_Border { val = ST_Border.none };
        tblBorders.insideH = new CT_Border { val = ST_Border.none };
        tblBorders.insideV = new CT_Border { val = ST_Border.none };

        table.GetCTTbl().tblPr = new CT_TblPr();
        table.GetCTTbl().tblPr.tblBorders = tblBorders;


        for (int i = 0; i <= columnCount; i++)
        {
            table.SetColumnWidth(i, 1200);

        }

        try
        {
            foreach (var participant in participants)
            {
                var columnIndex = 1;
                table.GetRow(counter).GetCell(0).SetText((counter + 1).ToString() + '.');
                foreach (var field in participant)
                {
                    var tabeRow = table.GetRow(counter);
                    if (tabeRow != null)
                    {
                        tabeRow.GetCell(columnIndex).SetText(field);
                    }
                    else
                    {
                        return "Table not contain this column";
                    }
                    columnIndex++;
                }
                counter++;
            }
        }
        catch (Exception e)
        {
            return e.Message;
        }

        return "Ok";
    }


    private static void FillLablel(this XWPFDocument doc, string label)
    {
        XWPFParagraph p1 = doc.CreateParagraph();
        p1.Alignment = ParagraphAlignment.LEFT;
        p1.BorderBottom = Borders.Double;
        p1.BorderTop = Borders.Double;

        p1.BorderRight = Borders.Double;
        p1.BorderLeft = Borders.Double;
        p1.BorderBetween = Borders.Single;

        p1.VerticalAlignment = TextAlignment.TOP;

        XWPFRun r1 = p1.CreateRun();
        r1.SetText(label);
        r1.IsBold = true;
        r1.FontFamily = "Times New Roman";
        r1.FontSize = 15;
    }

    private static void FillDescription(this XWPFDocument doc, string description)
    {
        XWPFParagraph p1 = doc.CreateParagraph();
        p1.Alignment = ParagraphAlignment.LEFT;
        p1.BorderBottom = Borders.Double;
        p1.BorderTop = Borders.Double;

        p1.BorderRight = Borders.Double;
        p1.BorderLeft = Borders.Double;
        p1.BorderBetween = Borders.Single;

        p1.VerticalAlignment = TextAlignment.TOP;

        XWPFRun r1 = p1.CreateRun();
        r1.SetText(description);
        r1.IsBold = true;
        r1.FontFamily = "Times New Roman";
        r1.FontSize = 15;
    }

    private static void FillLocation(this XWPFDocument doc, string location)
    {
        XWPFParagraph p1 = doc.CreateParagraph();
        p1.Alignment = ParagraphAlignment.LEFT;
        p1.BorderBottom = Borders.Double;
        p1.BorderTop = Borders.Double;

        p1.BorderRight = Borders.Double;
        p1.BorderLeft = Borders.Double;
        p1.BorderBetween = Borders.Single;

        p1.VerticalAlignment = TextAlignment.TOP;

        XWPFRun r1 = p1.CreateRun();
        r1.SetText(location);
        r1.IsBold = false;
        r1.FontFamily = "Times New Roman";
        r1.FontSize = 15;
    }
}
