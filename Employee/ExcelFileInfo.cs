namespace Employee;

public class ExcelFileInfo
{
   public IList<string> Header { get; set; }
   public List<IDictionary<string?, string>> TableData { get; set; }
   public ExcelFileInfo()
   {
       TableData = new List<IDictionary<string?, string>>();
   }
}