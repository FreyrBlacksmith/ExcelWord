using Employee;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace ExscelToWord.Service;

public class ExcelService
{
    private XSSFWorkbook workbook;


    public void SetupService(string excelFilePath)
    {
        using (var stream = new FileStream(excelFilePath, FileMode.Open, FileAccess.Read))
        {
            workbook = new XSSFWorkbook(stream);
        }
    }

    public IEnumerable<string> GetAllSheets()
    {
        for (int i = 0; i < workbook.NumberOfSheets; i++)
        {
            ISheet sheet = workbook.GetSheetAt(i);
            yield return sheet.SheetName;
        }
    }

    /// <summary>  
    /// Method to Get All the Records from Excel  
    /// </summary>  
    /// <returns></returns>  
    public ExcelFileInfo ReadRecordsFromExcel(string sheetName)
    {
        var result = new ExcelFileInfo();

        try
        {
            var worksheet = workbook.GetSheet(sheetName);
            result.Header = new List<string>(worksheet.GetRow(0).Cells.Select(x => x?.ToString()));
            var emptyCellHeaderCounter = 1;
            for (int rowIndex = 0; rowIndex < worksheet.LastRowNum; rowIndex++)
            {
                var row = worksheet.GetRow(rowIndex);

                var rowData = new Dictionary<string, string>();
                if (row != null)
                {

                    for (int columnIndex = 0; columnIndex < row.LastCellNum; columnIndex++)
                    {

                        var value = row.GetCell(columnIndex)?.ToString();
                        var header = string.Empty;
                        if (result.Header.Count <= columnIndex)
                        {
                            var unnamedHeaderName = $"Unnamed{emptyCellHeaderCounter}";
                            header = unnamedHeaderName;
                            result.Header.Add(unnamedHeaderName);
                            emptyCellHeaderCounter++;
                        }
                        else
                        {
                            header = result.Header[columnIndex];
                        }
                        if (header != null)
                            rowData[header] = value;
                        
                    }
                    result.TableData.Add(rowData);
                }
            }

            result.TableData = result.TableData.Where(x => x.Values.First() != string.Empty).ToList();
            result=UpdateAllToMax(result);
        }
        catch (Exception e)
        {
            Console.WriteLine(e);
            throw;
        }
        return result;
    }

    private ExcelFileInfo UpdateAllToMax(ExcelFileInfo result)
    {
        var headersCount = result.Header.Count;
        foreach (var tableRow in result.TableData)
        {
            if (tableRow.Keys.Count < headersCount)
            {
                for (int i = tableRow.Keys.Count; i < result.Header.Count; i++)
                {
                    tableRow.Add(result.Header[i], string.Empty);
                }
            }
        }

        return result;
    }
}

