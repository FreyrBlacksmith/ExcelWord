using Employee;
using ExscelToWord.Helpers;
using ExscelToWord.Service;
using ExscelToWord.WordService.Templates;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Input;
using UiBase;
using Task = System.Threading.Tasks.Task;

namespace ExscelToWord.ViewModels;

public class ExcelConvertViewModel : BaseNotifyObject
{
    private string excelFilePath;
    private string excelSheetName;
    private bool isSheetSelected;
    private IEnumerable<string> excelSheets;
    private readonly ExcelService excelService;
    private ObservableCollection<TableColumnInfo> listOfFields;
    private ExcelFileInfo tableInfo;
    private string wordFileName;
    private const string assigmentPattern = @"\b(должность)\b";
    private const string companyPattern = @"\b(название предприятия|организации)\b";
    private const string compFormat = @"(.*)\b(ООО|ОАО|ЗАО|НКО|ТСЖ|ОДО|АО|ПАО|НПО|ИП)\b(.*)";

    public ICommand LoadExcelCommand { get; }
    public AsyncCommand LoadExcelColumnsCommand { get; }
    public AsyncCommand FillWordCommand { get; }
    public AsyncCommand LoadWordCommand { get; }
    public string ExcelFilePath
    {
        get => excelFilePath;
        set => SetProperty(ref excelFilePath, value);
    }

    public IEnumerable<string> ExcelSheets
    {
        get => excelSheets;
        set => SetProperty(ref excelSheets, value);
    }

    public string ExcelSheet
    {
        get => excelSheetName;
        set
        {
            if (!string.IsNullOrEmpty(value))
            {
                isSheetSelected = true;
                UpdateCommands();
            }
            SetProperty(ref excelSheetName, value);
        }
    }

    public string WordFileName
    {
        get => wordFileName;
        set => SetProperty(ref wordFileName, value);
    }
    public ObservableCollection<TableColumnInfo> ListOfFields
    {
        get => listOfFields;
        set => SetProperty(ref listOfFields, value);
    }
    public ExcelConvertViewModel()
    {
        isSheetSelected = false;
        LoadExcelCommand = new Command(LoadExcelName);
        LoadExcelColumnsCommand = new AsyncCommand(LoadExcelColumns, IsSheetSelected);
        FillWordCommand = new AsyncCommand(FillWord, IsSheetSelected);
        LoadWordCommand = new AsyncCommand(LoadWord, IsSheetSelected);
        excelService = new ExcelService();
        ExcelSheets = new List<string>();
        ListOfFields = new();
    }
    private bool IsSheetSelected()
    {
        return isSheetSelected;
    }
    private async Task LoadWord()
    {
        if (IsSheetSelected())
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            if (openFileDialog.ShowDialog() == true)
                WordFileName = openFileDialog.FileName;
        }
    }
    private void LoadExcelName()
    {
        OpenFileDialog openFileDialog = new OpenFileDialog();
        if (openFileDialog.ShowDialog() == true)
            ExcelFilePath = openFileDialog.FileName;
        excelService.SetupService(ExcelFilePath);
        ExcelSheets = excelService.GetAllSheets();
    }
    private void UpdateCommands()
    {
        LoadExcelColumnsCommand.RaiseCanExecuteChanged();
        FillWordCommand.RaiseCanExecuteChanged();
        LoadWordCommand.RaiseCanExecuteChanged();
    }
    private async Task LoadExcelColumns()
    {
        if (IsSheetSelected())
        {
            if (string.IsNullOrEmpty(ExcelSheet))
            {
                MessageBox.Show("Пожалуйста, выберите сначала необзодимую вкладку документа");
            }
            else
            {
                ListOfFields.Clear();
                tableInfo = excelService.ReadRecordsFromExcel(ExcelSheet);

                var columnIndex = 0;
                foreach (var columnName in tableInfo.Header)
                {
                    ListOfFields.Add(new TableColumnInfo()
                    { Name = columnName, IsSelected = false, Index = columnIndex });
                    columnIndex++;
                }
            }
        }
    }

    private async Task FillWord()
    {
        if (IsSheetSelected())
        {
            var template = new ConferenceTemplate()
            {
                Label = "",
                Description =
                    "",
                Location = "",
                Participants = LoadTableContent()
            };
            var res = WordService.WordService.LoadAndFillWordDocument(WordFileName, template);
            MessageBox.Show($"{res.Message}");
        }
    }

    private IEnumerable<IEnumerable<string>> LoadTableContent()
    {
        var result = new List<List<string>>();

        foreach (var tableRow in tableInfo.TableData.Skip(1))
        {
            var rowInfo = new List<string>();
            foreach (var columnInfo in ListOfFields.Where(x => x.IsSelected))
            {
                // if (columnInfo.Name == "ФИО участника визита (полностью)" || columnInfo.Name == "ФИО участника" || columnInfo.Name == "ФИО")
                if (columnInfo.Name.IndexOf("ФИО") >= 0)
                {
                    //Add opportunity to setup separator
                    rowInfo.Add(ConvertFioToUpperCase(tableRow[columnInfo.Name].Split(' ')));
                }else if (Regex.IsMatch(columnInfo.Name, assigmentPattern, RegexOptions.IgnoreCase))
                {
                    var organizationName = ListOfFields.First(item => Regex.IsMatch(item.Name, companyPattern));
                    var correctedCompName = tableRow[organizationName.Name];
                    Match match = Regex.Match(tableRow[organizationName.Name], compFormat);
                    if (match.Success)
                    {
                        string opf = match.Groups[2].Value;
                        string rest = match.Groups[1].Value + match.Groups[3].Value;
                        string output = opf + " " + rest;
                        correctedCompName = output;
                    }

                    rowInfo.Add("-" + tableRow[columnInfo.Name] + " "+ correctedCompName);
                }
                else if (Regex.IsMatch(columnInfo.Name, companyPattern, RegexOptions.IgnoreCase))
                {
                    Match match = Regex.Match(tableRow[columnInfo.Name], compFormat);
                    if (match.Success)
                    {
                        string opf = match.Groups[2].Value;
                        string rest = match.Groups[1].Value + match.Groups[3].Value;
                        string output = opf + " " + rest;
                        rowInfo.Add(output);
                    }

                }
                else
                {
                    rowInfo.Add(tableRow[columnInfo.Name]);
                }

            }
            result.Add(rowInfo);
        }


        return result;
    }



    private string ConvertFioToUpperCase(string[] fio)
    {
        var result = fio[0].ToUpper();
        for (int i = 1; i < fio.Length; i++)
        {
            result += " " + fio[i] ;
        }

        var words = result.Split(' ');
        return $"{words[0]}\n{string.Join("  ", words.Skip(1))}";
    }
}