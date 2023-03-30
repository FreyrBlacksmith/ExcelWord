using System.IO;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Windows.Input;
using Employee;
using Newtonsoft.Json.Linq;
using UiBase;

namespace EmployeeObjectCreator;

public class EmployeeCreatorViewModel : BaseNotifyObject
{
    public ICommand CreateEmployeePresetClassCommand { get; }
    public ICommand AddNewFieldToPresetCommand { get; }
    private JsonDocument schema;
    private string presetName;

    public string PresetName
    {
        get => presetName;
        set => SetProperty(ref presetName, value);
    }

    public EmployeeCreatorViewModel()
    {
        CreateEmployeePresetClassCommand = new Command(CreateEmployeePreset);
        AddNewFieldToPresetCommand = new Command(AddNewFieldToPreset);
    }

    private void AddNewFieldToPreset()
    {
        if (schema.ToString() == string.Empty)
        {
            //schema = new JsonDocument();
            var obj = new JsonObject();
            obj["name"] = PresetName;

        }
        if (PresetName!= string.Empty)
        {
           
        }
    }
    private void CreateEmployeePreset()
    {
        var converter = new JsonConverterService();
        var shemaObject=converter.ConvertStringToSchema(schema.ToString());
        if (shemaObject != null)
        {
            converter.TryGenerateClassFromSchema(shemaObject,
                AddressBuilder.Build(Directory.GetCurrentDirectory() + "Employee  presets", PresetName));
        }
        
    }
}