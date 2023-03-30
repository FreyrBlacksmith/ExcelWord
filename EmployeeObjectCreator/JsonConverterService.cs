using System.IO;
using NJsonSchema;
using NJsonSchema.CodeGeneration.CSharp;

namespace EmployeeObjectCreator;

public class JsonConverterService
{

    public JsonSchema? ConvertStringToSchema(string jsonString)
    {
        string json = File.ReadAllText(jsonString);
        return JsonSchema.FromSampleJson(json);
    }

    public bool TryGenerateClassFromSchema(JsonSchema classSchema, string outputFilePath)
    {
        var isCreted = false;
        if (classSchema != null)
        {
            var classGenerator = new CSharpGenerator(classSchema, new CSharpGeneratorSettings
            {
                ClassStyle = CSharpClassStyle.Poco,
            });
            var codeFile = classGenerator.GenerateFile();
            File.WriteAllText(outputFilePath, codeFile);
            isCreted= true;
        }
        return isCreted;
    }
}