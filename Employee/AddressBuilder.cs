using Fluid.Filters;

namespace Employee;

public static class AddressBuilder
{
    public static string Build(string folderPath, string fileName)
    {
        return folderPath + fileName;
    }
}