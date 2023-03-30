using System;
using System.Collections.Generic;

namespace ExscelToWord.Helpers;

public static class Extensions
{
    public static string Filter(this string str, List<string> charsToRemove)
    {
        foreach (string c in charsToRemove)
        {
            str = str.Replace(c, String.Empty);
        }

        return str;
    }
}