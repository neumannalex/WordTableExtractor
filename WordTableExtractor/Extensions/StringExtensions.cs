using System;
using System.Collections.Generic;
using System.Text;

namespace WordTableExtractor.Extensions
{
    public static class StringExtensions
    {
        public static string Repeat(this string s, int n)
        {
            return new StringBuilder(s.Length * n)
                            .AppendJoin(s, new string[n + 1])
                            .ToString();
        }
    }
}
