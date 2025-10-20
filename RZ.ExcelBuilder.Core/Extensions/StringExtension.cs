using System;
using System.Linq;

namespace RZ.ExcelBuilder.Core.Extensions
{
    internal static class StringExtension
    {
        public static string FormatRut(this string rut)
        {
            rut = new string(rut.Where(c => char.IsDigit(c) || c == 'K' || c == 'k').ToArray());

            if (rut.Length < 2)
                return rut;

            string body = rut[..^1];
            string dv = rut[^1].ToString().ToUpper();

            string formattedBody = string.Join(".",
                Enumerable.Range(0, (body.Length + 2) / 3)
                .Select(i => new string([.. body.Reverse().Skip(i * 3).Take(3).Reverse()]))
                .Reverse());

            return $"{formattedBody}-{dv}";
        }

    }
}
