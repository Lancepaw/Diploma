using System;

namespace DataBaseIO
{
    public static class clsTextProcessor
    {
        //
        private static string[] Split(string text)
        {
            char[] separators = new char[] { '\r', '\a', '\v' };

            return text.Split(separators, StringSplitOptions.RemoveEmptyEntries);
        }

        //
        private static string[] SplitSKOB(string text)
        {
            char[] separators = new char[] { '(', ')' };
            return text.Split(separators, StringSplitOptions.RemoveEmptyEntries);
        }

        //
        public static string WordCellToString(string text)
        {
            string[] subStrings = Split(text);
            string total = string.Empty;
            foreach (var substring in subStrings)
            {
                int result;
                if (int.TryParse(substring, out result))
                {
                    total += result;
                }
                else
                {
                    total = substring;
                    break;
                }
            }
            return total;
        }

        //
        public static int WordCellToSum(string text)
        {
            string[] subStrings = Split(text);
            int total = 0;
            foreach (var substring in subStrings)
            {
                if (!substring.Contains("*") || !substring.Contains("("))
                {
                    int result;
                    if (int.TryParse(substring, out result))
                    {
                        total += result;
                    }
                }
            }
            return total;
        }

        //
        public static int WordCellToSumSKOB(string text)
        {
            string[] subStrings = Split(text);
            int total = 0;
            foreach (var substring in subStrings)
            {

                if (substring.Contains("(") && substring.Contains(")"))
                {
                    subStrings = SplitSKOB(substring);
                    foreach (var substringi in subStrings)
                    {

                        int result;
                        if (int.TryParse(substringi, out result))
                        {
                            total += result;
                        }
                    }
                }
            }
            return total;
        }
    }
}
