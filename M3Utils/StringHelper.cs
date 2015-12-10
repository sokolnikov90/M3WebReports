using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace M3Utils
{
    public static class StringHelper
    {
        private static readonly Dictionary<string, string> Transliter = new Dictionary<string, string>
        {
            {"а", "a"},
            {"б", "b"},
            {"в", "v"},
            {"г", "g"},
            {"д", "d"},
            {"е", "ye"},
            {"ё", "ye"},
            {"ж", "zh"},
            {"з", "z"},
            {"и", "i"},
            {"й", "y"},
            {"к", "k"},
            {"л", "l"},
            {"м", "m"},
            {"н", "n"},
            {"о", "o"},
            {"п", "p"},
            {"р", "r"},
            {"с", "s"},
            {"т", "t"},
            {"у", "u"},
            {"ф", "f"},
            {"х", "kh"},
            {"ц", "ts"},
            {"ч", "ch"},
            {"ш", "sh"},
            {"щ", "shch"},
            {"ъ", ""},
            {"ы", "y"},
            {"ь", ""},
            {"э", "e"},
            {"ю", "yu"},
            {"я", "ya"},
            {"А", "A"},
            {"Б", "B"},
            {"В", "V"},
            {"Г", "G"},
            {"Д", "D"},
            {"Е", "Ye"},
            {"Ё", "Ye"},
            {"Ж", "Zh"},
            {"З", "Z"},
            {"И", "I"},
            {"Й", "J"},
            {"К", "K"},
            {"Л", "L"},
            {"М", "M"},
            {"Н", "N"},
            {"О", "O"},
            {"П", "P"},
            {"Р", "R"},
            {"С", "S"},
            {"Т", "T"},
            {"У", "U"},
            {"Ф", "F"},
            {"Х", "Kh"},
            {"Ц", "Ts"},
            {"Ч", "Ch"},
            {"Ш", "Sh"},
            {"Щ", "Shch"},
            {"Ъ", ""},
            {"Ы", "Y"},
            {"Ь", ""},
            {"Э", "E"},
            {"Ю", "Yu"},
            {"Я", "Ya"}
        };

        public static string GetTranslit(string sourceText)
        {
            StringBuilder ans = new StringBuilder();
            for (int i = 0; i < sourceText.Length; i++)
            {
                if (Transliter.ContainsKey(sourceText[i].ToString()))
                {
                    ans.Append(Transliter[sourceText[i].ToString()]);
                }
                else
                {
                    ans.Append(sourceText[i].ToString());
                }
            }
            return ans.ToString();
        }

        public static byte[] HexStr2ByteArr(string h)
        {
            if (h.Length % 2 != 0)
                throw new FormatException();

            List<byte> rv = new List<byte>();
            for (int i = 0; i < h.Length; i += 2)
            {
                string sub = h.Substring(i, 2);
                byte cc = Convert.ToByte(Convert.ToInt32(sub, 16));
                rv.Add(cc);
            }
            return rv.ToArray();
        }

        public static string HexToString(string hex)
        {
            string result = "";

            for (int i = 0; i < hex.Length; i = i + 2)
            {
                string subHex = hex.Substring(i, 2);

                result += Char.ConvertFromUtf32(Convert.ToInt32(subHex, 16));
            }

            return result;
        }

        // Замены непечатаемых символов в строке.
        public static string ReplaceNonprintingCharacters(string s)
        {
            char[] sCharArray = s.ToCharArray();

            for (int j = 0; j < sCharArray.Length; j++)
            {
                char ch = sCharArray[j];

                // непечатаемые символы
                // 11 | 12 | 14 | 15| | 27 | 28 | 29 | 30 | 31 | ...

                if (!((ch >= 161) || ((ch >= 32) && (ch <= 127))) )
                {
                    sCharArray[j] = '.';
                }
            }

            return new string(sCharArray);
        }

        // Удаление непечатаемых символов в строке.
        public static string RemoveNonprintingCharacters(string s)
        {
            if (s.Length == 0)
                return string.Empty;

            char[] sCharArray = s.ToCharArray();
            List<char> sCharList = new List<char>(s.Length);
            for (int j = 0; j < s.Length; j++)
            {
                char ch = sCharArray[j];
                if ((ch >=161) || ((ch >= 32) && (ch <= 127))) 
                {
                    sCharList.Add(ch);
                }
            }
            return new string(sCharList.ToArray());
        }

        public static IEnumerable<string> GenerateAtributes(string[] templates, int startIndex, int endIndex)
        {
            List<string> result = new List<string>();

            for (int i = 0; i < templates.Length; i++)
            {
                for (int j = startIndex; j <= endIndex; j++)
                {
                    result.Add(string.Format(templates[i], j));
                }
            }

            return result;
        }
    }
}
