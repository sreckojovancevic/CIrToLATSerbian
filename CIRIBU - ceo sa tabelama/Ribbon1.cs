using Microsoft.Office.Tools.Ribbon;
using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace CIRIBU
{
    public partial class Ribbon1 : RibbonBase, Microsoft.Office.Core.IRibbonExtensibility
    {
        public Ribbon1()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("CIRIBU.Ribbon1.xml");
        }

        private static string GetResourceText(string resourceName)
        {
            var asm = typeof(Ribbon1).Assembly;
            using (var stream = asm.GetManifestResourceStream(resourceName))
            {
                if (stream != null)
                {
                    using (var reader = new System.IO.StreamReader(stream))
                    {
                        return reader.ReadToEnd();
                    }
                }
            }
            return null;
        }

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            // Initialization code here
        }

        private void Button1_Click(object sender, RibbonControlEventArgs e)
        {
            Document doc = Globals.ThisAddIn.Application.ActiveDocument;
            Selection selection = Globals.ThisAddIn.Application.Selection;

            if (selection != null && selection.Range.Text != "")
            {
                // Check if the selection contains tables
                if (selection.Tables.Count > 0)
                {
                    foreach (Table table in selection.Tables)
                    {
                        foreach (Row row in table.Rows)
                        {
                            foreach (Cell cell in row.Cells)
                            {
                                string cellText = cell.Range.Text;
                                string convertedText = TransliterateToCyrillic(cellText);
                                cell.Range.Text = convertedText;
                            }
                        }
                    }
                }
                else
                {
                    // Transliterate the selected text
                    string selectedText = selection.Text;
                    string convertedText = TransliterateToCyrillic(selectedText);
                    selection.Text = convertedText;
                }
            }
        }

        private void Button2_Click(object sender, RibbonControlEventArgs e)
        {
            Document doc = Globals.ThisAddIn.Application.ActiveDocument;
            Selection selection = Globals.ThisAddIn.Application.Selection;

            if (selection != null && selection.Range.Text != "")
            {
                // Check if the selection contains tables
                if (selection.Tables.Count > 0)
                {
                    foreach (Table table in selection.Tables)
                    {
                        foreach (Row row in table.Rows)
                        {
                            foreach (Cell cell in row.Cells)
                            {
                                string cellText = cell.Range.Text;
                                string convertedText = TransliterateToLatin(cellText);
                                cell.Range.Text = convertedText;
                            }
                        }
                    }
                }
                else
                {
                    // Transliterate the selected text
                    string selectedText = selection.Text;
                    string convertedText = TransliterateToLatin(selectedText);
                    selection.Text = convertedText;
                }
            }
        }

        private string TransliterateToCyrillic(string input)
        {
            var latinToCyrillic = new Dictionary<string, string>
            {
                {"Lj", "Љ"}, {"Nj", "Њ"}, {"Dž", "Џ"}, {"A", "А"}, {"B", "Б"},
                {"V", "В"}, {"G", "Г"}, {"D", "Д"}, {"Đ", "Ђ"}, {"E", "Е"},
                {"Ž", "Ж"}, {"Z", "З"}, {"I", "И"}, {"J", "Ј"}, {"K", "К"},
                {"L", "Л"}, {"M", "М"}, {"N", "Н"}, {"O", "О"}, {"P", "П"},
                {"R", "Р"}, {"S", "С"}, {"T", "Т"}, {"Ć", "Ћ"}, {"U", "У"},
                {"F", "Ф"}, {"H", "Х"}, {"C", "Ц"}, {"Č", "Ч"}, {"Š", "Ш"},
                {"lj", "љ"}, {"nj", "њ"}, {"dž", "џ"}, {"a", "а"}, {"b", "б"},
                {"v", "в"}, {"g", "г"}, {"d", "д"}, {"đ", "ђ"}, {"e", "е"},
                {"ž", "ж"}, {"z", "з"}, {"i", "и"}, {"j", "ј"}, {"k", "к"},
                {"l", "л"}, {"m", "м"}, {"n", "н"}, {"o", "о"}, {"p", "п"},
                {"r", "р"}, {"s", "с"}, {"t", "т"}, {"ć", "ћ"}, {"u", "у"},
                {"f", "ф"}, {"h", "х"}, {"c", "ц"}, {"č", "ч"}, {"š", "ш"}
            };
            return Transliterate(input, latinToCyrillic);
        }

        private string TransliterateToLatin(string input)
        {
            var cyrillicToLatin = new Dictionary<string, string>
            {
                {"Љ", "Lj"}, {"Њ", "Nj"}, {"Џ", "Dž"}, {"А", "A"}, {"Б", "B"},
                {"В", "V"}, {"Г", "G"}, {"Д", "D"}, {"Ђ", "Đ"}, {"Е", "E"},
                {"Ж", "Ž"}, {"З", "Z"}, {"И", "I"}, {"Ј", "J"}, {"К", "K"},
                {"Л", "L"}, {"М", "M"}, {"Н", "N"}, {"О", "O"}, {"П", "P"},
                {"Р", "R"}, {"С", "S"}, {"Т", "T"}, {"Ћ", "Ć"}, {"У", "U"},
                {"Ф", "F"}, {"Х", "H"}, {"Ц", "C"}, {"Ч", "Č"}, {"Ш", "Š"},
                {"љ", "lj"}, {"њ", "nj"}, {"џ", "dž"}, {"а", "a"}, {"б", "b"},
                {"в", "v"}, {"г", "g"}, {"д", "d"}, {"ђ", "đ"}, {"е", "e"},
                {"ж", "ž"}, {"з", "z"}, {"и", "i"}, {"ј", "j"}, {"к", "k"},
                {"л", "l"}, {"м", "m"}, {"н", "n"}, {"о", "o"}, {"п", "p"},
                {"р", "r"}, {"с", "s"}, {"т", "t"}, {"ћ", "ć"}, {"у", "u"},
                {"ф", "f"}, {"х", "h"}, {"ц", "c"}, {"ч", "č"}, {"ш", "š"}
            };
            return Transliterate(input, cyrillicToLatin);
        }

        private string Transliterate(string input, Dictionary<string, string> translitMap)
        {
            StringBuilder result = new StringBuilder();
            int i = 0;

            while (i < input.Length)
            {
                bool matched = false;

                // Check for multi-character sequences
                foreach (var kvp in translitMap)
                {
                    if (input.Substring(i).StartsWith(kvp.Key))
                    {
                        result.Append(kvp.Value);
                        i += kvp.Key.Length;
                        matched = true;
                        break;
                    }
                }

                // If no multi-character sequence matched, process single character
                if (!matched)
                {
                    result.Append(input[i]);
                    i++;
                }
            }

            return result.ToString();
        }
    }
}
