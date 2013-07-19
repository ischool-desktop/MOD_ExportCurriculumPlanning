using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExportCurriculumPlanning
{
    static public class tool
    {
        static public Dictionary<string, int> PopSort { get; set; }
        static public Dictionary<string, string> PopDic { get; set; }

        static public void Run()
        {
            PopSort = new Dictionary<string, int>();
            PopSort.Add("Ⅰ", 1);
            PopSort.Add("Ⅱ", 2);
            PopSort.Add("Ⅲ", 3);
            PopSort.Add("Ⅳ", 4);
            PopSort.Add("Ⅴ", 5);
            PopSort.Add("Ⅵ", 6);
            PopSort.Add("Ⅶ", 7);
            PopSort.Add("Ⅷ", 8);

            PopDic = new Dictionary<string, string>();
            PopDic.Add("1", "Ⅰ");
            PopDic.Add("2", "Ⅱ");
            PopDic.Add("3", "Ⅲ");
            PopDic.Add("4", "Ⅳ");
            PopDic.Add("5", "Ⅴ");
            PopDic.Add("6", "Ⅵ");
            PopDic.Add("7", "Ⅶ");
            PopDic.Add("8", "Ⅷ");
        }


        static public int SortSubjectLevel(string a1, string a2)
        {
            int b1 = 0;
            int b2 = 0;
            if (PopSort.ContainsKey(a1))
                b1 = PopSort[a1];
            if (PopSort.ContainsKey(a2))
                b2 = PopSort[a2];

            return b1.CompareTo(b2);
        }
    }
}
