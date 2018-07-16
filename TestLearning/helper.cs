using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace TestLearning
{
    class helper
    {
        static public string clearURL(string badURL)
        {            
            if (badURL.StartsWith("https://"))
                badURL = badURL.Replace("https://", " ");
            if (badURL.StartsWith("http://"))
                badURL = badURL.Replace("http://", "");
            return badURL.Substring(0, badURL.IndexOf('/') + 1);
        }
        
        static public bool endsBad (string URL)
        {
            return (URL.EndsWith(".pdf") || URL.EndsWith(".xls") || URL.EndsWith(".xlsx") || URL.EndsWith(".doc") || URL.EndsWith(".docx") || URL.EndsWith(".zip") 
                || URL.EndsWith(".rtf") || URL.EndsWith(".rar"));
        }

        static public int keywordsChecker(String Text, String keywords)
        {
            int count = 0;
            String[] ary = keywords.Split(',');

            for (int i = 0; i < ary.Length - 1; i++)
            {
                while (ary.Length > 0 && char.IsWhiteSpace(ary[i][0]))
                {
                    ary[i] = ary[i].Remove(0, 1);
                }
                if (Text.Contains(ary[i].ToLower()))
                {
                    count++;
                }
            }
            return count;
        }

        static public string noHtml(string html)
        {
            string newhtml = Regex.Replace(html, "<[^>]+>", string.Empty);
            return newhtml;
        }
    }
}
