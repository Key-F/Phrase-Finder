using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestLearning
{
    class helper
    {
        static public string clearURL(string badURL)
        {
            string goodURL = "";
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
    }
}
