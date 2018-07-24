using HtmlAgilityPack;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace TestLearning
{
    public partial class Form1 : Form
    {
        int SubtokenLength = 2;
        Microsoft.Office.Interop.Excel.Application ObjExcel;
        public int MinWordLength { get; private set; }

        public Form1()
        {
            InitializeComponent();
        }                                      

                   
        private void button2_Click(object sender, EventArgs e)  // NOHTML
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {


                //Создаём приложение.
                Microsoft.Office.Interop.Excel.Application ObjExcel = new Microsoft.Office.Interop.Excel.Application();
                //Открываем книгу.       
                try
                {
                    Microsoft.Office.Interop.Excel.Workbook TestBook = ObjExcel.Workbooks.Open(openFileDialog1.FileName, 0, false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
                }
                catch
                {
                    MessageBox.Show("Файл не доступен");
                    ObjExcel.Quit();

                }
                Microsoft.Office.Interop.Excel.Workbook ObjWorkBook = ObjExcel.Workbooks.Open(openFileDialog1.FileName, 0, false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
                //
                //Выбираем таблицу(лист).
                Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet;
                ObjWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBook.Sheets[1];

                

                for (int i = Convert.ToInt32(firstN.Text); i <= Convert.ToInt32(lastN.Text); i++)
                {
                    Excel.Range range = ObjWorkSheet.get_Range(textBox3.Text + i, textBox3.Text + i); // Где искать текст
                  

                    string text;
                    //string keyword = textBox4.Text;
                    string keyword = richTextBox1.Text;

                    if (range.Text != null)
                    {
                        text = range.Text.ToLower();
                        range.Cells.Value = helper.noHtml(text);
                        
                    }                    
                }              
                MessageBox.Show("Все");                
                ObjExcel.Quit();  //Удаляем приложение (выходим из экселя) - а то будет висеть в процессах!
            }
        }

       

        private void button3_Click(object sender, EventArgs e) // Normalize
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {


                //Создаём приложение.
                ObjExcel = new Microsoft.Office.Interop.Excel.Application();
                //Открываем книгу.       
                try
                {
                    Microsoft.Office.Interop.Excel.Workbook TestBook = ObjExcel.Workbooks.Open(openFileDialog1.FileName, 0, false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
                }
                catch
                {
                    MessageBox.Show("Файл не доступен");
                    ObjExcel.Quit();

                }
                Microsoft.Office.Interop.Excel.Workbook ObjWorkBook = ObjExcel.Workbooks.Open(openFileDialog1.FileName, 0, false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
                //
                //Выбираем таблицу(лист).
                Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet;
                ObjWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBook.Sheets[1];



                int kolstrok = 0; // Количество строк в таблице, не прошедших проверки
                                  // try
                                  // {

                for (int i = Convert.ToInt32(firstN.Text); i <= Convert.ToInt32(lastN.Text); i++)
                {
                    Excel.Range range = ObjWorkSheet.get_Range(textBox4.Text + i, textBox4.Text + i); // Где искать текст
                    

                    string text;
                    //string keyword = textBox4.Text;
                    

                    if (range.Text != null)
                    {
                        text = range.Text.ToLower();
                        range.Cells.Value = NormalizeSentence(text);

                    }
                }
                MessageBox.Show("Все");
                ObjExcel.Quit();  //Удаляем приложение (выходим из экселя) - а то будет висеть в процессах!
            }
        }
       

        static private string NormalizeSentence(string sentence)
        {
            var resultContainer = new StringBuilder(100);
            var lowerSentece = sentence.ToLower();
            foreach (var c in lowerSentece)
            {
                if (IsNormalChar(c))
                {
                    resultContainer.Append(c);
                }
            }

            return resultContainer.ToString();
        }

        static private bool IsNormalChar(char c)
        {
            return char.IsLetterOrDigit(c) || c == ' ';
        }
                         
        

        private void button6_Click(object sender, EventArgs e) // Удаление лишнего
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                int tekind = Convert.ToInt32(firstN.Text); // текущий индекс
                //Создаём приложение.
                Microsoft.Office.Interop.Excel.Application ObjExcel = new Microsoft.Office.Interop.Excel.Application();
                //Открываем книгу.       
                try
                {
                    Microsoft.Office.Interop.Excel.Workbook TestBook = ObjExcel.Workbooks.Open(openFileDialog1.FileName, 0, false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
                }
                catch
                {
                    MessageBox.Show("Файл не доступен");
                    ObjExcel.Quit();

                }
                Microsoft.Office.Interop.Excel.Workbook ObjWorkBook = ObjExcel.Workbooks.Open(openFileDialog1.FileName, 0, false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
                //
                //Выбираем таблицу(лист).
                Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet;
                ObjWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBook.Sheets[1];
                Excel.Range range1 = ObjWorkSheet.get_Range("C" + tekind, "C" + tekind); // Начальное значение ссылки
                string goodURL = helper.clearURL(range1.Text);
                //Excel.Range range;
                
                int countodinakovih = 1;
                int i = tekind + 1;
                string goodURL2 = "";
                do
                {               
                    Excel.Range range = ObjWorkSheet.get_Range("C" + i, "C" + i);
                    goodURL2 = helper.clearURL(range.Text);
                    countodinakovih++;
                    i++;
                }
                while (goodURL2 == goodURL);
                MessageBox.Show(countodinakovih.ToString());
                int i2 = 0; // индекс по тексту
                int maxlen = 0; // максимальная длина совпадающих слов
                int len = 0; // длина совпалающих слов
                
                string maxtext = "";
                int[] bad = new int[Convert.ToInt32(lastN.Text)]; // будем отмечать 1, где есть вначале что-то плохое
                string deletetext = "";
                char letterfromfirststring = '0'; // ыы
                char letterfromsecondstring = '0';
                for (int i1 = tekind; i1 < countodinakovih; i1++)
                {
                    
                    do
                    {
                        len++;                        
                        letterfromfirststring = ObjWorkSheet.get_Range("Q" + i1, "Q" + i1).Text[i2];
                        letterfromsecondstring = ObjWorkSheet.get_Range("Q" + (i1 + 1).ToString(), "Q" + (i1 + 1).ToString()).Text[i2];
                        if (letterfromfirststring == letterfromsecondstring)
                            deletetext = deletetext + (ObjWorkSheet.get_Range("Q" + i1, "Q" + i1).Text[i2]);
                        i2++;
                    }
                    while (letterfromfirststring == letterfromsecondstring); // Сравниваем по словам

                    if (len >= maxtext.Length)
                    {
                        maxtext = deletetext;                        
                        bad[i1] = 1; // Помечаем, какие строки имеют такое-же название
                        bad[i1 + 1] = 1;
                    }
                    len = 0;
                    deletetext = "";
                    i2 = 0;
                } 
                
                MessageBox.Show(maxtext);
                ObjExcel.Quit();  //Удаляем приложение (выходим из экселя) - а то будет висеть в процессах!
            }
        }

        private void button7_Click(object sender, EventArgs e) // parse again
        {
            string[] spec = { "&larr;","&rarr", "&gt;", "&rdquo;", "&nbsp;", "&raquo;", "&laquo;", "&ndash;", "&qt;", "/n", "/t", "&mdash;", "&quot;" }; // Спецсимволы, которые нужно убрать
            string selector = "";
            string URL = "";
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {

                //Создаём приложение.
                ObjExcel = new Microsoft.Office.Interop.Excel.Application();
                //Открываем книгу.       
                try
                {
                    //Microsoft.Office.Interop.Excel.Workbook TestBook = ObjExcel.Workbooks.Open(openFileDialog1.FileName, 0, false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
                }
                catch
                {
                    MessageBox.Show("Файл не доступен");
                    ObjExcel.Quit();
                    return;

                }
                Microsoft.Office.Interop.Excel.Workbook ObjWorkBook = ObjExcel.Workbooks.Open(openFileDialog1.FileName, 0, false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
                //
                //Выбираем таблицу(лист).
                Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet;
                ObjWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBook.Sheets[1];

                for (int i = Convert.ToInt32(firstN.Text); i < Convert.ToInt32(lastN.Text); i++)
                {
                    Excel.Range range = ObjWorkSheet.get_Range("C" + i, "C" + i); // Где искать текст
                                                                                  //URL = range.Text.ToLower();
                    URL = range.Text;


                    //string URL = "https://ohranatruda.ru/news/896/166764/"; // ploho
                    //string URL = "http://baltic.transneft.ru/press/news/?id=45463"; // norm 
                    //string URL = "http://diascan.transneft.ru/press/news/?id=44797"; // ploho
                    //string URL = "http://rpn.gov.ru/news/22052017-v-moskve-startovala-nedelya-nauki-tehnologiy-i-innovaciy-geodata"; // ochen' ploho
                    //string URL = "http://rpn.gov.ru/news/baltiysko-arkticheskim-morskim-upravleniem-rosprirodnadzora-proveden-reydovyy-osmotr-uchastka"; // super ploho
                    //string URL = "http://rpn.gov.ru/news/baltiysko-arkticheskim-morskim-upravleniem-rosprirodnadzora-proveden-reydovyy-osmotr-uchastka"; // super ploho, и заголовок есть и инфа
                    //string URL = "http://vniiecology.ru/index.php/8-news/309-sapsany-v-cheremushkakh"; // и заголовок есть и инфа
                    //string URL = "http://vniiecology.ru/index.php/8-news/251-v-timiryazevke-rasskazali-o-vosstanovlenii-sapsana-v-moskve"; // infa only, ol ne rabotaet
                    //string URL = "http://voda.mnr.gov.ru/news/detail.php?ID=115796"; // +-
                    //string URL = "http://www.ecoindustry.ru/news/company/view/49739.html"; // Тут нужно удалять конец, тк часть текста ETO GOVNO
                    //URL = "http://www.gosnadzor.ru/news/64/2297/";
                    //URL = "http://www.mpe-sro.ru/news_3046.html";
                    //URL = "http://www.oaontc.ru/news/191";
                    //URL = "http://www.risk-news.ru/news/rostekhnadzor_otmenyaet_peregruppirovku_opasnykh_proizvodstvennykh_obektov/";
                    //URL = "http://www.vestipb.ru/indnews7952.html";
                    //URL = "https://www.orfi.ru/press/news/2017/?n=2017070601"; // idealno
                    //URL = "http://www.ecoindustry.ru/news/view/52792.html";
                    //URL = "https://www.gost.ru/portal/gost/home/presscenter/news?portal:isSecure=true&navigationalstate=JBPNS_rO0ABXcyAAZhY3Rpb24AAAABAA5zaW5nbGVOZXdzVmlldwACaWQAAAABAAM4NTIAB19fRU9GX18*&portal:componentId=88beae40-0e16-414c-b176-d0ab5de82e16";
                    //URL = "https://www.gost.ru/portal/gost/home/presscenter/news?portal:isSecure=true&navigationalstate=JBPNS_rO0ABXcxAAZhY3Rpb24AAAABAA5zaW5nbGVOZXdzVmlldwACaWQAAAABAAI0NgAHX19FT0ZfXw**&portal:componentId=88beae40-0e16-414c-b176-d0ab5de82e16";
                    HtmlWeb web = new HtmlWeb();
                    HtmlAgilityPack.HtmlDocument doc = null;

                    if (URL.Contains("ohranatruda.ru") && !helper.endsBad(URL))
                    {
                        selector = "//div[@class='ot-news-detail-text']";
                        web.OverrideEncoding = Encoding.GetEncoding(1251); // meta charset

                    }
                    else
                    if (URL.Contains("safety.ru") && !helper.endsBad(URL))
                    {
                        selector = "//div[@class='field-item even']";
                        web.OverrideEncoding = Encoding.UTF8;

                    }                    
                    else
                    if (URL.Contains("transneft.ru") && !helper.endsBad(URL))
                    {
                        selector = "//div[@class='content']";
                        web.OverrideEncoding = Encoding.UTF8;
                    }
                    else
                    if (URL.Contains("nbpo.ru") && !helper.endsBad(URL))
                    {
                        selector = "//div[@class='post_content']";
                        web.OverrideEncoding = Encoding.UTF8;
                    }
                    else
                    if (URL.Contains("rpn.gov.ru") && !helper.endsBad(URL))
                    {
                        selector = "//div[@property='content:encoded']";
                        web.OverrideEncoding = Encoding.UTF8;
                    }
                    else
                    if (URL.Contains("mnr.gov.ru") && !helper.endsBad(URL))
                    {
                        //selector = "//div[@class='news-detail']";
                        selector = "//div[@class='default-text-block']";
                        
                        web.OverrideEncoding = Encoding.UTF8;
                    }
                    else
                    if (URL.Contains("gosnadzor.ru") && !helper.endsBad(URL))
                    {
                        selector = "(//*[@class='news-detail'])";
                        web.OverrideEncoding = Encoding.UTF8;
                    }                                       
                    

                    else
                    if (URL.Contains("tehnoprogress.ru") && !helper.endsBad(URL))
                    {
                        selector = "//div[@class='content-column']//p";
                        web.OverrideEncoding = Encoding.UTF8;
                        doc = web.Load(URL);
                        var trynode = doc.DocumentNode.SelectSingleNode(selector);
                        var nodes = doc.DocumentNode.SelectNodes(selector);
                        string hh = "";
                        for (int ui = 0; ui < nodes.Count; ui++) 
                            hh = hh + nodes[ui].InnerText;

                        hh = helper.noHtml(hh);

                        foreach (string sl in spec)
                            hh = hh.Replace(sl, " ");
                        //MessageBox.Show(hh);
                        ObjWorkSheet.Cells[i, 6] = hh;
                        continue;
                    }
                    else
                    if (URL.Contains("risk-news.ru") && !helper.endsBad(URL)) // тут почти все норм с выгрузкой
                    {
                        try
                        {
                            ObjWorkSheet.Cells[i, 6] = ObjWorkSheet.Cells[i, 6].Text.Replace("С полным текстом документа можно ознакомиться в разделе «Официально»", " ");
                        }
                        catch { }
                            continue;
                    }
                    else
                    
                if (URL.Contains("oaontc.ru") && !helper.endsBad(URL))
                    {
                        ObjWorkSheet.Cells[i, 6] = ObjWorkSheet.Cells[i, 6].Text.Replace("К списку новостей", " ");
                        continue;
                    }
                    else
                    if (URL.Contains("www.mpe-sro.ru") && !helper.endsBad(URL)) // тут почти все норм с выгрузкой
                    {
                        ObjWorkSheet.Cells[i, 6] = ObjWorkSheet.Cells[i, 6].Text.Replace("Другие новости", " ");
                        continue;
                    }
                    else
                    if (URL.Contains("vernadsky.ru") && !helper.endsBad(URL)) // тут почти все норм с выгрузкой
                    {
                        foreach (string sl in spec)
                            ObjWorkSheet.Cells[i, 6] = ObjWorkSheet.Cells[i, 6].Text.Replace(sl, " ");                       
                        continue;
                    }
                    else
                    if (URL.Contains("ronktd.ru") && !helper.endsBad(URL))
                    {
                        selector = "//div[@class='news-detail']";
                        web.OverrideEncoding = Encoding.UTF8;
                    }
                    else
                    if (URL.Contains("vestipb.ru") && !helper.endsBad(URL))
                    {
                        selector = "//td[2] //p";
                        web.OverrideEncoding = Encoding.GetEncoding(1251); // meta charset
                        doc = web.Load(URL);
                        var trynode = doc.DocumentNode.SelectSingleNode(selector);
                        var nodes = doc.DocumentNode.SelectNodes(selector);
                        string hh = "";
                        for (int ui = 0; ui < nodes.Count; ui++)
                            hh = hh + nodes[ui].InnerText;


                        try
                        {
                            hh = hh.Substring(0, hh.LastIndexOf("Информационный портал")); //ыыы !!!!!!!!!!!!
                        }
                        catch { }
                        hh = helper.noHtml(hh);

                        foreach (string sl in spec)
                            hh = hh.Replace(sl, " ");
                        //MessageBox.Show(hh);
                        ObjWorkSheet.Cells[i, 6] = hh;
                        continue;

                    }
                    else
                    if (URL.Contains("ecolawyer.ru") && !helper.endsBad(URL))
                    {
                        selector = "//td[@class='content']/text()";
                        web.OverrideEncoding = Encoding.GetEncoding(1251); // meta charset
                        doc = web.Load(URL);
                        var trynode = doc.DocumentNode.SelectSingleNode(selector);
                        var nodes = doc.DocumentNode.SelectNodes(selector);
                        string hh = "";
                        for (int ui = 0; ui < nodes.Count; ui++)
                            hh = hh + nodes[ui].InnerText;


                        
                        hh = helper.noHtml(hh);

                        foreach (string sl in spec)
                            hh = hh.Replace(sl, " ");
                        //MessageBox.Show(hh);
                        ObjWorkSheet.Cells[i, 6] = hh;
                        ObjExcel.ActiveWorkbook.Save();
                        continue;

                    }
                    else
                    if (URL.Contains("orfi.ru") && !helper.endsBad(URL))
                    {
                        selector = "//div[@class='content col-sm-8']//p";
                        web.OverrideEncoding = Encoding.UTF8;
                        doc = web.Load(URL);
                        
                        var nodes = doc.DocumentNode.SelectNodes(selector);
                        string hh = "";
                        for (int ui = 0; ui < nodes.Count - 1; ui++)
                            hh = hh + nodes[ui].InnerText;



                        hh = helper.noHtml(hh);

                        foreach (string sl in spec)
                            hh = hh.Replace(sl, " ");
                        //MessageBox.Show(hh);
                        ObjWorkSheet.Cells[i, 6] = hh;
                        ObjExcel.ActiveWorkbook.Save();
                        continue;
                    }
                    else
                    if (URL.Contains("gost.ru") && !helper.endsBad(URL))
                    {
                        selector = "//div[@class='news-view__content']//p";
                        web.OverrideEncoding = Encoding.UTF8;
                        doc = web.Load(URL);
                        
                        var nodes = doc.DocumentNode.SelectNodes(selector);
                        string hh = "";
                        for (int ui = 1; ui < nodes.Count; ui++)
                        if (nodes != null)
                            hh = hh + nodes[ui].InnerText;
                        else continue;


                        hh = helper.noHtml(hh);

                        foreach (string sl in spec)
                            hh = hh.Replace(sl, " ");
                       // MessageBox.Show(hh);
                        ObjWorkSheet.Cells[i, 6] = hh;
                        ObjExcel.ActiveWorkbook.Save();
                        continue;
                    }
                    else
                    if (URL.Contains("блог-инженера.рф") && !helper.endsBad(URL)) // тут почти все норм с выгрузкой
                    {
                        ObjWorkSheet.Cells[i, 6] = ObjWorkSheet.Cells[i, 5].Text.Replace("Спасибо за участие! Продолжение следует ... Получайте анонсы новых заметок сразу на свой E-MAIL", " ");
                        continue;
                    }


                    else
                    if (URL.Contains("solidwaste.ru") && !helper.endsBad(URL))
                    {
                        web.OverrideEncoding = Encoding.GetEncoding(1251);
                        HtmlNode.ElementsFlags.Remove("p");
                        
                        selector = "//p[@align='justify']";
                        try
                        {
                            doc = web.Load(URL);
                        }
                        catch
                        {
                            MessageBox.Show("Сайт недоступен");
                            continue;
                        }
                        var nodet = doc.DocumentNode.SelectNodes(selector);
                        string itogt = "";
                        try
                        {
                            for (int ui = 0; ui < nodet.Count; ui++)
                                itogt = itogt + nodet[ui].InnerText;
                            itogt = helper.noHtml(itogt);

                        }
                        catch
                        {
                            richTextBox1.AppendText(URL + " " + selector);
                            richTextBox1.ScrollToCaret();
                            continue;
                        }
                        foreach (string sl in spec)
                            itogt = itogt.Replace(sl, " ");
                        try
                        {
                            itogt = itogt.Substring(0, itogt.IndexOf("Обсудить новость в форуме"));
                        }
                        catch { }
                        //MessageBox.Show(itogt);
                        HtmlNode.ElementsFlags.Add("p", HtmlElementFlag.Empty);
                        ObjWorkSheet.Cells[i, 6] = itogt;

                        if (i % 20 == 0) ObjExcel.ActiveWorkbook.Save();
                        continue;

                    } 
                    else
                     if (URL.Contains("ecoindustry.ru") && !helper.endsBad(URL))
                    {
                        web.OverrideEncoding = Encoding.GetEncoding(1251);
                        HtmlNode.ElementsFlags.Remove("p");
                        //selector = "(//td[@class='main_txt']//text())";
                        selector = "//p[@align='justify']";
                        try
                        {
                            doc = web.Load(URL);
                        }
                        catch
                        {
                            MessageBox.Show("Сайт недоступен");
                            continue;
                        }
                        var nodet = doc.DocumentNode.SelectNodes(selector);
                        string itogt = "";
                                            
                        try
                        {
                            for (int ui = 0; ui < nodet.Count; ui++)
                                itogt = itogt + nodet[ui].InnerText;
                            itogt = helper.noHtml(itogt);

                        }
                        catch
                        {
                            richTextBox1.AppendText(URL + " " + selector);
                            richTextBox1.ScrollToCaret();
                            continue;
                        }
                        foreach (string sl in spec)
                            itogt = itogt.Replace(sl, " ");
                        try
                        {
                            itogt = itogt.Substring(0, itogt.IndexOf("Чтобы добавить комментарий"));
                        }
                        catch { }
                        //MessageBox.Show(itogt);
                        HtmlNode.ElementsFlags.Add("p", HtmlElementFlag.Empty);
                         ObjWorkSheet.Cells[i, 6] = itogt;

                        if (i % 10 == 0) ObjExcel.ActiveWorkbook.Save();
                        continue;
                    } 


                    else
                    if (URL.Contains("vniiecology.ru") && !helper.endsBad(URL))
                    {
                        selector = "(//div[@class='content clearfix'] //p)  [position()=1]";
                        web.OverrideEncoding = Encoding.UTF8;
                        doc = web.Load(URL);
                        var trynode = doc.DocumentNode.SelectSingleNode(selector);
                        if (helper.noHtml(trynode.InnerText).Contains("Текст, фото")) // Когда нет заголовка, а сразу идет инфа о фото и тексте
                            selector = "(//div[@class='content clearfix'] //p)  [position()>1]";
                        else
                            selector = "(//div[@class='content clearfix'] //p)  [position()>2]"; // Когда есть и заголовок и инфа о фото и тексте  
                        var nodes = doc.DocumentNode.SelectNodes(selector);
                        string hh = "";
                        for (int ui = 0; ui < nodes.Count; ui++)
                            hh = hh + nodes[ui].InnerText;

                        hh = helper.noHtml(hh);

                        foreach (string sl in spec)
                            hh = hh.Replace(sl, " ");
                        //MessageBox.Show(hh);
                        ObjWorkSheet.Cells[i, 6] = hh;
                        continue;
                    }
                    else
                    {
                        
                       // MessageBox.Show("tut");
                        continue;
                    }
                    try
                    {
                        doc = web.Load(URL);
                    }
                    catch
                    {
                        MessageBox.Show("Сайт недоступен");
                        continue;
                    }

                    var node = doc.DocumentNode.SelectSingleNode(selector);
                    //var node = doc.GetElementbyId("justify");

                    /*  if (URL.Contains("ecoindustry.ru"))
                      {
                          string itog1 = noHtml(node.NextSibling.InnerText);
                          foreach (string sl in spec)
                              itog1 = itog1.Replace(sl, "");
                          itog1 = itog1.Replace("Чтобы добавить комментарий, надо ", "");

                          MessageBox.Show(itog1);
                      } */
                    string itog = "";
                    try
                    {
                        itog = helper.noHtml(node.InnerText);

                    }
                    catch
                    {
                        richTextBox1.AppendText(URL + " " + selector);
                        richTextBox1.ScrollToCaret();
                        continue;
                    }


                    foreach (string sl in spec)
                        itog = itog.Replace(sl, " ");                
                    //MessageBox.Show(itog);
                    richTextBox1.AppendText(itog);
                    richTextBox1.ScrollToCaret();
                    ObjWorkSheet.Cells[i, 6] = itog;
                    if (i / 10 == 0) 
                    ObjExcel.ActiveWorkbook.Save();
                    
                }
                ObjExcel.ActiveWorkbook.Save();
                MessageBox.Show("vse");
                
                ObjExcel.Quit();
                //
            }
        }

        private void button8_Click(object sender, EventArgs e)
        {
            ObjExcel.ActiveWorkbook.Save();
        }
    }
}

