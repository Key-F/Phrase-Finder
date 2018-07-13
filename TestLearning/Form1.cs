﻿using HtmlAgilityPack;
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

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        
        
        static int keywordsChecker(String Text, String keywords)
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

        static string noHtml(string html)
        {
            //string html = "<html><head><title>Home Page</title></head><body>Welcome</body></html>";
            string newhtml = Regex.Replace(html, "<[^>]+>", string.Empty);
            return newhtml; }

        private void button1_Click(object sender, EventArgs e)
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

            }
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
                        range.Cells.Value = noHtml(text);
                        
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
            string[] spec = { "&rdquo;", "&nbsp;", "&raquo;", "&laquo;", "&ndash;", "&qt;", "/n", "/t", "&mdash", "&quot" }; // Спецсимволы, которые нужно убрать
            string selector = "";
            string URL = "";
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

                for (int i = Convert.ToInt32(firstN.Text); i < Convert.ToInt32(lastN.Text); i++)
                {
                    Excel.Range range = ObjWorkSheet.get_Range("B" + i, "B" + i); // Где искать текст
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
                    HtmlWeb web = new HtmlWeb();
                    HtmlAgilityPack.HtmlDocument doc = null;

                    if (URL.Contains("ohranatruda.ru") && !helper.endsBad(URL))
                    {
                        selector = "//div[@class='ot-news-detail-text']";
                        web.OverrideEncoding = Encoding.GetEncoding(1251); // meta charset
                    }
                    else
                    if (URL.Contains("transneft.ru") && !helper.endsBad(URL))
                    {
                        selector = "//div[@class='content']";
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
                        selector = "//div[@class='news-detail']";
                        web.OverrideEncoding = Encoding.UTF8;
                    }
                    else continue;

                    /*if (URL.Contains("ecoindustry.ru") && !helper.endsBad(URL))
                    {
                        //selector = "//p[@align='justify']";
                        //selector = "//p[@align='justify']//br";
                        //selector = "//div[@class='date']";
                        //selector = "//*[@id='td_main_center']/p/text()";
                        //selector = "(//*[@id='td_main_center']//p)  [position()>1] ";
                        //selector = "(//*[@align='justify']//br)  [position()>0] ";
                        //selector = "//*[@id='td_main_center']/p[2]";
                        //selector = "/html[1]/body[1]/div[2]/table[3]/tbody/tr/td[2]/p[2]";
                        //selector = "//td[@id='td_main_center']" //p[2];
                        //selector = "html[1]/body[1]/div[2]/table[3]/tr[1]/td[2]/#text[8]";
                        selector = "//*[@id='td_main_center']/p[2]/text()";
                        web.OverrideEncoding = Encoding.GetEncoding(1251);
                    } */
                    // gisnadzor "(//*[@class='news-detail']//p)  [position()>0] "
                    // ecolawer (//*[@class='content']//br)  [position()>3] 
                    // mnr.gov class="item-section section--lined
                    // http://www.mpe-sro.ru ро p пройтись
                    // nbpo class = 'Post_content'
                    // oaontc по p и ul 
                    // http://www.risk-news.ru/ источник обрезать
                    //http://www.ronktd.ru/news/2017/699/ класс news-detail и по p
                    // http://www.solidwaste.ru/news/view/21524.html как и в ecoindustry
                    // http://www.vernadsky.ru/news/news/?ELEMENT_ID=747 только p
                    // http://www.vestipb.ru/indnews8024.html jeppa
                    // rosmintrud vse ok https://ohranatruda.ru/news/2845/166801/
                    // gost class news-view__content--text и по p
                    // https://www.orfi.ru/press/news/2017/?n=2017080901 по p, кроме последнего
                    // safework = все ок
                    // vniims = все ок
                    // блог инженера - убрать  😉 Спасибо за участие! Продолжение следует ... Получайте анонсы новых заметок сразу на свой E-MAIL

                    if (URL.Contains("vniiecology.ru") && !helper.endsBad(URL))
                    {
                        selector = "(//div[@class='content clearfix'] //p)  [position()=1]";
                        web.OverrideEncoding = Encoding.UTF8;
                        doc = web.Load(URL);
                        var trynode = doc.DocumentNode.SelectSingleNode(selector);
                        if (noHtml(trynode.InnerText).Contains("Текст, фото")) // Когда нет заголовка, а сразу идет инфа о фото и тексте
                            selector = "(//div[@class='content clearfix'] //p)  [position()>1]";
                        else
                            selector = "(//div[@class='content clearfix'] //p)  [position()>2]"; // Когда есть и заголовок и инфа о фото и тексте  
                        var nodes = doc.DocumentNode.SelectNodes(selector);
                        string hh = "";
                        for (int ui = 0; ui <= nodes.Count; ui++)
                            hh = hh + nodes[ui].InnerText;

                        hh = noHtml(hh);

                        foreach (string sl in spec)
                            hh = hh.Replace(sl, "");
                        MessageBox.Show(hh);
                    }


                    try
                    {
                        doc = web.Load(URL);
                    }
                    catch
                    {
                        MessageBox.Show("Сайт недоступен");
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
                        itog = noHtml(node.InnerText);

                    }
                    catch
                    {
                        richTextBox1.AppendText(URL + " " + selector);
                        richTextBox1.ScrollToCaret();
                        continue;
                    }


                    foreach (string sl in spec)
                        itog = itog.Replace(sl, "");
                    if (URL.Contains("ecoindustry.ru"))
                        itog = itog.Replace("Чтобы добавить комментарий, надо ", "");
                    //MessageBox.Show(itog);
                    //richTextBox1.AppendText(itog);
                    //richTextBox1.ScrollToCaret();
                    //Excel.Range rangetext = ObjWorkSheet.get_Range("E" + i, "E" + i);
                    ObjWorkSheet.Cells[i, 5] = itog;
                    }
                //ObjExcel.Quit();
                //ObjExcel.ActiveWorkbook.Save();
            }
        }

        private void button8_Click(object sender, EventArgs e)
        {
            ObjExcel.ActiveWorkbook.Save();
        }
    }
}

