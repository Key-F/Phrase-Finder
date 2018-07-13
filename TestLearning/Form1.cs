using HtmlAgilityPack;
using System;
using System.Drawing;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace TestLearning
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void LoadFile_Click(object sender, EventArgs e)
        {




            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {

                this.Text = "ИДЕТ РАБОТА";
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
                    this.Text = "Phrase Finder";
                    return;

                }
                Microsoft.Office.Interop.Excel.Workbook ObjWorkBook = ObjExcel.Workbooks.Open(openFileDialog1.FileName, 0, false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
                //
                //Выбираем таблицу(лист).
                Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet;
                ObjWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBook.Sheets[1];



                int kolstrok = 0; // Количество строк в таблице, не прошедших проверки
               
                    for (int i = Convert.ToInt32(firstN.Text); i <= Convert.ToInt32(lastN.Text); i++)
                    {
                        Excel.Range range = ObjWorkSheet.get_Range(textBox6.Text + i, textBox6.Text + i); // Где искать текст
                        Excel.Range range1 = ObjWorkSheet.get_Range(textBox2.Text + i, textBox2.Text + i); // Где выделять текст
                       
                        string text;                        
                        string keyword = richTextBox1.Text;

                    if (range.Text != null)
                    {
                        text = range.Text.ToLower();
                        int k = keywordsChecker(text, keyword);
                        if (k > 0)
                        {
                            kolstrok++;
                            if (checkBox1.CheckState == CheckState.Checked)
                            {
                                switch (comboBox1.Text) // Цвет текста
                                {
                                    case "Черный":
                                        range1.Cells.Font.Color = ColorTranslator.ToOle(Color.Black);
                                        break;
                                    case "Красный":
                                        range1.Cells.Font.Color = ColorTranslator.ToOle(Color.Red);
                                        break;
                                    case "Синий":
                                        range1.Cells.Font.Color = ColorTranslator.ToOle(Color.Blue);
                                        break;
                                    case "Желтый":
                                        range1.Cells.Font.Color = ColorTranslator.ToOle(Color.Yellow);
                                        break;
                                    case "Оранжевый":
                                        range1.Cells.Font.Color = ColorTranslator.ToOle(Color.Orange);
                                        break;
                                    case "Зеленый":
                                        range1.Cells.Font.Color = ColorTranslator.ToOle(Color.Green);
                                        break;
                                }
                            }
                            if (checkBox2.CheckState == CheckState.Checked)
                            {
                                switch (comboBox2.Text) //Фоновый цвет

                                {
                                    case "Белый":
                                        range1.Interior.Color = ColorTranslator.ToOle(Color.White);
                                        break;
                                    case "Красный":
                                        range1.Interior.Color = ColorTranslator.ToOle(Color.Red);
                                        break;
                                    case "Синий":
                                        range1.Interior.Color = ColorTranslator.ToOle(Color.Blue);
                                        break;
                                    case "Желтый":
                                        range1.Interior.Color = ColorTranslator.ToOle(Color.Yellow);
                                        break;
                                    case "Оранжевый":
                                        range1.Interior.Color = ColorTranslator.ToOle(Color.Orange);
                                        break;
                                    case "Зеленый":
                                        range1.Interior.Color = ColorTranslator.ToOle(Color.Green);
                                        break;
                                }
                            }
                        }

                    }
                        

                    }
             
                MessageBox.Show("Было найдено " + kolstrok + " строк, в которых содержится хотя бы одна из фраз");
                this.Text = "Phrase Finder";
                ObjExcel.Quit();  //Удаляем приложение (выходим из экселя) - а то будет висеть в процессах!
            }

        }
        static int keywordsChecker(String Text, String keywords)
        {
            int count = 0;
            String[] ary = keywords.Split(',');

            for (int i = 0; i < ary.Length; i++)
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
    }
}

