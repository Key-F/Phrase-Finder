using HtmlAgilityPack;
using System;
using System.Drawing;
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

                //Создаём приложение.
                Microsoft.Office.Interop.Excel.Application ObjExcel = new Microsoft.Office.Interop.Excel.Application();
                //Открываем книгу.                                                                                                                                                        
                Microsoft.Office.Interop.Excel.Workbook ObjWorkBook = ObjExcel.Workbooks.Open(openFileDialog1.FileName, 0, false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
                //Выбираем таблицу(лист).
                Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet;
                ObjWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBook.Sheets[1];

                try
                {

                    for (int i = Convert.ToInt32(firstN.Text); i <= Convert.ToInt32(lastN.Text); i++)
                    {
                        Excel.Range range = ObjWorkSheet.get_Range(textBox6.Text + i, textBox6.Text + i); // Заголовки
                        Form1.ActiveForm.Text = "ИДЕТ РАБОТА";
                        
                        string text;
                        string keyword = textBox4.Text;
                        if (range.Text != null)
                        {
                            text = range.Text.ToLower();
                            int k = keywordsChecker(text, keyword);
                            if (k > 0)
                            {
                                
                                switch (comboBox1.Text) // Цвет текста
                                {
                                    case "Черный": range.Cells.Font.Color = ColorTranslator.ToOle(Color.Black);
                                        break;
                                    case "Красный":
                                        range.Cells.Font.Color = ColorTranslator.ToOle(Color.Red);
                                        break;
                                    case "Синий":
                                        range.Cells.Font.Color = ColorTranslator.ToOle(Color.Blue);
                                        break;
                                    case "Желтый":
                                        range.Cells.Font.Color = ColorTranslator.ToOle(Color.Yellow);
                                        break;
                                    case "Оранжевый":
                                        range.Cells.Font.Color = ColorTranslator.ToOle(Color.Orange);
                                        break;
                                    case "Зеленый":
                                        range.Cells.Font.Color = ColorTranslator.ToOle(Color.Green);
                                        break;
                                }
                                switch (comboBox2.Text) //Фоновый цвет

                                {
                                    case "Белый":
                                        range.Interior.Color = ColorTranslator.ToOle(Color.White);
                                        break;
                                    case "Красный":
                                        range.Interior.Color = ColorTranslator.ToOle(Color.Red);
                                        break;
                                    case "Синий":
                                        range.Interior.Color = ColorTranslator.ToOle(Color.Blue);
                                        break;
                                    case "Желтый":
                                        range.Interior.Color = ColorTranslator.ToOle(Color.Yellow);
                                        break;
                                    case "Оранжевый":
                                        range.Interior.Color = ColorTranslator.ToOle(Color.Orange);
                                        break;
                                    case "Зеленый":
                                        range.Interior.Color = ColorTranslator.ToOle(Color.Green);
                                        break;
                                }
                                
                            }


                        }

                    }
                }
                catch
                {
                    ObjExcel.Quit();
                    Form1.ActiveForm.Text = "Phrase Finder";
                   MessageBox.Show("что-то не так");
                  
                }


                Form1.ActiveForm.Text = "Phrase Finder";
                ObjExcel.Quit();  //Удаляем приложение (выходим из экселя) - а то будет висеть в процессах!
            }

            }
            static int keywordsChecker(String shortEssay, String keywords)
            {
                int count = 0;
                String[] ary = keywords.Split(',');
                for (int i = 0; i < ary.Length; i++)
                {
                while (ary.Length > 0 && char.IsWhiteSpace(ary[i][0]))
                {
                    ary[i] = ary[i].Remove(0, 1);
                }
                if (shortEssay.Contains(ary[i].ToLower()))
                    {
                        count++;
                    }
                }
                return count;
            }
        }
    }

