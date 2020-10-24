using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;

namespace window3
{

    public partial class mainForm : Form
    {
        //Добавляю свое, для Word
        object oMissing = System.Reflection.Missing.Value;
        object oEndOfDoc = "\\endofdoc"; // \endofdoc - предустановленная закладка

        // Запускаем Word и создаём новый документ
        Word._Application oWord;
        Word._Document oDoc;



        bool flag = false;
        regForm reg = new regForm();
        idForm idform = new idForm();
        teleForm tele = new teleForm();

        public mainForm()
        {
            InitializeComponent();
        }

        //Закрытие формы входа
        public void loginForm_FormClosed(object sender, EventArgs e)
        {
            this.Close();
        }


        private void mainForm_Load(object sender, EventArgs e)
        {
/*            Excel.Application excelApp = new Excel.Application();

            if (excelApp == null)
            {
                Console.WriteLine("Excel is not installed!!");
                return;
            }

            Excel.Workbook excelBook = excelApp.Workbooks.Open(@"E:\readExample.xlsx");
            Excel._Worksheet excelSheet = (Excel._Worksheet)excelBook.Sheets[1];
            Excel.Range excelRange = excelSheet.UsedRange;

            int rows = excelRange.Rows.Count;
            int cols = excelRange.Columns.Count;

            for (int i = 1; i <= rows; i++)
            {
                // read new line
                for (int j = 1; j <= cols; j++)
                {
                    //write to cell
                    //if (excelRange.Cells[i, j] != null);
                    //

                }
            }
            excelApp.Quit();*/
        }

        //Разворачивание из трея при двойной нажатии на иконку
        private void notifyIcon1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            Show();
            WindowState = FormWindowState.Normal;
        }

        //Сворачивание приложения в трей при нажатии на крестик
        public void mainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            this.Hide();
            this.ShowInTaskbar = false;
            e.Cancel = true;
            if (flag == true)
            {
                e.Cancel = false;
                flag = false;
            }
        }

        //При нажатии на кнопку открытие формы регистрации 

        //При входе в учетную запись на главном экране переопределяются кнопки
        //"Вход" => "Телеметрия"
        //"Регистрация" => "Личный кабинет"
        private void button1_Click(object sender, EventArgs e)
        {
        }

        //При нажатии открывается форма Телеметрия
        private void next_button1_Click(object sender, EventArgs e)
        {
            tele.ShowDialog();
        }

        //Завершение работы из главного экрана 
        private void ВыходToolStripMenuItem_Click(object sender, EventArgs e)
        {
            flag = true;
            this.Close();
        }

        //При нажатии открывается личный кабинет при условии входа в учетную запись
        private void idButton_Click_1(object sender, EventArgs e)
        {
        }

        //Завершение работы из трея
        private void закрытьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            flag = true;
            this.Close();
        }

        //Кнопка получить отчет
        private void buttonGetRep_Click(object sender, EventArgs e)
        {
            oWord = new Word.Application(); // Запускаем Word
            oWord.Visible = true; // Делаем окно Word видимым

            // Старый способ: здесь отражено наличие ключевого слова ref 
            //и параметра oMissing, которые можно не использовать
            oDoc = oWord.Documents.Add(ref oMissing, ref oMissing, ref oMissing, 
                ref oMissing); // Создаём новый документ

            // Вставка текста в начало документа и отступа после
            Word.Paragraph oPara1;

            // Старый способ
            //oPara1 = oDoc.Content.Paragraphs.Add(ref oMissing);
            oPara1 = oDoc.Content.Paragraphs.Add();

            oPara1.Range.Text = "Заголовок № 1 с тенью";

            oPara1.Range.Font.Size = 20; // Размер шрифта: 20

            oPara1.Range.Font.Shadow = 1; // Тенью от шрифта

            oPara1.Range.Font.Bold = 1; // "Жирный" шрифт

            oPara1.Format.SpaceAfter = 24; // 24 пт.: оступ после параграфа

            oPara1.Range.InsertParagraphAfter();

            oPara1.Range.Font.Size = 12; // Размер шрифта: 12

            oPara1.Range.Font.Shadow = 0; // Тенью от шрифта: выключаем


            // Вставка текста и отступа после (для последующих частей документа)
            Word.Paragraph oPara2;

            object oRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;

            oPara2 = oDoc.Content.Paragraphs.Add(ref oRng);

            oPara2.Range.Text = "Заголовок № 2";

            oPara2.Format.SpaceAfter = 6; // Отступ после

            oPara2.Range.InsertParagraphAfter();


            // Вставка текста
            Word.Paragraph oPara3;

            oRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;

            oPara3 = oDoc.Content.Paragraphs.Add(ref oRng);

            oPara3.Range.Text = "Обычный текст. Дальше идёт таблица:";

            oPara3.Range.Font.Bold = 0;

            oPara3.Format.SpaceAfter = 24;

            oPara3.Range.InsertParagraphAfter();


            // Вставка таблицы 3 на 5, заполнение данными, и изменение первой строки: "жирный" и "курсив".
            Word.Table oTable;

            Word.Range wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;

            oTable = oDoc.Tables.Add(wrdRng, 3, 5); // 3 строки, 5 столбцов (Add требует 5 параметров, но мы записываем без двух последних параметров oMissing)

            oTable.Range.ParagraphFormat.SpaceAfter = 6;

            int r, c;

            string strText;

            for (r = 1; r <= 3; r++) // Заполняем строки
            {
                for (c = 1; c <= 5; c++) // Заполняем столбцы
                {
                    strText = "r" + r + "c" + c;
                    oTable.Cell(r, c).Range.Text = strText;
                }
                oTable.Rows[1].Range.Font.Bold = 1; // Меняем стиль первой строки: "жирный"
                oTable.Rows[1].Range.Font.Italic = 1; // Меняем стиль первой строки: "курсив"
            }


            // Вставка текста после таблицы
            Word.Paragraph oPara4;

            oRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;

            oPara4 = oDoc.Content.Paragraphs.Add(ref oRng);

            oPara4.Range.InsertParagraphBefore(); // Вставка отступ до с параметром 24 пт. (подтягиваем из oPara3 по умолчанию)

            oPara4.Range.Text = "Вставляем другую таблицу:";

            oPara4.Format.SpaceAfter = 24;

            oPara4.Range.InsertParagraphAfter(); // Вставка оступа после с параметром 24 пт.



            // Вставка таблицы 5 на 2, заполнение данными, и изменение размера ширины столбцов
            wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;

            oTable = oDoc.Tables.Add(wrdRng, 5, 2);

            oTable.Range.ParagraphFormat.SpaceAfter = 6;

            for (r = 1; r <= 5; r++)
            {
                for (c = 1; c <= 2; c++)
                { 
                    strText = "r" + r + "c" + c;
                    oTable.Cell(r, c).Range.Text = strText;
                }
                oTable.Columns[1].Width = oWord.InchesToPoints(2); // Изменение ширины столбца 1
                oTable.Columns[2].Width = oWord.InchesToPoints(3); // Изменение ширины столбца 2
            }


        }
    }
}
