using System;
using System.ComponentModel;
using System.Data;
using System.IO;
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


        Word._Application oWord;
        Word._Document oDoc;
        String pathLEP = "";
        String pathPS = "";
        String DB = "";
        SQLRequests localDB = new SQLRequests();
        SQLRequests remoteDB = new SQLRequests();
        int pathButt = 0;

        public mainForm()
        {
            InitializeComponent();
        }


        private void mainForm_Load(object sender, EventArgs e)
        {
            ClientSize = new System.Drawing.Size(535, ClientSize.Height);
            localDB.Connect("local.db");
            if (localDB.isThereRes("SELECT name from sqlite_master where type= \"table\"")){
                DataTable res = localDB.getResTable("SELECT path FROM config");
                DB = res.Rows[0].ItemArray[0].ToString();
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            ClientSize = new System.Drawing.Size(1095, ClientSize.Height);
        }

        private void next_button1_Click(object sender, EventArgs e)
        {
            
        }

        private void ВыходToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }
        
        private void idButton_Click_1(object sender, EventArgs e)
        {
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            pathButt = 0;
            tablePicker.ShowDialog();
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            button2.Enabled = !((CheckBox)sender).Checked;
        }

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
            string str = "Анализ технического состояния и возрастная " +
    "структура линий электропередачи и подстанций";
            oPara1 = oDoc.Content.Paragraphs.Add();
            oPara1.Range.Text = str;
            oPara1.Range.Font.Size = 14; // Размер шрифта
            oPara1.Range.Font.Bold = 1; // "Жирный" шрифт
            oPara1.Format.SpaceAfter = 20; // оступ после параграфа
            oPara1.Range.InsertParagraphAfter();

            //тут обращение к БД название компании
            string CompName = "Имя компании";

            Word.Paragraph oPara2;
            object oRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            oPara2 = oDoc.Content.Paragraphs.Add(ref oRng);
            oPara2.Range.Text = CompName;
            oPara2.Format.SpaceAfter = 1; // Отступ после
            oPara2.Range.InsertParagraphAfter();

            str = "\t" + "Протяженность ВЛ 110 кВ и КЛ 110 кВ, количество и суммарная мощность ПС " +
                "110 кВ, находящихся в собственности " + CompName + ", по состоянию на " +
                "01.01.2019 г. составили:" + "\n";

            Word.Paragraph oPara3;
            object oRng1 = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            oPara3 = oDoc.Content.Paragraphs.Add(ref oRng1);
            oPara3.Range.Text = str;
            oPara3.Format.SpaceAfter = 6; // Отступ после
            oPara3.Range.InsertParagraphAfter();







            /*// Вставка текста и отступа после (для последующих частей документа)
            Word.Paragraph oPara2;

            object oRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;

            oPara2 = oDoc.Content.Paragraphs.Add(ref oRng);

            oPara2.Range.Text = "Заголовок № 2";

            oPara2.Format.SpaceAfter = 6; // Отступ после

            oPara2.Range.InsertParagraphAfter();
*/

            // Вставка текста
            /*Word.Paragraph oPara3;

            oRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;

            oPara3 = oDoc.Content.Paragraphs.Add(ref oRng);

            oPara3.Range.Text = "Обычный текст. Дальше идёт таблица:";

            oPara3.Range.Font.Bold = 0;

            oPara3.Format.SpaceAfter = 24;

            oPara3.Range.InsertParagraphAfter();*/


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

        private void button3_Click(object sender, EventArgs e)
        {
            if (!(File.Exists(pathPS) && File.Exists(pathLEP) && 
                (Path.GetExtension(pathPS) == ".xls" || Path.GetExtension(pathPS) == ".xlsx") &&
                (Path.GetExtension(pathLEP) == ".xls" || Path.GetExtension(pathLEP) == ".xlsx")))
            {
                labelError.Text = "Проблема с выбранными файлами или файлы не выбраны вовсе";
                return;
            }
            labelError.Text = "";

            Excel.Application excelApp = new Excel.Application();
            if (excelApp == null)
            {
                MessageBox.Show("Excel is not installed!!");
                return;
            }
            Excel.Workbook excelBook = excelApp.Workbooks.Open(pathLEP);
            Excel.Worksheet excelSheet = excelBook.Sheets[1];
            Excel.Range excelRange = excelSheet.UsedRange;
            excelApp.Visible = true;
            int rows = excelRange.Rows.Count;
            int cols = excelRange.Columns.Count;

            //////////////          СЧИТЫВАЕМ ПЕРВУЮ СТРОКУ, ИЩЕМ НАЗВАНИЕ           //////////////

            //////////////          ИЩЕМ PT           //////////////
            for (int i = 1; i <= rows; i++)
            {
                for (int j = 1; j <= cols; j++)
                {
                    if (excelRange.Cells[i, j].Value2 != null)
                        excelSheet.Cells[1,1] = "1";
                }
            }
            excelApp.Quit();
            ClientSize = new System.Drawing.Size(535, ClientSize.Height);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            pathButt = 1;
            tablePicker.ShowDialog();
        }

        private void filePicker_FileOk(object sender, CancelEventArgs e)
        {
            switch (pathButt)
            {
                case 0:
                    pathLEP = ((OpenFileDialog)sender).FileName;
                    break;
                case 1:
                    pathPS = ((OpenFileDialog)sender).FileName;
                    break;
            }
        }

        private void выбратьСетевуюПапкуСБДToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DBPicker.ShowDialog();
        }

        private void DBPicker_FileOk(object sender, CancelEventArgs e)
        {
            DB = ((OpenFileDialog)sender).FileName;
            localDB.makeQuery("update config set path = \"" + DB + "\" where id = 1");
        }
    }
}
