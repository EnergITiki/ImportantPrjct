using System;
using System.ComponentModel;
using System.Data;
using System.IO;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;
using Microsoft.Office.Interop.Word;

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
            //button2.Enabled = !((CheckBox)sender).Checked;
        }

        public void PrintText(string text, int size, int bold, int Ital, int SpaceAfter)
        {
            Word.Paragraph oPara1;
            oPara1 = oDoc.Content.Paragraphs.Add();
            oPara1.Range.Text = text;
            oPara1.Range.Font.Name = "Times New Roman";
            oPara1.Range.Font.Size = size; // Размер шрифта
            oPara1.Range.Font.Bold = bold; // "Жирный" шрифт
            oPara1.Range.Font.Italic = Ital; // "Курсив" шрифт
            oPara1.Format.SpaceAfter = SpaceAfter; // оступ после параграфа
            oPara1.Range.InsertParagraphAfter();

            oPara1.Range.Font.Bold = 0; // "Жирный" шрифт
            oPara1.Range.Font.Italic = 0; // "Курсив" шрифт
        }

        public void PrintTable(double VL, double KL, double P, int count)
        {
            //Формируем таблицу
            Word.Table oTable;
            Word.Range wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            oTable = oDoc.Tables.Add(wrdRng, 2, 2); // 2 строки, 2 столбцов
            oTable.Range.Rows.Alignment = WdRowAlignment.wdAlignRowCenter;
            oTable.Range.ParagraphFormat.SpaceAfter = 5;
            oTable.Cell(1, 1).Range.Text = "Протяженность действующих ВЛ и КЛ (в одноцепном исчислении), км";
            oTable.Cell(1, 2).Range.Text = "ВЛ - " + VL.ToString() + "\n" + "КЛ - " + KL.ToString();
            oTable.Cell(2, 1).Range.Text = "Количество и суммарная установленная мощность ПС, шт./ МВА";
            oTable.Cell(2, 2).Range.Text = count.ToString() + "/" + P.ToString();
            oTable.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleEmboss3D;//стиль границ
        }

        public void RES( string CompName, string FilialName) //РЭСы филиалов 
        {
            //Обращаемся к БД нахуй




            /*


На рисунке 4.3 приведена возрастная структура линий электропередачи и подстанций
110 кВ [энергосистемы] по состоянию на 01.01.2019 г., обслуживаемых [сетевым предприятием
1].
 * 
 */


        }
        public void Filials(string CompName)//Филиалы
        {
            string FilialName = "Биба и Боба";//тут в БД обращение!
            string str = "\t" + "Филиал компании \"" + CompName + "\" \"" + FilialName + "\"";
            PrintText(str, 12, 1, 0, 1);

            str = "\t" + "Протяженность ВЛ 110 кВ и КЛ 110 кВ, количество и суммарная мощность ПС " +
            "110 кВ, находящихся в собственности " + FilialName + ", по состоянию на " +
            "01.01.2019 г. составили:";
            PrintText(str, 12, 0, 0, 5);

            //обращаемся к БД
            double VL = 88005, KL = 896.5, P = 2548;
            int count = 100;
            //Рисуем таблицу
            PrintTable(VL, KL, P, count);

            str = "\t" + "Анализ технического состояния электросетевых объектов " +
                "напряжением 110 кВ \"" + FilialName + "\" показал: ";
            PrintText(str, 12, 0, 0, 6);

            //Делаем маркированный список

            //Обращаемся к базам данных
            int countPods = 15;
            double pocOldPods = 28.8;
            int countP = 122;
            double procOldTrans = 3.6;
            double lengthV = 1396.6;
            double procAllLenV = 84.2;
            double lengthK = 82.1;
            double procAllLenK = 100;

            str = countPods.ToString() + " подстанций ( " + pocOldPods.ToString() + "% " +
                "от общего числа ПС 110 кВ) отработали более 50 лет;";
            Word.Paragraph list;
            list = oDoc.Content.Paragraphs.Add();
            list.Range.Text = str;
            list.Range.SetListLevel(1);
            list.Range.ListFormat.ApplyBulletDefault(Word.WdListGalleryType.wdBulletGallery);
            list.Range.InsertParagraphAfter();

            str = countP.ToString() + " МВА трансформаторной мощности (" + procOldTrans.ToString() +
                "% от общей трансформаторной мощности напряжением 110 кВ) отработало более 50 лет;";
            PrintText(str, 12, 0, 0, 0);
            str = "воздушные линии электропередачи 110 кВ протяженностью " + lengthV.ToString() +
                " км в одноцепном исчислении(" + procAllLenV.ToString() + "% от общей " +
                "протяженности ВЛ 110 кВ) отработали более 50 лет;";
            PrintText(str, 12, 0, 0, 0);
            str = "кабельные линии электропередачи 110 кВ протяженностью " + lengthK.ToString() +
                " км (" + procAllLenK.ToString() + "% от общей протяженности КЛ 110 кВ) " +
                "находятся в эксплуатации от 2 до 14 лет.";
            list.OutlineDemoteToBody();

            //Обращаемся к БД ПС
            int[] arrYearPS = new int[4];
            arrYearPS[0] = 68;
            arrYearPS[1] = 67;
            arrYearPS[2] = 64;
            arrYearPS[3] = 63;

            string[] arrNamePS = new string[4];
            arrNamePS[0] = "Северная";
            arrNamePS[1] = "Западная";
            arrNamePS[2] = "Восточная";
            arrNamePS[3] = "Южная";

            str = "Наиболее продолжительное время эксплуатируются ПС 110 кВ " + arrNamePS[0] +
                " - срок службы " + arrYearPS[0].ToString() + " лет, ПС 110 кВ " + arrNamePS[1] +
                " – " + arrYearPS[1].ToString() + " лет, ПС 110 кВ " + arrNamePS[2] + " – " +
                arrYearPS[2].ToString() + " года, ПС 110 кВ " + arrNamePS[3] + " – " + arrYearPS[3].ToString() +
                " года.";
            PrintText(str, 12, 0, 0, 5);

            //Обращаемся к БД
            int[] arrYearLP = new int[2];
            arrYearLP[0] = 68;
            arrYearLP[1] = 67;


            string[] arrNameLP = new string[2];//Дохуя вопросов хуячим, и склеиваем в одно
            arrNameLP[0] = "ВЛ 110 кВ ПС1 – ПС2 I цепь (А-1), ВЛ 110 кВ ПС1 – ПС2 II цепь (А-2)";
            arrNameLP[1] = "ХУЙНЯ КАКАЯ ТО ВАШ КЕЙС, ВЫ БЫ ДАННЫЕ НОРМАЛЬНЫЕ ДАЛИ СНАЧАЛА!!! " +
                "ЭТО Ж БЛЯТЬ ТАК СЛОЖНО ВСЕ ПО ОДНОМУ ШАБЛОНУ ДЕЛАТЬ ПРОСТО ПИЗДЕЦ";

            str = "Наиболее продолжительное время находятся в эксплуатации следующие линии " +
                "электропередачи 110 кВ: " + arrNameLP[0] + ", " + arrYearLP[0].ToString() + " лет, " +
                arrNameLP[1] + ", " + arrYearLP[1].ToString() + "лет.";
            PrintText(str, 12, 0, 0, 5);



            //Еще один цикл нахуй, все РЭСы просматриваем
            RES(CompName, FilialName);




        }

        public void Companies()// компании
        {
            // Вставка текста в начало документа и отступа послe
            string str = "Анализ технического состояния и возрастная " +
            "структура линий электропередачи и подстанций";
            PrintText(str, 12, 1, 0, 10);


            //тут обращение к БД название компании
            string CompName = NameCompany.Text;//Пока так, потом из БД надо будет считывать нормально!
            PrintText("\t\"" + CompName + "\"", 12, 1, 0, 1);


            str = "\t" + "Протяженность ВЛ 110 кВ и КЛ 110 кВ, количество и суммарная мощность ПС " +
            "110 кВ, находящихся в собственности " + CompName + ", по состоянию на " +
            "01.01.2019 г. составили:";
            PrintText(str, 12, 0, 0, 5);

            //обращаемся к БД
            double VL = 55555.5, KL = 785.6, P = 8000;
            int count = 59;
            //Рисуем таблицу
            PrintTable(VL, KL, P, count);

            str = "Далее приведена возрастная структура линий электропередачи и подстанций 110 кВ " +
                CompName + " по состоянию на 01.01.2019 г.с разбивкой по электросетевым " +
                "предприятиям.";
            PrintText(str, 12, 0, 0, 10);

            //Сложная какая та логика в цикле, чтобы все  филиалы компании перебрать
            Filials(CompName);



        }

        private void buttonGetRep_Click(object sender, EventArgs e)
        {
            oWord = new Word.Application(); // Запускаем Word
            oWord.Visible = true; // Делаем окно Word видимым
           

            // Старый способ: здесь отражено наличие ключевого слова ref 
            //и параметра oMissing, которые можно не использовать
            oDoc = oWord.Documents.Add(ref oMissing, ref oMissing, ref oMissing, 
                ref oMissing); // Создаём новый документ

            //в цикле блять хуячим компании
            Companies();







            // Вставка текста после таблицы
            /* Word.Paragraph oPara4;

             oRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;

             oPara4 = oDoc.Content.Paragraphs.Add(ref oRng);

             oPara4.Range.InsertParagraphBefore(); // Вставка отступ до с параметром 24 пт. (подтягиваем из oPara3 по умолчанию)

             oPara4.Range.Text = "Вставляем другую таблицу:";

             oPara4.Format.SpaceAfter = 24;

             oPara4.Range.InsertParagraphAfter(); // Вставка оступа после с параметром 24 пт.*/



            // Вставка таблицы 5 на 2, заполнение данными, и изменение размера ширины столбцов
            /*wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;

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
            }*/


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
