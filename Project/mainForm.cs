using System;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;
using System.Reflection;
using System.Windows.Forms.DataVisualization.Charting;
//using Microsoft.Office.Interop.Word;

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
            oTable.Range.Rows.Alignment = Word.WdRowAlignment.wdAlignRowCenter;
            oTable.Range.ParagraphFormat.SpaceAfter = 5;
            oTable.Cell(1, 1).Range.Text = "Протяженность действующих ВЛ и КЛ (в одноцепном исчислении), км";
            oTable.Cell(1, 2).Range.Text = "ВЛ - " + VL.ToString() + "\n" + "КЛ - " + KL.ToString();
            oTable.Cell(2, 1).Range.Text = "Количество и суммарная установленная мощность ПС, шт./ МВА";
            oTable.Cell(2, 2).Range.Text = count.ToString() + "/" + P.ToString();
            oTable.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleEmboss3D;//стиль границ
        }

        double PictureNum = 1.1;

        public void RES( string CompName, string FilialName) //РЭСы филиалов 
        {
            //Обращаемся к БД нахуй
            string[] NameRES = new string[1];//Имена РЭСов
            NameRES[0] = "RAS";

            //


            string str = "На рисунке " + PictureNum.ToString() + " приведена возрастная структура " +
                "линий электропередачи и подстанций 110 кВ " + NameRES[0] + " по состоянию на 01.01.2019 " +
                "г., обслуживаемых " + FilialName;
            PrintText(str, 12, 0, 0, 5);


            //Рисуем Диаграммы нахуй
            //Обращаемся к базам данных блять!!!!
            //Подстанции
            double[] Pods = new double[3];
            Pods[0] = 60;
            Pods[1] = 20;
            Pods[2] = 20;
            double PodsAll = Pods[0] + Pods[1] + Pods[2];


            //Трансформаторы
            double[] Trans = new double[3];
            Trans[0] = 60;
            Trans[1] = 20;
            Trans[2] = 20;
            double TransAll = Trans[0] + Trans[1] + Trans[2];

            //ВЛ
            double[] VL = new double[3];
            VL[0] = 60;
            VL[1] = 20;
            VL[2] = 20;
            double VLAll = VL[0] + VL[1] + VL[2];

            //КЛ
            double[] KL = new double[3];
            KL[0] = 60;
            KL[1] = 20;
            KL[2] = 20;
            double KLAll = KL[0] + KL[1] + KL[2];

            //Подстанции**********************
            chart1.Series[0].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Pie;
            chart1.Series[0].IsValueShownAsLabel = true; // метки становятся видимыми
            chart1.Series[0].LabelFormat = "{#0}"; // формат отображения "{#0}"
            chart1.Titles.Add("Подстанции (штук, %) 110 кВ");
            chart1.Titles[1].Font = new Font("Times New Roman", 12, FontStyle.Bold);
            chart1.Legends[0].Font = new Font("Times New Roman", 9);
          
            chart1.Series[0].Points.AddY(Pods[0]);
            chart1.Series[0].Points.AddY(Pods[1]);
            chart1.Series[0].Points.AddY(Pods[2]);

            chart1.Series[0].Points[0].YValues[0] = Pods[0];
            chart1.Series[0].Points[1].YValues[0] = Pods[1];
            chart1.Series[0].Points[2].YValues[0] = Pods[2];

            chart1.Series[0].Points[0].LegendText = "свыше 50 лет";
            chart1.Series[0].Points[1].LegendText = "от 26 до 50 лет";
            chart1.Series[0].Points[2].LegendText = "до 25 лет";

            chart1.ChartAreas[0].Area3DStyle.Enable3D = true;



            //Сохраняем епта
            chart1.SaveImage(@"D:\\test1.jpeg", System.Drawing.Imaging.ImageFormat.Jpeg);

            string pathFileImage = "D:\\test1.jpeg";

            // Загружаем исходное изображение
            var image1 = new Bitmap(pathFileImage);

            // Масштабируем до нужного размера
            var image2 = new Bitmap(image1, 600, 400);
            image2.Save("D:\\test2.jpeg");

            var pPicture = oDoc.Paragraphs.Last.Range;
            oDoc.InlineShapes.AddPicture("D:\\test2.jpeg", Range: pPicture);

            image1.Dispose();
            image2.Dispose();

            //Трансформаторы*******************
            chart2.Series[0].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Pie;
            chart2.Series[0].IsValueShownAsLabel = true; // метки становятся видимыми
            chart2.Series[0].LabelFormat = "{#0}"; // формат отображения "{#0}"
            chart2.Titles.Add("Трансформаторы (МВА, %) 110 кВ");
            chart2.Titles[1].Font = new Font("Times New Roman", 12, FontStyle.Bold);
            chart2.Legends[0].Font = new Font("Times New Roman", 9);

            chart2.Series[0].Points.AddY(Trans[0]);
            chart2.Series[0].Points.AddY(Trans[1]);
            chart2.Series[0].Points.AddY(Trans[2]);

            chart2.Series[0].Points[0].YValues[0] = Trans[0];
            chart2.Series[0].Points[1].YValues[0] = Trans[1];
            chart2.Series[0].Points[2].YValues[0] = Trans[2];

            chart2.Series[0].Points[0].LegendText = "свыше 50 лет";
            chart2.Series[0].Points[1].LegendText = "от 26 до 50 лет";
            chart2.Series[0].Points[2].LegendText = "до 25 лет";

            chart2.ChartAreas[0].Area3DStyle.Enable3D = true;

            //Сохраняем епта
            chart2.SaveImage(@"D:\\test1.jpeg", System.Drawing.Imaging.ImageFormat.Jpeg);

            pathFileImage = "D:\\test1.jpeg";

            // Загружаем исходное изображение
            image1 = new Bitmap(pathFileImage);

            // Масштабируем до нужного размера
            image2 = new Bitmap(image1, 600, 400);
            image2.Save("D:\\test2.jpeg");

            pPicture = oDoc.Paragraphs.Last.Range;
            oDoc.InlineShapes.AddPicture("D:\\test2.jpeg", Range: pPicture);

            image1.Dispose();
            image2.Dispose();

            //ВЛ************************
            chart3.Series[0].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Pie;
            chart3.Series[0].IsValueShownAsLabel = true; // метки становятся видимыми
            chart3.Series[0].LabelFormat = "{#0}"; // формат отображения "{#0}"
            chart3.Titles.Add("Воздушные линии электропередачи в одноцепном исчислении(км, %) 110 кВ");
            chart3.Titles[1].Font = new Font("Times New Roman", 12, FontStyle.Bold);
            chart3.Legends[0].Font = new Font("Times New Roman", 9);

            chart3.Series[0].Points.AddY(VL[0]);
            chart3.Series[0].Points.AddY(VL[1]);
            chart3.Series[0].Points.AddY(VL[2]);

            chart3.Series[0].Points[0].YValues[0] = VL[0];
            chart3.Series[0].Points[1].YValues[0] = VL[1];
            chart3.Series[0].Points[2].YValues[0] = VL[2];

            chart3.Series[0].Points[0].LegendText = "свыше 50 лет";
            chart3.Series[0].Points[1].LegendText = "от 36 до 50 лет";
            chart3.Series[0].Points[2].LegendText = "до 35 лет";

            chart3.ChartAreas[0].Area3DStyle.Enable3D = true;

            //Сохраняем епта
            chart3.SaveImage(@"D:\\test1.jpeg", System.Drawing.Imaging.ImageFormat.Jpeg);

            pathFileImage = "D:\\test1.jpeg";

            // Загружаем исходное изображение
            image1 = new Bitmap(pathFileImage);

            // Масштабируем до нужного размера
            image2 = new Bitmap(image1, 600, 400);
            image2.Save("D:\\test2.jpeg");

            pPicture = oDoc.Paragraphs.Last.Range;
            oDoc.InlineShapes.AddPicture("D:\\test2.jpeg", Range: pPicture);

            image1.Dispose();
            image2.Dispose();

            //КЛ************************
            chart4.Series[0].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Pie;
            chart4.Series[0].IsValueShownAsLabel = true; // метки становятся видимыми
            chart4.Series[0].LabelFormat = "{#0}"; // формат отображения "{#0}"
            chart4.Titles.Add("Кабельные линии электропередачи в одноцепном исчислении(км, %) 110 кВ");
            chart4.Titles[1].Font = new Font("Times New Roman", 12, FontStyle.Bold);
            chart4.Legends[0].Font = new Font("Times New Roman", 9);

            chart4.Series[0].Points.AddY(KL[0]);
            chart4.Series[0].Points.AddY(KL[1]);
            chart4.Series[0].Points.AddY(KL[2]);

            chart4.Series[0].Points[0].YValues[0] = KL[0];
            chart4.Series[0].Points[1].YValues[0] = KL[1];
            chart4.Series[0].Points[2].YValues[0] = KL[2];

            chart4.Series[0].Points[0].LegendText = "свыше 50 лет";
            chart4.Series[0].Points[1].LegendText = "от 36 до 50 лет";
            chart4.Series[0].Points[2].LegendText = "до 35 лет";

            chart4.ChartAreas[0].Area3DStyle.Enable3D = true;

            //Сохраняем епта
            chart4.SaveImage(@"D:\\test1.jpeg", System.Drawing.Imaging.ImageFormat.Jpeg);

            pathFileImage = "D:\\test1.jpeg";

            // Загружаем исходное изображение
            image1 = new Bitmap(pathFileImage);

            // Масштабируем до нужного размера
            image2 = new Bitmap(image1, 600, 400);
            image2.Save("D:\\test2.jpeg");

            pPicture = oDoc.Paragraphs.Last.Range;
            oDoc.InlineShapes.AddPicture("D:\\test2.jpeg", Range: pPicture);

            image1.Dispose();
            image2.Dispose();

            System.IO.File.Delete(@"D:\\test1.jpeg");
            System.IO.File.Delete(@"D:\\test2.jpeg");
      

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
            String str;
            int rows = excelRange.Rows.Count;
            int cols = excelRange.Columns.Count;

            //////////////          СЧИТЫВАЕМ ПЕРВУЮ СТРОКУ, ИЩЕМ НАЗВАНИЕ           //////////////

            //////////////          ИЩЕМ PT           //////////////
            double num = 0;
            int rowOfNum = 0;
            String cell = "";
            for (int i = 1; i <= rows; i++)
            {
                cell = Convert.ToString(excelRange.Cells[i, 1].Value2);
                if (cell != null && cell.StartsWith("№"))
                {
                    rowOfNum = i;
                    break;
                }
            }
            cell = "1";
            remoteDB.Connect(DB);
            int lastWeb = rowOfNum;
            String nameOfWeb = "Сеть", nameOfLEP = "", nameOfTable = "";
            String[] pole = { "num", "name", "voltage", "checkNumber", "countOfChains", "length_all_oneChain", "length_all_allChain", "length_region_oneChain", "length_region_allChain", "", "", "stamp", "", "", "", "", "", "", "", "year", "isWork", "type" };
            String query = "CREATE TABLE ";
            for (int i = lastWeb + 1; i <= rows; i++)
            {
                cell = Convert.ToString(excelRange.Cells[i, 2].Value2);
                if (i == lastWeb + 1)
                    while (i <= rows)
                    {
                        cell = Convert.ToString(excelRange.Cells[i, 2].Value2);
                        if (cell == null && Convert.ToString(excelRange.Cells[i, 1].Value2) != null)
                        {
                            nameOfWeb = Convert.ToString(excelRange.Cells[i, 1].Value2);
                            nameOfTable = NameCompany.Text + '_' + "2019" + '_' + "ЛЭП" + '_' + nameOfWeb;
                            query = "CREATE TABLE \"" + nameOfTable + "\" (id INTEGER PRIMARY KEY, num DOUBLE, name STRING, voltage INTEGER, checkNumber STRING,countOfChains INTEGER, length_all_oneChain DOUBLE,length_all_allChain DOUBLE, length_region_oneChain DOUBLE, length_region_allChain DOUBLE,stamp STRING, year INTEGER, isWork BOOLEAN,type INTEGER);";
                            remoteDB.makeQuery(query);
                            break;
                        }
                        i++;
                    }
                if (cell == null && Convert.ToString(excelRange.Cells[i, 1].Value2) != null)////////ЕСЛИ ЕСТЬ СЕТЬ
                {
                    nameOfWeb = Convert.ToString(excelRange.Cells[i, 1].Value2);
                    nameOfTable = NameCompany.Text + '_' + "2019" + '_' + "ЛЭП" + '_' + nameOfWeb;
                    query = "CREATE TABLE \"" + nameOfTable + "\" (id INTEGER PRIMARY KEY, num DOUBLE, name STRING, voltage INTEGER, checkNumber STRING,countOfChains INTEGER, length_all_oneChain DOUBLE,length_all_allChain DOUBLE, length_region_oneChain DOUBLE, length_region_allChain DOUBLE,stamp STRING, year INTEGER, isWork BOOLEAN,type INTEGER);";
                    remoteDB.makeQuery(query);
                    continue;
                }
                if (cell != null)
                {
                    if (cell.StartsWith("ВЛ") || cell.StartsWith("КЛ") || cell.StartsWith("КВЛ") || cell.StartsWith("ВКЛ")) num = excelRange.Cells[i, 1].Value2;
                    nameOfLEP = cell;
                    remoteDB.makeQuery("INSERT INTO \"" + nameOfTable + "\" (\"" + pole[0] + "\",\"" + pole[1] + "\") " + "VALUES (\"" + num + "\", \"" + nameOfLEP + "\")");
                }
                if (nameOfLEP != "")
                {
                    DataTable res = remoteDB.getResTable("SELECT * FROM \"" + nameOfTable + "\" WHERE name = \"" + nameOfLEP + "\"");
                    for (int k = 3; k <= cols && k < pole.Length; k++)
                    {
                        cell = Convert.ToString(excelRange.Cells[i, k].Value2);
                        if (cell != null)
                        {
                            query = "UPDATE \"" + nameOfTable + "\" SET \"" + pole[k - 1] + "\" = \"";
                            if (res.Rows[0].ItemArray[k].ToString() == "")
                                query += cell + "\" WHERE name = \"" + nameOfLEP;
                            else
                                query += pole[k - 1] + "\" + \"" + cell + "\" WHERE name = \"" + nameOfLEP + "\"";
                            remoteDB.makeQuery(query);
                        }
                    }
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

        private void DBPicker_FileOk(object sender, CancelEventArgs e)
        {
            DB = ((OpenFileDialog)sender).FileName;
            localDB.makeQuery("update config set path = \"" + DB + "\" where id = 1");
        }

        private void выбратьБДToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DBPicker.ShowDialog();
        }
    }
}
