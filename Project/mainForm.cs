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
using System.Collections.Generic;
using System.Linq;

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
        String[] VLstamps = { "М-",
                            "А-" ,
                            "АС-" ,
                            "АСО-" ,
                            "АСУ-" ,
                            "АСК-" ,
                            "АН-" ,
                            "АЖ-" ,
                            "АСКП-" ,
                            "АСКС-" ,
                            "АССС " ,
                            "АПС-"}; 
        int pathButt = 0;

        public mainForm()
        {
            InitializeComponent();
        }


        private void mainForm_Load(object sender, EventArgs e)
        {
            ClientSize = new System.Drawing.Size(535, ClientSize.Height);
            localDB.Connect("local.db");
            if (localDB.isThereRes("SELECT name from sqlite_master where type= 'table'")){
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
            object oRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            oPara1 = oDoc.Content.Paragraphs.Add(ref oRng);
            oPara1.Range.Font.Name = "Times New Roman";
            oPara1.Range.Font.Size = size; // Размер шрифта
            oPara1.Range.Font.Bold = bold; // "Жирный" шрифт
            oPara1.Range.Font.Italic = Ital; // "Курсив" шрифт
            oPara1.Format.SpaceAfter = SpaceAfter; // оступ после параграфа
            oPara1.Range.Text = text;
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

        public void RES(string CompName, string FilialName) //РЭСы филиалов 
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

            chart1.Series[0].Points.Clear();

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
            image2.Save("D:\\test1_1.jpeg");

            var pPicture = oDoc.Paragraphs.Last.Range;

            image1.Dispose();
            image2.Dispose();

            //Трансформаторы*******************
            chart2.Series[0].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Pie;
            chart2.Series[0].IsValueShownAsLabel = true; // метки становятся видимыми
            chart2.Series[0].LabelFormat = "{#0}"; // формат отображения "{#0}"
            chart2.Titles.Add("Трансформаторы (МВА, %) 110 кВ");
            chart2.Titles[1].Font = new Font("Times New Roman", 12, FontStyle.Bold);
            chart2.Legends[0].Font = new Font("Times New Roman", 9);

            chart2.Series[0].Points.Clear();

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
            chart2.SaveImage(@"D:\\test2.jpeg", System.Drawing.Imaging.ImageFormat.Jpeg);

            pathFileImage = "D:\\test2.jpeg";

            // Загружаем исходное изображение
            image1 = new Bitmap(pathFileImage);

            // Масштабируем до нужного размера
            image2 = new Bitmap(image1, 600, 400);
            image2.Save("D:\\test2_1.jpeg");

            pPicture = oDoc.Paragraphs.Last.Range;

            image1.Dispose();
            image2.Dispose();

            //ВЛ************************
            chart3.Series[0].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Pie;
            chart3.Series[0].IsValueShownAsLabel = true; // метки становятся видимыми
            chart3.Series[0].LabelFormat = "{#0}"; // формат отображения "{#0}"
            chart3.Titles.Add("Воздушные линии электропередачи в одноцепном исчислении(км, %) 110 кВ");
            chart3.Titles[1].Font = new Font("Times New Roman", 12, FontStyle.Bold);
            chart3.Legends[0].Font = new Font("Times New Roman", 9);

            chart3.Series[0].Points.Clear();

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
            chart3.SaveImage(@"D:\\test3.jpeg", System.Drawing.Imaging.ImageFormat.Jpeg);

            pathFileImage = "D:\\test3.jpeg";

            // Загружаем исходное изображение
            image1 = new Bitmap(pathFileImage);

            // Масштабируем до нужного размера
            image2 = new Bitmap(image1, 600, 400);
            image2.Save("D:\\test3_1.jpeg");

            pPicture = oDoc.Paragraphs.Last.Range;

            image1.Dispose();
            image2.Dispose();

            //КЛ************************
            chart4.Series[0].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Pie;
            chart4.Series[0].IsValueShownAsLabel = true; // метки становятся видимыми
            chart4.Series[0].LabelFormat = "{#0}"; // формат отображения "{#0}"
            chart4.Titles.Add("Кабельные линии электропередачи в одноцепном исчислении(км, %) 110 кВ");
            chart4.Titles[1].Font = new Font("Times New Roman", 12, FontStyle.Bold);
            chart4.Legends[0].Font = new Font("Times New Roman", 9);

            chart4.Series[0].Points.Clear();

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
            chart4.SaveImage(@"D:\\test4.jpeg", System.Drawing.Imaging.ImageFormat.Jpeg);

            pathFileImage = "D:\\test4.jpeg";

            // Загружаем исходное изображение
            image1 = new Bitmap(pathFileImage);

            // Масштабируем до нужного размера
            image2 = new Bitmap(image1, 600, 400);
            image2.Save("D:\\test4_1.jpeg");

            pPicture = oDoc.Paragraphs.Last.Range;

            image1.Dispose();
            image2.Dispose();

            Bitmap bmp1 = new Bitmap(@"D:\\test1_1.jpeg"); //путь к твоей картинке
            int bmp1_width = bmp1.Width;
            int bmp1_height = bmp1.Height;

            Bitmap bmp2 = new Bitmap(@"D:\\test2_1.jpeg"); //путь к твоей картинке
            Bitmap bmp3 = new Bitmap(@"D:\\test3_1.jpeg"); //путь к твоей картинке
            Bitmap bmp4 = new Bitmap(@"D:\\test4_1.jpeg"); //путь к твоей картинке
            Bitmap final_bmp = new Bitmap(bmp1_width * 2, bmp1_height * 2);


            Graphics g = Graphics.FromImage(final_bmp);
            g.DrawImage(bmp1, 0, 0, bmp1_width, bmp1_height);
            g.DrawImage(bmp2, bmp1_width, 0, bmp1_width, bmp1_height);
            g.DrawImage(bmp3, 0, bmp1_height, bmp1_width, bmp1_height);
            g.DrawImage(bmp4, bmp1_width, bmp1_height, bmp1_width, bmp1_height);
            g.Dispose();

            final_bmp.Save("D:\\Result.jpeg");
            Image newImage = Image.FromFile("D:\\Result.jpeg");
           
            Clipboard.SetImage(newImage);

            Word.Range wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;

            // Вставит изображение из буфера обмена в конец документа

            wrdRng.Paste();

            bmp1.Dispose();
            bmp2.Dispose();
            bmp3.Dispose();
            bmp4.Dispose();
            final_bmp.Dispose();

            Word.Paragraph oPara1;
            object oRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            oPara1 = oDoc.Content.Paragraphs.Add(ref oRng);
            oPara1.Indent();

            PrintText("\n", 1, 0, 0, 1);

            str = "Рисунок " + PictureNum.ToString() + " - возрастная характеристика ВЛ, КЛ и ПС 110 кВ '" +
                FilialName + "' на 01.01.2019 г.";
            PrintText(str, 12, 1, 0, 5);

            PictureNum += 0.1;
            newImage.Dispose();
            File.Delete("D:\\Result.jpeg");
        }
        public void Filials(string CompName)//Филиалы
        {
            DataTable res = remoteDB.getResTable("select * from sqlite_master where type = 'table'");
            List<string> temp = new List<string>();
            for (int i = 0; i < res.Rows.Count; ++i)
            {
                String g = res.Rows[i].ItemArray[1].ToString().Split('_')[0];
                if (g == CompName)
                {
                    temp.Add(res.Rows[i].ItemArray[1].ToString().Split('_')[res.Rows[i].ItemArray[1].ToString().Split('_').Length - 1]);
                }
            }


            for (int i = 0; i < temp.Count; i++)
            {
                string FilialName = temp[i];

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
                arrNameLP[1] = "Какие то другие цепи";

                str = "Наиболее продолжительное время находятся в эксплуатации следующие линии " +
                    "электропередачи 110 кВ: " + arrNameLP[0] + ", " + arrYearLP[0].ToString() + " лет, " +
                    arrNameLP[1] + ", " + arrYearLP[1].ToString() + "лет.";
                PrintText(str, 12, 0, 0, 5);



                //Еще один цикл нахуй, все РЭСы просматриваем
                RES(CompName, FilialName);
            }

        }
        bool was = false;

        public void Companies(string CompName)// компании
        {
            if (!was)
            {
                Word.Paragraph oPara1;
                oPara1 = oDoc.Content.Paragraphs.Add();
                oPara1.Range.Text = "\t'" + CompName + "'";
                oPara1.Range.Font.Name = "Times New Roman";
                oPara1.Range.Font.Size = 12; // Размер шрифта
                oPara1.Range.Font.Bold = 1; // "Жирный" шрифт
                oPara1.Range.Font.Italic = 0; // "Курсив" шрифт
                oPara1.Format.SpaceAfter = 1; // оступ после параграфа
                oPara1.Range.InsertParagraphAfter();
                oPara1.Range.Font.Bold = 0; // "Жирный" шрифт
                oPara1.Range.Font.Italic = 0; // "Курсив" шрифт
                was = true;
            }
            else
            {
                PrintText("\t'" + CompName + "'", 12, 1, 0, 1);
            }


            string str = "\t" + "Протяженность ВЛ 110 кВ и КЛ 110 кВ, количество и суммарная мощность ПС " +
            "110 кВ, находящихся в собственности " + CompName + ", по состоянию на " +
            "01.01.2019 г. составили:";

            PrintText(str, 12, 0, 0, 5);

            //обращаемся к БД
            double VL = 0, KL = 0, P = 0;
            int count = 0;

            DataTable res = remoteDB.getResTable("select * from sqlite_master where type = 'table'");
            List<string> temp = new List<string>();
            for (int i = 0; i < res.Rows.Count; ++i)
            {
                String g = res.Rows[i].ItemArray[1].ToString().Split('_')[0];
                if (g == CompName)
                {
                    temp.Add(res.Rows[i].ItemArray[1].ToString());
                }
            }

            //VL
            for (int i = 0; i < temp.Count; i++)
            {
                DataTable lenTable = remoteDB.getResTable("select length_all_oneChain from \"" + temp[i] + "\" where type = 1");
                for (int j = 0; j < lenTable.Rows.Count; j++)
                {
                    if (lenTable.Rows[j].ItemArray[0].ToString() != "")
                    {
                        VL += Convert.ToDouble(lenTable.Rows[j].ItemArray[0].ToString());
                    }
                }
            }

            for (int i = 0; i < temp.Count; i++)
            {
                DataTable lenTable = remoteDB.getResTable("select length_region_oneChain from \"" + temp[i] + "\" where type = 0");
                for (int j = 0; j < lenTable.Rows.Count; j++)
                {
                    if (lenTable.Rows[j].ItemArray[0].ToString() != "")
                    {
                        VL += Convert.ToDouble(lenTable.Rows[j].ItemArray[0].ToString());
                    }
                }
            }

            //KL
            /*for (int i = 0; i < temp.Count; i++)
            {
                DataTable lenTable = remoteDB.getResTable("select length_all_oneChain from " + temp[i] + " where type = 2");
                for (int j = 0; j < lenTable.Rows.Count; j++)
                {
                    VL += Convert.ToDouble(lenTable.Rows[j].ItemArray[0].ToString());
                }
            }

            for (int i = 0; i < temp.Count; i++)
            {
                DataTable lenTable = remoteDB.getResTable("select length_region_oneChain from " + temp[i] + " where type = ''");
                for (int j = 0; j < lenTable.Rows.Count; j++)
                {
                    VL += Convert.ToDouble(lenTable.Rows[j].ItemArray[0].ToString());
                }
            }*/









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

            // Вставка текста в начало документа и отступа послe
            string str = "Анализ технического состояния и возрастная " +
            "структура линий электропередачи и подстанций";
            PrintText(str, 12, 1, 0, 10);

            //Обращаемся к БД, узнаем количество компаний
            remoteDB.Connect(DB);

            DataTable res = remoteDB.getResTable("select name from sqlite_master where type = 'table'");
            List<string> temp = new List<string>();
            for (int i = 0; i < res.Rows.Count; ++i)
            {
                String g = res.Rows[i].ItemArray[0].ToString().Split('_')[0];
                //String g = res.Rows[i].ItemArray[0].ToString();
                temp.Add(g);
            }

            List<string> compNames = temp.Distinct().ToList();

            for (int i = 0; i < compNames.Count; i++)
            {
                Companies(compNames[i]);
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
            int id = 1;
            String nameOfWeb = "Сеть", nameOfTable = "", nameOfLEP = "", stampOfLEP = "";
            int idOfLEP = 1;
            String[] pole = { "num", "name", "voltage", "checkNumber", "countOfChains", "length_all_oneChain", "length_all_allChain", "length_region_oneChain", "length_region_allChain", "", "", "stamp", "", "", "", "", "", "", "", "year", "", "", "", "", "", "isWork" };
            String query = "CREATE TABLE ";
            for (int i = lastWeb + 1; i <= rows; i++)
            {
                cell = Convert.ToString(excelRange.Cells[i, 2].Value2);
                if (i == lastWeb + 1)
                    while (i <= rows)/////////// ПЕРВОНАЧАЛЬНО ИЩЕМ СЕТЬ
                    {
                        cell = Convert.ToString(excelRange.Cells[i, 2].Value2);
                        if (cell == null && Convert.ToString(excelRange.Cells[i, 1].Value2) != null)
                        {
                            nameOfWeb = Convert.ToString(excelRange.Cells[i, 1].Value2);
                            nameOfTable = NameCompany.Text + '_' + "2019" + '_' + "ЛЭП" + '_' + nameOfWeb;
                            query = "CREATE TABLE '" + nameOfTable + "' (id INTEGER PRIMARY KEY, num DOUBLE, name STRING, voltage INTEGER, checkNumber STRING,countOfChains INTEGER, length_all_oneChain DOUBLE,length_all_allChain DOUBLE, length_region_oneChain DOUBLE, length_region_allChain DOUBLE,stamp STRING, year INTEGER, isWork BOOLEAN,type INTEGER);";
                            remoteDB.makeQuery(query);
                            break;
                        }
                        i++;
                    }
                else if (cell == null && Convert.ToString(excelRange.Cells[i, 1].Value2) != null)////////ЕСЛИ ЕСТЬ СЕТЬ
                {
                    nameOfWeb = Convert.ToString(excelRange.Cells[i, 1].Value2);
                    nameOfTable = NameCompany.Text + '_' + "2019" + '_' + "ЛЭП" + '_' + nameOfWeb;
                    query = "CREATE TABLE '" + nameOfTable + "' (id INTEGER PRIMARY KEY, num DOUBLE, name STRING, voltage INTEGER, checkNumber STRING,countOfChains INTEGER, length_all_oneChain DOUBLE,length_all_allChain DOUBLE, length_region_oneChain DOUBLE, length_region_allChain DOUBLE,stamp STRING, year INTEGER, isWork BOOLEAN,type INTEGER);";
                    nameOfLEP = "";

                    id = 1;
                    remoteDB.makeQuery(query);
                    continue;
                }
                else if (cell != null)
                {
                    try
                    {
                        if (Convert.ToString(excelRange.Cells[i, 1].Value2) != null) num = Convert.ToDouble(excelRange.Cells[i, 1].Value2);
                    }
                    catch (Exception qwe)
                    {
                        String temp = Convert.ToString(excelRange.Cells[i, 1].Value2);
                        num = Convert.ToDouble(temp.Split('.')[0] + ',' + temp.Split('.')[1]);
                    }
                    int type = -1;
                    if (cell.StartsWith("ВЛ")) type = 1;
                    else if (cell.StartsWith("КЛ")) type = 2;
                    else if (cell.StartsWith("ВКЛ") || cell.StartsWith("КВЛ")) 
                        type = 0;
                    nameOfLEP = cell;
                    idOfLEP = id;
                    if(type >= 0)
                        remoteDB.makeQuery("INSERT INTO '" + nameOfTable + "' ('" + pole[0] + "','" + pole[1] + "','type') " + "VALUES ('" + num + "', '" + nameOfLEP + "','" + type + "')");
                    else remoteDB.makeQuery("INSERT INTO '" + nameOfTable + "' ('" + pole[0] + "','" + pole[1] + "') " + "VALUES ('" + num + "', '" + nameOfLEP + "')");
                    id++;
                }
                if (nameOfLEP != "")
                {
                    DataTable res = remoteDB.getResTable("SELECT * FROM '" + nameOfTable + "' WHERE id = '" + idOfLEP + "'");
                    for (int k = 3, val = 0; k <= cols && k < pole.Length && res.Rows.Count > 0; k++)
                    {
                        if (pole[k - 1] == "") continue;
                        cell = Convert.ToString(excelRange.Cells[i, k].Value2);
                        if (cell != null)
                        {
                            query = "UPDATE '" + nameOfTable + "' SET " + pole[k - 1] + " = ";

                            if (res.Rows[0].ItemArray[val + 3].ToString() == "")
                            {
                                query += "'" + cell + "' WHERE id = '" + idOfLEP + "'";
                            }
                            else
                            {
                                String temp = res.Rows[0].ItemArray[val + 3].ToString() + ";" + cell;
                                query += "'" + temp + "' WHERE id = '" + idOfLEP + "'";
                            }
                            remoteDB.makeQuery(query);
                            if(pole[k - 1] == "stamp")
                            {
                                bool check = false;
                                foreach (String st in VLstamps)
                                {
                                    if (cell.Contains(st)) check = true;
                                }
                                if (nameOfLEP.StartsWith("ВКЛ") || nameOfLEP.StartsWith("КВЛ")) ;
                                else if (check)
                                    remoteDB.makeQuery("UPDATE '" + nameOfTable + "' SET type = 1 WHERE id = '" + idOfLEP + "'");
                                else
                                    remoteDB.makeQuery("UPDATE '" + nameOfTable + "' SET type = 2 WHERE id = '" + idOfLEP + "'");
                            }
                        }
                        val++;
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
            localDB.makeQuery("update config set path = '" + DB + "' where id = 1");
        }

        private void выбратьБДToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DBPicker.ShowDialog();
        }
    }
}
