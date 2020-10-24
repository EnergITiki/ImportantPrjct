using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;

namespace window3
{
    public partial class teleForm : Form
    {
        public user curuser = new user();
        List<devCommit> dataList = new List<devCommit>();
        bool flagToChange = true;
        reptForm rept = new reptForm();
        int a = 0, b = 0;

        public teleForm()
        {
            InitializeComponent();
        }

        //Закрытие окна
        public void teleForm_FormClosed(object sender, EventArgs e)
        {
            this.Hide();
        }


        private void draw()
        {
            int posY=0;

            //ПОИСК КРАЙНИХ ИНДЕКСОВ СПИСКА ПО ДАТЕ
            foreach (devCommit tempCom in dataList)
            {
                if (tempCom.dattim.CompareTo(dPicker1.Value) >= 0)
                {
                    a = dataList.IndexOf(tempCom);
                    break;
                }
            }
            foreach (devCommit tempCom in dataList)
            {
                if (tempCom.dattim.CompareTo(dPicker2.Value) == 0)
                {
                    b = dataList.IndexOf(tempCom);
                    break;
                }else if (tempCom.dattim.CompareTo(dPicker2.Value) > 0)
                {
                    b = dataList.IndexOf(tempCom)-1;
                }
            }
            infoBox.Controls.Clear();

            //ВЫВОД В ЗАВИСИМОСТИ ОТ ВЫБРАННОГО ТИПА
            //ВЫВОД ГРАФИКОВ/ДИАГРАММ
            if (typeCombo.SelectedIndex == 0 || typeCombo.SelectedIndex == 1)
            {
                foreach (ToolStripMenuItem curpar in par.DropDownItems)
                {
                    if (!curpar.Checked || curpar.Text == "Отметить все") continue;

                    Chart tempChart = new Chart();
                    ChartArea chartArea1 = new ChartArea();
                    Legend Legend1 = new Legend();
                    Title title = new Title();

                    title.Text = curpar.Text;
                    title.Font = new System.Drawing.Font("Century Gothic", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
                    chartArea1.Name = "chartArea_" + curpar.Name;
                    Legend1.Name = "legend_" + curpar.Name;
                    tempChart.Name = "chart_"+curpar.Name;
                    tempChart.Text = curpar.Text;

                    chartArea1.CursorX.IsUserEnabled = true;
                    chartArea1.CursorX.IsUserSelectionEnabled = true;
                    chartArea1.AxisX.ScaleView.Zoomable = true;
                    chartArea1.AxisX.ScrollBar.IsPositionedInside = false;

                    tempChart.Size = new System.Drawing.Size(infoBox.Width - 30,320);
                    tempChart.Location = new System.Drawing.Point(0,posY);posY += 350;
                    tempChart.Legends.Add(Legend1);
                    tempChart.Titles.Add(title);
                    tempChart.ChartAreas.Add(chartArea1);

                    infoBox.Controls.Add(tempChart);
                    
                    foreach (ToolStripMenuItem id in id_dev.DropDownItems)
                    {
                        if (!id.Checked || id.Text == "Отметить все") continue;
                        Series ser = new Series();
                        ser.Legend = "legend_" + curpar.Name;
                        ser.BorderWidth = 3;
                        ser.LegendText = id.Text;
                        ser.Name = "ser_"+curpar.Name+id.Text;
                        ser.ChartArea = "chartArea_" + curpar.Name;
                        if (typeCombo.SelectedIndex == 0)
                            ser.ChartType = SeriesChartType.Line;
                        else
                            ser.ChartType = SeriesChartType.Column;
                        tempChart.Series.Add(ser);
                    }
                }
                for(int i = a, count = 0; i < b+1; i++)
                {
                    devCommit tempCom = dataList[i];bool flag = false;
                    foreach (ToolStripMenuItem s  in id_dev.DropDown.Items) if (s.Checked && s.Text == tempCom.dev_id.ToString()) flag = true;
                    if (!flag) continue;
                    count++;
                    foreach (ToolStripMenuItem tempItem in par.DropDownItems) {
                        if (tempItem.Checked && tempItem.Text != "Отметить все") {
                            Chart ch = (infoBox.Controls["chart_" + tempItem.Name] as Chart);
                            Series s = (infoBox.Controls["chart_" + tempItem.Name] as Chart).Series["ser_" + tempItem.Name + tempCom.dev_id.ToString()];
                            s.Points.AddXY(tempCom.date, Convert.ToDouble(tempCom.GetField(tempItem.Name)));
                            if (tempItem.Name == "temp") Console.WriteLine(tempCom.temp);
                            if (s.Points.Count == 1)
                            {
                                //ch.ChartAreas["chartArea_" + tempItem.Name].AxisX.ScaleView.Zoom(0, 4);
                                ch.ChartAreas["chartArea_" + tempItem.Name].AxisY.Maximum = Convert.ToDouble(tempCom.GetField(tempItem.Name));
                                ch.ChartAreas["chartArea_" + tempItem.Name].AxisY.Minimum = Convert.ToDouble(tempCom.GetField(tempItem.Name));
                            }
                            if (ch.ChartAreas["chartArea_" + tempItem.Name].AxisY.Maximum < Convert.ToDouble(tempCom.GetField(tempItem.Name)))
                                ch.ChartAreas["chartArea_" + tempItem.Name].AxisY.Maximum = Convert.ToDouble(tempCom.GetField(tempItem.Name));
                            else if (ch.ChartAreas["chartArea_" + tempItem.Name].AxisY.Minimum > Convert.ToDouble(tempCom.GetField(tempItem.Name)))
                                ch.ChartAreas["chartArea_" + tempItem.Name].AxisY.Minimum = Convert.ToDouble(tempCom.GetField(tempItem.Name));
                        }
                    }
                }
            }

            //ВЫВОД ТАБЛИЦЫ С ИНФОРМЦИЕЙ

            if (typeCombo.SelectedIndex == 2)
            {
                DataGridView data = new DataGridView();
                data.Size = infoBox.Size;
                data.ColumnCount = 8;
                //data.RowCount = 1;
                data.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
                data.Columns[0].HeaderText = "ID"; data.Columns[0].Name = "dev_id";
                data.Columns[1].HeaderText = "Дата"; data.Columns[0].Name = "date";
                data.Columns[2].HeaderText = "Время"; data.Columns[0].Name = "time";
                data.ColumnHeadersHeight = 80;
                int count = 3;
                foreach (ToolStripMenuItem curpar in par.DropDownItems)
                {
                    if (!curpar.Checked || curpar.Text == "Отметить все") continue;
                    data.Columns[count].HeaderText = curpar.Text;
                    data.Columns[count].Name = curpar.Name;
                    count++;
                }
                data.ColumnCount = count;
                int rowCount = 0;
                for (int i = a; i < b + 1; i++)
                {
                    devCommit tempCom = dataList[i]; bool flag = false;
                    foreach (ToolStripMenuItem s in id_dev.DropDown.Items) if (s.Checked && s.Text == tempCom.dev_id.ToString()) flag = true;
                    if (!flag) continue;
                    data.RowCount++;
                    int colCount = 3;
                    data.Rows[rowCount].Cells[0].Value = tempCom.dev_id;
                    data.Rows[rowCount].Cells[1].Value = tempCom.date;
                    data.Rows[rowCount].Cells[2].Value = tempCom.time;
                    foreach (ToolStripMenuItem tempItem in par.DropDownItems)
                    {
                        if (tempItem.Checked && tempItem.Text != "Отметить все")
                        {
                            data.Rows[rowCount].Cells[colCount].Value = tempCom.GetField(tempItem.Name);
                            colCount++;
                        }
                    }
                    rowCount++;
                }
                //data.RowCount--;
                infoBox.Controls.Add(data);
            }

        }
        
        private void teleForm_Load(object sender, EventArgs e)
        {
            id_dev.DropDown.BackColor = Color.Black;
            par.DropDown.BackColor = Color.Black;
            par.Click += delegate
            {
                par.ForeColor = Color.Black;
            };
            id_dev.Click += delegate
            {
                id_dev.ForeColor = Color.Black;
            };
            //updTimer.Start();
            id_dev.DropDown.AutoClose = false;
            par.DropDown.AutoClose = false;
            id_dev.DropDown.MouseLeave += new System.EventHandler(TeleForm_Leave);
            par.DropDown.MouseLeave += new System.EventHandler(TeleForm_Leave);
            //dPicker1.Value = dataList[0].dattim;
            //dPicker2.Value = dataList[dataList.Count - 1].dattim;
        }

        private void par_checkAll_Click(object sender, EventArgs e)
        {
            flagToChange = false;
            foreach(ToolStripMenuItem someItem in par.DropDownItems)
            {
                if (someItem == par_checkAll) continue;
                if (par_checkAll.Checked)
                    someItem.Checked = true;
                else
                    someItem.Checked = false;
            }
            draw(); flagToChange = true;
        }

        private void dev_checkAll_Click(object sender, EventArgs e)
        {
            flagToChange = false;
            foreach (ToolStripMenuItem someItem in id_dev.DropDownItems)
            {
                if (someItem == dev_checkAll) continue;
                if (dev_checkAll.Checked)
                    someItem.Checked = true;
                else
                    someItem.Checked = false;
            }
            draw(); flagToChange = true;
        }

        private void всеПоля_CheckedChanged(object sender, EventArgs e)
        {
            if (flagToChange)
            {
                ToolStripMenuItem curItem = new ToolStripMenuItem();
                if (curItem.GetType() == sender.GetType())
                {
                    curItem = (ToolStripMenuItem)sender;
                    if (curItem.OwnerItem == id_dev) dev_checkAll.Checked = false;
                    if (curItem.OwnerItem == par) par_checkAll.Checked = false;
                }
                draw();
            }
        }

        private void devListUpdate()
        {
            List<string> somedevs = new List<string>();
            id_dev.DropDownItems.Clear();
            id_dev.DropDownItems.Add(dev_checkAll);
            for (int i = 0; i < dataList.Count; i++)
            {
                if (!somedevs.Contains(dataList[i].dev_id.ToString()))
                    somedevs.Add(dataList[i].dev_id.ToString());
            }
            for (int i = 0; i < somedevs.Count; i++)
            {
                ToolStripMenuItem newItem = new ToolStripMenuItem(somedevs[i]) { CheckOnClick = true, Name="id_"+somedevs[i], ForeColor = Color.White};
                newItem.CheckedChanged += new System.EventHandler(this.всеПоля_CheckedChanged);

                id_dev.DropDownItems.Add(newItem);
            }
        }

        private void TeleForm_Leave(object sender, EventArgs e)
        {
            par.ForeColor = Color.White;
            id_dev.ForeColor = Color.White;
            par.DropDown.Close();
            id_dev.DropDown.Close();
        }

        private void DPicker_ValueChanged(object sender, EventArgs e)
        {
            if (dPicker1.Value.CompareTo(dPicker2.Value) > 0) dPicker1.Value = dPicker2.Value;
            draw();
        }

        //Открытие формы отчета
        private void showButton_Click(object sender, EventArgs e)
        {
            rept.Close();
            rept = new reptForm();
            rept.dataList = dataList;
            rept.par = par;
            rept.id_dev = id_dev;
            rept.a = a; rept.b = b;
            WebRequests client = new WebRequests();
            rept.critParam = client.getCritData(curuser.login, curuser.password);
            rept.Show();
        }
    }
}
