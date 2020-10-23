using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace window3
{
    public partial class reptForm : Form
    {
        public List<devCommit> dataList = new List<devCommit>();
        public ToolStripMenuItem id_dev;
        public ToolStripMenuItem par;
        public int a, b;
        public devCommit critParam;
        public reptForm()
        {
            InitializeComponent();
        }

        private void Button1_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void reptForm_Load(object sender, EventArgs e)
        {
            data.Rows.Clear();
            data.ColumnCount = 5;
            data.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
            data.Columns[0].HeaderText = "ID"; data.Columns[1].HeaderText = "Дата";
            data.Columns[2].HeaderText = "Время"; data.Columns[3].HeaderText = "Событие";
            data.Columns[4].HeaderText = "Значение - Крит. значение";
            data.Columns[0].Width = 40; data.Columns[3].Width = 300;
            int rowCount = 0;
            for (int i = a; i < b+1; i++)
            {
                devCommit tempCom = dataList[i]; bool flag = false;
                foreach (ToolStripMenuItem s in id_dev.DropDown.Items) if (s.Checked && s.Text == tempCom.dev_id.ToString()) flag = true;
                if (!flag) continue;
                foreach (ToolStripMenuItem tempItem in par.DropDownItems)
                {
                    if (tempItem.Checked && tempItem.Text != "Отметить все")
                    {
                        int whatHappened = 0;
                        switch (tempItem.Name)
                        {
                            case "temp":
                                if (tempCom.temp >= 0.95 * critParam.temp && tempCom.temp < critParam.temp) whatHappened = 1;
                                else if (tempCom.temp >= critParam.temp) whatHappened = 2;
                                break;
                            case "vibr":
                                if (tempCom.vibr >= 0.95 * critParam.vibr && tempCom.vibr < critParam.vibr) whatHappened = 1;
                                else if (tempCom.vibr >= critParam.vibr) whatHappened = 2;
                                break;
                            case "loading":
                                if (tempCom.loading >= 0.80 * critParam.loading && tempCom.loading < 0.95 * critParam.loading) whatHappened = 1;
                                else if (tempCom.loading >= 0.95 * critParam.loading) whatHappened = 2;
                                break;
                            case "power":
                                if (tempCom.power >= 0.95 * critParam.power && tempCom.power < critParam.power) whatHappened = 1;
                                else if (tempCom.power >= critParam.power) whatHappened = 2;
                                break;
                            case "work_time":
                                if (tempCom.work_time >= 0.85 * critParam.work_time && tempCom.work_time < critParam.work_time) whatHappened = 1;
                                else if (tempCom.work_time >= critParam.work_time) whatHappened = 2;
                                break;
                        }
                        if (whatHappened > 0)
                        {
                            data.RowCount++;
                            data.Rows[rowCount].Height = 60;
                            data.Rows[rowCount].Cells[0].Value = tempCom.dev_id;
                            data.Rows[rowCount].Cells[1].Value = tempCom.date;
                            data.Rows[rowCount].Cells[2].Value = tempCom.time;
                            data.Rows[rowCount].Cells[3].Value = "Значение параметра \"" + tempItem.Text + "\" ";
                            if (whatHappened == 1) data.Rows[rowCount].Cells[3].Value += "приближается к критическому";
                            else
                            {
                                data.Rows[rowCount].Cells[3].Value += "превысило критическое";
                                data.Rows[rowCount].Cells[3].Style.BackColor = Color.Red;
                            }
                            data.Rows[rowCount].Cells[4].Value = tempCom.GetField(tempItem.Name) + " - " + critParam.GetField(tempItem.Name);
                            rowCount++;
                        }
                    }
                }
            }
        }
    }
}
