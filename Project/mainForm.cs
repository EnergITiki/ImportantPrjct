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

namespace window3
{

    public partial class mainForm : Form
    {
        
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
            Excel.Application excelApp = new Excel.Application();

            if (excelApp == null)
            {
                Console.WriteLine("Excel is not installed!!");
                return;
            }

            Excel.Workbook excelBook = excelApp.Workbooks.Open(@"E:\readExample.xlsx");
            Excel._Worksheet excelSheet = excelBook.Sheets[1];
            Excel.Range excelRange = excelSheet.UsedRange;

            int rows = excelRange.Rows.Count;
            int cols = excelRange.Columns.Count;

            for (int i = 1; i <= rows; i++)
            {
                // read new line
                for (int j = 1; j <= cols; j++)
                {
                    //write to cell
                    if (excelRange.Cells[i, j] != null && excelRange.Cells[i, j].Value2 != null) ;
                        //

                }
            }
            excelApp.Quit();
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
    }
}
