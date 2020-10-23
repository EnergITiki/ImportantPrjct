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

namespace window3
{

    public partial class mainForm : Form
    {
        bool flag = false;
        regForm reg = new regForm();
        loginForm login = new loginForm();
        idForm idform = new idForm();
        user curuser = new user();
        teleForm tele = new teleForm();
        mapForm map = new mapForm();

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
            mapButton.Visible = false;
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
            this.Hide();
            login.ShowDialog();          
            if (login.result.sex < 3)
            {
                curuser = login.result;
                
                buttonEnter.Text = "Телеметрия";
                buttonEnter.Click -= button1_Click;
                buttonEnter.Click += new System.EventHandler(next_button1_Click);

                idButton.Text = "Личный кабинет";
                idButton.Click += new System.EventHandler(idButton_Click_1);
                idButton.Visible = true;
                mapButton.Visible = true;
                if(curuser.perm_level==3)
                    regButton.Visible = true;
                nameLabel.Text = curuser.name;
            }
            this.Show();
        }

        //При нажатии открывается форма Телеметрия
        private void next_button1_Click(object sender, EventArgs e)
        {
            tele.curuser = curuser;
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
            idform.curuser = curuser;
            idform.ShowDialog();
        }

        //Завершение работы из трея
        private void закрытьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            flag = true;
            this.Close();
        }

        private void button1_Click_2(object sender, EventArgs e)
        {
            map.ShowDialog();
        }

        private void Button2_Click_1(object sender, EventArgs e)
        {
            this.Hide();
            reg.curuser = curuser;
            reg.ShowDialog();
            this.Show();
        }
    }
}
