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
    public partial class loginForm : Form
    {
        public user result = new user();

        public loginForm()
        {
            InitializeComponent();

            result.sex = 3;
            this.passText.AutoSize = false;
            this.passText.Size = new Size(this.passText.Size.Width, logText.Size.Height);
        }
        
        //Закрытие формы входа
        public void loginForm_FormClosed(object sender, EventArgs e)
        {
            this.Hide();    
        }

        //Вход в учетную запись
        //Проверка правильности введенных данных
        private void buttonEnter_Click(object sender, EventArgs e)
        {
            WebRequests client = new WebRequests();
            result = client.login(logText.Text, passText.Text);
            if (result.sex == 3) infoLabel.Text = "Неверный логин или пароль";
            else if (result.sex == 404) infoLabel.Text = "Ошибка в интернет соединении";
            else this.Hide();
        }

        //Изменение графических свойств полей ввода пароля и логина
        private void LoginForm_Load(object sender, EventArgs e)
        {
            Paint += new PaintEventHandler(loginForm_Paint);
        }

        void loginForm_Paint(object sender, PaintEventArgs e)
        {
            Graphics g = this.CreateGraphics();
            Pen pen = new Pen(Brushes.White, 7);
            pen.LineJoin = System.Drawing.Drawing2D.LineJoin.Bevel;//задаем скошенные углы
            pen.MiterLimit = 4;//задаем ограничение толщины скошенных углов
            g.DrawRectangle(pen, new Rectangle(logText.Location.X - 1, logText.Location.Y - 1, logText.Width , logText.Height ));

            g.DrawRectangle(pen, new Rectangle(passText.Location.X - 1, passText.Location.Y - 1, passText.Width, passText.Height));

        }

    }
}
