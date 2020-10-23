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
    public partial class idForm : Form
    {
        public user curuser;

        public idForm()
        {
            InitializeComponent();
        }

        //Вывод данных о пользователе
        private void idForm_Load(object sender, EventArgs e)
        {

            nameLabel.Text = "Имя: " + curuser.name;
            ageLabel.Text = "Возраст: " + curuser.age.ToString();
            statLabel.Text = "Должность: " + curuser.position;
            mailLabel.Text = "Почта: " + curuser.mail;
            expLabel.Text = "Стаж: " + curuser.experience.ToString();

            if (curuser.sex == 1)
                sexLabel.Text = "Пол: " + "Женский";
            else
                sexLabel.Text = "Пол: " + "Мужской";
        }
    }
}
