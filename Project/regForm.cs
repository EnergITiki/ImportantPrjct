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
    public partial class regForm : Form
    {
        public user curuser;
        public regForm()
        {
            InitializeComponent();
            this.nameText.Size = new System.Drawing.Size(nameText.Size.Width, postComboBox.Size.Height);
            this.ageText.Size = new System.Drawing.Size(ageText.Size.Width, postComboBox.Size.Height);
            this.ageText.MaxLength = 2;
            this.expText.MaxLength = 2;

        }

        //Закрытие формы регистрации
        public void regForm_FormClosed(object sender, EventArgs e)
        {
            this.Hide();
        }

        //При нажатии на кнопку происходит проверка введенных данных,
        //а также отправка их на сервер
        private void regButton2_Click(object sender, EventArgs e)
        {
            int result = 0;
            try
            {
                user newuser = new user();
                newuser.name = nameText.Text;
                newuser.age = Convert.ToInt32(ageText.Text);
                newuser.login = logText.Text;
                newuser.password = passText.Text;
                newuser.mail = mailText.Text;
                newuser.perm_level = postComboBox.SelectedIndex;
                newuser.position = postComboBox.Text;
                newuser.experience = Convert.ToInt32(expText.Text);
                newuser.sex = sexComboBox.SelectedIndex;
                WebRequests client = new WebRequests();
                result = client.registration(newuser, curuser);
            }

            catch (Exception exp)
            {
                warnLabel.Text = "Проверьте введенные данные";
            }

            if ((nameText.Text == "") || (logText.Text == "") || (passText.Text == null) || (mailText.Text == null) || (postComboBox.Text == null) || (sexComboBox.Text == null) || (ageText.Text.ToString() == null)||(expText.Text == null))
                warnLabel.Text = "Проверьте введенные данные";
            else if(result == 3)
                this.Hide();
        }

        //Ограничение на ввод более 2 чисел
        private void ageText_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!(Char.IsDigit(e.KeyChar)) && (e.KeyChar != (char)Keys.Back) && (e.KeyChar != (char)Keys.Delete))
            {
                e.Handled = true;
            }
        }

        //Ограничение на ввод более 2 чисел
        private void expText_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!(Char.IsDigit(e.KeyChar)) && (e.KeyChar != (char)Keys.Back) && (e.KeyChar != (char)Keys.Delete))
            {
                e.Handled = true;
            }
        }

        //Ограничение на ввод чисел
        private void nameText_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!(Char.IsLetter(e.KeyChar)) && (e.KeyChar != (char)Keys.Back) && (e.KeyChar != (char)Keys.Delete) && !(Char.IsWhiteSpace(e.KeyChar)))
            {
                e.Handled = true;
            }
        }

        //Ограничение на ввод пробела
        private void logText_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((Char.IsWhiteSpace(e.KeyChar)) && (e.KeyChar != (char)Keys.Back) && (e.KeyChar != (char)Keys.Delete))
            {
                e.Handled = true;
            }
        }

        //Автоматический выбор в комбобоксах
        private void RegForm_Load(object sender, EventArgs e)
        {
            postComboBox.SelectedIndex = 0;
            postComboBox.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            sexComboBox.SelectedIndex = 0;
            sexComboBox.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
        }              
    }
}
