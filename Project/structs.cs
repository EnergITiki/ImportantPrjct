using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace window3
{
    public struct user
    {
        public int id;
        public string name;
        public int age;
        public int sex;
        public string position;
        public int experience;
        public string password;
        public string login;
        public string mail;
        public int perm_level;
        public void nul() { id = 0; name = ""; sex = 0; position = ""; experience = 0; password = ""; login = ""; perm_level = 0; }
    };
    public struct devCommit
    {
        public int id;
        public int dev_id;
        public string date;
        public string time;
        public float temp;
        public float vibr;
        public float power;
        public float loading;
        public float work_time;
        public System.DateTime dattim;
        public void getDate()
        {
            string[] datArr = date.Split('.');
            string[] timArr = time.Split(':');
            dattim = new System.DateTime(Convert.ToInt32(datArr[2]), Convert.ToInt32(datArr[1]), Convert.ToInt32(datArr[0]), 
                Convert.ToInt32(timArr[0]), Convert.ToInt32(timArr[1]), 0);
        }
        public string GetField(string name)
        {
            return typeof(devCommit).GetField(name).GetValue(this).ToString();
        }
    }
}
