using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Net;
using System.Web;
using Newtonsoft.Json;

namespace window3
{
    class Analize
    {
        System.Net.WebRequest request;
        string postData;

        public Analize()
        {
        }
        public user login(string login, string password)
        {
            request = WebRequest.Create("http://nogy.onlinehacktomsk.ru/login.php");
            postData = "login=" + login + "&password=" + password;
            byte[] byteArray = Encoding.UTF8.GetBytes(postData);
            request.ContentLength = byteArray.Length;
            request.ContentType = "application/x-www-form-urlencoded";
            request.Method = "POST";
            user def = new user(); def.sex = 3;
            
            try
            {
                Stream dataStream = request.GetRequestStream();
                dataStream.Write(byteArray, 0, byteArray.Length);
                dataStream.Close();
                WebResponse response = request.GetResponse();
                using (dataStream = response.GetResponseStream())
                {
                    StreamReader reader = new StreamReader(dataStream);
                    string responseFromServer = reader.ReadToEnd();
                    response.Close();
                    reader.Close();
                    if (responseFromServer == "0") return def;
                    List<user> result = JsonConvert.DeserializeObject<List<user>>(responseFromServer);
                    return result[0];
                }
            }
            catch (System.Net.WebException e) {
                def.sex = 404;return def;
            };
        }
        public int registration(user nwusr, user curuser)
        {
            request = WebRequest.Create("http://nogy.onlinehacktomsk.ru/registration.php");
            postData = 
                   "name=" + nwusr.name
                + "&age=" + nwusr.age
                + "&sex=" + nwusr.sex
                + "&position=" + nwusr.position
                + "&experience=" + nwusr.experience
                + "&password=" + nwusr.password
                + "&login=" + nwusr.login
                + "&mail=" + nwusr.mail
                + "&perm_level=" + nwusr.perm_level
                + "&curlogin=" + curuser.login
                + "&curpassword=" + curuser.password;
            byte[] byteArray = Encoding.UTF8.GetBytes(postData);
            request.ContentLength = byteArray.Length;
            request.ContentType = "application/x-www-form-urlencoded";
            request.Method = "POST";
            try
            {
                Stream dataStream = request.GetRequestStream();
                dataStream.Write(byteArray, 0, byteArray.Length);
                dataStream.Close();
                WebResponse response = request.GetResponse();
                using (dataStream = response.GetResponseStream())
                {
                    StreamReader reader = new StreamReader(dataStream);
                    string responseFromServer = reader.ReadToEnd();
                    response.Close();
                    reader.Close();
                    if (responseFromServer == "0") return 0;
                    else if (responseFromServer == "1") return 1;
                    else if (responseFromServer == "2") return 2;
                    else if (responseFromServer == "OK") return 3;
                }
                return 5;
            }
            catch (System.Net.WebException e)
            {
                return 4;
            };
        }
        public List<devCommit> getDevData(string login, string password)
        {
            request = WebRequest.Create("http://nogy.onlinehacktomsk.ru/getdevdata.php");
            postData = "login=" + login + "&password=" + password;
            byte[] byteArray = Encoding.UTF8.GetBytes(postData);
            request.ContentLength = byteArray.Length;
            request.ContentType = "application/x-www-form-urlencoded";
            request.Method = "POST";
            List<devCommit> def = new List<devCommit>();

            try
            {
                Stream dataStream = request.GetRequestStream();
                dataStream.Write(byteArray, 0, byteArray.Length);
                dataStream.Close();
                WebResponse response = request.GetResponse();
                using (dataStream = response.GetResponseStream())
                {
                    StreamReader reader = new StreamReader(dataStream);
                    string responseFromServer = reader.ReadToEnd();
                    response.Close();
                    reader.Close();
                    if (responseFromServer == "0") return def;
                    List<devCommit> result = JsonConvert.DeserializeObject<List<devCommit>>(responseFromServer);                    return result;
                }
            }
            catch (System.Net.WebException e)
            {
                return def;
            };
        }

        public devCommit getCritData(string login, string password)
        {
            request = WebRequest.Create("http://nogy.onlinehacktomsk.ru/getcrit.php");
            postData = "login=" + login + "&password=" + password;
            byte[] byteArray = Encoding.UTF8.GetBytes(postData);
            request.ContentLength = byteArray.Length;
            request.ContentType = "application/x-www-form-urlencoded";
            request.Method = "POST";
            devCommit def = new devCommit();

            try
            {
                Stream dataStream = request.GetRequestStream();
                dataStream.Write(byteArray, 0, byteArray.Length);
                dataStream.Close();
                WebResponse response = request.GetResponse();
                using (dataStream = response.GetResponseStream())
                {
                    StreamReader reader = new StreamReader(dataStream);
                    string responseFromServer = reader.ReadToEnd();
                    response.Close();
                    reader.Close();
                    if (responseFromServer == "0") return def;
                    List<devCommit> result = JsonConvert.DeserializeObject<List<devCommit>>(responseFromServer);
                    return result[0];
                }
            }
            catch (System.Net.WebException e)
            {
                return def;
            };
        }
    }
}
