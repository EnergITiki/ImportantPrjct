using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using GMap.NET;
using GMap.NET.MapProviders;
using GMap.NET.WindowsForms;
using GMap.NET.WindowsForms.Markers;
using GMap.NET.WindowsForms.ToolTips;


namespace window3
{
    public partial class mapForm : Form
    {
        public mapForm()
        {
            InitializeComponent();
        }

        private void mapForm_Load(object sender, EventArgs e)
        {
            //Настройки для компонента GMap.
            map.Bearing = 0;

            //CanDragMap - Если параметр установлен в True,
            //пользователь может перетаскивать карту
            ///с помощью правой кнопки мыши.
            map.CanDragMap = true;

            //Указываем, что перетаскивание карты осуществляется
            //с использованием левой клавишей мыши.
            //По умолчанию - правая.
            map.DragButton = MouseButtons.Left;

            map.GrayScaleMode = true;

            //MarkersEnabled - Если параметр установлен в True,
            //любые маркеры, заданные вручную будет показаны.
            //Если нет, они не появятся.
            map.MarkersEnabled = true;

            //Указываем значение максимального приближения.
            map.MaxZoom = 18;

            //Указываем значение минимального приближения.
            map.MinZoom = 2;

            //Устанавливаем центр приближения/удаления
            //курсор мыши.
            map.MouseWheelZoomType = GMap.NET.MouseWheelZoomType.MousePositionAndCenter;

            //Отказываемся от негативного режима.
            map.NegativeMode = false;

            //Разрешаем полигоны.
            map.PolygonsEnabled = true;

            //Разрешаем маршруты
            map.RoutesEnabled = true;

            //Скрываем внешнюю сетку карты
            //с заголовками.
            map.ShowTileGridLines = false;

            //Указываем, что при загрузке карты будет использоваться
            //18ти кратное приближение.
            map.Zoom = 5;

            //Указываем что все края элемента управления
            //закрепляются у краев содержащего его элемента
            //управления(главной формы), а их размеры изменяются
            //соответствующим образом.
            map.Dock = DockStyle.Fill;

            //Указываем что будем использовать карты Google.
            map.MapProvider = GMap.NET.MapProviders.GMapProviders.GoogleMap;
            GMap.NET.GMaps.Instance.Mode = GMap.NET.AccessMode.ServerOnly;

            //Если вы используете интернет через прокси сервер,
            //указываем свои учетные данные.
            GMap.NET.MapProviders.GMapProvider.WebProxy = System.Net.WebRequest.GetSystemWebProxy();
            GMap.NET.MapProviders.GMapProvider.WebProxy.Credentials = System.Net.CredentialCache.DefaultCredentials;




        }



    }
}
