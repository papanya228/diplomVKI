using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using GMap.NET.MapProviders;
using GMap.NET.WindowsForms;
using GMap.NET.WindowsForms.Markers;

namespace soft1
{
    public partial class Form1 : Form
    {
        private Graphics gPanel2;
        private List<PointInfo> massPoint = new List<PointInfo>();
        private DataGridView dataStorage = new DataGridView();
        private DataGridView dataStorageAnalyzes = new DataGridView();
        private List<string> elements = new List<string>();
        private List<GMapMarker> selectedPoint = new List<GMapMarker>();
        private List<GMapMarker> allMarkers = new List<GMapMarker>();
        private bool optionChoice = true;
        private bool optionCalculateButton = false;

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            gMapControl1.Bearing = 0;
            gMapControl1.ShowCenter = false;
            gMapControl1.MapProvider = GMapProviders.GoogleMap;
            //GMap.NET.GMaps.Instance.Mode = GMap.NET.AccessMode.ServerOnly;
            gMapControl1.DragButton = MouseButtons.Left;
            gMapControl1.MouseWheelZoomType = GMap.NET.MouseWheelZoomType.MousePositionWithoutCenter;
            gMapControl1.MaxZoom = 20;

            gMapControl1.Position = new GMap.NET.PointLatLng(56.0184, 92.9672);
            //gMapControl1.Position = new GMap.NET.PointLatLng(43.9783, 15.3834);
            //gMapControl1.Zoom = 16;

            gPanel2 = panel2.CreateGraphics();

            try
            {
                using (DBhelper sqlH = new DBhelper())
                {

                    dataStorage = sqlH.ReadCommand("Select * from Points");
                    markerFromGrid();
                    dataStorageAnalyzes = sqlH.ReadCommand("Select * from Analyzes");
                    getElements();
                    setToolText();
                }
            }
            catch (Exception)
            {

            }

        }

        private void getElements()
        {
            if (dataStorageAnalyzes.Rows.Count > 0)
            {
                comboBox1.Items.Clear();
                elements.Clear();
                for (int i = 0; i < dataStorageAnalyzes.Rows.Count - 1; i++)
                {
                    string findedElement = elements.Find(item => item == dataStorageAnalyzes["Element", i].Value.ToString());
                    if (findedElement == null)
                    {
                        elements.Add(dataStorageAnalyzes["Element", i].Value.ToString());
                        comboBox1.Items.Add(dataStorageAnalyzes["Element", i].Value.ToString());
                    }

                }

            }

        }

       private void setToolText()
       {
            string selectedText = comboBox1.Text;
            if (selectedText == "")
            {
                foreach (var items in gMapControl1.Overlays)
                {
                    if (items.Markers.Count > 0)
                    {
                        items.Markers[0].ToolTipText = items.Markers[0].Tag.ToString();
                        for (int i = 0; i < dataStorageAnalyzes.Rows.Count - 1; i++)
                        {
                            if (dataStorageAnalyzes["Point", i].Value.ToString() == items.Markers[0].Tag.ToString())
                            {
                                items.Markers[0].ToolTipText += Environment.NewLine + dataStorageAnalyzes["Element", i].Value.ToString() + " = "
                                + dataStorageAnalyzes["Analyz", i].Value.ToString();
                            }

                        }
                    }
                }
            }
            else
            {
                foreach (var items in gMapControl1.Overlays)
                {
                    if (items.Markers.Count > 0)
                    {
                        items.Markers[0].ToolTipText = items.Markers[0].Tag.ToString();
                        for (int i = 0; i < dataStorageAnalyzes.Rows.Count - 1; i++)
                        {
                            if (dataStorageAnalyzes["Point", i].Value.ToString() == items.Markers[0].Tag.ToString() && dataStorageAnalyzes["Element", i].Value.ToString() == selectedText)
                            {
                                items.Markers[0].ToolTipText += Environment.NewLine + dataStorageAnalyzes["Element", i].Value.ToString() + " = "
                                + dataStorageAnalyzes["Analyz", i].Value.ToString();
                            }

                        }
                    }
                }
            }
       }

        private void clearFields()
        {
            List<GMapOverlay> deleteOverlays = new List<GMapOverlay>();
            foreach (var item in gMapControl1.Overlays)
            {
                if (item.Polygons.Count > 0)
                {
                    deleteOverlays.Add(item);
                }

                if (item.Markers.Count > 0)
                {
                    item.Markers[0].IsVisible = true;
                }
            }

            foreach (var item in deleteOverlays)
            {
                gMapControl1.Overlays.Remove(item);
            }

            button2.Text = "Расчитать";
            optionCalculateButton = false;

            gPanel2.Clear(Color.White);
            panel2.Controls.Clear();

            label3.Text = "В сезон: 0";
            label4.Text = "В год: 0";

            gMapControl1.Zoom++;
            gMapControl1.Zoom--;

        }

        private void PaintFields()
        {
            string element = "";
            if (elements.Find(str => str == comboBox1.Text) == null)
            {
                return;
            }
            else
            {
                element = comboBox1.Text;
            }

            List<List<double>> calcData = CalculationModule.Calc(selectedPoint, dataStorageAnalyzes, element);

            if (calcData != null)
            {


                double dist = CalculationModule.calculateTheDistance(CalculationModule.firstPos.Lat, CalculationModule.firstPos.Lng, CalculationModule.lastPos.Lat, CalculationModule.firstPos.Lng);
                double incrementX = Math.Abs((CalculationModule.firstPos.Lat - CalculationModule.lastPos.Lat) / (dist / CalculationModule.calcRangeMetr));
                dist = CalculationModule.calculateTheDistance(CalculationModule.firstPos.Lat, CalculationModule.firstPos.Lng, CalculationModule.firstPos.Lat, CalculationModule.lastPos.Lng);
                double incrementY = Math.Abs((CalculationModule.firstPos.Lng - CalculationModule.lastPos.Lng) / (dist / CalculationModule.calcRangeMetr));
                if (CalculationModule.firstPos.Lat > CalculationModule.lastPos.Lat) incrementX = -incrementX;
                if (CalculationModule.firstPos.Lng > CalculationModule.lastPos.Lng) incrementY = -incrementY;


                double maxValue = 0, minValue = 999999;

                List<Color> colorList = new List<Color>();

                colorList.Add(Color.Green);
                colorList.Add(System.Drawing.ColorTranslator.FromHtml("#00FF77"));
                colorList.Add(System.Drawing.ColorTranslator.FromHtml("#00FFEF"));
                colorList.Add(System.Drawing.ColorTranslator.FromHtml("#0089FF"));
                colorList.Add(System.Drawing.ColorTranslator.FromHtml("#0000FF"));
                colorList.Add(System.Drawing.ColorTranslator.FromHtml("#CD00FF"));
                colorList.Add(System.Drawing.ColorTranslator.FromHtml("#FF0080"));
                colorList.Add(Color.Red);

                int TransparencyValue = 80;

                int plusX = 0;
                for (double x = CalculationModule.firstPos.Lat; x > CalculationModule.lastPos.Lat; x = x + incrementX)
                {
                    int plusY = 0;
                    for (double y = CalculationModule.firstPos.Lng; y < CalculationModule.lastPos.Lng; y = y + incrementY)
                    {
                        double val = calcData[plusX][plusY];
                        if (val > maxValue) maxValue = val;
                        if (val < minValue) minValue = val;
                        plusY++;
                    }
                    plusX++;
                }

                double range = (maxValue - minValue) / colorList.Count;

                double sumSezon = 0;

                plusX = 0;
                for (double x = CalculationModule.firstPos.Lat; x > CalculationModule.lastPos.Lat; x = x + incrementX)
                {
                    int plusY = 0;
                    int color = -1;

                    for (int i = 0; i < colorList.Count; i++)
                    {
                        if(calcData[plusX][plusY] < minValue + (range * (i + 1)))
                        {
                            color = i;
                            break;
                        }
                    }
                    if(color == -1)
                    {
                        color = colorList.Count - 1;
                    }

                    double firstpos = CalculationModule.firstPos.Lng;

                    for (double y = CalculationModule.firstPos.Lng; y < CalculationModule.lastPos.Lng; y = y + incrementY)
                    {

                        int color2 = -1;

                        for (int i = 0; i < colorList.Count; i++)
                        {
                            if (calcData[plusX][plusY] < minValue + (range * (i + 1)))
                            {
                                color2 = i;
                                break;
                            }
                        }
                        if (color2 == -1)
                        {
                            color2 = colorList.Count - 1;
                        }

                        if (color != color2)
                        {
                            List<GMap.NET.PointLatLng> points2 = new List<GMap.NET.PointLatLng>();
                            points2.Add(new GMap.NET.PointLatLng(x - (incrementX / 2), firstpos - (incrementY / 2)));
                            points2.Add(new GMap.NET.PointLatLng(x + (incrementX / 2), firstpos - (incrementY / 2)));
                            points2.Add(new GMap.NET.PointLatLng(x + (incrementX / 2), y - incrementY + (incrementY / 2)));
                            points2.Add(new GMap.NET.PointLatLng(x - (incrementX / 2), y - incrementY + (incrementY / 2)));
                            GMapOverlay polygons2 = new GMapOverlay("polygons");
                            GMapPolygon polygon2 = new GMapPolygon(points2, "Запретный город");

                            polygon2.Fill = new SolidBrush(Color.FromArgb(TransparencyValue, colorList[color]));

                            polygon2.Stroke = new Pen(new SolidBrush(Color.FromArgb(0, Color.Red)));
                            polygons2.Polygons.Add(polygon2);
                            gMapControl1.Overlays.Add(polygons2);

                            color = color2;
                            firstpos = y;
                        }

                        sumSezon += calcData[plusX][plusY];
                        plusY++;
                    }

                    List<GMap.NET.PointLatLng> points = new List<GMap.NET.PointLatLng>();
                    points.Add(new GMap.NET.PointLatLng(x - (incrementX / 2), firstpos - (incrementY / 2)));
                    points.Add(new GMap.NET.PointLatLng(x + (incrementX / 2), firstpos - (incrementY / 2)));
                    points.Add(new GMap.NET.PointLatLng(x + (incrementX / 2), CalculationModule.lastPos.Lng + (incrementY / 2)));
                    points.Add(new GMap.NET.PointLatLng(x - (incrementX / 2), CalculationModule.lastPos.Lng + (incrementY / 2)));
                    GMapOverlay polygons = new GMapOverlay("polygons");
                    GMapPolygon polygon = new GMapPolygon(points, "Запретный город");

                    polygon.Fill = new SolidBrush(Color.FromArgb(TransparencyValue, colorList[color]));
                    polygon.Stroke = new Pen(new SolidBrush(Color.FromArgb(0, Color.Red)));
                    polygons.Polygons.Add(polygon);
                    gMapControl1.Overlays.Add(polygons);

                    plusX++;
                }

                calcData.Clear();
                calcData = null;

                foreach (var item in gMapControl1.Overlays)
                {
                    if (item.Markers.Count > 0 && selectedPoint.Find(itemm => itemm.Tag == item.Markers[0].Tag) == null)
                    {
                        item.Markers[0].IsVisible = false;
                    }
                }

                button2.Text = "Расчитать заново";
                optionCalculateButton = true;

                int positionY = 5;
                int positionX = 10;
                int widthrect = 20;
                int heightRect = 20;
                int incrementDraw = positionY;
                for (int i = 0; i < colorList.Count; i++)
                {
                    SolidBrush brush = new SolidBrush(Color.FromArgb(TransparencyValue, colorList[i]));
                    Rectangle rect = new Rectangle(positionX, incrementDraw, widthrect, heightRect);
                    gPanel2.FillRectangle(brush, rect);

                    Label labelBuf = new Label();
                    labelBuf.Location = new Point(positionX + 30, incrementDraw + 5);
                    labelBuf.Width = 160;
                     labelBuf.Text = Math.Round(minValue + (range * i), 4).ToString() + " < x < " + Math.Round(minValue + (range * (i + 1)), 4).ToString();
                    panel2.Controls.Add(labelBuf);

                    incrementDraw += 25;
                }

                sumSezon = Math.Round(sumSezon * Math.Pow(Convert.ToDouble(CalculationModule.calcRangeMetr) / 1000, 2) / 1000, 4);

                label3.Text = "В сезон: " + sumSezon.ToString() + " мг/дм^3";
                label4.Text = "В год: " + Math.Round((sumSezon * 365) / 120, 4).ToString() + " мг/дм^3";

                gMapControl1.Zoom++;
                gMapControl1.Zoom--;

            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            //Thread mythread = new Thread(Calc);
            //mythread.Start();
            //mythread.IsBackground = true;
            //gMapControl1.MapProvider = GMapProviders.GoogleSatelliteMap;



            //PaintFields();






            /*try
            {
                using (ExcelHelper exhelp = new ExcelHelper())
                {
                    if (exhelp.Open())
                    {

                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }*/
        }

        private void gMapControl1_OnMapClick(GMap.NET.PointLatLng PointClick, MouseEventArgs e)
        {
            //Console.WriteLine(Math.Round(PointClick.Lat, 4));
            //Console.WriteLine(Math.Round(PointClick.Lng, 4));
            //Console.WriteLine();

            //GMapMarker marker = new GMarkerGoogle(new GMap.NET.PointLatLng(Math.Round(PointClick.Lng, 4), Math.Round(PointClick.Lat, 4)), GMarkerGoogleType.pink);
            ////marker.ToolTipText = "1/1";
            //marker.ToolTipMode = MarkerTooltipMode.Always;
            //GMapOverlay markers = new GMapOverlay("markers");
            //markers.Markers.Add(marker);
            //gMapControl1.Overlays.Add(markers);

            //StreamWriter sw = new StreamWriter("C:\\Test.txt", true);

            //sw.Write(Math.Round(PointClick.Lng, 4));
            //sw.Write((char)9);
            //sw.Write(Math.Round(PointClick.Lat, 4));
            //sw.WriteLine("");

            //sw.Close();

            //gMapControl1.Zoom++;
            //gMapControl1.Zoom--;
        }


        private void markerFromGrid()
        {
            if (dataStorage != null)
            {
                for (int i = 0; i < dataStorage.Rows.Count - 1; i++)
                {
                    PointInfo poin = new PointInfo();
                    poin.Point = dataStorage["Point", i].Value.ToString();
                    poin.Lng = dataStorage["Lng", i].Value.ToString();
                    poin.Lat = dataStorage["Lat", i].Value.ToString();
                    addMarker(poin);
                }
            }
        }

        private void addMarker(PointInfo poin)
        {
            if(poin.Lat != null && poin.Lng != null && poin.Point != null)
            {
                GMapMarker marker = new GMarkerGoogle(new GMap.NET.PointLatLng(Math.Round(Convert.ToDouble(poin.Lng), 4), Math.Round(Convert.ToDouble(poin.Lat), 4)), GMarkerGoogleType.green);
                marker.ToolTipText = poin.Point;
                //marker.ToolTipMode = MarkerTooltipMode.Always;
                marker.Tag = poin.Point;
                GMapOverlay markers = new GMapOverlay("markers");
                markers.Markers.Add(marker);
                gMapControl1.Overlays.Add(markers);
                massPoint.Add(poin);

                allMarkers.Add(marker);

                gMapControl1.Zoom++;
                gMapControl1.Zoom--;
            }

        }

        private void gMapControl1_OnMarkerClick(GMapMarker item, MouseEventArgs e)
        {
            if(checkBox1.Checked == true)
            {
                selectMarker(item);
            }
            else
            {
                viewAnalyzes(item.Tag.ToString());
            }
        }

        private void viewAnalyzes(string Tag)
        {
            dataGridView1.Rows.Clear();
            for (int i = 0; i < dataStorageAnalyzes.Rows.Count - 1; i++)
            {
                if (dataStorageAnalyzes["Point", i].Value.ToString() == Tag)
                {
                    object[] bufData = new object[3];

                    bufData[0] = dataStorageAnalyzes["Point", i].Value.ToString();
                    bufData[1] = dataStorageAnalyzes["Element", i].Value.ToString();
                    bufData[2] = dataStorageAnalyzes["Analyz", i].Value.ToString();
                    dataGridView1.Rows.Add(bufData);
                }

            }
        }

        private void selectMarker(GMapMarker item = null)
        {
            if (item == null)
            {
                if (optionChoice)
                {
                    foreach (var items in gMapControl1.Overlays)
                    {
                        if (items.Markers.Count > 0 && selectedPoint.Find(T => T == items.Markers[0]) == null)
                        {
                            item = items.Markers[0];
                            choiceMarker(item, true);
                        }
                    }
                    optionChoice = false;
                }
                else
                {
                    foreach (var items in gMapControl1.Overlays)
                    {
                        if (items.Markers.Count > 0 && selectedPoint.Find(T => T == items.Markers[0]) != null)
                        {
                            item = items.Markers[0];
                            choiceMarker(item, false);
                        }
                    }
                    optionChoice = true;
                }
            }
            else
            {
                GMapMarker findedMarker = selectedPoint.Find(T => T == item);
                if (findedMarker == null)
                {
                    choiceMarker(item, true);
                }
                else
                {
                    choiceMarker(item, false);
                }
            }
        }

        private void choiceMarker(GMapMarker item, bool opt)
        {
            if (opt)
            {
                GMapMarker marker = new GMarkerGoogle(item.Position, GMarkerGoogleType.red);
                marker.ToolTipText = item.ToolTipText;
                marker.Tag = item.Tag;
                item.Overlay.Markers.Remove(item);
                item.Overlay.Markers.Add(marker);
                selectedPoint.Add(marker);
            }
            else
            {
                GMapMarker marker = new GMarkerGoogle(item.Position, GMarkerGoogleType.green);
                marker.ToolTipText = item.ToolTipText;
                marker.Tag = item.Tag;
                item.Overlay.Markers.Remove(item);
                item.Overlay.Markers.Add(marker);
                selectedPoint.Remove(item);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (optionCalculateButton)
            {
                clearFields();
            }
            else
            {
                PaintFields();
            }

        }

        private void button3_Click(object sender, EventArgs e)
        {
            ExcelHelper.SaveMap(gMapControl1.ToImage());
        }

        private void button5_Click(object sender, EventArgs e)
        {
            
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void comboBox1_TextChanged(object sender, EventArgs e)
        {
            setToolText();
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked == true)
            {
                button1.Visible = false;
                button2.Visible = true;
                button3.Visible = true;
                button6.Visible = true;
                button7.Visible = false;
                panel2.Visible = true;
                dataGridView1.Visible = false;
                label3.Visible = true;
                label4.Visible = true;
            }
            else
            {
                button1.Visible = true;
                button2.Visible = false;
                button3.Visible = false;
                button6.Visible = false;
                button7.Visible = true;
                panel2.Visible = false;
                dataGridView1.Visible = true;
                label3.Visible = false;
                label4.Visible = false;

                foreach (var items in gMapControl1.Overlays)
                {
                    if (items.Markers.Count > 0 && selectedPoint.Find(T => T == items.Markers[0]) != null)
                    {
                        GMapMarker item = items.Markers[0];
                        choiceMarker(item, false);
                    }
                }
                optionChoice = true;

                clearFields();
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                using (ExcelHelper exhelp = new ExcelHelper())
                {
                    if (exhelp.Open())
                    {
                        //string[,] data = exhelp.Read();
                        object[,] data = exhelp.Read();
                        if (data != null)
                        {
                            int row = data.GetLength(0);
                            int column = data.GetLength(1);
                            if (row < 2 || column < 1)
                            {
                                throw new Exception("В таблице нет данных!");
                            }
                            string bufCommand = "insert Points (";
                            List<string> bufColumn = new List<string>();

                            for (int j = 1; j <= column; j++)
                            {

                                switch (data[1, j])
                                {
                                    case "Point":
                                        bufColumn.Add(data[1, j].ToString());
                                        bufCommand += data[1, j];
                                        bufCommand += ",";
                                        break;
                                    case "Lng":
                                        bufColumn.Add(data[1, j].ToString());
                                        bufCommand += data[1, j];
                                        bufCommand += ",";
                                        break;
                                    case "Lat":
                                        bufColumn.Add(data[1, j].ToString());
                                        bufCommand += data[1, j];
                                        bufCommand += ",";
                                        break;
                                    default:
                                        bufColumn.Add(data[1, j].ToString());
                                        break;
                                }

                            }

                            char[] chars = bufCommand.ToCharArray();
                            chars[chars.Length - 1] = ')';
                            bufCommand = new string(chars);

                            if (dataStorage.Columns.Count == 0)
                            {
                                dataStorage.Columns.Add("Point", "Point");
                                dataStorage.Columns.Add("Lng", "Lng");
                                dataStorage.Columns.Add("Lat", "Lat");
                            }

                            bufCommand += " values(";
                            for (int i = 2; i <= row; i++)
                            {
                                string command = bufCommand;
                                //PointInfo poin = new PointInfo();
                                object[] bufData = new object[3];
                                for (int j = 1; j <= column; j++)
                                {

                                    switch (bufColumn[j - 1])
                                    {
                                        case "Point":
                                            // poin.Point = data[i, j].ToString();
                                            bufData[0] = data[i, j].ToString();
                                            command += "'" + data[i, j].ToString() + "'";
                                            command += ",";
                                            break;
                                        case "Lng":
                                            // poin.Lng = data[i, j].ToString();
                                            bufData[1] = data[i, j];
                                            command += data[i, j].ToString().Replace(",", ".");
                                            command += ",";
                                            break;
                                        case "Lat":
                                            // poin.Lat = data[i, j].ToString();
                                            bufData[2] = data[i, j];
                                            command += data[i, j].ToString().Replace(",", ".");
                                            command += ",";
                                            break;
                                        default:
                                            break;
                                    }

                                }

                                chars = command.ToCharArray();
                                chars[chars.Length - 1] = ')';
                                command = new string(chars);

                                dataStorage.Rows.Add(bufData);
                                //addMarker(poin);


                                using (DBhelper sqlH = new DBhelper())
                                {
                                    Exception ex, ex1;
                                    if (command != bufCommand)
                                    {
                                        ex = sqlH.WriteCommand(command);
                                    }
                                    else
                                    {
                                        ex = new Exception("В таблице нет полей с обозначением точек!");
                                    }

                                    if (ex != null)
                                    {
                                        //throw ex;
                                    }

                                }

                            }

                            markerFromGrid();

                            string bufCommandAnalyzes = "insert Analyzes (Point, Element, Analyz) values(";

                            if (dataStorageAnalyzes.Columns.Count == 0)
                            {
                                dataStorageAnalyzes.Columns.Add("Point", "Point");
                                dataStorageAnalyzes.Columns.Add("Element", "Element");
                                dataStorageAnalyzes.Columns.Add("Analyz", "Analyz");
                            }

                            for (int i = 2; i <= row; i++)
                            {
                                //string command = bufCommand;
                                string vremCommandAnalyzes = bufCommandAnalyzes;
                                string bufPoin = null;
                                for (int j = 1; j <= column; j++)
                                {

                                    switch (bufColumn[j - 1])
                                    {
                                        case "Point":
                                            vremCommandAnalyzes += "'" + data[i, j].ToString() + "'";
                                            vremCommandAnalyzes += ",";
                                            bufPoin = data[i, j].ToString();
                                            break;
                                        default:
                                            break;
                                    }

                                }


                                using (DBhelper sqlH = new DBhelper())
                                {

                                    for (int j = 1; j <= column; j++)
                                    {
                                        string commandAnalyzes = vremCommandAnalyzes;
                                        if (bufColumn[j - 1] != "Point" && bufColumn[j - 1] != "Lng" && bufColumn[j - 1] != "Lat")
                                        {
                                            commandAnalyzes += "'" + bufColumn[j - 1] + "'" + "," + data[i, j].ToString().Replace(",", ".") + ")";
                                            Exception ex = sqlH.WriteCommand(commandAnalyzes);

                                            object[] bufData = new object[3];
                                            bufData[0] = bufPoin;
                                            bufData[1] = bufColumn[j - 1];
                                            bufData[2] = data[i, j];
                                            dataStorageAnalyzes.Rows.Add(bufData);

                                        }

                                    }
                                }

                                getElements();


                            }


                        }
                        else
                        {
                            throw new Exception("Не удалось считать данные!");
                        }
                    }

                }

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            selectMarker();
        }
    }
}
