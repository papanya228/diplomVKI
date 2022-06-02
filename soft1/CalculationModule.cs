using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using GMap.NET.MapProviders;
using GMap.NET.WindowsForms;
using GMap.NET.WindowsForms.Markers;

namespace soft1
{
    class CalculationModule
    {
        public static GMap.NET.PointLatLng firstPos = new GMap.NET.PointLatLng(56.141, 92.6889);
        public static GMap.NET.PointLatLng lastPos = new GMap.NET.PointLatLng(55.9677, 93.1586);
        public static int calcRangeMetr = 100;

        //На входе 4 параметра: Широта, долгота точки A(От которой считаем расстояние); Широта, Долгота точки B(До которой считаем расстояние)
        //На выходе растояние от точки A до точки B в формате double в !!!!!!МЕТРАХ!!!!!!
        //Не перепутай lat(Широту) и lng(Долготу), lat(Широта) скорее всего меньше чем lng(Долгота)
        public static double calculateTheDistance(double latA, double lngA, double latB, double lngB)
        {
            //Радиус земли
            double earthRadius = 6372795;

            // перевести координаты в радианы
            double lat1 = latA * Math.PI / 180;
            double lng1 = lngA * Math.PI / 180;
            double lat2 = latB * Math.PI / 180;
            double lng2 = lngB * Math.PI / 180;

            //double dy = lat2 - lat1;
            //double dx = Math.Cos(Math.PI / 180 * lat1) * (lng2 - lng1);
            //double angel = Math.Atan2(dy, dx);
            //double gAngel = angel * 180 / Math.PI;

            // косинусы и синусы широт и разницы долгот
            double cl1 = Math.Cos(lat1);
            double cl2 = Math.Cos(lat2);
            double sl1 = Math.Sin(lat1);
            double sl2 = Math.Sin(lat2);
            double delta = lng2 - lng1;
            double cdelta = Math.Cos(delta);
            double sdelta = Math.Sin(delta);

            //double angel = Math.Atan2(cl1 * sdelta, cl1 * sl2 - sl1 * cl2 * cdelta);
            //double gAngel = angel * 180 / Math.PI;

            // вычисления длины большого круга
            double y = Math.Sqrt(Math.Pow(cl1 * sdelta, 2) + Math.Pow(cl1 * sl2- sl1 * cl2 * cdelta, 2));
            double x = sl1 * sl2 + cl1 * cl2 * cdelta;

            double ad = Math.Atan2(y, x);
            double dist = ad * earthRadius;

            return dist;
        }

        public static List<List<double>> Calc(List<GMapMarker> selectedPoint, DataGridView dataStorageAnalyzes, string element)
        {
            if(selectedPoint.Count != 0 && dataStorageAnalyzes.Rows.Count != 0)
            {
                double dist = calculateTheDistance(firstPos.Lat, firstPos.Lng, lastPos.Lat, firstPos.Lng);
                double incrementX = Math.Abs((firstPos.Lat - lastPos.Lat) / (dist / calcRangeMetr));
                dist = calculateTheDistance(firstPos.Lat, firstPos.Lng, firstPos.Lat, lastPos.Lng);
                double incrementY = Math.Abs((firstPos.Lng - lastPos.Lng) / (dist / calcRangeMetr));
                if(firstPos.Lat > lastPos.Lat) incrementX = -incrementX;
                if(firstPos.Lng > lastPos.Lng) incrementY = -incrementY;


                System.Data.DataTable table = new System.Data.DataTable("ParentTable");
                System.Data.DataColumn column;
                System.Data.DataRow row;

                column = new System.Data.DataColumn();
                column.DataType = System.Type.GetType("System.Double");
                column.ColumnName = "lat";
                column.AutoIncrement = false;
                column.ReadOnly = true;
                column.Unique = false;
                table.Columns.Add(column);

                column = new System.Data.DataColumn();
                column.DataType = System.Type.GetType("System.Double");
                column.ColumnName = "lng";
                column.AutoIncrement = false;
                column.ReadOnly = true;
                column.Unique = false;
                table.Columns.Add(column);

                column = new System.Data.DataColumn();
                column.DataType = System.Type.GetType("System.Double");
                column.ColumnName = "value";
                column.AutoIncrement = false;
                column.ReadOnly = true;
                column.Unique = false;
                table.Columns.Add(column);

                for (int i = 0; i < dataStorageAnalyzes.Rows.Count - 1; i++)
                {
                    GMapMarker findedPoint = selectedPoint.Find(item => item.Tag.ToString() == dataStorageAnalyzes["Point", i].Value.ToString());
                    if (findedPoint != null && dataStorageAnalyzes["Element", i].Value.ToString() == element)
                    {
                        row = table.NewRow();
                        row["lat"] = findedPoint.Position.Lat;
                        row["lng"] = findedPoint.Position.Lng;
                        row["value"] = Convert.ToDouble(dataStorageAnalyzes["Analyz", i].Value);
                        table.Rows.Add(row);
                    }

                }

                List<List<double>> res = new List<List<double>>();

                int numberX = 0;
                for (double x = firstPos.Lat; x > lastPos.Lat; x = x + incrementX)
                {
                    res.Add(new List<double>());
                    int numberY = 0;
                    for (double y = firstPos.Lng; y < lastPos.Lng; y = y + incrementY)
                    {
                        double qq1 = 0, qq2 = 0;
                        for (int i = 0; i < table.Rows.Count; i++)
                        {
                            //int rBuf = calculateTheDistance(x, y, Convert.ToDouble(bufferData["lat", i].Value), Convert.ToDouble(bufferData["lng", i].Value));
                            //double r = calculateTheDistance(x, y, Convert.ToDouble(table.Rows[i]["lat"]), Convert.ToDouble(table.Rows[i]["lng"])) / 1000;
                            double r = calculateTheDistance(Convert.ToDouble(table.Rows[i]["lat"]), Convert.ToDouble(table.Rows[i]["lng"]), x, y) / 1000;
                            qq1 = qq1 + (Convert.ToDouble(table.Rows[i]["value"]) / Math.Pow(r, 2));
                            qq2 = qq2 + (1 / Math.Pow(r, 2));
                        }
                        res[numberX].Add(qq1/qq2);
                        numberY++;
                    }
                    numberX++;
                }
                return res;
            }

            return null;
        }
    }
}
