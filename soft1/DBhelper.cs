using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace soft1
{
    class DBhelper : IDisposable
    {
        public static string connString = @"Data Source = HOME-PC\SQLEXPRESS; Initial Catalog = diplom; Integrated Security = True;";
        private SqlConnection sqlConnection;
        public DBhelper()
        {
            try
            {
                sqlConnection = new SqlConnection(connString);
                sqlConnection.Open();
            }
            catch (Exception ex)
            {
                
            }
        }

        public Exception WriteCommand(string command)
        {
            try
            {
                SqlCommand sqlCommand = new SqlCommand(command, sqlConnection);
                sqlCommand.ExecuteNonQuery();
                return null;
            }
            catch (Exception ex)
            {
                return ex;
            }

        }

        public DataGridView ReadCommand(string command)
        {
            try
            {
                DataGridView dataStorage = new DataGridView();
                SqlCommand sqlCommand = new SqlCommand(command, sqlConnection);
                SqlDataReader sqlDataReader = sqlCommand.ExecuteReader();

                var columns = new List<string>();
                for (int i = 0; i < sqlDataReader.FieldCount; i++)
                {
                    columns.Add(sqlDataReader.GetName(i));
                    dataStorage.Columns.Add(sqlDataReader.GetName(i), sqlDataReader.GetName(i));
                }

                while (sqlDataReader.Read())
                {
                    string[] bufData = new string[sqlDataReader.FieldCount];
                    int schet = 0;
                    foreach (var fiel in columns)
                    {
                        bufData[schet] = sqlDataReader[fiel].ToString().Replace(" ", "");
                        schet++;
                    }
                    dataStorage.Rows.Add(bufData);


                    //PointInfo poin = new PointInfo();
                    //poin.Point = sqlDataReader["Point"].ToString();
                    //poin.Lng = sqlDataReader["Lng"].ToString();
                    //poin.Lat = sqlDataReader["Lat"].ToString();
                    //addMarker(poin);
                }

                sqlDataReader.Close();

                return dataStorage;

            }
            catch (Exception ex)
            {

            }

            return null;
        }

        public void Dispose()
        {
            try
            {
                sqlConnection.Close();
            }
            catch (Exception ex)
            {

            }
        }
    }
}
