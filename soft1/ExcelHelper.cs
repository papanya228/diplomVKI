using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using ExcelDataReader;
using System.Data;
using System.Windows.Forms;
using System.Drawing.Drawing2D;
using System.Drawing;
using System.Drawing.Imaging;

namespace soft1
{
    class ExcelHelper : IDisposable
    {
        private Excel.Application _excel;
        private Workbook _Workbook;
        private string _filePath;
        private string[] columnNames = {"A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"};

        public ExcelHelper()
        {
            _excel = new Excel.Application(); 
        }


        public static Exception SaveMap(Image image)
        {
            try
            {
                using (SaveFileDialog dialog = new SaveFileDialog())
                {
                    dialog.Filter = "PNG (*.png)|*.png";
                    dialog.FileName = "GMap.NET image";
                    //Image image = gMapControl1.ToImage();
                    if (image != null)
                    {
                        int difX = 1920 - image.Width;
                        int difY = 1080 - image.Height;
                        if (difX < difY)
                        {
                            int proc = image.Height / (image.Width / 100);
                            image = ResizeImage(image, 1920, Convert.ToInt32(19.2 * proc));
                        }
                        else
                        {
                            int proc = image.Width / (image.Height / 100);
                            image = ResizeImage(image, Convert.ToInt32(10.8 * proc), 1080);
                        }

                        using (image)
                        {
                            if (dialog.ShowDialog() == DialogResult.OK)
                            {
                                string fileName = dialog.FileName;
                                if (!fileName.EndsWith(".png", StringComparison.OrdinalIgnoreCase))
                                {
                                    fileName += ".png";
                                }
                                image.Save(fileName);
                                // MessageBox.Show("Image saved: " + dialog.FileName, "GMap.NET", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                                return null;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                //MessageBox.Show("Image save failed: " + exception.Message, "GMap.NET", MessageBoxButtons.OK, MessageBoxIcon.Hand);
                return ex;
            }

            return null;
        }

        public static Bitmap ResizeImage(Image image, int width, int height)
        {
            var destRect = new System.Drawing.Rectangle(0, 0, width, height);
            var destImage = new Bitmap(width, height);

            destImage.SetResolution(image.HorizontalResolution, image.VerticalResolution);

            using (var graphics = Graphics.FromImage(destImage))
            {
                graphics.CompositingMode = CompositingMode.SourceCopy;
                graphics.CompositingQuality = CompositingQuality.HighQuality;
                graphics.InterpolationMode = InterpolationMode.HighQualityBicubic;
                graphics.SmoothingMode = SmoothingMode.HighQuality;
                graphics.PixelOffsetMode = PixelOffsetMode.HighQuality;

                using (var wrapMode = new ImageAttributes())
                {
                    wrapMode.SetWrapMode(WrapMode.TileFlipXY);
                    graphics.DrawImage(image, destRect, 0, 0, image.Width, image.Height, GraphicsUnit.Pixel, wrapMode);
                }
            }

            return destImage;
        }

        public void Dispose()
        {
            try
            {
                _Workbook.Close();
                _excel.Quit();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        public bool Open()
        {
            try
            {

                OpenFileDialog fileDialog = new OpenFileDialog();
                DialogResult res = fileDialog.ShowDialog();
                string filePath;

                if (res == DialogResult.OK)
                {
                    filePath = fileDialog.FileName;

                    if (File.Exists(filePath))
                    {
                        _Workbook = _excel.Workbooks.Open(filePath);
                    }
                    else
                    {
                        _Workbook = _excel.Workbooks.Add(filePath);
                        _filePath = filePath;
                    }

                    return true;

                }
                else
                {
                    throw new Exception("Файл не найден");
                }

            }
            catch(Exception ex) { Console.WriteLine(ex.Message); }

            return false;
        }

        public bool Set(string column, int row, object data)
        {

            try
            {
                ((Excel.Worksheet)_excel.ActiveSheet).Cells[row, column] = data;
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

            return false;
        }

        public object[,] Read()
        {
            try
            {
                //Excel.Worksheet ObjWorkSheet = (Excel.Worksheet)_Workbook.Sheets[1];
                //var lastCell = ((Excel.Worksheet)_excel.ActiveSheet).Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);
                //string[,] list = new string[lastCell.Column, lastCell.Row];
                //for (int i = 0; i < (int)lastCell.Column; i++)
                //    for (int j = 0; j < (int)lastCell.Row; j++)
                //        list[j, i] = ((Excel.Worksheet)_excel.ActiveSheet).Cells[j + 1, i + 1].Text.ToString();

                //return list;

                //int iLastRow = ((Excel.Worksheet)_excel.ActiveSheet).Cells[((Excel.Worksheet)_excel.ActiveSheet).Rows.Count, "A"].End[Excel.XlDirection.xlUp].Row;  //последняя заполненная строка в столбце А
                var lastCell = ((Excel.Worksheet)_excel.ActiveSheet).Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);
                int iLastRow = lastCell.Row;
                string iLastColumn;
                if(lastCell.Column / 26 == 0)
                {
                    iLastColumn = columnNames[lastCell.Column-1];
                }
                else if(lastCell.Column / 26 == 1)
                {
                    iLastColumn = "A" + columnNames[lastCell.Column - 27];
                }
                else
                {
                    throw new Exception("Слишком много колонок в таблице");
                }

               
                var arrData = (object[,])((Excel.Worksheet)_excel.ActiveSheet).Range["A1:" + iLastColumn + iLastRow].Value;

                return arrData;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

            return null;
        }

        public void Save()
        {
            try
            {
                if (!string.IsNullOrEmpty(_filePath))
                {
                    _Workbook.SaveAs(_filePath);
                }
                else
                {
                    _Workbook.Save();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
           
        }



    }
}
