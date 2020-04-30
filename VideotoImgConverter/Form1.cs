using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using System.Collections;
using System.Reflection;

namespace VideotoImgConverter
{
    public partial class Form1 : Form
    {

        public Form1()
        {
            InitializeComponent();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            axWindowsMediaPlayer1.URL = textBox1.Text;
            axWindowsMediaPlayer1.Ctlcontrols.play();

        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();

            if(openFileDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                this.textBox1.Text = openFileDialog.FileName;
            }

        }

        private void button3_Click(object sender, EventArgs e)
        {
            axWindowsMediaPlayer1.Ctlcontrols.stop();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            axWindowsMediaPlayer1.Ctlcontrols.pause();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            
            VideotoImage();
            MetaData();
           // DeleteImages();

            DialogResult result = MessageBox.Show("Photos are created.", "Confirmation", MessageBoxButtons.OK);
            if (result == DialogResult.OK)
            {
                Process.Start("explorer.exe", this.textBox5.Text);
            }
        }
        private void button6_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();

            if (openFileDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                this.textBox4.Text = openFileDialog.FileName;
            }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog fbd = new FolderBrowserDialog();
            fbd.Description = "Custom Description";

            if (fbd.ShowDialog() == DialogResult.OK)
            {
                string sSelectedPath = fbd.SelectedPath;
                this.textBox5.Text = sSelectedPath;
            }
        }


        private void VideotoImage()
        {
            string filePath = this.textBox1.Text;

            string videoName = Path.GetFileName(filePath);

            videoName = videoName.Remove(videoName.Length - 4);

            var path = Path.GetDirectoryName(Assembly.GetExecutingAssembly().GetName().CodeBase);

            path = path.Substring(6);

            Console.WriteLine("FilePath = " + filePath);
            Console.WriteLine("Path of Bin = " + path);

            string videoPath = path + "\\abc.MP4";

            Console.WriteLine("videoPath = " + videoPath);

            Console.WriteLine("Video Name = " + videoName);

            File.Copy(filePath, videoPath, true);

            //File.Copy(@filePath, @"D:\Bilal Salim Folder\November 25 (Windows Form)\VideotoImgConverter\VideotoImgConverter\bin\Debug\abc.MP4", true);

            string pathVideo = "abc.MP4";

            //string root = @"D:\Temp";
            //string subdir = @"D:\Temp\Images";

            string root = this.textBox5.Text;
            string subdir = this.textBox5.Text + "\\Images";


            if (!Directory.Exists(root))
            {
                Directory.CreateDirectory(root);
            }

            if (!Directory.Exists(subdir))
            {
                Directory.CreateDirectory(subdir);
            }

            string pathImage = subdir + "\\" + videoName + "_FR";

            Console.WriteLine("pathImage = " + pathImage);

            //string pathImage = subdir + "\\GOPR6277_FR";


            //Console.WriteLine("pathImage = " + pathImage);

            //string pathImage = @"\Temp\Images\GOPR6277_FR";

            //string pathImage = @"\VideoPhotos\Images";

            double a = textBox2.Text != "" ? Convert.ToInt32(textBox2.Text) : 0;
            double b = textBox3.Text != "" ? Convert.ToInt32(textBox3.Text) : 0;

            string fps = " fps=" + a / b + " ";

            Console.WriteLine(fps);

            Process p;
            ProcessStartInfo info = new ProcessStartInfo();
            info.FileName = "ffmpeg.exe";
            info.WindowStyle = ProcessWindowStyle.Hidden;
            info.Arguments = " -i " + pathVideo + " -vf " + fps + pathImage + "%d_.jpg" + " -hide_banner ";
            p = Process.Start(info);
            while (!p.HasExited) { Thread.Sleep(10); }

        }

        private void MetaData()
        {
            ArrayList arrayListLat = new ArrayList();
            ArrayList arrayListLon = new ArrayList();

            string filePath = this.textBox1.Text;

            string videoName = Path.GetFileName(filePath);

            videoName = videoName.Remove(videoName.Length - 4);


            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(this.textBox4.Text);
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;

            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;

            //Getting Columns from UI
            
            int columnLat;
            int columnLon;
            if (!String.IsNullOrWhiteSpace(txtbxLatColumn.Text) && !String.IsNullOrWhiteSpace(txtbxLonColumn.Text))
            {
                if (int.TryParse(txtbxLatColumn.Text, out columnLat) &&
                    int.TryParse(txtbxLonColumn.Text, out columnLon))
                {
                    
                }
                else
                {
                    MessageBox.Show("Invalid Column Number");
                    return;
                }
            }
            else
            {
                MessageBox.Show("Invalid Column Number");
                return;
            }
            //Adding plus 1 to remove 0 index
            columnLat++;
            columnLon++;
            
            
            if (textBox3.Text == "1")
            {
                for (int i = 2; i <= rowCount; i = i + 20)
                {
                    for (int j = columnLat; j <= colCount; j++)
                    {

                        //write the value to the console
                        if (j == columnLat)
                        {
                            if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                            {
                                arrayListLat.Add(xlRange.Cells[i, j].Value2.ToString());
                            }
                        }

                        else if (j == columnLon)
                        {
                            if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                            {
                                arrayListLon.Add(xlRange.Cells[i, j].Value2.ToString());
                            }
                        }
                    }
                }
            }

            if (textBox3.Text == "2")
            {
                for (int i = 2; i <= rowCount; i = i + 40)
                {
                    for (int j = columnLat; j <= colCount; j++)
                    {

                        //write the value to the console
                        if (j == columnLat)
                        {
                            if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                            {
                                arrayListLat.Add(xlRange.Cells[i, j].Value2.ToString());
                            }

                        }

                        else if (j == columnLon)
                        {
                            if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                            {
                                arrayListLon.Add(xlRange.Cells[i, j].Value2.ToString());
                            }

                        }

                    }

                }

            }

            if (textBox3.Text == "3")
            {
                for (int i = 2; i <= rowCount; i = i + 60)
                {
                    for (int j = columnLat; j <= colCount; j++)
                    {

                        //write the value to the console
                        if (j == columnLat)
                        {
                            if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                            {
                                arrayListLat.Add(xlRange.Cells[i, j].Value2.ToString());
                            }

                        }

                        else if (j == columnLon)
                        {
                            if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                            {
                                arrayListLon.Add(xlRange.Cells[i, j].Value2.ToString());
                            }

                        }

                    }

                }
            }

            if (textBox3.Text == "4")
            {
                for (int i = 2; i <= rowCount; i = i + 80)
                {
                    for (int j = columnLat; j <= colCount; j++)
                    {

                        //write the value to the console
                        if (j == columnLat)
                        {
                            if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                            {
                                arrayListLat.Add(xlRange.Cells[i, j].Value2.ToString());
                            }

                        }

                        else if (j == columnLon)
                        {
                            if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                            {
                                arrayListLon.Add(xlRange.Cells[i, j].Value2.ToString());
                            }

                        }

                    }

                }
            }

            if (textBox3.Text == "5")
            {
                for (int i = 2; i <= rowCount; i = i + 100)
                {
                    for (int j = columnLat; j <= colCount; j++)
                    {

                        //write the value to the console
                        if (j == columnLat)
                        {
                            if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                            {
                                arrayListLat.Add(xlRange.Cells[i, j].Value2.ToString());
                            }

                        }

                        else if (j == columnLon)
                        {
                            if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                            {
                                arrayListLon.Add(xlRange.Cells[i, j].Value2.ToString());
                            }

                        }

                    }

                }
            }

            if (textBox3.Text == "6")
            {
                for (int i = 2; i <= rowCount; i = i + 120)
                {
                    for (int j = columnLat; j <= colCount; j++)
                    {

                        //write the value to the console
                        if (j == columnLat)
                        {
                            if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                            {
                                arrayListLat.Add(xlRange.Cells[i, j].Value2.ToString());
                            }

                        }

                        else if (j == columnLon)
                        {
                            if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                            {
                                arrayListLon.Add(xlRange.Cells[i, j].Value2.ToString());
                            }

                        }

                    }

                }
            }

            if (textBox3.Text == "7")
            {
                for (int i = 2; i <= rowCount; i = i + 140)
                {
                    for (int j = columnLat; j <= colCount; j++)
                    {

                        //write the value to the console
                        if (j == columnLat)
                        {
                            if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                            {
                                arrayListLat.Add(xlRange.Cells[i, j].Value2.ToString());
                            }

                        }

                        else if (j == columnLon)
                        {
                            if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                            {
                                arrayListLon.Add(xlRange.Cells[i, j].Value2.ToString());
                            }

                        }

                    }

                }
            }

            if (textBox3.Text == "8")
            {
                for (int i = 2; i <= rowCount; i = i + 160)
                {
                    for (int j = columnLat; j <= colCount; j++)
                    {

                        //write the value to the console
                        if (j == columnLat)
                        {
                            if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                            {
                                arrayListLat.Add(xlRange.Cells[i, j].Value2.ToString());
                            }

                        }

                        else if (j == columnLon)
                        {
                            if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                            {
                                arrayListLon.Add(xlRange.Cells[i, j].Value2.ToString());
                            }

                        }

                    }

                }
            }

            if (textBox3.Text == "9")
            {
                for (int i = 2; i <= rowCount; i = i + 180)
                {
                    for (int j = 35; j <= colCount; j++)
                    {

                        //write the value to the console
                        if (j == 35)
                        {
                            if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                            {
                                arrayListLat.Add(xlRange.Cells[i, j].Value2.ToString());
                            }

                        }

                        else if (j == 36)
                        {
                            if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                            {
                                arrayListLon.Add(xlRange.Cells[i, j].Value2.ToString());
                            }

                        }

                    }

                }
            }

            if (textBox3.Text == "10")
            {
                for (int i = 2; i <= rowCount; i = i + 200)
                {
                    for (int j = 35; j <= colCount; j++)
                    {

                        //write the value to the console
                        if (j == 35)
                        {
                            if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                            {
                                arrayListLat.Add(xlRange.Cells[i, j].Value2.ToString());
                            }

                        }

                        else if (j == 36)
                        {
                            if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                            {
                                arrayListLon.Add(xlRange.Cells[i, j].Value2.ToString());
                            }

                        }

                    }

                }
            }


                //for (int i = 2; i <= rowCount; i = i + 20)
                //{
                //    for (int j = 35; j <= colCount; j++)
                //    {

                //        //write the value to the console
                //        if (j == 35)
                //        {
                //            if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                //            {
                //                arrayListLat.Add(xlRange.Cells[i, j].Value2.ToString());
                //            }

                //        }

                //        else if (j == 36)
                //        {
                //            if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                //            {
                //                arrayListLon.Add(xlRange.Cells[i, j].Value2.ToString());
                //            }

                //        }

                //    }

                //}

            foreach (var item in arrayListLat)
            {
                Console.WriteLine("Latitude: " + item);
            }

            foreach (var item in arrayListLon)
            {
                Console.WriteLine("Longtitude: " + item);
            }

            int length = arrayListLat.Count;

            double[] valueLat = new double[length];
            double[] valueLon = new double[length];

            Console.WriteLine("Count Value = " + length);

            for (int i = 0; i <= length; i++)
            {

                try
                {
                    string strLat = arrayListLat[i].ToString();
                    string strLon = arrayListLon[i].ToString();



                    Geotag(new Bitmap(this.textBox5.Text + "\\Images\\" + videoName + "_FR" + i + "_.JPG"), double.Parse(strLat), double.Parse(strLon))
                    .Save(this.textBox5.Text + "\\Images\\" + videoName + "_FRNew" + i + ".JPG", ImageFormat.Jpeg);


                    //Geotag(new Bitmap(this.textBox5.Text + "\\Images\\GOPR6277_FR" + i + ".JPG"), double.Parse(strLat), double.Parse(strLon))
                    //.Save(this.textBox5.Text + "\\Images\\GOPR6277_FRNew" + i + ".JPG", ImageFormat.Jpeg);


                    //string rootFolder = @"D:\Temp\Images\";

                    //string author = "Images" + i + ".JPG";

                    //File.Delete(Path.Combine(rootFolder, author));

                    GC.Collect();
                    GC.WaitForPendingFinalizers();

                    //release com objects to fully kill excel process from running in the background
                    Marshal.ReleaseComObject(xlRange);
                    Marshal.ReleaseComObject(xlWorksheet);

                    //close and release
                    xlWorkbook.Close();
                    Marshal.ReleaseComObject(xlWorkbook);

                    //quit and release
                    xlApp.Quit();
                    Marshal.ReleaseComObject(xlApp);
                }

                catch (Exception)
                {

                }

            }

        }

        static Image Geotag(Image original, double lat, double lng)
        {
            // These constants come from the CIPA DC-008 standard for EXIF 2.3
            const short ExifTypeByte = 1;
            const short ExifTypeAscii = 2;
            const short ExifTypeRational = 5;

            const int ExifTagGPSVersionID = 0x0000;
            const int ExifTagGPSLatitudeRef = 0x0001;
            const int ExifTagGPSLatitude = 0x0002;
            const int ExifTagGPSLongitudeRef = 0x0003;
            const int ExifTagGPSLongitude = 0x0004;

            char latHemisphere = 'N';
            if (lat < 0)
            {
                latHemisphere = 'S';
                lat = -lat;
            }
            char lngHemisphere = 'E';
            if (lng < 0)
            {
                lngHemisphere = 'W';
                lng = -lng;
            }

            MemoryStream ms = new MemoryStream();
            original.Save(ms, ImageFormat.Jpeg);
            ms.Seek(0, SeekOrigin.Begin);

            Image img = Image.FromStream(ms);
            AddProperty(img, ExifTagGPSVersionID, ExifTypeByte, new byte[] { 2, 3, 0, 0 });
            AddProperty(img, ExifTagGPSLatitudeRef, ExifTypeAscii, new byte[] { (byte)latHemisphere, 0 });
            AddProperty(img, ExifTagGPSLatitude, ExifTypeRational, ConvertToRationalTriplet(lat));
            AddProperty(img, ExifTagGPSLongitudeRef, ExifTypeAscii, new byte[] { (byte)lngHemisphere, 0 });
            AddProperty(img, ExifTagGPSLongitude, ExifTypeRational, ConvertToRationalTriplet(lng));

            return img;
        }

        static byte[] ConvertToRationalTriplet(double value)
        {
            int degrees = (int)Math.Floor(value);
            value = (value - degrees) * 60;
            int minutes = (int)Math.Floor(value);
            value = (value - minutes) * 60 * 100;
            int seconds = (int)Math.Round(value);
            byte[] bytes = new byte[3 * 2 * 4]; // Degrees, minutes, and seconds, each with a numerator and a denominator, each composed of 4 bytes
            int i = 0;
            Array.Copy(BitConverter.GetBytes(degrees), 0, bytes, i, 4); i += 4;
            Array.Copy(BitConverter.GetBytes(1), 0, bytes, i, 4); i += 4;
            Array.Copy(BitConverter.GetBytes(minutes), 0, bytes, i, 4); i += 4;
            Array.Copy(BitConverter.GetBytes(1), 0, bytes, i, 4); i += 4;
            Array.Copy(BitConverter.GetBytes(seconds), 0, bytes, i, 4); i += 4;
            Array.Copy(BitConverter.GetBytes(100), 0, bytes, i, 4);
            return bytes;
        }

        static void AddProperty(Image img, int id, short type, byte[] value)
        {
            PropertyItem pi = img.PropertyItems[0];
            pi.Id = id;
            pi.Type = type;
            pi.Len = value.Length;
            pi.Value = value;
            img.SetPropertyItem(pi);
        }

        private void DeleteImages()
        {
            string filePath = this.textBox1.Text;

            string videoName = Path.GetFileName(filePath);

            videoName = videoName.Remove(videoName.Length - 4);

            //string deleteFiles = this.textBox5.Text + "\\Images\\" + videoName + "_FR" + ".JPG";

            string deleteFiles = this.textBox5.Text + "\\Images";

            Console.WriteLine("Delete File Names: "+ deleteFiles);

            string[] directoryFiles = Directory.GetFiles(deleteFiles);

            foreach (string directoryFile in directoryFiles)
            {
                try
                {
                    if (directoryFile.EndsWith("_.jpg"))
                    {
                        File.Delete(directoryFile);
                    }
                }

                catch(Exception ex)
                {
                    Console.WriteLine("Exception:" +ex);
                }

            }

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

    }
}
