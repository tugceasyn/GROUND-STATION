using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO.Ports;
using System.Drawing.Drawing2D;
using System.Globalization;
using System.Threading;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using OpenTK;
using OpenTK.Graphics.OpenGL;

namespace YER_İSTASYONU_ATIŞ
{
    public partial class Form1 : Form
    {
        DateTime yeni = DateTime.Now;
        double enlem;
        double boylam;
        int zaman = 0;
        int satir = 1;
        int SatirNo = 1;
        string dataVeri;

        float x = 0, y = 0, z = 0;

        public Form1()
        {
            InitializeComponent();
        }

        private void silindir(float step, float topla, float radius, float dikey1, float dikey2)
        {
            float eski = 0.1f;
            GL.Begin(BeginMode.Quads); // Cylinder

             while (step <= 360)
            {

               GL.Color3(Color.FromArgb(0,0,0));           

              float ciz1_x = (float)(radius * Math.Cos(step * Math.PI / 90F));
              float ciz1_y = (float)(radius * Math.Sin(step * Math.PI / 90F));
              GL.Vertex3(ciz1_x, dikey1, ciz1_y);

              float ciz2_x = (float)(radius * Math.Cos((step + 2) * Math.PI / 90F));
              float ciz2_y = (float)(radius * Math.Sin((step + 2) * Math.PI / 90F));
               GL.Vertex3(ciz2_x, dikey1, ciz2_y);

               GL.Vertex3(ciz1_x, dikey2, ciz1_y);
              GL.Vertex3(ciz2_x, dikey2, ciz2_y);
               step += topla;
            }
            

            GL.End();
            GL.Begin(BeginMode.Lines);
            step = eski;
            topla = step;

            while (step <= 180)  // Top Cover
            {
                GL.Color3(Color.FromArgb(0,0,0));
              
                float ciz1_x = (float)(radius * Math.Cos(step * Math.PI / 180F));
                float ciz1_y = (float)(radius * Math.Sin(step * Math.PI / 180F));
                GL.Vertex3(ciz1_x, dikey1, ciz1_y);

                float ciz2_x = (float)(radius * Math.Cos((step + 180) * Math.PI / 180F));
                float ciz2_y = (float)(radius * Math.Sin((step + 180) * Math.PI / 180F));
                GL.Vertex3(ciz2_x, dikey1, ciz2_y);

                GL.Vertex3(ciz1_x, dikey1, ciz1_y);
                GL.Vertex3(ciz2_x, dikey1, ciz2_y);
                step += topla;
            }
            step = eski;
            topla = step;


            while (step <= 180)  // Bottom Cover
            {

                GL.Color3(Color.FromArgb(0,0,0));
               
                float ciz1_x = (float)(radius * Math.Cos(step * Math.PI / 180F));
                float ciz1_y = (float)(radius * Math.Sin(step * Math.PI / 180F));
                GL.Vertex3(ciz1_x, dikey2, ciz1_y);

                float ciz2_x = (float)(radius * Math.Cos((step + 180) * Math.PI / 180F));
                float ciz2_y = (float)(radius * Math.Sin((step + 180) * Math.PI / 180F));
                GL.Vertex3(ciz2_x, dikey2, ciz2_y);

                GL.Vertex3(ciz1_x, dikey2, ciz1_y);
                GL.Vertex3(ciz2_x, dikey2, ciz2_y);
                step += topla;
            }
            GL.End();

        }
        private void burunkonisi(float step, float topla, float radius1, float radius2, float dikey1, float dikey2)
        {
            float eski_step = 0.1f;
            GL.Begin(BeginMode.Lines);
            while (step <= 360)
            {
                GL.Color3(Color.FromArgb(105, 105, 105));

                float ciz1_x = (float)(radius1 * Math.Cos(step * Math.PI / 180F));
                float ciz1_y = (float)(radius1 * Math.Sin(step * Math.PI / 180F));
                GL.Vertex3(ciz1_x, dikey1, ciz1_y);

                float ciz2_x = (float)(radius2 * Math.Cos(step * Math.PI / 180F));
                float ciz2_y = (float)(radius2 * Math.Sin(step * Math.PI / 180F));
                GL.Vertex3(ciz2_x, dikey2, ciz2_y);
                step += topla;
            }
            GL.End();
        }

        private void kanat(float yukseklik, float uzunluk, float kalinlik, float egiklik)
        {
            float yaricap = 10, aci = 90.0f;
            GL.Begin(BeginMode.Quads);

            GL.Color3(Color.FromArgb(105, 105, 105));
            GL.Vertex3(0.0, yukseklik - egiklik, kalinlik);
            GL.Vertex3(0.0, yukseklik, +kalinlik);
            GL.Vertex3(uzunluk, yukseklik, kalinlik);
            GL.Vertex3(0.0, egiklik, kalinlik);

            GL.End();

        }
        private void kanat1(float yukseklik, float uzunluk, float kalinlik, float egiklik)
        {
            float yaricap = 10, aci = 90.0f;
            GL.Begin(BeginMode.Quads);

            GL.Vertex3(0.0, yukseklik + egiklik, -kalinlik);
            GL.Vertex3(-uzunluk, yukseklik, -kalinlik);
            GL.Vertex3(0.0, egiklik, -kalinlik);
            GL.Vertex3(0.0, yukseklik + egiklik, -kalinlik);

            GL.End();

        }
        private void kanat2(float yukseklik, float uzunluk, float kalinlik, float egiklik)
        {
            float yaricap = 10, aci = 90.0f;
            GL.Begin(BeginMode.Quads);

            GL.Vertex3(-kalinlik, yukseklik, -uzunluk);
            GL.Vertex3(-kalinlik, yukseklik, uzunluk);
            GL.Vertex3(kalinlik, -0.0, 0.0);
            GL.Vertex3(kalinlik, 0.0, 0.0);
            GL.End();

        }
        private void kanat3(float yukseklik, float uzunluk, float kalinlik, float egiklik)
        {
            //float yaricap = 10, aci = 90.0f;
            //GL.Begin(BeginMode.Quads);

            // GL.Color3(Color.Gray);
            // GL.Vertex3(0.0, yukseklik - egiklik, kalinlik);
            // GL.Vertex3(0.0, yukseklik, +kalinlik);
            // GL.Vertex3(uzunluk, yukseklik, kalinlik);
            // GL.Vertex3(0.0, egiklik, kalinlik);

            GL.End();

        }
        private void Form1_Load(object sender, EventArgs e)
        {
            Thread.CurrentThread.CurrentCulture = CultureInfo.CreateSpecificCulture("en-US");

            string[] ports = SerialPort.GetPortNames();
            foreach (String port in ports)
            {
                comboBox1.Items.Add(port);
            }

            gmap.MinZoom = -100;
            gmap.MaxZoom = 200;

            gmap.MapProvider = GMap.NET.MapProviders.ArcGIS_Imagery_World_2D_MapProvider.Instance;
            GMap.NET.GMaps.Instance.Mode = GMap.NET.AccessMode.ServerOnly;
            //gmap.Position = new GMap.NET.PointLatLng(enlem,boylam);
            gmap.Position = new GMap.NET.PointLatLng(0, 0);

            DateTime yeni = DateTime.Now;

            GL.ClearColor(Color.FromArgb(25, 25, 112));
        }
        private void button1_Click(object sender, EventArgs e)
        {

            try
            {
                serialPort1.PortName = comboBox1.Text;
                serialPort1.Open();                     // Open serial port
                timer1.Start();                         // Zamanlayıcıyı başlat
                button2.Enabled = true;                  // Active stop button
                button1.Enabled = false;                 // Inactive start buttom

            }
            catch (Exception hata)
            {
                MessageBox.Show("Lütfen Port Seçiniz");
            }
        }
        private void button2_Click(object sender, EventArgs e)
        {
            timer1.Stop();
            serialPort1.Close();
            if (serialPort1.IsOpen == true) ;
            {
                serialPort1.Close();
            }

            button1.Enabled = true;
            button2.Enabled = false;
        }

        private void glControl1_Load(object sender, EventArgs e)
        {
            GL.ClearColor(0.0f, 0.0f, 0.0f, 0.0f);
            GL.Enable(EnableCap.DepthTest);
        }

        private void glControl1_Paint(object sender, PaintEventArgs e)
        {
            float step = 1.0f; //Adım genişliği çözünürlük
            float topla = step; //Tanpon
            float yaricap = 5.0f; 
            float dikey1 = yaricap, dikey2 = -yaricap;
            GL.Clear(ClearBufferMask.ColorBufferBit); // Buffer temizlenmez ise görüntüler üst üste biner o yüzden temizliyoruz.
            GL.Clear(ClearBufferMask.DepthBufferBit);

            Matrix4 perspective = Matrix4.CreatePerspectiveFieldOfView(1.04f, 4 / 3, 1, 800); 
            Matrix4 lookat = Matrix4.LookAt(40, 0, 0, 0, 0, 0, 0, 1, 0); // görünütü uzaklık-yakınlık
            GL.MatrixMode(MatrixMode.Projection);
            GL.LoadIdentity();
            GL.LoadMatrix(ref perspective);
            GL.MatrixMode(MatrixMode.Modelview);
            GL.LoadIdentity();
            GL.LoadMatrix(ref lookat);
            GL.Viewport(0, 0, glControl1.Width, glControl1.Height);
            GL.Enable(EnableCap.DepthTest);
            GL.DepthFunc(DepthFunction.Less);


            //koordinatlar nesneyi hareket ettirmemizi sağlıyor.
            GL.Rotate(x, 0.0, 0.0, -1.0);
            GL.Rotate(y, 0.0, 1.0, 0.0);
            GL.Rotate(z, -1.0, 0.0, 0.0);

            //Çizim Fonksiyonları
            silindir(step, topla, yaricap, 10, -13);
            burunkonisi(0.01f, 0.01f, 5.0f, 0.0f, 10.0f, 18.0f);
            //silindir(0.01f, topla, 0.5f, 9, 9.7f);
            //silindir(0.01f, topla, 0.1f, 5, 5);
            //(Yükseklik,Uzunluk,Genişlik,Açı)
            kanat( -13.0f,  -12.0f, 0.1f, 0.1f);
            kanat1(-13.0f, -12.0f, 0.1f, 0.1f);
            kanat2(-13.0f, -12.0f, 0.1f, 0.1f);
            kanat3(-13.0f, -12.0f, 0.1f, 0.1f);
            

            // AŞAĞIDA X, Y, Z EKSEN CİZGELERİ ÇİZDİRİLİYOR

            GL.Begin(BeginMode.Lines);

            GL.Color3(Color.FromArgb(105, 105, 105));
            GL.Vertex3(-30.0, 0.0, 0.0);
            GL.Vertex3(30.0, 0.0, 0.0);

            GL.Color3(Color.FromArgb(105, 105, 105));
            GL.Vertex3(0.0, 30.0, 0.0);
            GL.Vertex3(0.0, -30.0, 0.0);

            //GL.Color3(Color.FromArgb(0, 0, 0));
            //GL.Vertex3(0.0, 0.0, 30.0);
            //GL.Vertex3(0.0, 0.0, -30.0);

            GL.End();
            //GraphicsContext.CurrentContext.VSync = true;
            glControl1.SwapBuffers();

        }

        //Butonlara 3D görüntü sağlıyor
        private void button1_Paint(object sender, PaintEventArgs e)
        {
            ControlPaint.DrawBorder(e.Graphics, button1.ClientRectangle,
            SystemColors.ControlLightLight, 5, ButtonBorderStyle.Outset,
            SystemColors.ControlLightLight, 5, ButtonBorderStyle.Outset,
            SystemColors.ControlLightLight, 5, ButtonBorderStyle.Outset,
            SystemColors.ControlLightLight, 5, ButtonBorderStyle.Outset);
        }

        private void button2_Paint(object sender, PaintEventArgs e)
        {
            ControlPaint.DrawBorder(e.Graphics, button1.ClientRectangle,
            SystemColors.ControlLightLight, 5, ButtonBorderStyle.Outset,
            SystemColors.ControlLightLight, 5, ButtonBorderStyle.Outset,
            SystemColors.ControlLightLight, 5, ButtonBorderStyle.Outset,
            SystemColors.ControlLightLight, 5, ButtonBorderStyle.Outset);
        }

        private void button3_Paint(object sender, PaintEventArgs e)
        {
            ControlPaint.DrawBorder(e.Graphics, button1.ClientRectangle,
            SystemColors.ControlLightLight, 5, ButtonBorderStyle.Outset,
            SystemColors.ControlLightLight, 5, ButtonBorderStyle.Outset,
            SystemColors.ControlLightLight, 5, ButtonBorderStyle.Outset,
            SystemColors.ControlLightLight, 5, ButtonBorderStyle.Outset);
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            try
            {
                dataVeri = serialPort1.ReadLine();
                string[] parcala = dataVeri.Split('*');
                    textBox8.Text = parcala[0];
                    textBox3.Text = parcala[1];
                    textBox2.Text = parcala[2];
                    textBox1.Text = parcala[3];
                    textBox6.Text = parcala[4];
                    textBox7.Text = parcala[5];
                    textBox4.Text = parcala[6]; 
                    textBox5.Text = parcala[7];
          
                
            }

            catch (Exception hata)
            {

            }

            //Gmap
            try
            {
                enlem = double.Parse(textBox4.Text);
                boylam = double.Parse(textBox5.Text);
                gmap.MapProvider = GMap.NET.MapProviders.ArcGIS_Imagery_World_2D_MapProvider.Instance;
                GMap.NET.GMaps.Instance.Mode = GMap.NET.AccessMode.ServerOnly;
                gmap.Position = new GMap.NET.PointLatLng(enlem, boylam);
            }

            catch (Exception hata)
            {

            }
            try
            {
                    chart1.Series["HIZ"].Points.AddXY(zaman, textBox2.Text);
                    chart2.Series["YÜKSEKLİK"].Points.AddXY(zaman, textBox3.Text);
                    zaman = zaman + 1;

                //ekseni hareketi

                    x = Convert.ToInt32(textBox1.Text);
                    y = Convert.ToInt32(textBox6.Text);
                    z = Convert.ToInt32(textBox7.Text);
                    glControl1.Invalidate();
            }
            catch
            { }
            
            //dataGridWiew
            try
            {
                satir = dataGridView1.Rows.Add();

                dataGridView1.Rows[satir].Cells[0].Value = SatirNo;
                dataGridView1.Rows[satir].Cells[1].Value = textBox2.Text; // Velocity
                dataGridView1.Rows[satir].Cells[2].Value = textBox3.Text; // Altitude
                dataGridView1.Rows[satir].Cells[3].Value = textBox1.Text; // X axis
                dataGridView1.Rows[satir].Cells[4].Value = textBox6.Text; // Y axis
                dataGridView1.Rows[satir].Cells[5].Value = textBox7.Text; // Z axis
                dataGridView1.Rows[satir].Cells[6].Value = textBox4.Text; // Latitude
                dataGridView1.Rows[satir].Cells[7].Value = textBox5.Text; // Longitude
                dataGridView1.Rows[satir].Cells[8].Value = yeni.ToLongTimeString();
                dataGridView1.Rows[satir].Cells[9].Value = yeni.ToLongDateString();
                satir++;
                SatirNo++;
            }
            catch
            {

            }
        }

        //Excel
        private void button3_Click(object sender, EventArgs e)
        {
            Microsoft.Office.Interop.Excel.Application objExcel = new Microsoft.Office.Interop.Excel.Application();
            objExcel.Visible = true;
            Microsoft.Office.Interop.Excel.Workbook objbook = objExcel.Workbooks.Add(System.Reflection.Missing.Value);
            Microsoft.Office.Interop.Excel.Worksheet objSheet = (Microsoft.Office.Interop.Excel.Worksheet)objbook.Worksheets.get_Item(1);

            for (int s = 0; s < dataGridView1.Columns.Count; s++)
            {
                Microsoft.Office.Interop.Excel.Range myrange = (Microsoft.Office.Interop.Excel.Range)objSheet.Cells[1, s + 1];
                myrange.Value2 = dataGridView1.Columns[s].HeaderText;
            }

            for (int s = 0; s < dataGridView1.Columns.Count; s++)
            {
                for (int j = 0; j < dataGridView1.Rows.Count; j++)
                {
                    Microsoft.Office.Interop.Excel.Range myrange = (Microsoft.Office.Interop.Excel.Range)objSheet.Cells[j + 2, s + 1];
                    myrange.Value2 = dataGridView1[s, j].Value;
                }}}}}







