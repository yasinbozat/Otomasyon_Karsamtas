using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Office = Microsoft.Office.Interop.Excel;

namespace karsamtas
{
    public partial class anasayfa : Form
    {
        public anasayfa()
        {
            InitializeComponent();
        }

       

        private void kayitGor_Click(object sender, EventArgs e)
        {
            
            try
            {

            
            Form kayitGoruntule = new kayitGoruntule();
            kayitGoruntule.Show();
            this.Hide();

           }
            catch (Exception)
            {
                MessageBox.Show("Kayıt Görüntüleme Başarısız! Lütfen internet bağlantınızı kontrol edip programı yeniden başlatınız." );

            }
           
        }

        private void kayitEkle_Click(object sender, EventArgs e)
        {
           
        }

        private void anasayfa_Load(object sender, EventArgs e)
        {

        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {
         
        }

        private void button1_Click(object sender, EventArgs e)
        {
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            Form tasimaIrsaliyesi = new tasimaIrsaliyesi();
            tasimaIrsaliyesi.Show();
            this.Hide();
            
        }

        private void button2_Click_1(object sender, EventArgs e)
        {
           
        }

        private void pictureBox1_Click_1(object sender, EventArgs e)
        {
            this.Close();
            Application.Exit();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Form fatura = new fatura();
            fatura.Show();
            this.Hide();

        }

        private void button3_Click(object sender, EventArgs e)
        {
            Form topluTasima = new topluTasima();
            topluTasima.Show();
            this.Hide();
        }


       
    }
}
