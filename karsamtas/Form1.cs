using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using MySql.Data;
using MySql.Data.MySqlClient;

namespace karsamtas
{
    public partial class Form1 : Form
    {
        public MySqlConnection mysqlbaglan = new MySqlConnection("Server=karsamtas.com;Database=karsamtasProg;Uid=karsamtas;Pwd='852456456';");
        
        public Form1()
        {
            InitializeComponent();
            
        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
            Application.Exit();
        }

        private void button1_Click(object sender, EventArgs e)
        {
          
           
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void button2_Click_1(object sender, EventArgs e)
        {
           
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {
            this.Close();
            Application.Exit();
        }

        private void pictureBox2_Click(object sender, EventArgs e)
        {

 
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            try
            {
              
            

            mysqlbaglan.Open();
            MySqlCommand komut = new MySqlCommand("select * from kullanici where kullaniciAdi='" + textBox1.Text.Trim() + "' and sifre='" + textBox2.Text.Trim() + "'", mysqlbaglan);
            MySqlDataReader dr = komut.ExecuteReader();
            if (dr.Read())
            {

             
                Form arkaplan = new arkaplan();
                arkaplan.Show();
                this.Hide();

                Form anasayfa = new anasayfa();
                anasayfa.Show();


            }
            else
            {
                MessageBox.Show("Kullanıcı Adı Veya Şifre Hatalı","Giriş Hatası",MessageBoxButtons.OK,MessageBoxIcon.Stop);
            }
            mysqlbaglan.Close();
            }
            catch (Exception)
            {

                MessageBox.Show("Giriş Başarısız! Lütfen internet bağlantınızı kontrol edip programı yeniden başlatınız.", "Bağlantı Hatası", MessageBoxButtons.OK,MessageBoxIcon.Error);
        
            }
        }

        private void button2_Click_2(object sender, EventArgs e)
        {

        }
    }
}
