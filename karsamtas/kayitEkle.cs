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

    public partial class kayitEkle : Form
    {
        public MySqlConnection mysqlbaglan = new MySqlConnection("Server=karsamtas.com;Database=karsamtasProg;Uid=karsamtas;Pwd='852456456';");
        
        public kayitEkle()
        {
            InitializeComponent();
        }

        private void kayitEkle_Load(object sender, EventArgs e)
        {
            
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            Form anasayfa = new anasayfa();
            anasayfa.Show();
            this.Hide();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            {
                try
                {
                    mysqlbaglan.Open();
                    string komut = "insert into kayitlar(urun,cikisTarihi,miktar,aracPlaka,aracModel,musteriTC,musteriAd,musteriSoyad,ucret,odemeTuru) values('" + textBox1.Text + "', '" + dateTimePicker1.Text + "', '" + textBox2.Text + comboBox1.Text + "','" + textBox3.Text + "','" + textBox4.Text + "','" + textBox5.Text + "','" + textBox6.Text + "','" + textBox7.Text + "','" + textBox8.Text + comboBox3.Text + "','" + comboBox2.Text + "')";
                    MySqlCommand kmt = new MySqlCommand(komut, mysqlbaglan);
                    kmt.ExecuteNonQuery();
                    MessageBox.Show("Kayıt Başarıyla Gerçekleştirildi.");

                    Form anasayfa = new anasayfa();
                    anasayfa.Show();
                    this.Hide();
                  
                }

                catch (Exception  )
                {
                    MessageBox.Show("Beklenmeyen Bir Hata Ouştu Lütfen İnternet Bağlantınızı Kontrol Ediniz Ve Yeniden Deneyiniz. ", "Bağlantı Hatası" ,MessageBoxButtons.OK, MessageBoxIcon.Stop);
                }
            }
        }

        private void textBox2_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (char.IsLetter(e.KeyChar))
            {
                e.Handled = true;
            }
            else
            {
                e.Handled = false;
            }
        }

        private void textBox8_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (char.IsLetter(e.KeyChar))
            {
                e.Handled = true;
            }
            else
            {
                e.Handled = false;
            }
        }

        private void textBox6_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (char.IsNumber(e.KeyChar))
            {
                e.Handled = true;
            }
            else
            {
                e.Handled = false;
            }
        }

        private void textBox7_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (char.IsNumber(e.KeyChar))
            {
                e.Handled = true;
            }
            else
            {
                e.Handled = false;
            }
        }

        private void textBox5_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (char.IsLetter(e.KeyChar))
            {
                e.Handled = true;
            }
            else
            {
                e.Handled = false;
            }
        }
    }
}
