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

    public partial class kayitGoruntule : Form
    {
        public MySqlConnection mysqlbaglan = new MySqlConnection("Server=karsamtas.com;Database=karsamtasProg;Uid=karsamtas;Pwd='852456456';");

        DialogResult sil = new DialogResult();

        public kayitGoruntule()
        {
            InitializeComponent();
        }

        public void baslik()
        {


            dataGridView1.Columns[0].HeaderText = "Seri No";
            dataGridView1.Columns[1].HeaderText = "Alıcı Ünvanı";
            dataGridView1.Columns[2].HeaderText = "Alıcı Adresi";
            dataGridView1.Columns[3].HeaderText = "Alıcı V.D No";
            dataGridView1.Columns[4].HeaderText = "Alıcı Telefon No";
            dataGridView1.Columns[5].HeaderText = "Gönderen Ünvanı";
            dataGridView1.Columns[6].HeaderText = "Gönderen Adresi";
            dataGridView1.Columns[7].HeaderText = "Gönderen V.D. No";
            dataGridView1.Columns[8].HeaderText = "Gönderen Telefon No";
            dataGridView1.Columns[9].HeaderText = "Adet 1";
            dataGridView1.Columns[10].HeaderText = "Kap 1";
            dataGridView1.Columns[11].HeaderText = "Cinsi 1";
            dataGridView1.Columns[12].HeaderText = "KG 1";
            dataGridView1.Columns[13].HeaderText = "Adet 2";
            dataGridView1.Columns[14].HeaderText = "Kap 2";
            dataGridView1.Columns[15].HeaderText = "Cinsi 2";
            dataGridView1.Columns[16].HeaderText = "KG  2";
            dataGridView1.Columns[17].HeaderText = "Adet 3";
            dataGridView1.Columns[18].HeaderText = "Kap 3";
            dataGridView1.Columns[19].HeaderText = "Cinsi 3";
            dataGridView1.Columns[20].HeaderText = "KG 3";
            dataGridView1.Columns[21].HeaderText = "Adet 4";
            dataGridView1.Columns[22].HeaderText = "Kap 4";
            dataGridView1.Columns[23].HeaderText = "Cinsi 4";
            dataGridView1.Columns[24].HeaderText = "KG 4";
            dataGridView1.Columns[25].HeaderText = "Ücret";
            dataGridView1.Columns[26].HeaderText = "Havale";
            dataGridView1.Columns[27].HeaderText = "Peşin";
            dataGridView1.Columns[28].HeaderText = "Tarih";
            dataGridView1.Columns[29].HeaderText = "Yer";
        }

        public void baslik2()
        {


            dataGridView2.Columns[0].HeaderText = "Seri No";
            dataGridView2.Columns[1].HeaderText = "Alıcı Ünvanı";
            dataGridView2.Columns[2].HeaderText = "Alıcı Adresi";
            dataGridView2.Columns[3].HeaderText = "Alıcı V.D No";
            dataGridView2.Columns[4].HeaderText = "Alıcı Telefon No";
            dataGridView2.Columns[5].HeaderText = "Gönderen Ünvanı";
            dataGridView2.Columns[6].HeaderText = "Gönderen Adresi";
            dataGridView2.Columns[7].HeaderText = "Gönderen V.D. No";
            dataGridView2.Columns[8].HeaderText = "Gönderen Telefon No";
            dataGridView2.Columns[9].HeaderText = "Adet 1";
            dataGridView2.Columns[10].HeaderText = "Kap 1";
            dataGridView2.Columns[11].HeaderText = "Cinsi 1";
            dataGridView2.Columns[12].HeaderText = "KG 1";
            dataGridView2.Columns[13].HeaderText = "Adet 2";
            dataGridView2.Columns[14].HeaderText = "Kap 2";
            dataGridView2.Columns[15].HeaderText = "Cinsi 2";
            dataGridView2.Columns[16].HeaderText = "KG  2";
            dataGridView2.Columns[17].HeaderText = "Adet 3";
            dataGridView2.Columns[18].HeaderText = "Kap 3";
            dataGridView2.Columns[19].HeaderText = "Cinsi 3";
            dataGridView2.Columns[20].HeaderText = "KG 3";
            dataGridView2.Columns[21].HeaderText = "Adet 4";
            dataGridView2.Columns[22].HeaderText = "Kap 4";
            dataGridView2.Columns[23].HeaderText = "Cinsi 4";
            dataGridView2.Columns[24].HeaderText = "KG 4";
            dataGridView2.Columns[25].HeaderText = "Ücret";
            dataGridView2.Columns[26].HeaderText = "Havale";
            dataGridView2.Columns[27].HeaderText = "Peşin";
            dataGridView2.Columns[28].HeaderText = "Tarih";
            dataGridView2.Columns[29].HeaderText = "Yer";
        }


        private void kayitGoruntule_Load(object sender, EventArgs e)
        {


            try
            {


                mysqlbaglan.Open();
                string komut = "select *from tasimaIrsaliyeleri";
                MySqlDataAdapter da = new MySqlDataAdapter(komut, mysqlbaglan);
                DataTable dt = new DataTable();
                da.Fill(dt);
                dataGridView1.DataSource = dt;
                dataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect;

                baslik();



                string komut2 = "select *from BtasimaIrsaliyeleri";
                MySqlDataAdapter da2 = new MySqlDataAdapter(komut2, mysqlbaglan);
                DataTable dt2 = new DataTable();
                da2.Fill(dt2);
                dataGridView2.DataSource = dt2;
                dataGridView2.SelectionMode = DataGridViewSelectionMode.FullRowSelect;

                baslik2();


            }
            catch (Exception)
            {
                MessageBox.Show("Kayıt Görüntüleme Başarısız! Lütfen internet bağlantınızı kontrol edip programı yeniden başlatınız.");


            }




        }

        private void button1_Click(object sender, EventArgs e)
        {

        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {
            Form anasayfa = new anasayfa();
            anasayfa.Show();
            this.Hide();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                string komut = "select *from tasimaIrsaliyeleri where irsaliyeNo='" + textBox1.Text + "'";
                MySqlDataAdapter da = new MySqlDataAdapter(komut, mysqlbaglan);
                DataTable dt = new DataTable();
                da.Fill(dt);
                dataGridView1.DataSource = dt;
                baslik();

            }
            catch (Exception)
            {
                MessageBox.Show("İŞLEM BAŞARISIZ! Lütfen Tekrar Deneyin");

            }

        }

        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                string komut = "select *from tasimaIrsaliyeleri where tarih='" + dateTimePicker1.Text + "'";
                MySqlDataAdapter da = new MySqlDataAdapter(komut, mysqlbaglan);
                DataTable dt = new DataTable();
                da.Fill(dt);
                dataGridView1.DataSource = dt;
                baslik();

            }
            catch (Exception)
            {
                MessageBox.Show("İŞLEM BAŞARISIZ! Lütfen Tekrar Deneyin ");

            }

        }

        private void button4_Click(object sender, EventArgs e)
        {
            sil = MessageBox.Show("Gerçekten Silmek İstiyor musunuz ?", "Sil", MessageBoxButtons.YesNo);

            if (sil == DialogResult.Yes)
            {
                try
                {
                    string komut = "delete from tasimaIrsaliyeleri where irsaliyeNo='" + textBox3.Text + "'";
                    MySqlDataAdapter da = new MySqlDataAdapter(komut, mysqlbaglan);
                    DataTable dt = new DataTable();
                    da.Fill(dt);
                    dataGridView1.DataSource = dt;

                    string komut2 = "select *from tasimaIrsaliyeleri";
                    MySqlDataAdapter da2 = new MySqlDataAdapter(komut2, mysqlbaglan);
                    DataTable dt2 = new DataTable();
                    da2.Fill(dt2);
                    dataGridView1.DataSource = dt2;
                    baslik();
                }
                catch (Exception)
                {
                    MessageBox.Show("İŞLEM BAŞARISIZ! Lütfen Tekrar Deneyin ");

                }


            }



        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            try
            {
                string komut = "select *from tasimaIrsaliyeleri where aliciUnvani='" + textBox2.Text + "'";
                MySqlDataAdapter da = new MySqlDataAdapter(komut, mysqlbaglan);
                DataTable dt = new DataTable();
                da.Fill(dt);
                dataGridView1.DataSource = dt;
                baslik();

            }
            catch (Exception)
            {
                MessageBox.Show("İŞLEM BAŞARISIZ! Lütfen Tekrar Deneyin ");

            }
        }

        private void groupBox5_Enter(object sender, EventArgs e)
        {

        }

        private void button5_Click(object sender, EventArgs e)
        {
            try
            {
                string komut = "select *from tasimaIrsaliyeleri";
                MySqlDataAdapter da = new MySqlDataAdapter(komut, mysqlbaglan);
                DataTable dt = new DataTable();
                da.Fill(dt);
                dataGridView1.DataSource = dt;
                baslik();


                string komut2 = "select *from BtasimaIrsaliyeleri";
                MySqlDataAdapter da2 = new MySqlDataAdapter(komut2, mysqlbaglan);
                DataTable dt2 = new DataTable();
                da2.Fill(dt2);
                dataGridView2.DataSource = dt2;
                dataGridView2.SelectionMode = DataGridViewSelectionMode.FullRowSelect;

                baslik2();

            }
            catch (Exception)
            {
                MessageBox.Show("İŞLEM BAŞARISIZ! Lütfen Tekrar Deneyin ");

            }
        }

        private void dataGridView1_RowEnter(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            textBox3.Text = dataGridView1.CurrentRow.Cells[0].Value.ToString();
        }

        private void groupBox5_Enter_1(object sender, EventArgs e)
        {

        }

        private void button8_Click(object sender, EventArgs e)
        {
            try
            {
                string komut = "select *from BtasimaIrsaliyeleri where irsaliyeNo='" + textBox5.Text + "'";
                MySqlDataAdapter da = new MySqlDataAdapter(komut, mysqlbaglan);
                DataTable dt = new DataTable();
                da.Fill(dt);
                dataGridView2.DataSource = dt;
                baslik2();

            }
            catch (Exception)
            {
                MessageBox.Show("İŞLEM BAŞARISIZ! Lütfen Tekrar Deneyin");

            }
        }

        private void button9_Click(object sender, EventArgs e)
        {
            try
            {
                string komut = "select *from BtasimaIrsaliyeleri where tarih='" + dateTimePicker2.Text + "'";
                MySqlDataAdapter da = new MySqlDataAdapter(komut, mysqlbaglan);
                DataTable dt = new DataTable();
                da.Fill(dt);
                dataGridView2.DataSource = dt;
                baslik2();

            }
            catch (Exception)
            {
                MessageBox.Show("İŞLEM BAŞARISIZ! Lütfen Tekrar Deneyin ");

            }
        }

        private void button6_Click(object sender, EventArgs e)
        {

            try
            {
                string komut = "select *from BtasimaIrsaliyeleri where aliciUnvani='" + textBox4.Text + "'";
                MySqlDataAdapter da = new MySqlDataAdapter(komut, mysqlbaglan);
                DataTable dt = new DataTable();
                da.Fill(dt);
                dataGridView2.DataSource = dt;
                baslik2();

            }
            catch (Exception)
            {
                MessageBox.Show("İŞLEM BAŞARISIZ! Lütfen Tekrar Deneyin ");

            }
        }

        private void button10_Click(object sender, EventArgs e)
        {
            sil = MessageBox.Show("Gerçekten Silmek İstiyor musunuz ?", "Sil", MessageBoxButtons.YesNo);

            if (sil == DialogResult.Yes)
            {
                try
                {
                    string komut = "delete from BtasimaIrsaliyeleri where irsaliyeNo='" + textBox6.Text + "'";
                    MySqlDataAdapter da = new MySqlDataAdapter(komut, mysqlbaglan);
                    DataTable dt = new DataTable();
                    da.Fill(dt);
                    dataGridView2.DataSource = dt;

                    string komut2 = "select *from BtasimaIrsaliyeleri";
                    MySqlDataAdapter da2 = new MySqlDataAdapter(komut2, mysqlbaglan);
                    DataTable dt2 = new DataTable();
                    da2.Fill(dt2);
                    dataGridView2.DataSource = dt2;
                    baslik2();
                }
                catch (Exception)
                {
                    MessageBox.Show("İŞLEM BAŞARISIZ! Lütfen Tekrar Deneyin ");

                }

            }
        }

        private void groupBox3_Enter(object sender, EventArgs e)
        {

        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {

        }

        private void dataGridView2_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            textBox6.Text = dataGridView2.CurrentRow.Cells[0].Value.ToString();
        }
    }
}
