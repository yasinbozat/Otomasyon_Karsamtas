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
    public partial class tasimaIrsaliyesi : Form
    {
        public tasimaIrsaliyesi()
        {
            InitializeComponent();
        }

        string deneme, deneme2;

        public MySqlConnection mysqlbaglan = new MySqlConnection("Server=karsamtas.com;Database=karsamtasProg;Uid=karsamtas;Pwd='852456456';");


        public void yazdir()
        {

            Microsoft.Office.Interop.Excel.Application objExcel = new Microsoft.Office.Interop.Excel.Application();
            objExcel.Visible = true;
            Microsoft.Office.Interop.Excel.Workbook objBook = objExcel.Workbooks.Open("c:\\ambartesellum.xlsx", 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            Microsoft.Office.Interop.Excel.Worksheet objSheet = (Microsoft.Office.Interop.Excel.Worksheet)objBook.Worksheets.get_Item(1);
            Microsoft.Office.Interop.Excel.Range objRange;


            objRange = objSheet.get_Range("J4", System.Reflection.Missing.Value);
            objRange.set_Value(System.Reflection.Missing.Value, dateTimePicker2.Text);
            objRange = objSheet.get_Range("C4", System.Reflection.Missing.Value);
            objRange.set_Value(System.Reflection.Missing.Value, textBox7.Text);

            //ALICI BİLGİLERİ
            objRange = objSheet.get_Range("A7", System.Reflection.Missing.Value);
            objRange.set_Value(System.Reflection.Missing.Value, comboBox2.Text);
            objRange = objSheet.get_Range("A8", System.Reflection.Missing.Value);
            objRange.set_Value(System.Reflection.Missing.Value, textBox1.Text);
            objRange = objSheet.get_Range("C9", System.Reflection.Missing.Value);
            objRange.set_Value(System.Reflection.Missing.Value, textBox2.Text);
            objRange = objSheet.get_Range("C10", System.Reflection.Missing.Value);
            objRange.set_Value(System.Reflection.Missing.Value, textBox4.Text);
            ///////////////////////////////////////////////////////////////////

            //GÖNDEREN BİLGİLERİ
            objRange = objSheet.get_Range("G7", System.Reflection.Missing.Value);
            objRange.set_Value(System.Reflection.Missing.Value, comboBox4.Text);
            objRange = objSheet.get_Range("G8", System.Reflection.Missing.Value);
            objRange.set_Value(System.Reflection.Missing.Value, textBox6.Text);
            objRange = objSheet.get_Range("I9", System.Reflection.Missing.Value);
            objRange.set_Value(System.Reflection.Missing.Value, textBox8.Text);
            objRange = objSheet.get_Range("I10", System.Reflection.Missing.Value);
            objRange.set_Value(System.Reflection.Missing.Value, textBox5.Text);
            ///////////////////////////////////////////////////////////////////

            //MALIN ÖZELLİKLERİ
            //ADET
            objRange = objSheet.get_Range("A13", System.Reflection.Missing.Value);
            objRange.set_Value(System.Reflection.Missing.Value, textBox9.Text);
            objRange = objSheet.get_Range("A14", System.Reflection.Missing.Value);
            objRange.set_Value(System.Reflection.Missing.Value, textBox10.Text);
            objRange = objSheet.get_Range("A15", System.Reflection.Missing.Value);
            objRange.set_Value(System.Reflection.Missing.Value, textBox11.Text);
            objRange = objSheet.get_Range("A16", System.Reflection.Missing.Value);
            objRange.set_Value(System.Reflection.Missing.Value, textBox12.Text);
            ////////////////////////////////////////////////////////////////////
            //KAB
            objRange = objSheet.get_Range("B13", System.Reflection.Missing.Value);
            objRange.set_Value(System.Reflection.Missing.Value, textBox14.Text);
            objRange = objSheet.get_Range("B14", System.Reflection.Missing.Value);
            objRange.set_Value(System.Reflection.Missing.Value, textBox15.Text);
            objRange = objSheet.get_Range("B15", System.Reflection.Missing.Value);
            objRange.set_Value(System.Reflection.Missing.Value, textBox16.Text);
            objRange = objSheet.get_Range("B16", System.Reflection.Missing.Value);
            objRange.set_Value(System.Reflection.Missing.Value, textBox17.Text);
            ////////////////////////////////////////////////////////////////////
            //CİNSİ
            objRange = objSheet.get_Range("C13", System.Reflection.Missing.Value);
            objRange.set_Value(System.Reflection.Missing.Value, textBox19.Text);
            objRange = objSheet.get_Range("C14", System.Reflection.Missing.Value);
            objRange.set_Value(System.Reflection.Missing.Value, textBox20.Text);
            objRange = objSheet.get_Range("C15", System.Reflection.Missing.Value);
            objRange.set_Value(System.Reflection.Missing.Value, textBox21.Text);
            objRange = objSheet.get_Range("C16", System.Reflection.Missing.Value);
            objRange.set_Value(System.Reflection.Missing.Value, textBox22.Text);
            ////////////////////////////////////////////////////////////////////
            //KG
            objRange = objSheet.get_Range("F13", System.Reflection.Missing.Value);
            objRange.set_Value(System.Reflection.Missing.Value, textBox24.Text);
            objRange = objSheet.get_Range("F14", System.Reflection.Missing.Value);
            objRange.set_Value(System.Reflection.Missing.Value, textBox25.Text);
            objRange = objSheet.get_Range("F15", System.Reflection.Missing.Value);
            objRange.set_Value(System.Reflection.Missing.Value, textBox26.Text);
            objRange = objSheet.get_Range("F16", System.Reflection.Missing.Value);
            objRange.set_Value(System.Reflection.Missing.Value, textBox27.Text);
            ////////////////////////////////////////////////////////////////////
            objRange = objSheet.get_Range("K16", System.Reflection.Missing.Value);
            objRange.set_Value(System.Reflection.Missing.Value, textBox40.Text);
            objRange = objSheet.get_Range("G16", System.Reflection.Missing.Value);
            objRange.set_Value(System.Reflection.Missing.Value, textBox39.Text);
            objRange = objSheet.get_Range("I16", System.Reflection.Missing.Value);
            objRange.set_Value(System.Reflection.Missing.Value, textBox3.Text);

            objSheet.PrintOutEx(1, 1, 2, true);
            Form anasayfa = new anasayfa();
            anasayfa.Show();
            this.Hide();


        }


        public void aliciUnvan()
        {
            try
            {



                string komut3 = "SELECT *from tasimaIrsaliyeleri";
                MySqlCommand kmt3 = new MySqlCommand(komut3, mysqlbaglan);
                mysqlbaglan.Open();
                MySqlDataReader dr = kmt3.ExecuteReader();

                while (dr.Read())
                {

                    comboBox2.Items.Add(dr["aliciUnvani"].ToString());

                }
                dr.Close();
                dr.Dispose();
                mysqlbaglan.Close();
            }
            catch (Exception)
            {

                MessageBox.Show("Bir Hata Oluştu", "HATA", MessageBoxButtons.OK, MessageBoxIcon.Stop);
            }

        }

        public void aliciGetir()
        {
            try
            {


                string komut4 = "SELECT *from tasimaIrsaliyeleri where aliciUnvani ='" + comboBox2.Text + "'  ";
                MySqlCommand kmt4 = new MySqlCommand(komut4, mysqlbaglan);
                mysqlbaglan.Open();
                MySqlDataReader dr = kmt4.ExecuteReader();

                while (dr.Read())
                {
                    textBox1.Text = dr["aliciAdres"].ToString();

                }
                dr.Close();
                dr.Dispose();
                mysqlbaglan.Close();
            }
            catch (Exception)
            {

                MessageBox.Show("Bir Hata Oluştu", "HATA", MessageBoxButtons.OK, MessageBoxIcon.Stop);
            }


        }

        public void aliciGetirVDNO()
        {
            try
            {



                string komut5 = "SELECT *from tasimaIrsaliyeleri where aliciUnvani ='" + comboBox2.Text + "'  ";
                MySqlCommand kmt5 = new MySqlCommand(komut5, mysqlbaglan);
                mysqlbaglan.Open();
                MySqlDataReader dr = kmt5.ExecuteReader();

                while (dr.Read())
                {
                    textBox2.Text = dr["aliciVDNo"].ToString();

                }
                dr.Close();
                dr.Dispose();
                mysqlbaglan.Close();
            }
            catch (Exception)
            {

                MessageBox.Show("Bir Hata Oluştu", "HATA", MessageBoxButtons.OK, MessageBoxIcon.Stop);
            }
        }
        public void aliciTel()
        {
            try
            {


                string komut4 = "SELECT *from tasimaIrsaliyeleri where aliciUnvani ='" + comboBox2.Text + "'  ";
                MySqlCommand kmt4 = new MySqlCommand(komut4, mysqlbaglan);
                mysqlbaglan.Open();
                MySqlDataReader dr = kmt4.ExecuteReader();

                while (dr.Read())
                {
                    textBox4.Text = dr["aliciTel"].ToString();

                }
                dr.Close();
                dr.Dispose();
                mysqlbaglan.Close();
            }
            catch (Exception)
            {

                MessageBox.Show("Bir Hata Oluştu", "HATA", MessageBoxButtons.OK, MessageBoxIcon.Stop);
            }


        }

        ///////////////////////GÖNDEREN//////////////////////////////////

        public void gonderenUnvan()
        {
            try
            {


                string komut3 = "SELECT *from tasimaIrsaliyeleri";
                MySqlCommand kmt3 = new MySqlCommand(komut3, mysqlbaglan);
                mysqlbaglan.Open();
                MySqlDataReader dr = kmt3.ExecuteReader();

                while (dr.Read())
                {
                    comboBox4.Items.Add(dr["gonderenUnvani"].ToString());

                }
                dr.Close();
                dr.Dispose();
                mysqlbaglan.Close();

                
            }
            catch (Exception)
            {

                MessageBox.Show("Bir Hata Oluştu", "HATA", MessageBoxButtons.OK, MessageBoxIcon.Stop);
            }



        }
        public void gonderenGetir()
        {
            try
            {


                string komut4 = "SELECT *from tasimaIrsaliyeleri where gonderenUnvani ='" + comboBox4.Text + "'  ";
                MySqlCommand kmt4 = new MySqlCommand(komut4, mysqlbaglan);
                mysqlbaglan.Open();
                MySqlDataReader dr = kmt4.ExecuteReader();

                while (dr.Read())
                {
                    textBox6.Text = dr["gonderenAdres"].ToString();

                }
                dr.Close();
                dr.Dispose();
                mysqlbaglan.Close();
            }
            catch (Exception)
            {

                MessageBox.Show("Bir Hata Oluştu", "HATA", MessageBoxButtons.OK, MessageBoxIcon.Stop);
            }

        }

        public void gonderenGetirVDNO()
        {
            try
            {



                string komut5 = "SELECT *from tasimaIrsaliyeleri where gonderenUnvani ='" + comboBox4.Text + "'  ";
                MySqlCommand kmt5 = new MySqlCommand(komut5, mysqlbaglan);
                mysqlbaglan.Open();
                MySqlDataReader dr = kmt5.ExecuteReader();

                while (dr.Read())
                {
                    textBox8.Text = dr["gonderenVDNo"].ToString();

                }
                dr.Close();
                dr.Dispose();
                mysqlbaglan.Close();
            }
            catch (Exception)
            {

                MessageBox.Show("Bir Hata Oluştu", "HATA", MessageBoxButtons.OK, MessageBoxIcon.Stop);
            }


        }
        public void gonderenTel()
        {
            try
            {


                string komut6 = "SELECT *from tasimaIrsaliyeleri where gonderenUnvani ='" + comboBox4.Text + "'  ";
                MySqlCommand kmt6 = new MySqlCommand(komut6, mysqlbaglan);
                mysqlbaglan.Open();
                MySqlDataReader dr = kmt6.ExecuteReader();

                while (dr.Read())
                {
                    textBox5.Text = dr["gonderenTel"].ToString();

                }
                dr.Close();
                dr.Dispose();
                mysqlbaglan.Close();
            }
            catch (Exception)
            {

                MessageBox.Show("Bir Hata Oluştu", "HATA", MessageBoxButtons.OK, MessageBoxIcon.Stop);
            }

        }

        /////////////////////////////////////////////////////////////////
        


        /////////////////////////////////////////////////////////////////
        /// </summary>


        DialogResult onay = new DialogResult();


        private void tasimaIrsaliyesi_Load(object sender, EventArgs e)
        {
            comboBox1.SelectedItem = "A";
            aliciUnvan();
            gonderenUnvan();
            

           
            for (int i = 0; i <comboBox2.Items.Count; i++)
            {
                
                deneme = Convert.ToString(comboBox2.Items[i]);
                int f = i  + 1; 
                for (int j = f ; j < comboBox2.Items.Count; j++)
                {
                    

                    deneme2 = Convert.ToString(comboBox2.Items[j]);
                    if (deneme == deneme2)
                    {
                        comboBox2.Items.RemoveAt(j);

                    }

                }

               
            }


            for (int i = 0; i < comboBox2.Items.Count; i++)
            {

                deneme = Convert.ToString(comboBox2.Items[i]);
                int f = i + 1;
                for (int j = f; j < comboBox2.Items.Count; j++)
                {


                    deneme2 = Convert.ToString(comboBox2.Items[j]);
                    if (deneme == deneme2)
                    {
                        comboBox2.Items.RemoveAt(j);

                    }

                }


            }

            for (int i = 0; i < comboBox4.Items.Count; i++)
            {

                deneme = Convert.ToString(comboBox4.Items[i]);
                int f = i + 1;
                for (int j = f; j < comboBox4.Items.Count; j++)
                {


                    deneme2 = Convert.ToString(comboBox4.Items[j]);
                    if (deneme == deneme2)
                    {
                        comboBox4.Items.RemoveAt(j);

                    }

                }


            }
            for (int i = 0; i < comboBox4.Items.Count; i++)
            {

                deneme = Convert.ToString(comboBox4.Items[i]);
                int f = i + 1;
                for (int j = f; j < comboBox4.Items.Count; j++)
                {


                    deneme2 = Convert.ToString(comboBox4.Items[j]);
                    if (deneme == deneme2)
                    {
                        comboBox4.Items.RemoveAt(j);

                    }

                }


            }


            



        }

        private void button1_Click(object sender, EventArgs e)
        {
            onay = MessageBox.Show("Onaylıyor musunuz?", "Onay", MessageBoxButtons.YesNo);
            if (onay == DialogResult.Yes)
            {
                if (comboBox1.Text == "B")
                {
                    try
                    {
                        if (textBox41.Text != "" && comboBox1.Text != "")
                        {
                            mysqlbaglan.Open();
                            string komut = "insert into BtasimaIrsaliyeleri(irsaliyeNo,aliciUnvani,aliciAdres,aliciVDNo,aliciTel,gonderenUnvani,gonderenAdres,gonderenVDNo,gonderenTel,adet,kap,cinsi,kgDesi,adet2,kap2,cinsi2,kgDesi2,adet3,kap3,cinsi3,kgDesi3,adet4,kap4,cinsi4,kgDesi4,ucret,havale,pesin,tarih,yer) values('" + textBox41.Text + "', '" + comboBox2.Text + "','" + textBox1.Text + "','" + textBox2.Text + "','" + textBox4.Text + "', '" + comboBox4.Text + "','" + textBox6.Text + "','" + textBox8.Text + "','" + textBox5.Text + "', '" + textBox9.Text + "','" + textBox14.Text + "', '" + textBox19.Text + "','" + textBox24.Text + "', '" + textBox10.Text + "','" + textBox15.Text + "', '" + textBox20.Text + "','" + textBox25.Text + "', '" + textBox11.Text + "','" + textBox16.Text + "', '" + textBox21.Text + "','" + textBox26.Text + "','" + textBox12.Text + "','" + textBox17.Text + "', '" + textBox22.Text + "', '" + textBox27.Text + "','" + textBox39.Text + "','" + textBox40.Text + "', '" + textBox3.Text + "', '" + dateTimePicker2.Text + "','" + textBox7.Text + "')";
                            MySqlCommand kmt = new MySqlCommand(komut, mysqlbaglan);
                            kmt.ExecuteNonQuery();

                            yazdir();


                        }
                        else
                        {

                            MessageBox.Show(" 'İrsaliye No' boş olamaz.", "HATA", MessageBoxButtons.OK, MessageBoxIcon.Stop);

                        }
                    }
                    catch (Exception h)
                    {
                        MessageBox.Show("Beklenmeyen bir hata oluştu lütfen tekrar deneyiniz" + h, "HATA", MessageBoxButtons.OK, MessageBoxIcon.Stop);

                    }
                }

                if (comboBox1.Text == "A")
                {

                    if (textBox41.Text != "" && comboBox1.Text != "")
                    {
                        try
                        {
                            mysqlbaglan.Open();
                            string komut = "insert into tasimaIrsaliyeleri(irsaliyeNo,aliciUnvani,aliciAdres,aliciVDNo,aliciTel,gonderenUnvani,gonderenAdres,gonderenVDNo,gonderenTel,adet,kap,cinsi,kgDesi,adet2,kap2,cinsi2,kgDesi2,adet3,kap3,cinsi3,kgDesi3,adet4,kap4,cinsi4,kgDesi4,ucret,havale,pesin,tarih,yer) values('" + textBox41.Text + "', '" + comboBox2.Text + "','" + textBox1.Text + "','" + textBox2.Text + "','" + textBox4.Text + "', '" + comboBox4.Text + "','" + textBox6.Text + "','" + textBox8.Text + "','" + textBox5.Text + "', '" + textBox9.Text + "','" + textBox14.Text + "', '" + textBox19.Text + "','" + textBox24.Text + "', '" + textBox10.Text + "','" + textBox15.Text + "', '" + textBox20.Text + "','" + textBox25.Text + "', '" + textBox11.Text + "','" + textBox16.Text + "', '" + textBox21.Text + "','" + textBox26.Text + "','" + textBox12.Text + "','" + textBox17.Text + "', '" + textBox22.Text + "', '" + textBox27.Text + "','" + textBox39.Text + "','" + textBox40.Text + "', '" + textBox3.Text + "', '" + dateTimePicker2.Text + "','" + textBox7.Text + "')";
                            MySqlCommand kmt = new MySqlCommand(komut, mysqlbaglan);
                            kmt.ExecuteNonQuery();

                            yazdir();

                        }

                        catch (Exception h)
                        {

                            MessageBox.Show("Beklenmeyen bir hata oluştu lütfen internet bağlantınızı kontrol edip tekrar deneyin." + h);
                        }
                    }
                    else
                    {
                        MessageBox.Show(" 'İrsaliye No' boş olamaz.", "HATA", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    }
                }

                if (comboBox1.Text == "")
                {
                    MessageBox.Show(" 'Dosya Tipi' boş olamaz.", "HATA", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                }
            }
        }



        private void textBox3_TextChanged(object sender, EventArgs e)
        {

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {

            Form anasayfa = new anasayfa();
            anasayfa.Show();
            this.Hide();
        }

        private void textBox2_KeyPress(object sender, KeyPressEventArgs e)
        {

        }

        private void textBox6_KeyPress(object sender, KeyPressEventArgs e)
        {
            
        }

        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {

        }

        private void comboBox1_KeyPress(object sender, KeyPressEventArgs e)
        {

        }



        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            aliciGetir();
            aliciGetirVDNO();
            aliciTel();

        }

        private void textBox1_KeyPress_1(object sender, KeyPressEventArgs e)
        {
          
        }

        private void textBox2_KeyPress_1(object sender, KeyPressEventArgs e)
        {
          
        }

        private void textBox4_KeyPress(object sender, KeyPressEventArgs e)
        {
           
        }

        private void textBox8_KeyPress(object sender, KeyPressEventArgs e)
        {
           
        }

        private void textBox5_KeyPress(object sender, KeyPressEventArgs e)
        {
           
        }

        private void textBox9_KeyPress(object sender, KeyPressEventArgs e)
        {
           
        }

        private void textBox10_KeyPress(object sender, KeyPressEventArgs e)
        {
            
        }

        private void textBox11_KeyPress(object sender, KeyPressEventArgs e)
        {
            
        }

        private void textBox12_KeyPress(object sender, KeyPressEventArgs e)
        {
           
        }

        private void textBox14_KeyPress(object sender, KeyPressEventArgs e)
        {
           
        }

        private void textBox15_KeyPress(object sender, KeyPressEventArgs e)
        {
            
        }

        private void textBox16_KeyPress(object sender, KeyPressEventArgs e)
        {
           
        }

        private void textBox17_KeyPress(object sender, KeyPressEventArgs e)
        {
           
        }

        private void textBox19_KeyPress(object sender, KeyPressEventArgs e)
        {
            
        }

        private void textBox20_KeyPress(object sender, KeyPressEventArgs e)
        {
            
        }

        private void textBox21_KeyPress(object sender, KeyPressEventArgs e)
        {
           
        }

        private void textBox22_KeyPress(object sender, KeyPressEventArgs e)
        {
           
        }

        private void textBox24_KeyPress(object sender, KeyPressEventArgs e)
        {
           
        }

        private void textBox25_KeyPress(object sender, KeyPressEventArgs e)
        {

        }

        private void textBox26_KeyPress(object sender, KeyPressEventArgs e)
        {
            
        }

        private void textBox27_KeyPress(object sender, KeyPressEventArgs e)
        {
            
        }

        private void textBox7_KeyPress(object sender, KeyPressEventArgs e)
        {
            
        }

        private void textBox39_KeyPress(object sender, KeyPressEventArgs e)
        {
           
        }

        private void textBox3_KeyPress(object sender, KeyPressEventArgs e)
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

        private void textBox40_KeyPress(object sender, KeyPressEventArgs e)
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

        private void textBox3_TextChanged_1(object sender, EventArgs e)
        {

        }

        private void comboBox4_KeyPress(object sender, KeyPressEventArgs e)
        {
          
        }

        private void textBox41_KeyPress(object sender, KeyPressEventArgs e)
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

        private void comboBox4_SelectedIndexChanged(object sender, EventArgs e)
        {
            gonderenGetir();
            gonderenGetirVDNO();
            gonderenTel();
        }
    }
}



