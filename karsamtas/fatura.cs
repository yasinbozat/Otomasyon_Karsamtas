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
    public partial class fatura : Form
    {
        string deneme, deneme2;
        public void tesellum()
        {
            try
            {
                string komut3 = "SELECT *from tasimaIrsaliyeleri";
                MySqlCommand kmt3 = new MySqlCommand(komut3, mysqlbaglan);
                mysqlbaglan.Open();
                MySqlDataReader dr = kmt3.ExecuteReader();

                while (dr.Read())
                {
                    comboBox2.Items.Add(dr["irsaliyeNo"].ToString());
                    comboBox3.Items.Add(dr["irsaliyeNo"].ToString());
                    comboBox4.Items.Add(dr["irsaliyeNo"].ToString());
                    comboBox5.Items.Add(dr["irsaliyeNo"].ToString());
                    comboBox6.Items.Add(dr["irsaliyeNo"].ToString());
                    comboBox7.Items.Add(dr["irsaliyeNo"].ToString());
                    
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


        public MySqlConnection mysqlbaglan = new MySqlConnection("Server=karsamtas.com;Database=karsamtasProg;Uid=karsamtas;Pwd='852456456';");

        double sayi1, sayi2, sayi3, sayi4, sayi5, sayi6;
        double kdv;


        public void toplam() {

            try
            {

           
            kdv = 0.18;
            if (textBox56.Text != "")
            {
                sayi1 = Convert.ToInt32(textBox56.Text);
            }
            if (textBox55.Text != "")
            {
                sayi2 = Convert.ToInt32(textBox55.Text);
            }
            if (textBox54.Text != "")
            {
                sayi3 = Convert.ToInt32(textBox54.Text);

            }
            if (textBox53.Text != "")
            {
                sayi4 = Convert.ToInt32(textBox53.Text);

            }
            if (textBox52.Text != "")
            {
                sayi5 = Convert.ToInt32(textBox52.Text);
            }
            if (textBox51.Text != "")
            {
                sayi6 = Convert.ToInt32(textBox51.Text);
            }
  
            double   toplam1 = sayi1 + sayi2 + sayi3 + sayi4 + sayi5 + sayi6;
            double   toplam2 = (toplam1*kdv)+toplam1;
            double   toplam3 = toplam1 * kdv;

            label1.Text = Convert.ToString(toplam1) + " TL";
            label2.Text = Convert.ToString(toplam3) + " TL";
            label3.Text = Convert.ToString(toplam2) + " TL";
            label4.Text = label3.Text;
            }
            catch (Exception)
            {

                MessageBox.Show("Çok Fazla Veya Yanlış Bir Değer Girildi (Lütfen Küsürat Girmeyin)");
            }

        }

        public void sayin()
        {
            try
            {


                string komut = "SELECT *from fatura";
                MySqlCommand kmt = new MySqlCommand(komut, mysqlbaglan);
                mysqlbaglan.Open();
                MySqlDataReader dr = kmt.ExecuteReader();

                while (dr.Read())
                {
                    comboBox1.Items.Add(dr["sayin"].ToString());

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

        public void vdGetir()
        {
            try
            {

            

            string komut = "SELECT *from fatura where sayin ='" + comboBox1.Text + "'  ";
            MySqlCommand kmt = new MySqlCommand(komut, mysqlbaglan);
            mysqlbaglan.Open();
            MySqlDataReader dr = kmt.ExecuteReader();

            while (dr.Read())
            {
                textBox2.Text = dr["vdno"].ToString();
                textBox1.Text = dr["no"].ToString();
                textBox16.Text = dr["adres"].ToString();
                textBox23.Text = dr["adres2"].ToString();


            }
            dr.Close();
            dr.Dispose();
            mysqlbaglan.Close();
            }
            catch (Exception)
            {

                MessageBox.Show("Bir Hata Oluştu","HATA",MessageBoxButtons.OK,MessageBoxIcon.Stop);
            }


        }
       
        ////////////////////////////////////////////////////////////////////////////////////

        string tarihText;

        public void tesellumGetir(string tesellumNo,string tarih)
        {
            try
            {
                string komut = "SELECT *from tasimaIrsaliyeleri where irsaliyeNo ='" + tesellumNo + "'  ";
                MySqlCommand kmt = new MySqlCommand(komut, mysqlbaglan);
                mysqlbaglan.Open();
                MySqlDataReader dr = kmt.ExecuteReader();

                while (dr.Read())
                {
                    tarih = dr["tarih"].ToString();

                    tarihText = tarih;

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
        
       ////////////////////////////////////////////////////////////////////////////////////
     

        public fatura()
        {
            InitializeComponent();
        }

        DialogResult onay = new DialogResult();

        private void fatura_Load(object sender, EventArgs e)
        {
            tesellum();

            
            sayin();

            for (int i = 0; i < comboBox1.Items.Count; i++)
            {

                deneme = Convert.ToString(comboBox1.Items[i]);
                int f = i + 1;
                for (int j = f; j < comboBox1.Items.Count; j++)
                {


                    deneme2 = Convert.ToString(comboBox1.Items[j]);
                    if (deneme == deneme2)
                    {
                        comboBox1.Items.RemoveAt(j);

                    }

                }


            }
            for (int i = 0; i < comboBox1.Items.Count; i++)
            {

                deneme = Convert.ToString(comboBox1.Items[i]);
                int f = i + 1;
                for (int j = f; j < comboBox1.Items.Count; j++)
                {


                    deneme2 = Convert.ToString(comboBox1.Items[j]);
                    if (deneme == deneme2)
                    {
                        comboBox1.Items.RemoveAt(j);

                    }

                }


            }

        }
       
        /////////////////////////////////////////////////////////////////////////////////////////////////
        /////////////////////////////////////////////////////////////////////////////////////////////////
        private void button1_Click(object sender, EventArgs e)
        {

            onay = MessageBox.Show("Onaylıyor musunuz?", "Onay", MessageBoxButtons.YesNo);

            if (onay == DialogResult.Yes)
            {

               
                try
                {
                    mysqlbaglan.Open();
                    string komut = "insert into fatura(sayin,vdno,no,adres,adres2,tarih,fisNo,tarih1,aciklama,adet,kab,cins,kilo,fiyat,tutar,fisNo2,tarih2,aciklama2,adet2,kab2,cins2,kilo2,fiyat2,tutar2,fisNo3,tarih3,aciklama3,adet3,kab3,cins3,kilo3,fiyat3,tutar3,fisNo4,tarih4,aciklama4,adet4,kab4,cins4,kilo4,fiyat4,tutar4,fisNo5,tarih5,aciklama5,adet5,kab5,cins5,kilo5,fiyat5,tutar5,fisNo6,tarih6,aciklama6,adet6,kab6,cins6,kilo6,fiyat6,tutar6,toplam,genelToplam,tesellumNo) values('" + comboBox1.Text + "', '" + textBox2.Text + "','" + textBox1.Text + "','" + textBox16.Text + "','" + textBox23.Text + "', '" + dateTimePicker1.Text + "', '" + comboBox2.Text + "', '" + textBox4.Text + "', '" + textBox5.Text + "', '" + textBox6.Text + "', '" + textBox7.Text + "', '" + textBox8.Text + "', '" + textBox9.Text + "','" + textBox50.Text + "','" + textBox56.Text + "','" + comboBox3.Text + "', '" + textBox15.Text + "', '" + textBox14.Text + "', '" + textBox13.Text + "', '" + textBox12.Text + "', '" + textBox11.Text + "', '" + textBox10.Text + "','" + textBox49.Text + "','" + textBox55.Text + "','" + comboBox4.Text + "', '" + textBox22.Text + "', '" + textBox21.Text + "', '" + textBox20.Text + "', '" + textBox19.Text + "', '" + textBox18.Text + "', '" + textBox17.Text + "','" + textBox48.Text + "','" + textBox54.Text + "','" + comboBox5.Text + "', '" + textBox29.Text + "', '" + textBox28.Text + "', '" + textBox27.Text + "', '" + textBox26.Text + "', '" + textBox25.Text + "', '" + textBox24.Text + "','" + textBox47.Text + "','" + textBox53.Text + "','" + comboBox6.Text + "', '" + textBox36.Text + "', '" + textBox35.Text + "', '" + textBox34.Text + "', '" + textBox33.Text + "', '" + textBox32.Text + "', '" + textBox31.Text + "','" + textBox46.Text + "','" + textBox52.Text + "','" + comboBox7.Text + "', '" + textBox43.Text + "', '" + textBox42.Text + "', '" + textBox41.Text + "', '" + textBox40.Text + "', '" + textBox39.Text + "', '" + textBox38.Text + "','" + textBox45.Text + "','" + textBox51.Text + "','" + label1.Text + "','" + label2.Text + "','" + textBox3.Text + "')";
                    MySqlCommand kmt = new MySqlCommand(komut, mysqlbaglan);
                    kmt.ExecuteNonQuery();
                    
                    Microsoft.Office.Interop.Excel.Application objExcel = new Microsoft.Office.Interop.Excel.Application();
                    objExcel.Visible = true;
                    Microsoft.Office.Interop.Excel.Workbook objBook = objExcel.Workbooks.Open("c:\\fatura.xlsx", 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                    Microsoft.Office.Interop.Excel.Worksheet objSheet = (Microsoft.Office.Interop.Excel.Worksheet)objBook.Worksheets.get_Item(1);
                    Microsoft.Office.Interop.Excel.Range objRange;


                    objRange = objSheet.get_Range("K7", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, dateTimePicker1.Text);
                    objRange = objSheet.get_Range("b7", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox2.Text);
                    objRange = objSheet.get_Range("e7", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox1.Text);
                    objRange = objSheet.get_Range("k4", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox3.Text);

                    objRange = objSheet.get_Range("b3", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, comboBox1.Text);
                    
                    objRange = objSheet.get_Range("A4", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox16.Text);
                   
                    objRange = objSheet.get_Range("A5", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox23.Text);
                    
                    objRange = objSheet.get_Range("A23", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox30.Text);
                    // 1. SIRA
                    objRange = objSheet.get_Range("A10", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, comboBox2.Text);
                    objRange = objSheet.get_Range("B10", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox4.Text);
                    objRange = objSheet.get_Range("C10", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox5.Text);
                    objRange = objSheet.get_Range("F10", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox6.Text);
                    objRange = objSheet.get_Range("G10", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox7.Text);
                    objRange = objSheet.get_Range("H10", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox8.Text);
                    objRange = objSheet.get_Range("I10", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox9.Text);
                    objRange = objSheet.get_Range("J10", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox50.Text);
                    objRange = objSheet.get_Range("K10", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox56.Text);

                    // 2. SIRA
                    objRange = objSheet.get_Range("A12", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, comboBox3.Text);
                    objRange = objSheet.get_Range("B12", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox15.Text);
                    objRange = objSheet.get_Range("C12", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox14.Text);
                    objRange = objSheet.get_Range("F12", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox13.Text);
                    objRange = objSheet.get_Range("G12", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox12.Text);
                    objRange = objSheet.get_Range("H12", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox11.Text);
                    objRange = objSheet.get_Range("I12", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox10.Text);
                    objRange = objSheet.get_Range("J12", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox49.Text);
                    objRange = objSheet.get_Range("K12", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox55.Text);
                     
                    //3. SIRA      
                    objRange = objSheet.get_Range("A14", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, comboBox4.Text);
                    objRange = objSheet.get_Range("B14", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox22.Text);
                    objRange = objSheet.get_Range("C14", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox21.Text);
                    objRange = objSheet.get_Range("F14", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox20.Text);
                    objRange = objSheet.get_Range("G14", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox19.Text);
                    objRange = objSheet.get_Range("H14", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox18.Text);
                    objRange = objSheet.get_Range("I14", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox17.Text);
                    objRange = objSheet.get_Range("J14", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox48.Text);
                    objRange = objSheet.get_Range("K14", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox54.Text);

                    //4. SIRA
                    objRange = objSheet.get_Range("A16", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, comboBox5.Text);
                    objRange = objSheet.get_Range("B16", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox29.Text);
                    objRange = objSheet.get_Range("C16", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox28.Text);
                    objRange = objSheet.get_Range("F16", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox27.Text);
                    objRange = objSheet.get_Range("G16", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox26.Text);
                    objRange = objSheet.get_Range("H16", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox25.Text);
                    objRange = objSheet.get_Range("I16", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox24.Text);
                    objRange = objSheet.get_Range("J16", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox47.Text);
                    objRange = objSheet.get_Range("K16", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox53.Text);

                    //5. SIRA
                    objRange = objSheet.get_Range("A18", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, comboBox6.Text);
                    objRange = objSheet.get_Range("B18", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox36.Text);
                    objRange = objSheet.get_Range("C18", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox35.Text);
                    objRange = objSheet.get_Range("F18", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox34.Text);
                    objRange = objSheet.get_Range("G18", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox33.Text);
                    objRange = objSheet.get_Range("H18", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox32.Text);
                    objRange = objSheet.get_Range("I18", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox31.Text);
                    objRange = objSheet.get_Range("J18", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox46.Text);
                    objRange = objSheet.get_Range("K18", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox52.Text);

                    //6. SIRA
                    objRange = objSheet.get_Range("A20", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, comboBox7.Text);
                    objRange = objSheet.get_Range("B20", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox43.Text);
                    objRange = objSheet.get_Range("C20", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox42.Text);
                    objRange = objSheet.get_Range("F20", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox41.Text);
                    objRange = objSheet.get_Range("G20", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox40.Text);
                    objRange = objSheet.get_Range("H20", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox39.Text);
                    objRange = objSheet.get_Range("I20", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox38.Text);
                    objRange = objSheet.get_Range("J20", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox45.Text);
                    objRange = objSheet.get_Range("K20", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox51.Text);

                    objSheet.PrintOutEx(1, 1, 2, true);
                    Form anasayfa = new anasayfa();
                    anasayfa.Show();
                    this.Hide();

                }
              catch (Exception h)
                {
                    MessageBox.Show("Beklenmeyen bir hata oluştu lütfen tekrar deneyiniz"+h, "HATA", MessageBoxButtons.OK, MessageBoxIcon.Stop);

                }
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {

            Form anasayfa = new anasayfa();
            anasayfa.Show();
            this.Hide();

        }

        private void button2_Click(object sender, EventArgs e)
        {
           
        }

        private void textBox9_TextChanged(object sender, EventArgs e)
        {

        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            tesellumGetir(comboBox2.Text,textBox4.Text);
            textBox4.Text = tarihText;

        }

        private void button4_Click(object sender, EventArgs e)
        {
            
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            vdGetir();
        }

        private void textBox50_TextChanged(object sender, EventArgs e)
        {

          
           
           
            
        }

        private void textBox56_TextChanged(object sender, EventArgs e)
        {
            toplam();
        }

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            tesellumGetir(comboBox3.Text, textBox15.Text);
            textBox15.Text = tarihText;
        }

        private void comboBox4_SelectedIndexChanged(object sender, EventArgs e)
        {
            tesellumGetir(comboBox4.Text, textBox22.Text);
            textBox22.Text = tarihText;
        }

        private void comboBox5_SelectedIndexChanged(object sender, EventArgs e)
        {
            tesellumGetir(comboBox5.Text, textBox29.Text);
            textBox29.Text = tarihText;
        }

        private void comboBox6_SelectedIndexChanged(object sender, EventArgs e)
        {
            tesellumGetir(comboBox6.Text, textBox36.Text);
            textBox36.Text = tarihText;
        }

        private void comboBox7_SelectedIndexChanged(object sender, EventArgs e)
        {
            tesellumGetir(comboBox7.Text, textBox43.Text);
            textBox43.Text = tarihText;
        }

        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
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

        private void textBox56_KeyPress(object sender, KeyPressEventArgs e)
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

        private void textBox55_KeyPress(object sender, KeyPressEventArgs e)
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

        private void textBox54_KeyPress(object sender, KeyPressEventArgs e)
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

        private void textBox53_KeyPress(object sender, KeyPressEventArgs e)
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

        private void textBox52_KeyPress(object sender, KeyPressEventArgs e)
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

        private void textBox51_KeyPress(object sender, KeyPressEventArgs e)
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

        private void textBox55_TextChanged(object sender, EventArgs e)
        {
            toplam();
        }

        private void textBox54_TextChanged(object sender, EventArgs e)
        {
            toplam();
        }

        private void textBox53_TextChanged(object sender, EventArgs e)
        {
            toplam();
        }

        private void textBox52_TextChanged(object sender, EventArgs e)
        {
            toplam();
        }

        private void textBox51_TextChanged(object sender, EventArgs e)
        {
            toplam();
        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

       
    }
}
