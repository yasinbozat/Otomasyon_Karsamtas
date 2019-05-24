using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using MySql.Data.MySqlClient;
using MySql.Data;

namespace karsamtas
{

    public partial class topluTasima : Form
    {
        string irsaliyeNo;
        string gonderenText;
        string aliciText;
        string yerText;
        string ucretText;
       
        
        public MySqlConnection mysqlbaglan = new MySqlConnection("Server=karsamtas.com;Database=karsamtasProg;Uid=karsamtas;Pwd='852456456';");


        public void tesellumNo()
        {
            try
            {
                string komut = "SELECT *from tasimaIrsaliyeleri";
                MySqlCommand kmt = new MySqlCommand(komut, mysqlbaglan);
                mysqlbaglan.Open();
                MySqlDataReader dr = kmt.ExecuteReader();

                while (dr.Read())
                {
                    textBox1.Items.Add(dr["irsaliyeNo"].ToString());
                    textBox22.Items.Add(dr["irsaliyeNo"].ToString());
                    textBox33.Items.Add(dr["irsaliyeNo"].ToString());
                    textBox44.Items.Add(dr["irsaliyeNo"].ToString());
                    textBox55.Items.Add(dr["irsaliyeNo"].ToString());
                    textBox66.Items.Add(dr["irsaliyeNo"].ToString());
                    textBox77.Items.Add(dr["irsaliyeNo"].ToString());
                    textBox88.Items.Add(dr["irsaliyeNo"].ToString());
                    textBox99.Items.Add(dr["irsaliyeNo"].ToString());
                    textBox110.Items.Add(dr["irsaliyeNo"].ToString());
                    textBox121.Items.Add(dr["irsaliyeNo"].ToString());
                    textBox132.Items.Add(dr["irsaliyeNo"].ToString());
                    textBox143.Items.Add(dr["irsaliyeNo"].ToString());
                    textBox154.Items.Add(dr["irsaliyeNo"].ToString());
                    textBox165.Items.Add(dr["irsaliyeNo"].ToString());
                    textBox176.Items.Add(dr["irsaliyeNo"].ToString());
                    textBox187.Items.Add(dr["irsaliyeNo"].ToString());
                    textBox198.Items.Add(dr["irsaliyeNo"].ToString()); 
                    textBox209.Items.Add(dr["irsaliyeNo"].ToString());
                    
    
    
    
    
    
    
    
    
    
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
        ///////////////////////////////////////////////////////////////////////////////////////////////////

        public void bilgiGetir(string irsaliye ,string gonderen ,string alici, string yer,string ucret)
        {
            try
            {

              

                string komut = "SELECT *from tasimaIrsaliyeleri where irsaliyeNo = '"+irsaliye+"'";
                MySqlCommand kmt = new MySqlCommand(komut, mysqlbaglan);
                mysqlbaglan.Open();
                MySqlDataReader dr = kmt.ExecuteReader();


                while (dr.Read())
                {
                    gonderen = dr["gonderenUnvani"].ToString();
                    alici = dr["aliciUnvani"].ToString();
                    ucret = dr["ucret"].ToString();
                    yer = dr["yer"].ToString();

                    gonderenText = gonderen;
                    aliciText = alici;
                    yerText = yer;
                    ucretText = ucret;


                }
                dr.Close();
                dr.Dispose();
                mysqlbaglan.Close();
            }
            catch (Exception e )
            {

                MessageBox.Show("Bir Hata Oluştu"+e, "HATA", MessageBoxButtons.OK, MessageBoxIcon.Stop);
            }


        }




        /// <summary>
        /// ///////////////////////////////////////////////////////////////////////////////////////////////
        /// </summary>

        public topluTasima()
        {
            InitializeComponent();
        }

        DialogResult onay = new DialogResult();

        private void topluTasima_Load(object sender, EventArgs e)
        {

            tesellumNo();

        


        }

        private void button1_Click(object sender, EventArgs e)
        {
            onay = MessageBox.Show("Onaylıyor musunuz?", "Onay", MessageBoxButtons.YesNo);
            if (onay == DialogResult.Yes)
            {
                try
                {


                    Microsoft.Office.Interop.Excel.Application objExcel = new Microsoft.Office.Interop.Excel.Application();
                    objExcel.Visible = true;
                    Microsoft.Office.Interop.Excel.Workbook objBook = objExcel.Workbooks.Open("c:\\toplutasima.xlsx", 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                    Microsoft.Office.Interop.Excel.Worksheet objSheet = (Microsoft.Office.Interop.Excel.Worksheet)objBook.Worksheets.get_Item(1);
                    Microsoft.Office.Interop.Excel.Range objRange;

                    objRange = objSheet.get_Range("A12", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox1.Text);
                    objRange = objSheet.get_Range("B12", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox2.Text);
                    objRange = objSheet.get_Range("E12", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox3.Text);
                    objRange = objSheet.get_Range("H12", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox4.Text);
                    objRange = objSheet.get_Range("J12", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox5.Text);
                    objRange = objSheet.get_Range("K12", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox6.Text);
                    objRange = objSheet.get_Range("P12", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox7.Text);
                    objRange = objSheet.get_Range("Q12", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox8.Text);
                    objRange = objSheet.get_Range("R12", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox9.Text);
                    objRange = objSheet.get_Range("S12", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox10.Text);


                    objRange = objSheet.get_Range("R2", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox11.Text);
                    objRange = objSheet.get_Range("R3", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, dateTimePicker1.Text);
                    objRange = objSheet.get_Range("R4", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, dateTimePicker2.Text);
                    objRange = objSheet.get_Range("R5", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox12.Text);
                    objRange = objSheet.get_Range("R6", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox23.Text);
                    objRange = objSheet.get_Range("R7", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox45.Text);
                    objRange = objSheet.get_Range("R8", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox34.Text);



                    objRange = objSheet.get_Range("A14", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox22.Text);
                    objRange = objSheet.get_Range("B14", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox21.Text);
                    objRange = objSheet.get_Range("E14", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox20.Text);
                    objRange = objSheet.get_Range("H14", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox19.Text);
                    objRange = objSheet.get_Range("J14", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox18.Text);
                    objRange = objSheet.get_Range("K14", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox17.Text);
                    objRange = objSheet.get_Range("P14", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox16.Text);
                    objRange = objSheet.get_Range("Q14", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox15.Text);
                    objRange = objSheet.get_Range("R14", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox14.Text);
                    objRange = objSheet.get_Range("S14", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox13.Text);
       

                    objRange = objSheet.get_Range("A16", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox33.Text);
                    objRange = objSheet.get_Range("B16", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox32.Text);
                    objRange = objSheet.get_Range("E16", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox31.Text);
                    objRange = objSheet.get_Range("H16", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox30.Text);
                    objRange = objSheet.get_Range("J16", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox29.Text);
                    objRange = objSheet.get_Range("K16", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox28.Text);
                    objRange = objSheet.get_Range("P16", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox27.Text);
                    objRange = objSheet.get_Range("Q16", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox26.Text);
                    objRange = objSheet.get_Range("R16", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox25.Text);
                    objRange = objSheet.get_Range("S16", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox24.Text);


                    objRange = objSheet.get_Range("A18", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox44.Text);
                    objRange = objSheet.get_Range("B18", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox43.Text);
                    objRange = objSheet.get_Range("E18", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox42.Text);
                    objRange = objSheet.get_Range("H18", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox41.Text);
                    objRange = objSheet.get_Range("J18", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox40.Text);
                    objRange = objSheet.get_Range("K18", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox39.Text);
                    objRange = objSheet.get_Range("P18", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox38.Text);
                    objRange = objSheet.get_Range("Q18", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox37.Text);
                    objRange = objSheet.get_Range("R18", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox36.Text);
                    objRange = objSheet.get_Range("S18", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox35.Text);


                    objRange = objSheet.get_Range("A20", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox55.Text);
                    objRange = objSheet.get_Range("B20", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox54.Text);
                    objRange = objSheet.get_Range("E20", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox53.Text);
                    objRange = objSheet.get_Range("H20", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox52.Text);
                    objRange = objSheet.get_Range("J20", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox51.Text);
                    objRange = objSheet.get_Range("K20", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox50.Text);
                    objRange = objSheet.get_Range("P20", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox49.Text);
                    objRange = objSheet.get_Range("Q20", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox48.Text);
                    objRange = objSheet.get_Range("R20", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox47.Text);
                    objRange = objSheet.get_Range("S20", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox46.Text);


                    objRange = objSheet.get_Range("A22", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox66.Text);
                    objRange = objSheet.get_Range("B22", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox65.Text);
                    objRange = objSheet.get_Range("E22", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox64.Text);
                    objRange = objSheet.get_Range("H22", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox63.Text);
                    objRange = objSheet.get_Range("J22", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox62.Text);
                    objRange = objSheet.get_Range("K22", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox61.Text);
                    objRange = objSheet.get_Range("P22", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox60.Text);
                    objRange = objSheet.get_Range("Q22", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox59.Text);
                    objRange = objSheet.get_Range("R22", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox58.Text);
                    objRange = objSheet.get_Range("S22", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox57.Text);


                    objRange = objSheet.get_Range("A24", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox77.Text);
                    objRange = objSheet.get_Range("B24", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox76.Text);
                    objRange = objSheet.get_Range("E24", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox75.Text);
                    objRange = objSheet.get_Range("H24", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox74.Text);
                    objRange = objSheet.get_Range("J24", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox73.Text);
                    objRange = objSheet.get_Range("K24", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox72.Text);
                    objRange = objSheet.get_Range("P24", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox71.Text);
                    objRange = objSheet.get_Range("Q24", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox70.Text);
                    objRange = objSheet.get_Range("R24", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox69.Text);
                    objRange = objSheet.get_Range("S24", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox68.Text);

            
                    
                    objRange = objSheet.get_Range("A26", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox88.Text);
                    objRange = objSheet.get_Range("B26", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox87.Text);
                    objRange = objSheet.get_Range("E26", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox86.Text);
                    objRange = objSheet.get_Range("H26", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox85.Text);
                    objRange = objSheet.get_Range("J26", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox84.Text);
                    objRange = objSheet.get_Range("K26", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox83.Text);
                    objRange = objSheet.get_Range("P26", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox82.Text);
                    objRange = objSheet.get_Range("Q26", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox81.Text);
                    objRange = objSheet.get_Range("R26", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox80.Text);
                    objRange = objSheet.get_Range("S26", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox79.Text);
     

                    objRange = objSheet.get_Range("A28", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox99.Text);
                    objRange = objSheet.get_Range("B28", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox98.Text);
                    objRange = objSheet.get_Range("E28", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox97.Text);
                    objRange = objSheet.get_Range("H28", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox96.Text);
                    objRange = objSheet.get_Range("J28", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox95.Text);
                    objRange = objSheet.get_Range("K28", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox94.Text);
                    objRange = objSheet.get_Range("P28", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox93.Text);
                    objRange = objSheet.get_Range("Q28", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox92.Text);
                    objRange = objSheet.get_Range("R28", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox91.Text);
                    objRange = objSheet.get_Range("S28", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox90.Text);


                    objRange = objSheet.get_Range("a30", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox110.Text);
                    objRange = objSheet.get_Range("b30", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox109.Text);
                    objRange = objSheet.get_Range("E30", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox108.Text);
                    objRange = objSheet.get_Range("H30", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox107.Text);
                    objRange = objSheet.get_Range("J30", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox106.Text);
                    objRange = objSheet.get_Range("K30", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox105.Text);
                    objRange = objSheet.get_Range("P30", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox104.Text);
                    objRange = objSheet.get_Range("Q30", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox103.Text);
                    objRange = objSheet.get_Range("R30", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox102.Text);
                    objRange = objSheet.get_Range("S30", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox101.Text);


                    objRange = objSheet.get_Range("a32", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox121.Text);
                    objRange = objSheet.get_Range("b32", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox120.Text);
                    objRange = objSheet.get_Range("E32", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox119.Text);
                    objRange = objSheet.get_Range("H32", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox118.Text);
                    objRange = objSheet.get_Range("J32", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox117.Text);
                    objRange = objSheet.get_Range("K32", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox116.Text);
                    objRange = objSheet.get_Range("P32", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox115.Text);
                    objRange = objSheet.get_Range("Q32", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox114.Text);
                    objRange = objSheet.get_Range("R32", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox113.Text);
                    objRange = objSheet.get_Range("S32", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox112.Text);


                    objRange = objSheet.get_Range("a34", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox132.Text);
                    objRange = objSheet.get_Range("b34", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox131.Text);
                    objRange = objSheet.get_Range("E34", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox130.Text);
                    objRange = objSheet.get_Range("H34", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox129.Text);
                    objRange = objSheet.get_Range("J34", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox128.Text);
                    objRange = objSheet.get_Range("K34", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox127.Text);
                    objRange = objSheet.get_Range("P34", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox126.Text);
                    objRange = objSheet.get_Range("Q34", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox125.Text);
                    objRange = objSheet.get_Range("R34", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox124.Text);
                    objRange = objSheet.get_Range("S34", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox123.Text);
    

                    objRange = objSheet.get_Range("a36", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox143.Text);
                    objRange = objSheet.get_Range("b36", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox142.Text);
                    objRange = objSheet.get_Range("E36", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox141.Text);
                    objRange = objSheet.get_Range("H36", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox140.Text);
                    objRange = objSheet.get_Range("J36", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox139.Text);
                    objRange = objSheet.get_Range("K36", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox138.Text);
                    objRange = objSheet.get_Range("P36", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox137.Text);
                    objRange = objSheet.get_Range("Q36", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox136.Text);
                    objRange = objSheet.get_Range("R36", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox135.Text);
                    objRange = objSheet.get_Range("S36", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox134.Text);
       

                    objRange = objSheet.get_Range("a38", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox154.Text);
                    objRange = objSheet.get_Range("b38", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox153.Text);
                    objRange = objSheet.get_Range("E38", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox152.Text);
                    objRange = objSheet.get_Range("H38", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox151.Text);
                    objRange = objSheet.get_Range("J38", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox150.Text);
                    objRange = objSheet.get_Range("K38", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox149.Text);
                    objRange = objSheet.get_Range("P38", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox148.Text);
                    objRange = objSheet.get_Range("Q38", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox147.Text);
                    objRange = objSheet.get_Range("R38", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox146.Text);
                    objRange = objSheet.get_Range("S38", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox145.Text);
                   

                    objRange = objSheet.get_Range("a40", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox165.Text);
                    objRange = objSheet.get_Range("b40", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox164.Text);
                    objRange = objSheet.get_Range("E40", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox163.Text);
                    objRange = objSheet.get_Range("H40", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox162.Text);
                    objRange = objSheet.get_Range("J40", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox161.Text);
                    objRange = objSheet.get_Range("K40", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox160.Text);
                    objRange = objSheet.get_Range("P40", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox159.Text);
                    objRange = objSheet.get_Range("Q40", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox158.Text);
                    objRange = objSheet.get_Range("R40", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox157.Text);
                    objRange = objSheet.get_Range("S40", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox156.Text);
         

                    objRange = objSheet.get_Range("a42", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox176.Text);
                    objRange = objSheet.get_Range("b42", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox175.Text);
                    objRange = objSheet.get_Range("E42", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox174.Text);
                    objRange = objSheet.get_Range("H42", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox173.Text);
                    objRange = objSheet.get_Range("J42", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox172.Text);
                    objRange = objSheet.get_Range("K42", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox171.Text);
                    objRange = objSheet.get_Range("P42", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox170.Text);
                    objRange = objSheet.get_Range("Q42", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox169.Text);
                    objRange = objSheet.get_Range("R42", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox168.Text);
                    objRange = objSheet.get_Range("S42", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox167.Text);
              

                    objRange = objSheet.get_Range("a44", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox187.Text);
                    objRange = objSheet.get_Range("b44", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox186.Text);
                    objRange = objSheet.get_Range("E44", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox185.Text);
                    objRange = objSheet.get_Range("H44", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox184.Text);
                    objRange = objSheet.get_Range("J44", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox183.Text);
                    objRange = objSheet.get_Range("K44", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox182.Text);
                    objRange = objSheet.get_Range("P44", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox181.Text);
                    objRange = objSheet.get_Range("Q44", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox180.Text);
                    objRange = objSheet.get_Range("R44", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox179.Text);
                    objRange = objSheet.get_Range("S44", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox178.Text);

                    objRange = objSheet.get_Range("a46", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox198.Text);
                    objRange = objSheet.get_Range("b46", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox197.Text);
                    objRange = objSheet.get_Range("E46", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox196.Text);
                    objRange = objSheet.get_Range("H46", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox195.Text);
                    objRange = objSheet.get_Range("J46", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox194.Text);
                    objRange = objSheet.get_Range("K46", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox193.Text);
                    objRange = objSheet.get_Range("P46", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox192.Text);
                    objRange = objSheet.get_Range("Q46", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox191.Text);
                    objRange = objSheet.get_Range("R46", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox190.Text);
                    objRange = objSheet.get_Range("S46", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox189.Text);
        

                    objRange = objSheet.get_Range("a48", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox209.Text);
                    objRange = objSheet.get_Range("b48", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox208.Text);
                    objRange = objSheet.get_Range("E48", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox207.Text);
                    objRange = objSheet.get_Range("H48", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox206.Text);
                    objRange = objSheet.get_Range("J48", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox205.Text);
                    objRange = objSheet.get_Range("K48", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox204.Text);
                    objRange = objSheet.get_Range("P48", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox203.Text);
                    objRange = objSheet.get_Range("Q48", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox202.Text);
                    objRange = objSheet.get_Range("R48", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox201.Text);
                    objRange = objSheet.get_Range("S48", System.Reflection.Missing.Value);
                    objRange.set_Value(System.Reflection.Missing.Value, textBox200.Text);
         

                    

                    objSheet.PrintOutEx(1, 1, 2, true);

                    Form anasayfa = new anasayfa();
                    anasayfa.Show();
                    this.Hide();
                }
                catch (Exception)
                {


                }
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Form anasayfa = new anasayfa();
            anasayfa.Show();
            this.Hide();
        }

        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
           
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            irsaliyeNo = textBox1.Text;
            bilgiGetir(irsaliyeNo, textBox2.Text, textBox3.Text, textBox4.Text, textBox10.Text);

            textBox2.Text = gonderenText;
            textBox3.Text = aliciText;
            textBox4.Text = yerText;
            textBox10.Text =  ucretText;
            
          
        }

        private void textBox22_SelectedIndexChanged(object sender, EventArgs e)
        {
            irsaliyeNo = textBox22.Text;
            bilgiGetir(irsaliyeNo, textBox21.Text, textBox20.Text, textBox19.Text, textBox13.Text);

            textBox21.Text = gonderenText;
            textBox20.Text = aliciText;
            textBox19.Text = yerText;
            textBox13.Text = ucretText;
        }

        private void textBox33_SelectedIndexChanged(object sender, EventArgs e)
        {
            irsaliyeNo = textBox33.Text;
            bilgiGetir(irsaliyeNo, textBox32.Text, textBox31.Text, textBox29.Text, textBox24.Text);

            textBox32.Text = gonderenText;
            textBox31.Text = aliciText;
            textBox30.Text = yerText;
            textBox24.Text = ucretText;
        }

        private void textBox44_SelectedIndexChanged(object sender, EventArgs e)
        {
            irsaliyeNo = textBox44.Text;
            bilgiGetir(irsaliyeNo, textBox43.Text, textBox42.Text, textBox41.Text, textBox35.Text);

            textBox43.Text = gonderenText;
            textBox42.Text = aliciText;
            textBox41.Text = yerText;
            textBox35.Text = ucretText;
        }

        private void textBox55_SelectedIndexChanged(object sender, EventArgs e)
        {
            irsaliyeNo = textBox55.Text;
            bilgiGetir(irsaliyeNo, textBox54.Text, textBox53.Text, textBox52.Text, textBox46.Text);

            textBox54.Text = gonderenText;
            textBox53.Text = aliciText;
            textBox52.Text = yerText;
            textBox46.Text = ucretText;
        }

        private void textBox66_SelectedIndexChanged(object sender, EventArgs e)
        {
            irsaliyeNo = textBox66.Text;
            bilgiGetir(irsaliyeNo, textBox65.Text, textBox64.Text, textBox63.Text, textBox57.Text);

            textBox65.Text = gonderenText;
            textBox64.Text = aliciText;
            textBox63.Text = yerText;
            textBox57.Text = ucretText;
        }

        private void textBox77_SelectedIndexChanged(object sender, EventArgs e)
        {
            irsaliyeNo = textBox77.Text;
            bilgiGetir(irsaliyeNo, textBox76.Text, textBox75.Text, textBox74.Text, textBox68.Text);

            textBox76.Text = gonderenText;
            textBox75.Text = aliciText;
            textBox74.Text = yerText;
            textBox68.Text = ucretText;
        }

        private void textBox88_SelectedIndexChanged(object sender, EventArgs e)
        {
            irsaliyeNo = textBox88.Text;
            bilgiGetir(irsaliyeNo, textBox87.Text, textBox86.Text, textBox85.Text, textBox79.Text);

            textBox87.Text = gonderenText;
            textBox86.Text = aliciText;
            textBox85.Text = yerText;
            textBox79.Text = ucretText;
        }

        private void textBox99_SelectedIndexChanged(object sender, EventArgs e)
        {
            irsaliyeNo = textBox99.Text;
            bilgiGetir(irsaliyeNo, textBox98.Text, textBox97.Text, textBox96.Text, textBox90.Text);

            textBox98.Text = gonderenText;
            textBox97.Text = aliciText;
            textBox96.Text = yerText;
            textBox90.Text = ucretText;
        }

        private void textBox110_SelectedIndexChanged(object sender, EventArgs e)
        {
            irsaliyeNo = textBox110.Text;
            bilgiGetir(irsaliyeNo, textBox109.Text, textBox108.Text, textBox107.Text, textBox101.Text);

            textBox109.Text = gonderenText;
            textBox108.Text = aliciText;
            textBox107.Text = yerText;
            textBox101.Text = ucretText;
        }

        private void textBox121_SelectedIndexChanged(object sender, EventArgs e)
        {
            irsaliyeNo = textBox121.Text;
            bilgiGetir(irsaliyeNo, textBox120.Text, textBox119.Text, textBox118.Text, textBox112.Text);

            textBox120.Text = gonderenText;
            textBox119.Text = aliciText;
            textBox118.Text = yerText;
            textBox112.Text = ucretText;
        }

        private void textBox132_SelectedIndexChanged(object sender, EventArgs e)
        {
            irsaliyeNo = textBox132.Text;
            bilgiGetir(irsaliyeNo, textBox131.Text, textBox130.Text, textBox129.Text, textBox123.Text);

            textBox131.Text = gonderenText;
            textBox130.Text = aliciText;
            textBox129.Text = yerText;
            textBox123.Text = ucretText;
        }

        private void textBox143_SelectedIndexChanged(object sender, EventArgs e)
        {
            irsaliyeNo = textBox143.Text;
            bilgiGetir(irsaliyeNo, textBox142.Text, textBox141.Text, textBox140.Text, textBox134.Text);

            textBox142.Text = gonderenText;
            textBox141.Text = aliciText;
            textBox140.Text = yerText;
            textBox134.Text = ucretText;
        }

        private void textBox154_SelectedIndexChanged(object sender, EventArgs e)
        {
            irsaliyeNo = textBox154.Text;
            bilgiGetir(irsaliyeNo, textBox153.Text, textBox152.Text, textBox151.Text, textBox145.Text);

            textBox153.Text = gonderenText;
            textBox152.Text = aliciText;
            textBox151.Text = yerText;
            textBox145.Text = ucretText;
        }

        private void textBox165_SelectedIndexChanged(object sender, EventArgs e)
        {
            irsaliyeNo = textBox165.Text;
            bilgiGetir(irsaliyeNo, textBox164.Text, textBox163.Text, textBox162.Text, textBox156.Text);

            textBox164.Text = gonderenText;
            textBox163.Text = aliciText;
            textBox162.Text = yerText;
            textBox156.Text = ucretText;
        }

        private void textBox176_SelectedIndexChanged(object sender, EventArgs e)
        {
            irsaliyeNo = textBox176.Text;
            bilgiGetir(irsaliyeNo, textBox175.Text, textBox174.Text, textBox173.Text, textBox167.Text);

            textBox175.Text = gonderenText;
            textBox174.Text = aliciText;
            textBox173.Text = yerText;
            textBox167.Text = ucretText;
        }

        private void textBox187_SelectedIndexChanged(object sender, EventArgs e)
        {
            irsaliyeNo = textBox187.Text;
            bilgiGetir(irsaliyeNo, textBox186.Text, textBox185.Text, textBox184.Text, textBox178.Text);

            textBox186.Text = gonderenText;
            textBox185.Text = aliciText;
            textBox184.Text = yerText;
            textBox178.Text = ucretText;
        }

        private void textBox198_SelectedIndexChanged(object sender, EventArgs e)
        {
            irsaliyeNo = textBox198.Text;
            bilgiGetir(irsaliyeNo, textBox197.Text, textBox196.Text, textBox195.Text, textBox189.Text);

            textBox197.Text = gonderenText;
            textBox196.Text = aliciText;
            textBox195.Text = yerText;
            textBox189.Text = ucretText;
        }

        private void textBox209_SelectedIndexChanged(object sender, EventArgs e)
        {
            irsaliyeNo = textBox209.Text;
            bilgiGetir(irsaliyeNo, textBox208.Text, textBox207.Text, textBox206.Text, textBox200.Text);

            textBox208.Text = gonderenText;
            textBox207.Text = aliciText;
            textBox206.Text = yerText;
            textBox200.Text = ucretText;
        }

        

    }
}
