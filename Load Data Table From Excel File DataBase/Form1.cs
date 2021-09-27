using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.IO;

namespace Load_Data_Table_From_Excel_File_DataBase
{
    public partial class Form1 : Form
    {
        SqlConnection SqlConn = new SqlConnection();
        SqlCommand SqlCmd = new SqlCommand();
        SqlDataAdapter SqlDatA = new SqlDataAdapter();        
        OleDbConnection OlDbConn = new OleDbConnection();
        OleDbDataAdapter OlDbDatA = new OleDbDataAdapter();
        DataTable Dt = new DataTable();
    
        public Form1()
        {
            InitializeComponent();
        }
        private void button1_Click(object sender, EventArgs e)
        {
            int coun = 0;
            try
            {            
            SqlConn = new SqlConnection("Server = DESKTOP-7O52JPG ; DataBase = My_Database ; Integrated Security = true ");
            OlDbConn = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source =E:\Employees.xlsx; Extended Properties= 'Excel 12.0 Xml;HDR=YES;IMEX=1;MAXSCANROWS=0'");
            OlDbDatA = new OleDbDataAdapter("select * from [Sheet1$]", OlDbConn);            
            SqlConn.Open();
            OlDbDatA.Fill(Dt);
            for (int i = 0; i < Dt.Rows.Count; i++)
            {
                try
                {
                    SqlCmd = new SqlCommand("insert into Employees values('" + Dt.Rows[i][0] + "','" + Dt.Rows[i][1] + "','" + Dt.Rows[i][2] + "','" + Dt.Rows[i][3] + "','" + Dt.Rows[i][4] + "','" + Dt.Rows[i][5] + "','" + Dt.Rows[i][6] + "','" + Dt.Rows[i][7] + "')", SqlConn);
                    SqlCmd.ExecuteNonQuery();
                    coun += 1;
                }
                catch (Exception )
                {
                    continue;
                }
            }
                if (coun == Dt.Rows.Count)
                {
                    lbl_error.Text = " Sql تم التحميل بنجاح الى \n" + "      عدد الصفوف : " + coun;
                }
                else
                {
                    lbl_error.Text = " ! خطأ فى التحميل "; 
                }                                                          
            }
            catch (SqlException ex)
            {
                MessageBox.Show("Error : " + ex.Message); 
            }
            finally
            {
                SqlConn.Close(); 
            }                     
        }
        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                OleDbConnection OlDbCon = new OleDbConnection();
                SqlConn = new SqlConnection("Server = DESKTOP-7O52JPG ; DataBase = My_Database ; Integrated Security = true ");
                OlDbCon = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source =E:\\Data.xlsx; Extended Properties= 'Excel 12.0 Xml;HDR=YES;'");   //////// YES تم حذف كل ما هو بعد 
                SqlDatA = new SqlDataAdapter("select * from Employees", SqlConn);
                Dt = new DataTable();
                SqlConn.Open();
                OlDbCon.Open();
                SqlDatA.Fill(Dt);
                int cont = 0;
                string nameFile = ""+ txt_lod.Text +"" + DateTime.Now.ToString("dd-MM-yyyy");
                StreamWriter file = new StreamWriter(@"E:\"+nameFile+".txt");
                for (int i = 0; i < Dt.Rows.Count; i++)
                {
                    try
                    {
                        OleDbCommand ocmd = new OleDbCommand("insert into [Sheet1$]  values(" + Dt.Rows[i][0] + ",'" + Dt.Rows[i][1] + "','" + Dt.Rows[i][2] + "','" + Dt.Rows[i][3] + "','" + Dt.Rows[i][4] + "','" + Dt.Rows[i][5] + "','" + Dt.Rows[i][6] + "'," + Dt.Rows[i][7] + ")", OlDbCon);
                        ocmd.ExecuteNonQuery();
                        cont += 1;
                        file.WriteLine(Dt.Rows[i][0].ToString() + "|" + Dt.Rows[i][1].ToString() + "|" + Dt.Rows[i][2].ToString() + "|" + Dt.Rows[i][3].ToString() + "|" + Dt.Rows[i][4].ToString() + "|" + Dt.Rows[i][5].ToString() + "|" + Dt.Rows[i][6].ToString() + "|" + Dt.Rows[i][7].ToString());                    
                    }
                    catch (SqlException ex)
                    {
                        MessageBox.Show("Error : " + ex.Message);
                    }
                }
                if (cont == Dt.Rows.Count)
                {
                    lbl_error.Text = "text و Excel تم التحميل بنجاح الى \n" + "      عدد الصفوف : " + cont;
                }
                else
                {
                    lbl_error.Text = " ! خطأ فى التحميل ";
                }
                file.Close();
            }
            catch (SqlException ex)
            {
                MessageBox.Show("Error : " + ex.Message);
            }
            finally
            {
                SqlConn.Close();
                OlDbConn.Close();
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
   
            SqlConn = new SqlConnection("Server = DESKTOP-7O52JPG ; DataBase = My_Database ; Integrated Security = true ");
             DataTable dt = new DataTable();
            string path = "E:\\nader24-09-2021.txt";
             dt = CreateDataTableFromFile(path,25,1);
            for (int i = 0; i < Convert.ToInt32(dt.Rows.Count); i++)
            {
                decimal sal;
                int id;
                string name,ad,ph,dat,dep,job;
                id =Convert.ToInt32(dt.Rows[i][0]);
                name = dt.Rows[i][1].ToString();
                ad = dt.Rows[i][2].ToString();
                ph = dt.Rows[i][3].ToString();
                dat = dt.Rows[i][4].ToString();
                dep = dt.Rows[i][5].ToString();
                job = dt.Rows[i][6].ToString();
                sal = Convert.ToDecimal( dt.Rows[i][7]);
                SqlConn.Open();
                SqlCmd = new SqlCommand("insert into Employees values ( '"+id+"','"+name+"','"+ad+"','"+ph+"','"+dat+"','"+dep+"','"+job+"','"+sal+"' )", SqlConn);
                SqlCmd.ExecuteNonQuery();
                SqlConn.Close();     
            }
        }

  //-------------------------------------------------------------------------------------------
        private DataTable CreateDataTableFromFile(string path, int count, int file_extension)
        {

            int lines;
            DataTable dt = new DataTable();
            DataColumn dc;
            DataRow dr;
            if (file_extension == 1)
            {
                lines = count;
                for (int i = 1; i <= lines; i++)
                {
                    dc = new DataColumn();
                    dc.DataType = System.Type.GetType("System.String");
                    dc.ColumnName = string.Format("f{0}", i);
                    dc.Unique = false;
                    dt.Columns.Add(dc);
                }

                StreamReader sr = new StreamReader(@path, Encoding.Default);
                string input;
                while ((input = sr.ReadLine()) != null)
                {
                    string[] s = input.Split(new string[] { "||" }, StringSplitOptions.None);
                    //string[] s = input.Split(new char[] { '|' });
                    dr = dt.NewRow();

                    for (int i = 1; i <= lines; i++)
                    {
                        string col = string.Format("f{0}", i);
                        try
                        {
                            dr[col] = s[i - 1];

                        }
                        catch (Exception e) { }

                    }
                    dt.Rows.Add(dr);

                }
                sr.Close();
            }
            else
            {
                OleDbConnection oconn = new OleDbConnection(string.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + @path + ";Extended Properties=\"Excel 12.0 Xml;HDR=YES;IMEX=0\";"));

                OleDbDataAdapter da = new OleDbDataAdapter("select * from [sheet1$]", oconn);
                da.Fill(dt);
            }
            return dt;
        }
       
    }
}
