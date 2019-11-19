using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using MySql.Data.MySqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Data.SQLite;

namespace Data_base
{
    public partial class Form1 : Form
    {
        public List<RoomType> roomTypes { get; set; }

        public Form1()
        {
            InitializeComponent();
            InsertData();
            roomTypes = new List<RoomType>();
            MySqlConnection connection = new MySqlConnection(myConnectionString);
            connection.Open();
            try
            {
                //在这里使用代码对数据库进行增删查改
                string sql1 = "select * from roomtype_table";
                MySqlCommand cmd = new MySqlCommand(sql1, connection);
                MySqlDataReader reader = cmd.ExecuteReader();
                if (reader.HasRows)
                {
                    //读取数据
                    //如果读取到数据返回true，否则false
                    while (reader.Read())
                    {
                        //在数据集合加入数据，
                        roomTypes.Add(
                        new RoomType(reader["room_type"].ToString(),
                            reader["room_name"].ToString(),
                            reader["room_spe"].ToString(),
                            reader["room_price"].ToString()));
                    }
                }
                bindingSource1.DataSource = roomTypes;
                bindingSource1.ResetBindings(true);
                dataGridView1.Columns["room_type"].DataPropertyName = "room_type";
                dataGridView1.Columns["room_name"].DataPropertyName = "room_name";
                dataGridView1.Columns["room_spe"].DataPropertyName = "room_spe";
                dataGridView1.Columns["room_price"].DataPropertyName = "room_price";



            }
            catch (MySqlException ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {
                connection.Close();

            }
            SQLiteCommand command = GetSQLiteCommand();
            command.CommandText = "select * from tbl_fgwj";
            SQLiteDataAdapter da = new SQLiteDataAdapter();
            DataSet dt = new DataSet();
            da.SelectCommand = command;
            da.Fill(dt);
            bindingSource3.DataSource = dt.Tables[0];
            dataGridView3.Columns["Column6"].DataPropertyName = "FILE_NO";
            dataGridView3.Columns["Column7"].DataPropertyName = "FILE_NAME";
            dataGridView3.Columns["Column8"].DataPropertyName = "SUBJECT";
            dataGridView3.Columns["Column9"].DataPropertyName = "PUBLISH_DATE";
            dataGridView3.Columns["Column10"].DataPropertyName = "IMPLEMENT_DATE";



        }
        public static SQLiteCommand GetSQLiteCommand()
        {
            string connectionString = @"data source=E:\Course\windowsProgramDesign\H4_6\WindowsFormsApp1\Data_base\demo.db";
            SQLiteConnection con = new SQLiteConnection(connectionString);
            con.Open();
            return new SQLiteCommand(con);

        }
        private string myConnectionString = "server=localhost;User Id=root;password=blaveer;Database=hotel";


        public DataSet ExcelToDS(string path, string no)
        {
            DataSet ds = null;
            try
            {
                string strConn = "Provider=Microsoft.Jet.OLEDB.4.0;"
                    + "Data Source=" + path + ";" + "Extended Properties=Excel 8.0;";
                OleDbConnection conn = new OleDbConnection(strConn);
                conn.Open();
                string strExcel = "";
                OleDbDataAdapter myCommand = null;
                strExcel = "select * from [sheet" + no + "$]";
                myCommand = new OleDbDataAdapter(strExcel, strConn);
                DataTable table1 = new DataTable();
                ds = new DataSet();
                myCommand.Fill(table1);
                myCommand.Fill(ds);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            return ds;
        }

        private void InsertData()
        {
            DataTable dataTable = ExcelToDS(@"E:\Course\windowsProgramDesign\H4_6\WindowsFormsApp1\Data_base\hotel.xls", "1").Tables[0];
            foreach (DataRow r in dataTable.Rows)
            {
                string type = r["房型编号"].ToString();
                string name = r["房型名称"].ToString();
                string special = r["房型特征"].ToString();
                string price = r["单价"].ToString();
                insertMysqlDB(type, name, special, price);//更新数据到数据库
            }



        }

        private void insertMysqlDB(string type, string name, string special, string price)
        {
            MySqlConnection connection = new MySqlConnection(myConnectionString);
            connection.Open();

            try
            {

                string sql1 = "insert into roomtype_table set room_type='" + type + "',room_name='" + name + "',room_spe='" + special + "',room_price='" + price + "'";
                MySqlCommand cmd = new MySqlCommand(sql1, connection);
                cmd.ExecuteNonQuery();
            }
            catch (MySqlException ex)
            {
                Console.WriteLine(ex.Message);

            }
            finally
            {

                connection.Close();

            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            bindingSource2.DataSource = ExcelToDS(@"E:\Course\windowsProgramDesign\H4_6\WindowsFormsApp1\Data_base\hotel.xls", "2").Tables[0];
            dataGridView2.Columns["Column1"].DataPropertyName = "房型编号";
            dataGridView2.Columns["Column2"].DataPropertyName = "房号";
            dataGridView2.Columns["Column3"].DataPropertyName = "所在地址";
            dataGridView2.Columns["Column4"].DataPropertyName = "房型特征";
            dataGridView2.Columns["Column5"].DataPropertyName = "单价";


        }

        private void button2_Click(object sender, EventArgs e)
        {
            New_one n1 = new New_one(bindingSource3);
            n1.Show();


        }

        private void button4_Click(object sender, EventArgs e)
        {
            int rowNum = dataGridView3.CurrentRow.Index;
            SQLiteCommand command = GetSQLiteCommand();
            command.CommandText = "delete from tbl_fgwj where FILE_NO =@FILE_NO";
            command.Parameters.Add("FILE_NO", DbType.String).Value = dataGridView3.Rows[rowNum].Cells[0].Value;
            command.ExecuteNonQuery();
            MessageBox.Show("删除成功");
            command.CommandText = "select * from tbl_fgwj";
            SQLiteDataAdapter da = new SQLiteDataAdapter();
            DataSet dt = new DataSet();
            da.SelectCommand = command;
            da.Fill(dt);
            bindingSource3.DataSource = dt.Tables[0];
            bindingSource3.ResetBindings(true);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            int rowNum = dataGridView3.CurrentRow.Index;
            SQLiteCommand command = GetSQLiteCommand();
            command.CommandText = "select * from tbl_fgwj";
            SQLiteDataAdapter da = new SQLiteDataAdapter();
            DataSet dt = new DataSet();
            da.SelectCommand = command;
            da.Fill(dt);
            DataTable table = dt.Tables[0];
            string key = table.Rows[rowNum].ItemArray[0].ToString();
            command.CommandText = "UPDATE tbl_fgwj set FILE_NO=@FILE_NO,SUBJECT=@SUBJECT,PUBLISH_DATE=@PUBLISH_DATE,IMPLEMENT_DATE=@IMPLEMENT_DATE,FILE_NAME=@FILE_NAME where ID_KEY =@ID_KEY";
            command.Parameters.Add("FILE_NO", DbType.String).Value = dataGridView3.Rows[rowNum].Cells[0].Value;
            command.Parameters.Add("SUBJECT", DbType.String).Value = dataGridView3.Rows[rowNum].Cells[2].Value;
            command.Parameters.Add("PUBLISH_DATE", DbType.String).Value = dataGridView3.Rows[rowNum].Cells[3].Value;
            command.Parameters.Add("IMPLEMENT_DATE", DbType.String).Value = dataGridView3.Rows[rowNum].Cells[4].Value;
            command.Parameters.Add("ID_KEY", DbType.String).Value = key;
            command.Parameters.Add("FILE_NAME", DbType.String).Value = dataGridView3.Rows[rowNum].Cells[1].Value;
            command.ExecuteNonQuery();
            MessageBox.Show("更新成功");
            command.CommandText = "select * from tbl_fgwj";
            da = new SQLiteDataAdapter();
            dt = new DataSet();
            da.SelectCommand = command;
            da.Fill(dt);
            bindingSource3.DataSource = dt.Tables[0];
            bindingSource3.ResetBindings(true);
        }

        private void button5_Click(object sender, EventArgs e)
        {
            MySqlConnection connection = new MySqlConnection("server=localhost;User Id=root;password=blaveer;Database=hotel");
            connection.Open();
            DataTable DBNameTable = new DataTable();


            MySqlDataAdapter Adapter = new MySqlDataAdapter("show databases;", connection);
            lock (Adapter)
            {
                Adapter.Fill(DBNameTable);
            }

            foreach (DataRow row in DBNameTable.Rows)
            {

                comboBox1.Items.Insert(0, row.ItemArray[0]);
            }

        }

        private void button6_Click(object sender, EventArgs e)
        {
            try
            {
                string sqlCon = "server=localhost;User Id=" + textBox1.Text + ";password=" + textBox2.Text + ";Database=" + comboBox1.SelectedText;
                MySqlConnection connection = new MySqlConnection(sqlCon);
                connection.Open();
                MessageBox.Show("连接成功");
            }
            catch (Exception e1)
            {
                MessageBox.Show("连接失败！\r\n" + e1.Message);
            }

        }

        private void button7_Click(object sender, EventArgs e)
        {
            string sqlCon = "server=localhost;User Id=" + textBox1.Text + ";password=" + textBox2.Text + ";Database=" + comboBox1.Text;
            MySqlConnection connection = new MySqlConnection(sqlCon);
            connection.Open();
            string sql1 = "select * from books";
            MySqlCommand cmd = new MySqlCommand(sql1, connection);
            MySqlDataReader reader = cmd.ExecuteReader();
            List<book> books = new List<book>();
            try
            {
                if (reader.HasRows)
                {
                    //读取数据
                    //如果读取到数据返回true，否则false
                    while (reader.Read())
                    {
                        //在数据集合加入数据，
                        books.Add(
                        new book(int.Parse(reader["id"].ToString()),
                            reader["name"].ToString(),
                            reader["authorFirstName"].ToString(),
                            reader["authorLastName"].ToString(),
                            reader["press"].ToString(),
                            double.Parse(reader["price"].ToString())
                            ));
                    }
                }
            }catch(Exception ex)
            {
                MessageBox.Show(ex.ToString());
                return;
            }
           
            bindingSource4.DataSource = books;
            bindingSource4.ResetBindings(true);
            dataGridView4.Columns["Column11"].DataPropertyName = "id";
            dataGridView4.Columns["Column12"].DataPropertyName = "name";
            dataGridView4.Columns["Column13"].DataPropertyName = "authorFirstName";
            dataGridView4.Columns["Column14"].DataPropertyName = "authroLastName";
            dataGridView4.Columns["Column15"].DataPropertyName = "press";
            dataGridView4.Columns["Column16"].DataPropertyName = "price";

        }

        private void button8_Click(object sender, EventArgs e)
        {
            string sqlCon = "server=localhost;User Id=" + textBox1.Text + ";password=" + textBox2.Text + ";Database=" + comboBox1.Text;
            MySqlConnection connection = new MySqlConnection(sqlCon);
            connection.Open();
            string sql1 = "select * from books";
            MySqlDataAdapter da = new MySqlDataAdapter(sql1, connection);
            DataSet dt = new DataSet();
            da.Fill(dt);
            bindingSource4.DataSource = dt.Tables[0];
            dataGridView4.Columns["Column11"].DataPropertyName = "id";
            dataGridView4.Columns["Column12"].DataPropertyName = "name";
            dataGridView4.Columns["Column13"].DataPropertyName = "authorFirstName";
            dataGridView4.Columns["Column14"].DataPropertyName = "authroLastName";
            dataGridView4.Columns["Column15"].DataPropertyName = "press";
            dataGridView4.Columns["Column16"].DataPropertyName = "price";
        }

        private void button9_Click(object sender, EventArgs e)
        {
            string sqlCon = "server=localhost;User Id=" + textBox1.Text + ";password=" + textBox2.Text + ";Database=" + comboBox1.Text;

            MultiUpateData((DataTable)bindingSource4.DataSource, sqlCon);
        }
        public static int MultiUpateData(DataTable dtInfor, string sqlCon)
        {

            if (dtInfor.Rows.Count == 0)
            {
                return -1;
            }

            string sqlStr = "SELECT * FROM books";
            using (MySqlConnection con = new MySqlConnection(sqlCon))
            {
                using (MySqlCommand cmd = new MySqlCommand(sqlStr, con))
                {
                    con.Open();
                    MySqlTransaction transction = con.BeginTransaction(IsolationLevel.ReadCommitted);
                    try
                    {
                        int count = 0;
                        MySqlDataAdapter dataAdapter = new MySqlDataAdapter(cmd);
                        dataAdapter.SelectCommand = new MySqlCommand(sqlStr, con);
                        MySqlCommandBuilder builder = new MySqlCommandBuilder(dataAdapter);
                        builder.ConflictOption = ConflictOption.OverwriteChanges;
                        builder.SetAllValues = true;
                        count = dataAdapter.Update(dtInfor);
                        transction.Commit();
                        dtInfor.AcceptChanges();
                        dataAdapter.Dispose();
                        builder.Dispose();
                        MessageBox.Show("批量更新成功");
                        return count;
                    }
                    catch (Exception)
                    {
                        transction.Rollback();
                        throw;
                    }
                }
            }
        }

        private void button10_Click(object sender, EventArgs e)
        {
            try
            {
                string sqlCon = "server=localhost;User Id=" + textBox1.Text + ";password=" + textBox2.Text + ";Database=" + comboBox1.Text;

                MultiUpateData((DataTable)bindingSource4.DataSource, sqlCon);
            } catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
          
        }

        private void button11_Click(object sender, EventArgs e)
        {
            int rowNum = dataGridView4.CurrentRow.Index;
            string sqlCon = "server=localhost;User Id=" + textBox1.Text + ";password=" + textBox2.Text + ";Database=" + comboBox1.Text;
            MySqlConnection connection = new MySqlConnection(sqlCon);
            connection.Open();
            MySqlCommand command = new MySqlCommand("delete from books where id =@id", connection);
            command.CommandText = "delete from books where id =@id";
            command.Parameters.Add("id", MySqlDbType.Int32).Value = dataGridView4.Rows[rowNum].Cells[0].Value;
            command.ExecuteNonQuery();
            MessageBox.Show("删除成功");

            MySqlDataAdapter da = new MySqlDataAdapter("select * from books", connection);
            DataSet dt = new DataSet();

            da.Fill(dt);
            bindingSource4.DataSource = dt.Tables[0];
            bindingSource4.ResetBindings(true);
        }

        private void button12_Click(object sender, EventArgs e)
        {
            string sqlCon = "server=localhost;User Id=" + textBox1.Text + ";password=" + textBox2.Text + ";Database=" + comboBox1.Text;

            MultiUpateData((DataTable)bindingSource4.DataSource, sqlCon);
        }

        private void bindingSource4_CurrentChanged(object sender, EventArgs e)
        {

        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }
    }
}
