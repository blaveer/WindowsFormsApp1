using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SQLite;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Data_base
{
    public partial class New_one : Form
    {
        public BindingSource bindingSource3;
        public string fileType = "";
        public New_one(BindingSource bindingSource3)
        {
            InitializeComponent();
            this.bindingSource3 = bindingSource3;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog fileDialog = new OpenFileDialog();
            fileDialog.Multiselect = true;
            fileDialog.Title = "请选择文件";
            fileDialog.Filter = "所有文件(*.*)|*.*";
            if (fileDialog.ShowDialog() == DialogResult.OK)
            {
                string file = fileDialog.SafeFileName;
                fileNameBox.Text = file.Split('.')[0];
                fileType = file.Split('.')[1];
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (fileType == "")
            {
                MessageBox.Show("尚未选择文件");
                return;
            }
            SQLiteCommand command = Form1.GetSQLiteCommand();
            string insertSql = "insert into tbl_fgwj(ID_KEY,FILE_NO,SUBJECT,PUBLISH_DATE,IMPLEMENT_DATE,PUBLISH_ORG,FILE_NAME,FILE_TYPE) values(@ID_KEY,@FILE_NO,@SUBJECT,@PUBLISH_DATE,@IMPLEMENT_DATE,@PUBLISH_ORG,@FILE_NAME,@FILE_TYPE)";
            command.Parameters.Add("ID_KEY", DbType.String).Value = System.Guid.NewGuid().ToString("N");
            command.Parameters.Add("FILE_NO", DbType.String).Value = fileNoBox.Text;
            command.Parameters.Add("SUBJECT", DbType.String).Value = subjectBox.Text;
            command.Parameters.Add("PUBLISH_DATE", DbType.String).Value = publishBox.Text;
            command.Parameters.Add("IMPLEMENT_DATE", DbType.String).Value = implementBox.Text;
            command.Parameters.Add("PUBLISH_ORG", DbType.String).Value = orgBox.Text;
            command.Parameters.Add("FILE_NAME", DbType.String).Value = fileNameBox.Text;
            command.Parameters.Add("FILE_TYPE", DbType.String).Value = fileType;



            command.CommandText = insertSql;
            command.ExecuteNonQuery();
            MessageBox.Show("插入成功");

            command.CommandText = "select * from tbl_fgwj";
            SQLiteDataAdapter da = new SQLiteDataAdapter();
            DataSet dt = new DataSet();
            da.SelectCommand = command;
            da.Fill(dt);
            bindingSource3.DataSource = dt.Tables[0];
            bindingSource3.ResetBindings(true);
            this.Close();

        }

        private void button3_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
