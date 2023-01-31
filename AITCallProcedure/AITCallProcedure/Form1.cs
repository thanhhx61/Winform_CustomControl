using System;
using System.Windows.Forms;

namespace AITCallProcedure
{
    public partial class Form1 : Form
    {
        AITConnect aITConect = new AITConnect();
        public Form1()
        {
            string notifi = "ERROR : ";
            try
            {
                //Call Function Procedure 
                string querytable = "SELECT * FROM getStaffCount()";
                string counttable = "select * from getStaffCountOne()";
                var namecodeObj = new nameCode();
                var namecodecountObj = new nameCode();
                InitializeComponent();
                //show all list column
                var lisstdata = aITConect.ConnectSqlPostgres<model>(querytable, (object)namecodeObj);
                bindingSource1.DataSource = lisstdata;
                dataGridView1.DataSource = bindingSource1;
                //Count column 
                var countdata = aITConect.ConnectSqlPostgres<int>(counttable, (object)namecodecountObj, "getStaffcountone");
                bindingSource2.DataSource = countdata;
                textBox1.Text = bindingSource2.DataSource.ToString();
            }
            catch (Exception e)
            {
                Message message = new Message();
                message.message = e.Message;
                MessageBox.Show((notifi + message.message));
            }
        }
        class nameCode
        {
            public string name { get; set; }
            public int? code { get; set; }
        }
        class model
        {
            public string name { get; set; }
            public string code { get; set; }
            public DateTime birthDay { get; set; }
        }
    }
}