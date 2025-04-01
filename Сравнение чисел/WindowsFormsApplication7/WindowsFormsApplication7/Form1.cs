using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;

namespace WindowsFormsApplication9
{
    public partial class Form1 : Form
    {
        private OleDbConnection myConnection;
        public static string connectString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Workers.mdb;";

        public Form1()
        {
            InitializeComponent();

            myConnection = new OleDbConnection(connectString);


            myConnection.Open();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string query = "SELECT W_name FROM Workers WHERE W_id = 3";


            OleDbCommand command = new OleDbCommand(query, myConnection);


            textBox1.Text = command.ExecuteScalar().ToString();
        }

        private void button2_Click(object sender, EventArgs e)
        {

            string query = "SELECT W_name, W_position, W_salary FROM Workers ORDER BY W_salary";


            OleDbCommand command = new OleDbCommand(query, myConnection);


            OleDbDataReader reader = command.ExecuteReader();


            listBox1.Items.Clear();


            while (reader.Read())
            {

                listBox1.Items.Add(reader[0].ToString() + " " + reader[1].ToString() + " " + reader[2].ToString() + " ");
            }


            reader.Close();
        }

        private void button3_Click(object sender, EventArgs e)
        {

            string query = "INSERT INTO Worker (W_name, W_position, W_salary) VALUES ('Андрей', 'Танкист', 142)";

            OleDbCommand command = new OleDbCommand(query, myConnection);


            command.ExecuteNonQuery();
        }

        private void button4_Click(object sender, EventArgs e)
        {

            string query = "UPDATE Worker SET w_salary = 123456 WHERE W_id = 3";


            OleDbCommand command = new OleDbCommand(query, myConnection);


            command.ExecuteNonQuery();
        }

        private void button5_Click(object sender, EventArgs e)
        {

            string query = "DELETE FROM Worker WHERE W_id < 3";


            OleDbCommand command = new OleDbCommand(query, myConnection);


            command.ExecuteNonQuery();
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            myConnection.Close();
        }
    }
}
