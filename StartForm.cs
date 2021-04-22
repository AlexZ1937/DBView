using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DBView
{
    public partial class StartForm : Form
    {

        private string stringConnectionContext = "DBView.Properties.Settings.DB_A72B1E_SecureConnectionString";
        public StartForm()
        {
            InitializeComponent();
          
            LoadDataClients();
          
        }

        private void ResizeGrid(object sender, EventArgs e)
        {
            for (int k = 0; k < 2; k++)
            {
                dataGridView1.Columns[k].Width = this.Width / 2;
            }
        }


        private void LoadDataClients()
        {
            try
            {
                dataGridView1.Rows.Clear();
                string connectString = System.Configuration.ConfigurationManager.ConnectionStrings[stringConnectionContext].ConnectionString;
                SqlConnection myConnection = new SqlConnection(connectString);

                myConnection.Open();

                string query = "Select*FROM Clients";
                SqlCommand command = new SqlCommand(query, myConnection);
                SqlDataReader reader = command.ExecuteReader();
                List<string[]> data = new List<string[]>();

                while (reader.Read())
                {
                    data.Add(new string[2]);

                    for (int k = 0; k < 2; k++)
                    {
                        dataGridView1.Columns[k].Width = this.Width / 2;
                        data[data.Count - 1][k] = reader[k].ToString();
                    
                    }
                }

                reader.Close();

                myConnection.Close();
                foreach (string[] s in data)
                    dataGridView1.Rows.Add(s);


                for(int k=0;k<dataGridView1.Rows.Count;k++)
                {
                    dataGridView1.CellMouseClick += LoadToFoundation;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void LoadToFoundation(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.RowIndex <= dataGridView1.Rows.Count-2)
            {
                try
                {

                    textBox1.Text = dataGridView1.Rows[e.RowIndex].Cells[0].Value.ToString();
                    comboBox1.SelectedItem = comboBox1.Items[1];
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }

        private void LoadSecondForm(object sender, EventArgs e)
        {
            if (textBox1.Text.Length > 0)
            {
                int mycase = 0;
                switch (this.comboBox1.SelectedItem.ToString())
                {
                    
                    case "номеру карты":
                        mycase = 1;
                        break;
                    case "id":
                        mycase = 0;
                        break;
                }
                string connectString = System.Configuration.ConfigurationManager.ConnectionStrings[stringConnectionContext].ConnectionString;
                SqlConnection myConnection = new SqlConnection(connectString);
                myConnection.Open();

                string query = "Select*FROM Clients";
                SqlCommand command = new SqlCommand(query, myConnection);
                SqlDataReader reader = command.ExecuteReader();

                int userId = 0;
                string cardnumber = "";

                while (reader.Read())
                {
                    if (mycase != 0)
                    {
                        if (textBox1.Text.ToString() == reader.GetString(mycase))
                        {
                            cardnumber = reader.GetString(1);
                            userId = reader.GetInt32(0);
                            break;
                        }
                    }
                    else
                    {
                        try
                        {
                            if (Convert.ToInt32(textBox1.Text) == reader.GetInt32(mycase))
                            {
                                cardnumber = reader.GetString(1);
                                userId = Convert.ToInt32(textBox1.Text);
                                break;
                            }
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.Message);
                            break;
                        }
                    }
                }

                reader.Close();

                myConnection.Close();

                if (userId == 0)
                {
                    MessageBox.Show("Пользователь не найден!");
                }
                else
                {

                    Form1 form = new Form1();
                    this.Visible = false;
                    form.IDUser = userId;
                    form.CardNumber = cardnumber;
                    form.ShowDialog();
                    this.Visible = true;
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            LoadDataClients();
        }
    }
}
