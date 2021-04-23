using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace DBView
{
    public partial class Form1 : Form
    {
        int columncount = 2;
       public int IDUser=0;
        public string CardNumber = "";
        private string stringConnectionContext = "DBView.Properties.Settings.DB_A72B1E_SecureConnectionString";
        public Form1()
        {
            InitializeComponent();

            this.SizeChanged += ResizeGrid;
            this.Shown += LoadFromDB;

        }

        private void LoadFromDB(object sender, EventArgs e)
        {
            LoadData();
        }

        private void ResizeGrid(object sender, EventArgs e)
        {
            dataGridView3.Width = this.Width / 3;
            dataGridView2.Width = (this.Width / 3)*2;
       

            dataGridView3.Height = this.Height- dataGridView2.Location.Y-10;
            dataGridView2.Height = this.Height- dataGridView2.Location.Y - 10;
         

            for (int k = 0; k < 2; k++)
            {
                dataGridView2.Columns[k].Width = dataGridView2.Width / 2 - 30;
            }

            dataGridView3.Columns[0].Width = dataGridView3.Width - 30;
           
            dataGridView3.Location = new Point(dataGridView2.Location.X + dataGridView2.Width, dataGridView2.Location.Y);
        }
 

        private void CreateExcel(string cardnumber, List<Contact> contacts)
        {
            Excel.Application excel_app = null;
            try
            {
              
                excel_app = new Excel.ApplicationClass();
                excel_app.Visible = false;
                excel_app.SheetsInNewWorkbook = 1;
                excel_app.Interactive = false;
                excel_app.Workbooks.Add(Type.Missing);
                Excel.Workbook workbook = null;

                  bool retry = true;
                int counter = 0;
                //System.Runtime.InteropServices.COMException: "Исключение из HRESULT: 0x800AC472"
            
                do
                {
                    string mesage="";
                    try
                    {
                        workbook = excel_app.Workbooks[1];
                        workbook.Sheets.Add(Type.Missing);
                        Excel.Worksheet sheet = (Excel.Worksheet)workbook.Sheets[1];
                        Excel.Range excelcells = null;
                        excelcells = (Excel.Range)sheet.Cells[1, 2];
                        excelcells.ColumnWidth = 20;
                        
                        excelcells.Value2 = "Номер";
                        excelcells = (Excel.Range)sheet.Cells[1, 1];
                        excelcells.ColumnWidth = 40;
                        excelcells.Value2 = "Контактное имя";
                        for (int k = 0; k < contacts.Count; k++)
                        {
                            excelcells = (Excel.Range)sheet.Cells[k + 2, 1];
                           
                            excelcells.Value2 = contacts[k].GetName();
                            excelcells = (Excel.Range)sheet.Cells[k + 2, 2];
                            excelcells.NumberFormat = "#########";
                            excelcells.Value2 = contacts[k].GetNumber();
                        }
                        retry = false;
                    }
                    catch (System.Runtime.InteropServices.COMException ex)
                    {
                        counter++;
                        retry = true;
                        mesage = ex.Message;
                       


                    }
                    //if(counter>2500)
                    //{
                    //    MessageBox.Show("Ошибка записи в exel "+mesage);
                    //    break;
                    //}
                }
                while (retry);
              
                if (File.Exists(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "/Client_" + CardNumber + ".xlsx"))
                {
                   
                    File.Delete(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "/Client_" + CardNumber + ".xlsx");
                }
                workbook.SaveAs("Client_"+CardNumber+".xlsx");
              
               
                ////if (File.Exists("Clientcard " + cardnumber+".xsln")!=true)
                ////{
                //    workbook = excel_app.Workbooks.Add(Type.Missing);

                ////}
                ////else
                ////{
                ////    workbook = excel_app.Workbooks.Open("Clientcard " + cardnumber + ".xsln");
                ////}
                //////workbook = excel_app.Workbooks.Open("Client " + cardnumber);
                //sheet = (Excel.Worksheet)workbook.Sheets[1];
                //sheet.Cells[1, 1] = "Контактное имя";
                //sheet.Cells[1, 2] = "Номер";
                //for(int k=0;k<contacts.Count;k++)
                //{
                //    sheet.Cells[k + 2, 1] = contacts[k].GetName();
                //    sheet.Cells[k + 2, 2] = contacts[k].GetNumber();
                //}
                //workbook.SaveAs("Client " + cardnumber+".xsln");
                //workbook.Close(true, Type.Missing, Type.Missing);

            }
            catch(Exception ex)
            {
                if (ex.Message.Contains("Исключение из HRESULT: 0x800AC472")!=true)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            finally
            {
                excel_app.Quit();
            }
           
        }

  private void LoadData()
        {


            try
            {
                dataGridView2.Rows.Clear();
                
                string connectString = System.Configuration.ConfigurationManager.ConnectionStrings[stringConnectionContext].ConnectionString;
                SqlConnection myConnection = new SqlConnection(connectString);
                
                myConnection.Open();

                string query = "Select MessageFrom, MessageText FROM Messages Where ClientID =" + IDUser;

                SqlCommand command = new SqlCommand(query, myConnection);

                SqlDataReader reader = command.ExecuteReader();

                List<string[]> data = new List<string[]>();

                while (reader.Read())
                {
                    data.Add(new string[columncount]);

                    for (int k = 0; k < columncount; k++)
                    {
                        data[data.Count - 1][k] = reader[k].ToString();
                     
                    }
                }
                reader.Close();
                for (int k = 0; k < data.Count; k++)
                {
                    dataGridView2.Rows.Add(data[k]);
                }
                query = "Select ContactName, ContactNumber FROM Contacts Where ClientID =" + IDUser;
                command = new SqlCommand(query, myConnection);
                reader = command.ExecuteReader();
                List<Contact> contacts = new List<Contact>();
                data.Clear();
                while (reader.Read())
                {
                    contacts.Add(new Contact(reader[0].ToString(), reader[1].ToString()));
                }       
                reader.Close();
          
       

                query = "Select ApplicationName FROM Applications Where ClientID =" + IDUser;
                command = new SqlCommand(query, myConnection);
                reader = command.ExecuteReader();

                while (reader.Read())
                {
                    data.Add(new string[1]);
                    data[data.Count - 1][0] = reader[0].ToString();
                }
              
                for (int k = 0; k < data.Count; k++)
                {
                    dataGridView3.Rows.Add(data[k]);
                }
                EventArgs arg = new EventArgs();
                this.OnSizeChanged(arg);

                reader.Close();
                myConnection.Close();


                CreateExcel(CardNumber, contacts);

            }
            catch (Exception ex)
            {
                if (ex.Message.Contains("Исключение из HRESULT: 0x800AC472") != true)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            LoadData();
        }

   
    }
}
