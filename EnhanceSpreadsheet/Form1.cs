using System;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Windows.Forms;

namespace EnhanceSpreadsheet
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            DataTable sheet1 = new DataTable();
            OleDbConnectionStringBuilder csBuilder = new OleDbConnectionStringBuilder();
            csBuilder.Provider = "Microsoft.ACE.OLEDB.12.0";
            csBuilder.Add("Extended Properties", "Excel 12.0 Xml;HDR=YES");
            OpenFileDialog openFD = new OpenFileDialog();

            openFD.InitialDirectory = @"c:\";
            openFD.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
            openFD.FilterIndex = 2;
            openFD.RestoreDirectory = true;

            if (openFD.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    if ((csBuilder.DataSource = openFD.FileName) != null)
                    {
                        using (OleDbConnection connection = new OleDbConnection(csBuilder.ConnectionString))
                        {
                            connection.Open();
                            string selectSql = @"SELECT * FROM [Sheet1$]";
                            using (OleDbDataAdapter adapter = new OleDbDataAdapter(selectSql, connection))
                            {
                                adapter.UpdateCommand = new OleDbCommand("UPDATE [Sheet1$] SET CarNumber = ? WHERE TransactionNumber = ?", connection);
                                adapter.UpdateCommand.Parameters.Add("@CarNumber", OleDbType.Char, 255).SourceColumn = "CarNumber";
                                adapter.UpdateCommand.Parameters.Add("@TransactionNumber", OleDbType.Char, 255).SourceColumn = "TransactionNumber";
                                adapter.Fill(sheet1);
                                string _carNumber;
                                foreach (DataRow row in sheet1.Rows)
                                {
                                    _carNumber = GetCarNumber(row["TransactionNumber"].ToString());
                                    row["CarNumber"] = _carNumber;
                                    adapter.Update(sheet1);
                                }
                                dgvResults.DataSource = sheet1;
                            }
                            connection.Close();
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error: Could not read file from disk. Original error: " + ex.Message);
                    throw;
                }
            }
        }

        private string GetCarNumber(string bol)
        { 
            string DBConnectionString = @"Data Source=PRD-DBSVR-01;
                                                      Initial Catalog=Host32;
                                                      Integrated Security=True;
                                                      Connect Timeout=15;
                                                      Encrypt=False;
                                                      TrustServerCertificate=False;
                                                      ApplicationIntent=ReadWrite;
                                                      MultiSubnetFailover=False";

            string DBSql = "SELECT UnitID FROM Trips T JOIN TripsRef TR ON T.TripID = TR.TripID WHERE T.CustomerID = @CustomerID AND T.BOLCustomer = @BOLCustomer";
            string car = null;
            int customerID = 119;

            using (SqlConnection DBConnection = new SqlConnection(DBConnectionString))
            {
                using (SqlCommand DBCommand = new SqlCommand(DBSql, DBConnection))
                {
                    DBCommand.CommandType = CommandType.Text;
                    DBCommand.Parameters.AddWithValue("@BOLCustomer", bol);
                    DBCommand.Parameters.AddWithValue("@CustomerID", customerID);
                    DBConnection.Open();
                    object o = DBCommand.ExecuteScalar();
                    if (o != null)
                    {
                        car = o.ToString();
                    }
                    DBConnection.Close();
                }
            }
            return car;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Dispose();
        }
    }
}