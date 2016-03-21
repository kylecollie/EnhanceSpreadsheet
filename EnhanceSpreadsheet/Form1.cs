using System;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Windows.Forms;

namespace EnhanceSpreadsheet
{
    public partial class Form1 : Form
    {
        private string _carNumber;
        private string _currentDescription;
        private int customerID = 119;
        private string _DBUserID = Properties.Settings.Default.DBUserID;
        private string _DBPassword = Properties.Settings.Default.DBPassword;
        
        private string DBConnectionString;
        
        public Form1()
        {
            InitializeComponent();

            DBConnectionString = @"Data Source=PRD-DBSVR-01;Initial Catalog=Host32;Integrated Security=False;User ID=" 
                                              +_DBUserID +
                                              @";Password="
                                              + _DBPassword +
                                              @";Connect Timeout=15;Encrypt=False;TrustServerCertificate=False;ApplicationIntent=ReadWrite;MultiSubnetFailover=False";

        }

        private void btnSelectFile_Click(object sender, EventArgs e)
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
                                //adapter.UpdateCommand = new OleDbCommand("UPDATE [Sheet1$] SET CarNumber = ? WHERE TransactionNumber = ?", connection);
                                adapter.UpdateCommand = new OleDbCommand("UPDATE [Sheet1$] SET CurrentDescription = ? WHERE TransactionNumber = ?", connection);
                                //adapter.UpdateCommand.Parameters.Add("@CarNumber", OleDbType.Char, 255).SourceColumn = "CarNumber";
                                adapter.UpdateCommand.Parameters.Add("@CurrentDescription", OleDbType.Char, 255).SourceColumn = "CurrentDescription";
                                adapter.UpdateCommand.Parameters.Add("@TransactionNumber", OleDbType.Char, 255).SourceColumn = "TransactionNumber";
                                adapter.Fill(sheet1);

                                foreach (DataRow row in sheet1.Rows)
                                {
                                    _currentDescription = GetCurrentDescription(row["TransactionNumber"].ToString());
                                    if (UpdateDescription(row["TransactionNumber"].ToString(), row["MaterialDescription"].ToString()))
                                    {
                                        row["CurrentDescription"] = _currentDescription;
                                        adapter.Update(sheet1); 
                                    }
                                }
                                dgvResults.DataSource = sheet1;

                                //btnGetDescription.Enabled = true;
                                //foreach (DataRow row in sheet1.Rows)
                                //{
                                //    _carNumber = GetCarNumber(row["TransactionNumber"].ToString());
                                //    row["CarNumber"] = _carNumber;
                                //    adapter.Update(sheet1);
                                //}
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

        private bool UpdateDescription(string TransactionNumber, string MaterialDescription)
        {
            throw new NotImplementedException();
            //SqlConnection sqlConnection1 = new SqlConnection("Your Connection String");
            //SqlCommand cmd = new SqlCommand();
            //Int32 rowsAffected;

            //cmd.CommandText = "StoredProcedureName";
            //cmd.CommandType = CommandType.StoredProcedure;
            //cmd.Connection = sqlConnection1;

            //sqlConnection1.Open();

            //rowsAffected = cmd.ExecuteNonQuery();

            //sqlConnection1.Close();
        }

        private string GetCurrentDescription(string bol)
        {
            string DBSql = "SELECT TP.Description FROM Trips T JOIN TripProducts TP ON T.TripID = TP.TripID WHERE T.CustomerID = @CustomerID AND T.BOLCustomer = @BOLCustomer";
            string desc = null;

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
                        desc = o.ToString();
                    }
                    DBConnection.Close();
                }
            }
            return desc;
        }

        private string GetCarNumber(string bol) 
        { 
            string DBSql = "SELECT UnitID FROM Trips T JOIN TripsRef TR ON T.TripID = TR.TripID WHERE T.CustomerID = @CustomerID AND T.BOLCustomer = @BOLCustomer";
            string car = null;
            
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

        private void btnQuit_Click(object sender, EventArgs e)
        {
            Dispose();
        }

        private void btnGetDescription_Click(object sender, EventArgs e)
        {
            //string _transactionNumber = (string)dgvResults.SelectedRows[0].Cells["TransactionNumber"].Value;
            //string _currentDescription = getCurrentDescription(_transactionNumber);
            //dgvResults.SelectedRows[0].Cells["CurrentDescription"].Value = _currentDescription;
        }
    }
}