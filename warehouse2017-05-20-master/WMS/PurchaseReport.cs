using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.Shared;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WMS
{
    public partial class PurchaseReport : Form
    {
       
        ReportDocument cryRpt = new ReportDocument();
        public PurchaseReport()
        {
            InitializeComponent();           
        }
        
        private void PurchaseReport_Load(object sender, EventArgs e)
        {
            dateTimePicker1.CustomFormat = "yyyy-MM-dd";
            dateTimePicker2.CustomFormat = "yyyy-MM-dd";
            DateTime today = DateTime.Today;
            dateTimePicker1.Value = today;
            dateTimePicker2.Value = today;
        }
        private void DataGrid(string InvoiceNo)
        {
            SqlConnection con = new SqlConnection(ConfigurationManager.ConnectionStrings["PosConString"].ToString());
            con.Open();
            SqlCommand cmd = con.CreateCommand();
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = "SELECT PurchaseNo as 'Purchase_No', PurchaseSupplierInvoiceNo as 'Supplier_Invoice_No', PurchaseDate as 'Purchase_Date',  Product.Code as 'Product_Code', Product.Name as 'Product_Name', PurchaseProductPrice as 'Purchase_Price', PurchaseQuantity as 'Product_Quantity',  PurchaseTotal as 'Total' FROM Purchase, Product WHERE Purchase.PurchaseNo='"+ InvoiceNo + "' AND Product.ID=Purchase.PurchaseProductID";
            cmd.ExecuteNonQuery();
            DataTable dt = new DataTable();
            SqlDataAdapter da = new SqlDataAdapter(cmd);
            da.Fill(dt);
            dataGridViewPurchaseReport.DataSource = dt;
            con.Close();
        }
        private void Temp_Purchase_Report_Truncket()
        {
            SqlConnection con = new SqlConnection(ConfigurationManager.ConnectionStrings["PosConString"].ToString());
            con.Open();
            string sql = @"TRUNCATE TABLE TempPurchaseReport";
            SqlCommand cmd = new SqlCommand(sql, con);
            cmd.ExecuteNonQuery();
            con.Close();
        }
        private void Temp_Purchase_Report_Stock()
        {
           
            foreach (DataGridViewRow row in dataGridViewPurchaseReport.Rows)
            {
                //string constring = @"SERVER = Kamrul-pc; DATABASE =pos; user id = sa; PASSWORD = 123";
                using (SqlConnection con1 = new SqlConnection(ConfigurationManager.ConnectionStrings["PosConString"].ToString()))
                {
                    using (SqlCommand cmd = new SqlCommand("INSERT INTO TempPurchaseReport(CompanyID,PurchaseNo,SupplierInvoiceNo,PurchaseDate,ProductCode,ProductName,PurchasePrice,ProductQuantity,Total) VALUES(1,@Purchase_No,@Supplier_Invoice_No,@Purchase_Date,@Product_Code,@Product_Name,@Purchase_Price,@Product_Quantity,@Total)", con1))
                    {
                        cmd.Parameters.AddWithValue("@Purchase_No", row.Cells["Purchase_No"].Value);
                        cmd.Parameters.AddWithValue("@Supplier_Invoice_No", row.Cells["Supplier_Invoice_No"].Value);
                        cmd.Parameters.AddWithValue("@Purchase_Date", row.Cells["Purchase_Date"].Value);
                        cmd.Parameters.AddWithValue("@Product_Code", row.Cells["Product_Code"].Value);
                        cmd.Parameters.AddWithValue("@Product_Name", row.Cells["Product_Name"].Value);
                        cmd.Parameters.AddWithValue("@Purchase_Price", row.Cells["Purchase_Price"].Value);
                        cmd.Parameters.AddWithValue("@Product_Quantity", row.Cells["Product_Quantity"].Value);
                        cmd.Parameters.AddWithValue("@Total", row.Cells["Total"].Value);

                        con1.Open();
                        cmd.ExecuteNonQuery();
                        con1.Close();
                    }
                }
            }
            // MessageBox.Show("Records inserted.");
        }
        
        public string ReportPaths = ReportPath.rPath;
        private void Print_Purchase_Report()
        {
            TableLogOnInfos crtableLogoninfos = new TableLogOnInfos();
            TableLogOnInfo crtableLogoninfo = new TableLogOnInfo();
            ConnectionInfo crConnectionInfo = new ConnectionInfo();

            string rPath = ReportPaths + "CrystalReportPurchaseReport.rpt";
            cryRpt.Load(rPath);

            crConnectionInfo.ServerName = ConfigurationManager.ConnectionStrings["cryServer"].ToString();
            crConnectionInfo.DatabaseName = ConfigurationManager.ConnectionStrings["cryDatabase"].ToString();
            crConnectionInfo.UserID = ConfigurationManager.ConnectionStrings["cryUserID"].ToString();
            crConnectionInfo.Password = ConfigurationManager.ConnectionStrings["cryPass"].ToString();

            foreach (CrystalDecisions.CrystalReports.Engine.Table CrTable in cryRpt.Database.Tables)
            {
                crtableLogoninfo = CrTable.LogOnInfo;
                crtableLogoninfo.ConnectionInfo = crConnectionInfo;
                CrTable.ApplyLogOnInfo(crtableLogoninfo);
            }

            crystalReportViewer1.ReportSource = cryRpt;
            crystalReportViewer1.RefreshReport();
        }
        private void buttonSearch_Click_1(object sender, EventArgs e)
        {
            try
            {
                if (comboBoxInvoiceNo.Text == "")
                {
                    MessageBox.Show("Please Fill Invoice No....!!!");
                    return;
                }

                else
                {
                    DataGrid(comboBoxInvoiceNo.Text);
                    Temp_Purchase_Report_Truncket();
                    Temp_Purchase_Report_Stock();
                    Print_Purchase_Report();

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            }
        private void PurchaseReport_FormClosed(object sender, FormClosedEventArgs e)
        {
            Temp_Purchase_Report_Truncket();
            cryRpt.Close();
            cryRpt.Dispose();
            crystalReportViewer1.ReportSource = null;
            crystalReportViewer1.Dispose();
            crystalReportViewer1 = null;
        }
        SqlCommand command;
        SqlDataAdapter adapter;
        DataTable table;
        public void fillCombo(ComboBox combo, string query, string displayMember, string valueMember)
        {
            SqlConnection con = new SqlConnection(ConfigurationManager.ConnectionStrings["PosConString"].ToString());
            command = new SqlCommand(query, con);
            adapter = new SqlDataAdapter(command);
            table = new DataTable();
            adapter.Fill(table);
            combo.DataSource = table;
            combo.DisplayMember = displayMember;
            combo.ValueMember = valueMember;

        }
        private void dateTimePicker2_ValueChanged(object sender, EventArgs e)
        {
            string date1 = dateTimePicker1.Text;
            string date2 = dateTimePicker2.Text;
            string query = "SELECT DISTINCT(PurchaseNo) FROM Purchase WHERE PurchaseDate  BETWEEN '" + date1 + "' AND '" + date2 + "'";
            fillCombo(comboBoxInvoiceNo, query, "PurchaseNo", "PurchaseNo");

        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            string date1 = dateTimePicker1.Text;
            string date2 = dateTimePicker2.Text;
            string query = "SELECT DISTINCT(PurchaseNo) FROM Purchase WHERE PurchaseDate  BETWEEN '" + date1 + "' AND '" + date2 + "'";
            fillCombo(comboBoxInvoiceNo, query, "PurchaseNo", "PurchaseNo");
        }
    }
}
