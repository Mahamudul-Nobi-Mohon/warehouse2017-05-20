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
    public partial class Purchase : Form
    {
        DataTable dt = new DataTable();
        int indexRows;
        public Purchase()
        {
            InitializeComponent();
            textBoxQuantity.KeyPress += new KeyPressEventHandler(Quantity_KeyPress);
            textBoxProductSearch.KeyDown += new KeyEventHandler(textBoxProductSearch_KeyDown);
            textBoxQuantity.KeyDown += new KeyEventHandler(textBoxQuantity_KeyDown);
            ShowTreeViewItem();

            string query = "SELECT * FROM Supplier";
            fillCombo(comboBoxSupplierName, query, "SupplierName", "SupplierID");
        }
        public string SNo = "PUR01";
        SqlConnection con = new SqlConnection(ConfigurationManager.ConnectionStrings["PosConString"].ToString());
        SqlCommand command;
        SqlDataAdapter adapter;
        DataTable table;
        public void fillCombo(ComboBox combo, string query, string displayMember, string valueMember)
        {

            command = new SqlCommand(query, con);
            adapter = new SqlDataAdapter(command);
            table = new DataTable();
            adapter.Fill(table);
            combo.DataSource = table;
            combo.DisplayMember = displayMember;
            combo.ValueMember = valueMember;

        }
        public void ShowTreeViewItem()
        {
            treeViewPurchaseItem.Nodes.Clear();
            string conStr = ConfigurationManager.ConnectionStrings["PosConString"].ToString();
            SqlConnection connection = new SqlConnection(conStr);
            string query = "SELECT * FROM CategoryMain";
            SqlCommand command1 = new SqlCommand(query, connection);

            connection.Open();
            SqlDataReader reader = command1.ExecuteReader();
            int i = 0;
            while (reader.Read())
            {
                treeViewPurchaseItem.Nodes.Add(reader["MaincategoryName"].ToString());
                FirstChild(Convert.ToInt32(reader["MainCategoryID"]), i);
                i++;
            }
            treeViewPurchaseItem.TabStop = false;
            reader.Close();
            connection.Close();
        }

        string conStr = ConfigurationManager.ConnectionStrings["PosConString"].ToString();

        public void FirstChild(int mainID, int i)
        {
            string conStr = ConfigurationManager.ConnectionStrings["PosConString"].ToString();
            SqlConnection connection1 = new SqlConnection(conStr);
            string query1 = "SELECT * FROM Category WHERE MaincategoryID = '" + mainID + "'";
            SqlCommand command11 = new SqlCommand(query1, connection1);

            connection1.Open();
            SqlDataReader reader1 = command11.ExecuteReader();
            int j = 0;
            while (reader1.Read())
            {
                treeViewPurchaseItem.Nodes[i].Nodes.Add(reader1["CategoryName"].ToString());
                SecondChild(Convert.ToInt32(reader1["CategoryID"]), i, j);
                j++;
            }
            reader1.Close();
            connection1.Close();
        }

        public void SecondChild(int catID, int i, int j)
        {
            string conStr = ConfigurationManager.ConnectionStrings["PosConString"].ToString();
            SqlConnection connection12 = new SqlConnection(conStr);
            string query12 = "SELECT * FROM CategorySub WHERE CategoryID = '" + catID + "'";
            SqlCommand command112 = new SqlCommand(query12, connection12);

            connection12.Open();
            SqlDataReader reader12 = command112.ExecuteReader();
            int k = 0;
            while (reader12.Read())
            {
                treeViewPurchaseItem.Nodes[i].Nodes[j].Nodes.Add(reader12["SubCategoryName"].ToString());
                ThirdChild(Convert.ToInt32(reader12["SubCategoryID"]), i, j, k);
                k++;
            }
            reader12.Close();
            connection12.Close();
        }
        public void ThirdChild(int SubcatID, int i, int j, int k)
        {
            string conStr = ConfigurationManager.ConnectionStrings["PosConString"].ToString();
            SqlConnection connection12 = new SqlConnection(conStr);
            string query12 = "SELECT * FROM Product WHERE SubCategoryID = '" + SubcatID + "'";
            SqlCommand command112 = new SqlCommand(query12, connection12);

            connection12.Open();
            SqlDataReader reader12 = command112.ExecuteReader();

            while (reader12.Read())
            {

                TreeNode tn = new TreeNode();
                tn.Tag = reader12["ID"];
                tn.Text = reader12["Name"].ToString();
                treeViewPurchaseItem.Nodes[i].Nodes[j].Nodes[k].Nodes.Add(tn);

            }
            reader12.Close();
            connection12.Close();
        }

        private void treeViewPurchaseItem_AfterSelect(object sender, TreeViewEventArgs e)
        {
           
            try
            {
                textBoxPdoductCode.Text = "";
                textBoxUnitType.Text = "";
                textBoxPdoductCode.Text = treeViewPurchaseItem.SelectedNode.Tag.ToString();
                textBoxProductName.Text = "";
                int pro_id =Convert.ToInt32(textBoxPdoductCode.Text);
                string conStr = ConfigurationManager.ConnectionStrings["PosConString"].ToString();
                SqlConnection connection12 = new SqlConnection(conStr);
                string query12 = "SELECT * FROM Product WHERE ID = '" + pro_id + "'";
                SqlCommand command112 = new SqlCommand(query12, connection12);

                connection12.Open();
                SqlDataReader reader12 = command112.ExecuteReader();

                while (reader12.Read())
                {
                    textBoxProductName.Text = reader12["Name"].ToString();    
                }
                reader12.Close();
                connection12.Close();
                //unit name
                string conStr11 = ConfigurationManager.ConnectionStrings["PosConString"].ToString();
                SqlConnection connection1211 = new SqlConnection(conStr11);
                string query1211 = "SELECT * FROM Product WHERE ID = '" + pro_id + "'";
                SqlCommand command11211 = new SqlCommand(query1211, connection1211);

                connection1211.Open();
                SqlDataReader reader1211 = command11211.ExecuteReader();

                while (reader1211.Read())
                {
                    //textBoxPrice.Text = reader1211["PurchasePrice"].ToString();
                    int unit_id =Convert.ToInt32(reader1211["UnitID"]);
                    GetUnitName(unit_id);
                }
                reader1211.Close();
                connection1211.Close();

                //See Currently Stock

                string conStrPross = ConfigurationManager.ConnectionStrings["PosConString"].ToString();
                SqlConnection connectionPross = new SqlConnection(conStrPross);
                string queryPross = "SELECT * FROM Product WHERE ID = '" + Convert.ToInt32(textBoxPdoductCode.Text) + "'";
                SqlCommand commandPross = new SqlCommand(queryPross, connectionPross);
                connectionPross.Open();
                SqlDataReader readerPross = commandPross.ExecuteReader();

                while (readerPross.Read())
                {
                    textBoxCurrentStock.Text = readerPross["Stock"].ToString();
                }
                readerPross.Close();
                connectionPross.Close();

                // End Currently Stock
            }
            catch (Exception)
            {
                textBoxPdoductCode.Text = "";
                textBoxProductName.Text = "";
                //textBoxPrice.Text = "";
                textBoxUnitType.Text = "";
                textBoxQuantity.Text = "";
                textBoxSupplierInvoiceNo.Text = "";
                MessageBox.Show("Please Select the Product First..");
            }
        }
        public void GetUnitName(int unit_id)
        {
            string conStr111 = ConfigurationManager.ConnectionStrings["PosConString"].ToString();
            SqlConnection connection12111 = new SqlConnection(conStr111);
            string query12111 = "SELECT * FROM Unit WHERE UnitID = '" + unit_id + "'";
            SqlCommand command112111 = new SqlCommand(query12111, connection12111);

            connection12111.Open();
            SqlDataReader reader12111 = command112111.ExecuteReader();

            while (reader12111.Read())
            {
                textBoxUnitType.Text = reader12111["UnitValue"].ToString();
            }
            reader12111.Close();
            connection12111.Close();
        }
        private void Form_Load()
        {
            DateTime now = DateTime.Now;
            textBoxDate.Text = now.ToString("yyyy-MM-dd");
            textBoxDate.ReadOnly = true;
            textBoxTime.Text = now.ToLongTimeString();
            textBoxTime.ReadOnly = true;
            textBoxPurchaseNo.Text = DateTime.Now.ToString("yyyyMMddhhmmssfff");
        }
      //  SqlConnection con = new SqlConnection(ConfigurationManager.ConnectionStrings["PosConString"].ToString());
        //public void fillCombo(ComboBox combo, string query, string displayMember, string valueMember)
        //{
        //    SqlCommand command = new SqlCommand(query, con);
        //    SqlDataAdapter adapter = new SqlDataAdapter(command);
        //    DataTable table = new DataTable();
        //    adapter.Fill(table);
        //    combo.DataSource = table;
        //    combo.DisplayMember = displayMember;
        //    combo.ValueMember = valueMember;

        //}
        private void FormPurchaseItemBody_Load(object sender, EventArgs e)
        {
            string query = "SELECT * FROM LocationMain";
            comboBoxWarehouse.SelectedValue = -1;
            fillCombo(comboBoxWarehouse, query, "LocationMainName", "LocationMainID");


            //FormPurchaseClosed();
            Form_Load(); 
            Auto_Complete();
            this.ActiveControl = textBoxSupplierInvoiceNo; 


            dt.Columns.AddRange(new DataColumn[5] { new DataColumn("Code", typeof(string)),
                new DataColumn("Name", typeof(string)),new DataColumn("warehouse_name", typeof(string)),new DataColumn("comboBoxFloor", typeof(string)),new DataColumn("Quantity", typeof(string))});

        }
        private void Auto_Complete()
        {
            //Auto Complete search
            textBoxProductSearch.AutoCompleteMode = AutoCompleteMode.Suggest;
            textBoxProductSearch.AutoCompleteSource = AutoCompleteSource.CustomSource;
            string conStr = ConfigurationManager.ConnectionStrings["PosConString"].ToString();
            SqlConnection conSS = new SqlConnection(conStr);
            AutoCompleteStringCollection col = new AutoCompleteStringCollection();
            col.Clear();
            conSS.Open();
            string sql = "SELECT * FROM Product";
            SqlCommand cmd = new SqlCommand(sql, conSS);
            SqlDataReader sdr = null;
            sdr = cmd.ExecuteReader();
            while (sdr.Read())
            {
                col.Add(sdr["Code"].ToString());
                col.Add(sdr["Name"].ToString());

            }
            sdr.Close();
            textBoxProductSearch.AutoCompleteCustomSource = col;
            conSS.Close();
        }



        //public void FormPurchaseClosed()
        //{
        //    SqlConnection con = new SqlConnection(conStr);
        //    con.Open();
        //    string sql = @"TRUNCATE TABLE TempPurchase";
        //    SqlCommand cmd = new SqlCommand(sql, con);
        //    cmd.ExecuteNonQuery();
        //    con.Close();
        //}
       

        public void TempPurchase(int pro_id, string name, double qantity, double price, string purchase_no, double price_total)
        {
            SqlConnection connection = new SqlConnection(conStr);
            string query = "INSERT INTO TempPurchase(CompanyID,ProductID,ProductName,Quantity,PurchasePrice,PurchaseNo,PurchasePriceTotal,PurchaseDate) VALUES(1,'" + pro_id + "','" + name + "','" + qantity + "','" + price + "','" + purchase_no + "','" + price_total + "','" + textBoxDate.Text + "')";
            SqlCommand command = new SqlCommand(query, connection);
            connection.Open();
            int rowEffict = command.ExecuteNonQuery();
            connection.Close();
            
        }
        
        public double Temp_PreviousStock(int productcode)
        {
            double total_quantity = 0;
            string conStrCalq = ConfigurationManager.ConnectionStrings["PosConString"].ToString();
            SqlConnection connectionCalq = new SqlConnection(conStrCalq);
            string queryCalq = "SELECT * FROM TempPurchase WHERE ProductID = '" + productcode + "'";
            SqlCommand commandCalq = new SqlCommand(queryCalq, connectionCalq);
            connectionCalq.Open();
            SqlDataReader readerCalq = commandCalq.ExecuteReader();

            while (readerCalq.Read())
            {
                total_quantity = Convert.ToDouble(readerCalq["Quantity"]);
            }
            readerCalq.Close();
            connectionCalq.Close();
            return total_quantity;
        }
        

        public void DataGrid()
        {
            SqlConnection con = new SqlConnection(ConfigurationManager.ConnectionStrings["PosConString"].ToString());
            con.Open();
            SqlCommand cmd = con.CreateCommand();
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = "SELECT ID as 'ID', ProductID as 'ProductID', ProductName as 'Product Name', Quantity as 'Quantity', PurchasePrice as 'Price', PurchaseNo as 'Purchase No', PurchasePriceTotal as 'Total' FROM TempPurchase";
            cmd.ExecuteNonQuery();
            DataTable dt = new DataTable();
            SqlDataAdapter da = new SqlDataAdapter(cmd);
            da.Fill(dt);
            dataGridViewPurchase.DataSource = dt;
            con.Close();
        }
        public void UpdateTemp_ProductStock(double total_quantity,double total, int pro_id)
        {
            SqlConnection connection1 = new SqlConnection(conStr);
            string query1 = "UPDATE TempPurchase SET Quantity = '" + total_quantity + "', PurchasePriceTotal = '" + total + "' WHERE ProductID = '" + pro_id + "'";
            SqlCommand command1 = new SqlCommand(query1, connection1);
            connection1.Open();
            int rowEffict1 = command1.ExecuteNonQuery();
            connection1.Close();
            if (rowEffict1 > 0)
            {
                DataGrid();
            }
        }
        public void Display_Three_Item()
        {
            double ProTotal = 0;
            double AmountTotal = 0;
            int i = 0;
            string conStrs = ConfigurationManager.ConnectionStrings["PosConString"].ToString();
            SqlConnection connections = new SqlConnection(conStrs);
            string querys = "SELECT * FROM TempPurchase";
            SqlCommand commands = new SqlCommand(querys, connections);
            connections.Open();
            SqlDataReader readers = commands.ExecuteReader();

            while (readers.Read())
            {
                ProTotal = ProTotal + Convert.ToDouble(readers["Quantity"]);
                AmountTotal = AmountTotal + Convert.ToDouble(readers["PurchasePriceTotal"]);
                i++;
            }
            readers.Close();
            connections.Close();
            textBoxItemTotal.Text = i.ToString();
            textBoxProductTotal.Text = ProTotal.ToString();
            textBoxInvoiceTotalAmount.Text = AmountTotal.ToString();
        }
        private void buttonAdd_Click(object sender, EventArgs e)
        {

            if (textBoxSupplierInvoiceNo.Text == "")
            {
                MessageBox.Show("Please Fill the Supplier Invoice No..");
            }
            else if (textBoxPdoductCode.Text == "")
            {
                MessageBox.Show("Please Select a Product on Data Tree view..!!");
            }
            else if (textBoxQuantity.Text == "")
            {
                MessageBox.Show("Please Fill the quantity..");
            }

            else if (comboBoxFloor.Text == "")
            {
                MessageBox.Show("Please Select the Floor Name..");
            }

            else
            {
                int isEdited = 0;

                if (textBoxQuantity.Text == "")
                {
                    textBoxQuantity.Text = "0";

                }

                if (dataGridViewPurchase.Rows.Count > 0)
                {


                    String searchValue = textBoxPdoductCode.Text;
                    int rowIndex = -1;
                    if (isEdited == 0)
                    {
                        foreach (DataGridViewRow row in dataGridViewPurchase.Rows)
                        {
                            if (row.Cells["Code"].Value != null) // Need to check for null if new row is exposed
                            {
                                rowIndex = -1;
                                if (row.Cells["Code"].Value.ToString().Equals(searchValue) && row.Cells["warehouse_name"].Value.ToString().Equals(comboBoxWarehouse.Text))
                                {
                                    rowIndex = row.Index;
                                    // break;
                                    DataGridViewRow newDataRow = dataGridViewPurchase.Rows[rowIndex];
                                    double sum = Convert.ToDouble(newDataRow.Cells[4].Value) + Convert.ToDouble(textBoxQuantity.Text);
                                    newDataRow.Cells[0].Value = textBoxPdoductCode.Text;
                                    newDataRow.Cells[1].Value = textBoxProductName.Text;
                                    newDataRow.Cells[2].Value = comboBoxWarehouse.Text;
                                    newDataRow.Cells[3].Value = comboBoxFloor.Text;
                                    newDataRow.Cells[4].Value = sum;
                                    isEdited = 1;
                                    break;
                                }

                            }
                        }
                    }

                }
                if (isEdited != 1)
                {
                    dt.Rows.Add(textBoxPdoductCode.Text, textBoxProductName.Text, comboBoxWarehouse.Text, comboBoxFloor.Text, textBoxQuantity.Text);
                    this.dataGridViewPurchase.DataSource = dt;
                    //clearAll();
                }

            }

        }
       
        private void FromPurchaseClosed(object sender, FormClosedEventArgs e)
        {
            //FormPurchaseClosed();
        }

        //Validation only for numeric number
        private void Quantity_KeyPress(object sender, KeyPressEventArgs e)
        {

            if(e.KeyChar >= '0' && e.KeyChar <= '9' || e.KeyChar <= '.' || e.KeyChar == (char)Keys.Back) //The  character represents a backspace
              {
                e.Handled = false; //Do not reject the input
              }
            else
                {
                  e.Handled = true; //Reject the input
               }
        }

        private void buttonUpdate_Click(object sender, EventArgs e)
        {
            if (textBoxQuantity.Text == "")
            {
                textBoxQuantity.Text = "0";
            }
            if (dataGridViewPurchase.Rows.Count > 1)
            {
                indexRows = dataGridViewPurchase.CurrentCell.RowIndex;
                DataGridViewRow newDataRow = dataGridViewPurchase.Rows[indexRows];
                newDataRow.Cells[0].Value = textBoxPdoductCode.Text;
                newDataRow.Cells[1].Value = textBoxProductName.Text;
                newDataRow.Cells[2].Value = comboBoxWarehouse.Text;
                newDataRow.Cells[3].Value = comboBoxFloor.Text;
                newDataRow.Cells[4].Value = textBoxQuantity.Text;
                //clearAll();
            }
            else
            {
                MessageBox.Show("Please select a data. then click update button.");
            }
        }
        private void dataGridViewPurchase_CellMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            textBoxPdoductCode.Text = textBoxProductName.Text = textBoxQuantity.Text =  textBoxUnitType.Text = "";
            textBoxPdoductCode.Text = dataGridViewPurchase.SelectedRows[0].Cells[0].Value.ToString();
            textBoxProductName.Text = dataGridViewPurchase.SelectedRows[0].Cells[1].Value.ToString();
            comboBoxWarehouse.Text = dataGridViewPurchase.SelectedRows[0].Cells[2].Value.ToString();
            comboBoxFloor.Text = dataGridViewPurchase.SelectedRows[0].Cells[3].Value.ToString();
            textBoxQuantity.Text = dataGridViewPurchase.SelectedRows[0].Cells[4].Value.ToString();
            //textBoxPrice.Text = dataGridViewPurchase.SelectedRows[0].Cells[4].Value.ToString();
        }

        private void buttonDelete_Click(object sender, EventArgs e)
        {
            if (dataGridViewPurchase.Rows.Count >= 1)
            {
                int rowIndex = dataGridViewPurchase.CurrentCell.RowIndex;
                dataGridViewPurchase.Rows.RemoveAt(rowIndex);
                //clearAll();
            }
            else
            {
                MessageBox.Show("Please select a data. then click Remove button.");
            }
        }
        private void Purchase_Details_Insert()
        {
            string remarks = "Na";
            int supplier_id;
            Int32.TryParse(comboBoxSupplierName.SelectedValue.ToString(), out supplier_id);
            foreach (DataGridViewRow row in dataGridViewPurchase.Rows)
            {
                using (SqlConnection con = new SqlConnection(ConfigurationManager.ConnectionStrings["PosConString"].ToString()))
                {
                    using (SqlCommand cmd = new SqlCommand("INSERT INTO Purchase(PurchaseNo,CompanyID,PurchaseDate,SupplierID,PurchaseSupplierInvoiceNo,PurchaseRemarks,PurchaseProductID,PurchaseProductPrice,PurchaseQuantity,PurchaseTotal) VALUES('" + SNo + "', 1, '" + textBoxDate.Text + "','"+ supplier_id + "','"+textBoxSupplierInvoiceNo.Text+"','" + remarks + "',@ProductID,@Price,@Quantity,@Total)", con))
                    {
                        cmd.Parameters.AddWithValue("@ProductID", row.Cells["ProductID"].Value);
                        cmd.Parameters.AddWithValue("@Price", row.Cells["Price"].Value);
                        cmd.Parameters.AddWithValue("@Quantity", row.Cells["Quantity"].Value);
                        cmd.Parameters.AddWithValue("@Total", row.Cells["Total"].Value);

                        con.Open();
                        cmd.ExecuteNonQuery();
                        con.Close();
                    }
                }
            }
          
        }
        public void Clear_All_Last()
        {
            textBoxPdoductCode.Text = textBoxProductName.Text = textBoxInvoiceTotalAmount.Text = "";
        }
        private void Purchase_Insert()
        {
            if (textBoxPdoductCode.Text == "")
            {
                MessageBox.Show("Please Select a Product in Tree View....!!!");
            }
            else
            {
                try
                {
                    //SqlConnection connection = new SqlConnection(conStr);
                    //string query = "INSERT INTO purchase VALUES('" + textBoxPurchaseNo.Text + "','" + textBoxDate.Text + "','1','" + textBoxSupplierInvoiceNo.Text + "','Na','" + textBoxInvoiceTotalAmount.Text + "')";
                    //SqlCommand command = new SqlCommand(query, connection);
                    //connection.Open();
                    //int rowEffict = command.ExecuteNonQuery();
                    //connection.Close();
                    Clear_All_Last();
                    //FormPurchaseClosed();
                    DataGrid();
                   // Dispose();

                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }

        }

        private void Update_Product_Stock()
        {
            
            foreach (DataGridViewRow row in dataGridViewPurchase.Rows)
            {
                using (SqlConnection con = new SqlConnection(ConfigurationManager.ConnectionStrings["PosConString"].ToString()))
                {
                    using (SqlCommand cmd = new SqlCommand("UPDATE Product SET Stock =(SELECT Stock FROM Product WHERE ID = @ProductID) + @Quantity  WHERE ID = @ProductID", con))
                    {
                        
                        cmd.Parameters.AddWithValue("@ProductID", Convert.ToDouble(row.Cells["ProductID"].Value));
                        cmd.Parameters.AddWithValue("@Quantity", Convert.ToDouble(row.Cells["Quantity"].Value));
                        con.Open();
                        cmd.ExecuteNonQuery();
                        con.Close();
                    }
                }
            }
        }
        
        private void UpdateTempPurchaseNo(int countPurchase)
        {
            SNo = SNo + countPurchase.ToString("D6");
        }
        private int CountPurchaseNo()
        {
            int no = 0;
            SqlConnection con = new SqlConnection(conStr);
            string query = "SELECT COUNT(DISTINCT(PurchaseNo)) as 'No' FROM Purchase";
            SqlCommand command = new SqlCommand(query, con);
            con.Open();
            SqlDataReader reader = command.ExecuteReader();

            while (reader.Read())
            {
                no = Convert.ToInt32(reader["No"]);
            }

            reader.Close();
            con.Close();
            return no;

        }
        private void buttonPurchase_Click(object sender, EventArgs e)
        {
            DialogResult dr = MessageBox.Show("Are you sure, you will Purchase these product?", "Purchase Product", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (dr == DialogResult.No)
            {
                return;
            }
            else
            {
                try
                {
                          int countPurchase = CountPurchaseNo();
                          UpdateTempPurchaseNo(countPurchase + 1);
                          Purchase_Details_Insert();
                          Update_Product_Stock();
                          Purchase_Insert();

                        DataGrid();
                        Form_Load();
                        textBoxSupplierInvoiceNo.Text = "";
                        MessageBox.Show("Purchase Successfully....!!!!");
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
                
            }
            }

        private bool IsNumeric(string input)
        {
            int test;
            return int.TryParse(input, out test);
        }
        private double Currently_Stock()
        {
            double stock = 0.0;
            string conStrPross = ConfigurationManager.ConnectionStrings["PosConString"].ToString();
            SqlConnection connectionPross = new SqlConnection(conStrPross);
            // string queryPross = "SELECT * FROM Product WHERE ID = '" + Convert.ToInt32(textBoxPdoductCode.Text) + "'";
            string queryPross = " select stock from product where WarehouseID = '"+comboBoxWarehouse.SelectedValue+"' and Code = '"+textBoxPdoductCode.Text+"'";
            SqlCommand commandPross = new SqlCommand(queryPross, connectionPross);
            connectionPross.Open();
            SqlDataReader readerPross = commandPross.ExecuteReader();

            while (readerPross.Read())
            {
                stock = Convert.ToDouble(readerPross["Stock"]);
            }

            readerPross.Close();
            connectionPross.Close();
            return stock;
        }
        private void textBoxProductSearch_TextChanged(object sender, EventArgs e)
        {
            string pro_id = textBoxProductSearch.Text;
            string pro_code = "";
            
                SqlConnection conww = new SqlConnection(conStr);
                conww.Open();
                string sqlww = "SELECT * FROM Product WHERE Name='" + textBoxProductSearch.Text + "'OR Code='" + textBoxProductSearch.Text + "'";
                SqlCommand cmdww = new SqlCommand(sqlww, conww);
                SqlDataReader sdrww = null;
                sdrww = cmdww.ExecuteReader();
                while (sdrww.Read())
                {
                pro_id = sdrww["ID"].ToString();
                pro_code = sdrww["code"].ToString();
                }
                sdrww.Close();
                conww.Close();
            textBoxCurrentStock.Text = Currently_Stock().ToString();

            try
            {
                textBoxPdoductCode.Text = "";
                textBoxPdoductCode.Text = pro_code;
               // int pro_id = Convert.ToInt32(textBoxPdoductCode.Text);
                textBoxProductName.Text = "";
                SqlConnection connection12 = new SqlConnection(conStr);
                string query12 = "SELECT * FROM Product WHERE ID = '" + pro_id + "'";
                SqlCommand command112 = new SqlCommand(query12, connection12);

                connection12.Open();
                SqlDataReader reader12 = command112.ExecuteReader();

                while (reader12.Read())
                {

                    textBoxProductName.Text = reader12["Name"].ToString();
                }
                reader12.Close();
                connection12.Close();

                //unit name
                string conStr11 = ConfigurationManager.ConnectionStrings["PosConString"].ToString();
                SqlConnection connection1211 = new SqlConnection(conStr11);
                string query1211 = "SELECT * FROM Product WHERE ID = '" + pro_id + "'";
                SqlCommand command11211 = new SqlCommand(query1211, connection1211);

                connection1211.Open();
                SqlDataReader reader1211 = command11211.ExecuteReader();

                while (reader1211.Read())
                {
                   // textBoxPrice.Text = reader1211["PurchasePrice"].ToString();
                    int unit_id = Convert.ToInt32(reader1211["UnitID"]);
                    GetUnitName(unit_id);
                }
                reader1211.Close();
                connection1211.Close();

                //See Currently Stock
                textBoxCurrentStock.Text = Currently_Stock().ToString();

            }
            catch (Exception)
            {
                textBoxPdoductCode.Text = "";
                textBoxProductName.Text = "";
               // textBoxPrice.Text = "";
                textBoxCurrentStock.Text = "";
                textBoxUnitType.Text = "";
            }

      
        //
    }

        private void comboBoxSupplierName_SelectedIndexChanged(object sender, EventArgs e)
        {
            textBoxAddress.Text = "";
            int id;
            Int32.TryParse(comboBoxSupplierName.SelectedValue.ToString(), out id);
            string conStr = ConfigurationManager.ConnectionStrings["PosConString"].ToString();
            SqlConnection con = new SqlConnection(conStr);
            string query = "SELECT SupplierAddress FROM Supplier WHERE SupplierID = " + id;
            SqlCommand command112 = new SqlCommand(query, con);
            con.Open();
            SqlDataReader reader12 = command112.ExecuteReader();
            while (reader12.Read())
            {
                textBoxAddress.Text = reader12["SupplierAddress"].ToString();          
            }
            reader12.Close();
            con.Close();
        }

        private void label5_Click(object sender, EventArgs e)
        {

        }

        private void label17_Click(object sender, EventArgs e)
        {

        }

        private void textBoxQuantity_KeyDown(object sender, KeyEventArgs e)
        {
            
            if (e.KeyCode == Keys.Enter)
            {
                this.ActiveControl = buttonAdd;
            }
        }

        private void textBoxProductSearch_KeyDown(object sender, KeyEventArgs e)
        {
           
            if (e.KeyCode == Keys.Enter)
            {
                this.ActiveControl = textBoxQuantity;
            }
            if (e.KeyCode == Keys.Escape)
            {
                this.ActiveControl = buttonPurchase;
            }
        }

        private void buttonPurchase_KeyDown(object sender, KeyEventArgs e)
        {

        }

        private void buttonPurchase_MouseEnter(object sender, EventArgs e)
        {
            buttonPurchase.BackColor = Color.YellowGreen;
        }

        private void textBoxSupplierInvoiceNo_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                this.ActiveControl = textBoxProductSearch;
            }
        }

        private void buttonAddSupplier_Click(object sender, EventArgs e)
        {
            AddSupplier addsupp = new AddSupplier();
            addsupp.Show();
        }

        private void panel2_Paint(object sender, PaintEventArgs e)
        {

        }

        private void comboBoxWarehouse_SelectedIndexChanged(object sender, EventArgs e)
        {
            int val;
            Int32.TryParse(comboBoxWarehouse.SelectedValue.ToString(), out val);

            string query2 = "SELECT * FROM Location where LocationMainID = " + val + " ";
            //comboBoxRoomName.SelectedValue = -1;
            fillCombo(comboBoxFloor, query2, "LocationName", "LocationID");

            textBoxCurrentStock.Text =  Currently_Stock().ToString();
        }

        private void dataGridViewPurchase_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            textBoxPdoductCode.Text = dataGridViewPurchase.CurrentRow.Cells[0].Value.ToString();
            textBoxQuantity.Text = dataGridViewPurchase.CurrentRow.Cells[1].Value.ToString();
        }
    }
}
