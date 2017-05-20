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
    public partial class Sales : Form
    {
        public Sales()
        {
            InitializeComponent();
            
            textBoxQuantity.KeyPress += new KeyPressEventHandler(QuantityKeyPress);
            
            textBoxQuantity.KeyDown += new KeyEventHandler(textBoxQuantity_KeyDown);
           
           
            

            ShowTreeViewItem();

            //string query12 = "SELECT * FROM Customer";
            //fillCombo(textBoxCustomerName, query12, "CustomerName", "CustomerID");


           
            
        }
        public string SNo = "01";
        public string currentuser = Login.loguser;
        public int CompanyID = GetUserLogInfo.Company();

        private void TextboxValue(string ParameterName, int ParameterValue, ParameterField myParameterField, ParameterDiscreteValue myDiscreteValue, ParameterFields myParameterFields)
        {
            myParameterField.ParameterFieldName = ParameterName;
            myDiscreteValue.Value = ParameterValue;
            myParameterField.CurrentValues.Add(myDiscreteValue);
            myParameterFields.Add(myParameterField);
        }

        public void fillCombo(ComboBox combo, string query, string displayMember, string valueMember)
        {
            SqlCommand command;
            SqlDataAdapter adapter;
            DataTable table;
            SqlConnection conss = new SqlConnection(ConfigurationManager.ConnectionStrings["PosConString"].ToString());
            command = new SqlCommand(query, conss);
            adapter = new SqlDataAdapter(command);
            table = new DataTable();
            adapter.Fill(table);
            combo.DataSource = table;
            combo.DisplayMember = displayMember;
            combo.ValueMember = valueMember;

        }
        public void ShowTreeViewItem()
        {
            treeViewPurchaseItemSales.Nodes.Clear();
            string conStr = ConfigurationManager.ConnectionStrings["PosConString"].ToString();
            SqlConnection connection = new SqlConnection(conStr);
            string query = "SELECT * FROM CategoryMain";
            SqlCommand command1 = new SqlCommand(query, connection);

            connection.Open();
            SqlDataReader reader = command1.ExecuteReader();
            int i = 0;
            while (reader.Read())
            {
                treeViewPurchaseItemSales.Nodes.Add(reader["MaincategoryName"].ToString());
                FirstChild(Convert.ToInt32(reader["MainCategoryID"]), i);
                i++;


            }
            treeViewPurchaseItemSales.TabStop = false;
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
                treeViewPurchaseItemSales.Nodes[i].Nodes.Add(reader1["CategoryName"].ToString());
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
                treeViewPurchaseItemSales.Nodes[i].Nodes[j].Nodes.Add(reader12["SubCategoryName"].ToString());
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
                treeViewPurchaseItemSales.Nodes[i].Nodes[j].Nodes[k].Nodes.Add(tn);

            }
            reader12.Close();
            connection12.Close();
        }
        ////public void Temp_Sales_Truncate()
        ////{
           
        ////    SqlConnection con = new SqlConnection(conStr);
        ////    con.Open();
        ////    string sql = @"DELETE FROM TempSales WHERE TempSalesCompanyID='"+CompanyID+"'";
        ////    SqlCommand cmd = new SqlCommand(sql, con);
        ////    cmd.ExecuteNonQuery();
        ////    con.Close();

        ////    TempAmountTruncate();
        ////}

        private void Sales_Load(object sender, EventArgs e)
        {
            buttonReports.Enabled = false;
           // Temp_Sales_Truncate();

            Load_Form();
            Auto_Complete();

            this.ActiveControl = textBoxProductSearch;

            DateTime today = DateTime.Today;
            
        }
        private void Auto_Complete()
        {
            //Auto Complete search
            textBoxProductSearch.AutoCompleteMode = AutoCompleteMode.Suggest;
            textBoxProductSearch.AutoCompleteSource = AutoCompleteSource.CustomSource;

            textBoxCustomerName.AutoCompleteMode = AutoCompleteMode.Suggest;
            textBoxCustomerName.AutoCompleteSource = AutoCompleteSource.CustomSource;

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



            AutoCompleteStringCollection col2 = new AutoCompleteStringCollection();
            col2.Clear();

            string sql2 = "SELECT CustomerName FROM Customer";
            SqlCommand cmd2 = new SqlCommand(sql2, conSS);
            SqlDataReader sdr2 = null;
            sdr2 = cmd2.ExecuteReader();
            while (sdr2.Read())
            {
                col2.Add(sdr2["CustomerName"].ToString());

            }
            sdr2.Close();
            textBoxCustomerName.AutoCompleteCustomSource = col2;

            conSS.Close();
        }
        private string Get_VAT()
        {
            string vats = "";
            SqlConnection conn = new SqlConnection(conStr);
            string query1 = "SELECT * FROM [CompanyInformation]";
            SqlCommand command1 = new SqlCommand(query1, conn);

            conn.Open();
            SqlDataReader reader1 = command1.ExecuteReader();

            while (reader1.Read())
            {

                vats = reader1["VatRate"].ToString();
            }
            reader1.Close();
            conn.Close();
            return vats;
        }
        private void Load_Form()
        {
            DateTime now = DateTime.Now;
            textBoxDate.Text = now.ToString("yyyy-MM-dd");
            textBoxDate.ReadOnly = true;
            textBoxTime.Text = now.ToLongTimeString();
            textBoxTime.ReadOnly = true;
            textBoxInvoiceNo.Text = now.ToLocalTime().ToString("yyyyMMddhhmmssfff");
           
            string svat = Get_VAT(); 
            textBoxVAT.Text = svat;
           
        }
        
        SqlConnection con = new SqlConnection(ConfigurationManager.ConnectionStrings["PosConString"].ToString());
        

        //public void DataGrid()
        //{
           
        //    con.Open();
        //    SqlCommand cmd = con.CreateCommand();
        //    cmd.CommandType = CommandType.Text;
        //    cmd.CommandText = "SELECT TempSalesProductID as 'Product CODE', Product.Name as 'Name', TempSalesSalePrice as 'Price', TempSalesQuantity as 'Quantity', TempSalesProductDiscount as 'Discount', TempSalesTotal as 'Total' FROM TempSales, Product WHERE TempSales.TempSalesCompanyID='" + CompanyID + "' AND TempSales.TempSalesProductID=Product.ID";
        //    cmd.ExecuteNonQuery();
        //    DataTable dt = new DataTable();
        //    SqlDataAdapter da = new SqlDataAdapter(cmd);
        //    da.Fill(dt);
        //    dataGridViewSales.DataSource = dt;
        //    con.Close();
        //}
        
        

        private void buttonClear_Click_1(object sender, EventArgs e)
        {
            textBoxProductName.Text = textBoxQuantity.Text = "";
            textBoxInvoiceNo.Text = textBoxPdoductCode.Text = "";
        }

        private void buttonClose_Click(object sender, EventArgs e)
        {
            FormPurchaseClosed();
            Close();
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
                textBoxUnitType.Text = reader12111["UnitName"].ToString();
            }
            reader12111.Close();
            connection12111.Close();
        }
       
        private void treeViewPurchaseItemSales_AfterSelect(object sender, TreeViewEventArgs e)
        {
            try
            {
                textBoxPdoductCode.Text = "";
                textBoxPdoductCode.Text = treeViewPurchaseItemSales.SelectedNode.Tag.ToString();
                textBoxProductName.Text = "";
                int pro_id = Convert.ToInt32(textBoxPdoductCode.Text);
                string conStr = ConfigurationManager.ConnectionStrings["PosConString"].ToString();
                SqlConnection connection12 = new SqlConnection(conStr);
                string query12 = "SELECT * FROM Product WHERE ID = '" + pro_id + "'";
                SqlCommand command112 = new SqlCommand(query12, connection12);

                connection12.Open();
                SqlDataReader reader12 = command112.ExecuteReader();

                while (reader12.Read())
                {

                    textBoxProductName.Text = reader12["Name"].ToString();
                  //  textBoxProCode.Text = reader12["Code"].ToString();
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
                    
                    int unit_id = Convert.ToInt32(reader1211["UnitID"]);
                    GetUnitName(unit_id);
                }
                reader1211.Close();



                string query13 = "SELECT CategorySub.SubCategoryName FROM CategorySub INNER JOIN Product ON CategorySub.SubCategoryID = Product.SubCategoryID WHERE Product.ID = '" + pro_id + "'";
                SqlCommand command113 = new SqlCommand(query13, connection12);

                connection12.Open();
                SqlDataReader reader13 = command113.ExecuteReader();

                while (reader13.Read())
                {

                    // textBoxProductName.Text = reader12["Name"].ToString();
                    textBoxProCate.Text = reader13["SubCategoryName"].ToString();
                    // textBoxPrice.Text = reader12["SalePrice"].ToString();
                    // int unit_id = Convert.ToInt32(reader12["UnitID"]);
                    //  GetUnitName(unit_id);
                }
                reader13.Close();

                connection1211.Close();

                //See Currently Stock
                textBoxCurrentStock.Text = Currently_Stock().ToString();
                textBoxProductSearch.Text = "";

            }
            catch (Exception)
            {
                textBoxPdoductCode.Text = "";
                textBoxProductName.Text = "";
                textBoxProCate.Text = "";
               
                textBoxUnitType.Text = "";
                MessageBox.Show("Please Select the Product First..");
            }
        }
        double stock = 0;
        private double Currently_Stock()
        {
            string conStrPross = ConfigurationManager.ConnectionStrings["PosConString"].ToString();
            SqlConnection connectionPross = new SqlConnection(conStrPross);
            string queryPross = "SELECT * FROM Product WHERE ID = '" + Convert.ToInt32(textBoxPdoductCode.Text) + "'";
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
        public void FormPurchaseClosed()
        {
            //Temp_Sales_Truncate();
        }

        //Add Button Event


        private double previous_stock(int id)
        {
            double previous_stock = 0;

            string conStrCal1 = ConfigurationManager.ConnectionStrings["PosConString"].ToString();
            SqlConnection connectionCal1 = new SqlConnection(conStrCal1);
            string queryCal1 = "SELECT * FROM Product WHERE ID = '" + id+"'";
            SqlCommand commandCal1 = new SqlCommand(queryCal1, connectionCal1);
            connectionCal1.Open();
            SqlDataReader readerCal1 = commandCal1.ExecuteReader();

            while (readerCal1.Read())
            {
                previous_stock = Convert.ToDouble(readerCal1["Stock"]);
            }
            readerCal1.Close();
            connectionCal1.Close();
            return previous_stock;
        }
        private double SalesTotal()
        {
            double total = 0;
            double p_stock = 0;
            SqlConnection connection = new SqlConnection(conStr);
            string query = "SELECT * FROM TempSales WHERE TempSalesCompanyID='"+CompanyID+"'";
            SqlCommand command = new SqlCommand(query, connection);
            connection.Open();
            SqlDataReader reader = command.ExecuteReader();

            while (reader.Read())
            {
                total = total + Convert.ToDouble(reader["TempSalesTotal"]);
                p_stock = previous_stock(Convert.ToInt32(reader["TempSalesProductID"]));
                UpdateProductDetails(Convert.ToInt32(reader["TempSalesProductID"]), p_stock-Convert.ToDouble(reader["TempSalesQuantity"]));
            }
            reader.Close();
            connection.Close();
            return total;

        }
        private int GetCustId(string name)
        {
            int id = 0;
            SqlConnection connection = new SqlConnection(conStr);
            string query = "SELECT * FROM Customer Where CustomerName ='"+name+"'";
            SqlCommand command = new SqlCommand(query, connection);
            connection.Open();
            SqlDataReader reader = command.ExecuteReader();
            while (reader.Read())
            {
                id = Convert.ToInt32(reader["CustomerID"]);
            }
            reader.Close();
            connection.Close();
            return id;

        }

        public void TempSales(string temp_product_code,string name, double qantity, double temp_sale_price, double temp_purchase_price, string invoice_no, double discount, double vat, double temp_sales_with_vat, double price_total, int temp_pro_id)
        {
            string Custname = textBoxCustomerName.Text;
            int sales_cust_id= GetCustId(Custname);
            string date = textBoxDate.Text;
            SqlConnection connection = new SqlConnection(conStr);
            string query = "INSERT INTO TempSales(TempSalesCompanyID,TempSalesNo,TempSalesDate,TempSaleTime,TempSalesCustomerID,TempSalesRemarks,TempSalesProductID,TempSalesPurchasePrice,TempSalesSalePrice,TempSalesQuantity,TempSalesProductDiscount,TempSalesTotal,TempSalesCustomerName,TempSalesSoldBy, TempSalesVatRate, TempSalesVatTotal, TempSalesPuechaseBy, TempSalesPurchaseByContact) VALUES('" + CompanyID+"','" + invoice_no + "','" + date + "','" + textBoxTime.Text + "','" + sales_cust_id + "','" + textBoxRemarks.Text + "','" + textBoxPdoductCode.Text + "','" + temp_purchase_price + "','" + temp_sale_price + "','" + qantity + "','" + discount + "','" + price_total + "','" + textBoxCustomerName.Text + "','"+ currentuser + "', '"+ vat +"', '" + temp_sales_with_vat + "', '" + textBoxPurchaseBy.Text + "', '" + textBoxPruchaseByContact.Text + "')";
            SqlCommand command = new SqlCommand(query, connection);
            connection.Open();
            command.ExecuteNonQuery();
            connection.Close();
        }

        private void UpdateProductDetails(int pro_id, double quantity)
        {
            SqlConnection connection1 = new SqlConnection(conStr);
            string query1 = "UPDATE Product SET Stock = '" + quantity + "' WHERE ID = '" + pro_id + "'";
            SqlCommand command1 = new SqlCommand(query1, connection1);
            connection1.Open();
            command1.ExecuteNonQuery();
            connection1.Close();
           
        }
        private void SeeTotal_Component()
        {
            double total = 0;
            double total_product = 0;
            double total_item = 0;
            double total_dis = 0;

            SqlConnection connection = new SqlConnection(conStr);
            string query = "SELECT * FROM TempSales WHERE TempSalesCompanyID='" + CompanyID + "'";
            SqlCommand command = new SqlCommand(query, connection);
            connection.Open();
            SqlDataReader reader = command.ExecuteReader();

            while (reader.Read())
            {
                total_product = total_product + Convert.ToDouble(reader["TempSalesQuantity"]);
                total = total + (Convert.ToDouble(reader["TempSalesTotal"])+ Convert.ToDouble(reader["TempSalesProductDiscount"])); //Change
                total_dis = total_dis + Convert.ToDouble(reader["TempSalesProductDiscount"]);
                total_item++;
            }
            reader.Close();
            connection.Close();
            textBoxItemTotal.Text = total_item.ToString();
            textBoxProductTotal.Text = total_product.ToString();
          

          //  textBoxNetTotal.Text= (total - total_dis).ToString();


            //textBoxInvoiceTotalAmount.Text = Math.Ceiling(total - total_dis).ToString();
            textBoxInvoiceTotalAmount.Text = Math.Ceiling(total).ToString();

        }
        private double Add_Already_Temp(int id)
        {
            double exist = 0.0;
            SqlConnection con = new SqlConnection(conStr);
            string query = "SELECT * FROM TempSales WHERE TempSalesCompanyID='" + CompanyID + "' AND TempSalesProductID='" + id+"'";
            SqlCommand command = new SqlCommand(query, con);
            con.Open();
            SqlDataReader reader = command.ExecuteReader();
           
            while (reader.Read())
            {
                exist = Convert.ToDouble(reader["TempSalesQuantity"]);
            }

            reader.Close();
            con.Close();
            return exist;
        }

        private double Discount_Already_Temp(int id)
        {
            double discount = 0.0;
            SqlConnection con = new SqlConnection(conStr);
            string query = "SELECT * FROM TempSales WHERE TempSalesCompanyID='" + CompanyID + "' AND TempSalesProductID='" + id + "'";
            SqlCommand command = new SqlCommand(query, con);
            con.Open();
            SqlDataReader reader = command.ExecuteReader();

            while (reader.Read())
            {
                discount = Convert.ToDouble(reader["TempSalesProductDiscount"]);
            }

            reader.Close();
            con.Close();
            return discount;
        }
        private void Update_Temp_sales(double quantity, double total, int temp_pro_id, double temp_discount)
        {
            SqlConnection connection1 = new SqlConnection(conStr);
            string query1 = "UPDATE TempSales SET TempSalesQuantity = '" + quantity + "',TempSalesProductDiscount = '" + temp_discount + "', TempSalesTotal = '" + total + "'  WHERE TempSalesCompanyID = '"+CompanyID+"' AND TempSalesProductID = '" + temp_pro_id + "'";
            SqlCommand command1 = new SqlCommand(query1, connection1);
            connection1.Open();
            command1.ExecuteNonQuery();
            connection1.Close();
        }
        private double GetPurchasePrice(int id)
        {
            double purchase_price = 0.0;
            SqlConnection con = new SqlConnection(conStr);
            string query = "SELECT PurchasePrice FROM Product WHERE ID= '" + id + "'";
            SqlCommand command = new SqlCommand(query, con);
            con.Open();
            SqlDataReader reader = command.ExecuteReader();

            while (reader.Read())
            {
                purchase_price = Convert.ToDouble(reader["PurchasePrice"]);
            }

            reader.Close();
            con.Close();
            return purchase_price;

        }
        private double AlreadyProductStock(int id)
        {
            double Quantity = 0.0;
            SqlConnection con = new SqlConnection(conStr);
            string query = "SELECT * FROM TempSales WHERE TempSalesCompanyID='" + CompanyID + "' AND TempSalesProductID='" + id + "'";
            SqlCommand command = new SqlCommand(query, con);
            con.Open();
            SqlDataReader reader = command.ExecuteReader();

            while (reader.Read())
            {
                Quantity = Convert.ToDouble(reader["TempSalesQuantity"]);
            }

            reader.Close();
            con.Close();
            return Quantity;
        }
        private void buttonAdd_Click(object sender, EventArgs e)
        {
            double ExitQuantity = AlreadyProductStock(Convert.ToInt32(textBoxPdoductCode.Text));
            if (textBoxPdoductCode.Text == "")
            {
                MessageBox.Show("Please Select a product..");
            }
           
            else if(textBoxQuantity.Text == "")
            {
                MessageBox.Show("Please Fill the Quantity..");
            }
            else if (Convert.ToInt32( textBoxQuantity.Text) == 0)
            {
                MessageBox.Show("Quantity is not zero..");
            }
           

            else if ((Convert.ToDouble(textBoxQuantity.Text)+ ExitQuantity) > Currently_Stock())
            {
                MessageBox.Show("Out of Stock!!!..\n Currently Stock = " + textBoxCurrentStock.Text + " only.....!!!!!!");
            }

            else if (textBoxCustomerName.Text == "")
            {
                MessageBox.Show("Please Select A Customer....???");
            }

            else
            {
                double temp_quantity = Convert.ToDouble(textBoxQuantity.Text);
               
                double temp_discount = 0;
                string temp_name = textBoxProductName.Text;
                string temp_invoice_no = textBoxInvoiceNo.Text;
                string temp_product_code = textBoxProCate.Text;
                
                int temp_pro_id = Convert.ToInt32(textBoxPdoductCode.Text);
                double temp_purchase_price = GetPurchasePrice(Convert.ToInt32(textBoxPdoductCode.Text));
                double alredy = Add_Already_Temp(temp_pro_id);
                double alredy_discount = Discount_Already_Temp(temp_pro_id);
                double vat = Convert.ToDouble(textBoxVAT.Text);
              
                if (alredy > 0.0)
                {
                    double Quantity_total = alredy + temp_quantity;
                    double total_discount = temp_discount + alredy_discount;
                }
                SeeTotal_Component();
             
                //DataGrid();
                
                
                textBoxProductName.Text = textBoxPdoductCode.Text = textBoxUnitType.Text = textBoxProCate.Text = "";
                textBoxQuantity.Text = textBoxProductSearch.Text = textBoxCurrentStock.Text = "";

            



               // textBoxDiscountPercent.Text = textBoxDiscountTaka.Text = "0";
                this.ActiveControl = textBoxProductSearch;
                buttonReports.Enabled = true;

            }
        }

        private void SalesClosed(object sender, FormClosedEventArgs e)
        {
            FormPurchaseClosed();
        }

        private void QuantityKeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar >= '0' && e.KeyChar <= '9' || e.KeyChar == '.' || e.KeyChar == (char)Keys.Back) //The  character represents a backspace
            {
                e.Handled = false; //Do not reject the input
            }
            else
            {
                e.Handled = true; //Reject the input
            }
        }

        private void DiscountKeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar >= '0' && e.KeyChar <= '9' || e.KeyChar == '.' || e.KeyChar == (char)Keys.Back) //The  character represents a backspace
            {
                e.Handled = false; //Do not reject the input
            }
            else
            {
                e.Handled = true; //Reject the input
            }
        }


        private void DiscountPercent_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar >= '0' && e.KeyChar <= '9' || e.KeyChar=='.' || e.KeyChar == (char)Keys.Back) //The  character represents a backspace
            {
                e.Handled = false; //Do not reject the input
            }
            else
            {
                e.Handled = true; //Reject the input
            }

        }
        private void InsertSales()
        {
           SalesTotal();
            
        }
       int PaymentType = 0;
        private void AddSalesDetails()
        {
            SqlConnection con = new SqlConnection(conStr);
            con.Open();
            string sql = @"INSERT INTO Sales(CompanyID,SalesNo,SalesDate,SalesTime,SalesCustomerID,SalesRemarks,Reference,SalesProductID,SalesPurchasePrice,SalesSalePrice,SalesQuantity,SalesProductDiscount,SalesTotal,SalesCustomerName,SalesSoldBy,SalesReceivedAmount,SalesChangeAmount,SalesVatRate,SalesVatTotal,SalesPuechaseBy,SalesPurchaseByContact) SELECT TempSalesCompanyID,TempSalesNo,TempSalesDate,TempSaleTime,TempSalesCustomerID,TempSalesRemarks,TempSalesReference,TempSalesProductID,TempSalesPurchasePrice,TempSalesSalePrice,TempSalesQuantity,TempSalesProductDiscount,TempSalesTotal,TempSalesCustomerName,TempSalesSoldBy,TempSalesReceivedAmount,TempSalesChangeAmount,TempSalesVatRate,TempSalesVatTotal, TempSalesPuechaseBy, TempSalesPurchaseByContact FROM TempSales  WHERE TempSalesCompanyID='" + CompanyID + "'";
            SqlCommand cmd = new SqlCommand(sql, con);
            cmd.ExecuteNonQuery();
            con.Close();
            
            

            //int IsChequeReceive = 0;
            //if (checkBoxChequeReceive.Checked==true)
            //{
            //    IsChequeReceive = 1;
            //}

            con.Open();
            sql = @"UPDATE Sales SET SalesRemarks ='" + textBoxRemarks.Text + "', PaymentType = '" + PaymentType + "'  WHERE SalesNo='" + SerialNo + "'";
            cmd = new SqlCommand(sql, con);
            cmd.ExecuteNonQuery();
            con.Close();

        }
        public void TempAmountTruncate()
        {
            SqlConnection con = new SqlConnection(conStr);
            con.Open();
            string sql = @"DELETE FROM TempSalesAmount WHERE CompanyID='"+CompanyID+"'";
            SqlCommand cmd = new SqlCommand(sql, con);
            cmd.ExecuteNonQuery();
            con.Close();
        }
         
        private void InsertTempAmount(int Companyid, int CustomerID, double TotalVat,double NetPayable,double Cashpaid,double ReturnAmount,double DueAmount,string  CurrentUserSales, string Remarks)
        {
            TempAmountTruncate();
            SqlConnection con1 = new SqlConnection(conStr);
            con1.Open();
            string sql1 = @"INSERT INTO TempSalesAmount(CompanyID, CustomerID, TotalVat, NetPayable, CashPaid, ReturnAmount, DueAmount,CurrentUserSales,Remarks) VALUES('" + Companyid + "','" + CustomerID + "','" + TotalVat + "','" + NetPayable + "','" + Cashpaid + "','" + ReturnAmount + "','" + DueAmount + "','" + CurrentUserSales + "','" + Remarks + "')";
            SqlCommand cmd1 = new SqlCommand(sql1, con1);
            cmd1.ExecuteNonQuery();
            con1.Close();
        }

        public string ReportPaths = ReportPath.rPath;
        private void Print_report()
        {
            ReportDocument cryRpt = new ReportDocument();
            TableLogOnInfos crtableLogoninfos = new TableLogOnInfos();
            TableLogOnInfo crtableLogoninfo = new TableLogOnInfo();
            ConnectionInfo crConnectionInfo = new ConnectionInfo();
            
            string rPath = ReportPaths + "CrystalReportSalesReportInvoiceA4.rpt";
            cryRpt.Load(rPath);

            cryRpt.SetParameterValue("CompanyID", CompanyID);
           
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

            cryRpt.PrintOptions.PrinterName = "";
            cryRpt.PrintToPrinter(1, false, 0, 0);
            cryRpt.Close();
            cryRpt.Dispose();

        }
        private void Print_Challan_report()
        {
            ReportDocument cryRpt = new ReportDocument();
            TableLogOnInfos crtableLogoninfos = new TableLogOnInfos();
            TableLogOnInfo crtableLogoninfo = new TableLogOnInfo();
            ConnectionInfo crConnectionInfo = new ConnectionInfo();

            string rPath = ReportPaths + "CrystalReportSalesInvoiceChallanA4.rpt";
            cryRpt.Load(rPath);

            cryRpt.SetParameterValue("CompanyID", CompanyID);

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

            cryRpt.PrintOptions.PrinterName = "";
            cryRpt.PrintToPrinter(1, false, 0, 0);
            cryRpt.Close();
            cryRpt.Dispose();

        }

        public string SerialNo = "";
        private void UpdateTempSalesDiscount(int countsales)
        {
           

            SerialNo = SNo + countsales.ToString("D6");
           
            string Custname = textBoxCustomerName.Text;
            int id = GetCustId(Custname);
            SqlConnection con1 = new SqlConnection(conStr);
            con1.Open();
           // string sql1 = @"UPDATE TempSales SET TempSalesProductDiscount ='" + textBoxDisTotal.Text + "', TempSalesReference='" + textBoxReference.Text + "',TempSalesCustomerID='" + id+ "',TempSalesNo='" + SerialNo + "' WHERE TempSalesCompanyID = '" + CompanyID + "'";
            string sql1 = @"UPDATE TempSales SET TempSalesReference='" + textBoxReference.Text + "',TempSalesCustomerID='" + id + "',TempSalesNo='" + SerialNo + "', TempSalesPaymentType = '" + PaymentType +"' WHERE TempSalesCompanyID = '" + CompanyID + "'";    // Change
            SqlCommand cmd1 = new SqlCommand(sql1, con1);
            cmd1.ExecuteNonQuery();
            con1.Close();
        }


        private int CountSalesNo()
        {
            int no = 0;
            SqlConnection con = new SqlConnection(conStr);
            string query = "SELECT COUNT(DISTINCT(SalesNo)) as 'No' FROM Sales";
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
        private void buttonReports_Click(object sender, EventArgs e)
        {
          
            string Custname = textBoxCustomerName.Text;
            int id = GetCustId(Custname);
           

           if(textBoxInvoiceTotalAmount.Text == "")
            {
                MessageBox.Show("Please Add Product First....???");
            }
            
           
            else
            {
                //CountSalesNo
                int countsales= CountSalesNo();

                //UpdateTempsales
                UpdateTempSalesDiscount(countsales+1);
                //Add ReceivedAmount, ChangeAmount in TempSales
               
                //sales details table
                AddSalesDetails();
                
               
                //Insert Temp Amount
                string cust_name;
                cust_name= textBoxCustomerName.Text;

                //sales table
                InsertSales();

                // MessageBox.Show(cust_name);
                //string cust_name = comboBoxBuyerName.SelectedText;
                int Companyid = CompanyID;
                int CustomerID= id;
                string CurrentUserSales = currentuser;
                string Remarks = textBoxRemarks.Text;

               
              
                //customer ledger

                Insert_Customer_Ledger(id);

                if (checkBoxSalesWithoutPrint.Checked == false)
                {
                    Print_report();
                    Print_Challan_report();
                }
                 Load_Form();
                 //Temp_Sales_Truncate();
                 FormPurchaseClosed();
                 //DataGrid();

                 textBoxItemTotal.Text = "0";
                 textBoxProductTotal.Text = "0";
                
                 textBoxInvoiceTotalAmount.Text = "0";
                 
                 checkBoxSalesWithoutPrint.Checked = false;
                 SerialNo = "";
                 //Temp_Sales_Truncate();
                 textBoxRemarks.Text = textBoxPurchaseBy.Text = textBoxPruchaseByContact.Text = textBoxCustomerName.Text = textBoxGroupName.Text = textBoxCompanyAddress.Text = textBoxVatRegNo.Text = textBoxReference.Text = textBoxPdoductCode.Text = textBoxProCate.Text = string.Empty;
                 MessageBox.Show("Sales Successfully....!!!!");
                buttonReports.Enabled = false;

            }
        }
       
        private void Insert_Customer_Ledger(int id)
        {
            string date = textBoxDate.Text;
            string ledger_invoice_no = SerialNo;
            double ledger_credit = 0.0;
            string ledger_remarks = "";
            int zero = 0;
            int Adjustment = 0;
            string conStr = ConfigurationManager.ConnectionStrings["PosConString"].ToString();
            SqlConnection connection = new SqlConnection(conStr);
            string query11 = "INSERT INTO CustomerLedger(ReceiveDate,CustomerID,InvoiceNo,Debit,Credit,Adjustment,Remarks,NextPaymentDate,IsPreviousDue) VALUES('" + date + "','" + id + "','" + ledger_invoice_no + "','" + ledger_credit + "','"+ Adjustment + "','" + ledger_remarks + "','"+ zero + "')";
            SqlCommand command = new SqlCommand(query11, connection);
            connection.Open();
            command.ExecuteNonQuery();
            connection.Close();
           
        }

        private void Clear_all()
        {
            textBoxPdoductCode.Text = "";
            textBoxProductName.Text = "";
            textBoxQuantity.Text = "";
            textBoxUnitType.Text = "";
           
            //textBoxDiscountTaka.Text = "";
            textBoxProductSearch.Text = "";
            textBoxItemTotal.Text = "";
            textBoxProductTotal.Text = "";
           
            textBoxInvoiceTotalAmount.Text = "";
            //textBoxVAT.Text = "";
        }
        private void dataGridViewSales_DoubleClick_1(object sender, EventArgs e)
        {
            buttonAdd.Visible = false;
            Clear_all();
            textBoxPdoductCode.Text = dataGridViewSales.SelectedRows[0].Cells[0].Value.ToString();
            textBoxProductName.Text = dataGridViewSales.SelectedRows[0].Cells[1].Value.ToString();
           
            textBoxQuantity.Text = dataGridViewSales.SelectedRows[0].Cells[3].Value.ToString();
            
            //textBoxDiscountTaka.Text = dataGridViewSales.SelectedRows[0].Cells[4].Value.ToString();

            string conStr11 = ConfigurationManager.ConnectionStrings["PosConString"].ToString();
            SqlConnection connection1211 = new SqlConnection(conStr11);
            string query1211 = "SELECT * FROM Product WHERE Product.ID = '"+textBoxPdoductCode.Text+ "'";
            SqlCommand command11211 = new SqlCommand(query1211, connection1211);

            connection1211.Open();
            SqlDataReader reader1211 = command11211.ExecuteReader();

            while (reader1211.Read())
            {
                textBoxCurrentStock.Text = reader1211["Stock"].ToString();
                textBoxProductName.Text= reader1211["Name"].ToString();
                textBoxProCate.Text= reader1211["Code"].ToString();
                GetUnitName(Convert.ToInt32(reader1211["UnitID"])); 
            }
            reader1211.Close();
            connection1211.Close();

        }

       
        private void buttonUpdate_Click(object sender, EventArgs e)
        {
            try
            {
                double dis = 0.0;
                double total;
                total = (Convert.ToDouble(textBoxQuantity.Text));
                SqlConnection connection1 = new SqlConnection(conStr);
                string query1 = "UPDATE TempSales SET TempSalesQuantity = '" + textBoxQuantity.Text + "', TempSalesTotal = '" + total + "'  WHERE TempSalesProductID = '" + textBoxPdoductCode.Text + "' AND TempSalesCompanyID='"+CompanyID+"'";
                SqlCommand command1 = new SqlCommand(query1, connection1);
                connection1.Open();
                int rowEffict1 = command1.ExecuteNonQuery();
                connection1.Close();
                //DataGrid();
                SeeTotal_Component();

                double vat = Convert.ToDouble(textBoxVAT.Text);
               
                textBoxQuantity.Text = textBoxProductName.Text = textBoxProCate.Text = textBoxPdoductCode.Text = textBoxUnitType.Text = "";
             
                buttonAdd.Visible = true;

            }
            catch (Exception)
            {
                MessageBox.Show("Please Select a Product in DataGrid...");
            }
        }
        private void Temp_delete()
        {
            try
            {
                SqlConnection connection1 = new SqlConnection(conStr);
                string query1 = @"DELETE FROM TempSales WHERE TempSalesCompanyID='" + CompanyID + "' AND TempSalesProductID = '" + textBoxPdoductCode.Text + "'";
                SqlCommand command1 = new SqlCommand(query1, connection1);
                connection1.Open();
                int rowEffict1 = command1.ExecuteNonQuery();
                connection1.Close();
            }
            catch (Exception)
            {
                MessageBox.Show("Please Select a Product in DataGrid...");
            }
        }
        private void buttonDelete_Click(object sender, EventArgs e)
        {
            try
            {
                Temp_delete();
                //DataGrid();
                SeeTotal_Component();
                textBoxProCate.Text = textBoxProductName.Text = textBoxQuantity.Text = textBoxUnitType.Text = textBoxPdoductCode.Text = "";
            }
            catch (Exception)
            {
                MessageBox.Show("Please Select a Product in DataGrid...");
            }
        }
    
        private int IsProductExist(string proname)
        {
            int exist = 0;
            SqlConnection conww = new SqlConnection(conStr);
            conww.Open();
            string sqlww = "SELECT * FROM Product WHERE Name ='" + proname + "' OR Code ='" + proname + "'";
            SqlCommand cmdww = new SqlCommand(sqlww, conww);
            SqlDataReader sdrww = null;
            sdrww = cmdww.ExecuteReader();
            while (sdrww.Read())
            {
                exist = Convert.ToInt32(sdrww["ID"]);
            }
            sdrww.Close();
            conww.Close();
            return exist;
        }
        private void textBoxProductSearch_TextChanged(object sender, EventArgs e)
        {
            if (IsProductExist(textBoxProductSearch.Text)>0)
            {
                string spro_id = textBoxProductSearch.Text;
                SqlConnection conww = new SqlConnection(conStr);
                conww.Open();
                string sqlww = "SELECT * FROM Product WHERE Name ='" + textBoxProductSearch.Text + "' OR Code ='" + textBoxProductSearch.Text + "'";
                SqlCommand cmdww = new SqlCommand(sqlww, conww);
                SqlDataReader sdrww = null;
                sdrww = cmdww.ExecuteReader();
                while (sdrww.Read())
                {
                    spro_id = sdrww["ID"].ToString();
                }
                sdrww.Close();
                conww.Close();

                try
                {
                    textBoxPdoductCode.Text = "";
                    textBoxProductName.Text = "";
                    textBoxUnitType.Text = "";
                   
                    textBoxPdoductCode.Text = spro_id;
                    int pro_id = Convert.ToInt32(textBoxPdoductCode.Text);
                    textBoxProductName.Text = "";
                    SqlConnection connection12 = new SqlConnection(conStr);
                    string query12 = "SELECT * FROM Product WHERE ID = '" + pro_id + "'";
                    SqlCommand command112 = new SqlCommand(query12, connection12);

                    connection12.Open();
                    SqlDataReader reader12 = command112.ExecuteReader();

                    while (reader12.Read())
                    {

                        textBoxProductName.Text = reader12["Name"].ToString();
                       // textBoxProCode.Text = reader12["Code"].ToString();
                      int unit_id = Convert.ToInt32(reader12["UnitID"]);
                        GetUnitName(unit_id);
                    }
                    reader12.Close();
                    connection12.Close();




                    string query13 = "SELECT CategorySub.SubCategoryName FROM CategorySub INNER JOIN Product ON CategorySub.SubCategoryID = Product.SubCategoryID WHERE Product.ID = '" + pro_id + "'";
                    SqlCommand command113 = new SqlCommand(query13, connection12);

                    connection12.Open();
                    SqlDataReader reader13 = command113.ExecuteReader();

                    while (reader13.Read())
                    {

                       // textBoxProductName.Text = reader12["Name"].ToString();
                        textBoxProCate.Text = reader13["SubCategoryName"].ToString();
                       // textBoxPrice.Text = reader12["SalePrice"].ToString();
                       // int unit_id = Convert.ToInt32(reader12["UnitID"]);
                      //  GetUnitName(unit_id);
                    }
                    reader13.Close();
                    connection12.Close();




                    //See Currently Stock
                    textBoxCurrentStock.Text = Currently_Stock().ToString();
                    if (textBoxProCate.Text != "")
                    {
                        this.ActiveControl = textBoxQuantity;

                    }

                }
                catch (Exception)
                {
                    textBoxPdoductCode.Text = "";
                    textBoxProCate.Text = "";
                    textBoxProductName.Text = "";
                    
                    textBoxCurrentStock.Text = "";
                    textBoxUnitType.Text = "";
                }
            }
        }

        private void comboBoxBuyerName_SelectedIndexChanged(object sender, EventArgs e)
        {
           
            string Custname = textBoxCustomerName.Text;
            int id = GetCustId(Custname);

            double ledger_debit = 0.0;
            double ledger_credit = 0.0;
            string conStr = ConfigurationManager.ConnectionStrings["PosConString"].ToString();
            SqlConnection con = new SqlConnection(conStr);
            string query = "SELECT * FROM CustomerLedger WHERE CustomerID = " + id;
            SqlCommand command112 = new SqlCommand(query, con);
            con.Open();
            SqlDataReader reader12 = command112.ExecuteReader();
            while (reader12.Read())
            {

                ledger_debit = ledger_debit + Convert.ToDouble(reader12["Debit"]);
                ledger_credit = ledger_credit + Convert.ToDouble(reader12["Credit"]);

            }
            reader12.Close();
            con.Close();
           
            // Get Data from Customer Table
            
            con = new SqlConnection(conStr);
            query = "SELECT * FROM Customer WHERE CustomerID = " + id;
            command112 = new SqlCommand(query, con);
            con.Open();
            reader12 = command112.ExecuteReader();
            while (reader12.Read())
            {

                textBoxGroupName.Text = reader12["GroupName"].ToString();
                textBoxVatRegNo.Text = reader12["VatRegNo"].ToString();
                textBoxCompanyAddress.Text = reader12["Address"].ToString();

            }
            reader12.Close();
            con.Close();

            // End




        }

        private void buttonNewCustomer_Click(object sender, EventArgs e)
        {
            AddCustomer ac = new AddCustomer();
           ac.Show();
            
        }

        //private void comboBoxBuyerName_Click(object sender, EventArgs e)
        //{
        //    string query12 = "SELECT * FROM Customer";
        //    fillCombo(comboBoxBuyerName, query12, "CustomerName", "CustomerID");
        //}

        
        private void textBoxQuantity_KeyDown(object sender, KeyEventArgs e)
        {
           
            if (e.KeyCode == Keys.Enter && textBoxQuantity.Text!="")
            {
                this.ActiveControl = buttonAdd;
            }
        }

        

        private void Sales_KeyDown(object sender, KeyEventArgs e)
        {
           
        }



        private void textBoxProCode_TextChanged(object sender, EventArgs e)
        {
            if (textBoxProCate.Text != "")
            {
                buttonAdd.Enabled = true;
            }
        }


        private void textBoxReceiveAmount_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                this.ActiveControl = buttonReports;
            }
        }

        private void textBoxReceiveAmount_TextChanged(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar >= '0' && e.KeyChar <= '9' || e.KeyChar == '.' || e.KeyChar == (char)Keys.Back) //The  character represents a backspace
            {
                e.Handled = false; //Do not reject the input
            }
            else
            {
                e.Handled = true; //Reject the input
            }
        }

        private void textBoxReceiveAmount_CursorChanged(object sender, EventArgs e)
        {
           

        }


        private void textBoxRemarks_TextChanged(object sender, EventArgs e)
        {

        }

       

       

       

       
       
        private void Demo1PrintReport()
        {
            ReportDocument cryRpt = new ReportDocument();
            TableLogOnInfos crtableLogoninfos = new TableLogOnInfos();
            TableLogOnInfo crtableLogoninfo = new TableLogOnInfo();
            ConnectionInfo crConnectionInfo = new ConnectionInfo();

            ParameterFields myParameterFields = new ParameterFields();

            ParameterField myParameterField1 = new ParameterField();
            ParameterDiscreteValue myDiscreteValue1 = new ParameterDiscreteValue();


            string rPath = ReportPaths + "CrystalReportSalesReportInvoiceA4.rpt";
            cryRpt.Load(rPath);
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

            TextboxValue("CompanyID", CompanyID, myParameterField1, myDiscreteValue1, myParameterFields);

            crystalReportViewer1.ParameterFieldInfo = myParameterFields;
            crystalReportViewer1.Refresh();
            crystalReportViewer1.ReportSource = cryRpt;
        }
        private void Demo2PrintReport()
        {
            ReportDocument cryRpt = new ReportDocument();
            TableLogOnInfos crtableLogoninfos = new TableLogOnInfos();
            TableLogOnInfo crtableLogoninfo = new TableLogOnInfo();
            ConnectionInfo crConnectionInfo = new ConnectionInfo();

            ParameterFields myParameterFields = new ParameterFields();

            ParameterField myParameterField1 = new ParameterField();
            ParameterDiscreteValue myDiscreteValue1 = new ParameterDiscreteValue();

            string rPath = ReportPaths + "CrystalReportSalesInvoiceChallanA4.rpt";
            cryRpt.Load(rPath);
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


            TextboxValue("CompanyID", CompanyID, myParameterField1, myDiscreteValue1, myParameterFields);

            crystalReportViewer1.ParameterFieldInfo = myParameterFields;
            crystalReportViewer1.Refresh();
            crystalReportViewer1.ReportSource = cryRpt;
        }
        private void buttonSalesInvoice_Click(object sender, EventArgs e)
        {
            if (textBoxItemTotal.Text == "")
            {
                MessageBox.Show("Please Select a Customer to due sell....!!!");
            }

            else if (textBoxInvoiceTotalAmount.Text == "")
            {
                MessageBox.Show("Please Add Product First....???");
            }

            else
            {

                Demo1PrintReport();
            }
        }

        private void buttonChallan_Click(object sender, EventArgs e)
        {
                if (textBoxItemTotal.Text == "")
                {
                    MessageBox.Show("Please Select a Customer to due sell....!!!");
                }

                else if (textBoxInvoiceTotalAmount.Text == "")
                {
                    MessageBox.Show("Please Add Product First....???");
                }
                
                else
                {
                    Demo2PrintReport();
                }
        }

        private void textBoxCustomerName_TextChanged(object sender, EventArgs e)
        {
          
            string Custname = textBoxCustomerName.Text;
            int id = GetCustId(Custname);

            double ledger_debit = 0.0;
            double ledger_credit = 0.0;
            string conStr = ConfigurationManager.ConnectionStrings["PosConString"].ToString();
            SqlConnection con = new SqlConnection(conStr);
            string query = "SELECT * FROM CustomerLedger WHERE CustomerID = " + id;
            SqlCommand command112 = new SqlCommand(query, con);
            con.Open();
            SqlDataReader reader12 = command112.ExecuteReader();
            while (reader12.Read())
            {

                ledger_debit = ledger_debit + Convert.ToDouble(reader12["Debit"]);
                ledger_credit = ledger_credit + Convert.ToDouble(reader12["Credit"]);

            }
            reader12.Close();
            con.Close();
          
            // Get Data from Customer Table

            con = new SqlConnection(conStr);
            query = "SELECT * FROM Customer WHERE CustomerID = " + id;
            command112 = new SqlCommand(query, con);
            con.Open();
            reader12 = command112.ExecuteReader();
            while (reader12.Read())
            {

                textBoxGroupName.Text = reader12["GroupName"].ToString();
                textBoxVatRegNo.Text = reader12["VatRegNo"].ToString();
                textBoxCompanyAddress.Text = reader12["Address"].ToString();

            }
            reader12.Close();
            con.Close();
        }

        private void textBoxCustomerName_MouseClick(object sender, MouseEventArgs e)
        {
            Auto_Complete();
        }
    }
}
