using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using BIG;
using System.Configuration;

namespace OrderDiscount
{
    public class Events
    {
        internal BIG.Application bigApp; //?

        public static Configuration appConfig = ConfigurationManager.OpenExeConfiguration(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "OrderDiscount.dll"));

        public static string discountProductNo = appConfig.AppSettings.Settings["DiscountProductNo"].Value;

        public static int discountUnit = Convert.ToInt32(appConfig.AppSettings.Settings["DiscountUnit"].Value);

        public static string connectionString = appConfig.ConnectionStrings.ConnectionStrings["vbConnection"].ConnectionString;

        public Events(BIG.Application bigApplication) //?  EV54545
        {
            this.bigApp = bigApplication; //? EV54545

            this.bigApp.SessionLoggedOn += SessionLoggedOn;  // Egne metoder som knyttes til hendelser i Visma Business

            this.bigApp.SessionLoggingOff += SessionLoggingOff; 

            this.bigApp.ButtonClick += ButtonClick;

        }

        private void ButtonClick(BIG.Button button, ref bool SkipRecording) 
        {
            try
            {
                if (button.Caption.ToLower() == "rabatt")
                {
                    int orderNo = 0;

                    BIG.PageElement pageElement = button.PageElement;

                    BIG.Document bigDocument = pageElement.Document;

                    foreach (BIG.Table table in bigDocument.Tables)
                    {
                        //string title = table.Title.ToLower();

                        if (table.TableNo == (int)BIG.TableNo.TableNo_Order)  // If table.TableNo = 127 ...
                        {
                            string tableName = table.Title;

                            Ts_RowGetValue ts_RowGetValue;

                            orderNo = table.ActiveRow.GetIntegerValue((int)C_Order.C_Order_OrderNo, out ts_RowGetValue);

                            if (orderNo > 0)
                            {
                                bigDocument.Save();

                                double discount = GetDiscount(orderNo, discountProductNo);

                                if (discount > 0)
                                {
                                    bool result = AddOrderLineDiscountRow(bigDocument, table, orderNo, discountProductNo, discount * -1);
                                }

                                break;

                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
            }
        }

        private bool AddOrderLineDiscountRow(Document bigDocument, Table orderTable, int orderNo, string productNo, double discount)
        {
            Ts_TableJoin ts_TableJoin; 

            Ts_TableGetEmptyRow ts_TableGetEmptyRow; 

            string message = "";

            var newOrderLineTable = orderTable.Join((int)Fk_OrderLine.Fk_OrderLine_Order, out ts_TableJoin);  // Join orderlinetable

            var newOrderLineRow = newOrderLineTable.GetEmptyRow(out ts_TableGetEmptyRow);  // Get a empty row in orderlinetable

            newOrderLineRow.SuggestValue((int)C_OrderLine.C_OrderLine_LineNo, out message); // suggest a orderline linenumber
            newOrderLineRow.SetStringValue((int)C_OrderLine.C_OrderLine_ProductNo, productNo, out message); // set productNo
            newOrderLineRow.SetDecimalValue((int)C_OrderLine.C_OrderLine_Quantity, 1, out message); // Set qauantity
            newOrderLineRow.SetDecimalValue((int)C_OrderLine.C_OrderLine_Price, discount, out message); // set price (amount)


            bigDocument.Save();
            bigDocument.Refresh();

            return true;
        }

        public static double GetDiscount(int orderNo, string productNo)
        {
            double discount = 0;

            string commandText = @"
                delete from OrdLn where OrdNo = @OrderNo and ProdNo = @ProductNo;
    
                declare @TotalKg decimal = (select sum(NoInvoAb) from OrdLn where OrdNo = @OrderNo and Un = @DiscountUnit) 

                select
	                case when @TotalKg between 3000 and 6999 then @TotalKg * 0.10
		                    when @TotalKg between 7000 and 19999 then @TotalKg * 0.35
		                    when @TotalKg between 20000 and 49999 then @TotalKg * 0.55
		                    when @TotalKg between 50000 and 74999 then @TotalKg * 0.75
		                    when @TotalKg between 75000 and 99999 then @TotalKg * 0.80
		                    when @TotalKg between 100000 and 149999 then @TotalKg * 1
		                    when @TotalKg between 150000 and 199999 then @TotalKg * 1.05
		                    when @TotalKg > 200000 then @TotalKg * 1.10
		                    else 0
	                    end as TotalDiscount
                ";

            SqlCommand sqlCommand = new SqlCommand();

            sqlCommand.CommandText = commandText;
            sqlCommand.Parameters.AddWithValue("OrderNo", orderNo);
            sqlCommand.Parameters.AddWithValue("ProductNo", productNo);
            sqlCommand.Parameters.AddWithValue("DiscountUnit", discountUnit);

            DataTable dataTable = SelectData(sqlCommand); 

            if (dataTable.Rows.Count > 0) 
            {
                discount = Convert.ToDouble(dataTable.Rows[0]["TotalDiscount"]);
            }

            return discount;
        }

        static DataTable SelectData(SqlCommand sqlCommand)
        {
            DataTable dataTable = new DataTable();

            using (SqlConnection sqlConnection = new SqlConnection(connectionString))
            {
                sqlConnection.Open();

                using (sqlCommand) //?
                {
                    sqlCommand.Connection = sqlConnection; //?

                    dataTable.Load(sqlCommand.ExecuteReader());  //?
                }
            }

            return dataTable;
        }

        // --------------------------

        private void SessionLoggingOff(ref bool Cancel)
        {
            BIG.Ts_ReadOnlyRowStringValue ts;
            string username = bigApp.User.GetStringValue((int)C_User.C_User_UserName, out ts);

            if (username == "Tobias")
            {
                Cancel = true;
            }

            Log(username + " logged off.");
        }

        private void SessionLoggedOn()
        {
            // Logged on user
            BIG.Ts_ReadOnlyRowStringValue ts;
            string username = bigApp.User.GetStringValue((int)C_User.C_User_UserName, out ts);

            Log(username + " logged on.");
        }

        public static void DebugLog(string message)
        {
            StreamWriter sw = null;

            try
            {
                sw = new StreamWriter("DiscountDebugLog.txt", true);
                sw.WriteLine(DateTime.Now.ToString() + ": " + message);
                sw.Flush();
                sw.Close();
            }
            catch
            {
            }
        }

        public static void Log(string message)
        {
            StreamWriter sw = null;

            try
            {
                sw = new StreamWriter("DiscountLog.txt", true);
                sw.WriteLine(DateTime.Now.ToString() + ": " + message);
                sw.Flush();
                sw.Close();
            }
            catch
            {
            }
        }

    }
}
