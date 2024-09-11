using Coffee_Shop_Management_System.Models;
using Microsoft.AspNetCore.Mvc;
using System.Data.SqlClient;
using System.Data;
using OfficeOpenXml;

namespace Coffee_Shop_Management_System.Controllers
{
    public class OrderDetailsController : Controller
    {

        private readonly IConfiguration _configuration;

        public OrderDetailsController(IConfiguration configuration)
        {
            _configuration = configuration;
        }
        //public static List<OrderDetailsModel> orderdetails = new List<OrderDetailsModel>
        //{
        //    new OrderDetailsModel {OrderDetailsID=1,Amount=2000,Quantity=5,TotalAmount=300,OrderID=1,ProductID=1,UserID=1},
        //    new OrderDetailsModel {OrderDetailsID=2,Amount=3000,Quantity=6,TotalAmount=200,OrderID=2,ProductID=2,UserID=2},
        //};
        public IActionResult OrderDetailsList()
        {
            string connectionString = _configuration.GetConnectionString("ConnectionString");
            SqlConnection connection = new SqlConnection(connectionString);
            connection.Open();
            SqlCommand command = connection.CreateCommand();
            command.CommandType = System.Data.CommandType.StoredProcedure;
            command.CommandText = "PR_OrderDetail_SelectAll";
            SqlDataReader reader = command.ExecuteReader();
            DataTable table = new DataTable();
            table.Load(reader);
            return View(table);
        }
        public IActionResult AddOrderDetails(int OrderDetailID)
        {
            string connectionString = this._configuration.GetConnectionString("ConnectionString");

            #region User Drop-Down

            SqlConnection connection1 = new SqlConnection(connectionString);
            connection1.Open();
            SqlCommand command1 = connection1.CreateCommand();
            command1.CommandType = System.Data.CommandType.StoredProcedure;
            command1.CommandText = "PR_User_DropDown";
            SqlDataReader reader1 = command1.ExecuteReader();
            DataTable dataTable1 = new DataTable();
            dataTable1.Load(reader1);
            connection1.Close();

            List<UserDropDownModel> users = new List<UserDropDownModel>();

            foreach (DataRow dataRow in dataTable1.Rows)
            {
                UserDropDownModel userDropDownModel = new UserDropDownModel();
                userDropDownModel.UserID = Convert.ToInt32(dataRow["UserID"]);
                userDropDownModel.UserName = dataRow["UserName"].ToString();
                users.Add(userDropDownModel);
            }

            ViewBag.UserList = users;

            #endregion
            #region Product Drop-Down

            SqlConnection connection2 = new SqlConnection(connectionString);
            connection2.Open();
            SqlCommand command2 = connection2.CreateCommand();
            command2.CommandType = System.Data.CommandType.StoredProcedure;
            command2.CommandText = "PR_Product_DropDown";
            SqlDataReader reader2 = command2.ExecuteReader();
            DataTable dataTable2 = new DataTable();
            dataTable2.Load(reader2);
            connection2.Close();

            List<ProductDropDown> products = new List<ProductDropDown>();

            foreach (DataRow dataRow in dataTable2.Rows)
            {
                ProductDropDown productDropDownModel = new ProductDropDown();
                productDropDownModel.ProductID = Convert.ToInt32(dataRow["ProductID"]);
                productDropDownModel.ProductName = dataRow["ProductName"].ToString();
                products.Add(productDropDownModel);
            }

            ViewBag.ProductList = products;

            #endregion
            #region Order Drop-Down

            SqlConnection connection3 = new SqlConnection(connectionString);
            connection3.Open();
            SqlCommand command3 = connection3.CreateCommand();
            command3.CommandType = System.Data.CommandType.StoredProcedure;
            command3.CommandText = "PR_Order_DropDown";
            SqlDataReader reader3 = command3.ExecuteReader();
            DataTable dataTable3 = new DataTable();
            dataTable3.Load(reader3);
            connection3.Close();

            List<OrderDropDown> orders = new List<OrderDropDown>();

            foreach (DataRow dataRow in dataTable3.Rows)
            {
                OrderDropDown orderDropDownModel = new OrderDropDown();
                orderDropDownModel.OrderID = Convert.ToInt32(dataRow["OrderID"]);
                orderDropDownModel.OrderNumber = dataRow["OrderNumber"].ToString();
                orders.Add(orderDropDownModel);
            }

            ViewBag.OrderList = orders;

            #endregion

            #region OrderDetailsByID

            SqlConnection connection = new SqlConnection(connectionString);
            connection.Open();
            SqlCommand command = connection.CreateCommand();
            command.CommandType = CommandType.StoredProcedure;
            command.CommandText = "PR_OrderDetail_SelectByID";
            command.Parameters.AddWithValue("@OrderDetailID", OrderDetailID);
            SqlDataReader reader = command.ExecuteReader();
            DataTable table = new DataTable();
            table.Load(reader);
            OrderDetailsModel orderDetailsModel = new OrderDetailsModel();

            foreach (DataRow dataRow in table.Rows)
            {
                orderDetailsModel.Quantity = Convert.ToInt32(@dataRow["Quantity"]);
                orderDetailsModel.Amount = Convert.ToDecimal(@dataRow["Amount"]);
                orderDetailsModel.TotalAmount = Convert.ToDecimal(@dataRow["TotalAmount"]);
                orderDetailsModel.OrderID = Convert.ToInt32(@dataRow["ProductID"]);
                orderDetailsModel.ProductID = Convert.ToInt32(@dataRow["ProductID"]);
                orderDetailsModel.UserID = Convert.ToInt32(@dataRow["UserID"]);
            }

            #endregion

            return View("AddOrderDetails", orderDetailsModel);
        }
        public IActionResult OrderDetailsSave(OrderDetailsModel orderdetailsModel)
        {
            if (orderdetailsModel.UserID <= 0)
            {
                ModelState.AddModelError("UserID", "A valid User is required.");
            }

            if (ModelState.IsValid)
            {
                string connectionString = this._configuration.GetConnectionString("ConnectionString");
                SqlConnection connection = new SqlConnection(connectionString);
                connection.Open();
                SqlCommand command = connection.CreateCommand();
                command.CommandType = CommandType.StoredProcedure;
                if (orderdetailsModel.OrderDetailID == null)
                {
                    command.CommandText = "PR_OrderDetail_Insert";
                }
                else
                {
                    command.CommandText = "PR_OrderDetail_Update";
                    command.Parameters.Add("@OrderDetailID", SqlDbType.Int).Value = orderdetailsModel.OrderDetailID;
                }
                command.Parameters.Add("@Quantity", SqlDbType.VarChar).Value = orderdetailsModel.Quantity;
                command.Parameters.Add("@Amount", SqlDbType.VarChar).Value = orderdetailsModel.Amount;
                command.Parameters.Add("@TotalAmount", SqlDbType.Decimal).Value = orderdetailsModel.TotalAmount;
                command.Parameters.Add("@OrderID", SqlDbType.Int).Value = orderdetailsModel.OrderID;
                command.Parameters.Add("@ProductID", SqlDbType.Int).Value = orderdetailsModel.ProductID;
                command.Parameters.Add("@UserID", SqlDbType.Int).Value = orderdetailsModel.UserID;
                command.ExecuteNonQuery();
                return RedirectToAction("OrderDetailsList");
            }

            return View("AddOrderDetails", orderdetailsModel);
        }
        public IActionResult DeleteOrderDetail(int OrderDetailID)
        {
            string connectionString = this._configuration.GetConnectionString("ConnectionString");
            SqlConnection connection = new SqlConnection(connectionString);
            connection.Open();
            SqlCommand command = connection.CreateCommand();
            command.CommandType = CommandType.StoredProcedure;
            command.CommandText = "PR_OrderDetail_Delete";
            command.Parameters.Add("@OrderDetailID", SqlDbType.Int).Value = OrderDetailID;
            command.ExecuteNonQuery();
            return RedirectToAction("OrderDetailsList");
        }
        public IActionResult ExportToExcel()
        {
            // Fetch the product data
            string connectionString = _configuration.GetConnectionString("ConnectionString");
            SqlConnection connection = new SqlConnection(connectionString);
            connection.Open();
            SqlCommand command = connection.CreateCommand();
            command.CommandType = CommandType.StoredProcedure;
            command.CommandText = "PR_OrderDetail_SelectAll";
            SqlDataReader reader = command.ExecuteReader();
            DataTable table = new DataTable();
            table.Load(reader);

            // Create the Excel file in memory
            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("OrderDetails");

                // Load the data from the DataTable into the worksheet, starting from cell A1.
                worksheet.Cells["A1"].LoadFromDataTable(table, true);

                // Format the header row
                using (var range = worksheet.Cells["A1:Z1"])
                {
                    range.Style.Font.Bold = true;
                    range.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightBlue);
                }

                // Convert the package to a byte array
                var excelData = package.GetAsByteArray();

                // Return the Excel file as a download
                return File(excelData, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "OrderDetailList.xlsx");
            }
        }
    }
}
