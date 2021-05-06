using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using MISA.Import.APITest.Model;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Data;
using Dapper;
using MySqlConnector;
using System.Text.RegularExpressions;

namespace MISA.Import.APITest.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class CustomerController : ControllerBase
    {
        static IConfiguration _configuration;
        static IDbConnection dbConnection;
        public CustomerController(IConfiguration configuration)
        {
            _configuration = configuration;
        }


        [HttpPost("import")]
        public IActionResult Import(IFormFile formFile, CancellationToken cancellationToken)
        {
            if (formFile == null || formFile.Length <= 0)
            {
                return NoContent();
            }

            if (!Path.GetExtension(formFile.FileName).Equals(".xlsx", StringComparison.OrdinalIgnoreCase))
            {
                return NoContent();
            }
            // Khởi tạo 1 danh sách khách hàng
            var customers = new List<Customer>();

            using (var stream = new MemoryStream())
            {
                formFile.CopyToAsync(stream, cancellationToken);

                using (var package = new ExcelPackage(stream))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                    var rowCount = worksheet.Dimension.Rows;
                    // Duyệt từ dòng 3 của file excel
                    for (int row = 3; row <= rowCount; row++)
                    {
                        // Khởi tạo giá trị datetime.
                        DateTime dateReturn = DateTime.Now;
                        var dateString = "";
                        Regex dateValidRegex = new Regex(@"^([0]?[1-9]|[1|2][0-9]|[3][0|1])[./-]([0]?[1-9]|[1][0-2])[./-]([0-9]{4}|[0-9]{2})$");
                        Regex dateValidRegexMonthAndYear = new Regex(@"([0]?[1-9]|[1][0-2])[./-]([0-9]{4}|[0-9]{2})$");
                        Regex yearValidRegex = new Regex(@"([0-9]{4}|[0-9]{2})$");

                        // Kiểm tra null của cột DateTime DateOfBirth và gán giá trị
                        if (worksheet.Cells[row, 6].Value == null)
                        {
                            dateReturn = DateTime.Now;
                        }
                        else
                        {
                            // Gán giá trị cột DateOfBirth cho dateString
                            dateString = worksheet.Cells[row, 6].Value.ToString();
                            // Kiểm tra nếu khớp với regex Date dd-mm-yyyy thì format
                            if (dateValidRegex.IsMatch(dateString))
                            {
                                var dateSplit = dateString.Split(new string[] { "/", ".", "-" }, StringSplitOptions.None);
                                var day = int.Parse(dateSplit[0]);
                                var month = int.Parse(dateSplit[1]);
                                var year = int.Parse(dateSplit[2]);
                                dateReturn = new DateTime(year, month, day);
                            }
                            else if (dateValidRegexMonthAndYear.IsMatch(dateString))
                            {
                                var dateSplit = dateString.Split(new string[] { "/", ".", "-" }, StringSplitOptions.None);
                                var month = int.Parse(dateSplit[0]);
                                var year = int.Parse(dateSplit[1]);
                                dateReturn = new DateTime(year, month, 1);
                            }
                            // Kiểm tra nếu khớp với regex Date yyyy thì format
                            else if (yearValidRegex.IsMatch(dateString))
                            {
                                var year = int.Parse(dateString);
                                dateReturn = new DateTime(year, 1, 1);
                            }
                            // Kiểm tra nếu khớp với regex Date mm-yyyy thì format
                        }
                        // Thêm đối tượng customer vào danh sách
                        // Mặc định trạng thái status là hợp lệ
                        customers.Add(new Customer
                        {
                            CustomerId = new Guid(),
                            CustomerCode = (worksheet.Cells[row, 1].Value == null) ? "" : worksheet.Cells[row, 1].Value.ToString().Trim(),
                            FullName = (worksheet.Cells[row, 2].Value == null) ? "" : worksheet.Cells[row, 2].Value.ToString().Trim(),
                            MemberCardCode = (worksheet.Cells[row, 3].Value == null) ? "" : worksheet.Cells[row, 3].Value.ToString().Trim(),
                            CustomerGroupId = new Guid(),
                            CustomerGroupName = (worksheet.Cells[row, 4].Value == null) ? "" : worksheet.Cells[row, 4].Value.ToString().Trim(),
                            PhoneNumber = (worksheet.Cells[row, 5].Value == null) ? "" : worksheet.Cells[row, 5].Value.ToString().Trim(),
                            DateOfBirth = dateReturn,
                            CompanyName = (worksheet.Cells[row, 7].Value == null) ? "" : worksheet.Cells[row, 7].Value.ToString().Trim(),
                            CompanyTaxCode = (worksheet.Cells[row, 8].Value == null) ? "" : worksheet.Cells[row, 8].Value.ToString().Trim(),
                            Email = (worksheet.Cells[row, 9].Value == null) ? "" : worksheet.Cells[row, 9].Value.ToString().Trim(),
                            Address = (worksheet.Cells[row, 10].Value == null) ? "" : worksheet.Cells[row, 10].Value.ToString().Trim(),
                            Note = (worksheet.Cells[row, 11].Value == null) ? "" : worksheet.Cells[row, 11].Value.ToString().Trim(),
                            Status = "Hợp lệ"
                        });
                    }
                }
            }

            // Duyệt tất cả khách hàng trong danh sách khách hàng
            foreach (Customer customer in customers)
            {
                // Kiểm tra customerGroupName tồn tại hay không.
                var customerGroup = CheckCustomerGroupExist(customer);
                if (customerGroup == null)
                {
                    customer.CustomerGroupId = null;
                    customer.Status = "";
                    customer.Status += "Nhóm khách hàng không có trong hệ thống. ";
                }
                else
                {
                    customer.CustomerGroupId = customerGroup.CustomerGroupId;
                }
                // Kiểm tra tồn tại trong database
                if (CheckPropertyExistDB("CustomerCode", customer.CustomerCode))
                {
                    customer.Status = "";
                    customer.Status += "Mã khách hàng đã tồn tại trong hệ thống. ";
                }
                if (CheckPropertyExistDB("PhoneNumber", customer.PhoneNumber))
                {
                    customer.Status = "";
                    customer.Status += "SĐT đã có trong hệ thống. ";
                }
                
            }
            // Kiểm tra trùng mã code và số điện thoại trên File
            for (int i = customers.Count - 1; i > 0; i--)
            {
                for (int j = 0; j < i; j++)
                {
                    //Kiểm tra nếu đúng thì gán status rỗng sau đó gán.
                    if (customers[i].CustomerCode == customers[j].CustomerCode)
                    {
                        customers[i].Status = "";
                        customers[i].Status += "Mã khách hàng đã trùng với khách hàng khác trong tệp nhập khẩu\n";
                        break;
                    }
                    if (customers[i].PhoneNumber == customers[j].PhoneNumber)
                    {
                        customers[i].Status = "";
                        customers[i].Status += "SĐT đã trùng với SĐT của khách hàng khác trong tệp nhập khẩu.\n";
                        break;
                    }
                }
            }

            // insert danh sách khách hàng vào DB
            using (dbConnection = new MySqlConnection(_configuration.GetConnectionString("connectionDB")))
            {
                var sql = "Proc_InsertCustomer";
                var rowAffects = dbConnection.Execute(sql, customers, commandType: CommandType.StoredProcedure);
                if (rowAffects > 0)
                {
                    return Ok(customers);
                }
                return NoContent();
            }

        }

        /// <summary>
        /// Kiểm tra tồn tại trong DB
        /// </summary>
        /// <param name="property">property của đối tượng Customer</param>
        /// <param name="propertyValue">Giá trị của các property</param>
        /// <returns>true or false</returns>
        public static bool CheckPropertyExistDB(string property, string propertyValue)
        {
            using (dbConnection = new MySqlConnection(_configuration.GetConnectionString("connectionDB")))
            {
                string sqlCommand = $"Proc_Check{property}Exist";
                DynamicParameters dynamicParameters = new DynamicParameters();
                dynamicParameters.Add($"@{property}", propertyValue);
                var check = dbConnection.QueryFirstOrDefault<bool>(sqlCommand, param: dynamicParameters, commandType: CommandType.StoredProcedure);
                return check;
            }
        }

        /// <summary>
        /// Kiểm tra tên nhóm khách hàng
        /// </summary>
        /// <param name="customer">Khách hàng</param>
        /// <returns>nhóm khách hàng</returns>
        public static CustomerGroup CheckCustomerGroupExist(Customer customer)
        {
            using (dbConnection = new MySqlConnection(_configuration.GetConnectionString("connectionDB")))
            {
                string sqlCommand = "Proc_GetCustomerGroupByName";
                DynamicParameters dynamicParameters = new DynamicParameters();
                dynamicParameters.Add("@customerGroupName", customer.CustomerGroupName);
                var customerGroup = dbConnection.QueryFirstOrDefault<CustomerGroup>(sqlCommand, param: dynamicParameters, commandType: CommandType.StoredProcedure);
                return customerGroup;
            }
        }
    }
}
