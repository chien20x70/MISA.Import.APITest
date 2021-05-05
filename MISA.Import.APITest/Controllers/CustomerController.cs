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

namespace MISA.Import.APITest.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class CustomerController : ControllerBase
    {
        IConfiguration _configuration;
        public CustomerController(IConfiguration configuration)
        {
            _configuration = configuration;
        }
        [HttpPost("import")]
        public async Task<IActionResult> Import(IFormFile formFile, CancellationToken cancellationToken)
        {
            if (formFile == null || formFile.Length <= 0)
            {
                return NoContent();
            }

            if (!Path.GetExtension(formFile.FileName).Equals(".xlsx", StringComparison.OrdinalIgnoreCase))
            {
                return NoContent();
            }

            var list = new List<Customer>();

            using (var stream = new MemoryStream())
            {
                await formFile.CopyToAsync(stream, cancellationToken);

                using (var package = new ExcelPackage(stream))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                    var rowCount = worksheet.Dimension.Rows; 
                    for (int row = 3; row <= rowCount; row++)
                    {
                        //var check = DateTime.TryParse(worksheet.Cells[row, 6].Value.ToString(), out DateTime dt);
                        list.Add(new Customer
                        {
                            CustomerCode = (worksheet.Cells[row, 1].Value == null) ? "" : worksheet.Cells[row, 1].Value.ToString().Trim(),
                            FullName = (worksheet.Cells[row, 2].Value == null) ? "" : worksheet.Cells[row, 2].Value.ToString().Trim(),
                            MemberCardCode = (worksheet.Cells[row, 3].Value == null) ? "" : worksheet.Cells[row, 3].Value.ToString().Trim(),
                            CustomerGroupName = (worksheet.Cells[row, 4].Value == null) ? "" : worksheet.Cells[row, 4].Value.ToString().Trim(),
                            PhoneNumber = (worksheet.Cells[row, 5].Value == null) ? "" : worksheet.Cells[row, 5].Value.ToString().Trim(),
                            DateOfBirth = (worksheet.Cells[row, 6].Value == null) ? "" : worksheet.Cells[row, 6].Value.ToString().Trim(),
                            //DateOfBirth = (!check && worksheet.Cells[row, 6].Value == null) ? DateTime.Now : dt,
                            CompanyName = (worksheet.Cells[row, 7].Value == null) ? "" : worksheet.Cells[row, 7].Value.ToString().Trim(),
                            CompanyTaxCode = (worksheet.Cells[row, 8].Value == null) ? "" : worksheet.Cells[row, 8].Value.ToString().Trim(),
                            Email = (worksheet.Cells[row, 9].Value == null) ? "" : worksheet.Cells[row, 9].Value.ToString().Trim(),
                            Address = (worksheet.Cells[row, 10].Value == null) ? "" : worksheet.Cells[row, 10].Value.ToString().Trim(),
                            Note = (worksheet.Cells[row, 11].Value == null) ? "" : worksheet.Cells[row, 11].Value.ToString().Trim(),
                        });
                    }
                }
            }

            IDbConnection dbConnection = new MySqlConnection(_configuration.GetConnectionString("connectionDB"));
            var sql = "Proc_InsertCustomer";
            var rowAffects = dbConnection.Execute(sql, list, commandType: CommandType.StoredProcedure);
            if(rowAffects > 0)
            {
                return Ok(list);
            }
            return NoContent();
        }
    }
}
