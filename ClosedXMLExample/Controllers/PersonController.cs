using ClosedXML.Excel;
using Microsoft.AspNetCore.Mvc;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;

namespace ClosedXMLExample.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class PersonController : ControllerBase
    {
        [Route("export-from-table")]
        [HttpGet]
        public IActionResult ExportFromTable()
        {
            DataTable dt = getData();
            //Name of File  
            string fileName = "ListPerson.xlsx";
            using (XLWorkbook wb = new XLWorkbook())
            {
                //Add DataTable in worksheet  
                wb.Worksheets.Add(dt);
                using (MemoryStream stream = new MemoryStream())
                {
                    wb.SaveAs(stream);
                    //Return xlsx Excel File  
                    return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
                }
            }
        }

        [Route("export-from-list")]
        [HttpGet]
        public IActionResult ExportFromList(string format)
        {
            var productList = GetProductList();

            if (format == "csv")
            {
                var csv = new StringBuilder();
                string line = $"{nameof(Product.Name)},{nameof(Product.Price)},{nameof(Product.Description)}";
                csv.AppendLine(line);
                foreach (var product in productList)
                {
                    string productName = "";
                    string productDescription = "";
                    if (product.Name.Contains(",")){
                        productName = String.Format("\"{0}\"", product.Name);
                    }
                    else
                    {
                        productName = product.Name;
                    }
                    if (product.Description.Contains(","))
                    {
                        productDescription = String.Format("\"{0}\"", product.Description);
                    }
                    else
                    {
                        productDescription = product.Description;
                    }
                    line = string.Format("{0},{1},{2}", productName, product.Price, productDescription);
                    csv.AppendLine(line);
                }
                return File(Encoding.ASCII.GetBytes(csv.ToString()), "text/csv", "ProductList.csv");
            }
            else
            {
                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add("Products");
                    worksheet.Cell("A1").Value = nameof(Product.Name);
                    worksheet.Cell("B1").Value = nameof(Product.Price);
                    worksheet.Cell("C1").Value = nameof(Product.Description);

                    int tableRow = 1;
                    foreach (var product in productList)
                    {
                        tableRow++;
                        worksheet.Cell($"A{tableRow}").Value = product.Name;
                        worksheet.Cell($"B{tableRow}").Value = product.Price;
                        worksheet.Cell($"C{tableRow}").Value = product.Description;
                    }

                    using (MemoryStream stream = new MemoryStream())
                    {
                        workbook.SaveAs(stream);
                        return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "ProductList.xlsx");
                    }
                }
            }
        }

        public DataTable getData()
        {
            //Creating DataTable  
            DataTable dt = new DataTable();
            //Setiing Table Name  
            dt.TableName = "EmployeeData";
            //Add Columns  
            dt.Columns.Add("ID", typeof(int));
            dt.Columns.Add("Name", typeof(string));
            dt.Columns.Add("City", typeof(string));
            //Add Rows in DataTable  
            dt.Rows.Add(1, "Anoop Kumar Sharma", "Delhi");
            dt.Rows.Add(2, "Andrew", "U.P.");
            dt.AcceptChanges();
            return dt;
        }

        public IList<Product> GetProductList()
        {
            var productList = new List<Product>();
            productList.Add(new Product { Name = "Acanon, hali moto", Price = 10, Description = "From china, Hung yen" });
            productList.Add(new Product { Name = "Monachi", Price = 50, Description = "From Brazil, Hai phong" });
            productList.Add(new Product { Name = "Acheno", Price = 100, Description = "From Amarica" });
            productList.Add(new Product { Name = "Natoshi", Price = 30, Description = "From Conton" });

            return productList;
        }
         
    }

    public class Product
    {
        public string Name { get; set; }

        public decimal Price { get; set; }

        public string Description { get; set; }
    }
}
