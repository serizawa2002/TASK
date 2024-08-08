using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using OfficeOpenXml;
using SalesOrderApp.Data;
using SalesOrderApp.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace SalesOrderApp.Controllers
{
    public class OrderController : Controller
    {
        private readonly SalesOrderContext _context;

        public OrderController(SalesOrderContext context)
        {
            _context = context;
        }

        public async Task<IActionResult> Index()
        {
            var orders = await _context.Orders
                               .Include(o => o.Customer)
                               .Include(o => o.OrderDetails)
                               .ThenInclude(od => od.Product)
                               .ToListAsync();

            return View(orders);
        }



        public IActionResult Create()
        {
            return View();
        }

        [HttpPost]
        public async Task<IActionResult> Create(Order order)
        {
            if (ModelState.IsValid)
            {
                order.OrderId = _context.Orders.Max(o => o.OrderId) + 1;
                order.NetAmount = order.OrderDetails.Sum(od => od.Amount);
                _context.Add(order);
                await _context.SaveChangesAsync();
                return RedirectToAction(nameof(Index));
            }
            return View(order);
        }

        public List<Order> ReadFromExcel(string filePath)
        {
            var orders = new List<Order>();

            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheet = package.Workbook.Worksheets[0];
                int rowCount = worksheet.Dimension.Rows;
                for (int row = 2; row <= rowCount; row++)
                {
                    var order = new Order
                    {
                        OrderId = int.Parse(worksheet.Cells[row, 1].Text),
                        Date = DateTime.Parse(worksheet.Cells[row, 2].Text),
                        Customer = new Customer { CustomerId = int.Parse(worksheet.Cells[row, 3].Text), Name = worksheet.Cells[row, 4].Text },
                        OrderDetails = new List<OrderDetail>(),
                        NetAmount = decimal.Parse(worksheet.Cells[row, 5].Text)
                    };

                    int orderDetailRowCount = worksheet.Dimension.Rows;
                    for (int odRow = row; odRow <= orderDetailRowCount; odRow++)
                    {
                        var product = new Product
                        {
                            ProductId = int.Parse(worksheet.Cells[odRow, 6].Text),
                            Name = worksheet.Cells[odRow, 7].Text,
                            Rate = decimal.Parse(worksheet.Cells[odRow, 8].Text)
                        };

                        var orderDetail = new OrderDetail
                        {
                            Product = product,
                            Qty = int.Parse(worksheet.Cells[odRow, 9].Text),
                            Rate = decimal.Parse(worksheet.Cells[odRow, 10].Text),
                            Amount = decimal.Parse(worksheet.Cells[odRow, 11].Text)
                        };

                        order.OrderDetails.Add(orderDetail);
                    }

                    orders.Add(order);
                }
            }

            return orders;
        }

        public void WriteToExcel(string filePath, List<Order> orders)
        {
            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Orders");

                worksheet.Cells[1, 1].Value = "Order ID";
                worksheet.Cells[1, 2].Value = "Date";
                worksheet.Cells[1, 3].Value = "Customer ID";
                worksheet.Cells[1, 4].Value = "Customer Name";
                worksheet.Cells[1, 5].Value = "Net Amount";
                worksheet.Cells[1, 6].Value = "Product ID";
                worksheet.Cells[1, 7].Value = "Product Name";
                worksheet.Cells[1, 8].Value = "Product Rate";
                worksheet.Cells[1, 9].Value = "Qty";
                worksheet.Cells[1, 10].Value = "Rate";
                worksheet.Cells[1, 11].Value = "Amount";

                int row = 2;
                foreach (var order in orders)
                {
                    worksheet.Cells[row, 1].Value = order.OrderId;
                    worksheet.Cells[row, 2].Value = order.Date;
                    worksheet.Cells[row, 3].Value = order.Customer.CustomerId;
                    worksheet.Cells[row, 4].Value = order.Customer.Name;
                    worksheet.Cells[row, 5].Value = order.NetAmount;

                    foreach (var detail in order.OrderDetails)
                    {
                        worksheet.Cells[row, 6].Value = detail.Product.ProductId;
                        worksheet.Cells[row, 7].Value = detail.Product.Name;
                        worksheet.Cells[row, 8].Value = detail.Product.Rate;
                        worksheet.Cells[row, 9].Value = detail.Qty;
                        worksheet.Cells[row, 10].Value = detail.Rate;
                        worksheet.Cells[row, 11].Value = detail.Amount;
                        row++;
                    }
                }

                FileInfo excelFile = new FileInfo(filePath);
                package.SaveAs(excelFile);
            }
        }

        public IActionResult ImportFromExcel()
        {
            string filePath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "files", "Orders.xlsx");
            var orders = ReadFromExcel(filePath);

            foreach (var order in orders)
            {
                if (!_context.Orders.Any(o => o.OrderId == order.OrderId))
                {
                    _context.Orders.Add(order);
                }
            }

            _context.SaveChanges();
            return RedirectToAction(nameof(Index));
        }

        public IActionResult ExportToExcel()
        {
            var orders = _context.Orders.Include(o => o.Customer).Include(o => o.OrderDetails).ThenInclude(od => od.Product).ToList();
            string filePath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "files", "ExportedOrders.xlsx");
            WriteToExcel(filePath, orders);
            return RedirectToAction(nameof(Index));
        }
    }
}
