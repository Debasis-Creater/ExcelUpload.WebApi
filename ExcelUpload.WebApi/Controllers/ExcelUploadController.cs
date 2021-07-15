using ClosedXML.Excel;
using ExcelUpload.WebApi.Models;
using Microsoft.AspNetCore.Cors;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace ExcelUpload.WebApi.Controllers
{
    [Route("webapi/[controller]/[action]")]
    [ApiController]
    [EnableCors("AllowOrigin")]
    public class ExcelUploadController : ControllerBase
    {
        private readonly TestdatabaseContext _context;
        public ExcelUploadController(TestdatabaseContext context)
        {
            _context = context;           
        }
        /// <summary>
        /// Upload ExcelFile
        /// </summary>
        /// <param name="file"></param>
        /// <returns></returns>
        [HttpPost]
        public IActionResult ImportExcel()
        {
            var file = Request.Form.Files[0];

            if (file != null)
            {
                FileInfo fi = new FileInfo(file.FileName);
                string extension = fi.Extension;
                if (extension.ToLower() == ".xlsx")
                {               
                    var list = new List<Patient>();
                    using (var stream = new MemoryStream())
                    {
                        file.CopyTo(stream);
                        using (var package = new ExcelPackage(stream))
                        {
                            ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                            var rowCount = worksheet.Dimension.Rows;
                            for (int row = 2; row <= rowCount; row++)
                            {
                                list.Add(new Patient()
                                {
                                    PatientId = worksheet.Cells[row, 1].Value.ToString().Trim(),
                                    Name = worksheet.Cells[row, 2].Value.ToString().Trim(),
                                    Age = worksheet.Cells[row, 3].Value.ToString().Trim(),
                                    Address = worksheet.Cells[row, 4].Value.ToString().Trim(),
                                });
                            }
                        }
                    }
                    foreach (var item in list)
                    {
                        var excelData = new ExcelTable()
                        {
                            PatientId = item.PatientId,
                            Name = item.Name,
                            Age = item.Age,
                            Address = item.Address
                        };

                        _context.ExcelTable.Add(excelData);
                        _context.SaveChanges();
                    }
                }
                else
                {
                    return BadRequest("Please upload excel file only");
                }
            }
            else
            {
                return BadRequest("File is empty");
            }
            return Ok();
        }
        /// <summary>
        /// Export To Excel
        /// </summary>
        /// <returns></returns>
        public IActionResult ExportToExcel()
        {
            var patientData = _context.ExcelTable.ToList();

            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("PatientDetails");
                var currentRow = 1;
                worksheet.Cell(currentRow, 1).Value = "PatientId";
                worksheet.Cell(currentRow, 2).Value = "PatientName";
                worksheet.Cell(currentRow, 3).Value = "Age";
                worksheet.Cell(currentRow, 4).Value = "Address";
                worksheet.Cell(currentRow, 1).Style.Fill.BackgroundColor = XLColor.AliceBlue;
                worksheet.Cell(currentRow, 2).Style.Fill.BackgroundColor = XLColor.Green;
                worksheet.Cell(currentRow, 3).Style.Fill.BackgroundColor = XLColor.Blue;
                worksheet.Cell(currentRow, 4).Style.Fill.BackgroundColor = XLColor.Orange;
                foreach (var patient in patientData)
                {
                    currentRow++;
                    worksheet.Cell(currentRow, 1).Value = patient.PatientId;
                    worksheet.Cell(currentRow, 2).Value = patient.Name;
                    worksheet.Cell(currentRow, 3).Value = patient.Age;
                    worksheet.Cell(currentRow, 4).Value = patient.Address;
                }

                using (var stream = new MemoryStream())
                {
                    workbook.SaveAs(stream);
                    var content = stream.ToArray();

                    return File(
                        content,
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        "PatientDetails.xlsx");
                }
            }
        }
    }
}
