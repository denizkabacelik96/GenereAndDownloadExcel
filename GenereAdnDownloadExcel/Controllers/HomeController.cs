using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using GenereAdnDownloadExcel.Models;
using ClosedXML.Excel;
using System.IO;

namespace GenereAdnDownloadExcel.Controllers
{
    public class HomeController : Controller
    {

        List<Student> _ostudents = new List<Student>();
        public HomeController()
        {
            for (int i = 0; i < 9; i++)
            {
                _ostudents.Add(new Student()
                {
                    StudentId = i,

                    Name = "Student" + i,
                    Roll = "100" + i

                });
            }
        }

        public IActionResult Index()
        {
            using (var workbook=new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Students");
                var currentRow = 1;
                worksheet.Cell(currentRow, 1).Value = "StudentId";
                worksheet.Cell(currentRow, 2).Value = "Name";
                worksheet.Cell(currentRow, 3).Value = "Surname";
                foreach (var student in _ostudents)
                {
                    currentRow++;
                    worksheet.Cell(currentRow, 1).Value = student.StudentId;
                    worksheet.Cell(currentRow, 2).Value = student.Name;
                    worksheet.Cell(currentRow, 3).Value = student.Roll;
                }
                using (var stream = new MemoryStream()) {

                    workbook.SaveAs(stream);
                    var content = stream.ToArray();
                    return File(
                         content,
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        "Students.xlsx"

                        );
                
                }
            }
        }


    }
}
