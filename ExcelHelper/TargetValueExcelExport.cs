using Microsoft.AspNetCore.Hosting;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace DowloadExcel.ExcelHelper
{
    public class TargetValueExcelExport : ExcelExportByTempleBase, ITargetValueExcelExport
    {
        private readonly IWebHostEnvironment _webHostEnvironment;

        public TargetValueExcelExport(IWebHostEnvironment webHostEnvironment)
        {
            _webHostEnvironment = webHostEnvironment;
            _templateWorkBook = CreateTemplateWorkbook();
        }
        public override IWorkbook CreateTemplateWorkbook()
        {
            XSSFWorkbook hssfworkbook = null;
            var path = _webHostEnvironment.ContentRootPath;
            string templetfilepath = @$"{ _webHostEnvironment.ContentRootPath}\\ExcelHelper\\TargetValueTemplate.xlsx";//模版Excel

            using (var fileRead = new FileStream(templetfilepath, FileMode.Open, FileAccess.Read))
                hssfworkbook = new XSSFWorkbook(fileRead);

            return hssfworkbook;
        }
    }
}
