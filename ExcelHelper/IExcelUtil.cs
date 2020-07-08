using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Threading.Tasks;

namespace DowloadExcel.ExcelHelper
{
    public interface IExcelUtil
    {
        IWorkbook CreateWorkbook(List<ExcelModel> excelModels) ;
    }
}
