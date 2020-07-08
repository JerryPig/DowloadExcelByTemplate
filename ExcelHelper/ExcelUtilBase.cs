using DowloadExcel.Model;
using NPOI;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;

namespace DowloadExcel.ExcelHelper
{
    public abstract class ExcelExportByTempleBase : IExcelUtil
    {
        public ExcelExportByTempleBase()
        {
            _currentWorkBook = new XSSFWorkbook();
        }
        public virtual IWorkbook CreateWorkbook(List<ExcelModel> excelModels)
        {
            //IWorkbook workBook = new XSSFWorkbook();
            var sheetModels = SetSheetModel();
            foreach (var item in excelModels)
            {
                var sheetName = item.SheetName;
                var dataSoure = item.DataSoure;
                var sheetModel = sheetModels.FirstOrDefault(e => e.SheetName == sheetName);
                if (sheetModel == null)
                    throw new Exception("error!");
                ////获得模板sheet
                //var CPS = (_templateWorkBook as XSSFWorkbook).GetSheet(sheetName);
                //CPS.CopyTo(workBook, sheetName, true, true);//将模板sheet复制到目标sheet
                //else
                //workBook.CreateSheet(sheetName);
                var sheets = _currentWorkBook.GetSheet(sheetName);
                ////创建头
                //var headerRow = sheets.CreateRow(0);
                //if (sheetModel.HeaderRowStyle != null)
                //{
                //    var newRowStyle = workBook.CreateCellStyle();
                //    newRowStyle.CloneStyleFrom(sheetModel.HeaderRowStyle);
                //    headerRow.RowStyle = newRowStyle;
                //}
                //for (int col = 0; col < sheetModel.ExcelHeaders.Count; col++)
                //{
                //    var cellInsert = headerRow.CreateCell(col);
                //    var cellStyle = sheetModel.ExcelHeaders[col].HeaderCellStyle;
                //    //设置单元格样式　
                //    if (cellStyle != null)
                //    {
                //        var newCelStyle = workBook.CreateCellStyle() as XSSFCellStyle;
                //        newCelStyle.CloneStyleFrom(cellStyle);
                //        cellInsert.CellStyle = newCelStyle;
                //    }
                //    cellInsert.SetCellValue(sheetModel.ExcelHeaders[col].Title);
                //}

                var rowIndex = 1;
                //设置数据
                Type type = item.Type;
                var titleDic = CreateInstance(type);
                //设置行样式
                var newContentRowStyle = _currentWorkBook.CreateCellStyle();
                if (sheetModel.ContentRowStyle != null)
                {
                    newContentRowStyle.CloneStyleFrom(sheetModel.HeaderRowStyle);
                }
                foreach (var data in dataSoure)
                {
                    var insertRow = sheets.CreateRow(rowIndex++);
                    if (sheetModel.ContentRowStyle != null)
                        insertRow.RowStyle = newContentRowStyle;
                    for (int col = 0; col < sheetModel.ExcelHeaders.Count; col++)
                    {
                        titleDic.TryGetValue(sheetModel.ExcelHeaders[col].Title, out string propertyName);
                        if (!string.IsNullOrEmpty(propertyName))
                        {
                            var cellInsert = insertRow.CreateCell(col);
                            //(cellInsert as XSSFCell).CopyCellFrom();
                            var cellStyle = sheetModel.ExcelHeaders[col].ContentCellStyle;
                            //设置单元格样式　　　　
                            if (cellStyle != null)
                            {
                                var newCelStyle = _currentWorkBook.CreateCellStyle();
                                newCelStyle.CloneStyleFrom(cellStyle);
                                cellInsert.CellStyle = newCelStyle;
                                var font = cellInsert.CellStyle.GetFont(_currentWorkBook);
                                //cellInsert.CellStyle.CloneStyleFrom(cellStyle);
                            }
                            var drValue = GetStringValue(data, propertyName);
                            cellInsert.SetCellValue(drValue);
                        }
                    }
                }
            }
            return _currentWorkBook;
        }

        /// <summary>
        /// 创建映射字段，键：Excel列名 值：DTO属性名
        /// 待优化可考虑写入缓存
        /// </summary>
        /// <param name="exportType"></param>
        /// <returns></returns>
        private Dictionary<string, string> CreateInstance(Type exportType)
        {

            Dictionary<string, string> dict = new Dictionary<string, string>();
            exportType.GetProperties().ToList().ForEach(p =>
            {
                if (p.IsDefined(typeof(ColNameAttribute)))
                {
                    dict.Add(p.GetCustomAttribute<ColNameAttribute>().ColName, p.Name);
                }
            });
            return dict;
        }
        /// <summary>
        /// 反射获取导出DTO某个属性的值
        /// </summary>
        /// <param name="export"></param>
        /// <param name="propertyName"></param>
        /// <returns></returns>
        private string GetStringValue<T>(T export, string propertyName)
        {
            string strVal = string.Empty;
            var prop = export.GetType().GetProperties().Where(p => p.Name.Equals(propertyName)).SingleOrDefault();
            if (prop != null)
            {
                strVal = prop.GetValue(export).ToString();
            }

            return strVal;
        }

        public abstract IWorkbook CreateTemplateWorkbook();

        public virtual List<SheetModel> SetSheetModel()
        {
            int sheetCount = _templateWorkBook.NumberOfSheets;
            var result = new List<SheetModel>(sheetCount);
            for (int i = 0; i < _templateWorkBook.NumberOfSheets; i++)
            {
                var sheet = _templateWorkBook.GetSheetAt(i);
                (sheet as XSSFSheet).CopyTo(_currentWorkBook, sheet.SheetName, true, true);
                var headerRow = sheet.GetRow(0);
                var contentRow = sheet.GetRow(1);
                var headderCells = headerRow.Cells;
                var contentCells = contentRow.Cells;
                var header = new List<ExcelHeader>(headderCells.Count);
                for (int col = 0; col < headderCells.Count; col++)
                {

                    header.Add(new ExcelHeader()
                    {
                        ContentCellStyle = contentCells[col].CellStyle,

                        Title = headderCells[col].StringCellValue,
                        HeaderCellStyle = headderCells[col].CellStyle,
                    });
                }

                result.Add(new SheetModel()
                {
                    ContentRowStyle = contentRow.RowStyle,
                    HeaderRowStyle = headerRow.RowStyle,
                    SheetName = sheet.SheetName,
                    ExcelHeaders = header
                });
            }
            return result;
        }

        protected IWorkbook _templateWorkBook { get; set; }

        public IWorkbook _currentWorkBook { get; private set; }
    }
    public class SheetModel
    {
        public List<ExcelHeader> ExcelHeaders { get; set; }
        public string SheetName { get; set; }
        public ICellStyle ContentRowStyle { get; set; }
        public ICellStyle HeaderRowStyle { get; set; }

    }
    public class ExcelHeader
    {
        public string Title { get; set; }
        public ICellStyle ContentCellStyle { get; set; }

        public ICellStyle HeaderCellStyle { get; set; }

    }
    public class ExcelModel
    {
        public List<object> DataSoure { get; set; }

        public string SheetName { get; set; }

        public Type Type { get; set; }

        public int Sort { get; set; }

        public Dictionary<string, dynamic> Dic { get; set; }
        public ExcelModel(IEnumerable<object> enumerator, Type type, string sheetName = null)
        {
            DataSoure = new List<object>(enumerator);
            SheetName = sheetName;
            Type = type;
        }
    }
}
