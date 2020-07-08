using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace DowloadExcel.Model
{
    [SheetName("Standard Target")]
    public class StandardTarget
    {
        [ColName("KPI Name")]
        public string KPIName { get; set; }
        [ColName("Line / Line Group")]
        public string LineOrGroup { get; set; }
        [ColName("Data Format")]
        public string DataFormat { get; set; }
        [ColName("PL Target")]
        public decimal PLTarget { get; set; }
        [ColName("CSP Target")]
        public decimal CSPTarget { get; set; }
        [ColName("Target")]
        public string Target { get; set; }
    }

    [SheetName("Main")]
    public class Main
    {
        [ColName("Year")]
        public int Year { get; set; }
        [ColName("Month")]
        public int Month { get; set; }
    }

    [SheetName("NPNT Target")]
    public class NPNTTarget
    {
        [ColName("KPI Name")]
        public string KPIName { get; set; }
        [ColName("Line / Line Group")]
        public string LineOrGroup { get; set; }
        [ColName("Data Format")]
        public string DataFormat { get; set; }
        [ColName("PL Target")]
        public decimal PLTarget { get; set; }
    }

    [AttributeUsage(AttributeTargets.Property)]
    public class ColNameAttribute : Attribute
    {
        public ColNameAttribute(string colName)
        {
            this.ColName = colName;
        }
        public string ColName { get; set; }
    }

    [AttributeUsage(AttributeTargets.Class)]
    public class SheetNameAttribute : Attribute
    {
        public SheetNameAttribute(string sheetName)
        {
            this.SheetName = sheetName;
        }
        public string SheetName { get; set; }
    }
}
