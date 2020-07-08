using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using DowloadExcel.ExcelHelper;
using DowloadExcel.Model;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;

namespace DowloadExcel.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class WeatherForecastController : ControllerBase
    {
        private static readonly string[] Summaries = new[]
        {
            "Freezing", "Bracing", "Chilly", "Cool", "Mild", "Warm", "Balmy", "Hot", "Sweltering", "Scorching"
        };

        private readonly ILogger<WeatherForecastController> _logger;
        private readonly ITargetValueExcelExport _targetValueExcelExport;
        private readonly IWebHostEnvironment _webHostEnvironment;
        public WeatherForecastController(ILogger<WeatherForecastController> logger, ITargetValueExcelExport targetValueExcelExport, IWebHostEnvironment webHostEnvironment)
        {
            _logger = logger;
            _targetValueExcelExport = targetValueExcelExport;
            _webHostEnvironment = webHostEnvironment;
        }

        [HttpGet]
        public IEnumerable<WeatherForecast> Get()
        {
            var rng = new Random();
            return Enumerable.Range(1, 5).Select(index => new WeatherForecast
            {
                Date = DateTime.Now.AddDays(index),
                TemperatureC = rng.Next(-20, 55),
                Summary = Summaries[rng.Next(Summaries.Length)]
            })
            .ToArray();
        }
        [HttpGet("Export")]
        public ActionResult Export()
        {
            var main = new List<Main>() { new Main() { Year = 2020, Month = 7 } };
            var target = new List<StandardTarget>(){
                    new StandardTarget(){
                        CSPTarget=1,
                        DataFormat="%",
                        KPIName="test",
                        LineOrGroup="test",
                        PLTarget=1,
                        Target="test",
                    },
                   new StandardTarget(){
                        CSPTarget=2,
                        DataFormat="%",
                        KPIName="test2",
                        LineOrGroup="test2",
                        PLTarget=2,
                        Target="test2",
                    },
                    new StandardTarget(){
                        CSPTarget=3,
                        DataFormat="%",
                        KPIName="test3",
                        LineOrGroup="test3",
                        PLTarget=3,
                        Target="test3",
                    },
            };
            var nPNTTarget = new List<NPNTTarget>()
            {
                new NPNTTarget(){
                        DataFormat="%",
                        KPIName="test",
                        LineOrGroup="test",
                        PLTarget=1,
                },
                new NPNTTarget(){
                        DataFormat="%",
                        KPIName="test2",
                        LineOrGroup="test2",
                        PLTarget=2,
                },
                new NPNTTarget(){
                        DataFormat="%",
                        KPIName="test3",
                        LineOrGroup="test3",
                        PLTarget=3,
                }
            };
            var excelModel = new List<ExcelModel>(){
                    new ExcelModel(target,typeof(StandardTarget),"Standard Target"),
                     new ExcelModel(nPNTTarget,typeof(NPNTTarget),"NPNT Target"),
                     new ExcelModel(main,typeof(Main),"Main"),
            };
            var workBook = _targetValueExcelExport.CreateWorkbook(excelModel);

            MemoryStream ms = new NPOIMemoryStream();
            workBook.Write(ms);
            ms.Flush();
            ms.Position = 0;

            return new FileStreamResult(ms, "application/ms-excel") { FileDownloadName = "test.xlsx" };
        }
        public class NPOIMemoryStream : MemoryStream
        {
            /// <summary>
            /// 获取流是否关闭
            /// </summary>
            public bool IsColse
            {
                get;
                private set;
            }
            public NPOIMemoryStream(bool colse = false)
            {
                IsColse = colse;
            }
            public override void Close()
            {
                if (IsColse)
                {
                    base.Close();
                }
            }
        }
    }
}
