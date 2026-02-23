using System.IO;
using System.Linq;
using System.Text;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using XlsxMaster.Charts;
using XlsxMaster.Extensions;

namespace XlsxMaster.Tests
{
    [TestClass]
    public class ValidationTests
    {
        [TestMethod]
        public void GeneratedXlsx_PassesOpenXmlValidation()
        {
            byte[] result;

            using (var workbook = new XLWorkbook())
            {
                var ws = workbook.Worksheets.Add("Sheet1");
                ws.Cell(1, 1).Value = "월";
                ws.Cell(1, 2).Value = "매출";
                for (int i = 0; i < 6; i++)
                {
                    ws.Cell(i + 2, 1).Value = (i + 1) + "월";
                    ws.Cell(i + 2, 2).Value = (i + 1) * 10;
                }

                ws.AddMasterChart("D1:L15")
                  .SetXAxis("A2:A7")
                  .AddSeries("매출", "B2:B7", ChartType.Column)
                  .SetTitle("테스트 차트")
                  .ShowLegend(true);

                using (var stream = workbook.SaveWithChartsToStream())
                    result = stream.ToArray();
            }

            // Open XML SDK 유효성 검사
            using (var ms = new MemoryStream(result))
            using (var doc = SpreadsheetDocument.Open(ms, isEditable: false))
            {
                var validator = new OpenXmlValidator();
                var errors = validator.Validate(doc).ToList();

                if (errors.Count > 0)
                {
                    var sb = new StringBuilder();
                    sb.AppendLine($"OpenXML 유효성 오류 {errors.Count}개:");
                    foreach (var e in errors)
                        sb.AppendLine($"  [{e.ErrorType}] {e.Description} (Part: {e.Part?.Uri})");
                    Assert.Fail(sb.ToString());
                }
            }
        }
    }
}
