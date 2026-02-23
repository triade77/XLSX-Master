using System.IO;
using System.Linq;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Packaging;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using XlsxMaster.Charts;
using XlsxMaster.Extensions;
using C = DocumentFormat.OpenXml.Drawing.Charts;
// ChartPart는 DocumentFormat.OpenXml.Packaging 소속

namespace XlsxMaster.Tests
{
    /// <summary>
    /// ChartXmlGenerator가 생성하는 차트 XML 구조를 통합 방식으로 검증합니다.
    /// ChartXmlGenerator는 internal이므로 SaveWithChartsToStream → SpreadsheetDocument 파싱 경로를 사용합니다.
    /// </summary>
    [TestClass]
    public class ChartXmlGeneratorTests
    {
        // ──────────────────────────────────────────────────────────────
        // 계열 수 검증
        // ──────────────────────────────────────────────────────────────

        [TestMethod]
        public void BarChart_TwoSeries_ProducesTwoBarChartSeries()
        {
            var result = BuildChart(ws =>
                ws.AddMasterChart("D1:L18")
                  .SetXAxis("A2:A7")
                  .AddSeries("매출", "B2:B7", ChartType.Column)
                  .AddSeries("성장률", "C2:C7", ChartType.Column));

            using (var ms = new MemoryStream(result))
            using (var doc = SpreadsheetDocument.Open(ms, isEditable: false))
            {
                var chartPart = GetChartPart(doc);
                var serList = chartPart.ChartSpace.Descendants<C.BarChartSeries>().ToList();
                Assert.AreEqual(2, serList.Count, "Column 계열 2개 → BarChartSeries 2개여야 합니다.");
            }
        }

        [TestMethod]
        public void LineChart_OneSeries_ProducesOneLineChartSeries()
        {
            var result = BuildChart(ws =>
                ws.AddMasterChart("D1:L18")
                  .SetXAxis("A2:A7")
                  .AddSeries("매출", "B2:B7", ChartType.Line));

            using (var ms = new MemoryStream(result))
            using (var doc = SpreadsheetDocument.Open(ms, isEditable: false))
            {
                var chartPart = GetChartPart(doc);
                var serList = chartPart.ChartSpace.Descendants<C.LineChartSeries>().ToList();
                Assert.AreEqual(1, serList.Count, "Line 계열 1개 → LineChartSeries 1개여야 합니다.");
            }
        }

        [TestMethod]
        public void AreaChart_OneSeries_ProducesOneAreaChartSeries()
        {
            var result = BuildChart(ws =>
                ws.AddMasterChart("D1:L18")
                  .SetXAxis("A2:A7")
                  .AddSeries("매출", "B2:B7", ChartType.Area));

            using (var ms = new MemoryStream(result))
            using (var doc = SpreadsheetDocument.Open(ms, isEditable: false))
            {
                var chartPart = GetChartPart(doc);
                var serList = chartPart.ChartSpace.Descendants<C.AreaChartSeries>().ToList();
                Assert.AreEqual(1, serList.Count, "Area 계열 1개 → AreaChartSeries 1개여야 합니다.");
            }
        }

        // ──────────────────────────────────────────────────────────────
        // 보조축 검증
        // ──────────────────────────────────────────────────────────────

        [TestMethod]
        public void ComboChart_WithSecondaryAxis_ProducesTwoValueAxes()
        {
            var result = BuildChart(ws =>
                ws.AddMasterChart("D1:L18")
                  .SetXAxis("A2:A7")
                  .AddSeries("매출", "B2:B7", ChartType.Column)
                  .AddSeries("성장률", "C2:C7", ChartType.Line, useSecondaryAxis: true));

            using (var ms = new MemoryStream(result))
            using (var doc = SpreadsheetDocument.Open(ms, isEditable: false))
            {
                var chartPart = GetChartPart(doc);
                var valAxes = chartPart.ChartSpace.Descendants<C.ValueAxis>().ToList();
                Assert.AreEqual(2, valAxes.Count, "보조축 사용 시 ValueAxis가 2개(주축+보조축)여야 합니다.");
            }
        }

        // ──────────────────────────────────────────────────────────────
        // 제목 요소 검증
        // ──────────────────────────────────────────────────────────────

        [TestMethod]
        public void ChartWithTitle_ProducesTitleElement()
        {
            var result = BuildChart(ws =>
                ws.AddMasterChart("D1:L18")
                  .SetXAxis("A2:A7")
                  .AddSeries("매출", "B2:B7", ChartType.Column)
                  .SetTitle("테스트 제목"));

            using (var ms = new MemoryStream(result))
            using (var doc = SpreadsheetDocument.Open(ms, isEditable: false))
            {
                var chartPart = GetChartPart(doc);
                var chartEl = chartPart.ChartSpace.Descendants<C.Chart>().First();
                // Chart 직계 하위의 Title만 확인 (축 Title 제외)
                var title = chartEl.Elements<C.Title>().FirstOrDefault();
                Assert.IsNotNull(title, "SetTitle() 호출 시 Chart 하위에 Title 요소가 있어야 합니다.");
            }
        }

        [TestMethod]
        public void ChartWithoutTitle_ProducesAutoTitleDeleted()
        {
            var result = BuildChart(ws =>
                ws.AddMasterChart("D1:L18")
                  .SetXAxis("A2:A7")
                  .AddSeries("매출", "B2:B7", ChartType.Column)
                // SetTitle() 생략
            );

            using (var ms = new MemoryStream(result))
            using (var doc = SpreadsheetDocument.Open(ms, isEditable: false))
            {
                var chartPart = GetChartPart(doc);
                var atd = chartPart.ChartSpace.Descendants<C.AutoTitleDeleted>().FirstOrDefault();
                Assert.IsNotNull(atd, "SetTitle() 생략 시 AutoTitleDeleted 요소가 있어야 합니다.");
                Assert.IsTrue(atd.Val?.Value == true, "AutoTitleDeleted.Val이 true여야 합니다.");
            }
        }

        // ──────────────────────────────────────────────────────────────
        // 헬퍼
        // ──────────────────────────────────────────────────────────────

        private static byte[] BuildChart(System.Action<IXLWorksheet> configure)
        {
            using (var workbook = new XLWorkbook())
            {
                var ws = workbook.Worksheets.Add("Sheet1");
                FillSampleData(ws);
                configure(ws);
                using (var stream = workbook.SaveWithChartsToStream())
                    return stream.ToArray();
            }
        }

        private static ChartPart GetChartPart(SpreadsheetDocument doc) =>
            doc.WorkbookPart.WorksheetParts.First().DrawingsPart.ChartParts.First();

        private static void FillSampleData(IXLWorksheet ws)
        {
            ws.Cell(1, 1).Value = "월";
            ws.Cell(1, 2).Value = "매출";
            ws.Cell(1, 3).Value = "성장률";

            for (int i = 0; i < 6; i++)
            {
                ws.Cell(i + 2, 1).Value = (i + 1) + "월";
                ws.Cell(i + 2, 2).Value = (i + 1) * 10;
                ws.Cell(i + 2, 3).Value = (i + 1) * 2.5;
            }
        }
    }
}
