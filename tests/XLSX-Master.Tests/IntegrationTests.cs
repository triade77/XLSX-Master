using System;
using System.IO;
using System.Linq;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Drawing.Charts;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using XlsxMaster.Charts;
using XlsxMaster.Extensions;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace XlsxMaster.Tests
{
    [TestClass]
    public class IntegrationTests
    {
        [TestMethod]
        public void SingleColumnChart_ProducesValidXlsx()
        {
            byte[] result;

            using (var workbook = new XLWorkbook())
            {
                var ws = workbook.Worksheets.Add("Sheet1");
                FillSampleData(ws);

                ws.AddMasterChart("D1:L15")
                  .SetXAxis("A2:A7")
                  .AddSeries("값", "B2:B7", ChartType.Column)
                  .SetTitle("테스트 차트")
                  .ShowLegend(true);

                using (var stream = workbook.SaveWithChartsToStream())
                    result = stream.ToArray();
            }

            Assert.IsTrue(result.Length > 0, "출력 파일이 비어있습니다.");

            // 생성된 파일을 다시 파싱하여 차트 XML 검증
            using (var ms = new MemoryStream(result))
            using (var doc = SpreadsheetDocument.Open(ms, isEditable: false))
            {
                var wsPart = doc.WorkbookPart.WorksheetParts.First();

                Assert.IsNotNull(wsPart.DrawingsPart, "DrawingsPart가 없습니다.");
                Assert.IsNotNull(wsPart.DrawingsPart.ChartParts.FirstOrDefault(), "ChartPart가 없습니다.");

                var chartPart = wsPart.DrawingsPart.ChartParts.First();
                var serList = chartPart.ChartSpace.Descendants<BarChartSeries>().ToList();
                Assert.AreEqual(1, serList.Count, "계열(Series) 수가 일치하지 않습니다.");
            }
        }

        [TestMethod]
        public void ComboChartWithSecondaryAxis_ProducesTwoChartElements()
        {
            byte[] result;

            using (var workbook = new XLWorkbook())
            {
                var ws = workbook.Worksheets.Add("Sheet1");
                FillSampleData(ws);

                ws.AddMasterChart("D1:L20")
                  .SetXAxis("A2:A7")
                  .AddSeries("매출", "B2:B7", ChartType.Column)
                  .AddSeries("성장률", "C2:C7", ChartType.Line, useSecondaryAxis: true)
                  .SetTitle("콤보 차트")
                  .ShowLegend(true);

                using (var stream = workbook.SaveWithChartsToStream())
                    result = stream.ToArray();
            }

            using (var ms = new MemoryStream(result))
            using (var doc = SpreadsheetDocument.Open(ms, isEditable: false))
            {
                var chartPart = doc.WorkbookPart.WorksheetParts
                    .First().DrawingsPart.ChartParts.First();

                var barSeries = chartPart.ChartSpace.Descendants<BarChartSeries>().ToList();
                var lineSeries = chartPart.ChartSpace.Descendants<LineChartSeries>().ToList();

                Assert.AreEqual(1, barSeries.Count, "막대 계열 수가 일치하지 않습니다.");
                Assert.AreEqual(1, lineSeries.Count, "꺾은선 계열 수가 일치하지 않습니다.");
            }
        }

        [TestMethod]
        public void MultipleChartsOnSameSheet_BothInjected()
        {
            byte[] result;

            using (var workbook = new XLWorkbook())
            {
                var ws = workbook.Worksheets.Add("Sheet1");
                FillSampleData(ws);

                ws.AddMasterChart("D1:L12")
                  .SetXAxis("A2:A7")
                  .AddSeries("매출", "B2:B7", ChartType.Column)
                  .SetTitle("차트1");

                ws.AddMasterChart("D14:L26")
                  .SetXAxis("A2:A7")
                  .AddSeries("성장률", "C2:C7", ChartType.Line)
                  .SetTitle("차트2");

                using (var stream = workbook.SaveWithChartsToStream())
                    result = stream.ToArray();
            }

            using (var ms = new MemoryStream(result))
            using (var doc = SpreadsheetDocument.Open(ms, isEditable: false))
            {
                var wsPart = doc.WorkbookPart.WorksheetParts.First();
                var chartCount = wsPart.DrawingsPart.ChartParts.Count();
                Assert.AreEqual(2, chartCount, "차트 수가 일치하지 않습니다.");
            }
        }

        // ──────────────────────────────────────────────────────────────
        // Fluent API 보완 테스트
        // ──────────────────────────────────────────────────────────────

        [TestMethod]
        public void SetTitle_Omitted_AutoTitleDeleted_IsTrue()
        {
            byte[] result;
            using (var workbook = new XLWorkbook())
            {
                var ws = workbook.Worksheets.Add("Sheet1");
                FillSampleData(ws);
                ws.AddMasterChart("D1:L15")
                  .SetXAxis("A2:A7")
                  .AddSeries("값", "B2:B7", ChartType.Column);
                // SetTitle() 생략

                using (var stream = workbook.SaveWithChartsToStream())
                    result = stream.ToArray();
            }

            using (var ms = new MemoryStream(result))
            using (var doc = SpreadsheetDocument.Open(ms, isEditable: false))
            {
                var chartPart = doc.WorkbookPart.WorksheetParts.First().DrawingsPart.ChartParts.First();
                var titles = chartPart.ChartSpace.Descendants<C.Title>().ToList();
                Assert.AreEqual(0, titles.Count, "SetTitle() 생략 시 제목 요소가 없어야 합니다.");

                var atd = chartPart.ChartSpace.Descendants<AutoTitleDeleted>().FirstOrDefault();
                Assert.IsTrue(atd?.Val?.Value == true, "SetTitle() 생략 시 AutoTitleDeleted=true여야 합니다.");
            }
        }

        [TestMethod]
        public void ShowLegend_False_NoLegendElement()
        {
            byte[] result;
            using (var workbook = new XLWorkbook())
            {
                var ws = workbook.Worksheets.Add("Sheet1");
                FillSampleData(ws);
                ws.AddMasterChart("D1:L15")
                  .SetXAxis("A2:A7")
                  .AddSeries("값", "B2:B7", ChartType.Column)
                  .ShowLegend(false);

                using (var stream = workbook.SaveWithChartsToStream())
                    result = stream.ToArray();
            }

            using (var ms = new MemoryStream(result))
            using (var doc = SpreadsheetDocument.Open(ms, isEditable: false))
            {
                var chartPart = doc.WorkbookPart.WorksheetParts.First().DrawingsPart.ChartParts.First();
                var legends = chartPart.ChartSpace.Descendants<Legend>().ToList();
                Assert.AreEqual(0, legends.Count, "ShowLegend(false) 시 Legend 요소가 없어야 합니다.");
            }
        }

        [TestMethod]
        public void NoSeries_ThrowsInvalidOperationException_WithAnchorInfo()
        {
            using (var workbook = new XLWorkbook())
            {
                var ws = workbook.Worksheets.Add("Sheet1");
                ws.AddMasterChart("D1:L15").SetXAxis("A2:A7");
                // AddSeries() 생략

                var ex = Assert.ThrowsException<InvalidOperationException>(() =>
                {
                    var stream = workbook.SaveWithChartsToStream();
                    stream.Dispose();
                });

                StringAssert.Contains(ex.Message, "AddSeries", "에러 메시지에 AddSeries 안내가 포함되어야 합니다.");
                StringAssert.Contains(ex.Message, "D1:L15", "에러 메시지에 앵커 정보가 포함되어야 합니다.");
            }
        }

        [TestMethod]
        public void NoXAxis_ThrowsInvalidOperationException_WithAnchorInfo()
        {
            using (var workbook = new XLWorkbook())
            {
                var ws = workbook.Worksheets.Add("Sheet1");
                ws.AddMasterChart("D1:L15")
                  .AddSeries("값", "B2:B7", ChartType.Column);
                // SetXAxis() 생략

                var ex = Assert.ThrowsException<InvalidOperationException>(() =>
                {
                    var stream = workbook.SaveWithChartsToStream();
                    stream.Dispose();
                });

                StringAssert.Contains(ex.Message, "SetXAxis", "에러 메시지에 SetXAxis 안내가 포함되어야 합니다.");
                StringAssert.Contains(ex.Message, "D1:L15", "에러 메시지에 앵커 정보가 포함되어야 합니다.");
            }
        }

        [TestMethod]
        public void SaveWithCharts_StreamOverload_WritesToDestinationStream()
        {
            using (var workbook = new XLWorkbook())
            {
                var ws = workbook.Worksheets.Add("Sheet1");
                FillSampleData(ws);
                ws.AddMasterChart("D1:L15")
                  .SetXAxis("A2:A7")
                  .AddSeries("값", "B2:B7", ChartType.Column);

                using (var destination = new MemoryStream())
                {
                    workbook.SaveWithCharts(destination);
                    Assert.IsTrue(destination.Length > 0, "대상 스트림에 데이터가 기록되어야 합니다.");
                }
            }
        }

        private static void FillSampleData(IXLWorksheet ws)
        {
            ws.Cell(1, 1).Value = "월";
            ws.Cell(1, 2).Value = "매출";
            ws.Cell(1, 3).Value = "성장률";

            var rows = new (string m, int s, double g)[]
            {
                ("1월", 100, 0),
                ("2월", 120, 20),
                ("3월", 150, 25),
                ("4월", 130, -13),
                ("5월", 170, 30),
                ("6월", 200, 17),
            };

            for (int i = 0; i < rows.Length; i++)
            {
                ws.Cell(i + 2, 1).Value = rows[i].m;
                ws.Cell(i + 2, 2).Value = rows[i].s;
                ws.Cell(i + 2, 3).Value = rows[i].g;
            }
        }
    }
}
