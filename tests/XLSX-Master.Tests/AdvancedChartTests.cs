using System.IO;
using System.Linq;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Packaging;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using XlsxMaster.Charts;
using XlsxMaster.Extensions;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace XlsxMaster.Tests
{
    [TestClass]
    public class AdvancedChartTests
    {
        // ──────────────────────────────────────────────────────────────
        // 차트 타입 확장
        // ──────────────────────────────────────────────────────────────

        [TestMethod]
        public void PieChart_ProducesPieChartSeries_NoAxes()
        {
            var result = BuildChart(ws =>
                ws.AddMasterChart("D1:L18")
                  .SetXAxis("A2:A7")
                  .AddSeries("점유율", "B2:B7", ChartType.Pie)
                  .SetTitle("시장 점유율"));

            using (var ms = new MemoryStream(result))
            using (var doc = SpreadsheetDocument.Open(ms, isEditable: false))
            {
                var chartPart = GetChartPart(doc);
                var pieSeries = chartPart.ChartSpace.Descendants<C.PieChartSeries>().ToList();
                Assert.AreEqual(1, pieSeries.Count, "PieChartSeries가 1개여야 합니다.");

                // Pie chart에는 CategoryAxis / ValueAxis 없어야 함
                var catAxes = chartPart.ChartSpace.Descendants<C.CategoryAxis>().ToList();
                var valAxes = chartPart.ChartSpace.Descendants<C.ValueAxis>().ToList();
                Assert.AreEqual(0, catAxes.Count, "Pie 차트에는 CategoryAxis가 없어야 합니다.");
                Assert.AreEqual(0, valAxes.Count, "Pie 차트에는 ValueAxis가 없어야 합니다.");
            }
        }

        [TestMethod]
        public void ScatterChart_ProducesScatterSeries_TwoValueAxes()
        {
            var result = BuildChart(ws =>
                ws.AddMasterChart("D1:L18")
                  .AddScatterSeries("상관관계", "A2:A7", "B2:B7")
                  .SetTitle("분산형 차트"));

            using (var ms = new MemoryStream(result))
            using (var doc = SpreadsheetDocument.Open(ms, isEditable: false))
            {
                var chartPart = GetChartPart(doc);
                var scatterSeries = chartPart.ChartSpace.Descendants<C.ScatterChartSeries>().ToList();
                Assert.AreEqual(1, scatterSeries.Count, "ScatterChartSeries가 1개여야 합니다.");

                // XValues, YValues 모두 존재해야 함
                var xVals = chartPart.ChartSpace.Descendants<C.XValues>().ToList();
                var yVals = chartPart.ChartSpace.Descendants<C.YValues>().ToList();
                Assert.AreEqual(1, xVals.Count, "XValues가 1개여야 합니다.");
                Assert.AreEqual(1, yVals.Count, "YValues가 1개여야 합니다.");

                // 두 개의 ValueAxis (X용 + Y용)
                var valAxes = chartPart.ChartSpace.Descendants<C.ValueAxis>().ToList();
                Assert.AreEqual(2, valAxes.Count, "Scatter 차트에는 ValueAxis가 2개여야 합니다.");
            }
        }

        [TestMethod]
        public void ScatterChart_MixedWithSeries_ThrowsInvalidOperation()
        {
            using (var workbook = new XLWorkbook())
            {
                var ws = workbook.Worksheets.Add("Sheet1");
                FillSampleData(ws);
                ws.AddMasterChart("D1:L18")
                  .SetXAxis("A2:A7")
                  .AddSeries("일반", "B2:B7", ChartType.Column)
                  .AddScatterSeries("분산", "A2:A7", "C2:C7");

                Assert.ThrowsException<System.InvalidOperationException>(() =>
                {
                    var s = workbook.SaveWithChartsToStream();
                    s.Dispose();
                });
            }
        }

        [TestMethod]
        public void AreaStacked_ProducesStackedGrouping()
        {
            var result = BuildChart(ws =>
                ws.AddMasterChart("D1:L18")
                  .SetXAxis("A2:A7")
                  .AddSeries("A", "B2:B7", ChartType.AreaStacked)
                  .AddSeries("B", "C2:C7", ChartType.AreaStacked));

            using (var ms = new MemoryStream(result))
            using (var doc = SpreadsheetDocument.Open(ms, isEditable: false))
            {
                var chartPart = GetChartPart(doc);
                var grouping = chartPart.ChartSpace.Descendants<C.Grouping>().FirstOrDefault();
                Assert.IsNotNull(grouping, "Grouping 요소가 있어야 합니다.");
                Assert.AreEqual(C.GroupingValues.Stacked, grouping.Val.Value,
                    "Grouping이 Stacked여야 합니다.");
            }
        }

        // ──────────────────────────────────────────────────────────────
        // 축 커스터마이징
        // ──────────────────────────────────────────────────────────────

        [TestMethod]
        public void SetYAxisMin_Max_ProducesScalingElement()
        {
            var result = BuildChart(ws =>
                ws.AddMasterChart("D1:L18")
                  .SetXAxis("A2:A7")
                  .AddSeries("값", "B2:B7", ChartType.Column)
                  .SetYAxisMin(0)
                  .SetYAxisMax(500));

            using (var ms = new MemoryStream(result))
            using (var doc = SpreadsheetDocument.Open(ms, isEditable: false))
            {
                var chartPart = GetChartPart(doc);
                var minEl = chartPart.ChartSpace.Descendants<C.MinAxisValue>().FirstOrDefault();
                var maxEl = chartPart.ChartSpace.Descendants<C.MaxAxisValue>().FirstOrDefault();
                Assert.IsNotNull(minEl, "MinAxisValue가 있어야 합니다.");
                Assert.IsNotNull(maxEl, "MaxAxisValue가 있어야 합니다.");
                Assert.AreEqual(0.0,   minEl.Val.Value, 1e-9);
                Assert.AreEqual(500.0, maxEl.Val.Value, 1e-9);
            }
        }

        [TestMethod]
        public void SetSecondaryYAxisMin_Max_ProducesSecondaryScaling()
        {
            var result = BuildChart(ws =>
                ws.AddMasterChart("D1:L18")
                  .SetXAxis("A2:A7")
                  .AddSeries("매출", "B2:B7", ChartType.Column)
                  .AddSeries("성장률", "C2:C7", ChartType.Line, useSecondaryAxis: true)
                  .SetSecondaryYAxisMin(-50)
                  .SetSecondaryYAxisMax(50));

            using (var ms = new MemoryStream(result))
            using (var doc = SpreadsheetDocument.Open(ms, isEditable: false))
            {
                var chartPart = GetChartPart(doc);
                // 보조 ValueAxis (두 번째 ValueAxis)에 Scaling min/max 확인
                var valAxes = chartPart.ChartSpace.Descendants<C.ValueAxis>().ToList();
                Assert.AreEqual(2, valAxes.Count, "보조축이 있어야 합니다.");

                var secondaryValAx = valAxes[1];
                var minEl = secondaryValAx.Descendants<C.MinAxisValue>().FirstOrDefault();
                var maxEl = secondaryValAx.Descendants<C.MaxAxisValue>().FirstOrDefault();
                Assert.IsNotNull(minEl, "보조축 MinAxisValue가 있어야 합니다.");
                Assert.IsNotNull(maxEl, "보조축 MaxAxisValue가 있어야 합니다.");
                Assert.AreEqual(-50.0, minEl.Val.Value, 1e-9);
                Assert.AreEqual( 50.0, maxEl.Val.Value, 1e-9);
            }
        }

        [TestMethod]
        public void SetYAxisTitle_SetXAxisTitle_ProduceTitleElements()
        {
            var result = BuildChart(ws =>
                ws.AddMasterChart("D1:L18")
                  .SetXAxis("A2:A7")
                  .AddSeries("값", "B2:B7", ChartType.Column)
                  .SetYAxisTitle("매출(억)")
                  .SetXAxisTitle("월"));

            using (var ms = new MemoryStream(result))
            using (var doc = SpreadsheetDocument.Open(ms, isEditable: false))
            {
                var chartPart = GetChartPart(doc);
                // 축 Title 텍스트 검증
                var catAx = chartPart.ChartSpace.Descendants<C.CategoryAxis>().First();
                var valAx = chartPart.ChartSpace.Descendants<C.ValueAxis>().First();

                Assert.IsNotNull(catAx.Descendants<C.Title>().FirstOrDefault(), "X축 제목이 있어야 합니다.");
                Assert.IsNotNull(valAx.Descendants<C.Title>().FirstOrDefault(), "Y축 제목이 있어야 합니다.");
            }
        }

        // ──────────────────────────────────────────────────────────────
        // 시각 스타일
        // ──────────────────────────────────────────────────────────────

        [TestMethod]
        public void SetSeriesColor_ProducesShapeProperties()
        {
            var result = BuildChart(ws =>
                ws.AddMasterChart("D1:L18")
                  .SetXAxis("A2:A7")
                  .AddSeries("값", "B2:B7", ChartType.Column)
                  .SetSeriesColor("값", "#4472C4"));

            using (var ms = new MemoryStream(result))
            using (var doc = SpreadsheetDocument.Open(ms, isEditable: false))
            {
                var chartPart = GetChartPart(doc);
                var spPr = chartPart.ChartSpace
                    .Descendants<C.ChartShapeProperties>().FirstOrDefault();
                Assert.IsNotNull(spPr, "ChartShapeProperties(spPr)가 있어야 합니다.");

                var solidFill = spPr.Descendants<DocumentFormat.OpenXml.Drawing.SolidFill>().FirstOrDefault();
                Assert.IsNotNull(solidFill, "SolidFill이 있어야 합니다.");

                var hexColor = spPr
                    .Descendants<DocumentFormat.OpenXml.Drawing.RgbColorModelHex>()
                    .FirstOrDefault();
                Assert.IsNotNull(hexColor, "RgbColorModelHex가 있어야 합니다.");
                Assert.AreEqual("4472C4", hexColor.Val.Value);
            }
        }

        [TestMethod]
        public void SetChartStyle_ProducesStyleElement()
        {
            var result = BuildChart(ws =>
                ws.AddMasterChart("D1:L18")
                  .SetXAxis("A2:A7")
                  .AddSeries("값", "B2:B7", ChartType.Column)
                  .SetChartStyle(10));

            using (var ms = new MemoryStream(result))
            using (var doc = SpreadsheetDocument.Open(ms, isEditable: false))
            {
                var chartPart = GetChartPart(doc);
                var style = chartPart.ChartSpace.Descendants<C.Style>().FirstOrDefault();
                Assert.IsNotNull(style, "Style 요소가 있어야 합니다.");
                Assert.AreEqual((byte)10, style.Val.Value);
            }
        }

        [TestMethod]
        public void ShowDataLabels_True_ProducesDataLabels()
        {
            var result = BuildChart(ws =>
                ws.AddMasterChart("D1:L18")
                  .SetXAxis("A2:A7")
                  .AddSeries("값", "B2:B7", ChartType.Column)
                  .ShowDataLabels(true));

            using (var ms = new MemoryStream(result))
            using (var doc = SpreadsheetDocument.Open(ms, isEditable: false))
            {
                var chartPart = GetChartPart(doc);
                var dLbls = chartPart.ChartSpace.Descendants<C.DataLabels>().ToList();
                Assert.IsTrue(dLbls.Count > 0, "DataLabels 요소가 있어야 합니다.");
            }
        }

        [TestMethod]
        public void SetMarkerStyle_Circle_ProducesMarkerElement()
        {
            var result = BuildChart(ws =>
                ws.AddMasterChart("D1:L18")
                  .SetXAxis("A2:A7")
                  .AddSeries("값", "B2:B7", ChartType.Line)
                  .SetMarkerStyle("값", MarkerStyle.Circle));

            using (var ms = new MemoryStream(result))
            using (var doc = SpreadsheetDocument.Open(ms, isEditable: false))
            {
                var chartPart = GetChartPart(doc);
                var marker = chartPart.ChartSpace.Descendants<C.Marker>().FirstOrDefault();
                Assert.IsNotNull(marker, "Marker 요소가 있어야 합니다.");

                var symbol = marker.Descendants<C.Symbol>().FirstOrDefault();
                Assert.IsNotNull(symbol, "Symbol 요소가 있어야 합니다.");
                Assert.AreEqual(C.MarkerStyleValues.Circle, symbol.Val.Value);
            }
        }

        [TestMethod]
        public void SetChartStyle_OutOfRange_ThrowsArgumentOutOfRange()
        {
            using (var workbook = new XLWorkbook())
            {
                var ws = workbook.Worksheets.Add("Sheet1");
                Assert.ThrowsException<System.ArgumentOutOfRangeException>(() =>
                    ws.AddMasterChart("D1:L18").SetChartStyle(49));
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

            var rows = new (string m, int s, double g)[]
            {
                ("1월", 100, 0), ("2월", 120, 20), ("3월", 150, 25),
                ("4월", 130, -13), ("5월", 170, 30), ("6월", 200, 17),
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
