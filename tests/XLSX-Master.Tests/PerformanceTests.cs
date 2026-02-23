using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Packaging;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using XlsxMaster.Charts;
using XlsxMaster.Extensions;

namespace XlsxMaster.Tests
{
    /// <summary>
    /// 대용량 데이터 및 다중 차트 주입 시 성능을 검증합니다.
    /// </summary>
    [TestClass]
    public class PerformanceTests
    {
        private const int RowCount = 100_000;
        private const int TimeoutSeconds = 5;

        [TestMethod]
        public void LargeWorksheet_100kRows_WithChart_CompletesWithin5Seconds()
        {
            var sw = Stopwatch.StartNew();
            byte[] result;

            using (var workbook = new XLWorkbook())
            {
                var ws = workbook.Worksheets.Add("BigData");

                // 헤더
                ws.Cell(1, 1).Value = "인덱스";
                ws.Cell(1, 2).Value = "값";

                // 10만 행 데이터
                for (int i = 1; i <= RowCount; i++)
                {
                    ws.Cell(i + 1, 1).Value = i;
                    ws.Cell(i + 1, 2).Value = i * 1.5;
                }

                // 차트는 첫 6행만 참조 (대용량 범위가 아닌 실제 크기 차트)
                ws.AddMasterChart("D1:L18")
                  .SetXAxis("A2:A7")
                  .AddSeries("샘플", "B2:B7", ChartType.Column)
                  .SetTitle("대용량 데이터 시트 차트");

                using (var stream = workbook.SaveWithChartsToStream())
                    result = stream.ToArray();
            }

            sw.Stop();

            Assert.IsTrue(result.Length > 0, "결과 파일이 비어있습니다.");
            Assert.IsTrue(sw.Elapsed.TotalSeconds < TimeoutSeconds,
                $"10만 행 + 차트 생성이 {TimeoutSeconds}초를 초과했습니다. " +
                $"실제 소요: {sw.Elapsed.TotalSeconds:F2}초");
        }

        [TestMethod]
        public void TenCharts_SameWorkbook_CompletesSuccessfully()
        {
            const int ChartCount = 10;
            byte[] result;

            using (var workbook = new XLWorkbook())
            {
                var ws = workbook.Worksheets.Add("Sheet1");

                // 공통 데이터
                ws.Cell(1, 1).Value = "월";
                ws.Cell(1, 2).Value = "매출";
                for (int i = 0; i < 6; i++)
                {
                    ws.Cell(i + 2, 1).Value = (i + 1) + "월";
                    ws.Cell(i + 2, 2).Value = (i + 1) * 10;
                }

                // 10개 차트: 세로로 쌓이도록 앵커 배치
                for (int c = 0; c < ChartCount; c++)
                {
                    int rowStart = c * 16 + 1;
                    int rowEnd   = rowStart + 14;
                    string anchor = $"D{rowStart}:L{rowEnd}";

                    ws.AddMasterChart(anchor)
                      .SetXAxis("A2:A7")
                      .AddSeries($"계열{c + 1}", "B2:B7", ChartType.Column)
                      .SetTitle($"차트 {c + 1}");
                }

                using (var stream = workbook.SaveWithChartsToStream())
                    result = stream.ToArray();
            }

            Assert.IsTrue(result.Length > 0, "결과 파일이 비어있습니다.");

            // 차트 10개가 모두 주입되었는지 확인
            using (var ms = new MemoryStream(result))
            using (var doc = SpreadsheetDocument.Open(ms, isEditable: false))
            {
                var wsPart = doc.WorkbookPart.WorksheetParts.First();
                int injectedCount = wsPart.DrawingsPart.ChartParts.Count();
                Assert.AreEqual(ChartCount, injectedCount,
                    $"차트 {ChartCount}개가 모두 주입되어야 합니다. 실제: {injectedCount}개");
            }
        }
    }
}
