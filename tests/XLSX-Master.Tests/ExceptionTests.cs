using System;
using System.Collections.Generic;
using System.IO;
using ClosedXML.Excel;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using XlsxMaster.Charts;
using XlsxMaster.Core;
using XlsxMaster.Extensions;

namespace XlsxMaster.Tests
{
    /// <summary>
    /// 잘못된 입력에 대한 예외 동작을 검증합니다.
    /// </summary>
    [TestClass]
    public class ExceptionTests
    {
        // ──────────────────────────────────────────────────────────────
        // 앵커 형식 오류
        // ──────────────────────────────────────────────────────────────

        [TestMethod]
        public void AddMasterChart_InvalidAnchorFormat_ThrowsArgumentException()
        {
            using (var workbook = new XLWorkbook())
            {
                var ws = workbook.Worksheets.Add("Sheet1");
                ws.Cell(1, 1).Value = "데이터";
                for (int i = 2; i <= 7; i++) ws.Cell(i, 1).Value = i;

                ws.AddMasterChart("INVALID_FORMAT")
                  .SetXAxis("A2:A7")
                  .AddSeries("값", "B2:B7", ChartType.Column);

                Assert.ThrowsException<ArgumentException>(() =>
                {
                    var s = workbook.SaveWithChartsToStream();
                    s.Dispose();
                }, "앵커 형식이 잘못된 경우 ArgumentException이 발생해야 합니다.");
            }
        }

        [TestMethod]
        public void AddMasterChart_ReversedAnchor_ThrowsArgumentException()
        {
            using (var workbook = new XLWorkbook())
            {
                var ws = workbook.Worksheets.Add("Sheet1");
                ws.Cell(1, 1).Value = "데이터";
                for (int i = 2; i <= 7; i++) ws.Cell(i, 1).Value = i;

                // 끝 셀(E1)이 시작 셀(M20)보다 앞 — 역순 앵커
                ws.AddMasterChart("M20:E1")
                  .SetXAxis("A2:A7")
                  .AddSeries("값", "B2:B7", ChartType.Column);

                var ex = Assert.ThrowsException<ArgumentException>(() =>
                {
                    var s = workbook.SaveWithChartsToStream();
                    s.Dispose();
                });

                StringAssert.Contains(ex.Message, "역순",
                    "역순 앵커 오류 메시지에 '역순' 설명이 포함되어야 합니다.");
            }
        }

        // ──────────────────────────────────────────────────────────────
        // 시트 미존재 오류
        // ──────────────────────────────────────────────────────────────

        [TestMethod]
        public void SaveWithCharts_UnknownSheetName_ThrowsKeyNotFoundWithSheetList()
        {
            // ClosedXML으로 Sheet1이 포함된 표준 xlsx 스트림 생성
            byte[] xlsxBytes;
            using (var workbook = new XLWorkbook())
            {
                var ws = workbook.Worksheets.Add("Sheet1");
                ws.Cell(1, 1).Value = "데이터";
                using (var tmp = new MemoryStream())
                {
                    workbook.SaveAs(tmp);
                    xlsxBytes = tmp.ToArray();
                }
            }

            // 존재하지 않는 시트 이름으로 직접 주입 시도
            var options = new ChartBuildOptions
            {
                CategoryFormula = "Sheet1!$A$2:$A$7",
                Series = new List<SeriesDefinition>
                {
                    new SeriesDefinition("값", "B2:B7", ChartType.Column,
                        AxisPosition.Primary, "Sheet1!$B$2:$B$7")
                },
                ScatterSeries = new List<ScatterSeriesDefinition>(),
                ShowLegend = true,
            };

            using (var ms = new MemoryStream(xlsxBytes))
            {
                var ex = Assert.ThrowsException<KeyNotFoundException>(() =>
                    XlsxChartInjector.Inject(ms, "NonExistentSheet", "D1:L15", options));

                StringAssert.Contains(ex.Message, "NonExistentSheet",
                    "오류 메시지에 요청한 시트명이 포함되어야 합니다.");
                StringAssert.Contains(ex.Message, "Sheet1",
                    "오류 메시지에 사용 가능한 시트 목록이 포함되어야 합니다.");
            }
        }

        // ──────────────────────────────────────────────────────────────
        // 계열/X축 누락 오류
        // ──────────────────────────────────────────────────────────────

        [TestMethod]
        public void SaveWithCharts_NoSeries_ThrowsInvalidOperationException()
        {
            using (var workbook = new XLWorkbook())
            {
                var ws = workbook.Worksheets.Add("Sheet1");
                ws.AddMasterChart("D1:L15").SetXAxis("A2:A7");
                // AddSeries() 생략

                var ex = Assert.ThrowsException<InvalidOperationException>(() =>
                {
                    var s = workbook.SaveWithChartsToStream();
                    s.Dispose();
                });

                StringAssert.Contains(ex.Message, "AddSeries",
                    "오류 메시지에 AddSeries() 안내가 포함되어야 합니다.");
            }
        }

        [TestMethod]
        public void SaveWithCharts_NoXAxis_NormalSeries_ThrowsInvalidOperationException()
        {
            using (var workbook = new XLWorkbook())
            {
                var ws = workbook.Worksheets.Add("Sheet1");
                ws.AddMasterChart("D1:L15")
                  .AddSeries("값", "B2:B7", ChartType.Column);
                // SetXAxis() 생략

                var ex = Assert.ThrowsException<InvalidOperationException>(() =>
                {
                    var s = workbook.SaveWithChartsToStream();
                    s.Dispose();
                });

                StringAssert.Contains(ex.Message, "SetXAxis",
                    "오류 메시지에 SetXAxis() 안내가 포함되어야 합니다.");
            }
        }
    }
}
