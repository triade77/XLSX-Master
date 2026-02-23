using System;
using System.Collections.Generic;
using System.Data;
using ClosedXML.Excel;
using XlsxMaster.Charts;
using XlsxMaster.Extensions;

namespace XlsxMaster.Samples
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("=== XLSX-Master 샘플 ===");
            Console.WriteLine();

            try
            {
                Demo_SingleColumnChart();
                Demo_ComboChartWithSecondaryAxis();
                Demo_DataBindingWithChart();
                Demo_AdvancedVisualization();
                Console.WriteLine();
                Console.WriteLine("모든 샘플 파일이 생성되었습니다.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"오류: {ex.Message}");
                Console.WriteLine(ex.StackTrace);
            }

            Console.WriteLine("계속하려면 아무 키나 누르세요...");
            Console.ReadKey();
        }

        /// <summary>3단계: 고급 시각화 기능 모음</summary>
        private static void Demo_AdvancedVisualization()
        {
            Console.WriteLine("[4] 고급 시각화(Pie/Scatter/AreaStacked/축 커스터마이징/스타일) 생성 중...");

            using (var workbook = new XLWorkbook())
            {
                // ── Sheet 1: Pie 차트 ────────────────────────────────
                var wsPie = workbook.Worksheets.Add("Pie");
                wsPie.Cell(1, 1).Value = "제품";
                wsPie.Cell(1, 2).Value = "점유율(%)";
                var pieData = new[] { ("스마트폰", 42), ("노트북", 28), ("태블릿", 18), ("웨어러블", 12) };
                for (int i = 0; i < pieData.Length; i++)
                {
                    wsPie.Cell(i + 2, 1).Value = pieData[i].Item1;
                    wsPie.Cell(i + 2, 2).Value = pieData[i].Item2;
                }
                wsPie.AddMasterChart("D1:L16")
                     .SetXAxis("A2:A5")
                     .AddSeries("점유율", "B2:B5", ChartType.Pie)
                     .SetTitle("제품별 시장 점유율")
                     .ShowLegend(true)
                     .ShowDataLabels(true);

                // ── Sheet 2: Scatter 차트 ────────────────────────────
                var wsSc = workbook.Worksheets.Add("Scatter");
                wsSc.Cell(1, 1).Value = "광고비(억)";
                wsSc.Cell(1, 2).Value = "매출(억)";
                var scData = new[] { (1.2, 12.0), (2.5, 25.5), (3.1, 31.0), (4.0, 38.5), (5.2, 52.0), (6.0, 58.0) };
                for (int i = 0; i < scData.Length; i++)
                {
                    wsSc.Cell(i + 2, 1).Value = scData[i].Item1;
                    wsSc.Cell(i + 2, 2).Value = scData[i].Item2;
                }
                wsSc.AddMasterChart("D1:L18")
                    .AddScatterSeries("광고비 vs 매출", "A2:A7", "B2:B7", hexColor: "ED7D31")
                    .SetTitle("광고비-매출 상관관계")
                    .SetXAxisTitle("광고비(억)")
                    .SetYAxisTitle("매출(억)")
                    .ShowLegend(true);

                // ── Sheet 3: AreaStacked + 축 범위 + 스타일 ──────────
                var wsAs = workbook.Worksheets.Add("AreaStacked");
                wsAs.Cell(1, 1).Value = "분기";
                wsAs.Cell(1, 2).Value = "서울";
                wsAs.Cell(1, 3).Value = "부산";
                wsAs.Cell(1, 4).Value = "대구";
                var asData = new[] { ("Q1", 320, 180, 120), ("Q2", 380, 210, 140), ("Q3", 420, 240, 160), ("Q4", 510, 280, 190) };
                for (int i = 0; i < asData.Length; i++)
                {
                    wsAs.Cell(i + 2, 1).Value = asData[i].Item1;
                    wsAs.Cell(i + 2, 2).Value = asData[i].Item2;
                    wsAs.Cell(i + 2, 3).Value = asData[i].Item3;
                    wsAs.Cell(i + 2, 4).Value = asData[i].Item4;
                }
                wsAs.AddMasterChart("F1:N18")
                    .SetXAxis("A2:A5")
                    .AddSeries("서울", "B2:B5", ChartType.AreaStacked)
                    .AddSeries("부산", "C2:C5", ChartType.AreaStacked)
                    .AddSeries("대구", "D2:D5", ChartType.AreaStacked)
                    .SetTitle("지역별 누적 매출")
                    .SetYAxisMin(0).SetYAxisMax(1200)
                    .SetYAxisTitle("매출(억)")
                    .SetXAxisTitle("분기")
                    .SetChartStyle(2)
                    .ShowLegend(true)
                    .ShowDataLabels(false);

                // ── Sheet 4: 계열 색상 + 마커 + 보조축 범위 ──────────
                var wsStyle = workbook.Worksheets.Add("Style");
                wsStyle.Cell(1, 1).Value = "월";
                wsStyle.Cell(1, 2).Value = "매출(억)";
                wsStyle.Cell(1, 3).Value = "성장률(%)";
                var styleData = new[] { ("1월", 120, 0.0), ("2월", 150, 25.0), ("3월", 180, 20.0), ("4월", 220, 22.2), ("5월", 190, -13.6), ("6월", 250, 31.6) };
                for (int i = 0; i < styleData.Length; i++)
                {
                    wsStyle.Cell(i + 2, 1).Value = styleData[i].Item1;
                    wsStyle.Cell(i + 2, 2).Value = styleData[i].Item2;
                    wsStyle.Cell(i + 2, 3).Value = styleData[i].Item3;
                }
                wsStyle.AddMasterChart("E1:M20")
                       .SetXAxis("A2:A7")
                       .AddSeries("매출", "B2:B7", ChartType.Column)
                       .AddSeries("성장률", "C2:C7", ChartType.Line, useSecondaryAxis: true)
                       .SetSeriesColor("매출", "#4472C4")
                       .SetMarkerStyle("성장률", MarkerStyle.Circle)
                       .SetSecondaryYAxisMin(-20)
                       .SetSecondaryYAxisMax(40)
                       .SetTitle("매출 + 성장률(보조축)")
                       .ShowLegend(true);

                workbook.SaveWithCharts("Sample4_Advanced.xlsx");
            }

            Console.WriteLine("    → Sample4_Advanced.xlsx 저장 완료");
        }

        // 샘플 데이터 모델
        private class MonthlySalesRecord
        {
            public string   Month      { get; set; }
            public int      Sales      { get; set; }
            public decimal  Revenue    { get; set; }
            public double   GrowthRate { get; set; }
            public DateTime ReportDate { get; set; }
        }

        /// <summary>1단계 PoC: 단일 Column 차트</summary>
        private static void Demo_SingleColumnChart()
        {
            Console.WriteLine("[1] 단일 막대 차트 생성 중...");

            using (var workbook = new XLWorkbook())
            {
                var ws = workbook.Worksheets.Add("월별매출");

                // 헤더
                ws.Cell(1, 1).Value = "월";
                ws.Cell(1, 2).Value = "매출액(억)";

                // 데이터
                var months = new[] { "1월", "2월", "3월", "4월", "5월", "6월",
                                     "7월", "8월", "9월", "10월", "11월", "12월" };
                var sales = new[] { 12, 15, 18, 22, 19, 25, 28, 30, 27, 32, 35, 40 };

                for (int i = 0; i < months.Length; i++)
                {
                    ws.Cell(i + 2, 1).Value = months[i];
                    ws.Cell(i + 2, 2).Value = sales[i];
                }

                // 테이블 서식
                ws.Range(1, 1, 13, 2).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                ws.Column(1).Width = 10;
                ws.Column(2).Width = 15;

                // 차트 추가
                ws.AddMasterChart("D1:L18")
                  .SetXAxis("A2:A13")
                  .AddSeries("매출액(억)", "B2:B13", ChartType.Column)
                  .SetTitle("월별 매출 현황")
                  .ShowLegend(true);

                workbook.SaveWithCharts("Sample1_SingleChart.xlsx");
            }

            Console.WriteLine("    → Sample1_SingleChart.xlsx 저장 완료");
        }

        /// <summary>2단계: IEnumerable&lt;T&gt; 바인딩 + Excel Table + 콤보 차트</summary>
        private static void Demo_DataBindingWithChart()
        {
            Console.WriteLine("[3] 데이터 바인딩 + Excel Table + 차트 생성 중...");

            using (var workbook = new XLWorkbook())
            {
                // ── Sheet 1: IEnumerable<T> 바인딩 + Excel Table + 콤보 차트 ──
                var ws1 = workbook.Worksheets.Add("월별실적");

                var records = new List<MonthlySalesRecord>
                {
                    new MonthlySalesRecord { Month = "1월",  Sales = 1200, Revenue = 960.00m,  GrowthRate = 0.0,   ReportDate = new DateTime(2024,  1, 31) },
                    new MonthlySalesRecord { Month = "2월",  Sales = 1500, Revenue = 1200.00m, GrowthRate = 25.0,  ReportDate = new DateTime(2024,  2, 29) },
                    new MonthlySalesRecord { Month = "3월",  Sales = 1800, Revenue = 1440.00m, GrowthRate = 20.0,  ReportDate = new DateTime(2024,  3, 31) },
                    new MonthlySalesRecord { Month = "4월",  Sales = 2200, Revenue = 1760.00m, GrowthRate = 22.2,  ReportDate = new DateTime(2024,  4, 30) },
                    new MonthlySalesRecord { Month = "5월",  Sales = 1900, Revenue = 1520.00m, GrowthRate = -13.6, ReportDate = new DateTime(2024,  5, 31) },
                    new MonthlySalesRecord { Month = "6월",  Sales = 2500, Revenue = 2000.00m, GrowthRate = 31.6,  ReportDate = new DateTime(2024,  6, 30) },
                    new MonthlySalesRecord { Month = "7월",  Sales = 2800, Revenue = 2240.00m, GrowthRate = 12.0,  ReportDate = new DateTime(2024,  7, 31) },
                    new MonthlySalesRecord { Month = "8월",  Sales = 3000, Revenue = 2400.00m, GrowthRate = 7.1,   ReportDate = new DateTime(2024,  8, 31) },
                    new MonthlySalesRecord { Month = "9월",  Sales = 2700, Revenue = 2160.00m, GrowthRate = -10.0, ReportDate = new DateTime(2024,  9, 30) },
                    new MonthlySalesRecord { Month = "10월", Sales = 3200, Revenue = 2560.00m, GrowthRate = 18.5,  ReportDate = new DateTime(2024, 10, 31) },
                    new MonthlySalesRecord { Month = "11월", Sales = 3500, Revenue = 2800.00m, GrowthRate = 9.4,   ReportDate = new DateTime(2024, 11, 30) },
                    new MonthlySalesRecord { Month = "12월", Sales = 4000, Revenue = 3200.00m, GrowthRate = 14.3,  ReportDate = new DateTime(2024, 12, 31) },
                };

                // IEnumerable<T> 바인딩 → Excel Table 변환
                ws1.InsertMasterTable(records)
                   .AddExcelTable("MonthlySales", XLTableTheme.TableStyleMedium9);

                // 열 너비 보정
                ws1.Column(1).Width = 8;   // Month
                ws1.Column(2).Width = 10;  // Sales
                ws1.Column(3).Width = 14;  // Revenue
                ws1.Column(4).Width = 13;  // GrowthRate
                ws1.Column(5).Width = 14;  // ReportDate

                // 콤보 차트: 판매량(막대) + 성장률(꺾은선, 보조축)
                ws1.AddMasterChart("G1:P22")
                   .SetXAxis("A2:A13")
                   .AddSeries("판매량(건)", "B2:B13", ChartType.Column)
                   .AddSeries("성장률(%)", "D2:D13", ChartType.Line, useSecondaryAxis: true)
                   .SetTitle("2024년 월별 실적 현황")
                   .ShowLegend(true);

                // ── Sheet 2: DataTable 바인딩 + Excel Table (차트 없음) ──
                var ws2 = workbook.Worksheets.Add("지역별매출");

                var dt = new DataTable();
                dt.Columns.Add("지역",     typeof(string));
                dt.Columns.Add("Q1매출",   typeof(int));
                dt.Columns.Add("Q2매출",   typeof(int));
                dt.Columns.Add("Q3매출",   typeof(int));
                dt.Columns.Add("Q4매출",   typeof(int));
                dt.Columns.Add("연간합계", typeof(int));
                dt.Rows.Add("서울",  8500,  9200,  9800, 11000, 38500);
                dt.Rows.Add("부산",  4200,  4800,  5100,  5900, 20000);
                dt.Rows.Add("인천",  3100,  3400,  3700,  4200, 14400);
                dt.Rows.Add("대구",  2800,  3000,  3200,  3600, 12600);
                dt.Rows.Add("광주",  2100,  2300,  2500,  2900,  9800);
                dt.Rows.Add("대전",  1900,  2100,  2300,  2700,  9000);

                ws2.InsertMasterTable(dt)
                   .AddExcelTable("RegionalSales", XLTableTheme.TableStyleMedium2);

                ws2.Column(1).Width = 8;
                for (int c = 2; c <= 6; c++) ws2.Column(c).Width = 12;

                workbook.SaveWithCharts("Sample3_DataBinding.xlsx");
            }

            Console.WriteLine("    → Sample3_DataBinding.xlsx 저장 완료");
        }

        /// <summary>PRD 코드 바이브: 콤보 차트 + 보조축</summary>
        private static void Demo_ComboChartWithSecondaryAxis()
        {
            Console.WriteLine("[2] 콤보 차트(막대+꺾은선, 보조축) 생성 중...");

            using (var workbook = new XLWorkbook())
            {
                var ws = workbook.Worksheets.Add("Sales Analysis");

                // 헤더
                ws.Cell(1, 1).Value = "월";
                ws.Cell(1, 2).Value = "매출액(억)";
                ws.Cell(1, 3).Value = "성장률(%)";

                // 데이터
                var months = new[] { "1월", "2월", "3월", "4월", "5월", "6월",
                                     "7월", "8월", "9월", "10월", "11월", "12월" };
                var sales  = new[] { 120, 150, 180, 220, 190, 250, 280, 300, 270, 320, 350, 400 };
                var growth = new[] { 0.0, 25.0, 20.0, 22.2, -13.6, 31.6, 12.0, 7.1, -10.0, 18.5, 9.4, 14.3 };

                for (int i = 0; i < months.Length; i++)
                {
                    ws.Cell(i + 2, 1).Value = months[i];
                    ws.Cell(i + 2, 2).Value = sales[i];
                    ws.Cell(i + 2, 3).Value = growth[i];
                    ws.Cell(i + 2, 3).Style.NumberFormat.Format = "0.0\"%\"";
                }

                // 열 너비 조정
                ws.Column(1).Width = 8;
                ws.Column(2).Width = 14;
                ws.Column(3).Width = 14;

                // PRD 코드 바이브 예시
                ws.AddMasterChart(anchor: "E1:M20")
                  .SetXAxis("A2:A13")
                  .AddSeries("매출액", "B2:B13", ChartType.Column)
                  .AddSeries("성장률", "C2:C13", ChartType.Line, useSecondaryAxis: true)
                  .SetTitle("월별 실적 분석")
                  .ShowLegend(true);

                workbook.SaveWithCharts("Sample2_ComboChart.xlsx");
            }

            Console.WriteLine("    → Sample2_ComboChart.xlsx 저장 완료");
        }
    }
}
