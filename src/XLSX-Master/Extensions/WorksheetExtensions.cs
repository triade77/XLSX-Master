using System;
using ClosedXML.Excel;
using XlsxMaster.Charts;
using XlsxMaster.Core;

namespace XlsxMaster.Extensions
{
    /// <summary>
    /// <see cref="IXLWorksheet"/>에 대한 XLSX-Master 확장 메서드입니다.
    /// </summary>
    public static class WorksheetExtensions
    {
        /// <summary>
        /// 워크시트에 차트를 추가하기 위한 Fluent 빌더를 생성합니다.
        /// </summary>
        /// <param name="worksheet">차트를 추가할 워크시트</param>
        /// <param name="anchor">차트 위치 (셀 범위 형식, 예: "E1:M20")</param>
        /// <returns>차트 설정을 위한 <see cref="MasterChartBuilder"/></returns>
        /// <example>
        /// <code>
        /// ws.AddMasterChart("E1:M20")
        ///   .SetXAxis("A2:A13")
        ///   .AddSeries("매출액", "B2:B13", ChartType.Column)
        ///   .AddSeries("성장률", "C2:C13", ChartType.Line, useSecondaryAxis: true)
        ///   .SetTitle("월별 실적 분석")
        ///   .ShowLegend(true);
        /// </code>
        /// </example>
        public static MasterChartBuilder AddMasterChart(this IXLWorksheet worksheet, string anchor)
        {
            if (worksheet == null)
                throw new ArgumentNullException(nameof(worksheet));
            if (string.IsNullOrWhiteSpace(anchor))
                throw new ArgumentException("앵커 범위를 지정해야 합니다.", nameof(anchor));

            var builder = new MasterChartBuilder(worksheet, anchor);
            ChartRegistry.Register(worksheet.Workbook, builder);
            return builder;
        }
    }
}
