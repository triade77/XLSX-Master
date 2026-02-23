using System.Collections.Generic;

namespace XlsxMaster.Charts
{
    /// <summary>
    /// ChartXmlGenerator에 전달되는 차트 빌드 옵션 집합입니다.
    /// MasterChartBuilder → XlsxChartInjector → ChartXmlGenerator 경로로 흐릅니다.
    /// </summary>
    internal sealed class ChartBuildOptions
    {
        // ── 공통 ──────────────────────────────────────────────────────
        public string                         CategoryFormula  { get; set; }
        public IReadOnlyList<SeriesDefinition> Series          { get; set; }
        public List<ScatterSeriesDefinition>   ScatterSeries   { get; set; }
        public string                         Title            { get; set; }
        public bool                           ShowLegend       { get; set; }
        public bool                           ShowDataLabels   { get; set; }
        public int?                           ChartStyle       { get; set; }

        // ── 주 Y축 ────────────────────────────────────────────────────
        public double? YAxisMin   { get; set; }
        public double? YAxisMax   { get; set; }
        public string  YAxisTitle { get; set; }

        // ── 보조 Y축 ──────────────────────────────────────────────────
        public double? SecondaryYAxisMin   { get; set; }
        public double? SecondaryYAxisMax   { get; set; }

        // ── X축 ───────────────────────────────────────────────────────
        public string XAxisTitle { get; set; }
    }
}
