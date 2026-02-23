using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ClosedXML.Excel;
using XlsxMaster.Core;

namespace XlsxMaster.Charts
{
    /// <summary>
    /// Fluent API로 차트를 정의하고 워크북에 주입하는 빌더 클래스입니다.
    /// <para>
    /// <c>ws.AddMasterChart("E1:M20")</c>로 인스턴스를 생성하세요.
    /// </para>
    /// </summary>
    public sealed class MasterChartBuilder
    {
        private readonly IXLWorksheet _worksheet;
        private readonly string       _anchor;

        // ── 기본 차트 옵션 ───────────────────────────────────────────
        private string _categoryRange;
        private string _title;
        private bool   _showLegend    = true;
        private bool   _showDataLabels;
        private int?   _chartStyle;

        // ── 계열 목록 ────────────────────────────────────────────────
        private readonly List<SeriesDefinition>        _series        = new List<SeriesDefinition>();
        private readonly List<ScatterSeriesDefinition> _scatterSeries = new List<ScatterSeriesDefinition>();

        // ── 주 Y축 ───────────────────────────────────────────────────
        private double? _yAxisMin;
        private double? _yAxisMax;
        private string  _yAxisTitle;

        // ── 보조 Y축 ─────────────────────────────────────────────────
        private double? _secondaryYAxisMin;
        private double? _secondaryYAxisMax;

        // ── X축 ─────────────────────────────────────────────────────
        private string _xAxisTitle;

        // ────────────────────────────────────────────────────────────

        internal MasterChartBuilder(IXLWorksheet worksheet, string anchor)
        {
            _worksheet = worksheet ?? throw new ArgumentNullException(nameof(worksheet));
            _anchor    = anchor    ?? throw new ArgumentNullException(nameof(anchor));
        }

        // ── 기본 Fluent API ─────────────────────────────────────────

        /// <summary>X축(카테고리) 데이터 범위를 지정합니다. 예: "A2:A13"</summary>
        public MasterChartBuilder SetXAxis(string categoryRange)
        {
            _categoryRange = categoryRange;
            return this;
        }

        /// <summary>차트에 일반 데이터 계열을 추가합니다.</summary>
        /// <param name="name">범례에 표시될 계열 이름</param>
        /// <param name="valuesRange">값 범위 (예: "B2:B13")</param>
        /// <param name="chartType">차트 유형 (기본값: Column)</param>
        /// <param name="useSecondaryAxis">보조 Y축(오른쪽) 사용 여부</param>
        public MasterChartBuilder AddSeries(
            string    name,
            string    valuesRange,
            ChartType chartType        = ChartType.Column,
            bool      useSecondaryAxis = false)
        {
            if (string.IsNullOrWhiteSpace(name))
                throw new ArgumentException("계열 이름은 비어있을 수 없습니다.", nameof(name));
            if (string.IsNullOrWhiteSpace(valuesRange))
                throw new ArgumentException("값 범위는 비어있을 수 없습니다.", nameof(valuesRange));

            _series.Add(new SeriesDefinition(
                name,
                valuesRange,
                chartType,
                useSecondaryAxis ? AxisPosition.Secondary : AxisPosition.Primary,
                ToAbsoluteFormula(_worksheet.Name, valuesRange)));

            return this;
        }

        /// <summary>
        /// Scatter(분산형) 계열을 추가합니다. X범위와 Y범위를 각각 지정합니다.
        /// <para>Scatter 계열과 일반 계열(<see cref="AddSeries"/>)은 같은 차트에 혼용할 수 없습니다.</para>
        /// </summary>
        /// <param name="name">범례에 표시될 계열 이름</param>
        /// <param name="xRange">X값 범위 (예: "A2:A13")</param>
        /// <param name="yRange">Y값 범위 (예: "B2:B13")</param>
        /// <param name="hexColor">계열 색상 hex 코드 (예: "FF0000"). null이면 기본색 사용.</param>
        public MasterChartBuilder AddScatterSeries(
            string name,
            string xRange,
            string yRange,
            string hexColor = null)
        {
            if (string.IsNullOrWhiteSpace(name))
                throw new ArgumentException("계열 이름은 비어있을 수 없습니다.", nameof(name));
            if (string.IsNullOrWhiteSpace(xRange))
                throw new ArgumentException("X값 범위는 비어있을 수 없습니다.", nameof(xRange));
            if (string.IsNullOrWhiteSpace(yRange))
                throw new ArgumentException("Y값 범위는 비어있을 수 없습니다.", nameof(yRange));

            _scatterSeries.Add(new ScatterSeriesDefinition(
                name,
                ToAbsoluteFormula(_worksheet.Name, xRange),
                ToAbsoluteFormula(_worksheet.Name, yRange),
                hexColor == null ? null : NormalizeHexColor(hexColor)));

            return this;
        }

        /// <summary>차트 제목을 설정합니다.</summary>
        public MasterChartBuilder SetTitle(string title)
        {
            _title = title;
            return this;
        }

        /// <summary>범례 표시 여부를 설정합니다. (기본값: true)</summary>
        public MasterChartBuilder ShowLegend(bool show)
        {
            _showLegend = show;
            return this;
        }

        /// <summary>모든 계열에 데이터 레이블 표시 여부를 설정합니다. (기본값: false)</summary>
        public MasterChartBuilder ShowDataLabels(bool show)
        {
            _showDataLabels = show;
            return this;
        }

        /// <summary>Excel 내장 차트 스타일을 적용합니다. (1~48)</summary>
        public MasterChartBuilder SetChartStyle(int styleId)
        {
            if (styleId < 1 || styleId > 48)
                throw new ArgumentOutOfRangeException(nameof(styleId), "차트 스타일은 1~48 범위여야 합니다.");
            _chartStyle = styleId;
            return this;
        }

        // ── 계열 커스터마이징 ────────────────────────────────────────

        /// <summary>특정 계열의 색상을 지정합니다.</summary>
        /// <param name="seriesName">대상 계열 이름</param>
        /// <param name="hexColor">색상 hex 코드 (예: "FF0000" 또는 "#FF0000")</param>
        public MasterChartBuilder SetSeriesColor(string seriesName, string hexColor)
        {
            var series = _series.FirstOrDefault(s => s.Name == seriesName);
            if (series == null)
                throw new ArgumentException($"계열 '{seriesName}'을(를) 찾을 수 없습니다.", nameof(seriesName));
            series.HexColor = NormalizeHexColor(hexColor);
            return this;
        }

        /// <summary>꺾은선 계열의 마커 모양을 지정합니다.</summary>
        /// <param name="seriesName">대상 계열 이름 (Line 타입이어야 합니다)</param>
        /// <param name="style">마커 모양</param>
        public MasterChartBuilder SetMarkerStyle(string seriesName, MarkerStyle style)
        {
            var series = _series.FirstOrDefault(s => s.Name == seriesName);
            if (series == null)
                throw new ArgumentException($"계열 '{seriesName}'을(를) 찾을 수 없습니다.", nameof(seriesName));
            series.MarkerStyle = style;
            return this;
        }

        // ── 주 Y축 ───────────────────────────────────────────────────

        /// <summary>주 Y축의 최솟값을 지정합니다.</summary>
        public MasterChartBuilder SetYAxisMin(double min) { _yAxisMin = min; return this; }

        /// <summary>주 Y축의 최댓값을 지정합니다.</summary>
        public MasterChartBuilder SetYAxisMax(double max) { _yAxisMax = max; return this; }

        /// <summary>주 Y축 제목 텍스트를 설정합니다.</summary>
        public MasterChartBuilder SetYAxisTitle(string title) { _yAxisTitle = title; return this; }

        // ── 보조 Y축 ─────────────────────────────────────────────────

        /// <summary>보조 Y축의 최솟값을 지정합니다.</summary>
        public MasterChartBuilder SetSecondaryYAxisMin(double min) { _secondaryYAxisMin = min; return this; }

        /// <summary>보조 Y축의 최댓값을 지정합니다.</summary>
        public MasterChartBuilder SetSecondaryYAxisMax(double max) { _secondaryYAxisMax = max; return this; }

        // ── X축 ─────────────────────────────────────────────────────

        /// <summary>X축(카테고리 축) 제목 텍스트를 설정합니다.</summary>
        public MasterChartBuilder SetXAxisTitle(string title) { _xAxisTitle = title; return this; }

        // ── 주입 ─────────────────────────────────────────────────────

        /// <summary>
        /// 이미 저장된 xlsx 스트림에 이 빌더의 차트를 주입하여 새 스트림으로 반환합니다.
        /// 일반적으로 <see cref="Extensions.WorkbookExtensions.SaveWithCharts"/>를 통해 호출됩니다.
        /// </summary>
        internal MemoryStream InjectInto(Stream xlsxStream)
        {
            bool isScatterMode = _scatterSeries.Count > 0;

            if (isScatterMode && _series.Count > 0)
                throw new InvalidOperationException(
                    $"차트 앵커 '{_anchor}': Scatter 계열과 일반 계열을 같은 차트에 혼용할 수 없습니다.");

            if (!isScatterMode)
            {
                if (_series.Count == 0)
                    throw new InvalidOperationException(
                        $"차트 앵커 '{_anchor}': SaveWithCharts() 전에 AddSeries()로 계열을 추가하세요.");
                if (string.IsNullOrWhiteSpace(_categoryRange))
                    throw new InvalidOperationException(
                        $"차트 앵커 '{_anchor}': SaveWithCharts() 전에 SetXAxis()로 카테고리 범위를 지정하세요.");
            }

            var options = new ChartBuildOptions
            {
                CategoryFormula  = isScatterMode ? null : ToAbsoluteFormula(_worksheet.Name, _categoryRange),
                Series           = _series,
                ScatterSeries    = _scatterSeries,
                Title            = _title,
                ShowLegend       = _showLegend,
                ShowDataLabels   = _showDataLabels,
                ChartStyle       = _chartStyle,
                YAxisMin         = _yAxisMin,
                YAxisMax         = _yAxisMax,
                YAxisTitle       = _yAxisTitle,
                SecondaryYAxisMin = _secondaryYAxisMin,
                SecondaryYAxisMax = _secondaryYAxisMax,
                XAxisTitle       = _xAxisTitle,
            };

            return XlsxChartInjector.Inject(xlsxStream, _worksheet.Name, _anchor, options);
        }

        // ── 내부 헬퍼 ────────────────────────────────────────────────

        // 예: sheetName="Sales Analysis", range="B2:B13"
        // → "'Sales Analysis'!$B$2:$B$13"
        private static string ToAbsoluteFormula(string sheetName, string range)
        {
            var quotedSheet = sheetName.Contains(" ") ? $"'{sheetName}'" : sheetName;
            var absolute = System.Text.RegularExpressions.Regex.Replace(
                range.ToUpperInvariant(),
                @"([A-Z]+)(\d+)",
                m => $"${m.Groups[1].Value}${m.Groups[2].Value}");
            return $"{quotedSheet}!{absolute}";
        }

        // "#FF0000" 또는 "FF0000" → "FF0000"
        private static string NormalizeHexColor(string hexColor)
        {
            if (string.IsNullOrWhiteSpace(hexColor))
                throw new ArgumentException("색상 코드는 비어있을 수 없습니다.", nameof(hexColor));
            var clean = hexColor.TrimStart('#').ToUpperInvariant();
            if (clean.Length != 6)
                throw new ArgumentException("색상은 6자리 hex 코드여야 합니다 (예: FF0000).", nameof(hexColor));
            foreach (char c in clean)
                if (!((c >= '0' && c <= '9') || (c >= 'A' && c <= 'F')))
                    throw new ArgumentException("유효하지 않은 hex 색상 코드입니다.", nameof(hexColor));
            return clean;
        }
    }
}
