using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace XlsxMaster.Charts
{
    /// <summary>
    /// Open XML SDK를 사용하여 ChartSpace XML 객체를 생성합니다.
    /// 단일/다중 계열, 콤보 차트, 보조축, Pie, Scatter, AreaStacked를 지원합니다.
    /// </summary>
    internal static class ChartXmlGenerator
    {
        // 주축 ID
        private const uint PrimaryAxId    = 1u;
        private const uint PrimaryValAxId = 2u;
        // 보조축 ID
        private const uint SecondaryAxId    = 3u;
        private const uint SecondaryValAxId = 4u;
        // Scatter 전용 축 ID
        private const uint ScatterXAxId = 201u;
        private const uint ScatterYAxId = 202u;

        // ── 공개 진입점 ─────────────────────────────────────────────

        /// <summary>ChartBuildOptions 기반으로 ChartSpace 객체를 빌드합니다.</summary>
        internal static ChartSpace BuildChartSpace(ChartBuildOptions options)
        {
            var chartSpace = new ChartSpace();
            PopulateChartSpace(chartSpace, options);
            return chartSpace;
        }

        /// <summary>ChartPart에 차트 XML을 생성하여 채웁니다 (레거시 호환).</summary>
        internal static void Build(
            ChartPart chartPart,
            string categoryFormula,
            IReadOnlyList<SeriesDefinition> seriesList,
            string title,
            bool showLegend)
        {
            chartPart.ChartSpace = BuildChartSpace(new ChartBuildOptions
            {
                CategoryFormula = categoryFormula,
                Series          = seriesList,
                ScatterSeries   = new List<ScatterSeriesDefinition>(),
                Title           = title,
                ShowLegend      = showLegend,
            });
        }

        // ── ChartSpace 구성 ──────────────────────────────────────────

        private static void PopulateChartSpace(ChartSpace chartSpace, ChartBuildOptions options)
        {
            chartSpace.AddNamespaceDeclaration("c", "http://schemas.openxmlformats.org/drawingml/2006/chart");
            chartSpace.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            chartSpace.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");

            chartSpace.Append(new Date1904 { Val = false });
            chartSpace.Append(new RoundedCorners { Val = false });

            // 차트 스타일 (옵션)
            if (options.ChartStyle.HasValue)
                chartSpace.Append(new Style { Val = (byte)options.ChartStyle.Value });

            var chart = new C.Chart();

            // 제목
            if (!string.IsNullOrEmpty(options.Title))
            {
                chart.Append(BuildTitleElement(options.Title));
                chart.Append(new AutoTitleDeleted { Val = false });
            }
            else
            {
                chart.Append(new AutoTitleDeleted { Val = true });
            }

            var plotArea = new PlotArea();
            plotArea.Append(new Layout());

            bool isScatterMode = options.ScatterSeries != null && options.ScatterSeries.Count > 0;

            if (isScatterMode)
            {
                // ── Scatter 전용 경로 ──────────────────────────────
                plotArea.Append(BuildScatterChart(options.ScatterSeries, options.ShowDataLabels));
                plotArea.Append(BuildScatterXValueAxis(ScatterXAxId, ScatterYAxId, options.XAxisTitle));
                plotArea.Append(BuildScatterYValueAxis(ScatterYAxId, ScatterXAxId,
                    options.YAxisMin, options.YAxisMax, options.YAxisTitle));
            }
            else
            {
                // ── 일반 차트 경로 ─────────────────────────────────
                bool hasSecondary = options.Series.Any(s => s.AxisPosition == AxisPosition.Secondary);
                bool isPieOnly    = options.Series.Count > 0
                                    && options.Series.All(s => s.ChartType == ChartType.Pie);

                var primaryGroups = options.Series
                    .Select((s, i) => (series: s, index: i))
                    .Where(x => x.series.AxisPosition == AxisPosition.Primary)
                    .GroupBy(x => x.series.ChartType);

                var secondaryGroups = options.Series
                    .Select((s, i) => (series: s, index: i))
                    .Where(x => x.series.AxisPosition == AxisPosition.Secondary)
                    .GroupBy(x => x.series.ChartType);

                foreach (var group in primaryGroups)
                    plotArea.Append(BuildChartElement(group.Key, group.ToList(),
                        options.CategoryFormula, PrimaryAxId, PrimaryValAxId,
                        isSecondary: false, options.ShowDataLabels));

                foreach (var group in secondaryGroups)
                    plotArea.Append(BuildChartElement(group.Key, group.ToList(),
                        options.CategoryFormula, SecondaryAxId, SecondaryValAxId,
                        isSecondary: true, options.ShowDataLabels));

                if (!isPieOnly)
                {
                    plotArea.Append(BuildCategoryAxis(PrimaryAxId, PrimaryValAxId,
                        options.XAxisTitle));
                    plotArea.Append(BuildValueAxis(PrimaryValAxId, PrimaryAxId,
                        isSecondary: false,
                        options.YAxisMin, options.YAxisMax, options.YAxisTitle));

                    if (hasSecondary)
                    {
                        plotArea.Append(BuildCategoryAxis(SecondaryAxId, SecondaryValAxId,
                            title: null, isSecondary: true));
                        plotArea.Append(BuildValueAxis(SecondaryValAxId, SecondaryAxId,
                            isSecondary: true,
                            options.SecondaryYAxisMin, options.SecondaryYAxisMax, title: null));
                    }
                }
            }

            chart.Append(plotArea);

            if (options.ShowLegend)
                chart.Append(BuildLegend());

            chart.Append(new PlotVisibleOnly { Val = true });
            chart.Append(new DisplayBlanksAs { Val = DisplayBlanksAsValues.Gap });
            chart.Append(new ShowDataLabelsOverMaximum { Val = false });

            chartSpace.Append(chart);
        }

        // ── 차트 요소 분기 ───────────────────────────────────────────

        private static OpenXmlCompositeElement BuildChartElement(
            ChartType type,
            List<(SeriesDefinition series, int index)> entries,
            string categoryFormula,
            uint axId, uint valAxId,
            bool isSecondary,
            bool showDataLabels)
        {
            switch (type)
            {
                case ChartType.Line:
                    return BuildLineChart(entries, categoryFormula, axId, valAxId, isSecondary, showDataLabels);
                case ChartType.Bar:
                    return BuildBarChart(entries, categoryFormula, axId, valAxId, isSecondary,
                        isHorizontal: true, showDataLabels);
                case ChartType.Area:
                    return BuildAreaChart(entries, categoryFormula, axId, valAxId, isSecondary,
                        GroupingValues.Standard, showDataLabels);
                case ChartType.AreaStacked:
                    return BuildAreaChart(entries, categoryFormula, axId, valAxId, isSecondary,
                        GroupingValues.Stacked, showDataLabels);
                case ChartType.Pie:
                    return BuildPieChart(entries, categoryFormula, showDataLabels);
                default: // Column
                    return BuildBarChart(entries, categoryFormula, axId, valAxId, isSecondary,
                        isHorizontal: false, showDataLabels);
            }
        }

        // ── BarChart / ColumnChart ───────────────────────────────────

        private static BarChart BuildBarChart(
            List<(SeriesDefinition series, int index)> entries,
            string categoryFormula,
            uint axId, uint valAxId,
            bool isSecondary, bool isHorizontal,
            bool showDataLabels)
        {
            var chart = new BarChart();
            chart.Append(new BarDirection
            {
                Val = isHorizontal ? BarDirectionValues.Bar : BarDirectionValues.Column
            });
            chart.Append(new BarGrouping { Val = BarGroupingValues.Clustered });
            chart.Append(new VaryColors { Val = false });

            foreach (var (series, index) in entries)
                chart.Append(BuildBarSeries(series, index, categoryFormula, showDataLabels));

            chart.Append(new AxisId { Val = axId });
            chart.Append(new AxisId { Val = valAxId });
            return chart;
        }

        private static BarChartSeries BuildBarSeries(
            SeriesDefinition series, int index, string categoryFormula, bool showDataLabels)
        {
            var ser = new BarChartSeries();
            ser.Append(new C.Index { Val = (uint)index });
            ser.Append(new Order { Val = (uint)index });
            ser.Append(BuildSeriesText(series.Name));
            if (series.HexColor != null)
                ser.Append(BuildShapeProperties(series.HexColor));
            if (showDataLabels)
                ser.Append(BuildDataLabels());
            ser.Append(BuildCategoryAxisData(categoryFormula));
            ser.Append(BuildValues(series.ValuesFormula));
            return ser;
        }

        // ── LineChart ────────────────────────────────────────────────

        private static LineChart BuildLineChart(
            List<(SeriesDefinition series, int index)> entries,
            string categoryFormula,
            uint axId, uint valAxId,
            bool isSecondary,
            bool showDataLabels)
        {
            var chart = new LineChart();
            chart.Append(new Grouping { Val = GroupingValues.Standard });
            chart.Append(new VaryColors { Val = false });

            foreach (var (series, index) in entries)
                chart.Append(BuildLineChartSeries(series, index, categoryFormula, showDataLabels));

            chart.Append(new ShowMarker { Val = true });
            chart.Append(new Smooth { Val = false });
            chart.Append(new AxisId { Val = axId });
            chart.Append(new AxisId { Val = valAxId });
            return chart;
        }

        private static LineChartSeries BuildLineChartSeries(
            SeriesDefinition series, int index, string categoryFormula, bool showDataLabels)
        {
            var ser = new LineChartSeries();
            ser.Append(new C.Index { Val = (uint)index });
            ser.Append(new Order { Val = (uint)index });
            ser.Append(BuildSeriesText(series.Name));
            if (series.HexColor != null)
                ser.Append(BuildShapeProperties(series.HexColor));
            if (series.MarkerStyle != MarkerStyle.Auto)
                ser.Append(BuildMarker(series.MarkerStyle));
            if (showDataLabels)
                ser.Append(BuildDataLabels());
            ser.Append(BuildCategoryAxisData(categoryFormula));
            ser.Append(BuildValues(series.ValuesFormula));
            return ser;
        }

        // ── AreaChart ────────────────────────────────────────────────

        private static AreaChart BuildAreaChart(
            List<(SeriesDefinition series, int index)> entries,
            string categoryFormula,
            uint axId, uint valAxId,
            bool isSecondary,
            GroupingValues grouping,
            bool showDataLabels)
        {
            var chart = new AreaChart();
            chart.Append(new Grouping { Val = grouping });
            chart.Append(new VaryColors { Val = false });

            foreach (var (series, index) in entries)
                chart.Append(BuildAreaChartSeries(series, index, categoryFormula, showDataLabels));

            chart.Append(new AxisId { Val = axId });
            chart.Append(new AxisId { Val = valAxId });
            return chart;
        }

        private static AreaChartSeries BuildAreaChartSeries(
            SeriesDefinition series, int index, string categoryFormula, bool showDataLabels)
        {
            var ser = new AreaChartSeries();
            ser.Append(new C.Index { Val = (uint)index });
            ser.Append(new Order { Val = (uint)index });
            ser.Append(BuildSeriesText(series.Name));
            if (series.HexColor != null)
                ser.Append(BuildShapeProperties(series.HexColor));
            if (showDataLabels)
                ser.Append(BuildDataLabels());
            ser.Append(BuildCategoryAxisData(categoryFormula));
            ser.Append(BuildValues(series.ValuesFormula));
            return ser;
        }

        // ── PieChart ─────────────────────────────────────────────────

        private static PieChart BuildPieChart(
            List<(SeriesDefinition series, int index)> entries,
            string categoryFormula,
            bool showDataLabels)
        {
            var chart = new PieChart();
            chart.Append(new VaryColors { Val = true });

            foreach (var (series, index) in entries)
                chart.Append(BuildPieSeries(series, index, categoryFormula, showDataLabels));

            chart.Append(new FirstSliceAngle { Val = 0 });
            // Pie chart에는 <c:axId> 없음
            return chart;
        }

        private static PieChartSeries BuildPieSeries(
            SeriesDefinition series, int index, string categoryFormula, bool showDataLabels)
        {
            var ser = new PieChartSeries();
            ser.Append(new C.Index { Val = (uint)index });
            ser.Append(new Order { Val = (uint)index });
            ser.Append(BuildSeriesText(series.Name));
            if (series.HexColor != null)
                ser.Append(BuildShapeProperties(series.HexColor));
            if (showDataLabels)
                ser.Append(BuildDataLabels());
            ser.Append(BuildCategoryAxisData(categoryFormula));
            ser.Append(BuildValues(series.ValuesFormula));
            return ser;
        }

        // ── ScatterChart ─────────────────────────────────────────────

        private static ScatterChart BuildScatterChart(
            List<ScatterSeriesDefinition> scatterSeries,
            bool showDataLabels)
        {
            var chart = new ScatterChart();
            chart.Append(new ScatterStyle { Val = ScatterStyleValues.LineMarker });
            chart.Append(new VaryColors { Val = false });

            for (int i = 0; i < scatterSeries.Count; i++)
                chart.Append(BuildScatterSeries(scatterSeries[i], i, showDataLabels));

            chart.Append(new AxisId { Val = ScatterXAxId });
            chart.Append(new AxisId { Val = ScatterYAxId });
            return chart;
        }

        private static ScatterChartSeries BuildScatterSeries(
            ScatterSeriesDefinition series, int index, bool showDataLabels)
        {
            var ser = new ScatterChartSeries();
            ser.Append(new C.Index { Val = (uint)index });
            ser.Append(new Order { Val = (uint)index });
            ser.Append(BuildSeriesText(series.Name));
            if (series.HexColor != null)
                ser.Append(BuildShapeProperties(series.HexColor));
            if (showDataLabels)
                ser.Append(BuildDataLabels());
            ser.Append(new XValues(new NumberReference(new Formula(series.XValuesFormula))));
            ser.Append(new YValues(new NumberReference(new Formula(series.YValuesFormula))));
            return ser;
        }

        // ── 축 ───────────────────────────────────────────────────────

        private static CategoryAxis BuildCategoryAxis(
            uint axId, uint crossAxId,
            string title,
            bool isSecondary = false)
        {
            var ax = new CategoryAxis();
            ax.Append(new AxisId { Val = axId });
            ax.Append(new Scaling(new Orientation { Val = OrientationValues.MinMax }));
            ax.Append(new Delete { Val = isSecondary });
            ax.Append(new C.AxisPosition
            {
                Val = isSecondary ? AxisPositionValues.Top : AxisPositionValues.Bottom
            });
            if (!string.IsNullOrEmpty(title))
                ax.Append(BuildTitleElement(title));
            ax.Append(new NumberingFormat { FormatCode = "General", SourceLinked = true });
            ax.Append(new MajorTickMark { Val = TickMarkValues.Outside });
            ax.Append(new MinorTickMark { Val = TickMarkValues.None });
            ax.Append(new TickLabelPosition { Val = TickLabelPositionValues.NextTo });
            ax.Append(new CrossingAxis { Val = crossAxId });
            ax.Append(new Crosses { Val = CrossesValues.AutoZero });
            ax.Append(new AutoLabeled { Val = true });
            ax.Append(new LabelAlignment { Val = LabelAlignmentValues.Center });
            return ax;
        }

        private static ValueAxis BuildValueAxis(
            uint axId, uint crossAxId,
            bool isSecondary,
            double? min, double? max,
            string title)
        {
            var ax = new ValueAxis();
            ax.Append(new AxisId { Val = axId });
            ax.Append(BuildScaling(min, max));
            ax.Append(new Delete { Val = false });
            ax.Append(new C.AxisPosition
            {
                Val = isSecondary ? AxisPositionValues.Right : AxisPositionValues.Left
            });
            if (!string.IsNullOrEmpty(title))
                ax.Append(BuildTitleElement(title));
            ax.Append(new NumberingFormat { FormatCode = "General", SourceLinked = true });
            ax.Append(new MajorTickMark { Val = TickMarkValues.Outside });
            ax.Append(new MinorTickMark { Val = TickMarkValues.None });
            ax.Append(new TickLabelPosition { Val = TickLabelPositionValues.NextTo });
            ax.Append(new CrossingAxis { Val = crossAxId });
            ax.Append(new Crosses { Val = CrossesValues.AutoZero });
            ax.Append(new CrossBetween { Val = CrossBetweenValues.Between });
            return ax;
        }

        // Scatter 전용 X ValueAxis (하단)
        private static ValueAxis BuildScatterXValueAxis(
            uint axId, uint crossAxId, string title)
        {
            var ax = new ValueAxis();
            ax.Append(new AxisId { Val = axId });
            ax.Append(BuildScaling(null, null));
            ax.Append(new Delete { Val = false });
            ax.Append(new C.AxisPosition { Val = AxisPositionValues.Bottom });
            if (!string.IsNullOrEmpty(title))
                ax.Append(BuildTitleElement(title));
            ax.Append(new NumberingFormat { FormatCode = "General", SourceLinked = true });
            ax.Append(new MajorTickMark { Val = TickMarkValues.Outside });
            ax.Append(new MinorTickMark { Val = TickMarkValues.None });
            ax.Append(new TickLabelPosition { Val = TickLabelPositionValues.NextTo });
            ax.Append(new CrossingAxis { Val = crossAxId });
            ax.Append(new Crosses { Val = CrossesValues.AutoZero });
            ax.Append(new CrossBetween { Val = CrossBetweenValues.Between });
            return ax;
        }

        // Scatter 전용 Y ValueAxis (좌측)
        private static ValueAxis BuildScatterYValueAxis(
            uint axId, uint crossAxId,
            double? min, double? max,
            string title)
        {
            var ax = new ValueAxis();
            ax.Append(new AxisId { Val = axId });
            ax.Append(BuildScaling(min, max));
            ax.Append(new Delete { Val = false });
            ax.Append(new C.AxisPosition { Val = AxisPositionValues.Left });
            if (!string.IsNullOrEmpty(title))
                ax.Append(BuildTitleElement(title));
            ax.Append(new NumberingFormat { FormatCode = "General", SourceLinked = true });
            ax.Append(new MajorTickMark { Val = TickMarkValues.Outside });
            ax.Append(new MinorTickMark { Val = TickMarkValues.None });
            ax.Append(new TickLabelPosition { Val = TickLabelPositionValues.NextTo });
            ax.Append(new CrossingAxis { Val = crossAxId });
            ax.Append(new Crosses { Val = CrossesValues.AutoZero });
            ax.Append(new CrossBetween { Val = CrossBetweenValues.Between });
            return ax;
        }

        // ── 공용 빌더 헬퍼 ───────────────────────────────────────────

        private static Scaling BuildScaling(double? min, double? max)
        {
            var scaling = new Scaling();
            scaling.Append(new Orientation { Val = OrientationValues.MinMax });
            if (max.HasValue)
                scaling.Append(new MaxAxisValue { Val = max.Value });
            if (min.HasValue)
                scaling.Append(new MinAxisValue { Val = min.Value });
            return scaling;
        }

        private static SeriesText BuildSeriesText(string name)
        {
            // CT_SerTx: c:v (NumericValue) 만 허용 — c:strLit 금지
            return new SeriesText(new NumericValue(name));
        }

        private static CategoryAxisData BuildCategoryAxisData(string formula)
        {
            return new CategoryAxisData(new StringReference(new Formula(formula)));
        }

        private static Values BuildValues(string formula)
        {
            return new Values(new NumberReference(new Formula(formula)));
        }

        private static C.ChartShapeProperties BuildShapeProperties(string hexColor)
        {
            return new C.ChartShapeProperties(
                new A.SolidFill(
                    new A.RgbColorModelHex { Val = hexColor }
                )
            );
        }

        private static DataLabels BuildDataLabels()
        {
            return new DataLabels(
                new NumberingFormat { FormatCode = "General", SourceLinked = true },
                new ShowLegendKey   { Val = false },
                new ShowValue       { Val = true  },
                new ShowCategoryName{ Val = false },
                new ShowSeriesName  { Val = false },
                new ShowPercent     { Val = false },
                new ShowBubbleSize  { Val = false }
            );
        }

        private static Marker BuildMarker(MarkerStyle style)
        {
            C.MarkerStyleValues val;
            switch (style)
            {
                case MarkerStyle.None:     val = C.MarkerStyleValues.None;     break;
                case MarkerStyle.Circle:   val = C.MarkerStyleValues.Circle;   break;
                case MarkerStyle.Square:   val = C.MarkerStyleValues.Square;   break;
                case MarkerStyle.Diamond:  val = C.MarkerStyleValues.Diamond;  break;
                case MarkerStyle.Triangle: val = C.MarkerStyleValues.Triangle; break;
                default:                   val = C.MarkerStyleValues.Auto;     break;
            }
            return new Marker(
                new C.Symbol { Val = val },
                new C.Size   { Val = 5  }
            );
        }

        private static C.Title BuildTitleElement(string text)
        {
            return new C.Title(
                new ChartText(
                    new RichText(
                        new A.BodyProperties(),
                        new A.ListStyle(),
                        new A.Paragraph(
                            new A.Run(
                                new A.RunProperties { Language = "ko-KR" },
                                new A.Text(text)
                            )
                        )
                    )
                ),
                new Overlay { Val = false }
            );
        }

        private static Legend BuildLegend()
        {
            return new Legend(
                new LegendPosition { Val = LegendPositionValues.Bottom },
                new Overlay { Val = false }
            );
        }
    }
}
