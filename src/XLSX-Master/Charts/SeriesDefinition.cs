namespace XlsxMaster.Charts
{
    /// <summary>
    /// 차트에 추가될 단일 데이터 계열의 정의입니다.
    /// </summary>
    public sealed class SeriesDefinition
    {
        /// <summary>범례에 표시될 계열 이름</summary>
        public string Name { get; set; }

        /// <summary>값 데이터 셀 범위 (예: "B2:B13")</summary>
        public string ValuesRange { get; set; }

        /// <summary>이 계열에 적용할 차트 유형</summary>
        public ChartType ChartType { get; set; }

        /// <summary>이 계열이 표시될 Y축 위치</summary>
        public AxisPosition AxisPosition { get; set; }

        /// <summary>시트명이 포함된 완전한 Values 수식 (예: "Sheet1!$B$2:$B$13")</summary>
        public string ValuesFormula { get; set; }

        /// <summary>계열 색상 hex 코드 (예: "FF0000"). null이면 Excel 기본색 사용.</summary>
        internal string HexColor { get; set; }

        /// <summary>꺾은선 계열 마커 스타일 (기본값: Auto)</summary>
        internal MarkerStyle MarkerStyle { get; set; } = MarkerStyle.Auto;

        public SeriesDefinition(
            string name,
            string valuesRange,
            ChartType chartType,
            AxisPosition axisPosition,
            string valuesFormula)
        {
            Name = name;
            ValuesRange = valuesRange;
            ChartType = chartType;
            AxisPosition = axisPosition;
            ValuesFormula = valuesFormula;
        }
    }
}
