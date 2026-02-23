namespace XlsxMaster.Charts
{
    /// <summary>
    /// 차트 계열에 적용할 차트 유형을 정의합니다.
    /// </summary>
    public enum ChartType
    {
        /// <summary>세로 막대형</summary>
        Column,

        /// <summary>가로 막대형</summary>
        Bar,

        /// <summary>꺾은선형</summary>
        Line,

        /// <summary>영역형</summary>
        Area,

        /// <summary>누적 영역형</summary>
        AreaStacked,

        /// <summary>원형</summary>
        Pie,

        /// <summary>분산형</summary>
        Scatter
    }
}
