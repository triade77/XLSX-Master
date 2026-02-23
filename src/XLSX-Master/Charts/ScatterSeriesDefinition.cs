namespace XlsxMaster.Charts
{
    /// <summary>
    /// Scatter(분산형) 차트의 단일 계열 정의입니다.
    /// X값과 Y값 범위를 각각 지정합니다.
    /// </summary>
    internal sealed class ScatterSeriesDefinition
    {
        public string Name           { get; }
        public string XValuesFormula { get; }
        public string YValuesFormula { get; }
        public string HexColor       { get; set; }  // null = 기본색

        internal ScatterSeriesDefinition(
            string name,
            string xValuesFormula,
            string yValuesFormula,
            string hexColor = null)
        {
            Name           = name;
            XValuesFormula = xValuesFormula;
            YValuesFormula = yValuesFormula;
            HexColor       = hexColor;
        }
    }
}
