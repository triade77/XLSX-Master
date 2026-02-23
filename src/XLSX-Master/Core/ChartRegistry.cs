using System.Collections.Generic;
using System.Runtime.CompilerServices;
using ClosedXML.Excel;
using XlsxMaster.Charts;

namespace XlsxMaster.Core
{
    /// <summary>
    /// IXLWorkbook 인스턴스에 등록된 <see cref="MasterChartBuilder"/> 목록을 관리합니다.
    /// ConditionalWeakTable을 사용하여 워크북 GC 시 자동 해제됩니다.
    /// </summary>
    internal static class ChartRegistry
    {
        private static readonly ConditionalWeakTable<IXLWorkbook, List<MasterChartBuilder>> _table
            = new ConditionalWeakTable<IXLWorkbook, List<MasterChartBuilder>>();

        public static void Register(IXLWorkbook workbook, MasterChartBuilder builder)
        {
            var list = _table.GetOrCreateValue(workbook);
            list.Add(builder);
        }

        public static IReadOnlyList<MasterChartBuilder> GetBuilders(IXLWorkbook workbook)
        {
            if (_table.TryGetValue(workbook, out var list))
                return list;
            return new List<MasterChartBuilder>();
        }
    }
}
