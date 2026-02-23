using System;
using System.Linq;
using ClosedXML.Excel;

namespace XlsxMaster.Extensions
{
    /// <summary>
    /// <see cref="IXLRange"/>에 대한 XLSX-Master 확장 메서드입니다.
    /// </summary>
    public static class RangeExtensions
    {
        /// <summary>
        /// 지정된 범위를 Excel Table 객체로 변환합니다.
        /// 자동 필터와 테이블 스타일이 적용됩니다.
        /// </summary>
        /// <param name="range">테이블로 변환할 범위 (헤더 행 포함)</param>
        /// <param name="name">테이블 이름. null이면 "Table{n}" 형식으로 자동 생성됩니다.</param>
        /// <param name="theme">테이블 스타일 (기본값: TableStyleMedium9)</param>
        /// <returns>생성된 <see cref="IXLTable"/></returns>
        /// <example>
        /// <code>
        /// ws.InsertMasterTable(salesList)
        ///   .AddExcelTable("SalesTable", XLTableTheme.TableStyleMedium2);
        /// </code>
        /// </example>
        public static IXLTable AddExcelTable(
            this IXLRange range,
            string name = null,
            XLTableTheme theme = null)
        {
            if (range == null) throw new ArgumentNullException(nameof(range));

            var tableName = string.IsNullOrWhiteSpace(name)
                ? $"Table{range.Worksheet.Tables.Count() + 1}"
                : name;

            var table = range.CreateTable(tableName);
            table.Theme = theme ?? XLTableTheme.TableStyleMedium9;
            table.ShowAutoFilter = true;
            return table;
        }
    }
}
