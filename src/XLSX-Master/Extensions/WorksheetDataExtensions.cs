using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Reflection;
using ClosedXML.Excel;

namespace XlsxMaster.Extensions
{
    /// <summary>
    /// <see cref="IXLWorksheet"/>에 데이터를 바인딩하는 XLSX-Master 확장 메서드입니다.
    /// </summary>
    public static class WorksheetDataExtensions
    {
        /// <summary>
        /// <c>IEnumerable&lt;T&gt;</c>를 워크시트에 삽입합니다.
        /// 퍼블릭 프로퍼티명이 헤더로, 값 타입에 따라 셀 서식이 자동 적용됩니다.
        /// </summary>
        /// <typeparam name="T">데이터 모델 타입. 읽기 가능한 퍼블릭 프로퍼티가 하나 이상 있어야 합니다.</typeparam>
        /// <param name="ws">삽입 대상 워크시트</param>
        /// <param name="data">삽입할 데이터</param>
        /// <param name="startRow">헤더 행 번호 (기본값: 1)</param>
        /// <param name="startColumn">시작 열 번호 (기본값: 1)</param>
        /// <returns>헤더 행을 포함한 삽입된 전체 범위</returns>
        /// <example>
        /// <code>
        /// var range = ws.InsertMasterTable(salesList, startRow: 1, startColumn: 1);
        /// range.AddExcelTable("SalesTable");
        /// </code>
        /// </example>
        public static IXLRange InsertMasterTable<T>(
            this IXLWorksheet ws,
            IEnumerable<T> data,
            int startRow = 1,
            int startColumn = 1)
        {
            if (ws == null) throw new ArgumentNullException(nameof(ws));
            if (data == null) throw new ArgumentNullException(nameof(data));

            var props = GetReadableProperties(typeof(T));
            if (props.Length == 0)
                throw new ArgumentException(
                    $"타입 '{typeof(T).Name}'에 읽기 가능한 공개 프로퍼티가 없습니다.", nameof(data));

            // 헤더 행
            for (int col = 0; col < props.Length; col++)
                ws.Cell(startRow, startColumn + col).Value = props[col].Name;

            // 데이터 행
            var items = data.ToList();
            for (int row = 0; row < items.Count; row++)
                for (int col = 0; col < props.Length; col++)
                    SetCellValue(
                        ws.Cell(startRow + 1 + row, startColumn + col),
                        props[col].GetValue(items[row]),
                        props[col].PropertyType);

            int lastRow = startRow + items.Count;
            int lastCol = startColumn + props.Length - 1;
            return ws.Range(startRow, startColumn, lastRow, lastCol);
        }

        /// <summary>
        /// <see cref="DataTable"/>을 워크시트에 삽입합니다.
        /// 열 이름이 헤더로, 열 타입에 따라 셀 서식이 자동 적용됩니다.
        /// </summary>
        /// <param name="ws">삽입 대상 워크시트</param>
        /// <param name="data">삽입할 DataTable</param>
        /// <param name="startRow">헤더 행 번호 (기본값: 1)</param>
        /// <param name="startColumn">시작 열 번호 (기본값: 1)</param>
        /// <returns>헤더 행을 포함한 삽입된 전체 범위</returns>
        public static IXLRange InsertMasterTable(
            this IXLWorksheet ws,
            DataTable data,
            int startRow = 1,
            int startColumn = 1)
        {
            if (ws == null) throw new ArgumentNullException(nameof(ws));
            if (data == null) throw new ArgumentNullException(nameof(data));

            // 헤더 행
            for (int col = 0; col < data.Columns.Count; col++)
                ws.Cell(startRow, startColumn + col).Value = data.Columns[col].ColumnName;

            // 데이터 행
            for (int row = 0; row < data.Rows.Count; row++)
                for (int col = 0; col < data.Columns.Count; col++)
                {
                    var raw = data.Rows[row][col];
                    SetCellValue(
                        ws.Cell(startRow + 1 + row, startColumn + col),
                        raw == DBNull.Value ? null : raw,
                        data.Columns[col].DataType);
                }

            int lastRow = startRow + data.Rows.Count;
            int lastCol = startColumn + data.Columns.Count - 1;
            return ws.Range(startRow, startColumn, lastRow, lastCol);
        }

        // ─────────────────────────────────────────────────────────────
        // 내부 헬퍼
        // ─────────────────────────────────────────────────────────────

        private static PropertyInfo[] GetReadableProperties(Type type) =>
            type.GetProperties(BindingFlags.Public | BindingFlags.Instance)
                .Where(p => p.CanRead && p.GetIndexParameters().Length == 0)
                .ToArray();

        private static void SetCellValue(IXLCell cell, object value, Type type)
        {
            if (value == null)
                return; // 빈 셀 유지

            var t = Nullable.GetUnderlyingType(type) ?? type;

            if (t == typeof(string) || t == typeof(char))
            {
                cell.Value = value.ToString();
            }
            else if (t == typeof(bool))
            {
                cell.Value = (bool)value;
            }
            else if (t == typeof(DateTime))
            {
                cell.Value = (DateTime)value;
                cell.Style.NumberFormat.Format = "yyyy-mm-dd";
            }
            else if (t == typeof(decimal))
            {
                cell.Value = (double)(decimal)value;
                cell.Style.NumberFormat.Format = "#,##0.00";
            }
            else if (t == typeof(double))
            {
                cell.Value = (double)value;
                cell.Style.NumberFormat.Format = "#,##0.00";
            }
            else if (t == typeof(float))
            {
                cell.Value = (double)(float)value;
                cell.Style.NumberFormat.Format = "#,##0.00";
            }
            else if (t == typeof(int) || t == typeof(long) || t == typeof(short) ||
                     t == typeof(byte) || t == typeof(uint) || t == typeof(ulong))
            {
                cell.Value = Convert.ToDouble(value);
                cell.Style.NumberFormat.Format = "#,##0";
            }
            else
            {
                cell.Value = value.ToString();
            }
        }
    }
}
