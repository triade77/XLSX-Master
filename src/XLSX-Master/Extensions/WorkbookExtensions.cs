using System;
using System.IO;
using ClosedXML.Excel;
using XlsxMaster.Core;

namespace XlsxMaster.Extensions
{
    /// <summary>
    /// <see cref="IXLWorkbook"/>에 대한 XLSX-Master 확장 메서드입니다.
    /// </summary>
    public static class WorkbookExtensions
    {
        /// <summary>
        /// 워크북 데이터와 등록된 모든 차트를 포함하여 지정된 경로에 파일을 저장합니다.
        /// </summary>
        /// <param name="workbook">저장할 워크북</param>
        /// <param name="filePath">저장 경로 (예: "C:\Report.xlsx")</param>
        public static void SaveWithCharts(this IXLWorkbook workbook, string filePath)
        {
            if (workbook == null) throw new ArgumentNullException(nameof(workbook));
            if (string.IsNullOrWhiteSpace(filePath)) throw new ArgumentException("파일 경로를 지정하세요.", nameof(filePath));

            var resultStream = BuildStream(workbook);
            File.WriteAllBytes(filePath, resultStream.ToArray());
            resultStream.Dispose();
        }

        /// <summary>
        /// 워크북 데이터와 등록된 모든 차트를 포함하여 지정된 스트림에 씁니다.
        /// </summary>
        /// <param name="workbook">저장할 워크북</param>
        /// <param name="destination">쓰기 가능한 대상 스트림</param>
        public static void SaveWithCharts(this IXLWorkbook workbook, Stream destination)
        {
            if (workbook == null) throw new ArgumentNullException(nameof(workbook));
            if (destination == null) throw new ArgumentNullException(nameof(destination));
            if (!destination.CanWrite) throw new ArgumentException("쓰기 가능한 스트림이 필요합니다.", nameof(destination));

            using (var ms = BuildStream(workbook))
                ms.WriteTo(destination);
        }

        /// <summary>
        /// 워크북 데이터와 등록된 모든 차트를 포함하는 <see cref="MemoryStream"/>을 반환합니다.
        /// </summary>
        public static MemoryStream SaveWithChartsToStream(this IXLWorkbook workbook)
        {
            if (workbook == null) throw new ArgumentNullException(nameof(workbook));
            return BuildStream(workbook);
        }

        private static MemoryStream BuildStream(IXLWorkbook workbook)
        {
            var builders = ChartRegistry.GetBuilders(workbook);

            // 1. ClosedXML로 기본 워크북 저장
            var current = new MemoryStream();
            workbook.SaveAs(current);
            current.Position = 0;

            // 2. 등록된 차트 빌더를 순서대로 주입
            foreach (var builder in builders)
            {
                var next = builder.InjectInto(current);
                current.Dispose();
                current = next;
                current.Position = 0;
            }

            return current;
        }
    }
}
