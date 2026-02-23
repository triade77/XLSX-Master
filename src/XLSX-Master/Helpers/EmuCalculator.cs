using System;
using System.Text.RegularExpressions;

namespace XlsxMaster.Helpers
{
    /// <summary>
    /// 셀 주소 기반 Anchor 문자열을 Open XML EMU(English Metric Unit) 좌표로 변환합니다.
    /// </summary>
    public static class EmuCalculator
    {
        // 기본 열/행 크기 (픽셀 기준 Excel 기본값 → EMU)
        // 1 pixel = 9525 EMU (96 DPI 기준)
        private const long DefaultColumnWidthPixels = 64;   // Excel 기본 열 너비 ≈ 8 characters
        private const long DefaultRowHeightPixels = 20;     // Excel 기본 행 높이 15pt ≈ 20px

        public const long EmuPerPixel = 9525L;
        public const long DefaultColumnWidthEmu = DefaultColumnWidthPixels * EmuPerPixel; // 609600
        public const long DefaultRowHeightEmu = DefaultRowHeightPixels * EmuPerPixel;     // 190500

        /// <summary>
        /// "E1:M20" 형식의 앵커 문자열을 파싱하여 (colFrom, rowFrom, colTo, rowTo) 인덱스로 반환합니다.
        /// 반환값은 0-based 인덱스입니다.
        /// </summary>
        public static (int colFrom, int rowFrom, int colTo, int rowTo) ParseAnchor(string anchor)
        {
            if (string.IsNullOrWhiteSpace(anchor))
                throw new ArgumentException("앵커 문자열이 비어있습니다.", nameof(anchor));

            var match = Regex.Match(anchor.Trim().ToUpperInvariant(),
                @"^([A-Z]+)(\d+):([A-Z]+)(\d+)$");

            if (!match.Success)
                throw new ArgumentException($"앵커 형식이 잘못되었습니다: '{anchor}'. 예: \"E1:M20\"", nameof(anchor));

            int colFrom = ColumnLetterToIndex(match.Groups[1].Value); // 0-based
            int rowFrom = int.Parse(match.Groups[2].Value) - 1;        // 0-based
            int colTo = ColumnLetterToIndex(match.Groups[3].Value);
            int rowTo = int.Parse(match.Groups[4].Value) - 1;

            if (colFrom > colTo || rowFrom > rowTo)
                throw new ArgumentException(
                    $"앵커 범위가 역순입니다: '{anchor}'. 시작 셀이 끝 셀보다 앞이어야 합니다.", nameof(anchor));

            return (colFrom, rowFrom, colTo, rowTo);
        }

        /// <summary>열 문자(A, B, AA, …)를 0-based 인덱스로 변환합니다.</summary>
        public static int ColumnLetterToIndex(string letters)
        {
            int result = 0;
            foreach (char c in letters)
                result = result * 26 + (c - 'A' + 1);
            return result - 1; // 0-based
        }

        /// <summary>0-based 열 인덱스를 기준으로 EMU X 오프셋을 계산합니다.</summary>
        public static long ColumnIndexToEmu(int colIndex)
            => colIndex * DefaultColumnWidthEmu;

        /// <summary>0-based 행 인덱스를 기준으로 EMU Y 오프셋을 계산합니다.</summary>
        public static long RowIndexToEmu(int rowIndex)
            => rowIndex * DefaultRowHeightEmu;
    }
}
