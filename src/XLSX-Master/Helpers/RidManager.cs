using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;

namespace XlsxMaster.Helpers
{
    /// <summary>
    /// Open XML 파트의 Relationship ID(rId) 충돌 없이 다음 번호를 할당합니다.
    /// </summary>
    public static class RidManager
    {
        private static readonly Regex RidPattern = new Regex(@"rId(\d+)", RegexOptions.Compiled);

        /// <summary>
        /// 지정된 파트의 기존 관계(Relationship)를 스캔하여 다음에 사용 가능한 rId를 반환합니다.
        /// </summary>
        public static string GetNextRid(OpenXmlPart part)
        {
            var existingIds = part.Parts
                .Select(p => p.RelationshipId)
                .Concat(part.ExternalRelationships.Select(r => r.Id))
                .Concat(part.HyperlinkRelationships.Select(r => r.Id));

            int maxNum = ExtractMaxNumber(existingIds);
            return $"rId{maxNum + 1}";
        }

        /// <summary>
        /// 지정된 SpreadsheetDocument 내 WorkbookPart에서 다음 rId를 반환합니다.
        /// </summary>
        public static string GetNextRid(WorkbookPart workbookPart)
        {
            var existingIds = workbookPart.Parts
                .Select(p => p.RelationshipId)
                .Concat(workbookPart.ExternalRelationships.Select(r => r.Id))
                .Concat(workbookPart.HyperlinkRelationships.Select(r => r.Id));

            int maxNum = ExtractMaxNumber(existingIds);
            return $"rId{maxNum + 1}";
        }

        private static int ExtractMaxNumber(IEnumerable<string> ids)
        {
            int max = 0;
            foreach (var id in ids)
            {
                var m = RidPattern.Match(id ?? string.Empty);
                if (m.Success && int.TryParse(m.Groups[1].Value, out int num))
                    max = Math.Max(max, num);
            }
            return max;
        }
    }
}
