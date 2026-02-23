using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using XlsxMaster.Charts;
using XlsxMaster.Helpers;

namespace XlsxMaster.Core
{
    /// <summary>
    /// ClosedXML이 생성한 xlsx 스트림에 차트를 주입하는 핵심 엔진입니다.
    /// ZipArchive를 직접 조작하여 Excel 표준 경로(xl/charts/)와 상대 URI를 보장합니다.
    /// </summary>
    internal static class XlsxChartInjector
    {
        // OPC 패키지 관계 네임스페이스
        private static readonly XNamespace PkgRels =
            "http://schemas.openxmlformats.org/package/2006/relationships";
        // 워크북/시트 XML 네임스페이스
        private static readonly XNamespace SS =
            "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        // r: 관계 속성 네임스페이스
        private static readonly XNamespace RAttr =
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        private const string RT_DRAWING =
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing";
        private const string RT_CHART =
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart";
        private const string CT_DRAWING =
            "application/vnd.openxmlformats-officedocument.drawing+xml";
        private const string CT_CHART =
            "application/vnd.openxmlformats-officedocument.drawingml.chart+xml";

        internal static MemoryStream Inject(
            Stream source,
            string sheetName,
            string anchor,
            ChartBuildOptions options)
        {
            var (colFrom, rowFrom, colTo, rowTo) = EmuCalculator.ParseAnchor(anchor);

            var output = new MemoryStream();
            source.Position = 0;
            source.CopyTo(output);

            using (var zip = new ZipArchive(output, ZipArchiveMode.Update, leaveOpen: true))
            {
                // 1. 시트 파일 경로 확인
                string sheetPath = FindSheetPath(zip, sheetName);
                string sheetRelsPath = GetRelsPath(sheetPath);

                // 2. 이 시트에 이미 Drawing이 있는지 확인
                string drawingPath = FindExistingDrawingPath(zip, sheetRelsPath, sheetPath);
                string drawingRelsPath;

                if (drawingPath == null)
                {
                    // 새 Drawing 생성
                    int drawingNum = GetNextNumber(zip, "xl/drawings/drawing");
                    drawingPath = $"xl/drawings/drawing{drawingNum}.xml";
                    drawingRelsPath = GetRelsPath(drawingPath);

                    // 빈 drawing XML 작성
                    WriteEntry(zip, drawingPath, BuildEmptyDrawingXml());
                    WriteEntry(zip, drawingRelsPath, BuildEmptyRelsXml());

                    // 시트 rels에 drawing 관계 추가
                    string drawingRelId = GetNextRelId(zip, sheetRelsPath);
                    string drawingTarget = MakeRelativePath(sheetPath, drawingPath);
                    AddRel(zip, sheetRelsPath, drawingRelId, RT_DRAWING, drawingTarget);

                    // 시트 XML에 <drawing r:id="..."/> 요소 추가
                    AddDrawingElementToSheet(zip, sheetPath, drawingRelId);

                    // Content_Types에 drawing 등록
                    EnsureContentType(zip, "/" + drawingPath, CT_DRAWING);
                }
                else
                {
                    drawingRelsPath = GetRelsPath(drawingPath);
                }

                // 3. 차트 파일 추가 (표준 위치: xl/charts/chartN.xml)
                int chartNum = GetNextNumber(zip, "xl/charts/chart");
                string chartPath = $"xl/charts/chart{chartNum}.xml";
                WriteEntry(zip, chartPath, BuildChartXml(options));
                EnsureContentType(zip, "/" + chartPath, CT_CHART);

                // 4. drawing rels에 chart 관계 추가 (상대 경로: ../charts/chartN.xml)
                string chartRelId = GetNextRelId(zip, drawingRelsPath);
                string chartTarget = MakeRelativePath(drawingPath, chartPath);
                AddRel(zip, drawingRelsPath, chartRelId, RT_CHART, chartTarget);

                // 5. drawing XML에 TwoCellAnchor 추가
                AppendAnchorToDrawing(zip, drawingPath, colFrom, rowFrom, colTo, rowTo, chartRelId, chartNum);
            }

            output.Position = 0;
            return output;
        }

        // ─── 시트 탐색 ──────────────────────────────────────────────────────

        private static string FindSheetPath(ZipArchive zip, string sheetName)
        {
            var wb = ParseEntry(zip, "xl/workbook.xml");

            var allSheets = wb.Descendants(SS + "sheet").ToList();
            var sheetEl = allSheets.FirstOrDefault(s => (string)s.Attribute("name") == sheetName);
            if (sheetEl == null)
            {
                var availableSheets = allSheets
                    .Select(s => (string)s.Attribute("name"))
                    .Where(n => n != null);
                throw new KeyNotFoundException(
                    $"시트를 찾을 수 없습니다: '{sheetName}'. " +
                    $"사용 가능한 시트: {string.Join(", ", availableSheets)}");
            }

            string rId = (string)sheetEl.Attribute(RAttr + "id");

            var wbRels = ParseEntry(zip, "xl/_rels/workbook.xml.rels");
            string target = wbRels.Descendants(PkgRels + "Relationship")
                .Where(r => (string)r.Attribute("Id") == rId)
                .Select(r => (string)r.Attribute("Target"))
                .FirstOrDefault()
                ?? throw new InvalidOperationException($"workbook.xml.rels에서 '{rId}'를 찾을 수 없습니다.");

            return NormalizePath("xl/", target);
        }

        private static string FindExistingDrawingPath(ZipArchive zip, string sheetRelsPath, string sheetPath)
        {
            string content = ReadEntry(zip, sheetRelsPath);
            if (content == null) return null;

            var doc = XDocument.Parse(content);
            string target = doc.Descendants(PkgRels + "Relationship")
                .Where(r => ((string)r.Attribute("Type") ?? "").EndsWith("/drawing"))
                .Select(r => (string)r.Attribute("Target"))
                .FirstOrDefault();

            return target == null ? null : NormalizePath(GetDir(sheetPath) + "/", target);
        }

        // ─── 번호 할당 ──────────────────────────────────────────────────────

        private static int GetNextNumber(ZipArchive zip, string prefix)
        {
            for (int i = 1; i <= 9999; i++)
                if (zip.GetEntry($"{prefix}{i}.xml") == null) return i;
            throw new InvalidOperationException("파일 번호 할당 실패");
        }

        private static string GetNextRelId(ZipArchive zip, string relsPath)
        {
            string content = ReadEntry(zip, relsPath);
            if (content == null) return "rId1";

            int max = 0;
            var doc = XDocument.Parse(content);
            foreach (var rel in doc.Descendants(PkgRels + "Relationship"))
            {
                string id = (string)rel.Attribute("Id") ?? "";
                if (id.StartsWith("rId") && int.TryParse(id.Substring(3), out int n))
                    max = Math.Max(max, n);
            }
            return $"rId{max + 1}";
        }

        // ─── Rels 조작 ──────────────────────────────────────────────────────

        private static void AddRel(ZipArchive zip, string relsPath, string id, string type, string target)
        {
            string content = ReadEntry(zip, relsPath) ?? BuildEmptyRelsXml();
            var doc = XDocument.Parse(content);
            doc.Root.Add(new XElement(PkgRels + "Relationship",
                new XAttribute("Id", id),
                new XAttribute("Type", type),
                new XAttribute("Target", target)));
            WriteEntry(zip, relsPath, ToXmlString(doc));
        }

        // ─── 시트 XML 조작 ──────────────────────────────────────────────────

        private static void AddDrawingElementToSheet(ZipArchive zip, string sheetPath, string drawingRelId)
        {
            // XDocument 라운드트립을 피하기 위해 문자열 기반으로 삽입합니다.
            string content = ReadEntry(zip, sheetPath)
                ?? throw new FileNotFoundException($"시트 파일을 찾을 수 없습니다: {sheetPath}");

            // ClosedXML은 x: 접두사를 사용하지만, 기본 네임스페이스도 처리합니다.
            string drawingXml;
            int insertIdx;

            if (content.Contains("<x:tableParts"))
            {
                drawingXml = $"<x:drawing r:id=\"{drawingRelId}\"/>";
                insertIdx = content.IndexOf("<x:tableParts", StringComparison.Ordinal);
            }
            else if (content.Contains("<tableParts"))
            {
                drawingXml = $"<drawing r:id=\"{drawingRelId}\"/>";
                insertIdx = content.IndexOf("<tableParts", StringComparison.Ordinal);
            }
            else
            {
                // tableParts 없으면 worksheet 닫는 태그 앞에 삽입
                drawingXml = "<x:drawing r:id=\"" + drawingRelId + "\"/>";
                insertIdx = content.LastIndexOf("</x:worksheet>", StringComparison.Ordinal);
                if (insertIdx < 0)
                    insertIdx = content.LastIndexOf("</worksheet>", StringComparison.Ordinal);
                if (insertIdx < 0)
                    throw new InvalidOperationException("worksheet 닫는 태그를 찾을 수 없습니다.");
            }

            WriteEntry(zip, sheetPath, content.Insert(insertIdx, drawingXml));
        }

        // ─── Content_Types 조작 ─────────────────────────────────────────────

        private static void EnsureContentType(ZipArchive zip, string partName, string contentType)
        {
            // XDocument 라운드트립을 피하기 위해 문자열 기반으로 삽입합니다.
            const string ctPath = "[Content_Types].xml";
            string content = ReadEntry(zip, ctPath)
                ?? throw new FileNotFoundException("[Content_Types].xml을 찾을 수 없습니다.");

            // 이미 존재하면 스킵
            if (content.Contains("PartName=\"" + partName + "\""))
                return;

            string overrideXml = $"<Override PartName=\"{partName}\" ContentType=\"{contentType}\"/>";

            const string closeTag = "</Types>";
            int idx = content.LastIndexOf(closeTag, StringComparison.Ordinal);
            if (idx < 0)
                throw new InvalidOperationException("[Content_Types].xml에서 </Types> 태그를 찾을 수 없습니다.");

            WriteEntry(zip, ctPath, content.Insert(idx, overrideXml));
        }

        // ─── Drawing XML 조작 ───────────────────────────────────────────────

        private static void AppendAnchorToDrawing(
            ZipArchive zip, string drawingPath,
            int colFrom, int rowFrom, int colTo, int rowTo,
            string chartRelId, int chartNum)
        {
            string content = ReadEntry(zip, drawingPath)
                ?? throw new InvalidOperationException("drawing XML을 찾을 수 없습니다.");

            string anchor = BuildAnchorXml(colFrom, rowFrom, colTo, rowTo, chartRelId, chartNum);

            const string closeTag = "</xdr:wsDr>";
            int idx = content.LastIndexOf(closeTag, StringComparison.Ordinal);
            if (idx < 0) throw new InvalidOperationException("</xdr:wsDr> 태그를 찾을 수 없습니다.");

            WriteEntry(zip, drawingPath, content.Insert(idx, anchor));
        }

        // ─── XML 문자열 빌더 ────────────────────────────────────────────────

        private static string BuildChartXml(ChartBuildOptions options)
        {
            var chartSpace = ChartXmlGenerator.BuildChartSpace(options);
            return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n"
                   + chartSpace.OuterXml;
        }

        private static string BuildEmptyDrawingXml() =>
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
            "<xdr:wsDr" +
            " xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\"" +
            " xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"" +
            " xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">" +
            "</xdr:wsDr>";

        private static string BuildEmptyRelsXml() =>
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
            "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
            "</Relationships>";

        private static string BuildAnchorXml(
            int colFrom, int rowFrom, int colTo, int rowTo,
            string chartRelId, int chartNum) =>
            $"<xdr:twoCellAnchor editAs=\"oneCell\">" +
            $"<xdr:from><xdr:col>{colFrom}</xdr:col><xdr:colOff>0</xdr:colOff>" +
            $"<xdr:row>{rowFrom}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>" +
            $"<xdr:to><xdr:col>{colTo}</xdr:col><xdr:colOff>0</xdr:colOff>" +
            $"<xdr:row>{rowTo}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>" +
            $"<xdr:graphicFrame macro=\"\">" +
            $"<xdr:nvGraphicFramePr>" +
            $"<xdr:cNvPr id=\"{chartNum + 1}\" name=\"Chart {chartNum}\"/>" +
            $"<xdr:cNvGraphicFramePr><a:graphicFrameLocks noGrp=\"1\"/></xdr:cNvGraphicFramePr>" +
            $"</xdr:nvGraphicFramePr>" +
            $"<xdr:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"0\" cy=\"0\"/></xdr:xfrm>" +
            $"<a:graphic>" +
            $"<a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/chart\">" +
            $"<c:chart xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\" r:id=\"{chartRelId}\"/>" +
            $"</a:graphicData></a:graphic>" +
            $"</xdr:graphicFrame>" +
            $"<xdr:clientData/>" +
            $"</xdr:twoCellAnchor>";

        // ─── ZIP / XML 유틸리티 ─────────────────────────────────────────────

        private static string ReadEntry(ZipArchive zip, string path)
        {
            var entry = zip.GetEntry(path);
            if (entry == null) return null;
            using (var sr = new StreamReader(entry.Open(), Encoding.UTF8))
                return sr.ReadToEnd();
        }

        private static XDocument ParseEntry(ZipArchive zip, string path)
        {
            string content = ReadEntry(zip, path)
                ?? throw new FileNotFoundException($"패키지에서 찾을 수 없습니다: {path}");
            return XDocument.Parse(content);
        }

        private static void WriteEntry(ZipArchive zip, string path, string xml)
        {
            zip.GetEntry(path)?.Delete();
            var entry = zip.CreateEntry(path, CompressionLevel.Optimal);
            using (var sw = new StreamWriter(entry.Open(), new UTF8Encoding(false)))
                sw.Write(xml);
        }

        private static string ToXmlString(XDocument doc)
        {
            using (var ms = new MemoryStream())
            {
                var settings = new XmlWriterSettings
                {
                    Encoding = new UTF8Encoding(false),
                    OmitXmlDeclaration = false
                };
                using (var xw = XmlWriter.Create(ms, settings))
                    doc.Save(xw);
                return Encoding.UTF8.GetString(ms.ToArray());
            }
        }

        // ─── 경로 유틸리티 ──────────────────────────────────────────────────

        private static string GetRelsPath(string partPath)
        {
            int slash = partPath.LastIndexOf('/');
            string dir = slash >= 0 ? partPath.Substring(0, slash) : "";
            string file = partPath.Substring(slash + 1);
            return string.IsNullOrEmpty(dir)
                ? $"_rels/{file}.rels"
                : $"{dir}/_rels/{file}.rels";
        }

        private static string GetDir(string filePath)
        {
            int slash = filePath.LastIndexOf('/');
            return slash >= 0 ? filePath.Substring(0, slash) : "";
        }

        /// <summary>절대(/로 시작) 또는 상대 target을 baseDir 기준으로 정규화합니다.</summary>
        private static string NormalizePath(string baseDir, string target)
        {
            if (target.StartsWith("/"))
                return target.TrimStart('/');

            var parts = new List<string>(
                baseDir.TrimEnd('/').Split(new[] { '/' }, StringSplitOptions.RemoveEmptyEntries));
            foreach (var seg in target.Split('/'))
            {
                if (seg == "..") { if (parts.Count > 0) parts.RemoveAt(parts.Count - 1); }
                else if (seg != ".") parts.Add(seg);
            }
            return string.Join("/", parts);
        }

        /// <summary>fromFile 기준으로 toFile 까지의 상대 경로를 계산합니다.</summary>
        private static string MakeRelativePath(string fromFile, string toFile)
        {
            string fromDir = GetDir(fromFile);
            string[] fromParts = fromDir.Length > 0
                ? fromDir.Split('/')
                : new string[0];
            string[] toParts = toFile.Split('/');

            int common = 0;
            while (common < fromParts.Length && common < toParts.Length
                   && fromParts[common] == toParts[common])
                common++;

            var sb = new StringBuilder();
            for (int i = common; i < fromParts.Length; i++) sb.Append("../");
            for (int i = common; i < toParts.Length; i++)
            {
                if (i > common) sb.Append('/');
                sb.Append(toParts[i]);
            }
            return sb.ToString();
        }
    }
}
