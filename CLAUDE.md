# CLAUDE.md — XLSX-Master

## 프로젝트 개요
.NET Framework 4.8 기반 Excel 리포트 생성 래퍼 라이브러리.
ClosedXML(데이터/셀/시트)과 Open XML SDK(chart XML 직접 주입)를 결합하여
Fluent API로 복잡한 차트(콤보/보조축/다중계열)를 생성한다.

---

## PRD 참조

**원본 문서:** `PRD_XLSX-Master.md` (프로젝트 루트)

새 기능을 구현하거나 설계 결정을 내릴 때는 **반드시 PRD를 먼저 확인**할 것.

| PRD 섹션 | 참조 시점 |
|----------|-----------|
| §2.1 데이터 핸들링 및 표 | `IEnumerable<T>` / `DataTable` 바인딩, Excel Table 구현 시 |
| §2.2 고도화된 차트 기능 | 보조축·콤보·동적범위 관련 기능 추가 시 |
| §3 기술 아키텍처 | 의존 패키지 변경 또는 설계 패턴 논의 시 |
| §4 유저 시나리오 (코드 바이브) | 공개 API 시그니처가 예시 코드와 일치하는지 검증 시 |
| §5 상세 설계 가이드 | chart XML 구조(`<c:ser>`, Drawing, RId) 작업 시 |
| §6 로드맵 | TODO 우선순위 결정 및 단계 범위 확인 시 |

### PRD와 코드 간 핵심 대응 관계

```
PRD §4 코드 바이브 예시
  ws.AddMasterChart(anchor: "E1:M20")
    → WorksheetExtensions.AddMasterChart()  +  MasterChartBuilder

  .SetXAxis("A2:A13")
    → MasterChartBuilder._categoryRange

  .AddSeries("매출액", "B2:B13", ChartType.Column)
  .AddSeries("성장률", "C2:C13", ChartType.Line, useSecondaryAxis: true)
    → SeriesDefinition  →  ChartXmlGenerator (BarChart / LineChart)

  workbook.SaveAs("BusinessReport.xlsx")   ← PRD 원문
    → 실제 구현은 workbook.SaveWithCharts("BusinessReport.xlsx")
       (PRD 대비 변경점: ClosedXML SaveAs만으로는 차트 미포함)
```

> **PRD 대비 변경점 기록**
> - PRD §4: `workbook.SaveAs(...)` 한 줄로 저장 — 실제로는 `SaveWithCharts()` 별도 호출 필요
>   (ClosedXML 내부 후킹이 불가하여 별도 확장 메서드로 분리)

---

## 빌드 & 테스트 명령어

```bash
# 전체 솔루션 빌드
dotnet build XLSX-Master.sln

# 라이브러리만 빌드
dotnet build src/XLSX-Master/XLSX-Master.csproj

# 테스트 실행
dotnet test tests/XLSX-Master.Tests/XLSX-Master.Tests.csproj

# 샘플 앱 실행 (xlsx 파일 생성)
dotnet run --project samples/XLSX-Master.Samples/XLSX-Master.Samples.csproj
```

---

## 아키텍처 핵심 흐름

```
사용자 코드
  └─ ws.AddMasterChart("E1:M20")          WorksheetExtensions
       .AddSeries(...)                     MasterChartBuilder (Fluent)
       ──등록──> ChartRegistry             ConditionalWeakTable<IXLWorkbook, List<Builder>>

workbook.SaveWithCharts("out.xlsx")        WorkbookExtensions
  1. workbook.SaveAs(MemoryStream)         ClosedXML가 데이터/서식 처리
  2. foreach builder → InjectInto(stream)  XlsxChartInjector
       ├─ DrawingsPart 추가/획득
       ├─ ChartPart 추가
       ├─ ChartXmlGenerator.Build()        Open XML SDK로 chartN.xml 생성
       └─ TwoCellAnchor 위치 지정
  3. File.WriteAllBytes(path, result)
```

---

## 파일 구조 & 역할

| 파일 | 역할 |
|------|------|
| `Charts/ChartType.cs` | 차트 종류 enum (Column/Bar/Line/Area/Pie/Scatter) |
| `Charts/AxisPosition.cs` | Primary / Secondary Y축 enum |
| `Charts/SeriesDefinition.cs` | 단일 계열 데이터 모델 |
| `Charts/MasterChartBuilder.cs` | Fluent 빌더 — 공개 API 진입점 |
| `Charts/ChartXmlGenerator.cs` | Open XML SDK chart XML 생성 (내부) |
| `Core/XlsxChartInjector.cs` | xlsx 스트림에 chart/drawing 주입 (내부) |
| `Core/ChartRegistry.cs` | 워크북별 빌더 목록 관리 (내부) |
| `Extensions/WorksheetExtensions.cs` | `IXLWorksheet.AddMasterChart()` |
| `Extensions/WorkbookExtensions.cs` | `IXLWorkbook.SaveWithCharts()` |
| `Helpers/EmuCalculator.cs` | 셀 주소 → EMU 좌표 변환 |
| `Helpers/RidManager.cs` | Relationship ID 충돌 없이 다음 번호 할당 |

---

## 패키지 버전 (변경 시 주의)

- `ClosedXML 0.102.3` — `DocumentFormat.OpenXml >= 2.16.0 && < 3.0.0` 요구
- `DocumentFormat.OpenXml 2.20.0` — **3.x로 올리면 버전 충돌 경고 발생**
- `MSTest.TestFramework / TestAdapter 3.7.3`

---

## 코딩 컨벤션

- **C# 7.3** (LangVersion): `ValueTuple`, `ref returns` 사용 가능, `nullable` 비활성화
- 공개 API(`public`)는 XML 문서 주석 필수
- 내부 구현(`internal`)은 주석 선택
- 네임스페이스 충돌 주의: `XlsxMaster.Charts.AxisPosition` ↔ `DocumentFormat.OpenXml.Drawing.Charts.AxisPosition`
  → `ChartXmlGenerator.cs`에서 반드시 `C.AxisPosition { Val = ... }` 형태로 사용
- 외부 스트림은 호출자가 Dispose — 라이브러리 내부에서 임의 Dispose 금지

---

## 새 차트 타입 추가 방법

1. `ChartType.cs` enum에 값 추가
2. `ChartXmlGenerator.cs`의 `BuildChartElement()` switch에 case 추가
3. `Build{Type}Chart()` / `Build{Type}ChartSeries()` 메서드 구현
4. `IntegrationTests.cs`에 검증 테스트 추가

---

## 알려진 제약 사항

- `dotnet new classlib -f net48` 미지원 (dotnet 9 CLI) → csproj 수동 작성
- ClosedXML 워크북 저장 후 재편집 불가 (스트림 기반 파이프라인)
- 차트 주입은 `SaveWithCharts()` 또는 `SaveWithChartsToStream()` 호출 시 일괄 처리
- 현재 EMU 계산은 Excel 기본 열/행 크기 기준 — 사용자가 열 너비를 변경해도 차트 위치에 반영 안 됨 (추후 개선 대상)
