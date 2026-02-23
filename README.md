# XLSX-Master

.NET Framework 4.8 기반 Excel 리포트 생성 래퍼 라이브러리.
**ClosedXML**(데이터/셀/시트)과 **Open XML SDK**(차트 XML 직접 주입)를 결합하여
Fluent API로 복잡한 차트(콤보/보조축/다중계열)를 손쉽게 생성합니다.

---

## 특징

- **Fluent API** — 메서드 체이닝으로 직관적인 차트 정의
- **콤보 차트** — 동일 차트에 Column + Line + Area 혼합
- **보조 Y축** — 단위가 다른 계열을 오른쪽 축에 배치
- **다중 차트** — 한 워크북에 차트 N개 동시 주입
- **다양한 차트 타입** — Column, Bar, Line, Area, AreaStacked, Pie, Scatter
- **시각 커스터마이징** — 계열 색상, 마커, 축 범위/제목, 데이터 레이블, 차트 스타일
- **OpenXML 스키마 준수** — `OpenXmlValidator` 기반 검증 테스트 포함

---

## 설치

상황에 따라 세 가지 방법 중 하나를 선택하세요.

---

### 방법 1 — 프로젝트 참조 (같은 솔루션)

소스를 함께 수정하면서 사용할 때 가장 적합합니다.

```xml
<!-- 사용할 프로젝트의 .csproj -->
<ItemGroup>
  <ProjectReference Include="..\XLSX-Master\src\XLSX-Master\XLSX-Master.csproj" />
</ItemGroup>
```

---

### 방법 2 — DLL 직접 참조

빌드된 DLL 파일만 배포받아 참조할 때 사용합니다.

**1. DLL 빌드**

```bash
dotnet build src/XLSX-Master/XLSX-Master.csproj -c Release
# 출력: src/XLSX-Master/bin/Release/net48/XLSX-Master.dll
```

**2. 사용할 프로젝트의 `.csproj`에 추가**

`XLSX-Master.dll`을 프로젝트의 `lib\` 폴더에 복사한 뒤:

```xml
<ItemGroup>
  <Reference Include="XLSX-Master">
    <HintPath>lib\XLSX-Master.dll</HintPath>
  </Reference>
</ItemGroup>

<!-- 의존 패키지는 NuGet으로 별도 설치 -->
<ItemGroup>
  <PackageReference Include="ClosedXML" Version="0.102.3" />
  <PackageReference Include="DocumentFormat.OpenXml" Version="2.20.0" />
</ItemGroup>
```

---

### 방법 3 — 로컬 NuGet 패키지

여러 프로젝트에서 패키지처럼 관리할 때 사용합니다.

**1. `.nupkg` 생성**

```bash
dotnet pack src/XLSX-Master/XLSX-Master.csproj -c Release -o ./nupkg
```

**2. 로컬 NuGet 피드 등록**

```bash
dotnet nuget add source "D:\1_Work\XLSX-Master\nupkg" --name LocalXlsxMaster
```

**3. 사용할 프로젝트에서 설치**

```bash
dotnet add package XLSX-Master
```

---

### 의존 패키지 (공통)

| 패키지 | 버전 |
|--------|------|
| ClosedXML | 0.102.3 |
| DocumentFormat.OpenXml | 2.20.0 |

> ClosedXML 0.102.3은 `DocumentFormat.OpenXml >= 2.16.0 && < 3.0.0`을 요구합니다.
> **3.x로 올리면 버전 충돌 경고가 발생합니다.**

---

## 빠른 시작

```csharp
using ClosedXML.Excel;
using XlsxMaster.Charts;
using XlsxMaster.Extensions;

using var workbook = new XLWorkbook();
var ws = workbook.Worksheets.Add("영업보고서");

// 데이터 입력
ws.Cell("A1").Value = "월";   ws.Cell("B1").Value = "매출";  ws.Cell("C1").Value = "성장률";
ws.Cell("A2").Value = "1월"; ws.Cell("B2").Value = 100;      ws.Cell("C2").Value = 0.0;
ws.Cell("A3").Value = "2월"; ws.Cell("B3").Value = 120;      ws.Cell("C3").Value = 20.0;
// ... 추가 데이터

// 차트 정의 (Fluent API)
ws.AddMasterChart(anchor: "E1:M20")
  .SetXAxis("A2:A7")
  .AddSeries("매출액", "B2:B7", ChartType.Column)
  .AddSeries("성장률", "C2:C7", ChartType.Line, useSecondaryAxis: true)
  .SetTitle("월별 영업 실적")
  .ShowLegend(true);

// 차트 포함 저장
workbook.SaveWithCharts("BusinessReport.xlsx");
```

---

## 주요 API 레퍼런스

### `IXLWorksheet.AddMasterChart(anchor)`

차트 빌더를 생성합니다.

| 파라미터 | 설명 |
|----------|------|
| `anchor` | 차트 위치 범위 (예: `"E1:M20"`) |

반환값: `MasterChartBuilder` (Fluent 체이닝)

---

### `MasterChartBuilder` 메서드

| 메서드 | 설명 |
|--------|------|
| `.SetXAxis(range)` | X축 카테고리 범위 (예: `"A2:A13"`) |
| `.AddSeries(name, range, type, useSecondaryAxis)` | 데이터 계열 추가 |
| `.AddScatterSeries(name, xRange, yRange, hexColor)` | 분산형 계열 추가 |
| `.SetTitle(title)` | 차트 제목 |
| `.ShowLegend(bool)` | 범례 표시 여부 (기본 `true`) |
| `.ShowDataLabels(bool)` | 데이터 레이블 표시 여부 (기본 `false`) |
| `.SetSeriesColor(name, hexColor)` | 계열 색상 지정 (예: `"4472C4"`) |
| `.SetMarkerStyle(name, style)` | 꺾은선 마커 모양 |
| `.SetYAxisMin(value)` / `.SetYAxisMax(value)` | 주 Y축 범위 |
| `.SetSecondaryYAxisMin(value)` / `.SetSecondaryYAxisMax(value)` | 보조 Y축 범위 |
| `.SetYAxisTitle(title)` | 주 Y축 제목 |
| `.SetXAxisTitle(title)` | X축 제목 |
| `.SetChartStyle(1~48)` | Excel 내장 차트 스타일 |

---

### `IXLWorkbook` 저장 메서드

| 메서드 | 설명 |
|--------|------|
| `.SaveWithCharts(filePath)` | 차트 포함 파일 저장 |
| `.SaveWithCharts(stream)` | 차트 포함 스트림 출력 |
| `.SaveWithChartsToStream()` | 차트 포함 `MemoryStream` 반환 |

> **주의:** ClosedXML의 기본 `SaveAs()`는 차트를 포함하지 않습니다.
> 반드시 위의 `SaveWithCharts*` 메서드를 사용하세요.

---

## 지원 차트 타입

| `ChartType` 열거값 | Excel 차트 종류 |
|--------------------|-----------------|
| `Column` (기본값) | 세로 막대형 |
| `Bar` | 가로 막대형 |
| `Line` | 꺾은선형 |
| `Area` | 영역형 |
| `AreaStacked` | 누적 영역형 |
| `Pie` | 원형 |
| `Scatter` (AddScatterSeries 사용) | 분산형 |

---

## 콤보 차트 예시

```csharp
ws.AddMasterChart("E1:M20")
  .SetXAxis("A2:A13")
  .AddSeries("매출액", "B2:B13", ChartType.Column)           // 주 Y축
  .AddSeries("성장률", "C2:C13", ChartType.Line,
             useSecondaryAxis: true)                           // 보조 Y축
  .SetYAxisTitle("매출 (억원)")
  .SetSecondaryYAxisMin(-50)
  .SetSecondaryYAxisMax(100)
  .SetTitle("매출 및 성장률");
```

---

## 알려진 제약 사항

- **`SaveWithCharts()`만 차트를 포함합니다** — ClosedXML의 `SaveAs()`는 차트 미포함
- **EMU 계산은 Excel 기본 열/행 크기 기준** — 열 너비 변경 시 차트 위치 불일치 가능 (추후 개선 예정)
- **Scatter 계열과 일반 계열 혼용 불가** — 같은 차트에 `AddSeries()`와 `AddScatterSeries()` 동시 사용 불가
- **`dotnet new classlib -f net48` 미지원** (dotnet 9 CLI) — `.csproj` 수동 작성 필요

---

## 빌드 & 테스트

```bash
# 전체 솔루션 빌드
dotnet build XLSX-Master.sln

# 테스트 실행 (57개 테스트)
dotnet test tests/XLSX-Master.Tests/XLSX-Master.Tests.csproj

# 샘플 실행 (xlsx 파일 생성)
dotnet run --project samples/XLSX-Master.Samples/XLSX-Master.Samples.csproj
```

---

## 라이선스

MIT
