# 📝 PRD: High-Performance Excel Export Wrapper (XLSX-Master)

## 1. 제품 개요 (Product Overview)

- **목적:** .NET Framework 4.8 환경에서 `ClosedXML`의 직관적인 사용성과 `Open XML SDK`의 강력한 시각화 기능을 결합하여, **표와 차트가 포함된 고성능 리포트**를 생성하는 래퍼 라이브러리 구축.

- **핵심 가치:** * **Developer Experience:** 복잡한 XML 조작 없이 Fluent API로 차트 생성.
  
  - **Visual Richness:** 단일/다중 계열 차트 및 보조축을 포함한 혼합형 차트 지원.
  
  - **Legacy Ready:** .NET 4.8 환경의 제약 사항을 완벽히 준수.

---

## 2. 핵심 기능 요구사항 (Functional Requirements)

### 2.1 데이터 핸들링 및 표 (Data & Table)

- **Smart Binding:** `IEnumerable<T>`, `DataTable` 등을 시트에 삽입 시 헤더 자동 생성 및 데이터 타입(숫자, 날짜, 문자열) 자동 매핑.

- **Excel Table Object:** 삽입된 데이터 범위를 엑셀 공식 `Table` 객체로 변환하여 자동 필터 및 테이블 스타일 적용.

### 2.2 고도화된 차트 기능 (Advanced Visualization)

- **Multi-Series Support:** 하나의 X축(Category)에 대해 여러 개의 Y축(Value) 데이터 계열을 표시.

- **Combo Charts:** 동일 차트 내에서 서로 다른 차트 타입 혼용 (예: 매출은 막대, 성장률은 꺾은선).

- **Dual Axis (Secondary Axis):** 데이터 단위가 다른 계열을 위해 보조 Y축(오른쪽 축) 제공.

- **Dynamic Range Linking:** 셀 데이터가 수정되면 차트가 자동으로 갱신되도록 XML Formula(`Sheet1!$A$1:$B$10`) 기반 연결.

- **Layout Control:** 셀 주소 기반으로 차트의 위치(Anchor)와 크기 지정.

---

## 3. 기술 아키텍처 및 스택 (Technical Stack)

- **Runtime:** .NET Framework 4.8

- **Base Engine:** * `ClosedXML`: 데이터 입출력, 셀 서식, 기본 시트 관리.
  
  - `Open XML SDK 2.5+`: `Drawing`, `ChartSpace` 등 차트 관련 XML 파트 직접 생성 및 주입.

- **Design Pattern:** **Wrapper & Extension Methods**. `IXLWorksheet` 객체에 `.AddMasterChart()` 같은 확장 메서드를 제공하여 기존 ClosedXML 코드와 이질감 없이 혼용.

---

## 4. 유저 시나리오 (Code Vibe)

C#

```
// XLSX-Master 바이브 코딩 예시
using (var workbook = new XLWorkbook()) {
    var ws = workbook.Worksheets.Add("Sales Analysis");

    // 1. 데이터 및 표 생성 (ClosedXML 기반)
    var table = ws.Cell(1, 1).InsertTable(salesData);

    // 2. 다중 계열 혼합 차트 추가 (XLSX-Master 확장 기능)
    ws.AddMasterChart(anchor: "E1:M20")
      .SetXAxis("A2:A13")                       // 월 데이터
      .AddSeries("매출액", "B2:B13", ChartType.Column) // 계열 1: 막대
      .AddSeries("성장률", "C2:C13", ChartType.Line, useSecondaryAxis: true) // 계열 2: 보조축 꺾은선
      .SetTitle("월별 실적 분석")
      .ShowLegend(true);

    workbook.SaveAs("BusinessReport.xlsx");
}
```

---

## 5. 상세 설계 가이드 (Implementation Details)

### 5.1 차트 XML 구조 관리

- `xl/charts/chartN.xml`: 차트의 본체. 다중 계열을 위해 `<c:ser>` 노드를 반복 생성.

- `xl/drawings/drawingN.xml`: 차트가 시트 위 어디에 위치할지 결정 (`TwoCellAnchor`).

- `xl/drawings/_rels/drawingN.xml.rels`: Drawing과 Chart 간의 관계 정의.

### 5.2 기술적 도전 과제

- **Relationship ID (RId) 관리:** ClosedXML이 생성한 파일 구조를 손상시키지 않고 새로운 RId를 할당하는 로직 필요.

- **Coordinate Calculation:** 엑셀의 EMU(English Metric Units) 단위를 사용하여 셀 위치를 정확한 픽셀 좌표로 계산하여 매핑.

---

## 6. 로드맵 (Roadmap)

- **1단계 (PoC):** Open XML SDK를 사용하여 단일 막대 차트가 포함된 파일 생성 검증.

- **2단계 (Wrapper):** ClosedXML 워크북 객체에서 메모리 스트림으로 넘겨받아 차트를 주입하는 핵심 엔진 개발.

- **3단계 (Advanced):** 멀티 계열, 보조축, 콤보 차트 기능을 Fluent API로 추상화.

- **4단계 (Final):** 예외 처리 및 대용량 데이터 테스트.
