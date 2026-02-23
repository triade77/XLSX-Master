using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Packaging;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using XlsxMaster.Extensions;

namespace XlsxMaster.Tests
{
    [TestClass]
    public class DataBindingTests
    {
        // ──────────────────────────────────────────────────────────────
        // IEnumerable<T> 바인딩
        // ──────────────────────────────────────────────────────────────

        [TestMethod]
        public void InsertMasterTable_Generic_WritesHeadersFromPropertyNames()
        {
            using (var workbook = new XLWorkbook())
            {
                var ws = workbook.Worksheets.Add("Sheet1");
                ws.InsertMasterTable(GetSampleRecords());

                Assert.AreEqual("Name",     ws.Cell(1, 1).GetString());
                Assert.AreEqual("Amount",   ws.Cell(1, 2).GetString());
                Assert.AreEqual("Rate",     ws.Cell(1, 3).GetString());
                Assert.AreEqual("Date",     ws.Cell(1, 4).GetString());
                Assert.AreEqual("IsActive", ws.Cell(1, 5).GetString());
            }
        }

        [TestMethod]
        public void InsertMasterTable_Generic_WritesDataRows()
        {
            using (var workbook = new XLWorkbook())
            {
                var ws = workbook.Worksheets.Add("Sheet1");
                var records = GetSampleRecords();
                ws.InsertMasterTable(records);

                // 첫 번째 데이터 행 확인 (헤더 = 1행, 데이터 = 2행~)
                Assert.AreEqual("Alpha",           ws.Cell(2, 1).GetString());
                Assert.AreEqual(1000.0,            ws.Cell(2, 2).GetDouble());
                Assert.AreEqual(0.05,              ws.Cell(2, 3).GetDouble(), 1e-9);
                Assert.AreEqual(new DateTime(2024, 1, 15), ws.Cell(2, 4).GetDateTime());
                Assert.AreEqual(true,              ws.Cell(2, 5).GetBoolean());
            }
        }

        [TestMethod]
        public void InsertMasterTable_Generic_AppliesNumberFormat_ForIntColumn()
        {
            using (var workbook = new XLWorkbook())
            {
                var ws = workbook.Worksheets.Add("Sheet1");
                ws.InsertMasterTable(new List<IntModel> { new IntModel { Count = 1234 } });

                var fmt = ws.Cell(2, 1).Style.NumberFormat.Format;
                Assert.AreEqual("#,##0", fmt, "정수 컬럼은 #,##0 포맷이어야 합니다.");
            }
        }

        [TestMethod]
        public void InsertMasterTable_Generic_AppliesNumberFormat_ForDecimalColumn()
        {
            using (var workbook = new XLWorkbook())
            {
                var ws = workbook.Worksheets.Add("Sheet1");
                ws.InsertMasterTable(new List<DecimalModel> { new DecimalModel { Price = 9.99m } });

                var fmt = ws.Cell(2, 1).Style.NumberFormat.Format;
                Assert.AreEqual("#,##0.00", fmt, "decimal 컬럼은 #,##0.00 포맷이어야 합니다.");
            }
        }

        [TestMethod]
        public void InsertMasterTable_Generic_AppliesDateFormat()
        {
            using (var workbook = new XLWorkbook())
            {
                var ws = workbook.Worksheets.Add("Sheet1");
                ws.InsertMasterTable(new List<DateModel>
                {
                    new DateModel { CreatedAt = new DateTime(2025, 6, 1) }
                });

                var fmt = ws.Cell(2, 1).Style.NumberFormat.Format;
                Assert.AreEqual("yyyy-mm-dd", fmt, "DateTime 컬럼은 yyyy-mm-dd 포맷이어야 합니다.");
            }
        }

        [TestMethod]
        public void InsertMasterTable_Generic_NullValue_LeavesBlank()
        {
            using (var workbook = new XLWorkbook())
            {
                var ws = workbook.Worksheets.Add("Sheet1");
                ws.InsertMasterTable(new List<NullableModel>
                {
                    new NullableModel { Value = null }
                });

                Assert.IsTrue(ws.Cell(2, 1).IsEmpty(), "null 값은 빈 셀이어야 합니다.");
            }
        }

        [TestMethod]
        public void InsertMasterTable_Generic_ReturnsCorrectRange()
        {
            using (var workbook = new XLWorkbook())
            {
                var ws = workbook.Worksheets.Add("Sheet1");
                var records = GetSampleRecords(); // 3건
                var range = ws.InsertMasterTable(records);

                // 헤더(1행) + 데이터(3행) = 4행, 5컬럼
                Assert.AreEqual(1, range.FirstRow().RowNumber());
                Assert.AreEqual(4, range.LastRow().RowNumber());
                Assert.AreEqual(5, range.ColumnCount());
            }
        }

        [TestMethod]
        public void InsertMasterTable_Generic_EmptyData_ReturnsHeaderOnlyRange()
        {
            using (var workbook = new XLWorkbook())
            {
                var ws = workbook.Worksheets.Add("Sheet1");
                var range = ws.InsertMasterTable(new List<SampleRecord>());

                Assert.AreEqual(1, range.RowCount(), "빈 데이터는 헤더 행 1개만 있어야 합니다.");
            }
        }

        [TestMethod]
        public void InsertMasterTable_Generic_NoReadableProps_ThrowsArgumentException()
        {
            using (var workbook = new XLWorkbook())
            {
                var ws = workbook.Worksheets.Add("Sheet1");
                Assert.ThrowsException<ArgumentException>(() =>
                    ws.InsertMasterTable(new List<NoProps>()));
            }
        }

        // ──────────────────────────────────────────────────────────────
        // DataTable 바인딩
        // ──────────────────────────────────────────────────────────────

        [TestMethod]
        public void InsertMasterTable_DataTable_WritesHeadersFromColumnNames()
        {
            using (var workbook = new XLWorkbook())
            {
                var ws = workbook.Worksheets.Add("Sheet1");
                ws.InsertMasterTable(BuildDataTable());

                Assert.AreEqual("Product",  ws.Cell(1, 1).GetString());
                Assert.AreEqual("Qty",      ws.Cell(1, 2).GetString());
                Assert.AreEqual("Price",    ws.Cell(1, 3).GetString());
            }
        }

        [TestMethod]
        public void InsertMasterTable_DataTable_WritesDataRows()
        {
            using (var workbook = new XLWorkbook())
            {
                var ws = workbook.Worksheets.Add("Sheet1");
                ws.InsertMasterTable(BuildDataTable());

                Assert.AreEqual("Widget", ws.Cell(2, 1).GetString());
                Assert.AreEqual(10.0,     ws.Cell(2, 2).GetDouble());
                Assert.AreEqual(5.5,      ws.Cell(2, 3).GetDouble(), 1e-9);
            }
        }

        [TestMethod]
        public void InsertMasterTable_DataTable_DbNull_LeavesBlank()
        {
            using (var workbook = new XLWorkbook())
            {
                var ws = workbook.Worksheets.Add("Sheet1");
                var dt = new DataTable();
                dt.Columns.Add("Val", typeof(string));
                dt.Rows.Add(DBNull.Value);

                ws.InsertMasterTable(dt);

                Assert.IsTrue(ws.Cell(2, 1).IsEmpty(), "DBNull은 빈 셀이어야 합니다.");
            }
        }

        // ──────────────────────────────────────────────────────────────
        // AddExcelTable
        // ──────────────────────────────────────────────────────────────

        [TestMethod]
        public void AddExcelTable_CreatesTableWithAutoFilter()
        {
            byte[] result;
            using (var workbook = new XLWorkbook())
            {
                var ws = workbook.Worksheets.Add("Sheet1");
                ws.InsertMasterTable(GetSampleRecords())
                  .AddExcelTable("SalesTable");

                using (var ms = new MemoryStream())
                {
                    workbook.SaveAs(ms);
                    result = ms.ToArray();
                }
            }

            using (var ms = new MemoryStream(result))
            using (var doc = SpreadsheetDocument.Open(ms, isEditable: false))
            {
                // xl/tables/tableN.xml 파트가 생성되었는지 확인
                var tableCount = doc.WorkbookPart.WorksheetParts
                    .SelectMany(wp => wp.TableDefinitionParts)
                    .Count();
                Assert.IsTrue(tableCount >= 1, "Excel Table 파트가 생성되어야 합니다.");
            }
        }

        [TestMethod]
        public void AddExcelTable_AutoGeneratesNameWhenNull()
        {
            using (var workbook = new XLWorkbook())
            {
                var ws = workbook.Worksheets.Add("Sheet1");
                var table = ws.InsertMasterTable(GetSampleRecords())
                               .AddExcelTable();  // name 생략

                Assert.IsFalse(string.IsNullOrEmpty(table.Name), "테이블 이름이 자동 생성되어야 합니다.");
            }
        }

        [TestMethod]
        public void AddExcelTable_AppliesSpecifiedTheme()
        {
            using (var workbook = new XLWorkbook())
            {
                var ws = workbook.Worksheets.Add("Sheet1");
                var table = ws.InsertMasterTable(GetSampleRecords())
                               .AddExcelTable("T1", XLTableTheme.TableStyleLight1);

                Assert.AreEqual(XLTableTheme.TableStyleLight1.Name, table.Theme.Name);
            }
        }

        // ──────────────────────────────────────────────────────────────
        // 테스트 데이터 모델 및 헬퍼
        // ──────────────────────────────────────────────────────────────

        private static List<SampleRecord> GetSampleRecords() => new List<SampleRecord>
        {
            new SampleRecord { Name = "Alpha", Amount = 1000,  Rate = 0.05, Date = new DateTime(2024, 1, 15), IsActive = true  },
            new SampleRecord { Name = "Beta",  Amount = 2000,  Rate = 0.10, Date = new DateTime(2024, 2, 20), IsActive = false },
            new SampleRecord { Name = "Gamma", Amount = 1500,  Rate = 0.07, Date = new DateTime(2024, 3, 10), IsActive = true  },
        };

        private static DataTable BuildDataTable()
        {
            var dt = new DataTable();
            dt.Columns.Add("Product", typeof(string));
            dt.Columns.Add("Qty",     typeof(int));
            dt.Columns.Add("Price",   typeof(double));
            dt.Rows.Add("Widget", 10, 5.5);
            dt.Rows.Add("Gadget", 5,  12.0);
            return dt;
        }

        public class SampleRecord
        {
            public string   Name     { get; set; }
            public int      Amount   { get; set; }
            public double   Rate     { get; set; }
            public DateTime Date     { get; set; }
            public bool     IsActive { get; set; }
        }

        public class IntModel     { public int     Count { get; set; } }
        public class DecimalModel { public decimal Price { get; set; } }
        public class DateModel    { public DateTime CreatedAt { get; set; } }
        public class NullableModel { public int?   Value { get; set; } }
        public class NoProps { } // 읽기 가능한 공개 프로퍼티 없음
    }
}
