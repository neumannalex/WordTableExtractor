using FluentAssertions;
using System;
using System.Data;
using WordTableExtractor;
using Xunit;

namespace Test
{
    public class ExcelTemplateWriterTests
    {
        private DataTable DataTable1;

        public ExcelTemplateWriterTests()
        {
            CreateDataTables();
        }

        private void CreateDataTables()
        {
            DataTable1 = new DataTable("Table1");
            DataTable1.Columns.Add("Name");
            DataTable1.Columns.Add("Vorname");
            DataTable1.Columns.Add("Email");

            for(int i = 0; i < 5; i++)
            {
                var row = DataTable1.NewRow();
                row["Name"] = $"Name_{i+1}";
                row["Vorname"] = $"Vorname_{i + 1}";
                row["Email"] = $"Vorname_{i + 1}.Name_{i + 1}@gmail.com";

                DataTable1.Rows.Add(row);
            }
        }


        [Fact]
        public void Can_Create_Objects_From_DataTable()
        {
            var writer = new ExcelTemplateWriter();

            var items = writer.CreateObjectsFromTable(DataTable1);

            items.Should().HaveCount(DataTable1.Rows.Count);
        }

        [Fact]
        public void Can_Sort_Objects()
        {
            var writer = new ExcelTemplateWriter();
            var items = writer.WriteDataToTemplate(DataTable1, "");

            items.Should().HaveCount(DataTable1.Rows.Count);

            items[DataTable1.Rows.Count - 1]["Name"].Should().Be(DataTable1.Rows[0]["Name"]);

            items[0]["Name"].Should().Be(DataTable1.Rows[DataTable1.Rows.Count - 1]["Name"]);
        }
    }
}
