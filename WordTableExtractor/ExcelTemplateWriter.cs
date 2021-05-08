using System;
using System.Collections.Generic;
using System.Data;
using System.Dynamic;
using System.Linq;
using System.Text;

namespace WordTableExtractor
{
    public class ExcelTemplateWriter
    {
        public ExcelTemplateWriter()
        {

        }

        public List<Dictionary<string, object>> WriteDataToTemplate(DataTable data, string templatePath)
        {
            var items = CreateObjectsFromTable(data);

            var orderedItems = items.OrderByDescending(x => x["Name"]).ToList();

            return orderedItems;
        }

        public List<Dictionary<string, object>> CreateObjectsFromTable(DataTable data)
        {
            // Get Column names
            var columnNames = new List<string>();
            for(int i = 0; i < data.Columns.Count; i++)
            {
                columnNames.Add(data.Columns[i].ColumnName);
            }

            var items = new List<Dictionary<string, object>>();

            for(int i = 0; i < data.Rows.Count; i++)
            {
                var dict = new Dictionary<string, object>();

                for(int j = 0; j < columnNames.Count; j++)
                {
                    dict.Add(columnNames[j], data.Rows[i].Field<object>(columnNames[j]));   
                }

                items.Add(dict);
            }

            return items;
        }
    }
}
