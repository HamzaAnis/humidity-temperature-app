﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ClosedXML.Attributes;
using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;

namespace Excel_Sample
{
    class Program
    {
         // Public
        public static void Create(String filePath)
        {
            String[] sensors =
            {
                "Sensor 1", "Sensor 2", "Sensor 3", "Sensor 4", "Sensor 5", "Sensor 6", "Sensor 7", "Sensor 8",
                "Sensor 9", "Sensor 10", "Sensor 11", "Sensor 12", "Sensor 13", "Sensor 14", "Sensor 15", "Sensor 16"
            };
            // From a DataTable
            var dataTable = GetTable();
            using (var wb = new XLWorkbook())
            {
                for (int i = 0; i < 16; i++)
            {
              
                    var ws = wb.Worksheets.Add(sensors[i]);

                    ws.Range(1, 1, 1, 4).Merge().AddToNamed("Titles");
                    ws.Cell(1, 1).InsertTable(dataTable[i]);

                    // Prepare the style for the titles
                  
                    ws.Columns().AdjustToContents();


                }
                wb.SaveAs(filePath);
            }

        }
        // Private
        private static DataTable[] GetTable()
        {
            DataTable[] table = new DataTable[16];
            for (int i = 0; i < 16; i++)
            {
                table[i] = new DataTable();
                table[i].Columns.Add("Dosage", typeof(int));
                table[i].Columns.Add("Drug", typeof(string));
                table[i].Columns.Add("Patient", typeof(string));
                table[i].Columns.Add("Date", typeof(DateTime));

                table[i].Rows.Add(25, "Indocin", "David", new DateTime(2000, 1, 1));
                table[i].Rows.Add(50, "Enebrel", "Sam", new DateTime(2000, 1, 2));
                table[i].Rows.Add(10, "Hydralazine", "Christoff", new DateTime(2000, 1, 3));
                table[i].Rows.Add(21, "Combivent", "Janet", new DateTime(2000, 1, 4));
                table[i].Rows.Add(100, "Dilantin", "Melanie", new DateTime(2000, 1, 5));
            }
            return table;
        }
        private static DataTable GetTable1()
        {
            DataTable table = new DataTable();
            table.Columns.Add("Dosage", typeof(int));
            table.Columns.Add("Drug", typeof(string));
            table.Columns.Add("Patient", typeof(string));
            table.Columns.Add("Date", typeof(DateTime));

            table.Rows.Add(25, "Indocin", "David", new DateTime(2000, 1, 1));
            table.Rows.Add(50, "Enebrel", "Sam", new DateTime(2000, 1, 2));
            table.Rows.Add(10, "Hydralazine", "Christoff", new DateTime(2000, 1, 3));
            table.Rows.Add(21, "Combivent", "Janet", new DateTime(2000, 1, 4));
            table.Rows.Add(100, "Dilantin", "Melanie", new DateTime(2000, 1, 5));
            return table;
        }
        public static void Create1(String filePath)
        {
            var workbook = new XLWorkbook();
            var ws = workbook.Worksheets.Add("Row Settings");

            var row1 = ws.Row(2);
            row1.Style.Fill.BackgroundColor = XLColor.Red;
            row1.Height = 30;

            var row2 = ws.Row(4);
            row2.Style.Fill.BackgroundColor = XLColor.DarkOrange;
            row2.Height = 3;

            workbook.SaveAs(filePath);
        }
        private class Person
        {
            [XLColumn(Header = "House Street")]
            public String House { get; set; }

            public String Name { get; set; }
            public Int32 Age { get; set; }

            [XLColumn(Header = "Class Type")]
            public static String ClassType { get { return nameof(Person); } }
        }

        static void Main(string[] args)
        { 
            Create1("Check2.xlsx");
        }
    }

}