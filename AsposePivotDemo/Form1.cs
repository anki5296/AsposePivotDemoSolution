using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Pivot;

namespace AsposePivotDemo
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Workbook workbook = new Workbook("C:/Users/Ankita/Documents/bc69.xlsx");
            workbook.Worksheets.RemoveAt("Pivot");
            Worksheet worksheet = workbook.Worksheets.Add("Pivot");


            Worksheet sheet0 = workbook.Worksheets[0];
            //Cells cells = sheet0.Cells;

            Cell cell = sheet0.Cells.LastCell;

            int col=cell.Column + 1;
            int row=cell.Row + 1;

         
            int dividend = col;
            string columnName = String.Empty;

            while (dividend > 0)
            {
                var modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                dividend = (dividend - modulo) / 26;
            }
            string sheetname = sheet0.Name;
            string datasource = sheetname+"!A1:"+ columnName + row.ToString();
            

        int iPivotIndex = worksheet.PivotTables.Add(datasource,"A1","PivotTable");
            PivotTable pt = worksheet.PivotTables[iPivotIndex];
            pt.RowGrand = true;
            pt.ColumnGrand = true;
            pt.IsAutoFormat = true;
            pt.AddFieldToArea(PivotFieldType.Row, 0);
            pt.AddFieldToArea(PivotFieldType.Row, 1);
            pt.AddFieldToArea(PivotFieldType.Data, 2);

            pt.PivotTableStyleType = PivotTableStyleType.PivotTableStyleDark1;

            //workbook.Worksheets[0].IsVisible = false;

            Style st = workbook.CreateStyle();
            pt.FormatAll(st);

            workbook.Save("C:/Users/Ankita/Documents/bc69.xlsx");




        }
    }
}
