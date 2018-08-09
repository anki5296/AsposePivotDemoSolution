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
            Workbook workbook = new Workbook("C:/Users/Ankita/Downloads/global.xlsx");
            workbook.Worksheets.RemoveAt("Pivot");
            Worksheet worksheet = workbook.Worksheets.Add("Pivot");


            Worksheet sheet0 = workbook.Worksheets[0];
            //Cells cells = sheet0.Cells;

            Cell cell = sheet0.Cells.LastCell;
            Cell cell_first = sheet0.Cells.FirstCell;

            #region Lastcell calculation

            
            int col_last = cell.Column + 1;
            int row_last = cell.Row + 1;

            #endregion

            #region First calculation


            int col_first = cell_first.Column + 1;
            int row_first = cell_first.Row + 1;

            #endregion



            #region Calculate column character
            int dividend = col_last;
            string columnName = String.Empty;

            while (dividend > 0)
            {
                var modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                dividend = (dividend - modulo) / 26;
            }

            #endregion

            string sheetname = sheet0.Name;
            string datasource = sheetname + "!A3:" + columnName + row_last.ToString();
            //string datasource = sheetname + "!A" + row_first.ToString() + ":" + columnName + row_last.ToString();



            int iPivotIndex = worksheet.PivotTables.Add(datasource, "A1", "PivotTable");
            PivotTable pt = worksheet.PivotTables[iPivotIndex];
            pt.RowGrand = true;
            pt.ColumnGrand = true;
            pt.IsAutoFormat = true;
            pt.AddFieldToArea(PivotFieldType.Column, 0);
            pt.AddFieldToArea(PivotFieldType.Row, 1);
            pt.AddFieldToArea(PivotFieldType.Data, 2);
            pt.AddFieldToArea(PivotFieldType.Data, 3);
            pt.PivotTableStyleType = PivotTableStyleType.PivotTableStyleDark1;
            pt.DataFields[0].NumberFormat = @"[>999999999]$#\,###\,###\,##0.00;[>=1000000]$###\,###\,##0.00;$#,###";
            pt.DataFields[1].NumberFormat = @"[>999999999]$#\,###\,###\,##0.00;[>=1000000]$###\,###\,##0.00;$#,###"; ;
            //workbook.Worksheets[0].IsVisible = false;

            Style st = workbook.CreateStyle();
            pt.FormatAll(st);

            workbook.Save("C:/Users/Ankita/Downloads/global.xlsx");




        }
    }
}
