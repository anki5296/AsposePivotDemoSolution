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


            int iPivotIndex = worksheet.PivotTables.Add("Data!A1:C39","A1","PivotTable");
            PivotTable pt = worksheet.PivotTables[iPivotIndex];
            pt.RowGrand = true;
            pt.ColumnGrand = true;
            pt.IsAutoFormat = true;
            pt.AddFieldToArea(PivotFieldType.Row, 0);
            pt.AddFieldToArea(PivotFieldType.Row, 1);
            pt.AddFieldToArea(PivotFieldType.Data, 2);



           workbook.Save("C:/Users/Ankita/Documents/bc69.xlsx");




        }
    }
}
