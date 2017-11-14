using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace TBATools
{
    public partial class TBARibbon
    {
        private void TBARibbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        public Excel.Range GetDataRange(string Prompt, object Default)
        {
            Excel.Range actRange = Globals.ThisAddIn.Application.ActiveWindow.RangeSelection;
            try
            {
                actRange = Globals.ThisAddIn.Application.InputBox(
                    Prompt,
                    "TBATools",
                    Default,
                    System.Type.Missing,
                    System.Type.Missing,
                    System.Type.Missing,
                    System.Type.Missing,
                    8);
            }
            catch
            {
                actRange = null;
            }
            return actRange;
        }

        private void btnToColByCol_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Range actRange = Globals.ThisAddIn.Application.ActiveWindow.RangeSelection;

            actRange = GetDataRange("已选择范围：", actRange.Address.ToString());

            if (actRange == null)
            {
                return;
            }

            Excel.Range resRange;

            resRange = GetDataRange("选择结果地址", System.Type.Missing);

            if (resRange == null)
            {
                return;
            }
            try
            {
                int nRow = actRange.Rows.Count;
                int nCol = actRange.Columns.Count;
                int iCount = 0;
                for (int iCol = 1; iCol <= nCol; iCol++)
                {
                    for (int iRow = 1; iRow <= nRow; iRow++)
                    {
                        iCount++;
                        resRange.Cells[iCount, 1] = actRange.Cells[iRow, iCol];

                    }
                }
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }
            resRange.Select();
        }

        private void btnAbout_Click(object sender, RibbonControlEventArgs e)
        {
            Worksheet currentSheet = Globals.ThisAddIn.GetActiveWorkSheet();
            //currentSheet.Range["A1"].Value = "Hello World";
            //currentSheet.Columns.AutoFit();
            int iCount = 0;
            for (int iRow = 1; iRow < 10; iRow++)
            {
                for (int iCol = 1; iCol < 3; iCol++)
                {
                    iCount++;
                    currentSheet.Cells[iRow, iCol] = iCount;
                }
            }

            string version = System.Reflection.Assembly.GetExecutingAssembly()
            .GetName()
            .Version
            .ToString();
            System.Windows.Forms.MessageBox.Show(version,"版本号");
        }

        private void btnToColByRow_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Range actRange = Globals.ThisAddIn.Application.ActiveWindow.RangeSelection;

            actRange = GetDataRange("已选择范围：", actRange.Address.ToString());

            if (actRange == null)
            {
                return;
            }

            Excel.Range resRange;

            resRange = GetDataRange("选择结果地址", System.Type.Missing);

            if (resRange == null)
            {
                return;
            }
            int nRow = actRange.Rows.Count;
            int nCol = actRange.Columns.Count;
            int iCount = 0;
            for (int iRow = 1; iRow <= nRow; iRow++)
            {
                for (int iCol = 1; iCol <= nCol; iCol++)
                {
                    iCount++;
                    resRange.Cells[iCount, 1] = actRange.Cells[iRow, iCol];

                }
            }
            resRange.Select();
        }
        private void btnToRowByCol_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Range actRange = Globals.ThisAddIn.Application.ActiveWindow.RangeSelection;

            actRange = GetDataRange("已选择范围：", actRange.Address.ToString());

            if (actRange == null)
            {
                return;
            }

            Excel.Range resRange;

            resRange = GetDataRange("选择结果地址", System.Type.Missing);

            if (resRange == null)
            {
                return;
            }
            int nRow = actRange.Rows.Count;
            int nCol = actRange.Columns.Count;
            int iCount = 0;
            for (int iCol = 1; iCol <= nCol; iCol++)
            {
                for (int iRow = 1; iRow <= nRow; iRow++)
                {
                    iCount++;
                    resRange.Cells[1, iCount] = actRange.Cells[iRow, iCol];

                }
            }
            resRange.Select();
        }

        private void btnToRowByRow_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Range actRange = Globals.ThisAddIn.Application.ActiveWindow.RangeSelection;

            actRange = GetDataRange("已选择范围：", actRange.Address.ToString());

            if (actRange == null)
            {
                return;
            }

            Excel.Range resRange;

            resRange = GetDataRange("选择结果地址", System.Type.Missing);

            if (resRange == null)
            {
                return;
            }

            int nRow = actRange.Rows.Count;
            int nCol = actRange.Columns.Count;
            int iCount = 0;
            for (int iRow = 1; iRow <= nRow; iRow++)
            {
                for (int iCol = 1; iCol <= nCol; iCol++)
                {
                    iCount++;
                    resRange.Cells[1, iCount] = actRange.Cells[iRow, iCol];

                }
            }
            resRange.Select();
        }
    }
}
