using System;
using System.Runtime.InteropServices;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using Excel = Microsoft.Office.Interop.Excel;

namespace SmartExcel
{
    [ComVisible(true)]
    public class RibbonController : ExcelRibbon
    {
        private IRibbonUI ribbon;

        public void OnLoad(IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        public void OnSplitRowsButton(IRibbonControl control)
        {
            try
            {
                var excelApp = (Excel.Application)ExcelDnaUtil.Application;
                var splitForm = new SplitRowsForm(excelApp);
                splitForm.ShowDialog();
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show($"Error: {ex.Message}", "Split Rows Error", 
                    System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }
        }

        public void OnChangeDelimiterButton(IRibbonControl control)
        {
            try
            {
                var excelApp = (Excel.Application)ExcelDnaUtil.Application;
                var changeDelimiterForm = new ChangeDelimiterForm(excelApp);
                changeDelimiterForm.ShowDialog();
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show($"Error: {ex.Message}", "Change Delimiter Error", 
                    System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }
        }
    }
}
