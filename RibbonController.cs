using System;
using System.Runtime.InteropServices;
using System.Drawing;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;

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
                dynamic excelApp = ExcelDnaUtil.Application;
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
                dynamic excelApp = ExcelDnaUtil.Application;
                var changeDelimiterForm = new ChangeDelimiterForm(excelApp);
                changeDelimiterForm.ShowDialog();
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show($"Error: {ex.Message}", "Change Delimiter Error", 
                    System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }
        }

        public void OnTrimColumnButton(IRibbonControl control)
        {
            try
            {
                dynamic excelApp = ExcelDnaUtil.Application;
                var trimForm = new TrimColumnForm(excelApp);
                trimForm.ShowDialog();
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show($"Error: {ex.Message}", "LTRIM / RTRIM Error", 
                    System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }
        }

        public Bitmap GetSplitRowsImage(IRibbonControl control)
        {
            return IconFactory.SplitRowsIcon;
        }

        public Bitmap GetChangeDelimiterImage(IRibbonControl control)
        {
            return IconFactory.ChangeDelimiterIcon;
        }

        public Bitmap GetTrimColumnImage(IRibbonControl control)
        {
            return IconFactory.TrimColumnIcon;
        }
    }
}
