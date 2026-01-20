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

        public override string GetCustomUI(string RibbonID)
        {
            return @"
                <customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui' loadImage='LoadImage'>
                    <ribbon>
                        <tabs>
                            <tab id='SmartExcelTab' label='SmartExcel'>
                                <group id='SmartExcelGroup' label='Data Tools'>
                                    <button id='SplitRowsButton' 
                                            label='Split Rows' 
                                            size='large' 
                                            onAction='OnSplitRowsButton'
                                            imageMso='TableSplitCells' />
                                    <button id='ChangeDelimiterButton' 
                                            label='Change Delimiter' 
                                            size='large' 
                                            onAction='OnChangeDelimiterButton'
                                            imageMso='ConvertTableToRange' />
                                </group>
                            </tab>
                        </tabs>
                    </ribbon>
                </customUI>";
        }

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
