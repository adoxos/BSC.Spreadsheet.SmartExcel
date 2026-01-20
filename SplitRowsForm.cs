using System;
using System.Windows.Forms;

namespace SmartExcel
{
    public partial class SplitRowsForm : Form
    {
        private readonly dynamic excelApp;
        private System.Windows.Forms.ComboBox columnComboBox;
        private System.Windows.Forms.ComboBox delimiterComboBox;
        private System.Windows.Forms.TextBox customDelimiterTextBox;
        private System.Windows.Forms.Button executeButton;
        private System.Windows.Forms.Button cancelButton;
        private System.Windows.Forms.Label columnLabel;
        private System.Windows.Forms.Label delimiterLabel;

        public SplitRowsForm(dynamic app)
        {
            this.excelApp = app;
            InitializeComponent();
            LoadColumns();
        }

        private void InitializeComponent()
        {
            this.Text = "Split Rows by Delimiter";
            this.Size = new System.Drawing.Size(400, 220);
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.StartPosition = FormStartPosition.CenterScreen;

            // Column selection
            columnLabel = new System.Windows.Forms.Label
            {
                Text = "Select Column:",
                Location = new System.Drawing.Point(20, 20),
                Size = new System.Drawing.Size(100, 20)
            };

            columnComboBox = new System.Windows.Forms.ComboBox
            {
                Location = new System.Drawing.Point(130, 20),
                Size = new System.Drawing.Size(230, 25),
                DropDownStyle = ComboBoxStyle.DropDownList
            };

            // Delimiter selection
            delimiterLabel = new System.Windows.Forms.Label
            {
                Text = "Select Delimiter:",
                Location = new System.Drawing.Point(20, 60),
                Size = new System.Drawing.Size(100, 20)
            };

            delimiterComboBox = new System.Windows.Forms.ComboBox
            {
                Location = new System.Drawing.Point(130, 60),
                Size = new System.Drawing.Size(230, 25),
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            delimiterComboBox.Items.AddRange(new object[] { "Comma (,)", "Semicolon (;)", "Pipe (|)", "Tab", "Line break", "Custom" });
            delimiterComboBox.SelectedIndex = 0;
            delimiterComboBox.SelectedIndexChanged += DelimiterComboBox_SelectedIndexChanged;

            // Custom delimiter textbox
            customDelimiterTextBox = new System.Windows.Forms.TextBox
            {
                Location = new System.Drawing.Point(130, 90),
                Size = new System.Drawing.Size(230, 25),
                Visible = false
            };

            // Execute button
            executeButton = new System.Windows.Forms.Button
            {
                Text = "Execute",
                Location = new System.Drawing.Point(180, 140),
                Size = new System.Drawing.Size(80, 30)
            };
            executeButton.Click += ExecuteButton_Click;

            // Cancel button
            cancelButton = new System.Windows.Forms.Button
            {
                Text = "Cancel",
                Location = new System.Drawing.Point(280, 140),
                Size = new System.Drawing.Size(80, 30)
            };
            cancelButton.Click += (s, e) => this.Close();

            // Add controls
            this.Controls.Add(columnLabel);
            this.Controls.Add(columnComboBox);
            this.Controls.Add(delimiterLabel);
            this.Controls.Add(delimiterComboBox);
            this.Controls.Add(customDelimiterTextBox);
            this.Controls.Add(executeButton);
            this.Controls.Add(cancelButton);
        }

        private void DelimiterComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            bool showCustom = string.Equals(delimiterComboBox.SelectedItem?.ToString(), "Custom", StringComparison.Ordinal);
            customDelimiterTextBox.Visible = showCustom;
            if (customDelimiterTextBox.Visible)
            {
                this.Size = new System.Drawing.Size(400, 250);
                executeButton.Location = new System.Drawing.Point(180, 170);
                cancelButton.Location = new System.Drawing.Point(280, 170);
            }
            else
            {
                this.Size = new System.Drawing.Size(400, 220);
                executeButton.Location = new System.Drawing.Point(180, 140);
                cancelButton.Location = new System.Drawing.Point(280, 140);
            }
        }

        private void LoadColumns()
        {
            try
            {
                dynamic activeSheet = excelApp.ActiveSheet;
                var usedRange = activeSheet.UsedRange;
                
                // Check if the worksheet has meaningful data (not just a single empty cell)
                if (usedRange.Rows.Count == 1 && usedRange.Columns.Count == 1)
                {
                    var value = usedRange.Value2;
                    if (value == null || string.IsNullOrWhiteSpace(value.ToString()))
                    {
                        MessageBox.Show("No data found in the active sheet.", "Information", 
                            MessageBoxButtons.OK, MessageBoxIcon.Information);
                        return;
                    }
                }

                for (int i = 1; i <= usedRange.Columns.Count; i++)
                {
                    string columnLetter = ExcelHelpers.GetColumnLetter(i);
                    columnComboBox.Items.Add($"Column {columnLetter}");
                }

                if (columnComboBox.Items.Count > 0)
                    columnComboBox.SelectedIndex = 0;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading columns: {ex.Message}", "Error", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ExecuteButton_Click(object sender, EventArgs e)
        {
            if (columnComboBox.SelectedIndex < 0)
            {
                MessageBox.Show("Please select a column.", "Validation Error", 
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            string delimiter = GetSelectedDelimiter();
            if (string.IsNullOrEmpty(delimiter))
            {
                MessageBox.Show("Please specify a delimiter.", "Validation Error", 
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                SplitRows(columnComboBox.SelectedIndex + 1, delimiter);
                MessageBox.Show("Rows split successfully!", "Success", 
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                this.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error splitting rows: {ex.Message}", "Error", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private string GetSelectedDelimiter()
        {
            string selected = delimiterComboBox.SelectedItem?.ToString();
            if (string.IsNullOrEmpty(selected))
                return ",";
                
            switch (selected)
            {
                case "Comma (,)":
                    return ",";
                case "Semicolon (;)":
                    return ";";
                case "Pipe (|)":
                    return "|";
                case "Tab":
                    return "\t";
                case "Line break":
                    return Environment.NewLine;
                case "Custom":
                    var customDelimiter = customDelimiterTextBox.Text;
                    if (string.IsNullOrWhiteSpace(customDelimiter))
                    {
                        // Return null to indicate missing custom delimiter
                        return null;
                    }
                    return customDelimiter;
                default:
                    return ",";
            }
        }

        private void SplitRows(int columnIndex, string delimiter)
        {
            dynamic activeSheet = excelApp.ActiveSheet;
            var usedRange = activeSheet.UsedRange;

            excelApp.ScreenUpdating = false;
            excelApp.Calculation = ExcelComConstants.CalculationManual;

            try
            {
                // Get all data first
                object[,] data = (object[,])usedRange.Value2;
                int rowCount = usedRange.Rows.Count;
                int colCount = usedRange.Columns.Count;

                // Process from bottom to top to avoid shifting issues
                for (int i = rowCount; i >= 1; i--)
                {
                    object cellValue = data[i, columnIndex];
                    if (cellValue != null)
                    {
                        string cellText = cellValue.ToString();
                        if (cellText.Contains(delimiter))
                        {
                            string[] parts = cellText.Split(new[] { delimiter }, StringSplitOptions.RemoveEmptyEntries);
                            
                            if (parts.Length > 1)
                            {
                                // Update the first row with the first part
                                dynamic cellToUpdate = activeSheet.Cells[i, columnIndex];
                                cellToUpdate.Value2 = parts[0];
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(cellToUpdate);

                                // Insert new rows for remaining parts
                                for (int j = 1; j < parts.Length; j++)
                                {
                                    dynamic rowToInsert = activeSheet.Rows[i + j];
                                    rowToInsert.Insert(ExcelComConstants.ShiftDown, ExcelComConstants.FormatFromLeftOrAbove);
                                    System.Runtime.InteropServices.Marshal.ReleaseComObject(rowToInsert);
                                    
                                    // Copy entire row data
                                    for (int k = 1; k <= colCount; k++)
                                    {
                                        dynamic targetCell = activeSheet.Cells[i + j, k];
                                        targetCell.Value2 = (k == columnIndex) ? (object)parts[j] : data[i, k];
                                        System.Runtime.InteropServices.Marshal.ReleaseComObject(targetCell);
                                    }
                                }
                            }
                        }
                    }
                }
            }
            finally
            {
                excelApp.Calculation = ExcelComConstants.CalculationAutomatic;
                excelApp.ScreenUpdating = true;
            }
        }
    }
}
