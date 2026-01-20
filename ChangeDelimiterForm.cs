using System;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace SmartExcel
{
    public partial class ChangeDelimiterForm : Form
    {
        private readonly Excel.Application excelApp;
        private System.Windows.Forms.ComboBox columnComboBox;
        private System.Windows.Forms.ComboBox fromDelimiterComboBox;
        private System.Windows.Forms.ComboBox toDelimiterComboBox;
        private System.Windows.Forms.TextBox customFromDelimiterTextBox;
        private System.Windows.Forms.TextBox customToDelimiterTextBox;
        private System.Windows.Forms.Button executeButton;
        private System.Windows.Forms.Button cancelButton;
        private System.Windows.Forms.Label columnLabel;
        private System.Windows.Forms.Label fromDelimiterLabel;
        private System.Windows.Forms.Label toDelimiterLabel;

        public ChangeDelimiterForm(Excel.Application app)
        {
            this.excelApp = app;
            InitializeComponent();
            LoadColumns();
        }

        private void InitializeComponent()
        {
            this.Text = "Change Delimiter";
            this.Size = new System.Drawing.Size(400, 280);
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

            // From delimiter selection
            fromDelimiterLabel = new System.Windows.Forms.Label
            {
                Text = "From Delimiter:",
                Location = new System.Drawing.Point(20, 60),
                Size = new System.Drawing.Size(100, 20)
            };

            fromDelimiterComboBox = new System.Windows.Forms.ComboBox
            {
                Location = new System.Drawing.Point(130, 60),
                Size = new System.Drawing.Size(230, 25),
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            fromDelimiterComboBox.Items.AddRange(new object[] { "Comma (,)", "Semicolon (;)", "Pipe (|)", "Tab", "Line break", "Custom" });
            fromDelimiterComboBox.SelectedIndex = 0;
            fromDelimiterComboBox.SelectedIndexChanged += FromDelimiterComboBox_SelectedIndexChanged;

            // Custom from delimiter textbox
            customFromDelimiterTextBox = new System.Windows.Forms.TextBox
            {
                Location = new System.Drawing.Point(130, 90),
                Size = new System.Drawing.Size(230, 25),
                Visible = false
            };

            // To delimiter selection
            toDelimiterLabel = new System.Windows.Forms.Label
            {
                Text = "To Delimiter:",
                Location = new System.Drawing.Point(20, 130),
                Size = new System.Drawing.Size(100, 20)
            };

            toDelimiterComboBox = new System.Windows.Forms.ComboBox
            {
                Location = new System.Drawing.Point(130, 130),
                Size = new System.Drawing.Size(230, 25),
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            toDelimiterComboBox.Items.AddRange(new object[] { "Comma (,)", "Semicolon (;)", "Pipe (|)", "Tab", "Line break", "Custom" });
            toDelimiterComboBox.SelectedIndex = 1;
            toDelimiterComboBox.SelectedIndexChanged += ToDelimiterComboBox_SelectedIndexChanged;

            // Custom to delimiter textbox
            customToDelimiterTextBox = new System.Windows.Forms.TextBox
            {
                Location = new System.Drawing.Point(130, 160),
                Size = new System.Drawing.Size(230, 25),
                Visible = false
            };

            // Execute button
            executeButton = new System.Windows.Forms.Button
            {
                Text = "Execute",
                Location = new System.Drawing.Point(180, 200),
                Size = new System.Drawing.Size(80, 30)
            };
            executeButton.Click += ExecuteButton_Click;

            // Cancel button
            cancelButton = new System.Windows.Forms.Button
            {
                Text = "Cancel",
                Location = new System.Drawing.Point(280, 200),
                Size = new System.Drawing.Size(80, 30)
            };
            cancelButton.Click += (s, e) => this.Close();

            // Add controls
            this.Controls.Add(columnLabel);
            this.Controls.Add(columnComboBox);
            this.Controls.Add(fromDelimiterLabel);
            this.Controls.Add(fromDelimiterComboBox);
            this.Controls.Add(customFromDelimiterTextBox);
            this.Controls.Add(toDelimiterLabel);
            this.Controls.Add(toDelimiterComboBox);
            this.Controls.Add(customToDelimiterTextBox);
            this.Controls.Add(executeButton);
            this.Controls.Add(cancelButton);
        }

        private void FromDelimiterComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            UpdateFormSize();
        }

        private void ToDelimiterComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            UpdateFormSize();
        }

        private void UpdateFormSize()
        {
            bool showFromCustom = fromDelimiterComboBox.SelectedItem?.ToString() == "Custom";
            bool showToCustom = toDelimiterComboBox.SelectedItem?.ToString() == "Custom";

            customFromDelimiterTextBox.Visible = showFromCustom;
            customToDelimiterTextBox.Visible = showToCustom;

            int baseHeight = 280;
            int additionalHeight = 0;

            if (showFromCustom)
            {
                additionalHeight += 30;
                toDelimiterLabel.Location = new System.Drawing.Point(20, 130);
                toDelimiterComboBox.Location = new System.Drawing.Point(130, 130);
            }
            else
            {
                toDelimiterLabel.Location = new System.Drawing.Point(20, 100);
                toDelimiterComboBox.Location = new System.Drawing.Point(130, 100);
            }

            if (showToCustom)
            {
                additionalHeight += 30;
                customToDelimiterTextBox.Location = new System.Drawing.Point(130, toDelimiterComboBox.Location.Y + 30);
            }

            this.Size = new System.Drawing.Size(400, baseHeight + additionalHeight);
            executeButton.Location = new System.Drawing.Point(180, this.ClientSize.Height - 50);
            cancelButton.Location = new System.Drawing.Point(280, this.ClientSize.Height - 50);
        }

        private void LoadColumns()
        {
            try
            {
                var activeSheet = (Excel.Worksheet)excelApp.ActiveSheet;
                var usedRange = activeSheet.UsedRange;
                
                // Check if the worksheet has meaningful data (not just a single empty cell)
                if (usedRange.Rows.Count == 1 && usedRange.Columns.Count == 1 && 
                    (usedRange.Value2 == null || string.IsNullOrWhiteSpace(usedRange.Value2.ToString())))
                {
                    MessageBox.Show("No data found in the active sheet.", "Information", 
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
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

            string fromDelimiter = GetDelimiter(fromDelimiterComboBox, customFromDelimiterTextBox);
            string toDelimiter = GetDelimiter(toDelimiterComboBox, customToDelimiterTextBox);

            if (string.IsNullOrEmpty(fromDelimiter))
            {
                MessageBox.Show("Please specify a source delimiter.", "Validation Error", 
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (string.IsNullOrEmpty(toDelimiter))
            {
                MessageBox.Show("Please specify a target delimiter.", "Validation Error", 
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (fromDelimiter == toDelimiter)
            {
                MessageBox.Show("Source and target delimiters are the same. No changes will be made.", "Validation Error", 
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                ChangeDelimiter(columnComboBox.SelectedIndex + 1, fromDelimiter, toDelimiter);
                MessageBox.Show("Delimiter changed successfully!", "Success", 
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                this.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error changing delimiter: {ex.Message}", "Error", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private string GetDelimiter(System.Windows.Forms.ComboBox comboBox, System.Windows.Forms.TextBox textBox)
        {
            string selected = comboBox.SelectedItem?.ToString();
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
                    var customDelimiter = textBox.Text;
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

        private void ChangeDelimiter(int columnIndex, string fromDelimiter, string toDelimiter)
        {
            var activeSheet = (Excel.Worksheet)excelApp.ActiveSheet;
            var usedRange = activeSheet.UsedRange;
            
            excelApp.ScreenUpdating = false;
            excelApp.Calculation = Excel.XlCalculation.xlCalculationManual;

            try
            {
                object[,] data = (object[,])usedRange.Value2;
                int rowCount = usedRange.Rows.Count;

                for (int i = 1; i <= rowCount; i++)
                {
                    object cellValue = data[i, columnIndex];
                    if (cellValue != null)
                    {
                        string cellText = cellValue.ToString();
                        if (cellText.Contains(fromDelimiter))
                        {
                            string newValue = cellText.Replace(fromDelimiter, toDelimiter);
                            Excel.Range cell = (Excel.Range)activeSheet.Cells[i, columnIndex];
                            cell.Value2 = newValue;
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(cell);
                        }
                    }
                }
            }
            finally
            {
                excelApp.Calculation = Excel.XlCalculation.xlCalculationAutomatic;
                excelApp.ScreenUpdating = true;
            }
        }
    }
}
