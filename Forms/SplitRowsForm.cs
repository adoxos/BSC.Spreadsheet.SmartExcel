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
        private System.Windows.Forms.CheckBox fileNameAwareCheckBox;
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
            this.Size = new System.Drawing.Size(400, 250);
            this.MinimumSize = new System.Drawing.Size(400, 250);
            this.FormBorderStyle = FormBorderStyle.Sizable;
            this.MaximizeBox = true;
            this.MinimizeBox = true;
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
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
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
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
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
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
                Visible = false
            };

            // File name aware checkbox
            fileNameAwareCheckBox = new System.Windows.Forms.CheckBox
            {
                Text = "File name aware (ignore delimiters in file paths)",
                Location = new System.Drawing.Point(20, 120),
                Size = new System.Drawing.Size(340, 20),
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
                Checked = false
            };

            // Execute button
            executeButton = new System.Windows.Forms.Button
            {
                Text = "Execute",
                Location = new System.Drawing.Point(180, 170),
                Size = new System.Drawing.Size(80, 30),
                Anchor = AnchorStyles.Bottom | AnchorStyles.Right
            };
            executeButton.Click += ExecuteButton_Click;

            // Cancel button
            cancelButton = new System.Windows.Forms.Button
            {
                Text = "Cancel",
                Location = new System.Drawing.Point(280, 170),
                Size = new System.Drawing.Size(80, 30),
                Anchor = AnchorStyles.Bottom | AnchorStyles.Right
            };
            cancelButton.Click += (s, e) => this.Close();

            // Add controls
            this.Controls.Add(columnLabel);
            this.Controls.Add(columnComboBox);
            this.Controls.Add(delimiterLabel);
            this.Controls.Add(delimiterComboBox);
            this.Controls.Add(customDelimiterTextBox);
            this.Controls.Add(fileNameAwareCheckBox);
            this.Controls.Add(executeButton);
            this.Controls.Add(cancelButton);
        }

        private void DelimiterComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            bool showCustom = string.Equals(delimiterComboBox.SelectedItem?.ToString(), "Custom", StringComparison.Ordinal);
            customDelimiterTextBox.Visible = showCustom;
            if (customDelimiterTextBox.Visible)
            {
                this.Size = new System.Drawing.Size(400, 280);
                fileNameAwareCheckBox.Location = new System.Drawing.Point(20, 150);
                executeButton.Location = new System.Drawing.Point(180, 200);
                cancelButton.Location = new System.Drawing.Point(280, 200);
            }
            else
            {
                this.Size = new System.Drawing.Size(400, 250);
                fileNameAwareCheckBox.Location = new System.Drawing.Point(20, 120);
                executeButton.Location = new System.Drawing.Point(180, 170);
                cancelButton.Location = new System.Drawing.Point(280, 170);
            }
        }

        private void LoadColumns()
        {
            ExcelHelpers.TryPopulateColumnComboBox(excelApp, columnComboBox);
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
                SplitRows(columnComboBox.SelectedIndex + 1, delimiter, fileNameAwareCheckBox.Checked);
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
                        return null;
                    }
                    return customDelimiter;
                default:
                    return ",";
            }
        }

        private void SplitRows(int columnIndex, string delimiter, bool fileNameAware)
        {
            dynamic activeSheet = excelApp.ActiveSheet;
            var usedRange = activeSheet.UsedRange;

            if (ExcelHelpers.IsUsedRangeEmpty(usedRange))
            {
                MessageBox.Show("No data found in the active sheet.", "Information",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            ExcelRangeData rangeData = ExcelHelpers.NormalizeRangeValues(usedRange);
            object[,] data = rangeData.Values;
            int rowCount = rangeData.RowCount;
            int colCount = rangeData.ColumnCount;

            excelApp.ScreenUpdating = false;
            excelApp.Calculation = ExcelComConstants.CalculationManual;

            try
            {
                for (int i = rowCount; i >= 1; i--)
                {
                    object cellValue = data[i, columnIndex];
                    if (cellValue == null)
                    {
                        continue;
                    }

                    string cellText = cellValue.ToString();
                    if (!cellText.Contains(delimiter))
                    {
                        continue;
                    }

                    string[] parts = fileNameAware
                        ? SplitFileNameAware(cellText, delimiter)
                        : cellText.Split(new[] { delimiter }, StringSplitOptions.RemoveEmptyEntries);

                    if (parts.Length <= 1)
                    {
                        continue;
                    }

                    dynamic cellToUpdate = activeSheet.Cells[i, columnIndex];
                    cellToUpdate.Value2 = parts[0];
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(cellToUpdate);

                    for (int j = 1; j < parts.Length; j++)
                    {
                        dynamic rowToInsert = activeSheet.Rows[i + j];
                        rowToInsert.Insert(ExcelComConstants.ShiftDown, ExcelComConstants.FormatFromLeftOrAbove);
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(rowToInsert);

                        for (int k = 1; k <= colCount; k++)
                        {
                            dynamic targetCell = activeSheet.Cells[i + j, k];
                            targetCell.Value2 = (k == columnIndex) ? (object)parts[j] : data[i, k];
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(targetCell);
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

        private string[] SplitFileNameAware(string text, string delimiter)
        {
            var parts = new System.Collections.Generic.List<string>();
            var currentPart = new System.Text.StringBuilder();
            bool inQuotes = false;
            bool inFileName = false;
            int lastDotPosition = -1;

            for (int i = 0; i < text.Length; i++)
            {
                char c = text[i];

                if (c == '"')
                {
                    inQuotes = !inQuotes;
                    currentPart.Append(c);
                    continue;
                }

                if (c == '.' && !inQuotes)
                {
                    lastDotPosition = currentPart.Length;
                }

                if (!inQuotes && !inFileName)
                {
                    if ((c == '\\' && i + 1 < text.Length && text[i + 1] == '\\') ||
                        (char.IsLetter(c) && i + 2 < text.Length && text[i + 1] == ':' && text[i + 2] == '\\') ||
                        (c == '/' && i > 0 && char.IsWhiteSpace(text[i - 1])))
                    {
                        inFileName = true;
                    }
                }

                if (!inQuotes && i + delimiter.Length <= text.Length)
                {
                    string segment = text.Substring(i, delimiter.Length);
                    if (segment == delimiter)
                    {
                        bool isRealDelimiter = true;

                        if (lastDotPosition >= 0)
                        {
                            int charsAfterDot = currentPart.Length - lastDotPosition - 1;

                            int extensionLength = charsAfterDot;
                            int lookAhead = i + delimiter.Length;
                            while (lookAhead < text.Length && char.IsLetterOrDigit(text[lookAhead]) && extensionLength < 6)
                            {
                                extensionLength++;
                                lookAhead++;
                            }

                            if (extensionLength >= 2 && extensionLength <= 5)
                            {
                                if (lookAhead >= text.Length || char.IsWhiteSpace(text[lookAhead]) ||
                                    text[lookAhead] == delimiter[0])
                                {
                                    isRealDelimiter = false;
                                }
                            }
                        }

                        if (inFileName)
                        {
                            int lookAheadStart = i + delimiter.Length;
                            if (lookAheadStart < text.Length)
                            {
                                char nextChar = text[lookAheadStart];
                                if (!char.IsWhiteSpace(nextChar))
                                {
                                    isRealDelimiter = false;
                                }
                                else
                                {
                                    inFileName = false;
                                }
                            }
                        }

                        if (isRealDelimiter)
                        {
                            string part = currentPart.ToString().Trim();
                            if (!string.IsNullOrEmpty(part))
                            {
                                parts.Add(part);
                            }
                            currentPart.Clear();
                            lastDotPosition = -1;
                            inFileName = false;
                            i += delimiter.Length - 1;
                            continue;
                        }
                    }
                }

                if ((inFileName || lastDotPosition >= 0) && !inQuotes && char.IsWhiteSpace(c))
                {
                    if (lastDotPosition >= 0)
                    {
                        int extensionLength = currentPart.Length - lastDotPosition - 1;
                        if (extensionLength >= 2 && extensionLength <= 5)
                        {
                            inFileName = false;
                            lastDotPosition = -1;
                        }
                    }
                    else
                    {
                        inFileName = false;
                    }
                }

                currentPart.Append(c);
            }

            string finalPart = currentPart.ToString().Trim();
            if (!string.IsNullOrEmpty(finalPart))
            {
                parts.Add(finalPart);
            }

            return parts.Count > 0 ? parts.ToArray() : new[] { text };
        }
    }
}
