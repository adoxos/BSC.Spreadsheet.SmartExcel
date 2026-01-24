using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace SmartExcel
{
    public partial class TrimColumnForm : Form
    {
        private readonly dynamic excelApp;
        private ComboBox columnComboBox;
        private RadioButton ltrimRadioButton;
        private RadioButton rtrimRadioButton;
        private RadioButton fullTrimRadioButton;
        private Button executeButton;
        private Button cancelButton;
        private Label columnLabel;
        private Label modeLabel;

        private enum TrimMode
        {
            Left,
            Right,
            Both
        }

        public TrimColumnForm(dynamic app)
        {
            excelApp = app;
            InitializeComponent();
            LoadColumns();
        }

        private void InitializeComponent()
        {
            Text = "LTRIM / RTRIM";
            Size = new System.Drawing.Size(380, 230);
            MinimumSize = new System.Drawing.Size(380, 230);
            FormBorderStyle = FormBorderStyle.Sizable;
            MaximizeBox = true;
            MinimizeBox = true;
            StartPosition = FormStartPosition.CenterScreen;

            columnLabel = new Label
            {
                Text = "Select Column:",
                Location = new System.Drawing.Point(20, 20),
                Size = new System.Drawing.Size(100, 20)
            };

            columnComboBox = new ComboBox
            {
                Location = new System.Drawing.Point(130, 18),
                Size = new System.Drawing.Size(210, 25),
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
                DropDownStyle = ComboBoxStyle.DropDownList
            };

            modeLabel = new Label
            {
                Text = "Trim Mode:",
                Location = new System.Drawing.Point(20, 65),
                Size = new System.Drawing.Size(100, 20)
            };

            ltrimRadioButton = new RadioButton
            {
                Text = "LTRIM (leading)",
                Location = new System.Drawing.Point(130, 60),
                Size = new System.Drawing.Size(150, 20)
            };

            rtrimRadioButton = new RadioButton
            {
                Text = "RTRIM (trailing)",
                Location = new System.Drawing.Point(130, 85),
                Size = new System.Drawing.Size(150, 20)
            };

            fullTrimRadioButton = new RadioButton
            {
                Text = "Both (trim)",
                Location = new System.Drawing.Point(130, 110),
                Size = new System.Drawing.Size(150, 20),
                Checked = true
            };

            executeButton = new Button
            {
                Text = "Execute",
                Location = new System.Drawing.Point(170, 155),
                Size = new System.Drawing.Size(80, 30),
                Anchor = AnchorStyles.Bottom | AnchorStyles.Right
            };
            executeButton.Click += ExecuteButton_Click;

            cancelButton = new Button
            {
                Text = "Cancel",
                Location = new System.Drawing.Point(260, 155),
                Size = new System.Drawing.Size(80, 30),
                Anchor = AnchorStyles.Bottom | AnchorStyles.Right
            };
            cancelButton.Click += (s, e) => Close();

            Controls.Add(columnLabel);
            Controls.Add(columnComboBox);
            Controls.Add(modeLabel);
            Controls.Add(ltrimRadioButton);
            Controls.Add(rtrimRadioButton);
            Controls.Add(fullTrimRadioButton);
            Controls.Add(executeButton);
            Controls.Add(cancelButton);
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

            TrimMode mode = GetSelectedTrimMode();

            try
            {
                int updatedCells = TrimColumn(columnComboBox.SelectedIndex + 1, mode);
                string message = updatedCells > 0
                    ? $"Trim completed. Updated {updatedCells} cell(s)."
                    : "Trim completed. No changes were necessary.";

                MessageBox.Show(message, "LTRIM / RTRIM",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error trimming column: {ex.Message}", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private TrimMode GetSelectedTrimMode()
        {
            if (ltrimRadioButton.Checked)
            {
                return TrimMode.Left;
            }

            if (rtrimRadioButton.Checked)
            {
                return TrimMode.Right;
            }

            return TrimMode.Both;
        }

        private int TrimColumn(int columnIndex, TrimMode mode)
        {
            dynamic activeSheet = excelApp.ActiveSheet;
            var usedRange = activeSheet.UsedRange;

            if (ExcelHelpers.IsUsedRangeEmpty(usedRange))
            {
                MessageBox.Show("No data found in the active sheet.", "Information",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return 0;
            }

            int updatedCells = 0;

            excelApp.ScreenUpdating = false;
            excelApp.Calculation = ExcelComConstants.CalculationManual;

            try
            {
                int rowCount = usedRange.Rows.Count;

                for (int i = 1; i <= rowCount; i++)
                {
                    dynamic cell = activeSheet.Cells[i, columnIndex];
                    object cellValue = cell.Value2;

                    if (cellValue != null)
                    {
                        string originalText = cellValue.ToString();
                        string trimmedText = ApplyTrim(originalText, mode);

                        if (!string.Equals(originalText, trimmedText, StringComparison.Ordinal))
                        {
                            cell.Value2 = trimmedText;
                            updatedCells++;
                        }
                    }

                    Marshal.ReleaseComObject(cell);
                }
            }
            finally
            {
                excelApp.Calculation = ExcelComConstants.CalculationAutomatic;
                excelApp.ScreenUpdating = true;
            }

            return updatedCells;
        }

        private static string ApplyTrim(string input, TrimMode mode)
        {
            return mode switch
            {
                TrimMode.Left => input.TrimStart(),
                TrimMode.Right => input.TrimEnd(),
                _ => input.Trim()
            };
        }
    }
}
