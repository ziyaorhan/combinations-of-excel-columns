using System;
using System.Collections.Generic;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
namespace CombinationOfExcelColumns
{
    public partial class FrmCombination : Form
    {
        #region Variables
        ExcelApp excelApp;
        Excel.Worksheet newSheet;
        List<int> filledColumnsIndex;
        int filledColumnsCount;
        List<string> rowValues;
        List<FilledColumnAndRows> filledColumnAndRows;
        List<HeaderCell> columnHeaderNames;
        List<DataCell> combinations;
        int currentRow;
        int currentColumn;
        string cellValue;
        #endregion

        public FrmCombination()
        {
            InitializeComponent();
            lblFileName.Text = String.Empty;
        }

        private void btnSelectFile_Click(object sender, EventArgs e)
        {
            ofdExcel.FileName = "*.xls";
            ofdExcel.Filter = "Excel File(*.xls)|*.xls";
            if (ofdExcel.ShowDialog() == DialogResult.OK)
            {
                lblFileName.Text = ofdExcel.FileName;
            }
        }

        private void btnCreateCombinations_Click(object sender, EventArgs e)
        {
            try
            {
                if (lblFileName.Text != String.Empty)
                {
                    Cursor.Current = Cursors.WaitCursor;
                    var startTime = DateTime.Now;
                    //1-
                    lblStatus.Text = "Opening excel application...";
                    excelApp = new ExcelApp(lblFileName.Text, Convert.ToInt32(nudWorkSheetNum.Value));
                    //2-
                    lblStatus.Text = "Collecting column headers ...";
                    var headers = GetColumnHeaderNames(Convert.ToInt32(nudHeaderStartRow.Value), Convert.ToInt32(nudHeaderEndRow.Value));
                    //3-
                    lblStatus.Text = "Collecting filled columns and rows ...";
                    GetFilledColumnsAndRows(GetFilledColumns(3), 3);
                    //4-
                    lblStatus.Text = "Creating new work sheet for combination...";
                    newSheet = excelApp.CreateNewSheet(string.Format("Combinations-{0}", Convert.ToInt32(nudWorkSheetNum.Value)));
                    //5-
                    lblStatus.Text = "Creating combinations...";
                    combinations = new List<DataCell>();
                    combinations = GetCombinationAsRecursive(0, new List<DataCell>());
                    //6-
                    lblStatus.Text = "Printing column headers to excel file...";
                    WriteColumnHeaderName(headers);
                    //7-
                    lblStatus.Text = "Printing combinations to excel file...";
                    WriteCombinationToNewSheet(Convert.ToInt32(nudDataStartRow.Value), combinations);
                    //8-
                    lblStatus.Text = "Closing excel...";
                    excelApp.CloseExcel();
                    var endTime = DateTime.Now;
                    TimeSpan timeSpan = (endTime - startTime);
                    lblStatus.Text = string.Format("The process was completed in {0} seconds.", timeSpan.TotalSeconds);
                    Cursor.Current = Cursors.Default;
                }
                else
                {
                    MessageBox.Show("Please select excel file.", "Info!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Hata!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                excelApp.CloseExcel();
            }
        }

        private List<FilledColumnAndRows> GetFilledColumnsAndRows(List<int> filledColumns, int startRowIndex)
        {
            filledColumnsCount = filledColumns.Count;
            filledColumnAndRows = new List<FilledColumnAndRows>();
            for (int i = 0; i < filledColumnsCount; i++)
            {
                lblStatus.Text = string.Format("Collecting rows of column {0} ...", filledColumns[i].ToString());
                filledColumnAndRows.Add(new FilledColumnAndRows
                {
                    FilledColumnIndeks = filledColumns[i],
                    RowsOfFilledColumn = GetRowsByColumnIndex(filledColumns[i], startRowIndex)
                });
            }
            return filledColumnAndRows;
        }

        private List<string> GetRowsByColumnIndex(int columnIndex, int startRowIndex)
        {
            try
            {
                rowValues = new List<string>();
                for (currentRow = startRowIndex; currentRow <= excelApp.RowCount; currentRow++)
                {
                    cellValue = Convert.ToString((excelApp.Range.Cells[currentRow, columnIndex] as Excel.Range).Value2);
                    if (!String.IsNullOrWhiteSpace(cellValue))
                    {
                        rowValues.Add(cellValue);
                    }
                }
                return rowValues;
            }
            catch
            {
                throw new Exception("An error occurred while gathering rows by column index.");
            }
        }

        private List<int> GetFilledColumns(int startRowIndex)
        {
            try
            {
                filledColumnsIndex = new List<int>();
                for (currentColumn = 1; currentColumn <= excelApp.ColumnCount; currentColumn++)
                {
                    for (currentRow = startRowIndex; currentRow <= excelApp.RowCount; currentRow++)
                    {
                        cellValue = Convert.ToString((excelApp.Range.Cells[currentRow, currentColumn] as Excel.Range).Value2);

                        if (String.IsNullOrWhiteSpace(cellValue))
                        {
                            break;
                        }
                        else
                        {
                            if (!filledColumnsIndex.Contains(currentColumn))
                            {
                                filledColumnsIndex.Add(currentColumn);
                            }
                        }
                    }
                }
                return filledColumnsIndex;
            }
            catch
            {
                throw new Exception("An error occurred while detecting full columns.");
            }
        }

        private List<HeaderCell> GetColumnHeaderNames(int startRowIndex, int endRowIndex)
        {
            try
            {
                columnHeaderNames = new List<HeaderCell>();
                for (currentColumn = 1; currentColumn <= excelApp.ColumnCount; currentColumn++)
                {
                    for (currentRow = startRowIndex; currentRow <= endRowIndex; currentRow++)
                    {
                        cellValue = Convert.ToString((excelApp.Range.Cells[currentRow, currentColumn] as Excel.Range).Value2);
                        if (!String.IsNullOrWhiteSpace(cellValue))
                        {
                            columnHeaderNames.Add(new HeaderCell
                            {
                                ColumnIndex = currentColumn,
                                RowIndex = currentRow,
                                Value = cellValue
                            });
                        }
                    }
                }
                return columnHeaderNames;
            }
            catch
            {
                throw new Exception("An error occurred while fetching column headers.");
            }
        }

        public List<DataCell> GetCombinationAsRecursive(int listStartIndeks, List<DataCell> output)
        {
            if (listStartIndeks < filledColumnAndRows.Count)
            {
                foreach (string row in filledColumnAndRows[listStartIndeks].RowsOfFilledColumn)
                {
                    List<DataCell> newList = new List<DataCell>();
                    newList.AddRange(output);
                    newList.Add(new DataCell
                    {
                        ColumnIndex = filledColumnAndRows[listStartIndeks].FilledColumnIndeks,
                        Value = row
                    });
                    if (newList.Count == filledColumnsCount)
                    {
                        combinations.AddRange(newList);
                    }
                    GetCombinationAsRecursive(listStartIndeks + 1, newList);
                }
            }
            return combinations;
        }

        private void WriteColumnHeaderName(List<HeaderCell> headerCells)
        {
            try
            {
                foreach (HeaderCell cell in headerCells)
                {
                    newSheet.Cells[cell.RowIndex, cell.ColumnIndex] = cell.Value;
                }
            }
            catch
            {
                Cursor.Current = Cursors.Default;
                throw new Exception("An error occurred while writing column headers.");
            }
        }

        private void WriteCombinationToNewSheet(int startRowIndex, List<DataCell> combinations)
        {
            try
            {
                int i = 0;
                int j = 0;
                foreach (DataCell cell in combinations)
                {
                    newSheet.Cells[startRowIndex + i, cell.ColumnIndex] = cell.Value;
                    j++;
                    if (j % filledColumnsCount == 0)
                    {
                        i++;
                    }
                }
            }
            catch
            {
                Cursor.Current = Cursors.Default;
                throw new Exception("An error occurred while writing combinations.");
            }
        }
    }
}
