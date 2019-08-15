using System;
using System.Collections.Generic;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
namespace CombinationOfExcelColumns
{
    public partial class FrmCombination : Form
    {
        #region Değişkenler
        ExcelApp excelApp;
        Excel.Worksheet newSheet;
        List<int> filledColumnsIndex;// dolu sütunların indeksleri
        int filledColumnsCount;// dolu sütun sayısı
        List<string> rowValues;// dolu sütunların satırlarını toplamak için
        List<FilledColumnAndRows> filledColumnAndRows; //dolu sütun indeksleri ve satırları
        List<HeaderCell> columnHeaderNames;//sütun başlıkları
        List<DataCell> combinations;//kombinasyonlar
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
                    lblStatus.Text = "Excel uygulaması açılıyor...";
                    excelApp = new ExcelApp(lblFileName.Text, Convert.ToInt32(nudWorkSheetNum.Value));
                    //2-
                    lblStatus.Text = "Dolu sütunlar saptanıyor...";
                    GetFilledColumnsAndRows(GetFilledColumns(3), 3);
                    //3-
                    lblStatus.Text = "Kombinasyonlar için yeni çalışma sayfası oluşturuluyor...";
                    newSheet = excelApp.CreateNewSheet(string.Format("Kombinasyonlar-{0}", Convert.ToInt32(nudWorkSheetNum.Value)));
                    //4-
                    lblStatus.Text = "Kombinasyonlar oluşturuluyor...";
                    combinations = new List<DataCell>();
                    combinations = GetCombinationAsRecursive(0, new List<DataCell>());
                    //5-
                    lblStatus.Text = "Sütun başlıkları yazdırılıyor...";
                    var headers = GetColumnHeaderNames(Convert.ToInt32(nudHeaderStartRow.Value), Convert.ToInt32(nudHeaderEndRow.Value));
                    WriteColumnHeaderName(headers);
                    //6-
                    lblStatus.Text = "Kombinasyonlar yazdırılıyor...";
                    WriteCombinationToNewSheet(Convert.ToInt32(nudDataStartRow.Value), combinations);
                    //7-
                    lblStatus.Text = "Excel kapatılıyor...";
                    excelApp.CloseExcel();
                    var endTime = DateTime.Now;
                    TimeSpan timeSpan = (endTime - startTime);
                    lblStatus.Text = string.Format("İşlem {0} saniyede tamamlandı.", timeSpan.TotalSeconds);
                    Cursor.Current = Cursors.Default;
                }
                else
                {
                    MessageBox.Show("Excel dosyası seçiniz!", "Bilgi!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Hata!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private List<FilledColumnAndRows> GetFilledColumnsAndRows(List<int> filledColumns, int startRowIndex)
        {
            filledColumnsCount = filledColumns.Count;
            filledColumnAndRows = new List<FilledColumnAndRows>();
            for (int i = 0; i < filledColumnsCount; i++)
            {
                lblStatus.Text = filledColumns[i].ToString() + ". sütun için satırlar toplanıyor...";
                filledColumnAndRows.Add(new FilledColumnAndRows
                {
                    FilledColumnIndeks = filledColumns[i],//dolu sütun indeksi
                    RowsOfFilledColumn = GetRowsByColumnIndex(filledColumns[i], startRowIndex)// dolu sütun indeksine göre satırlar.
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
                    cellValue = (string)(excelApp.Range.Cells[currentRow, columnIndex] as Excel.Range).Value2;
                    if (!String.IsNullOrWhiteSpace(cellValue))
                    {
                        rowValues.Add(cellValue);
                    }
                }
                return rowValues;
            }
            catch
            {
                throw new Exception("Sütun indeksine göre satırlar çekilirken bir hata oluştu.\r\nLütfen işlemi yeniden başlatınız.");
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
                        cellValue = (string)(excelApp.Range.Cells[currentRow, currentColumn] as Excel.Range).Value2;

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
                throw new Exception("Dolu sütunlar saptanırken bir hata oluştu.\r\nLütfen işlemi yeniden başlatınız.");
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
                        cellValue = (string)(excelApp.Range.Cells[currentRow, currentColumn] as Excel.Range).Value2;
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
                throw new Exception("Sütun başlıkları çekilirken bir hata oluştu.\r\nLütfen işlemi yeniden başlatınız.");
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
                throw new Exception("Sütun başlıkları yazılırken bir hata oluştu.\r\nLütfen işlemi yeniden başlatınız.");
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
                throw new Exception("Kombinasyonlar yazılırken bir hata oluştu.\r\nLütfen işlemi yeniden başlatınız.");
            }
        }
    }
}
