using System;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using LicenseContext = OfficeOpenXml.LicenseContext;

namespace ExcelColumnReorder
{
    public partial class MainForm : Form
    {
        private DataTable _dataTable;
        private string _filePath;

        private readonly string[] _orderedColumns = {
            "Comprobante",
            "Fecha elaboración",
            "Base gravada",
            "IVA",
            "Total",
            "Identificación",
            "Suc",
            "Nombre tercero"
        };

        private readonly string[] _columnsToRemove = {
            "Base exenta",
            "Impoconsumo",
            "AD-Valorem",
            "Cargo en totales",
            "Descuento en totales"
        };

        public MainForm()
        {
            InitializeComponent();
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }

        private void btnImport_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog { Filter = "Excel Files|*.xlsx;*.xls" })
            {
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    _filePath = openFileDialog.FileName;
                    LoadExcelFile();
                }
            }
        }

        private void LoadExcelFile()
        {
            using (var package = new ExcelPackage(new FileInfo(_filePath)))
            {
                var worksheet = package.Workbook.Worksheets[0];
                _dataTable = new DataTable();

                int headerRow = 1;
                for (int row = 1; row <= worksheet.Dimension.End.Row; row++)
                {
                    if (worksheet.Cells[row, 1].Text == "Comprobante")
                    {
                        headerRow = row;
                        break;
                    }
                }

                for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
                {
                    _dataTable.Columns.Add(worksheet.Cells[headerRow, col].Text);
                }

                for (int row = headerRow + 1; row <= worksheet.Dimension.End.Row; row++)
                {
                    if (string.IsNullOrWhiteSpace(worksheet.Cells[row, 1].Text))
                        continue;

                    var dataRow = _dataTable.NewRow();
                    for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
                    {
                        dataRow[col - 1] = worksheet.Cells[row, col].Text;
                    }
                    _dataTable.Rows.Add(dataRow);
                }
            }
            ReorderAndFilterColumns();
        }

        private void ReorderAndFilterColumns()
        {
            _dataTable = _dataTable.DefaultView.ToTable(false, _dataTable.Columns
                .Cast<DataColumn>()
                .Where(col => _orderedColumns.Contains(col.ColumnName))
                .OrderBy(col => Array.IndexOf(_orderedColumns, col.ColumnName))
                .Select(col => col.ColumnName)
                .ToArray());

            dataGridView1.DataSource = _dataTable;
        }

        private void btnExport_Click(object sender, EventArgs e)
        {
            using (SaveFileDialog saveFileDialog = new SaveFileDialog { Filter = "Excel Files|*.xlsx" })
            {
                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    ExportExcel(saveFileDialog.FileName);
                }
            }
        }

        private void ExportExcel(string filePath)
        {
            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Hoja1");

                // Apply document header formatting
                worksheet.Cells["A1"].Value = "Libro oficial de ventas";
                worksheet.Cells["A1"].Style.Font.Bold = true;
                worksheet.Cells["A1"].Style.Font.Size = 14;

                worksheet.Cells["A2"].Value = "IMPORTADORA DE INSERTOS SAS";
                worksheet.Cells["A2"].Style.Font.Bold = true;

                worksheet.Cells["A3"].Value = "900433608-0";
                worksheet.Cells["A3"].Style.Font.Bold = true;

                // Header row formatting
                int dataStartRow = 7;
                var headerRange = worksheet.Cells[dataStartRow, 1, dataStartRow, _orderedColumns.Length];

                headerRange.Style.Font.Bold = true;
                headerRange.Style.Fill.PatternType = ExcelFillStyle.Solid;
                headerRange.Style.Fill.BackgroundColor.SetColor(Color.LightGray);
                headerRange.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                // Add column headers
                for (int col = 0; col < _dataTable.Columns.Count; col++)
                {
                    worksheet.Cells[dataStartRow, col + 1].Value = _dataTable.Columns[col].ColumnName;
                }

                // Add data rows with formatting
                for (int row = 0; row < _dataTable.Rows.Count; row++)
                {
                    for (int col = 0; col < _dataTable.Columns.Count; col++)
                    {
                        var cell = worksheet.Cells[dataStartRow + row + 1, col + 1];
                        cell.Value = _dataTable.Rows[row][col];

                        // Format numeric columns
                        if (_dataTable.Columns[col].ColumnName == "Base gravada" ||
                            _dataTable.Columns[col].ColumnName == "IVA" ||
                            _dataTable.Columns[col].ColumnName == "Total")
                        {
                            cell.Style.Numberformat.Format = "#,##0";
                        }

                        // Format date column
                        if (_dataTable.Columns[col].ColumnName == "Fecha elaboración")
                        {
                            cell.Style.Numberformat.Format = "dd/MM/yyyy";
                        }
                    }
                }

                // Add borders to data
                var dataRange = worksheet.Cells[dataStartRow, 1, dataStartRow + _dataTable.Rows.Count, _orderedColumns.Length];
                dataRange.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                dataRange.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                dataRange.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                dataRange.Style.Border.Right.Style = ExcelBorderStyle.Thin;

                // Auto-fit columns
                worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();

                package.SaveAs(new FileInfo(filePath));
            }
            MessageBox.Show("Archivo exportado correctamente con el formato de referencia.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }
}