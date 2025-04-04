using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Data;
using LicenseContext = OfficeOpenXml.LicenseContext;

namespace ExcelColumnReorder
{
    public partial class MainForm : Form
    {
        private DataTable _dataTable;
        private string _filePath;

        private readonly string[] _orderedColumns = {
            "Comprobante",
            "Fecha elaboraci�n",
            "Base gravada",
            "IVA",
            "Total",
            "Identificaci�n",
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

            dataGridView1.ReadOnly = true;
            dataGridView1.AllowUserToAddRows = false;
            dataGridView1.AllowUserToDeleteRows = false;
            dataGridView1.AllowUserToOrderColumns = false;
            dataGridView1.AllowUserToResizeColumns = false;
            dataGridView1.AllowUserToResizeRows = false;

            dataGridView1.DataSource = _dataTable;
        }

        private void btnExport_Click(object sender, EventArgs e)
        {
            string defaultFileName = $"{DateTime.Now:dd MM yyyy} Libro Oficial de Ventas.xlsx";

            using (SaveFileDialog saveFileDialog = new SaveFileDialog
            {
                Filter = "Excel Files|*.xlsx",
                FileName = defaultFileName 
            })
            {
                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    ExportExcel(saveFileDialog.FileName);
                
                    // Limpiar el DataGridView despu�s de exportar
                    ClearDataGridView();
                }
            }
        }

        private void ClearDataGridView()
        {
            dataGridView1.DataSource = null;
            dataGridView1.Rows.Clear();
            dataGridView1.Columns.Clear();
            _dataTable?.Clear();
        }

        private void ExportExcel(string filePath)
        {
            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add($"Hoja {DateTime.Now:dd MM yyyy}");


                // Definir el rango de celdas a colorear
                var rango = worksheet.Cells["A1:H2"];
                // Configurar el color de fondo en una sola operaci�n
                rango.Style.Fill.PatternType = ExcelFillStyle.Solid;
                rango.Style.Fill.BackgroundColor.SetColor(Color.DodgerBlue);


                // Apply document header formatting

                // Unir celdas de A1 a H1
                worksheet.Cells["A1:H1"].Merge = true;
                // Establecer el valor de la celda
                worksheet.Cells["A1"].Value = "Libro oficial de ventas";
                // Centrar el texto horizontalmente y verticalmente
                worksheet.Cells["A1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["A1"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["A1"].Style.Font.Bold = true;
                worksheet.Cells["A1"].Style.Font.Size = 30;
                worksheet.Cells["A1"].Style.Font.Color.SetColor(Color.White);


                // Unir celdas de A2 a H2
                worksheet.Cells["A2:H2"].Merge = true;
                // Establecer el valor de la celda
                worksheet.Cells["A2"].Value = "IMPORTADORA DE INSERTOS SAS";
                worksheet.Cells["A2"].Style.Font.Size = 15;
                // Centrar el texto horizontalmente y verticalmente
                worksheet.Cells["A2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["A2"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["A2"].Style.Font.Bold = true;
                worksheet.Cells["A2"].Style.Font.Color.SetColor(Color.White);

                // Inmovilizar filas 1 a 3 (comienza la vista en A4)
                worksheet.View.FreezePanes(4, 1);


                // Header row formatting
                int dataStartRow = 3;
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
                        if (_dataTable.Columns[col].ColumnName == "Fecha elaboraci�n")
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
            MessageBox.Show("Archivo exportado correctamente con el formato de referencia.", "�xito", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }
}