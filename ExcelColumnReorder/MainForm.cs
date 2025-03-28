using System.ComponentModel;
using System.Data;
using System.Windows.Forms;
using System;
using System.Data;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using OfficeOpenXml;
using LicenseContext = OfficeOpenXml.LicenseContext;

namespace ExcelColumnReorder
{
    public partial class MainForm : Form
    {
        private DataTable _dataTable;
        private string _filePath;
        private readonly string[] _orderedColumns = { "Nombre", "Edad", "Ciudad" }; // Cambia según necesidad
        private readonly string[] _columnsToRemove = { "Dirección", "Teléfono" }; // Columnas a eliminar
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

                for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
                {
                    _dataTable.Columns.Add(worksheet.Cells[1, col].Text);
                }

                for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
                {
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
                .Where(col => !_columnsToRemove.Contains(col.ColumnName))
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
                var worksheet = package.Workbook.Worksheets.Add("Sheet1");

                for (int col = 0; col < _dataTable.Columns.Count; col++)
                {
                    worksheet.Cells[1, col + 1].Value = _dataTable.Columns[col].ColumnName;
                }

                for (int row = 0; row < _dataTable.Rows.Count; row++)
                {
                    for (int col = 0; col < _dataTable.Columns.Count; col++)
                    {
                        worksheet.Cells[row + 2, col + 1].Value = _dataTable.Rows[row][col];
                    }
                }
                package.SaveAs(new FileInfo(filePath));
            }
            MessageBox.Show("Archivo exportado correctamente.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void button1_Click(object sender, EventArgs e)
        {

        }
    }
}
