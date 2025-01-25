using ClosedXML.Excel;
using AdminSERMAC.Models;
using System.Data;

namespace AdminSERMAC.Services
{
    public class ExcelExportService
    {
        public void ExportarGuia(string numeroGuia, DataGridView ventasGrid, Cliente cliente, DateTime fechaEmision, string rutaArchivo)
        {
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Guía de Venta");

                // Estilo para títulos
                var tituloStyle = worksheet.Style;
                tituloStyle.Font.Bold = true;
                tituloStyle.Font.FontSize = 14;
                tituloStyle.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                // Encabezado
                worksheet.Cell("A1").Value = "GUÍA DE VENTA";
                worksheet.Range("A1:H1").Merge().Style = tituloStyle;

                // Información del cliente
                worksheet.Cell("A3").Value = "Número de Guía:";
                worksheet.Cell("B3").Value = numeroGuia;
                worksheet.Cell("A4").Value = "Fecha:";
                worksheet.Cell("B4").Value = fechaEmision.ToString("dd/MM/yyyy");
                worksheet.Cell("A5").Value = "Cliente:";
                worksheet.Cell("B5").Value = cliente.Nombre;
                worksheet.Cell("A6").Value = "RUT:";
                worksheet.Cell("B6").Value = cliente.RUT;
                worksheet.Cell("A7").Value = "Dirección:";
                worksheet.Cell("B7").Value = cliente.Direccion;

                // Encabezados de la tabla
                int rowIndex = 9;
                string[] headers = { "Código", "Descripción", "Unidades", "Bandejas", "Kilos Bruto", "Kilos Neto", "Precio", "Total" };
                for (int i = 0; i < headers.Length; i++)
                {
                    worksheet.Cell(rowIndex, i + 1).Value = headers[i];
                    worksheet.Cell(rowIndex, i + 1).Style.Font.Bold = true;
                }

                // Datos de la venta
                rowIndex++;
                double totalVenta = 0;

                foreach (DataGridViewRow row in ventasGrid.Rows)
                {
                    if (row.IsNewRow || row.Cells["Codigo"].Value == null) continue;

                    worksheet.Cell(rowIndex, 1).Value = row.Cells["Codigo"].Value?.ToString();
                    worksheet.Cell(rowIndex, 2).Value = row.Cells["Descripcion"].Value?.ToString();
                    worksheet.Cell(rowIndex, 3).Value = row.Cells["Unidades"].Value?.ToString();
                    worksheet.Cell(rowIndex, 4).Value = row.Cells["Bandejas"].Value?.ToString();
                    worksheet.Cell(rowIndex, 5).Value = Convert.ToDouble(row.Cells["KilosBruto"].Value);
                    worksheet.Cell(rowIndex, 6).Value = Convert.ToDouble(row.Cells["KilosNeto"].Value);
                    worksheet.Cell(rowIndex, 7).Value = Convert.ToDouble(row.Cells["Precio"].Value);
                    worksheet.Cell(rowIndex, 8).Value = Convert.ToDouble(row.Cells["Total"].Value);

                    totalVenta += Convert.ToDouble(row.Cells["Total"].Value);
                    rowIndex++;
                }

                // Total
                worksheet.Cell(rowIndex + 1, 7).Value = "Total:";
                worksheet.Cell(rowIndex + 1, 8).Value = totalVenta;
                worksheet.Cell(rowIndex + 1, 8).Style.NumberFormat.Format = "#,##0";

                // Formato de columnas
                worksheet.Columns().AdjustToContents();
                worksheet.Columns("E:H").Style.NumberFormat.Format = "#,##0.00";

                workbook.SaveAs(rutaArchivo);
            }
        }
    }
}
