using System.Drawing.Printing;
using System.Windows.Forms;

namespace AdminSERMAC.Services
{
    public class PrintService
    {
        private readonly Font titleFont = new Font("Arial", 14, FontStyle.Bold);
        private readonly Font subtitleFont = new Font("Arial", 12, FontStyle.Bold);
        private readonly Font normalFont = new Font("Arial", 10);
        private readonly int leftMargin = 50;
        private readonly DataGridView ventasGrid;
        private readonly Venta venta;
        private readonly AdminSERMAC.Models.Cliente cliente;
        private int currentPage = 1;
        private int rowsPerPage = 25;

        public PrintService(DataGridView ventasGrid, Venta venta, AdminSERMAC.Models.Cliente cliente)
        {
            this.ventasGrid = ventasGrid;
            this.venta = venta;
            this.cliente = cliente;
        }

        public void PrintGuia()
        {
            PrintDocument printDocument = new PrintDocument();
            printDocument.PrintPage += PrintPage;

            PrintPreviewDialog preview = new PrintPreviewDialog();
            preview.Document = printDocument;
            preview.ShowDialog();
        }

        private void PrintPage(object sender, PrintPageEventArgs e)
        {
            float yPos = 50;
            float rightMargin = e.MarginBounds.Right - 50;

            // Título
            e.Graphics.DrawString("GUÍA DE VENTA", titleFont, Brushes.Black, (e.PageBounds.Width - e.Graphics.MeasureString("GUÍA DE VENTA", titleFont).Width) / 2, yPos);
            yPos += 40;

            // Información del encabezado
            e.Graphics.DrawString($"N° Guía: {venta.NumeroGuia}", subtitleFont, Brushes.Black, leftMargin, yPos);
            e.Graphics.DrawString($"Fecha: {venta.FechaVenta}", subtitleFont, Brushes.Black, rightMargin - 200, yPos);
            yPos += 30;

            // Información del cliente
            e.Graphics.DrawString($"Cliente: {cliente.Nombre}", normalFont, Brushes.Black, leftMargin, yPos);
            yPos += 20;
            e.Graphics.DrawString($"RUT: {cliente.RUT}", normalFont, Brushes.Black, leftMargin, yPos);
            yPos += 20;
            e.Graphics.DrawString($"Dirección: {cliente.Direccion}", normalFont, Brushes.Black, leftMargin, yPos);
            yPos += 40;

            // Encabezados de la tabla
            string[] headers = { "Código", "Descripción", "Band.", "K.Bruto", "K.Neto", "Precio", "Total" };
            float[] columnWidths = { 80, 200, 60, 80, 80, 80, 100 };
            float xPos = leftMargin;

            for (int i = 0; i < headers.Length; i++)
            {
                e.Graphics.DrawString(headers[i], subtitleFont, Brushes.Black, xPos, yPos);
                xPos += columnWidths[i];
            }
            yPos += 30;

            // Línea separadora
            e.Graphics.DrawLine(Pens.Black, leftMargin, yPos - 5, rightMargin, yPos - 5);

            // Datos de los productos
            float startY = yPos;
            double totalVenta = 0;

            for (int i = 0; i < ventasGrid.Rows.Count - 1; i++)
            {
                var row = ventasGrid.Rows[i];
                if (row.IsNewRow) continue;

                xPos = leftMargin;
                e.Graphics.DrawString(row.Cells["Codigo"].Value?.ToString(), normalFont, Brushes.Black, xPos, yPos);
                xPos += columnWidths[0];

                e.Graphics.DrawString(row.Cells["Descripcion"].Value?.ToString(), normalFont, Brushes.Black, xPos, yPos);
                xPos += columnWidths[1];

                e.Graphics.DrawString(row.Cells["Bandejas"].Value?.ToString(), normalFont, Brushes.Black, xPos, yPos);
                xPos += columnWidths[2];

                e.Graphics.DrawString(Convert.ToDouble(row.Cells["KilosBruto"].Value).ToString("N2"), normalFont, Brushes.Black, xPos, yPos);
                xPos += columnWidths[3];

                e.Graphics.DrawString(Convert.ToDouble(row.Cells["KilosNeto"].Value).ToString("N2"), normalFont, Brushes.Black, xPos, yPos);
                xPos += columnWidths[4];

                e.Graphics.DrawString(Convert.ToDouble(row.Cells["Precio"].Value).ToString("N2"), normalFont, Brushes.Black, xPos, yPos);
                xPos += columnWidths[5];

                double total = Convert.ToDouble(row.Cells["Total"].Value);
                e.Graphics.DrawString(total.ToString("N2"), normalFont, Brushes.Black, xPos, yPos);

                totalVenta += total;
                yPos += 20;

                // Verificar si necesitamos una nueva página
                if (yPos > e.MarginBounds.Bottom - 50)
                {
                    e.HasMorePages = true;
                    return;
                }
            }

            // Total final
            yPos += 20;
            e.Graphics.DrawLine(Pens.Black, leftMargin, yPos - 5, rightMargin, yPos - 5);
            e.Graphics.DrawString("Total:", subtitleFont, Brushes.Black, rightMargin - 200, yPos);
            e.Graphics.DrawString(totalVenta.ToString("N2"), subtitleFont, Brushes.Black, rightMargin - 100, yPos);

            e.HasMorePages = false;
        }
    }
}