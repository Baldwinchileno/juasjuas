using System;
using System.Collections.Generic;

namespace AdminSERMAC.Models
{
    public class Compra
    {
        public int NumeroCompra { get; set; }
        public int ProveedorId { get; set; }
        public string FechaCompra { get; set; }
        public decimal Total { get; set; }
        public string Observaciones { get; set; }
        public string Estado { get; set; }
        public List<DetalleCompra> Detalles { get; set; } = new List<DetalleCompra>();
    }

    public class DetalleCompra
    {
        public int Id { get; set; }
        public int NumeroCompra { get; set; }
        public string CodigoProducto { get; set; }
        public int Cantidad { get; set; }
        public decimal PrecioUnitario { get; set; }
        public decimal Subtotal { get; set; }
        public string FechaVencimiento { get; set; }
        public double Kilos { get; set; }
    }
}
