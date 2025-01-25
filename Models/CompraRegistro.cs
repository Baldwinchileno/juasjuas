using System;

namespace AdminSERMAC.Models
{
    public class CompraRegistro
    {
        public int Id { get; set; }
        public DateTime FechaCompra { get; set; }
        public string Proveedor { get; set; }
        public string Producto { get; set; }
        public decimal Cantidad { get; set; }
        public decimal PrecioUnitario { get; set; }
        public decimal Total { get; set; }
        public string Observaciones { get; set; }
        public bool EstaProcesado { get; set; }  // Indica si ya se agregó al inventario

        // Constructor por defecto
        public CompraRegistro()
        {
            FechaCompra = DateTime.Now;
            EstaProcesado = false;
        }

        // Constructor con parámetros principales
        public CompraRegistro(string proveedor, string producto, decimal cantidad, decimal precioUnitario)
        {
            FechaCompra = DateTime.Now;
            Proveedor = proveedor;
            Producto = producto;
            Cantidad = cantidad;
            PrecioUnitario = precioUnitario;
            Total = cantidad * precioUnitario;
            EstaProcesado = false;
        }

        // Método para calcular el total
        public void CalcularTotal()
        {
            Total = Cantidad * PrecioUnitario;
        }

        // Método para marcar como procesado
        public void MarcarComoProcesado()
        {
            EstaProcesado = true;
        }
    }
}