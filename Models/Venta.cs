namespace AdminSERMAC.Models
{
    public class Venta
    {
        public string NumeroGuia { get; set; }
        public string? CodigoProducto { get; set; }
        public string? Descripcion { get; set; }
        public int Bandejas { get; set; }
        public double KilosNeto { get; set; }
        public DateTime FechaVenta { get; set; }
        public double Total { get; set; }
        public double Deuda { get; set; }
        public int PagadoConCredito { get; set; }
        public string? RUT { get; set; }
        public string? ClienteNombre { get; set; }
    }
}
