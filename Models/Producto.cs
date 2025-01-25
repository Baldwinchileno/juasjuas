namespace AdminSERMAC.Models
{
    public class Producto
    {
        public string? Codigo { get; set; }
        public string? Nombre { get; set; }
        public string? Marca { get; set; }
        public string? Categoria { get; set; }
        public string? SubCategoria { get; set; }
        public string? UnidadMedida { get; set; }
        public double Precio { get; set; }
        public int Unidades { get; set; }
        public double Kilos { get; set; }
        public string? FechaMasAntigua { get; set; }
        public string? FechaMasNueva { get; set; }
    }
}   

