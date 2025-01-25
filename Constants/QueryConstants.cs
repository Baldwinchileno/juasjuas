namespace AdminSERMAC.Constants
{
    public static class QueryConstants
    {
        public static class Cliente
        {
            public const string SELECT_ALL = @"
                SELECT * FROM Clientes";

            public const string SELECT_BY_RUT = @"
                SELECT * FROM Clientes 
                WHERE RUT = @RUT";

            public const string INSERT = @"
                INSERT INTO Clientes (
                    RUT, 
                    Nombre, 
                    Direccion, 
                    Giro, 
                    Deuda
                ) VALUES (
                    @RUT, 
                    @Nombre, 
                    @Direccion, 
                    @Giro, 
                    @Deuda
                )";

            public const string UPDATE = @"
                UPDATE Clientes 
                SET Nombre = @Nombre,
                    Direccion = @Direccion,
                    Giro = @Giro,
                    Deuda = @Deuda
                WHERE RUT = @RUT";

            public const string DELETE = @"
                DELETE FROM Clientes 
                WHERE RUT = @RUT";

            public const string UPDATE_DEUDA = @"
                UPDATE Clientes 
                SET Deuda = Deuda + @Monto 
                WHERE RUT = @RUT";

            public const string GET_CURRENT_DEUDA = @"
                SELECT Deuda 
                FROM Clientes 
                WHERE RUT = @RUT";
        }

        public static class Ventas
        {
            public const string SELECT_VENTAS = @"
                SELECT DISTINCT
                    v.NumeroGuia,
                    v.FechaVenta,
                    v.Descripcion,
                    v.KilosNeto,
                    SUM(v.Total) OVER (PARTITION BY v.NumeroGuia) as Total,
                    v.PagadoConCredito,
                    v.RUT,
                    c.Nombre as ClienteNombre
                FROM Ventas v
                JOIN Clientes c ON v.RUT = c.RUT
                WHERE v.RUT = @RUT
                ORDER BY v.NumeroGuia DESC";

            public const string SELECT_DETALLES_GUIA = @"
                SELECT 
                    v.NumeroGuia,
                    v.CodigoProducto,
                    v.Descripcion,
                    v.Bandejas,
                    v.KilosNeto,
                    v.Total,
                    v.FechaVenta,
                    v.PagadoConCredito,
                    c.Nombre as ClienteNombre,
                    c.RUT
                FROM Ventas v
                JOIN Clientes c ON v.RUT = c.RUT
                WHERE v.NumeroGuia = @NumeroGuia
                ORDER BY v.CodigoProducto";
        }

        public static class General
        {
            public const string CALCULAR_DEUDA_TOTAL = @"
                SELECT SUM(Deuda) 
                FROM Clientes 
                WHERE RUT = @RUT";
        }
    }
}
