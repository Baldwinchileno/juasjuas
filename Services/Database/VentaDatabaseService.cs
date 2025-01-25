using System.Data;
using System.Data.SQLite;
using Microsoft.Extensions.Logging;
using AdminSERMAC.Models;

namespace AdminSERMAC.Services.Database
{
    public interface IVentaDatabaseService
    {
        Task<int> GetUltimoNumeroGuiaAsync();
        Task<bool> IncrementarNumeroGuiaAsync();
        Task<bool> RegistrarVentaAsync(Venta venta);
        Task<IEnumerable<Venta>> GetVentasPorClienteAsync(string rut);
        Task<IEnumerable<Venta>> GetVentasPorFechaAsync(DateTime inicio, DateTime fin);
        Task<bool> MarcarVentaComoPagadaAsync(int numeroGuia, string codigoProducto);
        Task<bool> ActualizarVentaAsync(Venta venta);
        Task<Venta?> GetVentaPorGuiaYProductoAsync(int numeroGuia, string codigoProducto);
    }

    public class VentaDatabaseService : BaseSQLiteService, IVentaDatabaseService
    {
        private const string TableName = "Ventas";
        private readonly IClienteDatabaseService _clienteService;
        private readonly IInventarioDatabaseService _inventarioService;

        public VentaDatabaseService(
            ILogger<VentaDatabaseService> logger,
            string connectionString,
            IClienteDatabaseService clienteService,
            IInventarioDatabaseService inventarioService)
            : base(logger, connectionString)
        {
            _clienteService = clienteService ?? throw new ArgumentNullException(nameof(clienteService));
            _inventarioService = inventarioService ?? throw new ArgumentNullException(nameof(inventarioService));
            EnsureTableExists();
        }

        private void EnsureTableExists()
        {
            ExecuteInTransactionAsync(async (connection, transaction) =>
            {
                // Primero verificamos si la columna existe
                const string checkColumnSql = @"
            SELECT COUNT(*) 
            FROM pragma_table_info('Ventas') 
            WHERE name='Precio'";

                using (var checkCommand = new SQLiteCommand(checkColumnSql, connection, transaction))
                {
                    var result = await checkCommand.ExecuteScalarAsync();
                    if (Convert.ToInt32(result) == 0)
                    {
                        // Si la columna no existe, la agregamos
                        const string alterTableSql = "ALTER TABLE Ventas ADD COLUMN Precio REAL DEFAULT 0";
                        using var alterCommand = new SQLiteCommand(alterTableSql, connection, transaction);
                        await alterCommand.ExecuteNonQueryAsync();
                    }
                }
            }).Wait();
        }

        public async Task<double?> GetUltimoPrecioVentaAsync(string codigoProducto, string rut)
        {
            return await ExecuteInTransactionAsync(async (connection, transaction) =>
            {
                const string sql = @"
            SELECT Precio 
            FROM Ventas 
            WHERE CodigoProducto = @codigoProducto 
            AND RUT = @rut 
            AND Precio > 0
            ORDER BY NumeroGuia DESC 
            LIMIT 1";

                using var command = new SQLiteCommand(sql, connection, transaction);
                command.Parameters.AddWithValue("@codigoProducto", codigoProducto);
                command.Parameters.AddWithValue("@rut", rut);

                var result = await command.ExecuteScalarAsync();
                object value = result;
                return value == DBNull.Value ? (double?)null : Convert.ToDouble(value);
            });
        }

        public async Task<int> GetUltimoNumeroGuiaAsync()
        {
            return await ExecuteInTransactionAsync(async (connection, transaction) =>
            {
                const string sql = "SELECT Valor FROM Configuracion WHERE Clave = 'UltimoNumeroGuia'";
                using var command = new SQLiteCommand(sql, connection, transaction);

                var result = await command.ExecuteScalarAsync();
                return result != null ? Convert.ToInt32(result) : 0;
            });
        }

        public async Task<bool> IncrementarNumeroGuiaAsync()
        {
            return await ExecuteInTransactionAsync(async (connection, transaction) =>
            {
                const string sql = @"
                    UPDATE Configuracion
                    SET Valor = CAST((CAST(Valor AS INTEGER) + 1) AS TEXT)
                    WHERE Clave = 'UltimoNumeroGuia'";

                using var command = new SQLiteCommand(sql, connection, transaction);
                return await command.ExecuteNonQueryAsync() > 0;
            });
        }

        public async Task<bool> RegistrarVentaAsync(Venta venta)
        {
            return await ExecuteInTransactionAsync(async (connection, transaction) =>
            {
                // Validar si hay suficiente stock
                var inventarioActual = await _inventarioService.GetInventarioPorCodigoAsync(venta.CodigoProducto);
                if (inventarioActual.Rows.Count == 0 ||
                    Convert.ToInt32(inventarioActual.Rows[0]["Unidades"]) < venta.Bandejas ||
                    Convert.ToDouble(inventarioActual.Rows[0]["Kilos"]) < venta.KilosNeto)
                {
                    throw new InvalidOperationException("No hay suficiente stock para realizar la venta");
                }

                // Registrar la venta
                const string sqlVenta = @"
    INSERT INTO Ventas (
        NumeroGuia, 
        CodigoProducto, 
        Descripcion, 
        Bandejas, 
        KilosNeto, 
        FechaVenta, 
        PagadoConCredito, 
        RUT,
        Total,
        Precio  -- Agregamos el campo Precio
    ) VALUES (
        @numeroGuia,
        @codigoProducto,
        @descripcion,
        @bandejas,
        @kilosNeto,
        @fechaVenta,
        @pagadoConCredito,
        @rut,
        @total,
        @precio  -- Agregamos el parámetro
    )";

                using var command = new SQLiteCommand(sqlVenta, connection, transaction);
                command.Parameters.AddWithValue("@numeroGuia", venta.NumeroGuia);
                command.Parameters.AddWithValue("@codigoProducto", venta.CodigoProducto);
                command.Parameters.AddWithValue("@descripcion", venta.Descripcion);
                command.Parameters.AddWithValue("@bandejas", venta.Bandejas);
                command.Parameters.AddWithValue("@kilosNeto", venta.KilosNeto);
                command.Parameters.AddWithValue("@fechaVenta", venta.FechaVenta);
                command.Parameters.AddWithValue("@pagadoConCredito", venta.PagadoConCredito);
                command.Parameters.AddWithValue("@rut", venta.RUT);
                command.Parameters.AddWithValue("@total", venta.Total);
                command.Parameters.AddWithValue("@precio", venta.Total / venta.KilosNeto);

                var result = await command.ExecuteNonQueryAsync() > 0;

                // Actualizar el inventario
                if (result)
                {
                    await _inventarioService.ActualizarInventarioAsync(
                        venta.CodigoProducto,
                        venta.Bandejas,
                        venta.KilosNeto);

                    // Si es venta a crédito, actualizar la deuda del cliente
                    if (venta.PagadoConCredito == 1)
                    {
                        await _clienteService.UpdateDeudaAsync(venta.RUT, venta.Total);
                    }
                }

                return result;
            });
        }


        public async Task<bool> ActualizarVentaAsync(Venta venta)
        {
            return await ExecuteInTransactionAsync(async (connection, transaction) =>
            {
                const string sql = @"
        UPDATE Ventas 
        SET Descripcion = @descripcion,
            Bandejas = @bandejas,
            KilosNeto = @kilosNeto,
            FechaVenta = @fechaVenta,
            PagadoConCredito = @pagadoConCredito,
            RUT = @rut
        WHERE NumeroGuia = @numeroGuia 
        AND CodigoProducto = @codigoProducto";

                using var command = new SQLiteCommand(sql, connection, transaction);
                command.Parameters.AddWithValue("@numeroGuia", venta.NumeroGuia);
                command.Parameters.AddWithValue("@codigoProducto", venta.CodigoProducto);
                command.Parameters.AddWithValue("@descripcion", venta.Descripcion);
                command.Parameters.AddWithValue("@bandejas", venta.Bandejas);
                command.Parameters.AddWithValue("@kilosNeto", venta.KilosNeto);
                command.Parameters.AddWithValue("@fechaVenta", venta.FechaVenta);
                command.Parameters.AddWithValue("@pagadoConCredito", venta.PagadoConCredito);
                command.Parameters.AddWithValue("@rut", venta.RUT);

                return await command.ExecuteNonQueryAsync() > 0;
            });
        }


        public async Task<IEnumerable<Venta>> GetVentasPorClienteAsync(string rut)
        {
            return await ExecuteInTransactionAsync(async (connection, transaction) =>
            {
                var ventas = new List<Venta>();
                const string sql = @"
                    SELECT v.*, p.Precio * v.KilosNeto as Total, c.Deuda
                    FROM Ventas v
                    JOIN Clientes c ON c.RUT = @RUT
                    JOIN Productos p ON p.Codigo = v.CodigoProducto
                    WHERE c.RUT = @RUT AND v.PagadoConCredito = 1";

                using var command = new SQLiteCommand(sql, connection, transaction);
                command.Parameters.AddWithValue("@RUT", rut);

                using var reader = await command.ExecuteReaderAsync();
                while (await reader.ReadAsync())
                {
                    ventas.Add(new Venta
                    {
                        NumeroGuia = Convert.ToInt32(reader["NumeroGuia"]),
                        CodigoProducto = reader["CodigoProducto"].ToString(),
                        Descripcion = reader["Descripcion"].ToString(),
                        Bandejas = Convert.ToInt32(reader["Bandejas"]),
                        KilosNeto = Convert.ToDouble(reader["KilosNeto"]),
                        FechaVenta = reader["FechaVenta"].ToString(),
                        Total = Convert.ToDouble(reader["Total"]),
                        Deuda = Convert.ToDouble(reader["Deuda"]),
                        PagadoConCredito = Convert.ToInt32(reader["PagadoConCredito"]),
                        RUT = reader["RUT"].ToString()
                    });
                }
                return ventas;
            });
        }

        public async Task<IEnumerable<Venta>> GetVentasPorFechaAsync(DateTime inicio, DateTime fin)
        {
            return await ExecuteInTransactionAsync(async (connection, transaction) =>
            {
                var ventas = new List<Venta>();
                const string sql = @"
            SELECT v.*, p.Precio * v.KilosNeto as Total
            FROM Ventas v
            JOIN Productos p ON p.Codigo = v.CodigoProducto
            WHERE date(v.FechaVenta) BETWEEN date(@inicio) AND date(@fin)";

                using var command = new SQLiteCommand(sql, connection, transaction);
                command.Parameters.AddWithValue("@inicio", inicio.ToString("yyyy-MM-dd"));
                command.Parameters.AddWithValue("@fin", fin.ToString("yyyy-MM-dd"));

                using var reader = await command.ExecuteReaderAsync();
                while (await reader.ReadAsync())
                {
                    ventas.Add(new Venta
                    {
                        NumeroGuia = Convert.ToInt32(reader["NumeroGuia"]),
                        CodigoProducto = reader["CodigoProducto"].ToString(),
                        Descripcion = reader["Descripcion"].ToString(),
                        Bandejas = Convert.ToInt32(reader["Bandejas"]),
                        KilosNeto = Convert.ToDouble(reader["KilosNeto"]),
                        FechaVenta = reader["FechaVenta"].ToString(),
                        Total = Convert.ToDouble(reader["Total"]),
                        PagadoConCredito = Convert.ToInt32(reader["PagadoConCredito"]), // Fixed conversion
                        RUT = reader["RUT"].ToString()
                    });
                }
                return ventas;
            });
        }


        public async Task<bool> MarcarVentaComoPagadaAsync(int numeroGuia, string codigoProducto)
        {
            return await ExecuteInTransactionAsync(async (connection, transaction) =>
            {
                // Obtener la venta actual
                var venta = await GetVentaPorGuiaYProductoAsync(numeroGuia, codigoProducto);
                if (venta == null || Convert.ToInt32(venta.PagadoConCredito) == 0)
                {
                    return false;
                }

                // Actualizar el estado de la venta
                const string sqlVenta = @"
                    UPDATE Ventas 
                    SET PagadoConCredito = 0
                    WHERE NumeroGuia = @numeroGuia AND CodigoProducto = @codigoProducto";

                using var command = new SQLiteCommand(sqlVenta, connection, transaction);
                command.Parameters.AddWithValue("@numeroGuia", numeroGuia);
                command.Parameters.AddWithValue("@codigoProducto", codigoProducto);

                var result = await command.ExecuteNonQueryAsync() > 0;

                // Actualizar la deuda del cliente
                if (result)
                {
                    await _clienteService.UpdateDeudaAsync(venta.RUT, -venta.Total);
                }

                return result;
            });
        }

        public async Task<Venta?> GetVentaPorGuiaYProductoAsync(int numeroGuia, string codigoProducto)
        {
            return await ExecuteInTransactionAsync(async (connection, transaction) =>
            {
                const string sql = @"
                    SELECT v.*, p.Precio * v.KilosNeto as Total
                    FROM Ventas v
                    JOIN Productos p ON p.Codigo = v.CodigoProducto
                    WHERE v.NumeroGuia = @numeroGuia AND v.CodigoProducto = @codigoProducto";

                using var command = new SQLiteCommand(sql, connection, transaction);
                command.Parameters.AddWithValue("@numeroGuia", numeroGuia);
                command.Parameters.AddWithValue("@codigoProducto", codigoProducto);

                using var reader = await command.ExecuteReaderAsync();
                if (await reader.ReadAsync())
                {
                    return new Venta
                    {
                        NumeroGuia = Convert.ToInt32(reader["NumeroGuia"]),
                        CodigoProducto = reader["CodigoProducto"].ToString(),
                        Descripcion = reader["Descripcion"].ToString(),
                        Bandejas = Convert.ToInt32(reader["Bandejas"]),
                        KilosNeto = Convert.ToDouble(reader["KilosNeto"]),
                        FechaVenta = reader["FechaVenta"].ToString(),
                        Total = Convert.ToDouble(reader["Total"]),
                        PagadoConCredito = Convert.ToInt32(reader["PagadoConCredito"]), // Fixed conversion
                        RUT = reader["RUT"].ToString()
                    };

                }
                return null;
            });
        }
    }
}