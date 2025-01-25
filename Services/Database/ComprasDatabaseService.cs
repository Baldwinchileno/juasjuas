using System;
using System.Data.SQLite;
using System.Collections.Generic;
using System.Threading.Tasks;
using Microsoft.Extensions.Logging;
using AdminSERMAC.Models;
using AdminSERMAC.Core.Interfaces;

namespace AdminSERMAC.Services.Database
{
    public class ComprasDatabaseService : BaseSQLiteService, IComprasDatabaseService
    {
        private readonly ILogger<ComprasDatabaseService> _logger;

        public ComprasDatabaseService(ILogger<ComprasDatabaseService> logger, string connectionString)
            : base(logger, connectionString)
        {
            _logger = logger;
            EnsureTableExists();
        }

        private void EnsureTableExists()
        {
            const string createTablesSql = @"
        CREATE TABLE IF NOT EXISTS Compras (
            NumeroCompra INTEGER PRIMARY KEY AUTOINCREMENT,
            Proveedor_Id INTEGER NOT NULL,
            FechaCompra TEXT NOT NULL,
            Total REAL NOT NULL DEFAULT 0,
            Observaciones TEXT,
            Estado TEXT NOT NULL DEFAULT 'Pendiente',
            FOREIGN KEY (Proveedor_Id) REFERENCES Proveedores(Id)
        );

        CREATE TABLE IF NOT EXISTS DetallesCompra (
            Id INTEGER PRIMARY KEY AUTOINCREMENT,
            NumeroCompra INTEGER NOT NULL,
            Codigo_Producto TEXT NOT NULL,
            Cantidad INTEGER NOT NULL,
            PrecioUnitario REAL NOT NULL,
            Subtotal REAL NOT NULL,
            FechaVencimiento TEXT,
            FOREIGN KEY (NumeroCompra) REFERENCES Compras(NumeroCompra),
            FOREIGN KEY (Codigo_Producto) REFERENCES Productos(Codigo)
        );";

            ExecuteInTransactionAsync(async (connection, transaction) =>
            {
                using var command = new SQLiteCommand(createTablesSql, connection, transaction);
                await command.ExecuteNonQueryAsync();
            }).Wait();
        }

        public async Task<double?> GetUltimoPrecioCompraAsync(string codigoProducto)
        {
            const string sql = @"
        SELECT PrecioUnitario 
        FROM DetallesCompra 
        WHERE Codigo_Producto = @codigoProducto 
        ORDER BY NumeroCompra DESC 
        LIMIT 1";

            using var connection = new SQLiteConnection(_connectionString);
            await connection.OpenAsync();
            using var command = new SQLiteCommand(sql, connection);
            command.Parameters.AddWithValue("@codigoProducto", codigoProducto);

            var result = await command.ExecuteScalarAsync();
            return result != DBNull.Value ? Convert.ToDouble(result) : null;
        }




        public async Task<int> CreateCompra(Compra compra)
        {
            return await ExecuteInTransactionAsync(async (connection, transaction) =>
            {
                const string sql = @"
                    INSERT INTO Compras (Proveedor_Id, FechaCompra, Total, Observaciones, Estado)
                    VALUES (@ProveedorId, @FechaCompra, @Total, @Observaciones, @Estado);
                    SELECT last_insert_rowid();";

                using var command = new SQLiteCommand(sql, connection, transaction);
                command.Parameters.AddWithValue("@ProveedorId", compra.ProveedorId);
                command.Parameters.AddWithValue("@FechaCompra", compra.FechaCompra);
                command.Parameters.AddWithValue("@Total", compra.Total);
                command.Parameters.AddWithValue("@Observaciones", compra.Observaciones ?? (object)DBNull.Value);
                command.Parameters.AddWithValue("@Estado", compra.Estado);

                return Convert.ToInt32(await command.ExecuteScalarAsync());
            });
        }

        public async Task<bool> AddDetalleCompra(DetalleCompra detalle)
        {
            return await ExecuteInTransactionAsync(async (connection, transaction) =>
            {
                const string sql = @"
            INSERT INTO DetallesCompra (
                NumeroCompra, Codigo_Producto, Cantidad, 
                PrecioUnitario, Subtotal, FechaVencimiento, Kilos
            ) VALUES (
                @NumeroCompra, @CodigoProducto, @Cantidad,
                @PrecioUnitario, @Subtotal, @FechaVencimiento, @Kilos
            )";

                using var command = new SQLiteCommand(sql, connection, transaction);
                command.Parameters.AddWithValue("@NumeroCompra", detalle.NumeroCompra);
                command.Parameters.AddWithValue("@CodigoProducto", detalle.CodigoProducto);
                command.Parameters.AddWithValue("@Cantidad", detalle.Cantidad);
                command.Parameters.AddWithValue("@PrecioUnitario", detalle.PrecioUnitario);
                command.Parameters.AddWithValue("@Subtotal", detalle.Subtotal);
                command.Parameters.AddWithValue("@FechaVencimiento", detalle.FechaVencimiento ?? (object)DBNull.Value);
                command.Parameters.AddWithValue("@Kilos", detalle.Kilos);

                return await command.ExecuteNonQueryAsync() > 0;
            });
        }

        public async Task<Compra> GetCompraById(int numeroCompra)
        {
            return await ExecuteInTransactionAsync(async (connection, transaction) =>
            {
                const string sql = @"
                    SELECT c.*, p.Nombre as ProveedorNombre 
                    FROM Compras c
                    INNER JOIN Proveedores p ON c.Proveedor_Id = p.Id
                    WHERE c.NumeroCompra = @NumeroCompra";

                using var command = new SQLiteCommand(sql, connection, transaction);
                command.Parameters.AddWithValue("@NumeroCompra", numeroCompra);

                using var reader = await command.ExecuteReaderAsync();
                if (await reader.ReadAsync())
                {
                    return new Compra
                    {
                        NumeroCompra = reader.GetInt32(reader.GetOrdinal("NumeroCompra")),
                        ProveedorId = reader.GetInt32(reader.GetOrdinal("Proveedor_Id")),
                        FechaCompra = reader.GetString(reader.GetOrdinal("FechaCompra")),
                        Total = reader.GetDecimal(reader.GetOrdinal("Total")),
                        Observaciones = reader.IsDBNull(reader.GetOrdinal("Observaciones")) ? null : reader.GetString(reader.GetOrdinal("Observaciones")),
                        Estado = reader.GetString(reader.GetOrdinal("Estado"))
                    };
                }

                return null;
            });
        }

        public async Task<List<Compra>> GetAllCompras()
        {
            return await ExecuteInTransactionAsync(async (connection, transaction) =>
            {
                var compras = new List<Compra>();
                const string sql = @"
                    SELECT c.*, p.Nombre as ProveedorNombre 
                    FROM Compras c
                    INNER JOIN Proveedores p ON c.Proveedor_Id = p.Id
                    ORDER BY c.NumeroCompra DESC";

                using var command = new SQLiteCommand(sql, connection, transaction);
                using var reader = await command.ExecuteReaderAsync();

                while (await reader.ReadAsync())
                {
                    compras.Add(new Compra
                    {
                        NumeroCompra = reader.GetInt32(reader.GetOrdinal("NumeroCompra")),
                        ProveedorId = reader.GetInt32(reader.GetOrdinal("Proveedor_Id")),
                        FechaCompra = reader.GetString(reader.GetOrdinal("FechaCompra")),
                        Total = reader.GetDecimal(reader.GetOrdinal("Total")),
                        Observaciones = reader.IsDBNull(reader.GetOrdinal("Observaciones")) ? null : reader.GetString(reader.GetOrdinal("Observaciones")),
                        Estado = reader.GetString(reader.GetOrdinal("Estado"))
                    });
                }

                return compras;
            });
        }

        public async Task<List<DetalleCompra>> GetDetallesByCompra(int numeroCompra)
        {
            return await ExecuteInTransactionAsync(async (connection, transaction) =>
            {
                var detalles = new List<DetalleCompra>();
                const string sql = @"
                    SELECT * FROM DetallesCompra 
                    WHERE NumeroCompra = @NumeroCompra";

                using var command = new SQLiteCommand(sql, connection, transaction);
                command.Parameters.AddWithValue("@NumeroCompra", numeroCompra);

                using var reader = await command.ExecuteReaderAsync();
                while (await reader.ReadAsync())
                {
                    detalles.Add(new DetalleCompra
                    {
                        Id = reader.GetInt32(reader.GetOrdinal("Id")),
                        NumeroCompra = reader.GetInt32(reader.GetOrdinal("NumeroCompra")),
                        CodigoProducto = reader.GetString(reader.GetOrdinal("Codigo_Producto")),
                        Cantidad = reader.GetInt32(reader.GetOrdinal("Cantidad")),
                        PrecioUnitario = reader.GetDecimal(reader.GetOrdinal("PrecioUnitario")),
                        Subtotal = reader.GetDecimal(reader.GetOrdinal("Subtotal")),
                        FechaVencimiento = reader.IsDBNull(reader.GetOrdinal("FechaVencimiento")) ? null : reader.GetString(reader.GetOrdinal("FechaVencimiento"))
                    });
                }

                return detalles;
            });
        }

        public async Task<bool> UpdateCompraEstado(int numeroCompra, string estado)
        {
            return await ExecuteInTransactionAsync(async (connection, transaction) =>
            {
                const string sql = "UPDATE Compras SET Estado = @Estado WHERE NumeroCompra = @NumeroCompra";

                using var command = new SQLiteCommand(sql, connection, transaction);
                command.Parameters.AddWithValue("@NumeroCompra", numeroCompra);
                command.Parameters.AddWithValue("@Estado", estado);

                return await command.ExecuteNonQueryAsync() > 0;
            });
        }

        public async Task<int> GetNextNumeroCompra()
        {
            return await ExecuteInTransactionAsync(async (connection, transaction) =>
            {
                const string sql = "SELECT COALESCE(MAX(NumeroCompra), 0) + 1 FROM Compras";

                using var command = new SQLiteCommand(sql, connection, transaction);
                var result = await command.ExecuteScalarAsync();
                return Convert.ToInt32(result);
            });
        }
    }
}
