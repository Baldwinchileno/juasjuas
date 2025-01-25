using System;
using System.Collections.Generic;
using System.Data.SQLite;
using System.Threading.Tasks;
using AdminSERMAC.Models;
using Microsoft.Extensions.Logging;

namespace AdminSERMAC.Services.Database
{
    public interface ICompraRegistroDatabaseService
    {
        Task<CompraRegistro> GetCompraRegistroByIdAsync(int id);
        Task<List<CompraRegistro>> GetAllCompraRegistrosAsync();
        Task<List<CompraRegistro>> GetComprasNoProcesadasAsync();
        Task<int> CreateCompraRegistroAsync(CompraRegistro compraRegistro);
        Task<bool> UpdateCompraRegistroAsync(CompraRegistro compraRegistro);
        Task<bool> DeleteCompraRegistroAsync(int id);
        Task<bool> MarcarComoProcesadoAsync(int id);
    }

    public class CompraRegistroDatabaseService : SQLiteService, ICompraRegistroDatabaseService
    {
        private readonly string _connectionString;
        private readonly ILogger<SQLiteService> _logger;

        public CompraRegistroDatabaseService(ILogger<SQLiteService> logger) : base(logger)
        {
            _logger = logger;
            _connectionString = AppSettings.GetConnectionString();
            CreateTableIfNotExists();
        }

        private SQLiteConnection CreateConnection()
        {
            var connection = new SQLiteConnection(_connectionString);
            connection.Open();
            return connection;
        }

        private void CreateTableIfNotExists()
        {
            const string createTableQuery = @"
                CREATE TABLE IF NOT EXISTS CompraRegistros (
                    Id INTEGER PRIMARY KEY AUTOINCREMENT,
                    FechaCompra TEXT NOT NULL,
                    Proveedor TEXT NOT NULL,
                    Producto TEXT NOT NULL,
                    Cantidad DECIMAL NOT NULL,
                    PrecioUnitario DECIMAL NOT NULL,
                    Total DECIMAL NOT NULL,
                    Observaciones TEXT,
                    EstaProcesado INTEGER NOT NULL DEFAULT 0
                );";

            using var connection = CreateConnection();
            using var command = new SQLiteCommand(createTableQuery, connection);
            command.ExecuteNonQuery();
        }

        public async Task<CompraRegistro> GetCompraRegistroByIdAsync(int id)
        {
            const string query = "SELECT * FROM CompraRegistros WHERE Id = @Id;";

            try
            {
                using var connection = CreateConnection();
                using var command = new SQLiteCommand(query, connection);
                command.Parameters.AddWithValue("@Id", id);

                using var reader = await command.ExecuteReaderAsync();
                if (await reader.ReadAsync())
                {
                    return new CompraRegistro
                    {
                        Id = Convert.ToInt32(reader["Id"]),
                        FechaCompra = DateTime.Parse(reader["FechaCompra"].ToString()),
                        Proveedor = reader["Proveedor"].ToString(),
                        Producto = reader["Producto"].ToString(),
                        Cantidad = Convert.ToDecimal(reader["Cantidad"]),
                        PrecioUnitario = Convert.ToDecimal(reader["PrecioUnitario"]),
                        Total = Convert.ToDecimal(reader["Total"]),
                        Observaciones = reader["Observaciones"].ToString(),
                        EstaProcesado = Convert.ToBoolean(Convert.ToInt32(reader["EstaProcesado"]))
                    };
                }
                return null;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al obtener registro de compra por ID");
                throw;
            }
        }

        public async Task<List<CompraRegistro>> GetAllCompraRegistrosAsync()
        {
            const string query = "SELECT * FROM CompraRegistros ORDER BY FechaCompra DESC;";
            var registros = new List<CompraRegistro>();

            try
            {
                using var connection = CreateConnection();
                using var command = new SQLiteCommand(query, connection);
                using var reader = await command.ExecuteReaderAsync();

                while (await reader.ReadAsync())
                {
                    registros.Add(new CompraRegistro
                    {
                        Id = Convert.ToInt32(reader["Id"]),
                        FechaCompra = DateTime.Parse(reader["FechaCompra"].ToString()),
                        Proveedor = reader["Proveedor"].ToString(),
                        Producto = reader["Producto"].ToString(),
                        Cantidad = Convert.ToDecimal(reader["Cantidad"]),
                        PrecioUnitario = Convert.ToDecimal(reader["PrecioUnitario"]),
                        Total = Convert.ToDecimal(reader["Total"]),
                        Observaciones = reader["Observaciones"].ToString(),
                        EstaProcesado = Convert.ToBoolean(Convert.ToInt32(reader["EstaProcesado"]))
                    });
                }
                return registros;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al obtener todos los registros de compra");
                throw;
            }
        }

        public async Task<List<CompraRegistro>> GetComprasNoProcesadasAsync()
        {
            const string query = "SELECT * FROM CompraRegistros WHERE EstaProcesado = 0 ORDER BY FechaCompra DESC;";
            var registros = new List<CompraRegistro>();

            try
            {
                using var connection = CreateConnection();
                using var command = new SQLiteCommand(query, connection);
                using var reader = await command.ExecuteReaderAsync();

                while (await reader.ReadAsync())
                {
                    registros.Add(new CompraRegistro
                    {
                        Id = Convert.ToInt32(reader["Id"]),
                        FechaCompra = DateTime.Parse(reader["FechaCompra"].ToString()),
                        Proveedor = reader["Proveedor"].ToString(),
                        Producto = reader["Producto"].ToString(),
                        Cantidad = Convert.ToDecimal(reader["Cantidad"]),
                        PrecioUnitario = Convert.ToDecimal(reader["PrecioUnitario"]),
                        Total = Convert.ToDecimal(reader["Total"]),
                        Observaciones = reader["Observaciones"].ToString(),
                        EstaProcesado = false
                    });
                }
                return registros;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al obtener compras no procesadas");
                throw;
            }
        }

        public async Task<int> CreateCompraRegistroAsync(CompraRegistro compraRegistro)
        {
            const string query = @"
                INSERT INTO CompraRegistros 
                (FechaCompra, Proveedor, Producto, Cantidad, PrecioUnitario, Total, Observaciones, EstaProcesado)
                VALUES 
                (@FechaCompra, @Proveedor, @Producto, @Cantidad, @PrecioUnitario, @Total, @Observaciones, @EstaProcesado);
                SELECT last_insert_rowid();";

            try
            {
                using var connection = CreateConnection();
                using var command = new SQLiteCommand(query, connection);

                command.Parameters.AddWithValue("@FechaCompra", compraRegistro.FechaCompra.ToString("yyyy-MM-dd HH:mm:ss"));
                command.Parameters.AddWithValue("@Proveedor", compraRegistro.Proveedor);
                command.Parameters.AddWithValue("@Producto", compraRegistro.Producto);
                command.Parameters.AddWithValue("@Cantidad", compraRegistro.Cantidad);
                command.Parameters.AddWithValue("@PrecioUnitario", compraRegistro.PrecioUnitario);
                command.Parameters.AddWithValue("@Total", compraRegistro.Total);
                command.Parameters.AddWithValue("@Observaciones", compraRegistro.Observaciones ?? "");
                command.Parameters.AddWithValue("@EstaProcesado", compraRegistro.EstaProcesado ? 1 : 0);

                var id = Convert.ToInt32(await command.ExecuteScalarAsync());
                return id;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al crear registro de compra");
                throw;
            }
        }

        public async Task<bool> UpdateCompraRegistroAsync(CompraRegistro compraRegistro)
        {
            const string query = @"
                UPDATE CompraRegistros 
                SET FechaCompra = @FechaCompra,
                    Proveedor = @Proveedor,
                    Producto = @Producto,
                    Cantidad = @Cantidad,
                    PrecioUnitario = @PrecioUnitario,
                    Total = @Total,
                    Observaciones = @Observaciones,
                    EstaProcesado = @EstaProcesado
                WHERE Id = @Id;";

            try
            {
                using var connection = CreateConnection();
                using var command = new SQLiteCommand(query, connection);

                command.Parameters.AddWithValue("@Id", compraRegistro.Id);
                command.Parameters.AddWithValue("@FechaCompra", compraRegistro.FechaCompra.ToString("yyyy-MM-dd HH:mm:ss"));
                command.Parameters.AddWithValue("@Proveedor", compraRegistro.Proveedor);
                command.Parameters.AddWithValue("@Producto", compraRegistro.Producto);
                command.Parameters.AddWithValue("@Cantidad", compraRegistro.Cantidad);
                command.Parameters.AddWithValue("@PrecioUnitario", compraRegistro.PrecioUnitario);
                command.Parameters.AddWithValue("@Total", compraRegistro.Total);
                command.Parameters.AddWithValue("@Observaciones", compraRegistro.Observaciones ?? "");
                command.Parameters.AddWithValue("@EstaProcesado", compraRegistro.EstaProcesado ? 1 : 0);

                return await command.ExecuteNonQueryAsync() > 0;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al actualizar registro de compra");
                throw;
            }
        }

        public async Task<bool> DeleteCompraRegistroAsync(int id)
        {
            const string query = "DELETE FROM CompraRegistros WHERE Id = @Id;";

            try
            {
                using var connection = CreateConnection();
                using var command = new SQLiteCommand(query, connection);
                command.Parameters.AddWithValue("@Id", id);

                return await command.ExecuteNonQueryAsync() > 0;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al eliminar registro de compra");
                throw;
            }
        }

        public async Task<bool> MarcarComoProcesadoAsync(int id)
        {
            const string query = "UPDATE CompraRegistros SET EstaProcesado = 1 WHERE Id = @Id;";

            try
            {
                using var connection = CreateConnection();
                using var command = new SQLiteCommand(query, connection);
                command.Parameters.AddWithValue("@Id", id);

                return await command.ExecuteNonQueryAsync() > 0;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al marcar registro como procesado");
                throw;
            }
        }
    }
}