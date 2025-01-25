using System.Data.SQLite;
using Microsoft.Extensions.Logging;
using AdminSERMAC.Models;

namespace AdminSERMAC.Services.Database
{
    public interface IClienteDatabaseService
    {
        Task<Cliente?> GetByRutAsync(string rut);
        Task<IEnumerable<Cliente>> GetAllAsync();
        Task<bool> CreateAsync(Cliente cliente);
        Task<bool> UpdateAsync(Cliente cliente);
        Task<bool> UpdateDeudaAsync(string rut, double monto);
    }

    public class ClienteDatabaseService : BaseSQLiteService, IClienteDatabaseService
    {
        private const string TableName = "Clientes";

        public ClienteDatabaseService(ILogger<ClienteDatabaseService> logger, string connectionString)
            : base(logger, connectionString)
        {
            EnsureTableExists();
        }

        private void EnsureTableExists()
        {
            const string createTableSql = @"
                CREATE TABLE IF NOT EXISTS Clientes (
                    RUT TEXT PRIMARY KEY,
                    Nombre TEXT NOT NULL,
                    Direccion TEXT NOT NULL,
                    Giro TEXT NOT NULL,
                    Deuda DECIMAL(20,2) DEFAULT 0
                )";

            ExecuteInTransactionAsync(async (connection, transaction) =>
            {
                using var command = new SQLiteCommand(createTableSql, connection, transaction);
                await command.ExecuteNonQueryAsync();
            }).Wait();
        }

        public async Task<Cliente?> GetByRutAsync(string rut)
        {
            return await ExecuteInTransactionAsync(async (connection, transaction) =>
            {
                const string sql = "SELECT * FROM Clientes WHERE RUT = @rut";
                using var command = new SQLiteCommand(sql, connection, transaction);
                command.Parameters.AddWithValue("@rut", rut);

                using var reader = await command.ExecuteReaderAsync();
                if (await reader.ReadAsync())
                {
                    return new Cliente
                    {
                        RUT = reader["RUT"].ToString(),
                        Nombre = reader["Nombre"].ToString(),
                        Direccion = reader["Direccion"].ToString(),
                        Giro = reader["Giro"].ToString(),
                        Deuda = Convert.ToDouble(reader["Deuda"])
                    };
                }
                return null;
            });
        }

        public async Task<IEnumerable<Cliente>> GetAllAsync()
        {
            return await ExecuteInTransactionAsync(async (connection, transaction) =>
            {
                var clientes = new List<Cliente>();
                const string sql = "SELECT * FROM Clientes";
                using var command = new SQLiteCommand(sql, connection, transaction);
                using var reader = await command.ExecuteReaderAsync();

                while (await reader.ReadAsync())
                {
                    clientes.Add(new Cliente
                    {
                        RUT = reader["RUT"].ToString(),
                        Nombre = reader["Nombre"].ToString(),
                        Direccion = reader["Direccion"].ToString(),
                        Giro = reader["Giro"].ToString(),
                        Deuda = Convert.ToDouble(reader["Deuda"])
                    });
                }
                return clientes;
            });
        }

        public async Task<bool> CreateAsync(Cliente cliente)
        {
            if (!cliente.IsValid())
            {
                _logger.LogWarning("Attempted to create invalid cliente: {Rut}", cliente.RUT);
                return false;
            }

            return await ExecuteInTransactionAsync(async (connection, transaction) =>
            {
                const string sql = @"
                    INSERT INTO Clientes (RUT, Nombre, Direccion, Giro, Deuda)
                    VALUES (@rut, @nombre, @direccion, @giro, @deuda)";

                using var command = new SQLiteCommand(sql, connection, transaction);
                command.Parameters.AddWithValue("@rut", cliente.RUT);
                command.Parameters.AddWithValue("@nombre", cliente.Nombre);
                command.Parameters.AddWithValue("@direccion", cliente.Direccion);
                command.Parameters.AddWithValue("@giro", cliente.Giro);
                command.Parameters.AddWithValue("@deuda", cliente.Deuda);

                return await command.ExecuteNonQueryAsync() > 0;
            });
        }

        public async Task<bool> UpdateAsync(Cliente cliente)
        {
            if (!cliente.IsValid())
            {
                _logger.LogWarning("Attempted to update invalid cliente: {Rut}", cliente.RUT);
                return false;
            }

            return await ExecuteInTransactionAsync(async (connection, transaction) =>
            {
                const string sql = @"
                    UPDATE Clientes 
                    SET Nombre = @nombre,
                        Direccion = @direccion,
                        Giro = @giro,
                        Deuda = @deuda
                    WHERE RUT = @rut";

                using var command = new SQLiteCommand(sql, connection, transaction);
                command.Parameters.AddWithValue("@rut", cliente.RUT);
                command.Parameters.AddWithValue("@nombre", cliente.Nombre);
                command.Parameters.AddWithValue("@direccion", cliente.Direccion);
                command.Parameters.AddWithValue("@giro", cliente.Giro);
                command.Parameters.AddWithValue("@deuda", cliente.Deuda);

                return await command.ExecuteNonQueryAsync() > 0;
            });
        }

        public async Task<bool> UpdateDeudaAsync(string rut, double monto)
        {
            return await ExecuteInTransactionAsync(async (connection, transaction) =>
            {
                const string sql = @"
                    UPDATE Clientes 
                    SET Deuda = Deuda + @monto
                    WHERE RUT = @rut";

                using var command = new SQLiteCommand(sql, connection, transaction);
                command.Parameters.AddWithValue("@rut", rut);
                command.Parameters.AddWithValue("@monto", monto);

                return await command.ExecuteNonQueryAsync() > 0;
            });
        }
    }
}
