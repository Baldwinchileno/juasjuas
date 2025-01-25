using System;
using System.Data.SQLite;
using System.Collections.Generic;
using System.Threading.Tasks;
using Microsoft.Extensions.Logging;
using AdminSERMAC.Models;
using AdminSERMAC.Core.Interfaces;

namespace AdminSERMAC.Services.Database
{
    public class ProveedorDatabaseService : BaseSQLiteService, IProveedorService
    {
        private readonly ILogger<ProveedorDatabaseService> _logger;

        public ProveedorDatabaseService(ILogger<ProveedorDatabaseService> logger, string connectionString)
            : base(logger, connectionString)
        {
            _logger = logger;
        }

        public async Task<List<Proveedor>> GetAllProveedores()
        {
            return await ExecuteInTransactionAsync(async (connection, transaction) =>
            {
                var proveedores = new List<Proveedor>();
                const string sql = "SELECT * FROM Proveedores ORDER BY Nombre";

                using var command = new SQLiteCommand(sql, connection, transaction);
                using var reader = await command.ExecuteReaderAsync();

                while (await reader.ReadAsync())
                {
                    proveedores.Add(new Proveedor
                    {
                        Id = reader.GetInt32(reader.GetOrdinal("Id")),
                        Nombre = reader.GetString(reader.GetOrdinal("Nombre")),
                        Vendedor = reader.IsDBNull(reader.GetOrdinal("Vendedor")) ? null : reader.GetString(reader.GetOrdinal("Vendedor")),
                        Telefono = reader.IsDBNull(reader.GetOrdinal("Telefono")) ? null : reader.GetString(reader.GetOrdinal("Telefono")),
                        Email = reader.IsDBNull(reader.GetOrdinal("Email")) ? null : reader.GetString(reader.GetOrdinal("Email")),
                        FechaRegistro = reader.IsDBNull(reader.GetOrdinal("FechaRegistro")) ? null : reader.GetString(reader.GetOrdinal("FechaRegistro"))
                    });
                }

                return proveedores;
            });
        }

        public async Task<Proveedor> GetProveedorById(int id)
        {
            return await ExecuteInTransactionAsync(async (connection, transaction) =>
            {
                const string sql = "SELECT * FROM Proveedores WHERE Id = @Id";

                using var command = new SQLiteCommand(sql, connection, transaction);
                command.Parameters.AddWithValue("@Id", id);

                using var reader = await command.ExecuteReaderAsync();

                if (await reader.ReadAsync())
                {
                    return new Proveedor
                    {
                        Id = reader.GetInt32(reader.GetOrdinal("Id")),
                        Nombre = reader.GetString(reader.GetOrdinal("Nombre")),
                        Vendedor = reader.IsDBNull(reader.GetOrdinal("Vendedor")) ? null : reader.GetString(reader.GetOrdinal("Vendedor")),
                        Telefono = reader.IsDBNull(reader.GetOrdinal("Telefono")) ? null : reader.GetString(reader.GetOrdinal("Telefono")),
                        Email = reader.IsDBNull(reader.GetOrdinal("Email")) ? null : reader.GetString(reader.GetOrdinal("Email")),
                        FechaRegistro = reader.IsDBNull(reader.GetOrdinal("FechaRegistro")) ? null : reader.GetString(reader.GetOrdinal("FechaRegistro"))
                    };
                }

                return null;
            });
        }

        public async Task<bool> CreateProveedor(Proveedor proveedor)
        {
            return await ExecuteInTransactionAsync(async (connection, transaction) =>
            {
                const string sql = @"
                    INSERT INTO Proveedores (Nombre, Vendedor, Telefono, Email, FechaRegistro)
                    VALUES (@Nombre, @Vendedor, @Telefono, @Email, @FechaRegistro)";

                using var command = new SQLiteCommand(sql, connection, transaction);
                command.Parameters.AddWithValue("@Nombre", proveedor.Nombre);
                command.Parameters.AddWithValue("@Vendedor", proveedor.Vendedor ?? (object)DBNull.Value);
                command.Parameters.AddWithValue("@Telefono", proveedor.Telefono ?? (object)DBNull.Value);
                command.Parameters.AddWithValue("@Email", proveedor.Email ?? (object)DBNull.Value);
                command.Parameters.AddWithValue("@FechaRegistro", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));

                return await command.ExecuteNonQueryAsync() > 0;
            });
        }

        public async Task<bool> UpdateProveedor(Proveedor proveedor)
        {
            return await ExecuteInTransactionAsync(async (connection, transaction) =>
            {
                const string sql = @"
                    UPDATE Proveedores 
                    SET Nombre = @Nombre,
                        Vendedor = @Vendedor,
                        Telefono = @Telefono,
                        Email = @Email
                    WHERE Id = @Id";

                using var command = new SQLiteCommand(sql, connection, transaction);
                command.Parameters.AddWithValue("@Id", proveedor.Id);
                command.Parameters.AddWithValue("@Nombre", proveedor.Nombre);
                command.Parameters.AddWithValue("@Vendedor", proveedor.Vendedor ?? (object)DBNull.Value);
                command.Parameters.AddWithValue("@Telefono", proveedor.Telefono ?? (object)DBNull.Value);
                command.Parameters.AddWithValue("@Email", proveedor.Email ?? (object)DBNull.Value);

                return await command.ExecuteNonQueryAsync() > 0;
            });
        }

        public async Task<bool> DeleteProveedor(int id)
        {
            return await ExecuteInTransactionAsync(async (connection, transaction) =>
            {
                const string sql = "DELETE FROM Proveedores WHERE Id = @Id";

                using var command = new SQLiteCommand(sql, connection, transaction);
                command.Parameters.AddWithValue("@Id", id);

                return await command.ExecuteNonQueryAsync() > 0;
            });
        }
    }
}
