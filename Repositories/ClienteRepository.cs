using System;
using System.Collections.Generic;
using System.Data.SQLite;
using System.Linq;
using System.Linq.Expressions;
using System.Threading.Tasks;
using AdminSERMAC.Constants;
using AdminSERMAC.Exceptions;
using AdminSERMAC.Core.Interfaces;
using AdminSERMAC.Models;
using Microsoft.Extensions.Logging;
using Dapper;

namespace AdminSERMAC.Repositories
{
    public class ClienteRepository : IClienteRepository
    {
        private readonly string connectionString;
        private readonly ILogger<ClienteRepository> _logger;


        public ClienteRepository(string connectionString, ILogger<ClienteRepository> logger)
        {
            this.connectionString = connectionString ?? throw new ArgumentNullException(nameof(connectionString));
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
        }

        private SQLiteConnection CreateConnection()
        {
            if (string.IsNullOrEmpty(connectionString))
            {
                throw new InvalidOperationException("La cadena de conexión no está configurada.");
            }
            return new SQLiteConnection(connectionString);
        }



        // Implementación de IClienteRepository específico
        public List<Cliente> GetAll()
        {
            try
            {
                var clientes = new List<Cliente>();
                using (var connection = new SQLiteConnection(connectionString))
                {
                    connection.Open();
                    using (var command = new SQLiteCommand(QueryConstants.Cliente.SELECT_ALL, connection))
                    {
                        using (var reader = command.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                clientes.Add(MapClienteFromReader(reader));
                            }
                        }
                    }
                }
                return clientes;
            }
            catch (SQLiteException ex)
            {
                _logger.LogError(ex, "Error al obtener todos los clientes");
                throw new ClienteException("Error al obtener la lista de clientes", ex);
            }
        }

        public Cliente GetByRUT(string rut)
        {
            try
            {
                using (var connection = new SQLiteConnection(connectionString))
                {
                    connection.Open();
                    using (var command = new SQLiteCommand(QueryConstants.Cliente.SELECT_BY_RUT, connection))
                    {
                        command.Parameters.AddWithValue("@RUT", rut);
                        using (var reader = command.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                return MapClienteFromReader(reader);
                            }
                        }
                    }
                }
                throw new ClienteNotFoundException(rut);
            }
            catch (SQLiteException ex)
            {
                _logger.LogError(ex, "Error al buscar cliente por RUT: {RUT}", rut);
                throw new ClienteException($"Error al buscar cliente con RUT: {rut}", ex);
            }
        }

        // Implementaciones de IRepository<Cliente>
        public async Task<Cliente> GetByIdAsync(object id)
        {
            try
            {
                using (var connection = new SQLiteConnection(connectionString))
                {
                    await connection.OpenAsync();
                    var query = "SELECT * FROM Clientes WHERE RUT = @Id";
                    return await connection.QueryFirstOrDefaultAsync<Cliente>(query, new { Id = id });
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error getting cliente by id: {id}", id);
                throw;
            }
        }

        public async Task<IEnumerable<Abono>> GetAbonosPorClienteAsync(string rut)
        {
            using var connection = CreateConnection();
            var query = @"
        SELECT Id, RUT, Fecha, Monto, Descripcion
        FROM Abonos
        WHERE RUT = @RUT
        ORDER BY Fecha DESC";
            return await connection.QueryAsync<Abono>(query, new { RUT = rut });
        }

        public async Task<IEnumerable<Cliente>> GetAllAsync()
        {
            try
            {
                using (var connection = new SQLiteConnection(connectionString))
                {
                    await connection.OpenAsync();
                    return await connection.QueryAsync<Cliente>("SELECT * FROM Clientes");
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error getting all clientes");
                throw;
            }
        }

        public async Task<IEnumerable<Cliente>> FindAsync(Expression<Func<Cliente, bool>> predicate)
        {
            try
            {
                using (var connection = new SQLiteConnection(connectionString))
                {
                    await connection.OpenAsync();
                    var clientes = await connection.QueryAsync<Cliente>("SELECT * FROM Clientes");
                    return clientes.Where(predicate.Compile());
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error finding clientes");
                throw;
            }
        }

        public async Task AddAsync(Cliente entity)
        {
            try
            {
                using (var connection = new SQLiteConnection(connectionString))
                {
                    await connection.OpenAsync();
                    var query = "INSERT INTO Clientes (RUT, Nombre, Direccion, Giro, Deuda) VALUES (@RUT, @Nombre, @Direccion, @Giro, @Deuda)";
                    await connection.ExecuteAsync(query, entity);
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error adding cliente");
                throw;
            }
        }

        public async Task AddRangeAsync(IEnumerable<Cliente> entities)
        {
            try
            {
                using (var connection = new SQLiteConnection(connectionString))
                {
                    await connection.OpenAsync();
                    using (var transaction = connection.BeginTransaction())
                    {
                        try
                        {
                            var query = "INSERT INTO Clientes (RUT, Nombre, Direccion, Giro, Deuda) VALUES (@RUT, @Nombre, @Direccion, @Giro, @Deuda)";
                            await connection.ExecuteAsync(query, entities, transaction);
                            transaction.Commit();
                        }
                        catch
                        {
                            transaction.Rollback();
                            throw;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error adding range of clientes");
                throw;
            }
        }

        public async Task UpdateAsync(Cliente entity)
        {
            try
            {
                using (var connection = new SQLiteConnection(connectionString))
                {
                    await connection.OpenAsync();
                    var query = "UPDATE Clientes SET Nombre = @Nombre, Direccion = @Direccion, Giro = @Giro, Deuda = @Deuda WHERE RUT = @RUT";
                    await connection.ExecuteAsync(query, entity);
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error updating cliente");
                throw;
            }
        }

        public async Task DeleteAsync(Cliente entity)
        {
            try
            {
                using (var connection = new SQLiteConnection(connectionString))
                {
                    await connection.OpenAsync();
                    var query = "DELETE FROM Clientes WHERE RUT = @RUT";
                    await connection.ExecuteAsync(query, new { entity.RUT });
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error deleting cliente");
                throw;
            }
        }

        public async Task DeleteRangeAsync(IEnumerable<Cliente> entities)
        {
            try
            {
                using (var connection = new SQLiteConnection(connectionString))
                {
                    await connection.OpenAsync();
                    using (var transaction = connection.BeginTransaction())
                    {
                        try
                        {
                            var query = "DELETE FROM Clientes WHERE RUT = @RUT";
                            foreach (var entity in entities)
                            {
                                await connection.ExecuteAsync(query, new { entity.RUT }, transaction);
                            }
                            transaction.Commit();
                        }
                        catch
                        {
                            transaction.Rollback();
                            throw;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error deleting range of clientes");
                throw;
            }
        }

        public async Task<bool> ExistsAsync(Expression<Func<Cliente, bool>> predicate)
        {
            try
            {
                var items = await FindAsync(predicate);
                return items.Any();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error checking if cliente exists");
                throw;
            }
        }

        public async Task<int> CountAsync(Expression<Func<Cliente, bool>> predicate = null)
        {
            try
            {
                using (var connection = new SQLiteConnection(connectionString))
                {
                    await connection.OpenAsync();
                    if (predicate == null)
                    {
                        return await connection.ExecuteScalarAsync<int>("SELECT COUNT(*) FROM Clientes");
                    }
                    var items = await FindAsync(predicate);
                    return items.Count();
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error counting clientes");
                throw;
            }
        }

        public async Task<(IEnumerable<Cliente> Items, int TotalCount)> GetPagedAsync(
            int pageIndex,
            int pageSize,
            Expression<Func<Cliente, bool>> predicate = null,
            Func<IQueryable<Cliente>, IOrderedQueryable<Cliente>> orderBy = null)
        {
            try
            {
                using (var connection = new SQLiteConnection(connectionString))
                {
                    await connection.OpenAsync();
                    var query = "SELECT * FROM Clientes LIMIT @PageSize OFFSET @Offset";
                    var countQuery = "SELECT COUNT(*) FROM Clientes";

                    var totalCount = await connection.ExecuteScalarAsync<int>(countQuery);
                    var items = await connection.QueryAsync<Cliente>(
                        query,
                        new { PageSize = pageSize, Offset = pageIndex * pageSize });

                    return (items, totalCount);
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error getting paged clientes");
                throw;
            }
        }

        // Implementación de los métodos específicos de IClienteRepository
        public async Task<IEnumerable<Cliente>> GetClientesConDeudaAsync()
        {
            try
            {
                using (var connection = new SQLiteConnection(connectionString))
                {
                    await connection.OpenAsync();
                    return await connection.QueryAsync<Cliente>("SELECT * FROM Clientes WHERE Deuda > 0");
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error getting clientes con deuda");
                throw;
            }
        }

        public async Task<double> CalcularDeudaTotalAsync(string rut)
        {
            try
            {
                using (var connection = new SQLiteConnection(connectionString))
                {
                    await connection.OpenAsync();
                    return await connection.ExecuteScalarAsync<double>(
                        "SELECT Deuda FROM Clientes WHERE RUT = @RUT",
                        new { RUT = rut });
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error calculating deuda total");
                throw;
            }
        }

        public async Task ActualizarDeudaAsync(string rut, double monto)
        {
            try
            {
                using (var connection = new SQLiteConnection(connectionString))
                {
                    await connection.OpenAsync();
                    await connection.ExecuteAsync(
                        "UPDATE Clientes SET Deuda = Deuda + @Monto WHERE RUT = @RUT",
                        new { RUT = rut, Monto = monto });
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error updating deuda");
                throw;
            }
        }

        // Métodos auxiliares
        private Cliente MapClienteFromReader(SQLiteDataReader reader)
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

        public List<Venta> GetVentasPorCliente(string rut)
        {
            try
            {
                var ventas = new List<Venta>();
                using (var connection = new SQLiteConnection(connectionString))
                {
                    connection.Open();
                    using (var command = new SQLiteCommand(QueryConstants.Ventas.SELECT_VENTAS, connection))

                    {
                        command.Parameters.AddWithValue("@RUT", rut);
                        using (var reader = command.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                ventas.Add(MapVentaFromReader(reader));
                            }
                        }
                    }
                }
                return ventas;
            }
            catch (SQLiteException ex)
            {
                _logger.LogError(ex, "Error al obtener ventas del cliente: {RUT}", rut);
                throw new ClienteException($"Error al obtener ventas del cliente con RUT: {rut}", ex);
            }
        }

        private Venta MapVentaFromReader(SQLiteDataReader reader)
        {
            try
            {
                return new Venta
                {
                    NumeroGuia = reader.GetInt32(reader.GetOrdinal("NumeroGuia")),      // Entero
                    FechaVenta = reader.GetDateTime(reader.GetOrdinal("FechaVenta")).ToString("yyyy-MM-dd HH:mm:ss"),   // Fecha
                    Descripcion = reader.GetString(reader.GetOrdinal("Descripcion")),   // Texto
                    KilosNeto = reader.GetDouble(reader.GetOrdinal("KilosNeto")),       // Decimal
                    Total = reader.GetDouble(reader.GetOrdinal("Total")),               // Decimal
                    PagadoConCredito = reader.GetInt32(reader.GetOrdinal("PagadoConCredito")) // Entero
                };
            }
            catch (Exception ex)
            {
                throw new Exception($"Error al mapear la venta: {ex.Message}", ex);
            }
        }







        // Implementación de los métodos síncronos de IClienteRepository
        public void Add(Cliente cliente)
        {
            try
            {
                using (var connection = new SQLiteConnection(connectionString))
                {
                    connection.Open();
                    using (var transaction = connection.BeginTransaction())
                    {
                        try
                        {
                            using (var command = new SQLiteCommand(QueryConstants.Cliente.INSERT, connection))
                            {
                                SetClienteParameters(command, cliente);
                                command.ExecuteNonQuery();
                            }
                            transaction.Commit();
                            _logger.LogInformation("Cliente agregado: {RUT}", cliente.RUT);
                        }
                        catch (Exception)
                        {
                            transaction.Rollback();
                            throw;
                        }
                    }
                }
            }
            catch (SQLiteException ex)
            {
                _logger.LogError(ex, "Error al agregar cliente: {RUT}", cliente.RUT);
                throw new ClienteException($"Error al agregar cliente con RUT: {cliente.RUT}", ex);
            }
        }

        public void Update(Cliente cliente)
        {
            try
            {
                using (var connection = new SQLiteConnection(connectionString))
                {
                    connection.Open();
                    using (var transaction = connection.BeginTransaction())
                    {
                        try
                        {
                            using (var command = new SQLiteCommand(QueryConstants.Cliente.UPDATE, connection))
                            {
                                SetClienteParameters(command, cliente);
                                int rowsAffected = command.ExecuteNonQuery();
                                if (rowsAffected == 0)
                                {
                                    throw new ClienteNotFoundException(cliente.RUT);
                                }
                            }
                            transaction.Commit();
                            _logger.LogInformation("Cliente actualizado: {RUT}", cliente.RUT);
                        }
                        catch (Exception)
                        {
                            transaction.Rollback();
                            throw;
                        }
                    }
                }
            }
            catch (SQLiteException ex)
            {
                _logger.LogError(ex, "Error al actualizar cliente: {RUT}", cliente.RUT);
                throw new ClienteException($"Error al actualizar cliente con RUT: {cliente.RUT}", ex);
            }
        }

        public void Delete(string rut)
        {
            try
            {
                using (var connection = new SQLiteConnection(connectionString))
                {
                    connection.Open();
                    using (var transaction = connection.BeginTransaction())
                    {
                        try
                        {
                            using (var command = new SQLiteCommand(QueryConstants.Cliente.DELETE, connection))
                            {
                                command.Parameters.AddWithValue("@RUT", rut);
                                int rowsAffected = command.ExecuteNonQuery();
                                if (rowsAffected == 0)
                                {
                                    throw new ClienteNotFoundException(rut);
                                }
                            }
                            transaction.Commit();
                            _logger.LogInformation("Cliente eliminado: {RUT}", rut);
                        }
                        catch (Exception)
                        {
                            transaction.Rollback();
                            throw;
                        }
                    }
                }
            }
            catch (SQLiteException ex)
            {
                _logger.LogError(ex, "Error al eliminar cliente: {RUT}", rut);
                throw new ClienteException($"Error al eliminar cliente con RUT: {rut}", ex);
            }
        }

        public void UpdateDeuda(string rut, double monto)
        {
            try
            {
                using (var connection = new SQLiteConnection(connectionString))
                {
                    connection.Open();
                    using (var transaction = connection.BeginTransaction())
                    {
                        try
                        {
                            // 1. Actualizar la deuda del cliente
                            using (var command = new SQLiteCommand(QueryConstants.Cliente.UPDATE_DEUDA, connection))
                            {
                                command.Transaction = transaction;
                                command.Parameters.AddWithValue("@RUT", rut);
                                command.Parameters.AddWithValue("@Monto", monto);

                                int rowsAffected = command.ExecuteNonQuery();
                                if (rowsAffected == 0)
                                {
                                    throw new ClienteNotFoundException(rut);
                                }
                            }

                            // 2. Registrar en el historial de movimientos
                            using (var historialCommand = new SQLiteCommand(@"
                        INSERT INTO HistorialMovimientos (
                            RUT, 
                            Tipo, 
                            Monto, 
                            Fecha
                        ) VALUES (
                            @RUT,
                            @Tipo,
                            @Monto,
                            @Fecha
                        )", connection))
                            {
                                historialCommand.Transaction = transaction;
                                var fecha = DateTime.Now;
                                historialCommand.Parameters.AddWithValue("@RUT", rut);
                                historialCommand.Parameters.AddWithValue("@Tipo", monto < 0 ? "ABONO" : "CARGO");
                                historialCommand.Parameters.AddWithValue("@Monto", Math.Abs(monto));
                                historialCommand.Parameters.AddWithValue("@Fecha", fecha.ToString("yyyy-MM-dd HH:mm:ss"));

                                historialCommand.ExecuteNonQuery();
                            }

                            // 3. Si es un abono (monto negativo), registrar en la tabla Abonos
                            if (monto < 0)
                            {
                                using (var abonoCommand = new SQLiteCommand(@"
                            INSERT INTO Abonos (
                                RUT,
                                Fecha,
                                Monto,
                                Descripcion
                            ) VALUES (
                                @RUT,
                                @Fecha,
                                @Monto,
                                @Descripcion
                            )", connection))
                                {
                                    abonoCommand.Transaction = transaction;
                                    abonoCommand.Parameters.AddWithValue("@RUT", rut);
                                    abonoCommand.Parameters.AddWithValue("@Fecha", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
                                    abonoCommand.Parameters.AddWithValue("@Monto", Math.Abs(monto));
                                    abonoCommand.Parameters.AddWithValue("@Descripcion", $"Abono registrado el {DateTime.Now:dd/MM/yyyy HH:mm:ss}");

                                    abonoCommand.ExecuteNonQuery();
                                }
                            }

                            transaction.Commit();
                            _logger.LogInformation("Deuda actualizada para cliente: {RUT}, Monto: {Monto}", rut, monto);
                        }
                        catch
                        {
                            transaction.Rollback();
                            throw;
                        }
                    }
                }
            }
            catch (SQLiteException ex)
            {
                _logger.LogError(ex, "Error al actualizar deuda del cliente: {RUT}", rut);
                throw new ClienteException($"Error al actualizar deuda del cliente con RUT: {rut}", ex);
            }
        }

        private void SetClienteParameters(SQLiteCommand command, Cliente cliente)
        {
            command.Parameters.AddWithValue("@RUT", cliente.RUT);
            command.Parameters.AddWithValue("@Nombre", cliente.Nombre);
            command.Parameters.AddWithValue("@Direccion", cliente.Direccion);
            command.Parameters.AddWithValue("@Giro", cliente.Giro);
            command.Parameters.AddWithValue("@Deuda", cliente.Deuda);
        }
    }
}