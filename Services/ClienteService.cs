using System;
using System.Collections.Generic;
using System.Data.SQLite;
using System.Linq;
using AdminSERMAC.Models;
using AdminSERMAC.Core.Interfaces;
using AdminSERMAC.Exceptions;
using Microsoft.Extensions.Logging;

namespace AdminSERMAC.Services
{
    public class ClienteService : IClienteService
    {
        private readonly IClienteRepository _clienteRepository;
        private readonly ILogger<ClienteService> _logger;
        private readonly string _connectionString;

        public ClienteService(IClienteRepository clienteRepository, ILogger<ClienteService> logger, string connectionString)
        {
            _clienteRepository = clienteRepository;
            _logger = logger;
            _connectionString = connectionString;
        }

        public List<Cliente> ObtenerTodosLosClientes()
        {
            try
            {
                using (var connection = new SQLiteConnection(_connectionString))
                {
                    connection.Open();
                    var command = connection.CreateCommand();
                    command.CommandText = "SELECT * FROM Clientes ORDER BY Nombre COLLATE NOCASE"; // Ordenar por nombre ignorando mayúsculas/minúsculas

                    var clientes = new List<Cliente>();
                    using (var reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            var cliente = new Cliente
                            {
                                RUT = reader["RUT"].ToString(),
                                Nombre = reader["Nombre"].ToString(),
                                Direccion = reader["Direccion"].ToString(),
                                Giro = reader["Giro"].ToString(),
                                Deuda = Convert.ToDouble(reader["Deuda"])
                            };
                            clientes.Add(cliente);
                        }
                    }
                    return clientes;
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al obtener todos los clientes");
                throw new ClienteException("Error al obtener todos los clientes", ex);
            }
        }



        public Cliente ObtenerClientePorRUT(string rut)
        {
            ValidarRUT(rut);
            try
            {
                _logger.LogInformation("Buscando cliente con RUT: {RUT}", rut);
                return _clienteRepository.GetByRUT(rut);
            }
            catch (ClienteException ex)
            {
                _logger.LogError(ex, "Error al obtener cliente con RUT: {RUT}", rut);
                throw;
            }
        }

        public async Task<IEnumerable<Abono>> ObtenerAbonosPorClienteAsync(string rut)
        {
            if (string.IsNullOrWhiteSpace(rut))
            {
                throw new ArgumentException("El RUT no puede ser nulo o vacío.", nameof(rut));
            }

            try
            {
                return await _clienteRepository.GetAbonosPorClienteAsync(rut);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al obtener los abonos para el cliente con RUT: {RUT}", rut);
                throw new ClienteException($"Error al obtener los abonos para el cliente con RUT: {rut}", ex);
            }
        }

        public void AgregarCliente(Cliente cliente)
        {
            ValidarCliente(cliente);
            try
            {
                // Verificar si ya existe un cliente con ese RUT
                try
                {
                    var clienteExistente = _clienteRepository.GetByRUT(cliente.RUT ?? throw new ArgumentNullException(nameof(cliente.RUT)));
                    if (clienteExistente != null)
                    {
                        throw new ClienteDuplicadoException(cliente.RUT);
                    }
                }
                catch (ClienteNotFoundException)
                {
                    // Es correcto que no exista el cliente
                }

                _logger.LogInformation("Agregando nuevo cliente con RUT: {RUT}", cliente.RUT);
                _clienteRepository.Add(cliente);
            }
            catch (ClienteException ex)
            {
                _logger.LogError(ex, "Error al agregar cliente con RUT: {RUT}", cliente.RUT);
                throw;
            }
        }


        public void ActualizarCliente(Cliente cliente)
        {
            ValidarCliente(cliente);
            try
            {
                _logger.LogInformation("Actualizando cliente con RUT: {RUT}", cliente.RUT);
                _clienteRepository.Update(cliente);
            }
            catch (ClienteException ex)
            {
                _logger.LogError(ex, "Error al actualizar cliente con RUT: {RUT}", cliente.RUT);
                throw;
            }
        }

        public void EliminarCliente(string rut)
        {
            ValidarRUT(rut);
            try
            {
                // Verificar si el cliente tiene ventas
                var ventas = _clienteRepository.GetVentasPorCliente(rut);
                if (ventas.Any())
                {
                    throw new ClienteConVentasException(rut);
                }

                _logger.LogInformation("Eliminando cliente con RUT: {RUT}", rut);
                _clienteRepository.Delete(rut);
            }
            catch (ClienteException ex)
            {
                _logger.LogError(ex, "Error al eliminar cliente con RUT: {RUT}", rut);
                throw;
            }
        }

        public void ActualizarDeudaCliente(string rut, double monto)
        {
            using (var connection = new SQLiteConnection(_connectionString))
            {
                connection.Open();
                using (var transaction = connection.BeginTransaction())
                {
                    try
                    {
                        // 1. Verificar deuda actual
                        var commandGetDeuda = new SQLiteCommand(
                            "SELECT Deuda FROM Clientes WHERE RUT = @RUT",
                            connection,
                            transaction);
                        commandGetDeuda.Parameters.AddWithValue("@RUT", rut);
                        var deudaActual = Convert.ToDouble(commandGetDeuda.ExecuteScalar());

                        if (monto < 0 && Math.Abs(monto) > deudaActual)
                        {
                            throw new Exception("El abono no puede ser mayor que la deuda actual");
                        }

                        // 2. Actualizar deuda del cliente
                        var commandUpdateDeuda = new SQLiteCommand(
                            "UPDATE Clientes SET Deuda = Deuda + @Monto WHERE RUT = @RUT",
                            connection,
                            transaction);
                        commandUpdateDeuda.Parameters.AddWithValue("@RUT", rut);
                        commandUpdateDeuda.Parameters.AddWithValue("@Monto", monto);
                        commandUpdateDeuda.ExecuteNonQuery();

                        var fecha = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");

                        // 3. Registrar en HistorialMovimientos
                        var commandHistorial = new SQLiteCommand(@"
                    INSERT INTO HistorialMovimientos (
                        RUT, 
                        Tipo, 
                        Monto, 
                        Fecha
                    ) VALUES (
                        @RUT,
                        @Tipo,
                        @MontoAbs,
                        @Fecha
                    )", connection, transaction);

                        commandHistorial.Parameters.AddWithValue("@RUT", rut);
                        commandHistorial.Parameters.AddWithValue("@Tipo", monto < 0 ? "ABONO" : "CARGO");
                        commandHistorial.Parameters.AddWithValue("@MontoAbs", Math.Abs(monto));
                        commandHistorial.Parameters.AddWithValue("@Fecha", fecha);
                        commandHistorial.ExecuteNonQuery();

                        // 4. Si es un abono (monto negativo), registrar en la tabla Abonos
                        if (monto < 0)
                        {
                            var commandAbono = new SQLiteCommand(@"
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
                        )", connection, transaction);

                            commandAbono.Parameters.AddWithValue("@RUT", rut);
                            commandAbono.Parameters.AddWithValue("@Fecha", fecha);
                            commandAbono.Parameters.AddWithValue("@Monto", Math.Abs(monto));
                            commandAbono.Parameters.AddWithValue("@Descripcion", $"Abono registrado el {DateTime.Now:dd/MM/yyyy HH:mm:ss}");
                            commandAbono.ExecuteNonQuery();
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

        public IEnumerable<Abono> ObtenerAbonosPorCliente(string rut)
        {
            return _clienteRepository.GetAbonosPorClienteAsync(rut).Result;
        }

        public List<Venta> ObtenerVentasCliente(string rut)
        {
            ValidarRUT(rut);
            try
            {
                _logger.LogInformation("Obteniendo ventas del cliente: {RUT}", rut);
                return _clienteRepository.GetVentasPorCliente(rut);
            }
            catch (ClienteException ex)
            {
                _logger.LogError(ex, "Error al obtener ventas del cliente: {RUT}", rut);
                throw;
            }
        }

        public double CalcularDeudaTotal(string rut)
        {
            ValidarRUT(rut);
            try
            {
                using (var connection = new SQLiteConnection(_connectionString))
                {
                    connection.Open();
                    var command = new SQLiteCommand("SELECT Deuda FROM Clientes WHERE RUT = @RUT", connection);
                    command.Parameters.AddWithValue("@RUT", rut);
                    var result = command.ExecuteScalar();

                    if (result == null || result == DBNull.Value)
                    {
                        throw new ClienteNotFoundException(rut);
                    }

                    return Convert.ToDouble(result);
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al calcular deuda total del cliente: {RUT}", rut);
                throw new ClienteException($"Error al calcular deuda total del cliente con RUT: {rut}", ex);
            }
        }

        private void ValidarCliente(Cliente cliente)
        {
            if (cliente == null)
            {
                throw new ClienteValidationException("El cliente no puede ser nulo");
            }

            ValidarRUT(cliente.RUT ?? throw new ArgumentNullException(nameof(cliente.RUT)));

            if (string.IsNullOrWhiteSpace(cliente.Nombre))
            {
                throw new ClienteValidationException("El nombre del cliente es obligatorio");
            }

            if (string.IsNullOrWhiteSpace(cliente.Direccion))
            {
                throw new ClienteValidationException("La dirección del cliente es obligatoria");
            }

            if (string.IsNullOrWhiteSpace(cliente.Giro))
            {
                throw new ClienteValidationException("El giro del cliente es obligatorio");
            }

            if (cliente.Deuda < 0)
            {
                throw new ClienteValidationException("La deuda no puede ser negativa");
            }
        }

        private void ValidarRUT(string rut)
        {
            if (string.IsNullOrWhiteSpace(rut))
            {
                throw new ClienteValidationException("El RUT es obligatorio");
            }
        }
    }
}