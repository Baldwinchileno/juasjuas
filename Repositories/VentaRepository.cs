using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Threading.Tasks;
using AdminSERMAC.Core.Interfaces;
using AdminSERMAC.Models;
using Microsoft.Extensions.Logging;
using System.Data.SQLite;
using Dapper;

namespace AdminSERMAC.Repositories
{
    public class VentaRepository : IVentaRepository
    {
        private readonly string _connectionString;
        private readonly ILogger<VentaRepository> _logger;

        public VentaRepository(string connectionString, ILogger<VentaRepository> logger)
        {
            _connectionString = connectionString;
            _logger = logger;
        }

        public async Task<IEnumerable<Venta>> GetVentasPorClienteAsync(string rut)
        {
            try
            {
                using var connection = new SQLiteConnection(_connectionString);
                var query = @"
                    SELECT v.*, c.Nombre as ClienteNombre 
                    FROM Ventas v
                    JOIN Clientes c ON v.RUT = c.RUT
                    WHERE v.RUT = @rut";
                return await connection.QueryAsync<Venta>(query, new { rut });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error getting ventas for cliente");
                throw;
            }
        }

        public async Task<IEnumerable<Venta>> GetVentasEnRangoAsync(DateTime inicio, DateTime fin)
        {
            try
            {
                using var connection = new SQLiteConnection(_connectionString);
                var query = @"
                    SELECT v.*, c.Nombre as ClienteNombre 
                    FROM Ventas v
                    JOIN Clientes c ON v.RUT = c.RUT
                    WHERE date(v.FechaVenta) BETWEEN @inicio AND @fin";
                return await connection.QueryAsync<Venta>(query,
                    new
                    {
                        inicio = inicio.ToString("yyyy-MM-dd"),
                        fin = fin.ToString("yyyy-MM-dd")
                    });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error getting ventas in range");
                throw;
            }
        }

        public async Task<double> CalcularTotalVentasAsync(DateTime inicio, DateTime fin)
        {
            try
            {
                using var connection = new SQLiteConnection(_connectionString);
                var query = @"
                    SELECT COALESCE(SUM(Total), 0)
                    FROM Ventas
                    WHERE date(FechaVenta) BETWEEN @inicio AND @fin";
                return await connection.ExecuteScalarAsync<double>(query,
                    new
                    {
                        inicio = inicio.ToString("yyyy-MM-dd"),
                        fin = fin.ToString("yyyy-MM-dd")
                    });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error calculating total ventas");
                throw;
            }
        }

        public async Task<Venta> GetByIdAsync(object id)
        {
            try
            {
                using var connection = new SQLiteConnection(_connectionString);
                var query = @"
                    SELECT v.*, c.Nombre as ClienteNombre 
                    FROM Ventas v
                    JOIN Clientes c ON v.RUT = c.RUT
                    WHERE v.NumeroGuia = @Id";
                return await connection.QueryFirstOrDefaultAsync<Venta>(query, new { Id = id });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error getting venta by id");
                throw;
            }
        }

        public async Task<IEnumerable<Venta>> GetAllAsync()
        {
            try
            {
                using var connection = new SQLiteConnection(_connectionString);
                var query = @"
                    SELECT v.*, c.Nombre as ClienteNombre 
                    FROM Ventas v
                    JOIN Clientes c ON v.RUT = c.RUT";
                return await connection.QueryAsync<Venta>(query);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error getting all ventas");
                throw;
            }
        }

        public async Task AddAsync(Venta venta)
        {
            try
            {
                using var connection = new SQLiteConnection(_connectionString);
                var query = @"
                    INSERT INTO Ventas (
                        NumeroGuia, CodigoProducto, Descripcion, Bandejas,
                        KilosNeto, FechaVenta, PagadoConCredito, RUT, Total
                    ) VALUES (
                        @NumeroGuia, @CodigoProducto, @Descripcion, @Bandejas,
                        @KilosNeto, @FechaVenta, @PagadoConCredito, @RUT, @Total
                    )";
                await connection.ExecuteAsync(query, venta);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error adding venta");
                throw;
            }
        }

        public async Task UpdateAsync(Venta venta)
        {
            try
            {
                using var connection = new SQLiteConnection(_connectionString);
                var query = @"
                    UPDATE Ventas 
                    SET CodigoProducto = @CodigoProducto,
                        Descripcion = @Descripcion,
                        Bandejas = @Bandejas,
                        KilosNeto = @KilosNeto,
                        FechaVenta = @FechaVenta,
                        PagadoConCredito = @PagadoConCredito,
                        RUT = @RUT,
                        Total = @Total
                    WHERE NumeroGuia = @NumeroGuia";
                await connection.ExecuteAsync(query, venta);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error updating venta");
                throw;
            }
        }

        // Implementar los demás métodos según necesidad
        public Task<IEnumerable<Venta>> FindAsync(Expression<Func<Venta, bool>> predicate)
        {
            throw new NotImplementedException();
        }

        public Task AddRangeAsync(IEnumerable<Venta> entities)
        {
            throw new NotImplementedException();
        }

        public Task DeleteAsync(Venta entity)
        {
            throw new NotImplementedException();
        }

        public Task DeleteRangeAsync(IEnumerable<Venta> entities)
        {
            throw new NotImplementedException();
        }

        public Task<bool> ExistsAsync(Expression<Func<Venta, bool>> predicate)
        {
            throw new NotImplementedException();
        }

        public Task<int> CountAsync(Expression<Func<Venta, bool>> predicate = null)
        {
            throw new NotImplementedException();
        }

        public Task<(IEnumerable<Venta> Items, int TotalCount)> GetPagedAsync(
            int pageIndex, int pageSize, Expression<Func<Venta, bool>> predicate = null,
            Func<IQueryable<Venta>, IOrderedQueryable<Venta>> orderBy = null)
        {
            throw new NotImplementedException();
        }
    }
}