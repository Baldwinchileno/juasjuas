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
    public class InventarioRepository : IInventarioRepository
    {
        private readonly string _connectionString;
        private readonly ILogger<InventarioRepository> _logger;

        public InventarioRepository(string connectionString, ILogger<InventarioRepository> logger)
        {
            _connectionString = connectionString;
            _logger = logger;
        }

        public async Task<IEnumerable<Inventario>> GetInventarioPorFechaAsync(DateTime fecha)
        {
            try
            {
                using var connection = new SQLiteConnection(_connectionString);
                var query = @"
                    SELECT * FROM Inventario 
                    WHERE date(FechaMasAntigua) <= @fecha 
                    AND date(FechaMasNueva) >= @fecha";
                return await connection.QueryAsync<Inventario>(query,
                    new { fecha = fecha.ToString("yyyy-MM-dd") });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error getting inventario by date");
                throw;
            }
        }

        public async Task ActualizarInventarioAsync(string codigo, int unidades, double kilos)
        {
            try
            {
                using var connection = new SQLiteConnection(_connectionString);
                var query = @"
                    UPDATE Inventario 
                    SET Unidades = Unidades + @unidades,
                        Kilos = Kilos + @kilos
                    WHERE Codigo = @codigo";
                await connection.ExecuteAsync(query, new { codigo, unidades, kilos });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error updating inventario");
                throw;
            }
        }

        public async Task<IEnumerable<Inventario>> GetInventarioProximoAVencerAsync(int diasLimite)
        {
            try
            {
                using var connection = new SQLiteConnection(_connectionString);
                var query = @"
                    SELECT * FROM Inventario 
                    WHERE julianday('now') - julianday(FechaMasAntigua) >= @diasLimite
                    AND Unidades > 0";
                return await connection.QueryAsync<Inventario>(query, new { diasLimite });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error getting inventario próximo a vencer");
                throw;
            }
        }

        public async Task<Inventario> GetByIdAsync(object id)
        {
            try
            {
                using var connection = new SQLiteConnection(_connectionString);
                var query = "SELECT * FROM Inventario WHERE Codigo = @Id";
                return await connection.QueryFirstOrDefaultAsync<Inventario>(query, new { Id = id });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error getting inventario by id");
                throw;
            }
        }

        public async Task<IEnumerable<Inventario>> GetAllAsync()
        {
            try
            {
                using var connection = new SQLiteConnection(_connectionString);
                var query = "SELECT * FROM Inventario";
                return await connection.QueryAsync<Inventario>(query);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error getting all inventario");
                throw;
            }
        }

        public async Task AddAsync(Inventario inventario)
        {
            try
            {
                using var connection = new SQLiteConnection(_connectionString);
                var query = @"
                    INSERT INTO Inventario (
                        Codigo, Producto, Unidades, Kilos,
                        FechaMasAntigua, FechaMasNueva
                    ) VALUES (
                        @Codigo, @Producto, @Unidades, @Kilos,
                        @FechaMasAntigua, @FechaMasNueva
                    )";
                await connection.ExecuteAsync(query, inventario);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error adding inventario");
                throw;
            }
        }

        public async Task UpdateAsync(Inventario inventario)
        {
            try
            {
                using var connection = new SQLiteConnection(_connectionString);
                var query = @"
                    UPDATE Inventario 
                    SET Producto = @Producto,
                        Unidades = @Unidades,
                        Kilos = @Kilos,
                        FechaMasAntigua = @FechaMasAntigua,
                        FechaMasNueva = @FechaMasNueva
                    WHERE Codigo = @Codigo";
                await connection.ExecuteAsync(query, inventario);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error updating inventario");
                throw;
            }
        }

        // Implementar los demás métodos según necesidad
        public Task<IEnumerable<Inventario>> FindAsync(Expression<Func<Inventario, bool>> predicate)
        {
            throw new NotImplementedException();
        }

        public Task AddRangeAsync(IEnumerable<Inventario> entities)
        {
            throw new NotImplementedException();
        }

        public Task DeleteAsync(Inventario entity)
        {
            throw new NotImplementedException();
        }

        public Task DeleteRangeAsync(IEnumerable<Inventario> entities)
        {
            throw new NotImplementedException();
        }

        public Task<bool> ExistsAsync(Expression<Func<Inventario, bool>> predicate)
        {
            throw new NotImplementedException();
        }

        public Task<int> CountAsync(Expression<Func<Inventario, bool>> predicate = null)
        {
            throw new NotImplementedException();
        }

        public Task<(IEnumerable<Inventario> Items, int TotalCount)> GetPagedAsync(
            int pageIndex, int pageSize, Expression<Func<Inventario, bool>> predicate = null,
            Func<IQueryable<Inventario>, IOrderedQueryable<Inventario>> orderBy = null)
        {
            throw new NotImplementedException();
        }
    }
}