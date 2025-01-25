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
    public class ProductoRepository : IProductoRepository
    {
        private readonly string _connectionString;
        private readonly ILogger<ProductoRepository> _logger;

        public ProductoRepository(string connectionString, ILogger<ProductoRepository> logger)
        {
            _connectionString = connectionString;
            _logger = logger;
        }

        public async Task<IEnumerable<Producto>> GetProductosBajosDeStockAsync(int stockMinimo)
        {
            try
            {
                using var connection = new SQLiteConnection(_connectionString);
                var query = "SELECT * FROM Inventario WHERE Unidades <= @stockMinimo";
                return await connection.QueryAsync<Producto>(query, new { stockMinimo });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error getting products with low stock");
                throw;
            }
        }

        public async Task ActualizarStockAsync(string codigo, int unidades, double kilos)
        {
            try
            {
                using var connection = new SQLiteConnection(_connectionString);
                var query = @"UPDATE Inventario 
                            SET Unidades = Unidades + @unidades, 
                                Kilos = Kilos + @kilos 
                            WHERE Codigo = @codigo";
                await connection.ExecuteAsync(query, new { codigo, unidades, kilos });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error updating stock");
                throw;
            }
        }

        public async Task<bool> ValidarStockDisponibleAsync(string codigo, int unidades, double kilos)
        {
            try
            {
                using var connection = new SQLiteConnection(_connectionString);
                var query = @"SELECT COUNT(*) FROM Inventario 
                            WHERE Codigo = @codigo 
                            AND Unidades >= @unidades 
                            AND Kilos >= @kilos";
                var count = await connection.ExecuteScalarAsync<int>(query, new { codigo, unidades, kilos });
                return count > 0;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error validating stock");
                throw;
            }
        }

        // Implementación de IRepository<Producto>
        public async Task<Producto> GetByIdAsync(object id)
        {
            try
            {
                using var connection = new SQLiteConnection(_connectionString);
                var query = "SELECT * FROM Inventario WHERE Codigo = @Id";
                return await connection.QueryFirstOrDefaultAsync<Producto>(query, new { Id = id });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error getting product by id");
                throw;
            }
        }

        public async Task<IEnumerable<Producto>> GetAllAsync()
        {
            try
            {
                using var connection = new SQLiteConnection(_connectionString);
                var query = "SELECT * FROM Inventario";
                return await connection.QueryAsync<Producto>(query);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error getting all products");
                throw;
            }
        }

        // Implementar los demás métodos de IRepository<Producto> según sea necesario
        public Task<IEnumerable<Producto>> FindAsync(Expression<Func<Producto, bool>> predicate)
        {
            throw new NotImplementedException();
        }

        public Task AddAsync(Producto entity)
        {
            throw new NotImplementedException();
        }

        public Task AddRangeAsync(IEnumerable<Producto> entities)
        {
            throw new NotImplementedException();
        }

        public Task UpdateAsync(Producto entity)
        {
            throw new NotImplementedException();
        }

        public Task DeleteAsync(Producto entity)
        {
            throw new NotImplementedException();
        }

        public Task DeleteRangeAsync(IEnumerable<Producto> entities)
        {
            throw new NotImplementedException();
        }

        public Task<bool> ExistsAsync(Expression<Func<Producto, bool>> predicate)
        {
            throw new NotImplementedException();
        }

        public Task<int> CountAsync(Expression<Func<Producto, bool>> predicate = null)
        {
            throw new NotImplementedException();
        }

        public Task<(IEnumerable<Producto> Items, int TotalCount)> GetPagedAsync(int pageIndex, int pageSize, Expression<Func<Producto, bool>> predicate = null, Func<IQueryable<Producto>, IOrderedQueryable<Producto>> orderBy = null)
        {
            throw new NotImplementedException();
        }
    }
}