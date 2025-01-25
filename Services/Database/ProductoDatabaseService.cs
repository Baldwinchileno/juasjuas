using System.Data.SQLite;
using Microsoft.Extensions.Logging;
using AdminSERMAC.Models;

namespace AdminSERMAC.Services.Database
{
    public interface IProductoDatabaseService
    {
        Task<Producto?> GetByCodigo(string codigo);
        Task<IEnumerable<Producto>> GetAllAsync();
        Task<bool> CreateAsync(Producto producto);
        Task<bool> UpdateAsync(Producto producto);
        Task<bool> UpdatePrecioAsync(string codigo, double precio);
        Task<double> GetPrecioAsync(string codigo);
        Task<bool> DeleteAsync(string codigo);
        Task<bool> ImportarProductosAsync(string csvContent);
    }

    public class ProductoDatabaseService : BaseSQLiteService, IProductoDatabaseService
    {
        private const string TableName = "Productos";

        public ProductoDatabaseService(ILogger<ProductoDatabaseService> logger, string connectionString)
            : base(logger, connectionString)
        {
            EnsureTableExists();
        }

        private void EnsureTableExists()
        {
            const string createTableSql = @"
                CREATE TABLE IF NOT EXISTS Productos (
                    Codigo TEXT PRIMARY KEY,
                    Nombre TEXT NOT NULL,
                    Marca TEXT,
                    Categoria TEXT,
                    SubCategoria TEXT,
                    UnidadMedida TEXT,
                    Precio REAL DEFAULT 0
                )";

            ExecuteInTransactionAsync(async (connection, transaction) =>
            {
                using var command = new SQLiteCommand(createTableSql, connection, transaction);
                await command.ExecuteNonQueryAsync();
            }).Wait();
        }

        public async Task<Producto?> GetByCodigo(string codigo)
        {
            return await ExecuteInTransactionAsync(async (connection, transaction) =>
            {
                const string sql = @"
                    SELECT p.*, i.Unidades, i.Kilos, i.FechaMasAntigua, i.FechaMasNueva
                    FROM Productos p
                    LEFT JOIN Inventario i ON p.Codigo = i.Codigo
                    WHERE p.Codigo = @codigo";

                using var command = new SQLiteCommand(sql, connection, transaction);
                command.Parameters.AddWithValue("@codigo", codigo);

                using var reader = await command.ExecuteReaderAsync();
                if (await reader.ReadAsync())
                {
                    return new Producto
                    {
                        Codigo = reader["Codigo"].ToString(),
                        Nombre = reader["Nombre"].ToString(),
                        Marca = reader["Marca"]?.ToString(),
                        Categoria = reader["Categoria"]?.ToString(),
                        SubCategoria = reader["SubCategoria"]?.ToString(),
                        UnidadMedida = reader["UnidadMedida"]?.ToString(),
                        Precio = reader["Precio"] != DBNull.Value ? Convert.ToDouble(reader["Precio"]) : 0,
                        Unidades = reader["Unidades"] != DBNull.Value ? Convert.ToInt32(reader["Unidades"]) : 0,
                        Kilos = reader["Kilos"] != DBNull.Value ? Convert.ToDouble(reader["Kilos"]) : 0,
                        FechaMasAntigua = reader["FechaMasAntigua"]?.ToString(),
                        FechaMasNueva = reader["FechaMasNueva"]?.ToString()
                    };
                }
                return null;
            });
        }

        public async Task<IEnumerable<Producto>> GetAllAsync()
        {
            return await ExecuteInTransactionAsync(async (connection, transaction) =>
            {
                var productos = new List<Producto>();
                const string sql = @"
                    SELECT p.*, i.Unidades, i.Kilos, i.FechaMasAntigua, i.FechaMasNueva
                    FROM Productos p
                    LEFT JOIN Inventario i ON p.Codigo = i.Codigo";

                using var command = new SQLiteCommand(sql, connection, transaction);
                using var reader = await command.ExecuteReaderAsync();

                while (await reader.ReadAsync())
                {
                    productos.Add(new Producto
                    {
                        Codigo = reader["Codigo"].ToString(),
                        Nombre = reader["Nombre"].ToString(),
                        Marca = reader["Marca"]?.ToString(),
                        Categoria = reader["Categoria"]?.ToString(),
                        SubCategoria = reader["SubCategoria"]?.ToString(),
                        UnidadMedida = reader["UnidadMedida"]?.ToString(),
                        Precio = reader["Precio"] != DBNull.Value ? Convert.ToDouble(reader["Precio"]) : 0,
                        Unidades = reader["Unidades"] != DBNull.Value ? Convert.ToInt32(reader["Unidades"]) : 0,
                        Kilos = reader["Kilos"] != DBNull.Value ? Convert.ToDouble(reader["Kilos"]) : 0,
                        FechaMasAntigua = reader["FechaMasAntigua"]?.ToString(),
                        FechaMasNueva = reader["FechaMasNueva"]?.ToString()
                    });
                }
                return productos;
            });
        }

        public async Task<bool> CreateAsync(Producto producto)
        {
            return await ExecuteInTransactionAsync(async (connection, transaction) =>
            {
                const string sql = @"
                    INSERT INTO Productos (
                        Codigo, 
                        Nombre, 
                        Marca, 
                        Categoria, 
                        SubCategoria, 
                        UnidadMedida, 
                        Precio
                    ) VALUES (
                        @codigo,
                        @nombre,
                        @marca,
                        @categoria,
                        @subcategoria,
                        @unidadMedida,
                        @precio
                    )";

                using var command = new SQLiteCommand(sql, connection, transaction);
                command.Parameters.AddWithValue("@codigo", producto.Codigo);
                command.Parameters.AddWithValue("@nombre", producto.Nombre);
                command.Parameters.AddWithValue("@marca", (object?)producto.Marca ?? DBNull.Value);
                command.Parameters.AddWithValue("@categoria", (object?)producto.Categoria ?? DBNull.Value);
                command.Parameters.AddWithValue("@subcategoria", (object?)producto.SubCategoria ?? DBNull.Value);
                command.Parameters.AddWithValue("@unidadMedida", (object?)producto.UnidadMedida ?? DBNull.Value);
                command.Parameters.AddWithValue("@precio", producto.Precio);

                return await command.ExecuteNonQueryAsync() > 0;
            });
        }

        public async Task<bool> UpdateAsync(Producto producto)
        {
            return await ExecuteInTransactionAsync(async (connection, transaction) =>
            {
                const string sql = @"
                    UPDATE Productos 
                    SET Nombre = @nombre,
                        Marca = @marca,
                        Categoria = @categoria,
                        SubCategoria = @subcategoria,
                        UnidadMedida = @unidadMedida,
                        Precio = @precio
                    WHERE Codigo = @codigo";

                using var command = new SQLiteCommand(sql, connection, transaction);
                command.Parameters.AddWithValue("@codigo", producto.Codigo);
                command.Parameters.AddWithValue("@nombre", producto.Nombre);
                command.Parameters.AddWithValue("@marca", (object?)producto.Marca ?? DBNull.Value);
                command.Parameters.AddWithValue("@categoria", (object?)producto.Categoria ?? DBNull.Value);
                command.Parameters.AddWithValue("@subcategoria", (object?)producto.SubCategoria ?? DBNull.Value);
                command.Parameters.AddWithValue("@unidadMedida", (object?)producto.UnidadMedida ?? DBNull.Value);
                command.Parameters.AddWithValue("@precio", producto.Precio);

                return await command.ExecuteNonQueryAsync() > 0;
            });
        }

        public async Task<bool> UpdatePrecioAsync(string codigo, double precio)
        {
            return await ExecuteInTransactionAsync(async (connection, transaction) =>
            {
                const string sql = "UPDATE Productos SET Precio = @precio WHERE Codigo = @codigo";
                using var command = new SQLiteCommand(sql, connection, transaction);
                command.Parameters.AddWithValue("@codigo", codigo);
                command.Parameters.AddWithValue("@precio", precio);

                return await command.ExecuteNonQueryAsync() > 0;
            });
        }

        public async Task<double> GetPrecioAsync(string codigo)
        {
            return await ExecuteInTransactionAsync(async (connection, transaction) =>
            {
                const string sql = "SELECT Precio FROM Productos WHERE Codigo = @codigo";
                using var command = new SQLiteCommand(sql, connection, transaction);
                command.Parameters.AddWithValue("@codigo", codigo);

                var result = await command.ExecuteScalarAsync();
                return result != DBNull.Value ? Convert.ToDouble(result) : 0;
            });
        }

        public async Task<bool> DeleteAsync(string codigo)
        {
            return await ExecuteInTransactionAsync(async (connection, transaction) =>
            {
                const string sql = "DELETE FROM Productos WHERE Codigo = @codigo";
                using var command = new SQLiteCommand(sql, connection, transaction);
                command.Parameters.AddWithValue("@codigo", codigo);

                return await command.ExecuteNonQueryAsync() > 0;
            });
        }

        public async Task<bool> ImportarProductosAsync(string csvContent)
        {
            try
            {
                var lines = csvContent.Split(new[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries);
                foreach (var line in lines.Skip(1)) // Skip header
                {
                    var values = line.Split(',');
                    if (values.Length >= 7)
                    {
                        var producto = new Producto
                        {
                            Codigo = values[0].Trim(),
                            Nombre = values[1].Trim(),
                            Marca = values[2].Trim(),
                            Categoria = values[3].Trim(),
                            SubCategoria = values[4].Trim(),
                            UnidadMedida = values[5].Trim(),
                            Precio = double.TryParse(values[6].Trim(), out double precio) ? precio : 0
                        };

                        await CreateAsync(producto);
                    }
                }
                return true;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error importing products from CSV");
                return false;
            }
        }
    }
}
