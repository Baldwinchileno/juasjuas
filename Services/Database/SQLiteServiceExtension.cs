using System;
using System.Data.SQLite;
using System.IO;
using System.Collections.Generic;
using System.Text;
using Microsoft.Extensions.Logging;

namespace AdminSERMAC.Services.Database
{
    public class SQLiteServiceExtension
    {
        private readonly string _connectionString;
        private readonly ILogger<SQLiteServiceExtension> _logger;

        public SQLiteServiceExtension(string connectionString, ILogger<SQLiteServiceExtension> logger = null)
        {
            _connectionString = connectionString;
            _logger = logger;
        }

        public void CreateProductosTable()
        {
            try
            {
                using (var connection = new SQLiteConnection(_connectionString))
                {
                    connection.Open();
                    using (var command = new SQLiteCommand(connection))
                    {
                        // Crear tabla Productos si no existe
                        command.CommandText = @"
                    CREATE TABLE IF NOT EXISTS Productos (
                        Codigo INTEGER PRIMARY KEY,
                        Nombre TEXT NOT NULL,
                        Marca TEXT,
                        Categoria TEXT,
                        SubCategoria TEXT,
                        UnidadMedida TEXT,
                        Precio REAL DEFAULT 0
                    )";
                        command.ExecuteNonQuery();

                        _logger?.LogInformation("Tabla Productos creada o verificada exitosamente");
                    }
                }
            }
            catch (Exception ex)
            {
                _logger?.LogError(ex, "Error al crear la tabla Productos");
                throw new Exception("Error al crear la tabla Productos", ex);
            }
        }

        public void ImportarProductosDesdeCSV(string csvContent)
        {
            try
            {
                using (var connection = new SQLiteConnection(_connectionString))
                {
                    connection.Open();
                    using (var transaction = connection.BeginTransaction())
                    {
                        try
                        {
                            // Limpiamos la tabla primero
                            using (var cmdClear = new SQLiteCommand("DELETE FROM Productos", connection, transaction))
                            {
                                cmdClear.ExecuteNonQuery();
                            }

                            var lines = csvContent.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
                            int productosImportados = 0;

                            // Procesamos cada línea
                            for (int i = 1; i < lines.Length; i++) // Empezamos desde 1 para saltar los headers
                            {
                                var line = lines[i].Trim();
                                if (string.IsNullOrWhiteSpace(line)) continue;

                                var values = line.Split(',');
                                if (values.Length >= 6)
                                {
                                    try
                                    {
                                        // Limpiamos y convertimos el código
                                        var codigoStr = values[0].Trim().Replace("\"", "").Replace("'", "");
                                        if (!int.TryParse(codigoStr, out int codigo))
                                        {
                                            throw new Exception($"El código '{codigoStr}' en la línea {i + 1} no es un número válido");
                                        }

                                        var insertCmd = @"
                                    INSERT INTO Productos 
                                    (Codigo, Nombre, Marca, Categoria, SubCategoria, UnidadMedida)
                                    VALUES 
                                    (@Codigo, @Nombre, @Marca, @Categoria, @SubCategoria, @UnidadMedida)";

                                        using (var command = new SQLiteCommand(insertCmd, connection, transaction))
                                        {
                                            command.Parameters.AddWithValue("@Codigo", codigo);
                                            command.Parameters.AddWithValue("@Nombre", values[1].Trim().Replace("\"", ""));
                                            command.Parameters.AddWithValue("@Marca", values[2].Trim().Replace("\"", ""));
                                            command.Parameters.AddWithValue("@Categoria", values[3].Trim().Replace("\"", ""));
                                            command.Parameters.AddWithValue("@SubCategoria", values[4].Trim().Replace("\"", ""));
                                            command.Parameters.AddWithValue("@UnidadMedida", values[5].Trim().Replace("\"", ""));

                                            command.ExecuteNonQuery();
                                            productosImportados++;
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        throw new Exception($"Error en la línea {i + 1}: {ex.Message}");
                                    }
                                }
                                else
                                {
                                    throw new Exception($"La línea {i + 1} no tiene el formato correcto");
                                }
                            }

                            transaction.Commit();
                            _logger?.LogInformation($"Importación completada. {productosImportados} productos importados");
                        }
                        catch (Exception)
                        {
                            transaction.Rollback();
                            throw;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                _logger?.LogError(ex, "Error durante la importación de productos");
                throw;
            }
        }

        public void RecreateProductosTable()
        {
            try
            {
                using (var connection = new SQLiteConnection(_connectionString))
                {
                    connection.Open();
                    using (var transaction = connection.BeginTransaction())
                    {
                        try
                        {
                            // Eliminar tabla si existe
                            using (var command = new SQLiteCommand("DROP TABLE IF EXISTS Productos", connection, transaction))
                            {
                                command.ExecuteNonQuery();
                            }

                            // Crear tabla nueva
                            using (var command = new SQLiteCommand(@"
                        CREATE TABLE Productos (
                            Codigo INTEGER PRIMARY KEY,
                            Nombre TEXT NOT NULL,
                            Marca TEXT,
                            Categoria TEXT,
                            SubCategoria TEXT,
                            UnidadMedida TEXT,
                            Precio REAL DEFAULT 0
                        )", connection, transaction))
                            {
                                command.ExecuteNonQuery();
                            }

                            transaction.Commit();
                            MessageBox.Show("Tabla Productos recreada exitosamente");
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
                MessageBox.Show($"Error al recrear la tabla: {ex.Message}");
                throw;
            }
        }


        public List<string> GetCategoriasProductos()
        {
            var categorias = new List<string>();
            try
            {
                using (var connection = new SQLiteConnection(_connectionString))
                {
                    connection.Open();
                    using (var command = new SQLiteCommand("SELECT DISTINCT Categoria FROM Productos ORDER BY Categoria", connection))
                    {
                        using (var reader = command.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                categorias.Add(reader.GetString(0));
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                _logger?.LogError(ex, "Error al obtener categorías de productos");
                throw;
            }
            return categorias;
        }

        public List<string> GetSubCategoriasProductos(string categoria)
        {
            var subcategorias = new List<string>();
            try
            {
                using (var connection = new SQLiteConnection(_connectionString))
                {
                    connection.Open();
                    using (var command = new SQLiteCommand(
                        "SELECT DISTINCT SubCategoria FROM Productos WHERE Categoria = @categoria ORDER BY SubCategoria",
                        connection))
                    {
                        command.Parameters.AddWithValue("@categoria", categoria);
                        using (var reader = command.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                subcategorias.Add(reader.GetString(0));
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                _logger?.LogError(ex, "Error al obtener subcategorías de productos");
                throw;
            }
            return subcategorias;
        }

        public void ActualizarPrecioProducto(string codigo, decimal precio)
        {
            try
            {
                using (var connection = new SQLiteConnection(_connectionString))
                {
                    connection.Open();
                    using (var command = new SQLiteCommand(
                        "UPDATE Productos SET Precio = @precio WHERE Codigo = @codigo",
                        connection))
                    {
                        command.Parameters.AddWithValue("@precio", precio);
                        command.Parameters.AddWithValue("@codigo", codigo);
                        command.ExecuteNonQuery();
                    }
                }
            }
            catch (Exception ex)
            {
                _logger?.LogError(ex, $"Error al actualizar precio del producto {codigo}");
                throw;
            }
        }

        public void EliminarProducto(string codigo)
        {
            try
            {
                using (var connection = new SQLiteConnection(_connectionString))
                {
                    connection.Open();
                    using (var command = new SQLiteCommand(
                        "DELETE FROM Productos WHERE Codigo = @codigo",
                        connection))
                    {
                        command.Parameters.AddWithValue("@codigo", codigo);
                        command.ExecuteNonQuery();
                    }
                }
            }
            catch (Exception ex)
            {
                _logger?.LogError(ex, $"Error al eliminar el producto {codigo}");
                throw;
            }
        }
    }
}