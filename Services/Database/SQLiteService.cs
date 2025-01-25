using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SQLite;
using System.Net.NetworkInformation;
using AdminSERMAC.Models;
using System.Collections.Generic;


namespace AdminSERMAC.Services
{
    public class SQLiteService
    {
        public string connectionString { get; private set; }
        private object rut;
        private readonly string _connectionString;
        private readonly ILogger<SQLiteService> _logger;

        public SQLiteService(ILogger<SQLiteService> logger)
        {
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
            connectionString = "Data Source=AdminSERMAC.db;Version=3;";

            EnsureTablesExist();
            EnsureColumnsExist();
        }


        public SQLiteConnection GetConnection()
        {
            return new SQLiteConnection(connectionString); // Asegúrate de que "connectionString" esté definido en la clase.
        }

        private void EnsureTablesExist()
        {
            EnsureClientesTableExists();
            EnsureConfiguracionTableExists();
            EnsureProductosTableExists();
            EnsureInventarioTableExists();
            EnsureProveedoresTableExists();
            EnsureVentasTableExists();
            EnsureHistorialMovimientosTableExists(); // Agregar esta línea
        }

        public List<Producto> GetProductos()
        {
            List<Producto> productos = new List<Producto>();

            try
            {
                using (var connection = new SQLiteConnection(connectionString))
                {
                    connection.Open();
                    string query = @"
                SELECT 
                    p.Codigo,
                    p.Nombre,
                    p.Marca,
                    i.Kilos,
                    i.Unidades,
                    p.Categoria,
                    p.SubCategoria,
                    p.Precio,
                    p.UnidadMedida,
                    i.FechaMasAntigua,
                    i.FechaMasNueva
                FROM Productos p
                LEFT JOIN Inventario i ON p.Codigo = i.Codigo";

                    using (var command = new SQLiteCommand(query, connection))
                    using (var reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            productos.Add(new Producto
                            {
                                Codigo = reader["Codigo"].ToString(),
                                Nombre = reader["Nombre"].ToString(),
                                Marca = reader["Marca"].ToString(),        // Asegurarnos que esto está presente
                                Kilos = reader["Kilos"] != DBNull.Value ? Convert.ToDouble(reader["Kilos"]) : 0,
                                Unidades = reader["Unidades"] != DBNull.Value ? Convert.ToInt32(reader["Unidades"]) : 0,
                                Categoria = reader["Categoria"].ToString(),
                                SubCategoria = reader["SubCategoria"].ToString(),
                                Precio = reader["Precio"] != DBNull.Value ? Convert.ToDouble(reader["Precio"]) : 0,
                                UnidadMedida = reader["UnidadMedida"].ToString(),
                                FechaMasAntigua = reader["FechaMasAntigua"]?.ToString(),
                                FechaMasNueva = reader["FechaMasNueva"]?.ToString()
                            });
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al obtener productos");
                throw;
            }

            return productos;
        }

        public void EnsureHistorialMovimientosTableExists()
        {
            using (var connection = new SQLiteConnection(connectionString))
            {
                connection.Open();
                EnableForeignKeys(connection);
                string createTableQuery = @"
            CREATE TABLE IF NOT EXISTS HistorialMovimientos (
                Id INTEGER PRIMARY KEY AUTOINCREMENT,
                RUT TEXT NOT NULL,
                Tipo TEXT NOT NULL,
                Monto REAL NOT NULL,
                Fecha TEXT NOT NULL,
                FOREIGN KEY (RUT) REFERENCES Clientes(RUT)
            );";

                using (var command = new SQLiteCommand(createTableQuery, connection))
                {
                    command.ExecuteNonQuery();
                }
            }
        }

        public void ActualizarCliente(Cliente cliente)
        {
            using (var connection = GetConnection())
            {
                connection.Open();
                string updateQuery = @"
            UPDATE Clientes
            SET Nombre = @Nombre,
                Direccion = @Direccion,
                Giro = @Giro,
                Deuda = @Deuda
            WHERE RUT = @RUT";

                using (var command = new SQLiteCommand(updateQuery, connection))
                {
                    command.Parameters.AddWithValue("@RUT", cliente.RUT);
                    command.Parameters.AddWithValue("@Nombre", cliente.Nombre);
                    command.Parameters.AddWithValue("@Direccion", cliente.Direccion);
                    command.Parameters.AddWithValue("@Giro", cliente.Giro);
                    command.Parameters.AddWithValue("@Deuda", cliente.Deuda);
                    command.ExecuteNonQuery();
                }
            }
        }



        public List<string> GetCategorias()
        {
            var categorias = new List<string>();

            using (var connection = new SQLiteConnection(connectionString))
            {
                connection.Open();

                string query = "SELECT DISTINCT Categoria FROM Inventario WHERE Categoria IS NOT NULL ORDER BY Categoria";
                using (var command = new SQLiteCommand(query, connection))
                using (var reader = command.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        categorias.Add(reader["Categoria"].ToString());
                    }
                }
            }

            return categorias;
        }

        // Método para obtener las subcategorías de una categoría específica
        public List<string> GetSubCategorias(string categoria)
        {
            var subcategorias = new List<string>();

            using (var connection = new SQLiteConnection(connectionString))
            {
                connection.Open();

                string query = "SELECT DISTINCT SubCategoria FROM Inventario WHERE Categoria = @categoria AND SubCategoria IS NOT NULL ORDER BY SubCategoria";
                using (var command = new SQLiteCommand(query, connection))
                {
                    command.Parameters.AddWithValue("@categoria", categoria);

                    using (var reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            subcategorias.Add(reader["SubCategoria"].ToString());
                        }
                    }
                }
            }

            return subcategorias;
        }
    

        public void EnsureColumnsExist()
        {
            using (var connection = new SQLiteConnection(connectionString))
            {
                connection.Open();
                EnableForeignKeys(connection);

                // Verificar y agregar columna Deuda a la tabla Clientes si no existe
                if (!ColumnExists(connection, "Clientes", "Deuda"))
                {
                    using (var command = new SQLiteCommand("ALTER TABLE Clientes ADD COLUMN Deuda REAL DEFAULT 0", connection))
                    {
                        command.ExecuteNonQuery();
                    }
                }

                // Verificar y agregar columna Precio a la tabla Productos si no existe
                if (!ColumnExists(connection, "Productos", "Precio"))
                {
                    using (var command = new SQLiteCommand("ALTER TABLE Productos ADD COLUMN Precio REAL", connection))
                    {
                        command.ExecuteNonQuery();
                    }
                }

                // Verificar y agregar columna PagadoConCredito a la tabla Ventas si no existe
                if (!ColumnExists(connection, "Ventas", "PagadoConCredito"))
                {
                    using (var command = new SQLiteCommand("ALTER TABLE Ventas ADD COLUMN PagadoConCredito INTEGER DEFAULT 0", connection))
                    {
                        command.ExecuteNonQuery();
                    }
                }

                // Verificar y agregar columna RUT a la tabla Ventas si no existe
                if (!ColumnExists(connection, "Ventas", "RUT"))
                {
                    using (var command = new SQLiteCommand("ALTER TABLE Ventas ADD COLUMN RUT TEXT", connection))
                    {
                        command.ExecuteNonQuery();
                    }
                }
            }
        }

        private bool ColumnExists(SQLiteConnection connection, string tableName, string columnName)
        {
            var command = connection.CreateCommand();
            command.CommandText = $"PRAGMA table_info({tableName})";
            using (var reader = command.ExecuteReader())
            {
                while (reader.Read())
                {
                    if (reader["name"].ToString() == columnName)
                    {
                        return true;
                    }
                }
            }
            return false;
        }


        private void EnableForeignKeys(SQLiteConnection connection)
        {
            using (var command = new SQLiteCommand("PRAGMA foreign_keys = ON;", connection))
            {
                command.ExecuteNonQuery();
            }
        }

        public void EnsureClientesTableExists()
        {
            using (var connection = new SQLiteConnection(connectionString))
            {
                connection.Open();
                EnableForeignKeys(connection);
                using (var command = new SQLiteCommand(connection))
                {
                    command.CommandText = @"
                        CREATE TABLE IF NOT EXISTS Clientes (
                            RUT TEXT PRIMARY KEY,
                            Nombre TEXT,
                            Direccion TEXT,
                            Giro TEXT,
                            Deuda REAL DEFAULT 0
                        )";
                    command.ExecuteNonQuery();
                }
            }
        }

        public void EnsureConfiguracionTableExists()
        {
            using (var connection = new SQLiteConnection(connectionString))
            {
                connection.Open();
                EnableForeignKeys(connection);
                string createTableQuery = @"
                    CREATE TABLE IF NOT EXISTS Configuracion (
                        Clave TEXT PRIMARY KEY,
                        Valor TEXT NOT NULL
                    );";

                using (var command = new SQLiteCommand(createTableQuery, connection))
                {
                    command.ExecuteNonQuery();
                }

                string insertDefaultQuery = @"
                    INSERT OR IGNORE INTO Configuracion (Clave, Valor) VALUES ('UltimoNumeroGuia', '0');
                    INSERT OR IGNORE INTO Configuracion (Clave, Valor) VALUES ('UltimoNumeroCompra', '0');";

                using (var command = new SQLiteCommand(insertDefaultQuery, connection))
                {
                    command.ExecuteNonQuery();
                }
            }
        }

        public void EnsureInventarioTableExists()
        {
            using (var connection = new SQLiteConnection(connectionString))
            {
                connection.Open();
                var command = connection.CreateCommand();
                command.CommandText = @"
            CREATE TABLE IF NOT EXISTS Inventario (
                Id INTEGER PRIMARY KEY AUTOINCREMENT,
                Codigo TEXT NOT NULL,
                Unidades INTEGER NOT NULL,
                Kilos REAL NOT NULL,
                FechaCompra TEXT NOT NULL,
                FechaRegistro TEXT NOT NULL,
                FechaVencimiento TEXT
            )";
                command.ExecuteNonQuery();

                // Verificar si la columna FechaVencimiento existe, si no, agregarla
                if (!ColumnExists(connection, "Inventario", "FechaVencimiento"))
                {
                    command.CommandText = "ALTER TABLE Inventario ADD COLUMN FechaVencimiento TEXT";
                    command.ExecuteNonQuery();
                }
            }
        }

        public void EnsureProductosTableExists()
        {
            using (var connection = new SQLiteConnection(connectionString))
            {
                connection.Open();
                EnableForeignKeys(connection);
                string createTableQuery = @"
            CREATE TABLE IF NOT EXISTS Productos (
                Codigo TEXT PRIMARY KEY,
                Nombre TEXT NOT NULL,
                Marca TEXT,
                Categoria TEXT,
                SubCategoria TEXT,
                UnidadMedida TEXT,
                Precio REAL DEFAULT 0
            );";

                using (var command = new SQLiteCommand(createTableQuery, connection))
                {
                    command.ExecuteNonQuery();
                }
            }
        }
        public void EnsureProveedoresTableExists()
        {
            using (var connection = new SQLiteConnection(connectionString))
            {
                connection.Open();
                EnableForeignKeys(connection);
                string createTableQuery = @"
                    CREATE TABLE IF NOT EXISTS Proveedores (
                        Id INTEGER PRIMARY KEY AUTOINCREMENT,
                        Nombre TEXT NOT NULL,
                        Vendedor TEXT NOT NULL
                    );";

                using (var command = new SQLiteCommand(createTableQuery, connection))
                {
                    command.ExecuteNonQuery();
                }
            }
        }

        public void EnsureVentasTableExists()
        {
            using (var connection = new SQLiteConnection(connectionString))
            {
                connection.Open();
                EnableForeignKeys(connection);
                string createTableQuery = @"
                    CREATE TABLE IF NOT EXISTS Ventas (
                        NumeroGuia INTEGER,
                        CodigoProducto TEXT,
                        Descripcion TEXT,
                        Bandejas INTEGER,
                        KilosNeto REAL,
                        FechaVenta TEXT,
                        PagadoConCredito INTEGER DEFAULT 0,
                        RUT TEXT,
                        PRIMARY KEY (NumeroGuia, CodigoProducto),
                        FOREIGN KEY (CodigoProducto) REFERENCES Productos(Codigo),
                        FOREIGN KEY (RUT) REFERENCES Clientes(RUT)
                    );";

                using (var command = new SQLiteCommand(createTableQuery, connection))
                {
                    command.ExecuteNonQuery();
                }
            }
        }

        public List<string> GetProveedores()
        {
            List<string> proveedores = new List<string>();

            using (var connection = new SQLiteConnection(connectionString))
            {
                connection.Open();
                EnableForeignKeys(connection);
                string query = "SELECT DISTINCT Nombre FROM Proveedores";

                using (var command = new SQLiteCommand(query, connection))
                {
                    using (var reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            proveedores.Add(reader["Nombre"].ToString());
                        }
                    }
                }
            }

            return proveedores;
        }

        public List<string> GetVendedores()
        {
            List<string> vendedores = new List<string>();

            using (var connection = new SQLiteConnection(connectionString))
            {
                connection.Open();
                EnableForeignKeys(connection);
                string query = "SELECT DISTINCT Vendedor FROM Proveedores";

                using (var command = new SQLiteCommand(query, connection))
                {
                    using (var reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            vendedores.Add(reader["Vendedor"].ToString());
                        }
                    }
                }
            }

            return vendedores;
        }

        public List<Venta> GetVentasPorCliente(string rut)
        {
            var ventas = new List<Venta>();
            using (var connection = new SQLiteConnection(connectionString))
            {
                connection.Open();
                EnableForeignKeys(connection);
                using (var command = new SQLiteCommand(connection))
                {
                    command.CommandText = @"
                SELECT v.NumeroGuia, v.CodigoProducto, v.Descripcion, v.Bandejas, v.KilosNeto, v.FechaVenta, (v.KilosNeto * p.Precio) AS Total, c.Deuda
                FROM Ventas v
                JOIN Clientes c ON c.RUT = @RUT
                JOIN Productos p ON p.Codigo = v.CodigoProducto
                WHERE c.RUT = @RUT AND v.PagadoConCredito = 1";
                    command.Parameters.AddWithValue("@RUT", rut);
                    using (var reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            ventas.Add(new Venta
                            {
                                NumeroGuia = Convert.ToInt32(reader["NumeroGuia"]),
                                CodigoProducto = reader["CodigoProducto"].ToString(),
                                Descripcion = reader["Descripcion"].ToString(),
                                Bandejas = Convert.ToInt32(reader["Bandejas"]),
                                KilosNeto = Convert.ToDouble(reader["KilosNeto"]),
                                FechaVenta = reader["FechaVenta"].ToString(),
                                Total = Convert.ToDouble(reader["Total"]),
                                Deuda = Convert.ToDouble(reader["Deuda"]) // Asegúrate de agregar esta línea
                            });
                        }
                    }
                }
            }
            return ventas;
        }


        public int GetUltimoNumeroGuia()
        {
            int ultimoNumero = 0;

            using (var connection = new SQLiteConnection(connectionString))
            {
                connection.Open();
                EnableForeignKeys(connection);
                string query = "SELECT Valor FROM Configuracion WHERE Clave = 'UltimoNumeroGuia'";
                using (var command = new SQLiteCommand(query, connection))
                {
                    object result = command.ExecuteScalar();
                    if (result != null)
                    {
                        int.TryParse(result.ToString(), out ultimoNumero);
                    }
                }
            }

            return ultimoNumero;
        }

        public void IncrementarNumeroGuia()
        {
            using (var connection = new SQLiteConnection(connectionString))
            {
                connection.Open();
                EnableForeignKeys(connection);
                string query = @"
                    UPDATE Configuracion
                    SET Valor = CAST((CAST(Valor AS INTEGER) + 1) AS TEXT)
                    WHERE Clave = 'UltimoNumeroGuia';";

                using (var command = new SQLiteCommand(query, connection))
                {
                    command.ExecuteNonQuery();
                }
            }
        }

        public int GetUltimoNumeroCompra()
        {
            int ultimoNumero = 0;

            using (var connection = new SQLiteConnection(connectionString))
            {
                connection.Open();
                EnableForeignKeys(connection);
                string query = "SELECT Valor FROM Configuracion WHERE Clave = 'UltimoNumeroCompra'";
                using (var command = new SQLiteCommand(query, connection))
                {
                    object result = command.ExecuteScalar();
                    if (result != null)
                    {
                        int.TryParse(result.ToString(), out ultimoNumero);
                    }
                }
            }

            return ultimoNumero;
        }

        public void ImportarProductosDesdeCSV(string csvContent)
        {
            using (var connection = new SQLiteConnection(connectionString))
            {
                connection.Open();
                using (var transaction = connection.BeginTransaction())
                {
                    try
                    {
                        // Limpiamos las tablas
                        using (var command = new SQLiteCommand(@"
                    DELETE FROM Inventario;
                    DELETE FROM Productos;
                    DELETE FROM sqlite_sequence WHERE name IN ('Inventario');", connection, transaction))
                        {
                            command.ExecuteNonQuery();
                        }

                        // Procesamos el CSV
                        string[] lines = csvContent.Split(new[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries);

                        // Saltamos la primera línea (headers)
                        for (int i = 1; i < lines.Length; i++)
                        {
                            string line = lines[i].Trim();
                            if (string.IsNullOrWhiteSpace(line)) continue;

                            string[] values = line.Split(',');
                            if (values.Length >= 6) // Verificamos que tengamos todos los campos
                            {
                                string codigo = values[0].Trim();
                                string nombre = values[1].Trim();
                                string marca = values[2].Trim();
                                string categoria = !string.IsNullOrWhiteSpace(values[3]) ? values[3].Trim() : null;
                                string subcategoria = !string.IsNullOrWhiteSpace(values[4]) ? values[4].Trim() : null;
                                string unidadMedida = !string.IsNullOrWhiteSpace(values[5]) ? values[5].Trim() : "Kg";

                                // Insertamos en la tabla Productos
                                using (var command = new SQLiteCommand(@"
                            INSERT INTO Productos (Codigo, Nombre, Marca, Categoria, SubCategoria, UnidadMedida)
                            VALUES (@codigo, @nombre, @marca, @categoria, @subcategoria, @unidadMedida)",
                                    connection, transaction))
                                {
                                    command.Parameters.AddWithValue("@codigo", codigo);
                                    command.Parameters.AddWithValue("@nombre", nombre);
                                    command.Parameters.AddWithValue("@marca", marca);
                                    command.Parameters.AddWithValue("@categoria", (object)categoria ?? DBNull.Value);
                                    command.Parameters.AddWithValue("@subcategoria", (object)subcategoria ?? DBNull.Value);
                                    command.Parameters.AddWithValue("@unidadMedida", unidadMedida);

                                    command.ExecuteNonQuery();
                                }

                                // Creamos un registro inicial en Inventario
                                using (var command = new SQLiteCommand(@"
                            INSERT INTO Inventario (Codigo, Unidades, Kilos, FechaMasAntigua, FechaMasNueva)
                            VALUES (@codigo, 0, 0, @fecha, @fecha)",
                                    connection, transaction))
                                {
                                    string fechaActual = DateTime.Now.ToString("yyyy-MM-dd");
                                    command.Parameters.AddWithValue("@codigo", codigo);
                                    command.Parameters.AddWithValue("@fecha", fechaActual);
                                    command.ExecuteNonQuery();
                                }
                            }
                        }

                        transaction.Commit();
                        _logger.LogInformation($"Importación completada exitosamente. Se procesaron {lines.Length - 1} líneas.");
                    }
                    catch (Exception ex)
                    {
                        transaction.Rollback();
                        _logger.LogError(ex, "Error durante la importación");
                        throw;
                    }
                }
            }
        }

        public async Task<bool> BackupDatabaseAsync(string backupPath)
        {
            try
            {
                using var source = new SQLiteConnection(_connectionString);
                using var destination = new SQLiteConnection($"Data Source={backupPath}");
                await source.OpenAsync();
                source.BackupDatabase(destination, "main", "main", -1, null, 0);
                return true;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error while creating a database backup");
                return false;
            }
        }

        public async Task<double?> GetUltimoPrecioCompraAsync(string codigoProducto)
        {
            const string sql = @"
        SELECT PrecioUnitario 
        FROM DetallesCompra 
        WHERE Codigo_Producto = @codigoProducto 
        ORDER BY Id DESC 
        LIMIT 1";

            using var connection = new SQLiteConnection(connectionString);
            await connection.OpenAsync();
            using var command = new SQLiteCommand(sql, connection);
            command.Parameters.AddWithValue("@codigoProducto", codigoProducto);

            var result = await command.ExecuteScalarAsync();
            return result == DBNull.Value ? null : Convert.ToDouble(result);
        }

        public async Task<double?> GetUltimoPrecioVentaAsync(string codigoProducto, string rut)
        {
            const string sql = @"
        SELECT Precio 
        FROM Ventas 
        WHERE CodigoProducto = @codigoProducto 
        AND RUT = @rut 
        ORDER BY NumeroGuia DESC 
        LIMIT 1";

            using var connection = new SQLiteConnection(connectionString);
            await connection.OpenAsync();
            using var command = new SQLiteCommand(sql, connection);
            command.Parameters.AddWithValue("@codigoProducto", codigoProducto);
            command.Parameters.AddWithValue("@rut", rut);

            var result = await command.ExecuteScalarAsync();
            return result == DBNull.Value ? null : Convert.ToDouble(result);
        }


        // Método AddProducto ya existente en la clase
        public void AddProducto(string codigo, string producto, int unidades, double kilos, string fechaCompra, string fechaRegistro, string fechaVencimiento)
        {
            using (var connection = new SQLiteConnection(connectionString))
            {
                connection.Open();
                using (var transaction = connection.BeginTransaction())
                {
                    try
                    {
                        // Verificar si el producto existe en el inventario
                        var checkCommand = new SQLiteCommand(
                            "SELECT COUNT(*) FROM Inventario WHERE Codigo = @codigo",
                            connection,
                            transaction);
                        checkCommand.Parameters.AddWithValue("@codigo", codigo);

                        int exists = Convert.ToInt32(checkCommand.ExecuteScalar());

                        if (exists > 0)
                        {
                            // Actualizar inventario existente
                            var updateCommand = new SQLiteCommand(@"
                        UPDATE Inventario 
                        SET Unidades = Unidades + @unidades,
                            Kilos = Kilos + @kilos,
                            FechaMasNueva = @fechaRegistro,
                            FechaVencimiento = @fechaVencimiento
                        WHERE Codigo = @codigo",
                                connection,
                                transaction);

                            updateCommand.Parameters.AddWithValue("@codigo", codigo);
                            updateCommand.Parameters.AddWithValue("@unidades", unidades);
                            updateCommand.Parameters.AddWithValue("@kilos", kilos);
                            updateCommand.Parameters.AddWithValue("@fechaRegistro", fechaRegistro);
                            updateCommand.Parameters.AddWithValue("@fechaVencimiento", fechaVencimiento);
                            updateCommand.ExecuteNonQuery();
                        }
                        else
                        {
                            // Insertar nuevo registro
                            var insertCommand = new SQLiteCommand(@"
                        INSERT INTO Inventario (
                            Codigo, 
                            Producto, 
                            Unidades, 
                            Kilos, 
                            FechaMasAntigua,
                            FechaMasNueva,
                            FechaVencimiento
                        ) VALUES (
                            @codigo,
                            @producto,
                            @unidades,
                            @kilos,
                            @fechaCompra,
                            @fechaRegistro,
                            @fechaVencimiento
                        )",
                                connection,
                                transaction);

                            insertCommand.Parameters.AddWithValue("@codigo", codigo);
                            insertCommand.Parameters.AddWithValue("@producto", producto);
                            insertCommand.Parameters.AddWithValue("@unidades", unidades);
                            insertCommand.Parameters.AddWithValue("@kilos", kilos);
                            insertCommand.Parameters.AddWithValue("@fechaCompra", fechaCompra);
                            insertCommand.Parameters.AddWithValue("@fechaRegistro", fechaRegistro);
                            insertCommand.Parameters.AddWithValue("@fechaVencimiento", fechaVencimiento);
                            insertCommand.ExecuteNonQuery();
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



        public void IncrementarNumeroCompra()
        {
            using (var connection = new SQLiteConnection(connectionString))
            {
                connection.Open();
                EnableForeignKeys(connection);
                string query = @"
                    UPDATE Configuracion
                    SET Valor = CAST((CAST(Valor AS INTEGER) + 1) AS TEXT)
                    WHERE Clave = 'UltimoNumeroCompra';";

                using (var command = new SQLiteCommand(query, connection))
                {
                    command.ExecuteNonQuery();
                }
            }
        }

        public Producto GetProductoPorCodigo(string codigo)
        {
            Producto producto = null;

            using (var connection = new SQLiteConnection(connectionString))
            {
                connection.Open();
                string query = @"
        SELECT p.*, i.Unidades, i.Kilos, i.FechaMasAntigua, i.FechaMasNueva
        FROM Productos p
        LEFT JOIN Inventario i ON p.Codigo = i.Codigo
        WHERE p.Codigo = @codigo";

                using (var command = new SQLiteCommand(query, connection))
                {
                    command.Parameters.AddWithValue("@codigo", codigo);

                    using (var reader = command.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            producto = new Producto
                            {
                                Codigo = reader["Codigo"].ToString(),
                                Nombre = reader["Nombre"].ToString(),
                                Marca = reader["Marca"].ToString(),  // Asegurarnos que esto está presente
                                Unidades = reader["Unidades"] != DBNull.Value ? Convert.ToInt32(reader["Unidades"]) : 0,
                                Kilos = reader["Kilos"] != DBNull.Value ? Convert.ToDouble(reader["Kilos"]) : 0,
                                FechaMasAntigua = reader["FechaMasAntigua"]?.ToString(),
                                FechaMasNueva = reader["FechaMasNueva"]?.ToString()
                            };
                        }
                    }
                }
            }

            return producto;
        }

        public void ActualizarDeudaCliente(string rut, double monto)
        {
            using (var connection = new SQLiteConnection(connectionString))
            {
                connection.Open();
                EnableForeignKeys(connection);
                using (var command = new SQLiteCommand(connection))
                {
                    command.CommandText = "UPDATE Clientes SET Deuda = Deuda + @Monto WHERE RUT = @RUT";
                    command.Parameters.AddWithValue("@Monto", monto);
                    command.Parameters.AddWithValue("@RUT", rut);
                    command.ExecuteNonQuery();
                }
            }
        }

        public void DescontarInventario(int sucursalId, string codigoProducto, int unidades, double kilos, SQLiteTransaction transaction = null)
        {
            var query = @"
        UPDATE InventarioSucursal
        SET Unidades = Unidades - @Unidades, Kilos = Kilos - @Kilos
        WHERE SucursalId = @SucursalId AND Codigo = @CodigoProducto";

            using (var command = new SQLiteCommand(query, transaction?.Connection))
            {
                command.Parameters.AddWithValue("@Unidades", unidades);
                command.Parameters.AddWithValue("@Kilos", kilos);
                command.Parameters.AddWithValue("@SucursalId", sucursalId);
                command.Parameters.AddWithValue("@CodigoProducto", codigoProducto);

                if (transaction != null)
                {
                    command.Transaction = transaction;
                }

                command.ExecuteNonQuery();
            }
        }

        public string GetDatabaseFilePath()
        {
            var builder = new SQLiteConnectionStringBuilder(connectionString);
            return builder.DataSource;
        }


        public void AgregarVenta(int sucursalId, string codigoProducto, string descripcion, int unidades, double kilos, string cliente, SQLiteTransaction transaction = null)
        {
            var query = @"
        INSERT INTO Ventas (SucursalId, CodigoProducto, Descripcion, Unidades, Kilos, Cliente)
        VALUES (@SucursalId, @CodigoProducto, @Descripcion, @Unidades, @Kilos, @Cliente)";

            using (var command = new SQLiteCommand(query, transaction?.Connection))
            {
                command.Parameters.AddWithValue("@SucursalId", sucursalId);
                command.Parameters.AddWithValue("@CodigoProducto", codigoProducto);
                command.Parameters.AddWithValue("@Descripcion", descripcion);
                command.Parameters.AddWithValue("@Unidades", unidades);
                command.Parameters.AddWithValue("@Kilos", kilos);
                command.Parameters.AddWithValue("@Cliente", cliente);

                if (transaction != null)
                {
                    command.Transaction = transaction;
                }

                command.ExecuteNonQuery();
            }
        }



        public DataTable GetInventario()
        {
            DataTable inventario = new DataTable();

            using (var connection = new SQLiteConnection(connectionString))
            {
                connection.Open();
                EnableForeignKeys(connection);

                string query = "SELECT * FROM Inventario";
                using (var command = new SQLiteCommand(query, connection))
                {
                    using (var adapter = new SQLiteDataAdapter(command))
                    {
                        adapter.Fill(inventario);
                    }
                }
            }

            return inventario;
        }

        public double GetPrecioProducto(string codigo)
        {
            using (var connection = new SQLiteConnection(connectionString))
            {
                connection.Open();
                var command = new SQLiteCommand(
                    "SELECT Precio FROM Productos WHERE Codigo = @codigo",
                    connection);
                command.Parameters.AddWithValue("@codigo", codigo);

                var result = command.ExecuteScalar();
                return result != DBNull.Value ? Convert.ToDouble(result) : 0;
            }
        }

        public DataTable GetInventarioPorCodigo(string codigo)
        {
            DataTable inventario = new DataTable();

            using (var connection = new SQLiteConnection(connectionString))
            {
                connection.Open();
                EnableForeignKeys(connection);

                string query = "SELECT * FROM Inventario WHERE Codigo = @codigo";
                using (var command = new SQLiteCommand(query, connection))
                {
                    command.Parameters.AddWithValue("@codigo", codigo);
                    using (var adapter = new SQLiteDataAdapter(command))
                    {
                        adapter.Fill(inventario);
                    }
                }
            }

            return inventario;
        }

        public void ActualizarInventario(string codigo, int unidadesVendidas, double kilosVendidos)
        {
            using (var connection = new SQLiteConnection(connectionString))
            {
                connection.Open();

                var command = new SQLiteCommand(@"
            UPDATE Inventario
            SET 
                Unidades = Unidades - @unidadesVendidas,
                Kilos = Kilos - @kilosVendidos
            WHERE Codigo = @codigo", connection);

                command.Parameters.AddWithValue("@codigo", codigo);
                command.Parameters.AddWithValue("@unidadesVendidas", unidadesVendidas);
                command.Parameters.AddWithValue("@kilosVendidos", kilosVendidos);

                command.ExecuteNonQuery();
            }
        }

        public void RegistrarVenta(string codigo, int unidades, double kilos, double total)
        {
            try
            {
                using (var connection = new SQLiteConnection(connectionString))
                {
                    connection.Open();

                    string query = @"
                INSERT INTO Ventas (Codigo, Unidades, Kilos, FechaVenta, Total)
                VALUES (@codigo, @unidades, @kilos, @fechaVenta, @total)";

                    using (var command = new SQLiteCommand(query, connection))
                    {
                        command.Parameters.AddWithValue("@codigo", codigo);
                        command.Parameters.AddWithValue("@unidades", unidades);
                        command.Parameters.AddWithValue("@kilos", kilos);
                        command.Parameters.AddWithValue("@fechaVenta", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
                        command.Parameters.AddWithValue("@total", total);

                        command.ExecuteNonQuery();
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Error al registrar la venta: {ex.Message}");
            }
        }



        public void ActualizarFechasInventario(string codigo, DateTime fechaIngresada)
        {
            using (var connection = new SQLiteConnection(connectionString))
            {
                connection.Open();
                EnableForeignKeys(connection);
                string updateQuery = @"
                    UPDATE Inventario
                    SET 
                        FechaMasAntigua = CASE WHEN FechaMasAntigua > @fecha THEN @fecha ELSE FechaMasAntigua END,
                        FechaMasNueva = CASE WHEN FechaMasNueva < @fecha THEN @fecha ELSE FechaMasNueva END
                    WHERE Codigo = @codigo";

                using (var command = new SQLiteCommand(updateQuery, connection))
                {
                    command.Parameters.AddWithValue("@codigo", codigo);
                    command.Parameters.AddWithValue("@fecha", fechaIngresada.ToString("yyyy-MM-dd"));
                    command.ExecuteNonQuery();
                }
            }
        }

        public List<Cliente> GetClientes()
        {
            List<Cliente> clientes = new List<Cliente>();

            using (var connection = new SQLiteConnection(connectionString))
            {
                connection.Open();
                EnableForeignKeys(connection);
                string query = "SELECT * FROM Clientes";

                using (var command = new SQLiteCommand(query, connection))
                {
                    using (var reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            clientes.Add(new Cliente
                            {
                                RUT = reader["RUT"].ToString(),
                                Nombre = reader["Nombre"].ToString(),
                                Direccion = reader["Direccion"].ToString(),
                                Giro = reader["Giro"].ToString(),
                                Deuda = reader["Deuda"] != DBNull.Value ? Convert.ToDouble(reader["Deuda"]) : 0
                            });
                        }
                    }
                }

                return clientes;
            }

        }

        public Cliente GetClientePorRUT(string rut)
        {
            Cliente cliente = null;

            using (var connection = new SQLiteConnection(connectionString))
            {
                connection.Open();
                EnableForeignKeys(connection);
                string query = "SELECT * FROM Clientes WHERE RUT = @rut";

                using (var command = new SQLiteCommand(query, connection))
                {
                    command.Parameters.AddWithValue("@rut", rut);

                    using (var reader = command.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            cliente = new Cliente
                            {
                                RUT = reader["RUT"].ToString(),
                                Nombre = reader["Nombre"].ToString(),
                                Direccion = reader["Direccion"].ToString(),
                                Giro = reader["Giro"].ToString(),
                                Deuda = reader["Deuda"] != DBNull.Value ? Convert.ToDouble(reader["Deuda"]) : 0
                            };
                        }
                    }
                }
            }

            return cliente;
        }
    }
}


















