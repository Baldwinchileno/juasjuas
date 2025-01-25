using Microsoft.Extensions.Logging;
using System;
using System.Data.SQLite;

namespace AdminSERMAC.Services.Database
{
    public class DatabaseInitializer
    {
        private readonly string _connectionString;
        private readonly ILogger<DatabaseInitializer> _logger;

        public DatabaseInitializer(ILogger<DatabaseInitializer> logger, string connectionString)
        {
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
            _connectionString = connectionString ?? throw new ArgumentNullException(nameof(connectionString));
        }

        public void InitializeDatabase()
        {
            try
            {
                using (var connection = new SQLiteConnection(_connectionString))
                {
                    connection.Open();
                    CreateTables(connection);
                }
                _logger.LogInformation("Base de datos inicializada correctamente");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al inicializar la base de datos");
                throw;
            }
        }

        private void CreateTables(SQLiteConnection connection)
        {
            // Crear tabla de Clientes
            ExecuteCommand(connection, @"
                CREATE TABLE IF NOT EXISTS Clientes (
                    Id INTEGER PRIMARY KEY AUTOINCREMENT,
                    Nombre TEXT NOT NULL,
                    RUT TEXT UNIQUE NOT NULL,
                    Direccion TEXT,
                    Telefono TEXT,
                    Email TEXT,
                    FechaRegistro TEXT NOT NULL,
                    UltimaModificacion TEXT
                );");

            // Crear tabla de Inventario
            ExecuteCommand(connection, @"
                CREATE TABLE IF NOT EXISTS Inventario (
                    Id INTEGER PRIMARY KEY AUTOINCREMENT,
                    Codigo TEXT NOT NULL UNIQUE,
                    Producto TEXT NOT NULL,
                    Unidades INTEGER NOT NULL DEFAULT 0,
                    Kilos REAL NOT NULL DEFAULT 0,
                    FechaMasAntigua TEXT,
                    FechaMasNueva TEXT,
                    FechaVencimiento TEXT,
                    Categoria TEXT,
                    SubCategoria TEXT
                );");

            // Crear tabla de Ventas
            ExecuteCommand(connection, @"
                CREATE TABLE IF NOT EXISTS Ventas (
                    Id INTEGER PRIMARY KEY AUTOINCREMENT,
                    ClienteId INTEGER,
                    FechaVenta TEXT NOT NULL,
                    Total REAL NOT NULL,
                    Estado TEXT NOT NULL,
                    FOREIGN KEY(ClienteId) REFERENCES Clientes(Id)
                );");

            // Crear tabla de Detalles de Venta
            ExecuteCommand(connection, @"
                CREATE TABLE IF NOT EXISTS DetallesVenta (
                    Id INTEGER PRIMARY KEY AUTOINCREMENT,
                    VentaId INTEGER NOT NULL,
                    ProductoId INTEGER NOT NULL,
                    Cantidad INTEGER NOT NULL,
                    PrecioUnitario REAL NOT NULL,
                    Subtotal REAL NOT NULL,
                    FOREIGN KEY(VentaId) REFERENCES Ventas(Id),
                    FOREIGN KEY(ProductoId) REFERENCES Inventario(Id)
                );");

            // Crear tabla de Traspasos
            ExecuteCommand(connection, @"
                CREATE TABLE IF NOT EXISTS Traspasos (
                    Id INTEGER PRIMARY KEY AUTOINCREMENT,
                    ProductoId INTEGER NOT NULL,
                    CantidadUnidades INTEGER NOT NULL,
                    CantidadKilos REAL NOT NULL,
                    FechaTraspaso TEXT NOT NULL,
                    LocalOrigen TEXT NOT NULL,
                    LocalDestino TEXT NOT NULL,
                    FOREIGN KEY(ProductoId) REFERENCES Inventario(Id)
                );");

            // Crear tabla de Compras
            ExecuteCommand(connection, @"
                CREATE TABLE IF NOT EXISTS CompraRegistros (
                    Id INTEGER PRIMARY KEY AUTOINCREMENT,
                    Proveedor TEXT NOT NULL,
                    Producto TEXT NOT NULL,
                    Cantidad INTEGER NOT NULL,
                    PrecioUnitario REAL NOT NULL,
                    Total REAL NOT NULL,
                    Observaciones TEXT,
                    FechaCompra TEXT NOT NULL,
                    EstaProcesado INTEGER NOT NULL DEFAULT 0
                );");

            // Crear tabla de Abonos
            ExecuteCommand(connection, @"
                CREATE TABLE IF NOT EXISTS Abonos (
                    Id INTEGER PRIMARY KEY AUTOINCREMENT,
                    VentaId INTEGER NOT NULL,
                    Monto REAL NOT NULL,
                    FechaAbono TEXT NOT NULL,
                    FOREIGN KEY(VentaId) REFERENCES Ventas(Id)
                );");
        }

        private void ExecuteCommand(SQLiteConnection connection, string commandText)
        {
            try
            {
                using (var command = new SQLiteCommand(commandText, connection))
                {
                    command.ExecuteNonQuery();
                }
                _logger.LogInformation($"Comando SQL ejecutado exitosamente: {commandText.Substring(0, Math.Min(50, commandText.Length))}...");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"Error al ejecutar comando SQL: {commandText}");
                throw;
            }
        }
    }
}