using System.Data.SQLite;
using Microsoft.Extensions.Logging;

namespace AdminSERMAC.Services.Database
{
    public abstract class BaseSQLiteService
    {
        protected readonly string _connectionString;
        protected readonly ILogger _logger;

        protected BaseSQLiteService(ILogger logger, string connectionString)
        {
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
            _connectionString = connectionString ?? throw new ArgumentNullException(nameof(connectionString));
        }

        protected async Task<T> ExecuteInTransactionAsync<T>(Func<SQLiteConnection, SQLiteTransaction, Task<T>> operation)
        {
            using var connection = new SQLiteConnection(_connectionString);
            await connection.OpenAsync();

            using var transaction = connection.BeginTransaction();
            try
            {
                var result = await operation(connection, transaction);
                transaction.Commit();
                return result;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error executing database operation");
                transaction.Rollback();
                throw;
            }
        }

        protected async Task ExecuteInTransactionAsync(Func<SQLiteConnection, SQLiteTransaction, Task> operation)
        {
            await ExecuteInTransactionAsync<object>(async (conn, trans) =>
            {
                await operation(conn, trans);
                return null;
            });
        }

        protected void EnableForeignKeys(SQLiteConnection connection)
        {
            using var command = new SQLiteCommand("PRAGMA foreign_keys = ON;", connection);
            command.ExecuteNonQuery();
        }

        protected bool ColumnExists(SQLiteConnection connection, string tableName, string columnName)
        {
            using var command = connection.CreateCommand();
            command.CommandText = $"PRAGMA table_info({tableName})";
            using var reader = command.ExecuteReader();
            while (reader.Read())
            {
                if (reader["name"].ToString() == columnName)
                {
                    return true;
                }
            }
            return false;
        }
    }
}
