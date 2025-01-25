using System;
using System.Data.SQLite;
using System.Threading.Tasks;
using System.Collections.Generic;
using Microsoft.Extensions.Logging;

namespace AdminSERMAC.Services.Optimization
{
    public class DatabaseOptimizer
    {
        private readonly string _connectionString;
        private readonly ILogger<DatabaseOptimizer> _logger;

        public DatabaseOptimizer(string connectionString, ILogger<DatabaseOptimizer> logger)
        {
            _connectionString = connectionString;
            _logger = logger;
        }

        public async Task<bool> OptimizeDatabaseAsync()
        {
            try
            {
                using var connection = new SQLiteConnection(_connectionString);
                await connection.OpenAsync();

                // Run VACUUM to reclaim space and defrag
                using (var command = connection.CreateCommand())
                {
                    command.CommandText = "VACUUM;";
                    await command.ExecuteNonQueryAsync();
                }

                // Analyze tables for query optimization
                using (var command = connection.CreateCommand())
                {
                    command.CommandText = "ANALYZE;";
                    await command.ExecuteNonQueryAsync();
                }

                // Optimize indexes
                await OptimizeIndexesAsync(connection);

                _logger.LogInformation("Database optimization completed successfully");
                return true;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error during database optimization");
                return false;
            }
        }

        private async Task OptimizeIndexesAsync(SQLiteConnection connection)
        {
            var tables = await GetTableNamesAsync(connection);
            foreach (var table in tables)
            {
                await ReindexTableAsync(connection, table);
            }
        }

        private async Task<List<string>> GetTableNamesAsync(SQLiteConnection connection)
        {
            var tables = new List<string>();
            using var command = connection.CreateCommand();
            command.CommandText = "SELECT name FROM sqlite_master WHERE type='table';";

            using var reader = await command.ExecuteReaderAsync();
            while (await reader.ReadAsync())
            {
                tables.Add(reader.GetString(0));
            }
            return tables;
        }

        private async Task ReindexTableAsync(SQLiteConnection connection, string tableName)
        {
            using var command = connection.CreateCommand();
            command.CommandText = $"REINDEX '{tableName}';";
            await command.ExecuteNonQueryAsync();
        }

        public async Task<DatabaseStats> GetDatabaseStatsAsync()
        {
            try
            {
                using var connection = new SQLiteConnection(_connectionString);
                await connection.OpenAsync();

                var stats = new DatabaseStats
                {
                    DatabaseSize = GetDatabaseSize(connection),
                    TableStats = await GetTableStatsAsync(connection)
                };

                return stats;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error getting database statistics");
                throw;
            }
        }

        private long GetDatabaseSize(SQLiteConnection connection)
        {
            using var command = connection.CreateCommand();
            command.CommandText = "SELECT page_count * page_size as size FROM pragma_page_count(), pragma_page_size();";
            return Convert.ToInt64(command.ExecuteScalar());
        }

        private async Task<Dictionary<string, TableStats>> GetTableStatsAsync(SQLiteConnection connection)
        {
            var tableStats = new Dictionary<string, TableStats>();
            var tables = await GetTableNamesAsync(connection);

            foreach (var table in tables)
            {
                using var command = connection.CreateCommand();
                command.CommandText = $"SELECT COUNT(*) FROM '{table}';";
                var rowCount = Convert.ToInt64(await command.ExecuteScalarAsync());

                command.CommandText = $"SELECT COUNT(*) FROM pragma_index_list('{table}');";
                var indexCount = Convert.ToInt32(await command.ExecuteScalarAsync());

                tableStats[table] = new TableStats
                {
                    RowCount = rowCount,
                    IndexCount = indexCount
                };
            }

            return tableStats;
        }

        public async Task<bool> CompactDatabaseAsync()
        {
            try
            {
                using var connection = new SQLiteConnection(_connectionString);
                await connection.OpenAsync();

                // Clear unused space
                using (var command = connection.CreateCommand())
                {
                    command.CommandText = "VACUUM;";
                    await command.ExecuteNonQueryAsync();
                }

                // Clear statement cache
                using (var command = connection.CreateCommand())
                {
                    command.CommandText = "PRAGMA optimize;";
                    await command.ExecuteNonQueryAsync();
                }

                _logger.LogInformation("Database compaction completed successfully");
                return true;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error during database compaction");
                return false;
            }
        }

        public async Task<bool> CheckDatabaseIntegrityAsync()
        {
            try
            {
                using var connection = new SQLiteConnection(_connectionString);
                await connection.OpenAsync();
                using var command = connection.CreateCommand();

                command.CommandText = "PRAGMA integrity_check;";
                var result = await command.ExecuteScalarAsync() as string;

                var isValid = result == "ok";
                if (!isValid)
                {
                    _logger.LogError("Database integrity check failed: {Result}", result);
                }

                return isValid;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error checking database integrity");
                return false;
            }
        }

        public async Task<bool> RebuildIndexesAsync()
        {
            try
            {
                using var connection = new SQLiteConnection(_connectionString);
                await connection.OpenAsync();

                var tables = await GetTableNamesAsync(connection);
                foreach (var table in tables)
                {
                    using var command = connection.CreateCommand();
                    command.CommandText = $"REINDEX '{table}';";
                    await command.ExecuteNonQueryAsync();
                }

                _logger.LogInformation("Database indexes rebuilt successfully");
                return true;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error rebuilding database indexes");
                return false;
            }
        }
    }

    public class DatabaseStats
    {
        public long DatabaseSize { get; set; }
        public Dictionary<string, TableStats> TableStats { get; set; }
    }

    public class TableStats
    {
        public long RowCount { get; set; }
        public int IndexCount { get; set; }
    }
}
