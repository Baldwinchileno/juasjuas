namespace AdminSERMAC.Services.Database.Interfaces
{
    public interface IDatabaseService
    {
        Task<bool> EnsureTableExistsAsync(string tableName, string createTableSql);
        Task<bool> EnsureColumnExistsAsync(string tableName, string columnName, string columnType);
        Task<bool> BackupDatabaseAsync(string backupPath);
        Task<bool> ValidateConnectionAsync();
        string GetConnectionString();
    }

    public interface IDatabaseMigrationService
    {
        Task MigrateAsync();
        Task<int> GetCurrentVersionAsync();
        Task<bool> SetVersionAsync(int version);
    }
}
