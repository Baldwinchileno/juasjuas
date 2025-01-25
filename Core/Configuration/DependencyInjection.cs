using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;
using AdminSERMAC.Services.Database;
using AdminSERMAC.Core.Interfaces;
using AdminSERMAC.Repositories;
using AdminSERMAC.Services;
using AdminSERMAC.Core.Infrastructure;

namespace AdminSERMAC.Core.Configuration
{
    public static class DependencyInjection
    {
        public static IServiceCollection AddInfrastructure(this IServiceCollection services, string connectionString)
        {
            // Base Database Services
            services.AddSingleton<DatabaseService>();

            services.AddSingleton<SQLiteService>();

            // Database Services
            services.AddScoped<IClienteDatabaseService, ClienteDatabaseService>();
            services.AddScoped<IInventarioDatabaseService>(provider =>
            {
                var logger = provider.GetRequiredService<ILogger<InventarioDatabaseService>>();
                return new InventarioDatabaseService(logger, connectionString);
            });

            services.AddScoped<IProductoDatabaseService, ProductoDatabaseService>();
            services.AddScoped<IVentaDatabaseService, VentaDatabaseService>();
            services.AddScoped<CompraRegistroDatabaseService>();
            // Business Services
            services.AddScoped<IClienteService, ClienteService>();

            // Repositories (if still needed)
            services.AddScoped<IClienteRepository, ClienteRepository>();
            services.AddScoped<IProductoRepository, ProductoRepository>();
            services.AddScoped<IInventarioRepository, InventarioRepository>();
            services.AddScoped<IVentaRepository, VentaRepository>();

            // Unit of Work
            services.AddScoped<IUnitOfWork, UnitOfWork>();

            // Add connection string configuration
            services.AddSingleton(_ => connectionString);

            // Logging
            services.AddLogging(builder =>
            {
                builder.AddConsole();
                builder.AddDebug();
            });

            return services;
        }
    }
}
