﻿using Microsoft.Extensions.Configuration;

namespace AdminSERMAC;

public class AppSettings
{
    private static IConfiguration? _configuration;
    private const string DefaultConnectionString = "Data Source=AdminSERMAC.db;Version=3;";

    public static IConfiguration Configuration
    {
        get
        {
            if (_configuration == null)
            {
                var builder = new ConfigurationBuilder()
                    .SetBasePath(Directory.GetCurrentDirectory())
                    .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true);

                _configuration = builder.Build();
            }
            return _configuration;
        }
    }

    public static string GetConnectionString(string name = "DefaultConnection")
    {
        try
        {
            var connectionString = Configuration.GetConnectionString(name);
            return string.IsNullOrEmpty(connectionString) ? DefaultConnectionString : connectionString;
        }
        catch
        {
            return DefaultConnectionString;
        }
    }
}