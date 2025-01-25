using System;
using System.IO;
using Microsoft.Extensions.Configuration;

namespace AdminSERMAC.Core.Configuration
{
    public class ConfigurationService
    {
        private readonly IConfiguration _configuration;

        public ConfigurationService()
        {
            var builder = new ConfigurationBuilder()
                .SetBasePath(Directory.GetCurrentDirectory())
                .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true);

            _configuration = builder.Build();
        }

        public string GetConnectionString()
        {
            return _configuration.GetConnectionString("DefaultConnection");
        }

        public string GetValue(string key)
        {
            return _configuration[key];
        }

        public string this[string key]
        {
            get { return _configuration[key]; }
        }

        public IConfigurationSection GetSection(string key)
        {
            return _configuration.GetSection(key);
        }

        public bool HasValue(string key)
        {
            return !string.IsNullOrEmpty(_configuration[key]);
        }

        public T GetValueOrDefault<T>(string key, T defaultValue) where T : struct
        {
            var value = _configuration[key];
            if (string.IsNullOrEmpty(value))
            {
                return defaultValue;
            }

            try
            {
                return (T)Convert.ChangeType(value, typeof(T));
            }
            catch
            {
                return defaultValue;
            }
        }

        public string GetValueOrDefault(string key, string defaultValue)
        {
            var value = _configuration[key];
            return string.IsNullOrEmpty(value) ? defaultValue : value;
        }
    }
}