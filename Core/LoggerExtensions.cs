using Microsoft.Extensions.Logging;

namespace AdminSERMAC.Core
{
    public static class LoggerExtensions
    {
        public static ILogger<T> GetLogger<T>(this ILogger logger)
        {
            return (ILogger<T>)logger;
        }
    }
}