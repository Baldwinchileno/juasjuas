using System;
using System.Collections.Generic;
using System.Data;
using System.Threading.Tasks;
using AdminSERMAC.Models;

public interface IInventarioDatabaseService
{
    Task<bool> AddProductoAsync(string codigo, string producto, int unidades, double kilos, string fechaCompra, string fechaRegistro, string fechaVencimiento);
    Task<bool> ActualizarInventarioAsync(string codigo, int unidadesVendidas, double kilosVendidos);
    Task<IEnumerable<string>> GetCategoriasAsync();
    Task<IEnumerable<string>> GetSubCategoriasAsync(string categoria);
    Task<DataTable> GetInventarioAsync();
    Task<DataTable> GetInventarioPorCodigoAsync(string codigo);
    Task<bool> ActualizarFechasInventarioAsync(string codigo, DateTime fechaIngresada);

    Task<List<CompraRegistro>> GetAllCompraRegistrosAsync();
    Task AddCompraRegistroAsync(CompraRegistro registro);
    Task<CompraRegistro> GetCompraRegistroByIdAsync(int id);
    Task UpdateCompraRegistroAsync(CompraRegistro registro);
    Task DeleteCompraRegistroAsync(int id);
    Task ProcesarCompraRegistroAsync(int id);
}
