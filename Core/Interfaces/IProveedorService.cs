using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using AdminSERMAC.Models;

namespace AdminSERMAC.Core.Interfaces
{
    public interface IProveedorService
    {
        Task<List<Proveedor>> GetAllProveedores();
        Task<Proveedor> GetProveedorById(int id);
        Task<bool> CreateProveedor(Proveedor proveedor);
        Task<bool> UpdateProveedor(Proveedor proveedor);
        Task<bool> DeleteProveedor(int id);
    }
}
