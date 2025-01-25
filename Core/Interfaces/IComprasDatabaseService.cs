using System.Threading.Tasks;
using System.Collections.Generic;
using AdminSERMAC.Models;

namespace AdminSERMAC.Core.Interfaces
{
    public interface IComprasDatabaseService
    {
        Task<int> CreateCompra(Compra compra);
        Task<bool> AddDetalleCompra(DetalleCompra detalle);
        Task<Compra> GetCompraById(int numeroCompra);
        Task<List<Compra>> GetAllCompras();
        Task<List<DetalleCompra>> GetDetallesByCompra(int numeroCompra);
        Task<bool> UpdateCompraEstado(int numeroCompra, string estado);
        Task<int> GetNextNumeroCompra();
    }
}
