using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AdminSERMAC.Models
{
    public class Abono
    {
        public int Id { get; set; }
        public string RUT { get; set; } // RUT del cliente asociado
        public DateTime Fecha { get; set; } // Fecha del abono
        public double Monto { get; set; } // Monto del abono
        public string Descripcion { get; set; } // Opcional, descripción del abono
    }
}
