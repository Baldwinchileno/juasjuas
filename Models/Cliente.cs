using System.ComponentModel.DataAnnotations;

namespace AdminSERMAC.Models
{
    public class Cliente
    {
        [Required(ErrorMessage = "El RUT es obligatorio")]
        [RegularExpression(@"^(\d{1,2}\.\d{3}\.\d{3}[-][0-9kK]{1})$",
            ErrorMessage = "El formato del RUT no es válido")]
        public string RUT { get; set; } = string.Empty;

        [Required(ErrorMessage = "El nombre es obligatorio")]
        [StringLength(100, MinimumLength = 2,
            ErrorMessage = "El nombre debe tener entre 2 y 100 caracteres")]
        public string Nombre { get; set; } = string.Empty;

        [Required(ErrorMessage = "La dirección es obligatoria")]
        public string Direccion { get; set; } = string.Empty;

        [Required(ErrorMessage = "El giro es obligatorio")]
        public string Giro { get; set; } = string.Empty;

        public double Deuda { get; set; }

        public bool IsValid()
        {
            if (string.IsNullOrWhiteSpace(RUT) ||
                string.IsNullOrWhiteSpace(Nombre) ||
                string.IsNullOrWhiteSpace(Direccion) ||
                string.IsNullOrWhiteSpace(Giro))
                return false;

            return true;
        }

        public override string ToString()
        {
            return $"{RUT} - {Nombre}";
        }
    }
}
