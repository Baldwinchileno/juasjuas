using System;
using System.Drawing;
using System.Windows.Forms;
using AdminSERMAC.Models;
using AdminSERMAC.Services;
using Microsoft.Extensions.Logging;

namespace AdminSERMAC.Forms
{
    public class ModificarClienteForm : Form
    {
        private TextBox nombreTextBox;
        private TextBox direccionTextBox;
        private TextBox giroTextBox;
        private TextBox deudaTextBox;
        private Button guardarButton;
        private readonly SQLiteService sqliteService;
        private readonly ILogger<ClientesForm> _logger;
        private readonly Cliente cliente;

        public ModificarClienteForm(Cliente cliente, SQLiteService sqliteService, ILogger<ClientesForm> logger)
        {
            this.cliente = cliente ?? throw new ArgumentNullException(nameof(cliente));
            this.sqliteService = sqliteService ?? throw new ArgumentNullException(nameof(sqliteService));
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));

            InitializeComponents();
            CargarDatosCliente();
        }

        private void InitializeComponents()
        {
            this.Text = "Modificar Cliente";
            this.Width = 400;
            this.Height = 300;

            Label nombreLabel = new Label { Text = "Nombre:", Location = new Point(20, 20) };
            nombreTextBox = new TextBox { Location = new Point(100, 20), Width = 250 };

            Label direccionLabel = new Label { Text = "Dirección:", Location = new Point(20, 60) };
            direccionTextBox = new TextBox { Location = new Point(100, 60), Width = 250 };

            Label giroLabel = new Label { Text = "Giro:", Location = new Point(20, 100) };
            giroTextBox = new TextBox { Location = new Point(100, 100), Width = 250 };

            Label deudaLabel = new Label { Text = "Deuda:", Location = new Point(20, 140) };
            deudaTextBox = new TextBox { Location = new Point(100, 140), Width = 250 };

            guardarButton = new Button { Text = "Guardar", Location = new Point(150, 200), Width = 100 };
            guardarButton.Click += GuardarButton_Click;

            this.Controls.Add(nombreLabel);
            this.Controls.Add(nombreTextBox);
            this.Controls.Add(direccionLabel);
            this.Controls.Add(direccionTextBox);
            this.Controls.Add(giroLabel);
            this.Controls.Add(giroTextBox);
            this.Controls.Add(deudaLabel);
            this.Controls.Add(deudaTextBox);
            this.Controls.Add(guardarButton);
        }

        private void CargarDatosCliente()
        {
            nombreTextBox.Text = cliente.Nombre;
            direccionTextBox.Text = cliente.Direccion;
            giroTextBox.Text = cliente.Giro;
            deudaTextBox.Text = cliente.Deuda.ToString();
        }

        private void GuardarButton_Click(object sender, EventArgs e)
        {
            cliente.Nombre = nombreTextBox.Text;
            cliente.Direccion = direccionTextBox.Text;
            cliente.Giro = giroTextBox.Text;
            cliente.Deuda = Convert.ToDouble(deudaTextBox.Text);

            try
            {
                sqliteService.ActualizarCliente(cliente);
                MessageBox.Show("Cliente actualizado correctamente.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
                this.Close();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al actualizar el cliente.");
                MessageBox.Show($"Error al actualizar el cliente: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
