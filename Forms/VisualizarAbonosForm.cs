using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using AdminSERMAC.Core.Interfaces;

namespace AdminSERMAC.Forms
{
    public class VisualizarAbonosForm : Form
    {
        private readonly string _clienteRUT;
        private readonly IClienteService _clienteService;
        private DataGridView abonosDataGridView;

        public VisualizarAbonosForm(string clienteRUT, IClienteService clienteService)
        {
            _clienteRUT = clienteRUT ?? throw new ArgumentNullException(nameof(clienteRUT));
            _clienteService = clienteService ?? throw new ArgumentNullException(nameof(clienteService));

            InitializeComponents();
            LoadAbonos();
        }

        private void InitializeComponents()
        {
            this.Text = "Visualizar Abonos";
            this.Width = 800;
            this.Height = 600;

            abonosDataGridView = new DataGridView
            {
                Location = new Point(20, 20),
                Size = new Size(740, 500),
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                ReadOnly = true,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect
            };

            this.Controls.Add(abonosDataGridView);
        }

        private async void LoadAbonos()
        {
            try
            {
                var abonos = await _clienteService.ObtenerAbonosPorClienteAsync(_clienteRUT);
                abonosDataGridView.DataSource = abonos.ToList(); // Asegúrate de materializar los datos
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error al cargar los abonos: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}

