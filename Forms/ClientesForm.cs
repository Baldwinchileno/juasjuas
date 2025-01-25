using System;
using System.Data.SQLite;
using System.Drawing;
using System.Windows.Forms;
using AdminSERMAC.Core.Interfaces;
using AdminSERMAC.Exceptions;
using AdminSERMAC.Models;
using AdminSERMAC.Services;
using Microsoft.Extensions.Logging;

namespace AdminSERMAC.Forms
{
    public class ClientesForm : Form
    {
        #region Fields
        private readonly IClienteService _clienteService;
        private readonly SQLiteService sqliteService;
        private readonly ILogger<ClientesForm> _logger;

        // Controles para datos del cliente
        private GroupBox clienteGroupBox;
        private TextBox rutTextBox;
        private TextBox nombreTextBox;
        private TextBox direccionTextBox;
        private TextBox giroTextBox;
        private Button agregarButton;
        private Button actualizarButton;
        private Button eliminarButton;
        private Button limpiarButton;

        // Controles para abonos
        private GroupBox grupoAbono;
        private TextBox montoAbonoTextBox;
        private Button registrarAbonoButton;
        private Button visualizarAbonosButton;

        // Controles para selección y visualización
        private ComboBox clientesComboBox;
        private Label deudaLabel;
        private Label seleccionarClienteLabel;
        private DataGridView ventasDataGridView;
        private DataGridView clientesDataGridView;
        private Button modificarClienteButton;
        #endregion

        #region Constructor
        public ClientesForm(SQLiteService sqliteService, ILogger<ClientesForm> logger, IClienteService clienteService)
        {
            this.sqliteService = sqliteService ?? throw new ArgumentNullException(nameof(sqliteService));
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
            _clienteService = clienteService ?? throw new ArgumentNullException(nameof(clienteService));

            InitializeComponents();
            ConfigurarEventos();
            LoadClientes();
            AjustarEstilosBotones();
        }
        #endregion

        #region Inicialización y Configuración
        private void InitializeComponents()
        {
            // Ajustar tamaños y propiedades del form
            this.MinimumSize = new Size(1200, 800);
            this.AutoScaleMode = AutoScaleMode.None;

            this.Text = "Gestión de Clientes";
            this.Width = 1200;
            this.Height = 800;
            this.StartPosition = FormStartPosition.CenterScreen;

            // GroupBox para datos del cliente
            clienteGroupBox = new GroupBox
            {
                Text = "Datos del Cliente",
                Location = new Point(20, 20),
                Size = new Size(500, 250),
                Font = new Font("Segoe UI", 10F)
            };

            // Controles para datos del cliente
            var rutLabel = new Label
            {
                Text = "RUT:",
                Location = new Point(20, 35),
                Size = new Size(100, 23),
                Font = new Font("Segoe UI", 10F)
            };

            rutTextBox = new TextBox
            {
                Location = new Point(130, 35),
                Size = new Size(250, 23),
                Font = new Font("Segoe UI", 10F)
            };

            var nombreLabel = new Label
            {
                Text = "Nombre:",
                Location = new Point(20, 75),
                Size = new Size(100, 23),
                Font = new Font("Segoe UI", 10F)
            };

            nombreTextBox = new TextBox
            {
                Location = new Point(130, 75),
                Size = new Size(250, 23),
                Font = new Font("Segoe UI", 10F)
            };

            var direccionLabel = new Label
            {
                Text = "Dirección:",
                Location = new Point(20, 115),
                Size = new Size(100, 23),
                Font = new Font("Segoe UI", 10F)
            };

            direccionTextBox = new TextBox
            {
                Location = new Point(130, 115),
                Size = new Size(250, 23),
                Font = new Font("Segoe UI", 10F)
            };

            var giroLabel = new Label
            {
                Text = "Giro:",
                Location = new Point(20, 155),
                Size = new Size(100, 23),
                Font = new Font("Segoe UI", 10F)
            };

            giroTextBox = new TextBox
            {
                Location = new Point(130, 155),
                Size = new Size(250, 23),
                Font = new Font("Segoe UI", 10F)
            };

            // Botones de acción
            agregarButton = new Button
            {
                Text = "Agregar Cliente",
                Location = new Point(20, 195),
                Size = new Size(150, 35)
            };

            actualizarButton = new Button
            {
                Text = "Actualizar",
                Location = new Point(180, 195),
                Size = new Size(150, 35),
                Enabled = false
            };

            eliminarButton = new Button
            {
                Text = "Eliminar",
                Location = new Point(340, 195),
                Size = new Size(150, 35),
                Enabled = false
            };

            // Agregar controles al GroupBox
            clienteGroupBox.Controls.AddRange(new Control[] {
    rutLabel, rutTextBox,
    nombreLabel, nombreTextBox,
    direccionLabel, direccionTextBox,
    giroLabel, giroTextBox,
    agregarButton, actualizarButton, eliminarButton,
    deudaLabel
});

            // Grupo de abono
            grupoAbono = new GroupBox
            {
                Text = "Registrar Abono",
                Location = new Point(540, 20),
                Size = new Size(300, 150),
                Font = new Font("Segoe UI", 10F)
            };

            var montoAbonoLabel = new Label
            {
                Text = "Monto Abono ($):",
                Location = new Point(20, 35),
                Size = new Size(120, 23),
                Font = new Font("Segoe UI", 10F)
            };

            montoAbonoTextBox = new TextBox
            {
                Location = new Point(140, 35),
                Size = new Size(140, 23),
                Font = new Font("Segoe UI", 10F)
            };

            registrarAbonoButton = new Button
            {
                Text = "Registrar Abono",
                Location = new Point(140, 75),
                Size = new Size(140, 35),
                Enabled = false
            };

            grupoAbono.Controls.AddRange(new Control[] {
        montoAbonoLabel,
        montoAbonoTextBox,
        registrarAbonoButton
    });

            visualizarAbonosButton = new Button
            {
                Text = "Visualizar Abonos",
                Location = new Point(140, 120), // Justo debajo del botón "Registrar Abono"
                Size = new Size(140, 35),
                Enabled = true // Habilitarlo dinámicamente según la lógica
            };

            grupoAbono.Controls.Add(visualizarAbonosButton);

            visualizarAbonosButton.Click += VisualizarAbonosButton_Click;

            // ComboBox y Label de selección de cliente
            seleccionarClienteLabel = new Label
            {
                Text = "Seleccionar Cliente:",
                Location = new Point(20, 290),
                Size = new Size(200, 23),
                Font = new Font("Segoe UI", 10F)
            };

            clientesComboBox = new ComboBox
            {
                Location = new Point(20, 320),
                Size = new Size(500, 23),
                DropDownStyle = ComboBoxStyle.DropDownList,
                Font = new Font("Segoe UI", 10F)
            };

            // Label de deuda
            deudaLabel = new Label
{
    Location = new Point(340, 195),
    MinimumSize = new Size(350, 35),
    MaximumSize = new Size(500, 35),
    AutoSize = true,
    TextAlign = ContentAlignment.MiddleLeft,
    Font = new Font("Segoe UI", 10F, FontStyle.Bold)
};

            // DataGridView de ventas
            ventasDataGridView = new DataGridView
            {
                Location = new Point(20, 400),
                Size = new Size(1150, 300),
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                ReadOnly = true,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                BackgroundColor = Color.White,
                BorderStyle = BorderStyle.Fixed3D,
                Font = new Font("Segoe UI", 9F)
            };

            ConfigurarColumnas();

            // Agregar controles al formulario
            this.Controls.AddRange(new Control[] {
        clienteGroupBox,
        grupoAbono,
        seleccionarClienteLabel,
        clientesComboBox,
        deudaLabel,
        ventasDataGridView
    });

            // Ajustar tamaños y propiedades del form
            this.MinimumSize = new Size(1200, 800);
            this.AutoScaleMode = AutoScaleMode.None;

            // Modificar el DataGridView para que se expanda
            ventasDataGridView.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;

            // Anclar controles inferiores al fondo
            deudaLabel.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            clienteGroupBox.Anchor = AnchorStyles.Top | AnchorStyles.Left;
            grupoAbono.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            seleccionarClienteLabel.Anchor = AnchorStyles.Top | AnchorStyles.Left;
            clientesComboBox.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
        }



        private void ConfigurarEventos()
        {
            rutTextBox.Leave += ValidarRUT;
            clientesComboBox.SelectedIndexChanged += ClientesComboBox_SelectedIndexChanged;
            agregarButton.Click += AgregarButton_Click;
            actualizarButton.Click += ActualizarButton_Click;
            eliminarButton.Click += EliminarButton_Click;
            registrarAbonoButton.Click += RegistrarAbonoButton_Click;
            montoAbonoTextBox.KeyPress += ValidarSoloNumeros;
        }

        private void ConfigurarColumnas()
        {
            ventasDataGridView.Columns.AddRange(new DataGridViewColumn[] {
                new DataGridViewTextBoxColumn { Name = "FechaVenta", HeaderText = "Fecha", Width = 100 },
                new DataGridViewTextBoxColumn { Name = "NumeroGuia", HeaderText = "N° Guía", Width = 80 },
                new DataGridViewTextBoxColumn { Name = "Descripcion", HeaderText = "Descripción", Width = 200 },
                new DataGridViewTextBoxColumn { Name = "KilosNeto", HeaderText = "Kilos", Width = 80 },
                new DataGridViewTextBoxColumn { Name = "Total", HeaderText = "Monto", Width = 100 },
                new DataGridViewTextBoxColumn { Name = "Estado", HeaderText = "Estado", Width = 100 },
                new DataGridViewTextBoxColumn { Name = "DiasMora", HeaderText = "Días Mora", Width = 80 }
            });

            ventasDataGridView.EnableHeadersVisualStyles = false;
            ventasDataGridView.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(45, 66, 91);
            ventasDataGridView.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            ventasDataGridView.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 10F, FontStyle.Bold);
        }

        private void AjustarEstilosBotones()
        {
            var botones = new[] { agregarButton, actualizarButton, eliminarButton, registrarAbonoButton };

            foreach (var boton in botones)
            {
                boton.FlatStyle = FlatStyle.Flat;
                boton.FlatAppearance.BorderSize = 1;
                boton.Font = new Font("Segoe UI", 10F);
                boton.Cursor = Cursors.Hand;
            }

            agregarButton.BackColor = Color.FromArgb(0, 122, 204);
            agregarButton.ForeColor = Color.White;

            actualizarButton.BackColor = Color.FromArgb(40, 167, 69);
            actualizarButton.ForeColor = Color.White;

            eliminarButton.BackColor = Color.FromArgb(220, 53, 69);
            eliminarButton.ForeColor = Color.White;

            registrarAbonoButton.BackColor = Color.FromArgb(255, 193, 7);
            registrarAbonoButton.ForeColor = Color.Black;
        }
        #endregion

        #region Eventos
        private void ValidarRUT(object sender, EventArgs e)
        {
            string rut = rutTextBox.Text.Trim();
            if (!string.IsNullOrEmpty(rut))
            {
                try
                {
                    if (!ValidarFormatoRUT(rut))
                    {
                        MessageBox.Show("El formato del RUT no es válido.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        rutTextBox.Focus();
                    }
                }
                catch (Exception ex)
                {
                    _logger.LogError(ex, "Error validando RUT");
                    MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void VisualizarAbonosButton_Click(object sender, EventArgs e)
        {
            if (clientesComboBox.SelectedItem == null)
            {
                MessageBox.Show("Seleccione un cliente para visualizar los abonos.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var cliente = (Cliente)clientesComboBox.SelectedItem;

            // Abrir el formulario de visualización de abonos
            var visualizarAbonosForm = new VisualizarAbonosForm(cliente.RUT, _clienteService);
            visualizarAbonosForm.ShowDialog();
        }

        private void ClientesComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (clientesComboBox.SelectedItem == null) return;

            try
            {
                var cliente = (Cliente)clientesComboBox.SelectedItem;

                rutTextBox.Text = cliente.RUT;
                nombreTextBox.Text = cliente.Nombre;
                direccionTextBox.Text = cliente.Direccion;
                giroTextBox.Text = cliente.Giro;

                rutTextBox.Enabled = false;
                agregarButton.Enabled = false;
                actualizarButton.Enabled = true;
                eliminarButton.Enabled = true;
                registrarAbonoButton.Enabled = true;

                ActualizarInformacionCliente(cliente.RUT);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al cargar datos del cliente");
                MessageBox.Show("Error al cargar los datos del cliente.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ValidarSoloNumeros(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && e.KeyChar != '.')
            {
                e.Handled = true;
            }

            // Solo permitir un punto decimal
            if (e.KeyChar == '.' && (sender as TextBox).Text.IndexOf('.') > -1)
            {
                e.Handled = true;
            }
        }
        #endregion

        #region Operaciones CRUD
        private void AgregarButton_Click(object sender, EventArgs e)
        {
            try
            {
                if (!ValidarCampos()) return;

                var nuevoCliente = new Cliente
                {
                    RUT = rutTextBox.Text.Trim(),
                    Nombre = nombreTextBox.Text.Trim(),
                    Direccion = direccionTextBox.Text.Trim(),
                    Giro = giroTextBox.Text.Trim(),
                    Deuda = 0
                };

                _clienteService.AgregarCliente(nuevoCliente);
                MessageBox.Show("Cliente agregado exitosamente.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);

                LimpiarCampos();
                LoadClientes();
            }
            catch (ClienteDuplicadoException ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                rutTextBox.Focus();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al agregar cliente");
                MessageBox.Show("Error inesperado al agregar el cliente.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ActualizarButton_Click(object sender, EventArgs e)
        {
            try
            {
                if (!ValidarCampos()) return;

                var clienteActualizado = new Cliente
                {
                    RUT = rutTextBox.Text.Trim(),
                    Nombre = nombreTextBox.Text.Trim(),
                    Direccion = direccionTextBox.Text.Trim(),
                    Giro = giroTextBox.Text.Trim()
                };

                _clienteService.ActualizarCliente(clienteActualizado);
                MessageBox.Show("Cliente actualizado exitosamente.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);

                LimpiarCampos();
                LoadClientes();
            }
            catch (ClienteNotFoundException ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al actualizar cliente");
                MessageBox.Show("Error inesperado al actualizar el cliente.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void EliminarButton_Click(object sender, EventArgs e)
        {
            try
            {
                string rut = rutTextBox.Text.Trim();
                if (string.IsNullOrEmpty(rut))
                {
                    MessageBox.Show("Seleccione un cliente para eliminar.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                var resultado = MessageBox.Show(
                    "¿Está seguro que desea eliminar este cliente?\nEsta acción no se puede deshacer.",
                    "Confirmar Eliminación",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Warning);

                if (resultado == DialogResult.Yes)
                {
                    _clienteService.EliminarCliente(rut);
                    MessageBox.Show("Cliente eliminado exitosamente.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);

                    LimpiarCampos();
                    LoadClientes();
                }
            }
            catch (ClienteConVentasException ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al eliminar cliente");
                MessageBox.Show("Error inesperado al eliminar el cliente.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void RegistrarAbonoButton_Click(object sender, EventArgs e)
        {
            if (!double.TryParse(montoAbonoTextBox.Text, out double montoAbono) || montoAbono <= 0)
            {
                MessageBox.Show("Ingrese un monto válido mayor a cero.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var cliente = (Cliente)clientesComboBox.SelectedItem;
            if (cliente == null)
            {
                MessageBox.Show("Por favor seleccione un cliente.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                var deudaActual = _clienteService.CalcularDeudaTotal(cliente.RUT);
                if (montoAbono > deudaActual)
                {
                    MessageBox.Show($"El monto del abono (${montoAbono:N0}) no puede ser mayor que la deuda actual (${deudaActual:N0}).",
                        "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                double montoNegativo = -montoAbono;
                _clienteService.ActualizarDeudaCliente(cliente.RUT, montoNegativo);

                MessageBox.Show($"Abono por ${montoAbono:N0} registrado exitosamente.", "Éxito",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);

                montoAbonoTextBox.Clear();
                ActualizarInformacionCliente(cliente.RUT);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al registrar abono");
                MessageBox.Show($"Error al registrar abono: {ex.Message}", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        #endregion

        #region Métodos Auxiliares
        private bool ValidarCampos()
        {
            if (string.IsNullOrWhiteSpace(rutTextBox.Text))
            {
                MessageBox.Show("El RUT es obligatorio.", "Validación", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                rutTextBox.Focus();
                return false;
            }

            if (!ValidarFormatoRUT(rutTextBox.Text.Trim()))
            {
                MessageBox.Show("El formato del RUT no es válido.", "Validación", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                rutTextBox.Focus();
                return false;
            }

            if (string.IsNullOrWhiteSpace(nombreTextBox.Text))
            {
                MessageBox.Show("El nombre es obligatorio.", "Validación", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                nombreTextBox.Focus();
                return false;
            }

            if (string.IsNullOrWhiteSpace(direccionTextBox.Text))
            {
                MessageBox.Show("La dirección es obligatoria.", "Validación", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                direccionTextBox.Focus();
                return false;
            }

            if (string.IsNullOrWhiteSpace(giroTextBox.Text))
            {
                MessageBox.Show("El giro es obligatorio.", "Validación", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                giroTextBox.Focus();
                return false;
            }

            return true;
        }

        private bool ValidarFormatoRUT(string rut)
        {
            return System.Text.RegularExpressions.Regex.IsMatch(rut, @"^\d{1,2}\.\d{3}\.\d{3}[-][0-9kK]{1}$");
        }

        private void LimpiarCampos()
        {
            rutTextBox.Text = string.Empty;
            nombreTextBox.Text = string.Empty;
            direccionTextBox.Text = string.Empty;
            giroTextBox.Text = string.Empty;

            actualizarButton.Enabled = false;
            eliminarButton.Enabled = false;
            agregarButton.Enabled = true;
            rutTextBox.Enabled = true;
            registrarAbonoButton.Enabled = false;

            clientesComboBox.SelectedIndex = -1;
            deudaLabel.Text = string.Empty;
            ventasDataGridView.Rows.Clear();
            montoAbonoTextBox.Clear();
        }

        private void LoadClientes()
        {
            try
            {
                clientesComboBox.Items.Clear();
                var clientes = _clienteService.ObtenerTodosLosClientes(); // Los clientes ya vienen ordenados del servicio
                foreach (var cliente in clientes)
                {
                    clientesComboBox.Items.Add(cliente);
                }

                clientesComboBox.DisplayMember = "Nombre";
                clientesComboBox.ValueMember = "RUT";
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al cargar clientes");
                MessageBox.Show("Error al cargar la lista de clientes.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ActualizarInformacionCliente(string rut)
{
    try
    {
        var deudaTotal = _clienteService.CalcularDeudaTotal(rut);
        MessageBox.Show($"Deuda sin formatear: {deudaTotal}"); // Esta línea es temporal para debug
        deudaLabel.Text = $"Deuda Total: ${deudaTotal:N0}";
        CargarVentasCliente(rut);
    }
    catch (Exception ex)
    {
        _logger.LogError(ex, "Error al actualizar información del cliente");
        MessageBox.Show($"Error al actualizar información del cliente: {ex.Message}",
            "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
    }
}

        private void CargarVentasCliente(string rut)
        {
            try
            {
                ventasDataGridView.Rows.Clear();
                var ventas = _clienteService.ObtenerVentasCliente(rut);

                foreach (var venta in ventas)
                {
                    var fechaVenta = DateTime.Parse(venta.FechaVenta);
                    var diasMora = (DateTime.Now - fechaVenta).Days;
                    var estado = venta.PagadoConCredito == 1 ? "Pendiente" : "Pagado";

                    var rowIndex = ventasDataGridView.Rows.Add(
                        fechaVenta.ToString("dd/MM/yyyy"),
                        venta.NumeroGuia,
                        venta.Descripcion,
                        $"{venta.KilosNeto:N2}",
                        $"${venta.Total:N0}",
                        estado,
                        estado == "Pendiente" ? diasMora.ToString() : "-"
                    );

                    if (estado == "Pendiente")
                    {
                        if (diasMora > 30)
                            ventasDataGridView.Rows[rowIndex].DefaultCellStyle.BackColor = Color.LightCoral;
                        else if (diasMora > 15)
                            ventasDataGridView.Rows[rowIndex].DefaultCellStyle.BackColor = Color.LightYellow;
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al cargar ventas del cliente");
                MessageBox.Show($"Error al cargar las ventas del cliente: {ex.Message}",
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        #endregion
    }
}



