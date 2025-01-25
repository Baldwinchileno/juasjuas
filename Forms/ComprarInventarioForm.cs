using System;
using System.Drawing;
using System.Windows.Forms;
using System.Collections.Generic;
using AdminSERMAC.Models;
using AdminSERMAC.Services.Database;
using Microsoft.Extensions.Logging;
using AdminSERMAC.Core.Interfaces;
using AdminSERMAC.Services;

namespace AdminSERMAC.Forms
{
    public class ComprarInventarioForm : Form
    {
        private readonly IProductoDatabaseService _productoDatabaseService;
        private readonly IInventarioDatabaseService _inventarioDatabaseService;
        private readonly IProveedorService _proveedorService;
        private readonly ILogger<SQLiteService> _logger;
        private readonly IComprasDatabaseService _comprasService;

        // Controles del formulario
        private Label numeroCompraLabel;
        private TextBox numeroCompraTextBox;
        private Label proveedorLabel;
        private ComboBox proveedorComboBox;
        private Label fechaCompraLabel;
        private DateTimePicker fechaCompraPicker;
        private Label observacionesLabel;
        private TextBox observacionesTextBox;
        private DataGridView productosDataGridView;
        private Label totalLabel;
        private TextBox totalTextBox;
        private Button finalizarCompraButton;
        private Button cancelarButton;
        private decimal totalCompra = 0;

        private Dictionary<string, Producto> productosPorNombre = new Dictionary<string, Producto>();
        private Dictionary<string, Producto> productosPorCodigo = new Dictionary<string, Producto>();

        public ComprarInventarioForm(
            ILogger<SQLiteService> logger,
            IProductoDatabaseService productoDatabaseService,
            IInventarioDatabaseService inventarioDatabaseService,
            IProveedorService proveedorService,
            IComprasDatabaseService comprasService)
        {
            _logger = logger;
            _productoDatabaseService = productoDatabaseService;
            _inventarioDatabaseService = inventarioDatabaseService;
            _proveedorService = proveedorService;
            _comprasService = comprasService;

            CargarProductos();
            InitializeComponents();
            ConfigureEvents();
            CargarDatosIniciales();
        }

        private async void CargarProductos()
        {
            try
            {
                var productos = await _productoDatabaseService.GetAllAsync();
                productosPorNombre.Clear();
                productosPorCodigo.Clear();

                foreach (var producto in productos)
                {
                    productosPorNombre[producto.Nombre] = producto;
                    productosPorCodigo[producto.Codigo] = producto;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error al cargar productos: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void InitializeComponents()
        {
            this.Text = "Registro de Compras - SERMAC";
            this.Size = new Size(1000, 700);
            this.StartPosition = FormStartPosition.CenterScreen;

            // Panel principal que contendrá todo
            var mainPanel = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 3,
                Padding = new Padding(10),
            };

            // Configura las filas del TableLayoutPanel
            mainPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 150F)); // Header
            mainPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));  // Grid
            mainPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 60F));  // Footer

            // Panel superior
            var headerPanel = new Panel { Dock = DockStyle.Fill };

            numeroCompraLabel = new Label { Text = "N° Compra:", Location = new Point(20, 23), AutoSize = true };
            numeroCompraTextBox = new TextBox
            {
                Location = new Point(120, 20),
                Width = 100,
                ReadOnly = true
            };

            proveedorLabel = new Label { Text = "Proveedor:", Location = new Point(20, 53), AutoSize = true };
            proveedorComboBox = new ComboBox
            {
                Location = new Point(120, 50),
                Width = 300,
                DropDownStyle = ComboBoxStyle.DropDownList
            };

            fechaCompraLabel = new Label { Text = "Fecha:", Location = new Point(20, 83), AutoSize = true };
            fechaCompraPicker = new DateTimePicker
            {
                Location = new Point(120, 80),
                Width = 200,
                Format = DateTimePickerFormat.Short
            };

            observacionesLabel = new Label { Text = "Observaciones:", Location = new Point(20, 113), AutoSize = true };
            observacionesTextBox = new TextBox
            {
                Location = new Point(120, 110),
                Width = 300,
                Multiline = true,
                Height = 40
            };

            headerPanel.Controls.AddRange(new Control[] {
        numeroCompraLabel, numeroCompraTextBox,
        proveedorLabel, proveedorComboBox,
        fechaCompraLabel, fechaCompraPicker,
        observacionesLabel, observacionesTextBox
    });

            // DataGridView
            productosDataGridView = new DataGridView
            {
                Dock = DockStyle.Fill,
                AutoGenerateColumns = false,
                AllowUserToAddRows = true,
                AllowUserToDeleteRows = true,
                BackgroundColor = Color.White,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
            };

            ConfigurarColumnasGrid();

            // Panel inferior
            var bottomPanel = new Panel { Dock = DockStyle.Fill };

            totalLabel = new Label
            {
                Text = "Total:",
                Location = new Point(500, 15),
                AutoSize = true,
                Font = new Font("Segoe UI", 12, FontStyle.Bold)
            };

            totalTextBox = new TextBox
            {
                Location = new Point(580, 15),
                Width = 150,
                ReadOnly = true,
                Font = new Font("Segoe UI", 12)
            };

            finalizarCompraButton = new Button
            {
                Text = "Finalizar Compra",
                Location = new Point(750, 10),
                Size = new Size(120, 40),
                BackColor = Color.FromArgb(0, 122, 204),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand
            };

            cancelarButton = new Button
            {
                Text = "Cancelar",
                Location = new Point(880, 10),
                Size = new Size(90, 40),
                BackColor = Color.FromArgb(204, 0, 0),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand
            };

            bottomPanel.Controls.AddRange(new Control[] {
        totalLabel,
        totalTextBox,
        finalizarCompraButton,
        cancelarButton
    });

            // Agregar los paneles al TableLayoutPanel
            mainPanel.Controls.Add(headerPanel, 0, 0);
            mainPanel.Controls.Add(productosDataGridView, 0, 1);
            mainPanel.Controls.Add(bottomPanel, 0, 2);

            // Agregar el TableLayoutPanel al formulario
            this.Controls.Add(mainPanel);
        }

        private void ConfigurarColumnasGrid()
        {
            productosDataGridView.Columns.Clear();
            productosDataGridView.AutoGenerateColumns = false;

            // Agregar columnas
            productosDataGridView.Columns.AddRange(new DataGridViewColumn[]
            {
        new DataGridViewTextBoxColumn
        {
            Name = "Codigo",
            HeaderText = "Código",
            Width = 100
        },
        new DataGridViewTextBoxColumn
        {
            Name = "Descripcion",
            HeaderText = "Descripción",
            Width = 200
        },
        new DataGridViewTextBoxColumn
        {
            Name = "Cantidad",
            HeaderText = "Cantidad",
            Width = 80,
            DefaultCellStyle = new DataGridViewCellStyle
            {
                Alignment = DataGridViewContentAlignment.MiddleRight,
                ForeColor = Color.Black
            }
        },
        new DataGridViewTextBoxColumn
        {
            Name = "Kilos",
            HeaderText = "Kilos",
            Width = 80,
            DefaultCellStyle = new DataGridViewCellStyle
            {
                Alignment = DataGridViewContentAlignment.MiddleRight,
                Format = "N2",
                ForeColor = Color.Black
            }
        },
        new DataGridViewTextBoxColumn
        {
            Name = "PrecioUnitario",
            HeaderText = "Precio Unitario",
            Width = 100,
            DefaultCellStyle = new DataGridViewCellStyle
            {
                Alignment = DataGridViewContentAlignment.MiddleRight,
                Format = "N0",
                ForeColor = Color.Black
            }
        },
        new DataGridViewTextBoxColumn
        {
            Name = "Subtotal",
            HeaderText = "Subtotal",
            Width = 100,
            ReadOnly = true,
            DefaultCellStyle = new DataGridViewCellStyle
            {
                Alignment = DataGridViewContentAlignment.MiddleRight,
                Format = "N0",
                ForeColor = Color.Black
            }
        },
        new DataGridViewTextBoxColumn
        {
            Name = "FechaVencimiento",
            HeaderText = "Fecha Vencimiento",
            Width = 120,
            DefaultCellStyle = new DataGridViewCellStyle
            {
                ForeColor = Color.Black
            }
        }
            });

            // Estilos del DataGridView
            productosDataGridView.EnableHeadersVisualStyles = false;
            productosDataGridView.ColumnHeadersDefaultCellStyle = new DataGridViewCellStyle
            {
                BackColor = Color.FromArgb(50, 50, 50),
                ForeColor = Color.White,
                SelectionBackColor = Color.FromArgb(50, 50, 50),
                Font = new Font("Segoe UI", 10, FontStyle.Regular)
            };

            productosDataGridView.DefaultCellStyle = new DataGridViewCellStyle
            {
                BackColor = Color.White,
                ForeColor = Color.Black,
                SelectionBackColor = Color.FromArgb(220, 230, 240),
                SelectionForeColor = Color.Black,
                Font = new Font("Segoe UI", 9, FontStyle.Regular)
            };

            productosDataGridView.AlternatingRowsDefaultCellStyle = new DataGridViewCellStyle
            {
                BackColor = Color.FromArgb(240, 240, 240),
                ForeColor = Color.Black,
                SelectionBackColor = Color.FromArgb(220, 230, 240),
                SelectionForeColor = Color.Black
            };

            productosDataGridView.RowHeadersVisible = false;
            productosDataGridView.AllowUserToAddRows = true;
            productosDataGridView.AllowUserToDeleteRows = true;
            productosDataGridView.MultiSelect = false;
            productosDataGridView.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            productosDataGridView.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
            productosDataGridView.BorderStyle = BorderStyle.None;
        }

        private void ConfigureEvents()
        {
            finalizarCompraButton.Click += FinalizarCompraButton_Click;
            cancelarButton.Click += cancelarButton_Click;
            productosDataGridView.CellEndEdit += ProductosDataGridView_CellEndEdit;
            productosDataGridView.CellValidating += ProductosDataGridView_CellValidating;
            productosDataGridView.CellValueChanged += ProductosDataGridView_CellValueChanged;
            numeroCompraTextBox.Click += NumeroCompraTextBox_Click;
        }

        private async void CargarDatosIniciales()
        {
            try
            {
                // Cargar proveedores
                var proveedores = await _proveedorService.GetAllProveedores();
                proveedorComboBox.DataSource = proveedores;
                proveedorComboBox.DisplayMember = "Nombre";
                proveedorComboBox.ValueMember = "Id";

                // Establecer fecha actual
                fechaCompraPicker.Value = DateTime.Now;

                // Generar número de compra
                var ultimoNumero = await _comprasService.GetNextNumeroCompra();
                numeroCompraTextBox.Text = ultimoNumero.ToString();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error al cargar datos iniciales: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ProductosDataGridView_CellValidating(object sender, DataGridViewCellValidatingEventArgs e)
        {
            var row = productosDataGridView.Rows[e.RowIndex];

            if (e.ColumnIndex == productosDataGridView.Columns["Cantidad"].Index)
            {
                if (!int.TryParse(e.FormattedValue.ToString(), out int cantidad) || cantidad <= 0)
                {
                    e.Cancel = true;
                    MessageBox.Show("Por favor ingrese una cantidad válida mayor a 0.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            else if (e.ColumnIndex == productosDataGridView.Columns["Kilos"].Index)
            {
                if (!decimal.TryParse(e.FormattedValue.ToString(), out decimal kilos) || kilos <= 0)
                {
                    e.Cancel = true;
                    MessageBox.Show("Por favor ingrese un peso válido mayor a 0.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            else if (e.ColumnIndex == productosDataGridView.Columns["PrecioUnitario"].Index)
            {
                if (!decimal.TryParse(e.FormattedValue.ToString(), out decimal precio) || precio <= 0)
                {
                    e.Cancel = true;
                    MessageBox.Show("Por favor ingrese un precio válido mayor a 0.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            else if (e.ColumnIndex == productosDataGridView.Columns["FechaVencimiento"].Index)
            {
                if (!DateTime.TryParse(e.FormattedValue.ToString(), out DateTime fecha))
                {
                    e.Cancel = true;
                    MessageBox.Show("Por favor ingrese una fecha válida (dd/mm/yyyy).", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
        }

        private void ProductosDataGridView_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0) return;

            var row = productosDataGridView.Rows[e.RowIndex];

            // Actualizar subtotal cuando cambia kilos o precio
            if (e.ColumnIndex == productosDataGridView.Columns["Kilos"].Index ||
                e.ColumnIndex == productosDataGridView.Columns["PrecioUnitario"].Index)
            {
                if (decimal.TryParse(row.Cells["Kilos"].Value?.ToString(), out decimal kilos) &&
                    decimal.TryParse(row.Cells["PrecioUnitario"].Value?.ToString(), out decimal precio))
                {
                    row.Cells["Subtotal"].Value = (kilos * precio).ToString("N0");
                    CalcularTotal();
                }
            }
        }

        private void CalcularSubtotal(DataGridViewRow row)
        {
            if (int.TryParse(row.Cells["Cantidad"].Value?.ToString(), out int cantidad) &&
                decimal.TryParse(row.Cells["PrecioUnitario"].Value?.ToString(), out decimal precio))
            {
                row.Cells["Subtotal"].Value = (cantidad * precio).ToString("N2");
            }
        }

        private void CalcularTotal()
        {
            totalCompra = 0;
            foreach (DataGridViewRow row in productosDataGridView.Rows)
            {
                if (!row.IsNewRow && decimal.TryParse(row.Cells["Subtotal"].Value?.ToString(), out decimal subtotal))
                {
                    totalCompra += subtotal;
                }
            }
            totalTextBox.Text = totalCompra.ToString("N2");
        }

        private void cancelarButton_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private async void ProductosDataGridView_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0) return;

            var row = productosDataGridView.Rows[e.RowIndex];

            if (e.ColumnIndex == productosDataGridView.Columns["Codigo"].Index)
            {
                var codigo = row.Cells["Codigo"].Value?.ToString();
                if (!string.IsNullOrEmpty(codigo))
                {
                    var producto = await _productoDatabaseService.GetByCodigo(codigo);
                    if (producto != null)
                    {
                        row.Cells["Descripcion"].Value = producto.Nombre;
                    }
                    else
                    {
                        MessageBox.Show("Producto no encontrado.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }
            }
            else if (e.ColumnIndex == productosDataGridView.Columns["Descripcion"].Index)
            {
                var descripcion = row.Cells["Descripcion"].Value?.ToString();
                if (!string.IsNullOrEmpty(descripcion) && productosPorNombre.ContainsKey(descripcion))
                {
                    row.Cells["Codigo"].Value = productosPorNombre[descripcion].Codigo;
                }
            }
        }

        private void NumeroCompraTextBox_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("¿Desea limpiar el formulario para una nueva compra?",
                "Nueva Compra", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                LimpiarFormulario();
            }
        }

        private void LimpiarFormulario()
        {
            CargarDatosIniciales();
            proveedorComboBox.SelectedIndex = -1;
            observacionesTextBox.Clear();
            productosDataGridView.Rows.Clear();
            totalCompra = 0;
            totalTextBox.Text = "0.00";
        }

        private bool ValidarCompra()
        {
            if (proveedorComboBox.SelectedItem == null)
            {
                MessageBox.Show("Debe seleccionar un proveedor", "Validación", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }

            // Si solo existe la fila nueva (vacía)
            if (productosDataGridView.Rows.Count <= 1 ||
                productosDataGridView.Rows.Cast<DataGridViewRow>()
                    .All(row => row.IsNewRow || row.Cells["Codigo"].Value == null))
            {
                return true; // Permitir cerrar si no hay datos ingresados
            }

            // Validar solo las filas que tienen algún dato
            foreach (DataGridViewRow row in productosDataGridView.Rows)
            {
                if (row.IsNewRow || row.Cells["Codigo"].Value == null) continue;

                if (string.IsNullOrEmpty(row.Cells["Codigo"].Value?.ToString()))
                {
                    MessageBox.Show($"Falta el código en la línea {row.Index + 1}", "Validación", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return false;
                }

                if (!decimal.TryParse(row.Cells["Kilos"].Value?.ToString(), out decimal kilos) || kilos <= 0)
                {
                    MessageBox.Show($"Los kilos en la línea {row.Index + 1} deben ser un número mayor a 0",
                        "Validación", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return false;
                }

                if (!decimal.TryParse(row.Cells["PrecioUnitario"].Value?.ToString(), out decimal precio) || precio <= 0)
                {
                    MessageBox.Show($"El precio en la línea {row.Index + 1} debe ser un número mayor a 0",
                        "Validación", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return false;
                }
            }

            return true;
        }

        private async void FinalizarCompraButton_Click(object sender, EventArgs e)
        {
            if (!ValidarCompra())
            {
                // Si no hay datos y el usuario quiere salir
                if (productosDataGridView.Rows.Count <= 1 ||
                    productosDataGridView.Rows.Cast<DataGridViewRow>()
                        .All(row => row.IsNewRow || row.Cells["Codigo"].Value == null))
                {
                    this.Close();
                    return;
                }
                return;
            }

            // Si no hay datos que guardar, simplemente cerrar
            if (productosDataGridView.Rows.Count <= 1 ||
                productosDataGridView.Rows.Cast<DataGridViewRow>()
                    .All(row => row.IsNewRow || row.Cells["Codigo"].Value == null))
            {
                this.Close();
                return;
            }

            try
            {
                finalizarCompraButton.Enabled = false;
                Cursor = Cursors.WaitCursor;

                var compra = new Compra
                {
                    ProveedorId = (int)proveedorComboBox.SelectedValue,
                    FechaCompra = fechaCompraPicker.Value.ToString("yyyy-MM-dd"),
                    Total = totalCompra,
                    Observaciones = observacionesTextBox.Text,
                    Estado = "Completada",
                    Detalles = new List<DetalleCompra>()
                };

                // Obtener el número de compra
                compra.NumeroCompra = await _comprasService.GetNextNumeroCompra();

                foreach (DataGridViewRow row in productosDataGridView.Rows)
                {
                    // Saltar filas vacías o nuevas
                    if (row.IsNewRow || row.Cells["Codigo"].Value == null ||
                        string.IsNullOrEmpty(row.Cells["Codigo"].Value.ToString())) continue;

                    var detalle = new DetalleCompra
                    {
                        NumeroCompra = compra.NumeroCompra,
                        CodigoProducto = row.Cells["Codigo"].Value.ToString(),
                        Cantidad = Convert.ToInt32(row.Cells["Cantidad"].Value),
                        PrecioUnitario = Convert.ToDecimal(row.Cells["PrecioUnitario"].Value),
                        Subtotal = Convert.ToDecimal(row.Cells["Subtotal"].Value),
                        FechaVencimiento = row.Cells["FechaVencimiento"].Value?.ToString(),
                        Kilos = Convert.ToDouble(row.Cells["Kilos"].Value)
                    };

                    compra.Detalles.Add(detalle);

                    // Verificar y actualizar inventario
                    var inventarioActual = await _inventarioDatabaseService.GetInventarioPorCodigoAsync(detalle.CodigoProducto);
                    if (inventarioActual.Rows.Count > 0)
                    {
                        // Si el producto existe, actualizar sumando las cantidades (por eso el signo negativo)
                        await _inventarioDatabaseService.ActualizarInventarioAsync(
                            detalle.CodigoProducto,
                            -detalle.Cantidad, // Negativo porque ActualizarInventarioAsync resta
                            -Convert.ToDouble(row.Cells["Kilos"].Value) // Negativo por la misma razón
                        );
                    }
                    else
                    {
                        // Si el producto no existe, agregar como nuevo
                        await _inventarioDatabaseService.AddProductoAsync(
                            detalle.CodigoProducto,
                            row.Cells["Descripcion"].Value.ToString(),
                            detalle.Cantidad,
                            Convert.ToDouble(row.Cells["Kilos"].Value),
                            fechaCompraPicker.Value.ToString("yyyy-MM-dd"),
                            fechaCompraPicker.Value.ToString("yyyy-MM-dd"),
                            detalle.FechaVencimiento
                        );
                    }

                    // Agregar el detalle a la tabla DetallesCompra
                    await _comprasService.AddDetalleCompra(detalle);
                }

                if (compra.Detalles.Any())
                {
                    await _comprasService.CreateCompra(compra);
                    MessageBox.Show(
                        "Compra registrada exitosamente.\nHaga clic en el número de compra para iniciar una nueva.",
                        "Éxito",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                    LimpiarFormulario();
                }
                else
                {
                    this.Close();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"Error al registrar la compra: {ex.Message}",
                    "Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                _logger.LogError(ex, "Error al registrar la compra");
            }
            finally
            {
                finalizarCompraButton.Enabled = true;
                Cursor = Cursors.Default;
            }
        }
    }
}