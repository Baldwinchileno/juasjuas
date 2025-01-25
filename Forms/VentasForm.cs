using AdminSERMAC.Services;
using System.Data.SQLite;
using AdminSERMAC.Models;
using Microsoft.Extensions.Logging;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using ClosedXML.Excel;
using System.Drawing.Printing;
using System.Windows.Forms;
using System.Diagnostics;


public class VentasForm : Form
{
    private Label numeroGuiaLabel;
    private TextBox numeroGuiaTextBox;
    private Label rutLabel;
    private ComboBox rutComboBox;
    private Label clienteLabel;
    private ComboBox clienteComboBox;
    private Label direccionLabel;
    private TextBox direccionTextBox;
    private Label giroLabel;
    private TextBox giroTextBox;
    private Label fechaEmisionLabel;
    private DateTimePicker fechaEmisionPicker;
    private Label totalVentaLabel;
    private TextBox totalVentaTextBox;
    private Label stockLabel;
    private Label preciosLabel;

    private Button exportarExcelButton;
    private Button imprimirButton;


    private DataGridView ventasDataGridView;
    private Button finalizarButton;
    private CheckBox pagarConCreditoCheckBox;

    private SQLiteService sqliteService;
    private double totalVenta = 0;
    private Dictionary<string, Cliente> clientesDict = new Dictionary<string, Cliente>();
    private Dictionary<string, Producto> productosPorNombre = new Dictionary<string, Producto>();
    private Dictionary<string, Producto> productosPorCodigo = new Dictionary<string, Producto>();
    private readonly ILogger<VentasForm> _logger;



    public VentasForm(ILogger<SQLiteService> logger)
    {
        this.Text = "Gestión de Ventas";
        this.Width = 1000;
        this.Height = 800;

        _logger = LoggerFactory.Create(builder => builder.AddConsole())
                          .CreateLogger<VentasForm>();

        sqliteService = new SQLiteService(logger);

        CargarProductos(); // Añade esta línea antes de InitializeComponents
        InitializeComponents();

        exportarExcelButton.Enabled = true;
        imprimirButton.Enabled = true;

        ConfigureEvents();
        CargarNumeroGuia();
        CargarClientes();
    }

    private void InitializeComponents()
    {
        // Número de guía
        numeroGuiaLabel = new Label { Text = "Número de Guía:", Location = new Point(20, 20), Width = 100 };
        numeroGuiaTextBox = new TextBox
        {
            Location = new Point(150, 20),
            Width = 100,
            ReadOnly = true,
            Cursor = Cursors.Hand,
            BackColor = Color.LightYellow
        };

        // Configura el form para que permita resize
        this.MinimumSize = new Size(1000, 800);
        this.AutoScaleMode = AutoScaleMode.None;


        // Eventos del numeroGuiaTextBox
        numeroGuiaTextBox.MouseEnter += (s, e) => numeroGuiaTextBox.BackColor = Color.Yellow;
        numeroGuiaTextBox.MouseLeave += (s, e) => numeroGuiaTextBox.BackColor = Color.LightYellow;
        numeroGuiaTextBox.Click += NumeroGuiaTextBox_Click;

        // Fecha de emisión (mover a la derecha)
        fechaEmisionLabel = new Label { Text = "Fecha de Emisión:", Location = new Point(400, 20), Width = 120 };
        fechaEmisionPicker = new DateTimePicker { Location = new Point(520, 20), Width = 200 };

        // RUT (ajustar posición y estilo)
        rutLabel = new Label { Text = "RUT:", Location = new Point(20, 60), Width = 100 };
        rutComboBox = new ComboBox
        {
            Location = new Point(150, 60),
            Width = 200,
            DropDownStyle = ComboBoxStyle.DropDown,
            AutoCompleteMode = AutoCompleteMode.SuggestAppend,
            AutoCompleteSource = AutoCompleteSource.CustomSource
        };

        // Cliente
        clienteLabel = new Label { Text = "Nombre:", Location = new Point(20, 90), Width = 100 };
        clienteComboBox = new ComboBox
        {
            Location = new Point(150, 90),
            Width = 200,
            DropDownStyle = ComboBoxStyle.DropDown,
            AutoCompleteMode = AutoCompleteMode.SuggestAppend,
            AutoCompleteSource = AutoCompleteSource.CustomSource
        };

        // Dirección
        direccionLabel = new Label { Text = "Dirección:", Location = new Point(20, 120), Width = 100 };
        direccionTextBox = new TextBox { Location = new Point(150, 120), Width = 200, ReadOnly = true };

        // Giro
        giroLabel = new Label { Text = "Giro:", Location = new Point(20, 150), Width = 100 };
        giroTextBox = new TextBox { Location = new Point(150, 150), Width = 200, ReadOnly = true };

        // DataGridView
        ventasDataGridView = new DataGridView
        {
            Location = new Point(20, 190),
            Size = new Size(940, 350),
            AllowUserToAddRows = true,
            AllowUserToDeleteRows = true,
            AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        };

        stockLabel = new Label
        {
            Location = new Point(20, 550),
            Width = 300,
            Height = 30,
            Font = new Font("Arial", 12, FontStyle.Bold),
            ForeColor = Color.DarkGreen
        };

        this.Controls.Add(stockLabel);

        preciosLabel = new Label
        {
            Location = new Point(20, 580), // 30 pixels abajo del stockLabel
            Width = 300,
            Height = 30,
            Font = new Font("Arial", 12, FontStyle.Bold),
            ForeColor = Color.DarkBlue
        };

        this.Controls.Add(preciosLabel);



        exportarExcelButton = new Button
        {
            Text = "Exportar a Excel",
            Location = new Point(600, 620),
            Width = 150,
            Height = 30,
            BackColor = Color.FromArgb(76, 175, 80),
            ForeColor = Color.White
        };

        imprimirButton = new Button
        {
            Text = "Imprimir Guía",
            Location = new Point(800, 620),
            Width = 150,
            Height = 30,
            BackColor = Color.FromArgb(63, 81, 181),
            ForeColor = Color.White
        };

        this.Controls.AddRange(new Control[] { exportarExcelButton, imprimirButton });

        // Configurar las columnas del DataGridView
        ventasDataGridView.Columns.Add("Codigo", "Código");
        ventasDataGridView.Columns.Add("Descripcion", "Descripción");
        ventasDataGridView.Columns.Add("Marca", "Marca");
        ventasDataGridView.Columns.Add("Unidades", "Unidades");
        ventasDataGridView.Columns.Add("Bandejas", "Bandejas");
        ventasDataGridView.Columns.Add("KilosBruto", "Kilos Bruto");
        ventasDataGridView.Columns.Add("KilosNeto", "Kilos Neto");
        ventasDataGridView.Columns.Add("Precio", "Precio");
        ventasDataGridView.Columns.Add("Total", "Total");

        ventasDataGridView.DefaultValuesNeeded += VentasDataGridView_DefaultValuesNeeded;

        // Estilo del DataGridView
        ventasDataGridView.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
        ventasDataGridView.EnableHeadersVisualStyles = false;
        ventasDataGridView.ColumnHeadersDefaultCellStyle.BackColor = Color.LightGray;

        // Controles inferiores
        pagarConCreditoCheckBox = new CheckBox
        {
            Text = "Pago con crédito",
            Location = new Point(600, 560),
            Width = 150
        };

        totalVentaLabel = new Label { Text = "Total Venta:", Location = new Point(400, 560), Width = 100 };
        totalVentaTextBox = new TextBox { Location = new Point(500, 560), Width = 150, ReadOnly = true };

        finalizarButton = new Button
        {
            Text = "Finalizar Venta",
            Location = new Point(800, 560),
            Width = 150,
            Height = 40,
            BackColor = Color.FromArgb(0, 122, 204),
            ForeColor = Color.White
        };

        // Configurar formato de columnas
        ConfigurarColumnasGrid();

        // Ajustar tamaños y propiedades del form
        this.MinimumSize = new Size(1000, 800);
        this.AutoScaleMode = AutoScaleMode.None;

        // Modificar el DataGridView para que se expanda
        ventasDataGridView.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;

        // Anclar controles inferiores al fondo
        stockLabel.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
        preciosLabel.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
        pagarConCreditoCheckBox.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
        totalVentaLabel.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
        totalVentaTextBox.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
        finalizarButton.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
        exportarExcelButton.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
        imprimirButton.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;

        // Anclar controles superiores a la derecha
        fechaEmisionLabel.Anchor = AnchorStyles.Top | AnchorStyles.Right;
        fechaEmisionPicker.Anchor = AnchorStyles.Top | AnchorStyles.Right;

        // Agregar todos los controles al formulario
        this.Controls.AddRange(new Control[]
        {
        numeroGuiaLabel, numeroGuiaTextBox,
        rutLabel, rutComboBox,
        clienteLabel, clienteComboBox,
        direccionLabel, direccionTextBox,
        giroLabel, giroTextBox,
        fechaEmisionLabel, fechaEmisionPicker,
        ventasDataGridView,
        finalizarButton, pagarConCreditoCheckBox,
        totalVentaLabel, totalVentaTextBox
        });
    }

    private void CargarProductos()
    {
        try
        {
            var productos = sqliteService.GetProductos(); // Asegúrate de tener este método en SQLiteService
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

    private void ConfigureEvents()
    {
        rutComboBox.SelectedIndexChanged += RutComboBox_SelectedIndexChanged;
        clienteComboBox.SelectedIndexChanged += ClienteComboBox_SelectedIndexChanged;
        ventasDataGridView.CellEndEdit += VentasDataGridView_CellEndEdit;
        finalizarButton.Click += FinalizarButton_Click;
        ventasDataGridView.CellValueChanged += VentasDataGridView_CellValueChanged;
        exportarExcelButton.Click += ExportarExcelButton_Click;
        imprimirButton.Click += ImprimirButton_Click;
        ventasDataGridView.EditingControlShowing += VentasDataGridView_EditingControlShowing;
        ventasDataGridView.CellValueChanged += VentasDataGridView_CellValueChanged;
        ventasDataGridView.SelectionChanged += VentasDataGridView_SelectionChanged;
        numeroGuiaTextBox.Click += NumeroGuiaTextBox_Click;
    }

    private void NumeroGuiaTextBox_Click(object sender, EventArgs e)
    {
        if (MessageBox.Show("¿Desea limpiar el formulario para una nueva venta?",
            "Confirmar Nueva Venta", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
        {
            LimpiarFormulario();
        }
    }

    public class PrintService
    {
        private readonly DataGridView _gridView;
        private readonly Venta _venta;
        private readonly Cliente _cliente;
        private Font _titleFont = new Font("Arial", 16, FontStyle.Bold);
        private Font _headerFont = new Font("Arial", 12, FontStyle.Bold);
        private Font _normalFont = new Font("Arial", 10);
        private int _currentY = 40;

        public PrintService(DataGridView gridView, Venta venta, Cliente cliente)
        {
            _gridView = gridView;
            _venta = venta;
            _cliente = cliente;
        }

        public void PrintGuia()
        {
            PrintDocument pd = new PrintDocument();
            pd.PrintPage += PrintPage;

            PrintPreviewDialog printPreviewDialog = new PrintPreviewDialog();
            printPreviewDialog.Document = pd;
            printPreviewDialog.ShowDialog();
        }

        private void PrintPage(object sender, PrintPageEventArgs e)
        {
            Graphics g = e.Graphics;
            int leftMargin = 50;
            int width = 750;

            // Título
            g.DrawString("SERMAC", _titleFont, Brushes.Black, leftMargin, _currentY);
            _currentY += 30;

            // Información de la guía
            string[] headers = {
            $"Guía N°: {_venta.NumeroGuia}",
            $"Fecha: {_venta.FechaVenta}",
            $"Cliente: {_cliente.Nombre}",
            $"RUT: {_cliente.RUT}",
            $"Dirección: {_cliente.Direccion}",
            $"Giro: {_cliente.Giro}"
        };

            foreach (string header in headers)
            {
                g.DrawString(header, _headerFont, Brushes.Black, leftMargin, _currentY);
                _currentY += 25;
            }

            _currentY += 20;

            // Encabezados de la tabla
            string[] columnHeaders = { "Código", "Descripción", "Marca", "Bandejas", "Kilos", "Precio", "Total" };
            int[] columnWidths = { 80, 200, 100, 80, 80, 80, 100 };
            int xPos = leftMargin;

            // Dibujar línea superior
            g.DrawLine(Pens.Black, leftMargin, _currentY, leftMargin + width, _currentY);

            // Encabezados
            _currentY += 5;
            for (int i = 0; i < columnHeaders.Length; i++)
            {
                g.DrawString(columnHeaders[i], _headerFont, Brushes.Black, xPos, _currentY);
                xPos += columnWidths[i];
            }
            _currentY += 25;

            // Dibujar línea después de encabezados
            g.DrawLine(Pens.Black, leftMargin, _currentY, leftMargin + width, _currentY);
            _currentY += 5;

            // Datos
            foreach (DataGridViewRow row in _gridView.Rows)
            {
                if (!row.IsNewRow)
                {
                    xPos = leftMargin;
                    string[] values = {
                row.Cells["Codigo"].Value?.ToString(),
                row.Cells["Descripcion"].Value?.ToString(),
                row.Cells["Marca"].Value?.ToString(), // Agregada la marca
                row.Cells["Bandejas"].Value?.ToString(),
                $"{Convert.ToDouble(row.Cells["KilosNeto"].Value):N2}",
                $"${Convert.ToDouble(row.Cells["Precio"].Value):N0}",
                $"${Convert.ToDouble(row.Cells["Total"].Value):N0}"
                };

                    for (int i = 0; i < values.Length; i++)
                    {
                        g.DrawString(values[i], _normalFont, Brushes.Black, xPos, _currentY);
                        xPos += columnWidths[i];
                    }
                    _currentY += 20;
                }
            }

            // Línea final
            g.DrawLine(Pens.Black, leftMargin, _currentY, leftMargin + width, _currentY);
            _currentY += 20;

            // Total
            double total = _gridView.Rows.Cast<DataGridViewRow>()
                .Where(r => !r.IsNewRow)
                .Sum(r => Convert.ToDouble(r.Cells["Total"].Value));

            g.DrawString($"Total: ${total:N0}", _headerFont, Brushes.Black,
                leftMargin + width - 150, _currentY);
        }
    }

    private void VentasDataGridView_SelectionChanged(object sender, EventArgs e)
    {
        if (ventasDataGridView.CurrentRow != null && !ventasDataGridView.CurrentRow.IsNewRow)
        {
            string codigo = ventasDataGridView.CurrentRow.Cells["Codigo"].Value?.ToString();
            if (!string.IsNullOrEmpty(codigo) && productosPorCodigo.ContainsKey(codigo))
            {
                var producto = productosPorCodigo[codigo];
                if (rutComboBox.SelectedItem != null)
                {
                    MostrarPreciosHistoricos(codigo, rutComboBox.SelectedItem.ToString());
                }
            }
        }
    }

    private void ExportarExcelButton_Click(object sender, EventArgs e)
    {
        if (!ValidarVenta()) return;

        using (SaveFileDialog saveDialog = new SaveFileDialog())
        {
            saveDialog.Filter = "Excel files (*.xlsx)|*.xlsx";
            saveDialog.FileName = $"Guia_{numeroGuiaTextBox.Text}.xlsx";

            if (saveDialog.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    using (var workbook = new XLWorkbook())
                    {
                        var worksheet = workbook.Worksheets.Add("Detalle Venta");

                        // Encabezados
                        worksheet.Cell("A1").SetValue("SERMAC");
                        worksheet.Cell("A2").SetValue($"Guía N°: {numeroGuiaTextBox.Text}");
                        worksheet.Cell("A3").SetValue($"Fecha: {fechaEmisionPicker.Value:dd/MM/yyyy}");
                        worksheet.Cell("A4").SetValue($"Cliente: {clienteComboBox.Text}");
                        worksheet.Cell("A5").SetValue($"RUT: {rutComboBox.Text}");
                        worksheet.Cell("A6").SetValue($"Dirección: {direccionTextBox.Text}");
                        worksheet.Cell("A7").SetValue($"Giro: {giroTextBox.Text}");

                        // Encabezados de la tabla
                        worksheet.Cell("A9").SetValue("Código");
                        worksheet.Cell("B9").SetValue("Descripción");
                        worksheet.Cell("C9").SetValue("Bandejas");
                        worksheet.Cell("D9").SetValue("Kilos Neto");
                        worksheet.Cell("E9").SetValue("Precio sin IVA");
                        worksheet.Cell("F9").SetValue("IVA");
                        worksheet.Cell("G9").SetValue("Precio con IVA");
                        worksheet.Cell("H9").SetValue("Total");

                        // Datos
                        int row = 10;
                        foreach (DataGridViewRow dgvRow in ventasDataGridView.Rows)
                        {
                            if (!dgvRow.IsNewRow)
                            {
                                double precio = Convert.ToDouble(dgvRow.Cells["Precio"].Value);
                                double precioSinIVA = Math.Round(precio / 1.19, 2);
                                double iva = precio - precioSinIVA;
                                double kilosNeto = Convert.ToDouble(dgvRow.Cells["KilosNeto"].Value);
                                double total = precioSinIVA * kilosNeto;

                                worksheet.Cell($"A{row}").SetValue(dgvRow.Cells["Codigo"].Value?.ToString() ?? "");
                                worksheet.Cell($"B{row}").SetValue(dgvRow.Cells["Descripcion"].Value?.ToString() ?? "");
                                worksheet.Cell($"C{row}").SetValue(Convert.ToDouble(dgvRow.Cells["Bandejas"].Value));
                                worksheet.Cell($"D{row}").SetValue(kilosNeto);
                                worksheet.Cell($"E{row}").SetValue(precioSinIVA);
                                worksheet.Cell($"F{row}").SetValue(iva);
                                worksheet.Cell($"G{row}").SetValue(precio);
                                worksheet.Cell($"H{row}").SetValue(total);

                                row++;
                            }
                        }

                        // Formato
                        var rango = worksheet.Range("A9:H9");
                        rango.Style.Fill.BackgroundColor = XLColor.LightGray;
                        rango.Style.Font.Bold = true;

                        // Formato para columnas numéricas
                        worksheet.Range($"E10:H{row - 1}").Style.NumberFormat.NumberFormatId = 4; // Formato de moneda

                        worksheet.Columns().AdjustToContents();

                        workbook.SaveAs(saveDialog.FileName);
                        System.Diagnostics.Process.Start(new ProcessStartInfo(saveDialog.FileName) { UseShellExecute = true });
                    }

                    MessageBox.Show("Excel exportado exitosamente", "Éxito",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Error al exportar: {ex.Message}", "Error",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }
    }

    private void ImprimirButton_Click(object sender, EventArgs e)
    {
        if (!ValidarVenta()) return;

        try
        {
            var cliente = clientesDict[rutComboBox.SelectedItem.ToString()];
            var venta = new Venta
            {
                NumeroGuia = int.Parse(numeroGuiaTextBox.Text), // Convert string to int
                FechaVenta = fechaEmisionPicker.Value.ToString("dd/MM/yyyy"),
                RUT = cliente.RUT,
                ClienteNombre = cliente.Nombre,
                PagadoConCredito = pagarConCreditoCheckBox.Checked ? 1 : 0
            };

            var printService = new PrintService(ventasDataGridView, venta, cliente);
            printService.PrintGuia();
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Error al imprimir: {ex.Message}", "Error",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }


    private void CargarNumeroGuia()
    {
        int numeroGuia = sqliteService.GetUltimoNumeroGuia() + 1;
        numeroGuiaTextBox.Text = numeroGuia.ToString();
    }

    private void CargarClientes()
    {
        clientesDict.Clear();
        rutComboBox.Items.Clear();
        clienteComboBox.Items.Clear();

        // Crear las colecciones de autocompletado
        var rutAutoComplete = new AutoCompleteStringCollection();
        var nombreAutoComplete = new AutoCompleteStringCollection();

        var clientes = sqliteService.GetClientes();
        foreach (var cliente in clientes)
        {
            clientesDict[cliente.RUT] = cliente;
            rutComboBox.Items.Add(cliente.RUT);
            clienteComboBox.Items.Add(cliente.Nombre);

            // Agregar a las colecciones de autocompletado
            rutAutoComplete.Add(cliente.RUT);
            nombreAutoComplete.Add(cliente.Nombre);
        }

        // Asignar las colecciones de autocompletado
        rutComboBox.AutoCompleteCustomSource = rutAutoComplete;
        clienteComboBox.AutoCompleteCustomSource = nombreAutoComplete;
    }

    private void RutComboBox_SelectedIndexChanged(object sender, EventArgs e)
    {
        if (rutComboBox.SelectedItem != null)
        {
            string rut = rutComboBox.SelectedItem.ToString();
            if (clientesDict.ContainsKey(rut))
            {
                var cliente = clientesDict[rut];
                clienteComboBox.SelectedItem = cliente.Nombre;
                ActualizarDatosCliente(cliente);
            }
        }
    }

    private void ClienteComboBox_SelectedIndexChanged(object sender, EventArgs e)
    {
        if (clienteComboBox.SelectedItem != null)
        {
            string nombre = clienteComboBox.SelectedItem.ToString();
            var cliente = clientesDict.Values.FirstOrDefault(c => c.Nombre == nombre);
            if (cliente != null)
            {
                rutComboBox.SelectedItem = cliente.RUT;
                ActualizarDatosCliente(cliente);
            }
        }
    }

    private void ActualizarDatosCliente(Cliente cliente)
    {
        direccionTextBox.Text = cliente.Direccion;
        giroTextBox.Text = cliente.Giro;
    }

    private void ActualizarInformacionProducto(DataGridViewRow row, Producto producto)
    {
        stockLabel.Text = $"Stock disponible: {producto.Kilos:N2} kg";
        stockLabel.ForeColor = producto.Kilos > 0 ? Color.DarkGreen : Color.Red;
    }

    private async void VentasDataGridView_CellEndEdit(object sender, DataGridViewCellEventArgs e)
    {
        if (e.ColumnIndex == ventasDataGridView.Columns["Codigo"].Index)
        {
            var row = ventasDataGridView.Rows[e.RowIndex];
            string codigo = row.Cells["Codigo"].Value?.ToString();

            if (!string.IsNullOrEmpty(codigo))
            {
                var producto = sqliteService.GetProductoPorCodigo(codigo);
                if (producto != null)
                {
                    row.Cells["Descripcion"].Value = producto.Nombre;
                    row.Cells["Marca"].Value = producto.Marca;  // Agregar esta línea
                    stockLabel.Text = $"Stock disponible: {producto.Kilos:N2} kg";
                    stockLabel.ForeColor = producto.Kilos > 0 ? Color.DarkGreen : Color.Red;
                }
                else
                {
                    row.Cells["Codigo"].Value = null;
                    row.Cells["Descripcion"].Value = null;
                    row.Cells["Marca"].Value = null;  // Agregar esta línea
                    stockLabel.Text = "Producto no encontrado";
                    stockLabel.ForeColor = Color.Red;
                    MessageBox.Show("Producto no encontrado", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }
    }

    private void VentasDataGridView_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
    {
        if (ventasDataGridView.CurrentCell.ColumnIndex == ventasDataGridView.Columns["Descripcion"].Index)
        {
            if (e.Control is ComboBox combo)
            {
                combo.DropDownStyle = ComboBoxStyle.DropDown;
                combo.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
                combo.AutoCompleteSource = AutoCompleteSource.ListItems;
            }
        }
    }


    private async void VentasDataGridView_CellValueChanged(object sender, DataGridViewCellEventArgs e)
    {
        if (e.RowIndex >= 0)
        {
            var row = ventasDataGridView.Rows[e.RowIndex];

            // Si cambió la descripción
            if (e.ColumnIndex == ventasDataGridView.Columns["Descripcion"].Index)
            {
                string descripcion = row.Cells["Descripcion"].Value?.ToString();
                if (!string.IsNullOrEmpty(descripcion) && productosPorNombre.ContainsKey(descripcion))
                {
                    var producto = productosPorNombre[descripcion];
                    row.Cells["Codigo"].Value = producto.Codigo;
                    row.Cells["Marca"].Value = producto.Marca;
                    if (rutComboBox.SelectedItem != null)
                    {
                        var ultimoPrecio = await sqliteService.GetUltimoPrecioVentaAsync(producto.Codigo, rutComboBox.SelectedItem.ToString());
                        if (ultimoPrecio.HasValue)
                        {
                            row.Cells["Precio"].Value = ultimoPrecio.Value;
                        }
                        MostrarPreciosHistoricos(producto.Codigo, rutComboBox.SelectedItem.ToString());
                    }
                }
            }

            // Si cambió el código
            else if (e.ColumnIndex == ventasDataGridView.Columns["Codigo"].Index)
            {
                string codigo = row.Cells["Codigo"].Value?.ToString();
                if (!string.IsNullOrEmpty(codigo) && productosPorCodigo.ContainsKey(codigo))
                {
                    var producto = productosPorCodigo[codigo];
                    row.Cells["Descripcion"].Value = producto.Nombre;
                    row.Cells["Marca"].Value = producto.Marca;
                    if (rutComboBox.SelectedItem != null)
                    {
                        var ultimoPrecio = await sqliteService.GetUltimoPrecioVentaAsync(codigo, rutComboBox.SelectedItem.ToString());
                        if (ultimoPrecio.HasValue)
                        {
                            row.Cells["Precio"].Value = ultimoPrecio.Value;
                        }
                        MostrarPreciosHistoricos(codigo, rutComboBox.SelectedItem.ToString());
                    }
                }
            }
            // Si cambió el valor de Bandejas o Kilos Bruto
            else if (e.ColumnIndex == ventasDataGridView.Columns["Bandejas"].Index ||
                    e.ColumnIndex == ventasDataGridView.Columns["KilosBruto"].Index)
            {
                // Calcular Kilos Neto
                if (double.TryParse(row.Cells["KilosBruto"].Value?.ToString(), out double kilosBruto) &&
                    double.TryParse(row.Cells["Bandejas"].Value?.ToString(), out double bandejas))
                {
                    double descuentoBandejas = bandejas * 1.5;
                    double kilosNeto = kilosBruto - descuentoBandejas;

                    // Actualizar Kilos Neto
                    row.Cells["KilosNeto"].Value = kilosNeto.ToString("N2");

                    // Si hay precio, actualizar el total
                    if (double.TryParse(row.Cells["Precio"].Value?.ToString(), out double precio))
                    {
                        double total = kilosNeto * precio;
                        row.Cells["Total"].Value = total.ToString("N2");
                    }

                    // Verificar stock disponible
                    string codigo = row.Cells["Codigo"].Value?.ToString();
                    if (!string.IsNullOrEmpty(codigo) && productosPorCodigo.ContainsKey(codigo))
                    {
                        var producto = productosPorCodigo[codigo];
                        if (kilosNeto > producto.Kilos)
                        {
                            stockLabel.ForeColor = Color.Red;
                            MessageBox.Show($"¡Advertencia! La cantidad excede el stock disponible ({producto.Kilos:N2} kg)",
                                "Stock Insuficiente", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        }
                    }
                }
            }
            // Si cambió el Precio
            else if (e.ColumnIndex == ventasDataGridView.Columns["Precio"].Index)
            {
                if (double.TryParse(row.Cells["KilosNeto"].Value?.ToString(), out double kilosNeto) &&
                    double.TryParse(row.Cells["Precio"].Value?.ToString(), out double precio))
                {
                    double total = kilosNeto * precio;
                    row.Cells["Total"].Value = total.ToString("N2");
                }
            }

            // Recalcular el total general de la venta
            CalcularTotalVenta();

            // Validar si hay datos suficientes para habilitar los botones
            bool hayDatosValidos = ventasDataGridView.Rows.Count > 1 &&
                                  ventasDataGridView.Rows.Cast<DataGridViewRow>()
                                      .Any(r => !r.IsNewRow && r.Cells["Total"].Value != null);

            exportarExcelButton.Enabled = hayDatosValidos;
            imprimirButton.Enabled = hayDatosValidos;
        }
    }

    private void ConfigurarColumnasGrid()
    {
        // Crear un DataGridViewComboBoxColumn para la descripción
        var descripcionColumn = new DataGridViewComboBoxColumn
        {
            Name = "Descripcion",
            HeaderText = "Descripción",
            DataPropertyName = "Descripcion",
            AutoComplete = true,
            FlatStyle = FlatStyle.Flat,
            DisplayStyle = DataGridViewComboBoxDisplayStyle.ComboBox
        };
        descripcionColumn.Items.AddRange(productosPorNombre.Keys.ToArray());

        var ventaLocalColumn = new DataGridViewCheckBoxColumn
        {
            Name = "VentaLocal",
            HeaderText = "Venta Local",
            Width = 80,
            FlatStyle = FlatStyle.Flat  // Corregido de FalStyle a FlatStyle
        };

        ventasDataGridView.Columns.Add(ventaLocalColumn);

        // Reemplazar la columna de descripción existente
        int descripcionIndex = ventasDataGridView.Columns["Descripcion"].Index;
        ventasDataGridView.Columns.RemoveAt(descripcionIndex);
        ventasDataGridView.Columns.Insert(descripcionIndex, descripcionColumn);

        // Formato para columnas numéricas
        ventasDataGridView.Columns["Bandejas"].DefaultCellStyle.Format = "N0";
        ventasDataGridView.Columns["KilosBruto"].DefaultCellStyle.Format = "N2";
        ventasDataGridView.Columns["KilosNeto"].DefaultCellStyle.Format = "N2";
        ventasDataGridView.Columns["Precio"].DefaultCellStyle.Format = "N0"; // Cambiado a N0 para precios sin decimales
        ventasDataGridView.Columns["Total"].DefaultCellStyle.Format = "N0"; // Cambiado a N0 para totales sin decimales

        // Alineación a la derecha para columnas numéricas
        ventasDataGridView.Columns["Bandejas"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
        ventasDataGridView.Columns["KilosBruto"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
        ventasDataGridView.Columns["KilosNeto"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
        ventasDataGridView.Columns["Precio"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
        ventasDataGridView.Columns["Total"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

        // Hacer algunas columnas de solo lectura
        ventasDataGridView.Columns["KilosNeto"].ReadOnly = true;
        ventasDataGridView.Columns["Total"].ReadOnly = true;
        ventasDataGridView.Columns["Marca"].ReadOnly = true;
        ventasDataGridView.Columns["Precio"].ReadOnly = false; // Aseguramos que el precio se pueda editar

        // Cargar todos los productos sin filtrar
        var todosLosProductos = sqliteService.GetProductos();
        descripcionColumn.Items.Clear();
        foreach (var producto in todosLosProductos)
        {
            descripcionColumn.Items.Add(producto.Nombre);
        }

        // Establecer anchos de columna
        ventasDataGridView.Columns["Codigo"].Width = 80;
        ventasDataGridView.Columns["Descripcion"].Width = 250;
        ventasDataGridView.Columns["Marca"].Width = 100;
        ventasDataGridView.Columns["Bandejas"].Width = 80;
        ventasDataGridView.Columns["KilosBruto"].Width = 100;
        ventasDataGridView.Columns["KilosNeto"].Width = 100;
        ventasDataGridView.Columns["Precio"].Width = 100;
        ventasDataGridView.Columns["Total"].Width = 100;

        // Configurar colores de fondo para mejor visualización
        ventasDataGridView.DefaultCellStyle.BackColor = Color.White;
        ventasDataGridView.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(240, 240, 240);
        ventasDataGridView.EnableHeadersVisualStyles = false;
        ventasDataGridView.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(50, 50, 50);
        ventasDataGridView.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;

        // Configurar bordes
        ventasDataGridView.BorderStyle = BorderStyle.None;
        ventasDataGridView.CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal;
        ventasDataGridView.GridColor = Color.FromArgb(200, 200, 200);
    }

    private void RecargarProductos()
    {
        CargarProductos();
        var descripcionColumn = ventasDataGridView.Columns["Descripcion"] as DataGridViewComboBoxColumn;
        if (descripcionColumn != null)
        {
            descripcionColumn.Items.Clear();
            foreach (var producto in productosPorNombre.Values)
            {
                descripcionColumn.Items.Add(producto.Nombre);
            }
        }
    }

    private void CalcularTotalVenta()
    {
        totalVenta = 0;
        foreach (DataGridViewRow row in ventasDataGridView.Rows)
        {
            if (!row.IsNewRow && row.Cells["Total"].Value != null)
            {
                if (double.TryParse(row.Cells["Total"].Value.ToString(), out double total))
                {
                    totalVenta += total;
                }
            }
        }
        totalVentaTextBox.Text = totalVenta.ToString("C");
    }

    private void VentasDataGridView_DefaultValuesNeeded(object sender, DataGridViewRowEventArgs e)
    {
        e.Row.Cells["Bandejas"].Value = 0;
    }

    private async void MostrarPreciosHistoricos(string codigo, string rut)
    {
        try
        {
            // Mostrar stock
            if (productosPorCodigo.TryGetValue(codigo, out var producto))
            {
                stockLabel.Text = $"Stock disponible: {producto.Kilos:N2} kg";
                stockLabel.ForeColor = producto.Kilos > 0 ? Color.DarkGreen : Color.Red;
            }

            // Obtener precios
            var ultimoPrecioVenta = await sqliteService.GetUltimoPrecioVentaAsync(codigo, rut);
            var ultimoPrecioCompra = await sqliteService.GetUltimoPrecioCompraAsync(codigo);

            // Mostrar precios
            if (ultimoPrecioCompra.HasValue)
            {
                preciosLabel.Text = $"Último precio de compra: ${ultimoPrecioCompra.Value:N0}";
                preciosLabel.ForeColor = Color.DarkBlue;
                preciosLabel.Visible = true;
            }
            else
            {
                preciosLabel.Visible = false;
            }
        }
        catch (Exception ex)
        {
            preciosLabel.Text = "Error al cargar precios";
            preciosLabel.ForeColor = Color.Red;
            _logger.LogError(ex, "Error al obtener precios históricos");
        }
    }

    private async void FinalizarButton_Click(object sender, EventArgs e)
    {
        if (!ValidarVenta())
            return;

        SQLiteConnection connection = null;
        SQLiteTransaction transaction = null;

        try
        {
            // Deshabilitar el botón mientras se procesa
            finalizarButton.Enabled = false;

            connection = new SQLiteConnection(sqliteService.connectionString);
            await connection.OpenAsync();
            transaction = connection.BeginTransaction();

            var numeroGuia = int.Parse(numeroGuiaTextBox.Text);
            var rut = rutComboBox.SelectedItem.ToString();
            var fechaVenta = fechaEmisionPicker.Value.ToString("yyyy-MM-dd");
            var pagadoConCredito = pagarConCreditoCheckBox.Checked ? 1 : 0;
            var totalVentaActual = ventasDataGridView.Rows.Cast<DataGridViewRow>()
    .Where(r => !r.IsNewRow)
    .Sum(r => Convert.ToDouble(r.Cells["Total"].Value));

            // Registrar cada línea de la venta
            foreach (DataGridViewRow row in ventasDataGridView.Rows)
            {
                if (row.IsNewRow || row.Cells["Codigo"].Value == null)
                    continue;

                var codigo = row.Cells["Codigo"].Value.ToString();
                var descripcion = row.Cells["Descripcion"].Value?.ToString();
                var bandejas = Convert.ToInt32(row.Cells["Bandejas"].Value);
                var kilosNeto = Convert.ToDouble(row.Cells["KilosNeto"].Value);
                var precio = Convert.ToDouble(row.Cells["Precio"].Value);
                var total = Convert.ToDouble(row.Cells["Total"].Value);
                var esVentaLocal = Convert.ToBoolean(row.Cells["VentaLocal"].Value);

                // Insertar venta
                var command = new SQLiteCommand(@"
        INSERT INTO Ventas (
            NumeroGuia, CodigoProducto, Descripcion, Bandejas, 
            KilosNeto, FechaVenta, PagadoConCredito, RUT, Total, Precio
        ) VALUES (
            @numeroGuia, @codigo, @descripcion, @bandejas, 
            @kilosNeto, @fechaVenta, @pagadoConCredito, @rut, @total, @precio
        )", connection, transaction);

                command.Parameters.AddWithValue("@numeroGuia", numeroGuia);
                command.Parameters.AddWithValue("@codigo", codigo);
                command.Parameters.AddWithValue("@descripcion", descripcion);
                command.Parameters.AddWithValue("@bandejas", bandejas);
                command.Parameters.AddWithValue("@kilosNeto", kilosNeto);
                command.Parameters.AddWithValue("@fechaVenta", fechaVenta);
                command.Parameters.AddWithValue("@pagadoConCredito", pagadoConCredito);
                command.Parameters.AddWithValue("@rut", rut);
                command.Parameters.AddWithValue("@total", total);
                command.Parameters.AddWithValue("@precio", precio);

                await command.ExecuteNonQueryAsync();

                // Actualizar inventario solo si no es venta local
                if (!esVentaLocal)
                {
                    var updateInventarioCmd = new SQLiteCommand(@"
            UPDATE Inventario 
            SET Kilos = Kilos - @kilosNeto 
            WHERE Codigo = @codigo", connection, transaction);

                    updateInventarioCmd.Parameters.AddWithValue("@kilosNeto", kilosNeto);
                    updateInventarioCmd.Parameters.AddWithValue("@codigo", codigo);
                    await updateInventarioCmd.ExecuteNonQueryAsync();
                }
            }

            // Si es venta a crédito, actualizar deuda del cliente
            if (pagadoConCredito == 1)
            {
                var updateDeudaCmd = new SQLiteCommand(@"
                UPDATE Clientes 
                SET Deuda = Deuda + @monto 
                WHERE RUT = @rut", connection, transaction);

                updateDeudaCmd.Parameters.AddWithValue("@monto", totalVentaActual);
                updateDeudaCmd.Parameters.AddWithValue("@rut", rut);
                await updateDeudaCmd.ExecuteNonQueryAsync();
            }

            // Incrementar número de guía
            var updateGuiaCmd = new SQLiteCommand(@"
            UPDATE Configuracion 
            SET Valor = CAST((CAST(Valor AS INTEGER) + 1) AS TEXT) 
            WHERE Clave = 'UltimoNumeroGuia'", connection, transaction);

            await updateGuiaCmd.ExecuteNonQueryAsync();

            await transaction.CommitAsync();
            MessageBox.Show("Venta registrada exitosamente. Haga clic en el número de guía para iniciar una nueva venta.",
        "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);

        }
        catch (Exception ex)
        {
            transaction?.Rollback();
            MessageBox.Show($"Error al registrar la venta: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        finally
        {
            finalizarButton.Enabled = true;
            transaction?.Dispose();
            connection?.Close();
            connection?.Dispose();
        }
    }

    private bool ValidarVenta()
    {
        if (rutComboBox.SelectedItem == null)
        {
            MessageBox.Show("Debe seleccionar un cliente", "Validación", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return false;
        }

        if (ventasDataGridView.Rows.Count <= 1)
        {
            MessageBox.Show("Debe agregar al menos un producto", "Validación", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return false;
        }

        foreach (DataGridViewRow row in ventasDataGridView.Rows)
        {
            if (row.IsNewRow) continue;

            if (row.Cells["Codigo"].Value == null || string.IsNullOrEmpty(row.Cells["Codigo"].Value.ToString()))
            {
                MessageBox.Show($"Falta el código en la línea {row.Index + 1}", "Validación", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }

            if (row.Cells["KilosNeto"].Value == null ||
                !double.TryParse(row.Cells["KilosNeto"].Value.ToString(), out double kilos) ||
                kilos <= 0)
            {
                MessageBox.Show($"Los kilos en la línea {row.Index + 1} deben ser un número mayor a 0",
                    "Validación", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }

            if (row.Cells["Precio"].Value == null ||
                !double.TryParse(row.Cells["Precio"].Value.ToString(), out double precio) ||
                precio <= 0)
            {
                MessageBox.Show($"El precio en la línea {row.Index + 1} debe ser un número mayor a 0",
                    "Validación", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }

            string codigoProducto = row.Cells["Codigo"].Value.ToString();
            var productoActual = sqliteService.GetProductoPorCodigo(codigoProducto);

            if (productoActual == null)
            {
                MessageBox.Show($"El producto con código {codigoProducto} no existe",
                    "Validación", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }

            bool esVentaLocal = Convert.ToBoolean(row.Cells["VentaLocal"].Value);
            if (!esVentaLocal && productoActual.Kilos < kilos)
            {
                MessageBox.Show($"No hay suficiente stock para el producto {codigoProducto}. " +
                    $"Stock disponible: {productoActual.Kilos:N2} kg",
                    "Validación", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }
        }

        return true;
    }

    private void LimpiarFormulario()
    {
        // Cargar nuevo número de guía
        CargarNumeroGuia();

        stockLabel.Text = "";

        // Limpiar selecciones
        rutComboBox.SelectedIndex = -1;
        clienteComboBox.SelectedIndex = -1;
        direccionTextBox.Clear();
        giroTextBox.Clear();

        // Reiniciar fecha
        fechaEmisionPicker.Value = DateTime.Now;

        // Limpiar grid
        ventasDataGridView.Rows.Clear();

        // Reiniciar totales
        totalVenta = 0;
        totalVentaTextBox.Text = "0";

        // Desmarcar crédito
        pagarConCreditoCheckBox.Checked = false;
    }
}
