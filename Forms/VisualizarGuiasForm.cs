using System;
using System.Data;
using System.Windows.Forms;
using System.Drawing;
using AdminSERMAC.Services;
using AdminSERMAC.Constants;
using ClosedXML.Excel;
using System.Data.SQLite;

namespace AdminSERMAC.Forms
{
    public class VisualizarGuiasForm : Form
    {
        private TextBox numeroGuiaTextBox;
        private TextBox rutClienteTextBox;
        private DateTimePicker fechaDesdeTimePicker;
        private DateTimePicker fechaHastaTimePicker;
        private Button buscarButton;
        private Button limpiarFiltrosButton;
        private Button exportarExcelButton;
        private DataGridView guiasDataGridView;
        private DataGridView detalleGuiaDataGridView;
        private SQLiteService sqliteService;
        private readonly ILogger<SQLiteService> _logger;
        private Panel infoPanel;
        private Label clienteInfoLabel;
        private Label guiaInfoLabel;
        private Label fechaInfoLabel;
        private Label estadoInfoLabel;


        public VisualizarGuiasForm(ILogger<SQLiteService> logger)
        {
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
            sqliteService = new SQLiteService(_logger);

            this.Text = "Visualizar Guías";
            this.Width = 1200;
            this.Height = 800;

            InitializeComponents();
            ConfigurarDataGridViews();

            // Establecer fechas iniciales
            fechaDesdeTimePicker.Value = DateTime.Today.AddMonths(-1);
            fechaHastaTimePicker.Value = DateTime.Today;

            // Configurar AutoComplete para RUT
            rutClienteTextBox.AutoCompleteMode = AutoCompleteMode.Suggest;
            rutClienteTextBox.AutoCompleteSource = AutoCompleteSource.CustomSource;

            CargarRUTsAutoComplete();
        }


        private void CargarRUTsAutoComplete()
        {
            try
            {
                using (var connection = new SQLiteConnection(sqliteService.connectionString))
                {
                    connection.Open();
                    var command = new SQLiteCommand(
                        "SELECT DISTINCT RUT || ' - ' || Nombre as RutNombre FROM Clientes",
                        connection);

                    var autoComplete = new AutoCompleteStringCollection();
                    using (var reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            autoComplete.Add(reader["RutNombre"].ToString());
                        }
                    }

                    rutClienteTextBox.AutoCompleteCustomSource = autoComplete;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error cargando RUTs para autocompletado: {ex.Message}",
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void InitializeComponents()
        {
            // Configuración del formulario
            this.Text = "Visualizar Guías";
            this.Width = 1200;
            this.Height = 800;

            // Panel de filtros
            var filtrosPanel = new Panel
            {
                Dock = DockStyle.Top,
                Height = 80,  // Reducido de 100 a 80
                Padding = new Padding(10)
            };

            // Controles de búsqueda
            var numeroGuiaLabel = new Label
            {
                Text = "N° Guía:",
                AutoSize = true,
                Location = new Point(20, 25)
            };

            numeroGuiaTextBox = new TextBox
            {
                Location = new Point(80, 22),
                Width = 100
            };

            var rutClienteLabel = new Label
            {
                Text = "RUT Cliente:",
                AutoSize = true,
                Location = new Point(200, 25)
            };

            rutClienteTextBox = new TextBox
            {
                Location = new Point(280, 22),
                Width = 120
            };

            var fechaDesdeLabel = new Label
            {
                Text = "Desde:",
                AutoSize = true,
                Location = new Point(420, 25)
            };

            fechaDesdeTimePicker = new DateTimePicker
            {
                Location = new Point(470, 22),
                Width = 120,
                Format = DateTimePickerFormat.Short
            };

            var fechaHastaLabel = new Label
            {
                Text = "Hasta:",
                AutoSize = true,
                Location = new Point(600, 25)
            };

            fechaHastaTimePicker = new DateTimePicker
            {
                Location = new Point(650, 22),
                Width = 120,
                Format = DateTimePickerFormat.Short
            };

            buscarButton = new Button
            {
                Text = "Buscar",
                Location = new Point(790, 20),
                Width = 100,
                BackColor = Color.FromArgb(0, 122, 204),
                ForeColor = Color.White
            };

            limpiarFiltrosButton = new Button
            {
                Text = "Limpiar Filtros",
                Location = new Point(900, 20),
                Width = 100
            };

            exportarExcelButton = new Button
            {
                Text = "Exportar a Excel",
                Location = new Point(1010, 20),
                Width = 120,
                BackColor = Color.ForestGreen,
                ForeColor = Color.White
            };

            // Agregar controles al panel de filtros
            filtrosPanel.Controls.AddRange(new Control[] {
        numeroGuiaLabel, numeroGuiaTextBox,
        rutClienteLabel, rutClienteTextBox,
        fechaDesdeLabel, fechaDesdeTimePicker,
        fechaHastaLabel, fechaHastaTimePicker,
        buscarButton, limpiarFiltrosButton, exportarExcelButton
    });

            // Panel de información
            infoPanel = new Panel
            {
                Dock = DockStyle.Top,
                Height = 60,
                BackColor = Color.FromArgb(240, 240, 240),
                Visible = false,
                Padding = new Padding(5)
            };

            // Crear etiquetas de información
            guiaInfoLabel = new Label
            {
                Location = new Point(20, 5),
                AutoSize = true,
                Font = new Font("Segoe UI", 10, FontStyle.Bold)
            };

            clienteInfoLabel = new Label
            {
                Location = new Point(20, 30),
                AutoSize = true,
                Font = new Font("Segoe UI", 10)
            };

            fechaInfoLabel = new Label
            {
                Location = new Point(400, 5),
                AutoSize = true,
                Font = new Font("Segoe UI", 10)
            };

            estadoInfoLabel = new Label
            {
                Location = new Point(400, 30),
                AutoSize = true,
                Font = new Font("Segoe UI", 10)
            };

            infoPanel.Controls.AddRange(new Control[] {
        guiaInfoLabel,
        clienteInfoLabel,
        fechaInfoLabel,
        estadoInfoLabel
    });

            // TableLayoutPanel principal para organizar los elementos
            var mainTableLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 4,
                Padding = new Padding(10),
            };

            mainTableLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 80F));   // Filtros
            mainTableLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 60F));   // Info Panel
            mainTableLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 200F));  // Grid Superior
            mainTableLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));   // Grid Inferior

            // Configurar las grids
            guiasDataGridView = new DataGridView
            {
                Dock = DockStyle.Fill,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                AllowUserToAddRows = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                ReadOnly = true,
                BackgroundColor = Color.White,
                BorderStyle = BorderStyle.None,
                RowHeadersVisible = false,
                AllowUserToResizeRows = false
            };

            detalleGuiaDataGridView = new DataGridView
            {
                Dock = DockStyle.Fill,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                AllowUserToAddRows = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                ReadOnly = true,
                BackgroundColor = Color.White,
                BorderStyle = BorderStyle.None,
                RowHeadersVisible = false,
                AllowUserToResizeRows = false
            };

            // Configurar columnas de las grids
            ConfigurarColumnas();

            // Configurar estilos de las grids
            ConfigurarDataGridViews();

            // Agregar controles al TableLayoutPanel
            mainTableLayout.Controls.Add(filtrosPanel, 0, 0);
            mainTableLayout.Controls.Add(infoPanel, 0, 1);
            mainTableLayout.Controls.Add(guiasDataGridView, 0, 2);
            mainTableLayout.Controls.Add(detalleGuiaDataGridView, 0, 3);

            // Agregar el TableLayoutPanel al formulario
            this.Controls.Add(mainTableLayout);

            // Configurar eventos
            buscarButton.Click += BuscarButton_Click;
            limpiarFiltrosButton.Click += LimpiarFiltrosButton_Click;
            exportarExcelButton.Click += ExportarExcelButton_Click;
            guiasDataGridView.SelectionChanged += GuiasDataGridView_SelectionChanged;

            // Establecer fechas iniciales
            fechaDesdeTimePicker.Value = DateTime.Today.AddMonths(-1);
            fechaHastaTimePicker.Value = DateTime.Today;

            // Configurar AutoComplete para RUT
            rutClienteTextBox.AutoCompleteMode = AutoCompleteMode.Suggest;
            rutClienteTextBox.AutoCompleteSource = AutoCompleteSource.CustomSource;

            CargarRUTsAutoComplete();
        }

        private void ConfigurarColumnas()
        {
            // Columnas para la grid de guías
            guiasDataGridView.Columns.Add("NumeroGuia", "N° Guía");
            guiasDataGridView.Columns.Add("FechaVenta", "Fecha");
            guiasDataGridView.Columns.Add("RUT", "RUT");
            guiasDataGridView.Columns.Add("ClienteNombre", "Cliente");
            guiasDataGridView.Columns.Add("Total", "Total");
            guiasDataGridView.Columns.Add("Estado", "Estado");

            // Columnas para la grid de detalle
            detalleGuiaDataGridView.Columns.Add("CodigoProducto", "Código");
            detalleGuiaDataGridView.Columns.Add("Descripcion", "Descripción");
            detalleGuiaDataGridView.Columns.Add("Bandejas", "Bandejas");
            detalleGuiaDataGridView.Columns.Add("KilosNeto", "Kilos");
            detalleGuiaDataGridView.Columns.Add("Total", "Total");
        }

        private void BuscarButton_Click(object sender, EventArgs e)
        {
            try
            {
                using (var connection = new SQLiteConnection(sqliteService.connectionString))
                {
                    connection.Open();
                    var query = @"
                        SELECT 
                            v.NumeroGuia,
                            MIN(v.FechaVenta) as FechaVenta,
                            v.RUT,
                            c.Nombre as ClienteNombre,
                            SUM(v.Total) as Total,
                            v.PagadoConCredito
                        FROM Ventas v
                        JOIN Clientes c ON v.RUT = c.RUT
                        WHERE 1=1";

                    if (!string.IsNullOrEmpty(numeroGuiaTextBox.Text))
                        query += " AND v.NumeroGuia = @numeroGuia";

                    if (!string.IsNullOrEmpty(rutClienteTextBox.Text))
                        query += " AND v.RUT LIKE @rut";

                    query += @" AND date(v.FechaVenta) BETWEEN @fechaDesde AND @fechaHasta
                              GROUP BY v.NumeroGuia, v.RUT, c.Nombre, v.PagadoConCredito
                              ORDER BY v.NumeroGuia DESC";

                    using (var command = new SQLiteCommand(query, connection))
                    {
                        if (!string.IsNullOrEmpty(numeroGuiaTextBox.Text))
                            command.Parameters.AddWithValue("@numeroGuia", numeroGuiaTextBox.Text);

                        if (!string.IsNullOrEmpty(rutClienteTextBox.Text))
                            command.Parameters.AddWithValue("@rut", $"%{rutClienteTextBox.Text}%");

                        command.Parameters.AddWithValue("@fechaDesde", fechaDesdeTimePicker.Value.ToString("yyyy-MM-dd"));
                        command.Parameters.AddWithValue("@fechaHasta", fechaHastaTimePicker.Value.ToString("yyyy-MM-dd"));

                        using (var reader = command.ExecuteReader())
                        {
                            guiasDataGridView.Rows.Clear();
                            while (reader.Read())
                            {
                                guiasDataGridView.Rows.Add(
                                    reader["NumeroGuia"],
                                    Convert.ToDateTime(reader["FechaVenta"]).ToString("dd/MM/yyyy"),
                                    reader["RUT"],
                                    reader["ClienteNombre"],
                                    $"${Convert.ToDouble(reader["Total"]):N0}",
                                    Convert.ToInt32(reader["PagadoConCredito"]) == 1 ? "Crédito" : "Contado"
                                );
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error al buscar guías: {ex.Message}",
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void GuiasDataGridView_SelectionChanged(object sender, EventArgs e)
        {
            if (guiasDataGridView.SelectedRows.Count > 0)
            {
                var row = guiasDataGridView.SelectedRows[0];
                var numeroGuia = row.Cells["NumeroGuia"].Value.ToString();
                var fecha = row.Cells["FechaVenta"].Value.ToString();
                var cliente = row.Cells["ClienteNombre"].Value.ToString();
                var rut = row.Cells["RUT"].Value.ToString();
                var estado = row.Cells["Estado"].Value.ToString();

                // Actualizar panel de información
                guiaInfoLabel.Text = $"Guía N°: {numeroGuia}";
                clienteInfoLabel.Text = $"Cliente: {cliente} (RUT: {rut})";
                fechaInfoLabel.Text = $"Fecha: {fecha}";
                estadoInfoLabel.Text = $"Estado: {estado}";
                infoPanel.Visible = true;

                // Cargar detalles
                CargarDetalleGuia(numeroGuia);
            }
            else
            {
                infoPanel.Visible = false;
            }
        }

        private void CargarDetalleGuia(string numeroGuia)
        {
            try
            {
                using (var connection = new SQLiteConnection(sqliteService.connectionString))
                {
                    connection.Open();
                    var command = new SQLiteCommand(QueryConstants.Ventas.SELECT_DETALLES_GUIA, connection);
                    command.Parameters.AddWithValue("@NumeroGuia", numeroGuia);

                    detalleGuiaDataGridView.Rows.Clear();
                    using (var reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            detalleGuiaDataGridView.Rows.Add(
                                reader["CodigoProducto"],
                                reader["Descripcion"],
                                reader["Bandejas"],
                                $"{Convert.ToDouble(reader["KilosNeto"]):N2}",
                                $"${Convert.ToDouble(reader["Total"]):N0}"
                            );
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error al cargar detalle de guía: {ex.Message}",
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


        private void LimpiarFiltrosButton_Click(object sender, EventArgs e)
        {
            numeroGuiaTextBox.Clear();
            rutClienteTextBox.Clear();
            fechaDesdeTimePicker.Value = DateTime.Today.AddMonths(-1);
            fechaHastaTimePicker.Value = DateTime.Today;
            guiasDataGridView.Rows.Clear();
            detalleGuiaDataGridView.Rows.Clear();
            infoPanel.Visible = false;
        }

        private void ExportarExcelButton_Click(object sender, EventArgs e)
        {
            try
            {
                using (SaveFileDialog saveDialog = new SaveFileDialog())
                {
                    saveDialog.Filter = "Excel files (*.xlsx)|*.xlsx";
                    saveDialog.FilterIndex = 1;
                    saveDialog.RestoreDirectory = true;
                    saveDialog.FileName = $"ReporteGuias_{DateTime.Now:yyyyMMdd_HHmmss}";

                    if (saveDialog.ShowDialog() == DialogResult.OK)
                    {
                        using (var workbook = new XLWorkbook())
                        {
                            // Exportar guías
                            var worksheetGuias = workbook.Worksheets.Add("Guías");

                            // Encabezados de guías
                            for (int i = 0; i < guiasDataGridView.Columns.Count; i++)
                            {
                                worksheetGuias.Cell(1, i + 1).Value = guiasDataGridView.Columns[i].HeaderText;
                                worksheetGuias.Cell(1, i + 1).Style.Font.Bold = true;
                            }

                            // Datos de guías
                            for (int i = 0; i < guiasDataGridView.Rows.Count; i++)
                            {
                                for (int j = 0; j < guiasDataGridView.Columns.Count; j++)
                                {
                                    worksheetGuias.Cell(i + 2, j + 1).Value =
                                        guiasDataGridView.Rows[i].Cells[j].Value?.ToString() ?? "";
                                }
                            }

                            // Exportar detalles
                            var worksheetDetalles = workbook.Worksheets.Add("Detalles");

                            // Encabezados de detalles
                            for (int i = 0; i < detalleGuiaDataGridView.Columns.Count; i++)
                            {
                                worksheetDetalles.Cell(1, i + 1).Value = detalleGuiaDataGridView.Columns[i].HeaderText;
                                worksheetDetalles.Cell(1, i + 1).Style.Font.Bold = true;
                            }

                            // Datos de detalles
                            for (int i = 0; i < detalleGuiaDataGridView.Rows.Count; i++)
                            {
                                for (int j = 0; j < detalleGuiaDataGridView.Columns.Count; j++)
                                {
                                    worksheetDetalles.Cell(i + 2, j + 1).Value =
                                        detalleGuiaDataGridView.Rows[i].Cells[j].Value?.ToString() ?? "";
                                }
                            }

                            // Ajustar columnas
                            worksheetGuias.Columns().AdjustToContents();
                            worksheetDetalles.Columns().AdjustToContents();

                            // Guardar el archivo
                            workbook.SaveAs(saveDialog.FileName);
                        }

                        MessageBox.Show("Archivo exportado exitosamente", "Éxito",
                            MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error al exportar a Excel: {ex.Message}",
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // Agregar este método para formatear las columnas numéricas
        private void FormatearColumnasNumericas()
        {
            // Grid de guías
            guiasDataGridView.Columns["Total"].DefaultCellStyle.Format = "C0"; // Formato moneda
            guiasDataGridView.Columns["FechaVenta"].DefaultCellStyle.Format = "dd/MM/yyyy";

            // Grid de detalles
            detalleGuiaDataGridView.Columns["Bandejas"].DefaultCellStyle.Format = "N0";
            detalleGuiaDataGridView.Columns["KilosNeto"].DefaultCellStyle.Format = "N2";
            detalleGuiaDataGridView.Columns["Total"].DefaultCellStyle.Format = "C0";
        }

        private void ConfigurarDataGridViews()
        {
            // Configurar grid de guías
            guiasDataGridView.AllowUserToOrderColumns = true;
            guiasDataGridView.AllowUserToResizeColumns = true;
            guiasDataGridView.RowHeadersVisible = false;
            guiasDataGridView.DefaultCellStyle.SelectionBackColor = Color.LightSteelBlue;
            guiasDataGridView.DefaultCellStyle.SelectionForeColor = Color.Black;
            guiasDataGridView.BackgroundColor = Color.White;

            // Configurar grid de detalles
            detalleGuiaDataGridView.AllowUserToOrderColumns = true;
            detalleGuiaDataGridView.AllowUserToResizeColumns = true;
            detalleGuiaDataGridView.RowHeadersVisible = false;
            detalleGuiaDataGridView.DefaultCellStyle.SelectionBackColor = Color.LightSteelBlue;
            detalleGuiaDataGridView.DefaultCellStyle.SelectionForeColor = Color.Black;
            detalleGuiaDataGridView.BackgroundColor = Color.White;

            // Llamar al método de formateo
            FormatearColumnasNumericas();
        }

    }
}