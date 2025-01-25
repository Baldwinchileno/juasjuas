using System;
using System.Data;
using System.Data.SQLite;
using System.Drawing;
using System.Windows.Forms;
using AdminSERMAC.Models;
using AdminSERMAC.Services;
using ClosedXML.Excel; // Asegúrate de tener la referencia de ClosedXML en tu proyecto.
using Microsoft.Extensions.Logging;

namespace AdminSERMAC.Forms
{
    public class VisualizarInventarioForm : Form
    {
        private TextBox codigoProductoTextBox;
        private ComboBox categoriaComboBox;
        private ComboBox subcategoriaComboBox;
        private Button buscarButton;
        private Button limpiarFiltroButton;
        private Button exportarExcelButton;
        public DataGridView inventarioDataGridView;
        private readonly SQLiteService sqliteService;
        private readonly ILogger<VisualizarInventarioForm> _logger;
        private readonly bool _esCongelados;
        public Inventario ProductoSeleccionado { get; private set; }

        public VisualizarInventarioForm(SQLiteService sqliteService, ILogger<VisualizarInventarioForm> logger, bool esCongelados = false)
        {
            this.sqliteService = sqliteService ?? throw new ArgumentNullException(nameof(sqliteService));
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
            _esCongelados = esCongelados;

            this.Text = esCongelados ? "Visualizar Cámara de Congelados" : "Visualizar Inventario";
            this.Width = 1200;
            this.Height = 700;

            InitializeComponents();
            CargarCategorias();
            CargarInventario();
        }

        private void InitializeComponents()
        {
            // Panel de búsqueda
            Panel searchPanel = new Panel
            {
                Dock = DockStyle.Top,
                Height = 100,
                Padding = new Padding(10)
            };

            // Controles de búsqueda
            Label codigoProductoLabel = new Label
            {
                Text = "Código:",
                AutoSize = true,
                Location = new Point(20, 25)
            };

            codigoProductoTextBox = new TextBox
            {
                Location = new Point(80, 22),
                Width = 120
            };

            Label categoriaLabel = new Label
            {
                Text = "Categoría:",
                AutoSize = true,
                Location = new Point(220, 25)
            };

            categoriaComboBox = new ComboBox
            {
                Location = new Point(290, 22),
                Width = 150,
                DropDownStyle = ComboBoxStyle.DropDownList
            };

            Label subcategoriaLabel = new Label
            {
                Text = "Subcategoría:",
                AutoSize = true,
                Location = new Point(460, 25)
            };

            subcategoriaComboBox = new ComboBox
            {
                Location = new Point(550, 22),
                Width = 150,
                DropDownStyle = ComboBoxStyle.DropDownList
            };

            buscarButton = new Button
            {
                Text = "Buscar",
                Location = new Point(720, 20),
                Width = 100,
                BackColor = Color.FromArgb(0, 122, 204),
                ForeColor = Color.White
            };

            limpiarFiltroButton = new Button
            {
                Text = "Limpiar Filtros",
                Location = new Point(830, 20),
                Width = 100
            };

            exportarExcelButton = new Button
            {
                Text = "Exportar a Excel",
                Location = new Point(940, 20),
                Width = 120,
                BackColor = Color.ForestGreen,
                ForeColor = Color.White
            };

            // Agregar controles al panel de búsqueda
            searchPanel.Controls.AddRange(new Control[]
            {
                codigoProductoLabel,
                codigoProductoTextBox,
                categoriaLabel,
                categoriaComboBox,
                subcategoriaLabel,
                subcategoriaComboBox,
                buscarButton,
                limpiarFiltroButton,
                exportarExcelButton
            });

            // DataGridView
            inventarioDataGridView = new DataGridView
            {
                Dock = DockStyle.Fill,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                ReadOnly = true,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                RowHeadersVisible = false,
                BackgroundColor = Color.White,
                BorderStyle = BorderStyle.None,
                AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells,
                MultiSelect = false
            };

            

            // Agregar controles al formulario
            this.Controls.Add(inventarioDataGridView);
            this.Controls.Add(searchPanel);

            // Eventos
            buscarButton.Click += BuscarButton_Click;
            limpiarFiltroButton.Click += LimpiarFiltroButton_Click;
            exportarExcelButton.Click += ExportarExcelButton_Click;
            categoriaComboBox.SelectedIndexChanged += CategoriaComboBox_SelectedIndexChanged;
        }

        private void inventarioDataGridView_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                var row = inventarioDataGridView.Rows[e.RowIndex];
                ProductoSeleccionado = new Inventario
                {
                    Codigo = row.Cells["Codigo"].Value?.ToString(),
                    Producto = row.Cells["Descripcion"].Value?.ToString(),
                    Unidades = Convert.ToInt32(row.Cells["Unidades"].Value),
                    Kilos = Convert.ToDouble(row.Cells["Kilos"].Value),
                    FechaMasAntigua = row.Cells["FechaMasAntigua"].Value?.ToString(),
                    FechaMasNueva = row.Cells["FechaMasNueva"].Value?.ToString()
                };
                this.DialogResult = DialogResult.OK;
                this.Close();
            }
        }

        private void CargarInventario(string codigoProducto = null, string categoria = null, string subcategoria = null)
        {
            try
            {
                using (var connection = new SQLiteConnection(sqliteService.GetConnection().ConnectionString))
                {
                    connection.Open();

                    string tableName = _esCongelados ? "InventarioCongelados" : "Inventario";

                    string query = $@"
            SELECT 
                p.Codigo AS 'Código',
                p.Nombre AS 'Producto',
                p.Marca AS 'Marca',
                COALESCE(i.Unidades, 0) AS 'Unidades',
                COALESCE(i.Kilos, 0) AS 'Kilos',
                COALESCE(i.FechaCompra, '') AS 'Fecha de Compra',
                COALESCE(i.FechaRegistro, '') AS 'Fecha de Registro',
                COALESCE(i.FechaVencimiento, '') AS 'Fecha de Vencimiento',
                p.Categoria AS 'Categoría',
                p.SubCategoria AS 'SubCategoría'
            FROM Productos p
            LEFT JOIN {tableName} i ON p.Codigo = i.Codigo
            WHERE 1=1";

                    if (!string.IsNullOrEmpty(codigoProducto))
                    {
                        query += " AND p.Codigo = @codigoProducto";
                    }

                    if (!string.IsNullOrEmpty(categoria) && categoria != "Todas")
                    {
                        query += " AND p.Categoria = @categoria";
                    }

                    if (!string.IsNullOrEmpty(subcategoria) && subcategoria != "Todas")
                    {
                        query += " AND p.SubCategoria = @subcategoria";
                    }

                    query += " ORDER BY p.Codigo";

                    var command = new SQLiteCommand(query, connection);

                    if (!string.IsNullOrEmpty(codigoProducto))
                        command.Parameters.AddWithValue("@codigoProducto", codigoProducto);

                    if (!string.IsNullOrEmpty(categoria) && categoria != "Todas")
                        command.Parameters.AddWithValue("@categoria", categoria);

                    if (!string.IsNullOrEmpty(subcategoria) && subcategoria != "Todas")
                        command.Parameters.AddWithValue("@subcategoria", subcategoria);

                    var adapter = new SQLiteDataAdapter(command);
                    var dt = new DataTable();
                    adapter.Fill(dt);

                    inventarioDataGridView.DataSource = dt;
                    FormatearGrid();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error al cargar el inventario: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


        private void FormatearGrid()
        {
            if (inventarioDataGridView.Columns["Unidades"] != null)
                inventarioDataGridView.Columns["Unidades"].DefaultCellStyle.Format = "N0";

            if (inventarioDataGridView.Columns["Kilos"] != null)
                inventarioDataGridView.Columns["Kilos"].DefaultCellStyle.Format = "N2";

            if (inventarioDataGridView.Columns["Fecha de Compra"] != null)
                inventarioDataGridView.Columns["Fecha de Compra"].DefaultCellStyle.Format = "dd/MM/yyyy";

            if (inventarioDataGridView.Columns["Fecha de Registro"] != null)
                inventarioDataGridView.Columns["Fecha de Registro"].DefaultCellStyle.Format = "dd/MM/yyyy";

            if (inventarioDataGridView.Columns["Fecha de Vencimiento"] != null)
                inventarioDataGridView.Columns["Fecha de Vencimiento"].DefaultCellStyle.Format = "dd/MM/yyyy";
        }

        private void CargarCategorias()
        {
            categoriaComboBox.Items.Clear();
            categoriaComboBox.Items.Add("Todas");
            categoriaComboBox.Items.AddRange(sqliteService.GetCategorias().ToArray());
            categoriaComboBox.SelectedIndex = 0;
        }

        private void CategoriaComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            subcategoriaComboBox.Items.Clear();
            subcategoriaComboBox.Items.Add("Todas");

            if (categoriaComboBox.SelectedItem?.ToString() != "Todas")
            {
                subcategoriaComboBox.Items.AddRange(sqliteService.GetSubCategorias(categoriaComboBox.SelectedItem.ToString()).ToArray());
            }

            subcategoriaComboBox.SelectedIndex = 0;
        }

        private void BuscarButton_Click(object sender, EventArgs e)
        {
            string codigo = codigoProductoTextBox.Text.Trim();
            string categoria = categoriaComboBox.SelectedItem?.ToString();
            string subcategoria = subcategoriaComboBox.SelectedItem?.ToString();

            if (categoria == "Todas") categoria = null;
            if (subcategoria == "Todas") subcategoria = null;

            CargarInventario(codigo, categoria, subcategoria);
        }

        private void LimpiarFiltroButton_Click(object sender, EventArgs e)
        {
            codigoProductoTextBox.Clear();
            categoriaComboBox.SelectedIndex = 0;
            subcategoriaComboBox.SelectedIndex = 0;
            CargarInventario();
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

                    if (saveDialog.ShowDialog() == DialogResult.OK)
                    {
                        using (var workbook = new XLWorkbook())
                        {
                            DataTable dt = ((DataTable)inventarioDataGridView.DataSource).Copy();
                            workbook.Worksheets.Add(dt, "Inventario");
                            workbook.SaveAs(saveDialog.FileName);
                        }

                        MessageBox.Show("Archivo exportado exitosamente.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error al exportar a Excel: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

    }

}
