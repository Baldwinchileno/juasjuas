using System;
using System.Windows.Forms;
using System.Drawing;
using System.Data;
using AdminSERMAC.Models;
using AdminSERMAC.Services;
using System.Data.SQLite;

namespace AdminSERMAC.Forms
{
    public partial class ReportesForm : Form
    {
        private readonly ReportGenerator _reportGenerator;
        private readonly SQLiteService _sqliteService;
        private ComboBox cmbTipoReporte;
        private ComboBox cmbPeriodo;
        private ComboBox cmbCategoria;
        private ComboBox cmbProducto;
        private Button btnGenerar;
        private DataGridView dgvResultados;
        private Button btnExportar;

        public ReportesForm(ReportGenerator reportGenerator, SQLiteService sqliteService)
        {
            _reportGenerator = reportGenerator;
            _sqliteService = sqliteService;
            InitializeComponent();
            ConfigureComponents();
            LoadProductos();
            LoadCategorias();
        }

        private void InitializeComponent()
        {
            this.SuspendLayout();

            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Name = "ReportesForm";
            this.Text = "Reportes de Compras";
            this.StartPosition = FormStartPosition.CenterScreen;

            this.ResumeLayout(false);
        }

        private void ConfigureComponents()
        {
            // Tipo de Reporte
            cmbTipoReporte = new ComboBox
            {
                Location = new Point(20, 20),
                Size = new Size(150, 25),
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            cmbTipoReporte.Items.Clear();
            cmbTipoReporte.Items.AddRange(new object[] {
        "Por Categoría",
        "Por Producto",
        "Stock Critico",
        "Productos por vencer"
    });
            cmbTipoReporte.SelectedIndex = 0;
            cmbTipoReporte.SelectedIndexChanged += CmbTipoReporte_SelectedIndexChanged;

            // Periodo
            cmbPeriodo = new ComboBox
            {
                Location = new Point(190, 20),
                Size = new Size(150, 25),
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            cmbPeriodo.Items.AddRange(new object[] { "Semanal", "Mensual" });
            cmbPeriodo.SelectedIndex = 0;

            // Categoría
            cmbCategoria = new ComboBox
            {
                Location = new Point(360, 20),
                Size = new Size(150, 25),
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            cmbCategoria.SelectedIndexChanged += CmbCategoria_SelectedIndexChanged;

            



            // Producto
            cmbProducto = new ComboBox
            {
                Location = new Point(360, 20),
                Size = new Size(150, 25),
                DropDownStyle = ComboBoxStyle.DropDownList,
                Visible = false
            };

            // Botón Generar
            btnGenerar = new Button
            {
                Text = "Generar Reporte",
                Location = new Point(530, 20),
                Size = new Size(120, 25)
            };
            btnGenerar.Click += BtnGenerar_Click;

            // DataGridView
            dgvResultados = new DataGridView
            {
                Location = new Point(20, 60),
                Size = new Size(740, 300),
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                AllowUserToAddRows = false,
                ReadOnly = true
            };

            // Botón Exportar
            btnExportar = new Button
            {
                Text = "Exportar a Excel",
                Location = new Point(20, 370),
                Size = new Size(120, 25)
            };
            btnExportar.Click += BtnExportar_Click;

            // Agregar controles al formulario
            Controls.AddRange(new Control[] {
                cmbTipoReporte,
                cmbPeriodo,
                cmbCategoria,
                cmbProducto,
                btnGenerar,
                dgvResultados,
                btnExportar
            });

            // Controles superiores
            cmbTipoReporte.Anchor = AnchorStyles.Top | AnchorStyles.Left;
            cmbPeriodo.Anchor = AnchorStyles.Top | AnchorStyles.Left;
            cmbCategoria.Anchor = AnchorStyles.Top | AnchorStyles.Left;
            cmbProducto.Anchor = AnchorStyles.Top | AnchorStyles.Left;
            btnGenerar.Anchor = AnchorStyles.Top | AnchorStyles.Left;

            // DataGridView y TabControl
            dgvResultados.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;

            // Botón Exportar
            btnExportar.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;

            // Ajustar tamaño mínimo del formulario
            this.MinimumSize = new Size(800, 500);
        }

        private async void CmbCategoria_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cmbTipoReporte.SelectedItem?.ToString() == "Stock Critico")
            {
                string categoriaSeleccionada = cmbCategoria.SelectedItem?.ToString();
                var stockCritico = await _reportGenerator.GetStockCriticoAsync(categoriaSeleccionada);
                MostrarReporteStockCritico(stockCritico);
            }
        }


        private void LoadCategorias()
        {
            try
            {
                using var connection = new SQLiteConnection(_sqliteService.connectionString);
                connection.Open();
                var command = new SQLiteCommand(
                    "SELECT DISTINCT Categoria FROM Productos WHERE Categoria IS NOT NULL ORDER BY Categoria", connection);

                // Limpiar categorías existentes
                cmbCategoria.Items.Clear();

                using var reader = command.ExecuteReader();
                while (reader.Read())
                {
                    string categoria = reader["Categoria"].ToString();
                    if (!string.IsNullOrEmpty(categoria))
                    {
                        cmbCategoria.Items.Add(categoria);
                    }
                }

                if (cmbCategoria.Items.Count > 0)
                    cmbCategoria.SelectedIndex = 0;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error al cargar categorías: {ex.Message}", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void LoadProductos()
        {
            try
            {
                using var connection = new SQLiteConnection(_sqliteService.connectionString);
                connection.Open();
                var command = new SQLiteCommand(
                    "SELECT Codigo, Nombre FROM Productos ORDER BY Nombre", connection);

                using var reader = command.ExecuteReader();
                while (reader.Read())
                {
                    var item = new ComboBoxItem
                    {
                        Value = reader["Codigo"].ToString(),
                        Text = $"{reader["Codigo"]} - {reader["Nombre"]}"
                    };
                    cmbProducto.Items.Add(item);
                }

                if (cmbProducto.Items.Count > 0)
                    cmbProducto.SelectedIndex = 0;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error al cargar productos: {ex.Message}", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void CmbTipoReporte_SelectedIndexChanged(object sender, EventArgs e)
        {
            string tipoReporte = cmbTipoReporte.SelectedItem.ToString();

            // Determinar qué controles mostrar según el tipo de reporte
            switch (tipoReporte)
            {
                case "Stock Crítico":
                case "Productos por vencer":
                case "Inventario Crítico Completo":
                    cmbPeriodo.Visible = false;
                    cmbCategoria.Visible = true;
                    cmbProducto.Visible = false;
                    break;

                case "Por Categoría":
                    cmbPeriodo.Visible = true;
                    cmbCategoria.Visible = true;
                    cmbProducto.Visible = false;
                    break;

                case "Por Producto":
                    cmbPeriodo.Visible = true;
                    cmbCategoria.Visible = false;
                    cmbProducto.Visible = true;
                    break;
            }
        }

        private async void BtnGenerar_Click(object sender, EventArgs e)
        {
            try
            {
                string tipoReporte = cmbTipoReporte.SelectedItem.ToString();
                string categoriaSeleccionada = cmbCategoria.SelectedItem?.ToString();

                switch (tipoReporte)
                {
                    case "Stock Critico":
                        var stockCritico = await _reportGenerator.GetStockCriticoAsync(categoriaSeleccionada);
                        MostrarReporteStockCritico(stockCritico);
                        break;

                    case "Productos por vencer":
                        var porVencer = await _reportGenerator.GetProductosPorVencerAsync(categoriaSeleccionada);
                        MostrarReportePorVencer(porVencer);
                        break;

                    case "Por Categoría":
                        var periodo = cmbPeriodo.SelectedItem.ToString();
                        var reporteCategoria = await _reportGenerator.GenerateComprasPorCategoriaReport(periodo, categoriaSeleccionada);
                        dgvResultados.DataSource = reporteCategoria.Sections.First().Value;
                        break;

                    case "Por Producto":
                        periodo = cmbPeriodo.SelectedItem.ToString();
                        var productoSeleccionado = (ComboBoxItem)cmbProducto.SelectedItem;
                        var reporteProducto = await _reportGenerator.GenerateComprasPorProductoReport(periodo, productoSeleccionado.Value);
                        dgvResultados.DataSource = reporteProducto.Sections.First().Value;
                        break;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error al generar el reporte: {ex.Message}",
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void MostrarReporteStockCritico(DataTable data)
        {
            // Remover los controles anteriores
            foreach (Control control in Controls.OfType<TabControl>().ToList())
            {
                Controls.Remove(control);
                control.Dispose();
            }
            Controls.Remove(dgvResultados);

            // Crear un TabControl
            var tabControl = new TabControl
            {
                Location = new Point(20, 60),
                Size = new Size(740, 300)
            };

            // Tab para Stock Crítico Normal
            var tabNormal = new TabPage("Inventario Normal");
            var dgvNormal = new DataGridView
            {
                Dock = DockStyle.Fill,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                AllowUserToAddRows = false,
                ReadOnly = true
            };

            // Filtrar datos para inventario normal
            var dtNormal = data.Clone();
            foreach (DataRow row in data.Select("Estado = 'Inventario' OR Estado = 'Sin Stock'"))
            {
                dtNormal.ImportRow(row);
            }
            dgvNormal.DataSource = dtNormal;
            ConfigurarGridStockCritico(dgvNormal);
            tabNormal.Controls.Add(dgvNormal);

            // Tab para Stock Crítico Congelados
            var tabCongelados = new TabPage("Inventario Congelados");
            var dgvCongelados = new DataGridView
            {
                Dock = DockStyle.Fill,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                AllowUserToAddRows = false,
                ReadOnly = true
            };

            // Filtrar datos para inventario congelados
            var dtCongelados = data.Clone();
            foreach (DataRow row in data.Select("Estado = 'Congelados'"))
            {
                dtCongelados.ImportRow(row);
            }
            dgvCongelados.DataSource = dtCongelados;
            ConfigurarGridStockCritico(dgvCongelados);
            tabCongelados.Controls.Add(dgvCongelados);

            // Agregar las pestañas al TabControl
            tabControl.TabPages.Add(tabNormal);
            tabControl.TabPages.Add(tabCongelados);

            // Agregar el TabControl al formulario
            Controls.Add(tabControl);

            // Ajustar posición del botón Exportar si existe
            if (btnExportar != null)
            {
                btnExportar.Location = new Point(20, tabControl.Bottom + 10);
            }
        }

        private void MostrarReportePorVencer(DataTable data)
        {
            // Remover el DataGridView anterior
            Controls.Remove(dgvResultados);

            // Crear un TabControl
            var tabControl = new TabControl
            {
                Location = new Point(20, 60),
                Size = new Size(740, 300)
            };

            // Tab para Inventario Normal
            var tabNormal = new TabPage("Inventario Normal");
            var dgvNormal = new DataGridView
            {
                Dock = DockStyle.Fill,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                AllowUserToAddRows = false,
                ReadOnly = true
            };

            // Filtrar datos para inventario normal
            var dtNormal = data.Clone();
            foreach (DataRow row in data.Select("Ubicacion = 'Inventario'"))
            {
                dtNormal.ImportRow(row);
            }
            dgvNormal.DataSource = dtNormal;
            ConfigurarGridProductosVencer(dgvNormal);
            tabNormal.Controls.Add(dgvNormal);

            // Tab para Inventario Congelados
            var tabCongelados = new TabPage("Inventario Congelados");
            var dgvCongelados = new DataGridView
            {
                Dock = DockStyle.Fill,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                AllowUserToAddRows = false,
                ReadOnly = true
            };

            // Filtrar datos para inventario congelados
            var dtCongelados = data.Clone();
            foreach (DataRow row in data.Select("Ubicacion = 'Congelados'"))
            {
                dtCongelados.ImportRow(row);
            }
            dgvCongelados.DataSource = dtCongelados;
            ConfigurarGridProductosVencer(dgvCongelados);
            tabCongelados.Controls.Add(dgvCongelados);

            // Agregar las pestañas al TabControl
            tabControl.TabPages.Add(tabNormal);
            tabControl.TabPages.Add(tabCongelados);

            // Agregar el TabControl al formulario
            Controls.Add(tabControl);
        }

        private void ConfigurarGridStockCritico(DataGridView dgv)
        {
            // Configurar nombres de columnas
            if (dgv.Columns.Count > 0)
            {
                dgv.Columns["Codigo"].HeaderText = "Código";
                dgv.Columns["Producto"].HeaderText = "Producto";
                dgv.Columns["Unidades"].HeaderText = "Unidades";
                dgv.Columns["Kilos"].HeaderText = "Kilos";
                dgv.Columns["Categoria"].HeaderText = "Categoría";
                dgv.Columns["SubCategoria"].HeaderText = "Sub Categoría";
                dgv.Columns["Estado"].HeaderText = "Estado";

                if (dgv.Columns.Contains("FechaVencimiento"))
                    dgv.Columns["FechaVencimiento"].HeaderText = "Fecha Vencimiento";
            }

            // Formato numérico para kilos
            if (dgv.Columns.Contains("Kilos"))
            {
                dgv.Columns["Kilos"].DefaultCellStyle.Format = "N2";
            }

            // Colorear filas según el stock
            dgv.CellFormatting += (sender, e) =>
            {
                if (e.RowIndex >= 0)
                {
                    DataGridViewRow row = dgv.Rows[e.RowIndex];
                    if (row.Cells["Kilos"].Value != null && decimal.TryParse(row.Cells["Kilos"].Value.ToString(), out decimal kilos))
                    {
                        if (kilos == 0)
                        {
                            row.DefaultCellStyle.BackColor = Color.Red;
                            row.DefaultCellStyle.ForeColor = Color.White;
                        }
                        else if (kilos <= 200)
                        {
                            row.DefaultCellStyle.BackColor = Color.LightSalmon;
                            row.DefaultCellStyle.ForeColor = Color.Black;
                        }
                    }
                }
            };
        }

        private void ConfigurarGridProductosVencer(DataGridView dgv)
        {
            // Configurar nombres de columnas
            if (dgv.Columns.Count > 0)
            {
                dgv.Columns["Codigo"].HeaderText = "Código";
                dgv.Columns["Producto"].HeaderText = "Producto";
                dgv.Columns["Unidades"].HeaderText = "Unidades";
                dgv.Columns["Kilos"].HeaderText = "Kilos";
                dgv.Columns["FechaVencimiento"].HeaderText = "Fecha Vencimiento";
                dgv.Columns["Categoria"].HeaderText = "Categoría";
                dgv.Columns["SubCategoria"].HeaderText = "Sub Categoría";
                dgv.Columns["Ubicacion"].HeaderText = "Ubicación";
            }

            // Formato para fechas y números
            if (dgv.Columns.Contains("FechaVencimiento"))
            {
                dgv.Columns["FechaVencimiento"].DefaultCellStyle.Format = "dd/MM/yyyy";
            }
            if (dgv.Columns.Contains("Kilos"))
            {
                dgv.Columns["Kilos"].DefaultCellStyle.Format = "N2";
            }
        }

        private void MostrarReporteCompleto(Report reporte)
        {
            // Remover el DataGridView anterior
            Controls.Remove(dgvResultados);

            // Crear un TabControl dinámicamente
            var tabControl = new TabControl
            {
                Location = new Point(20, 60),
                Size = new Size(740, 300)
            };

            // Tab para Stock Crítico
            var tabStockCritico = new TabPage("Stock Crítico");
            var dgvStockCritico = new DataGridView
            {
                Dock = DockStyle.Fill,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                AllowUserToAddRows = false,
                ReadOnly = true
            };
            dgvStockCritico.DataSource = reporte.Sections["Productos con Stock Crítico (Menos de 200kg)"];
            ConfigurarGridStockCritico(dgvStockCritico);  // Quitamos el cast
            tabStockCritico.Controls.Add(dgvStockCritico);

            // Tab para Productos por Vencer
            var tabPorVencer = new TabPage("Por Vencer");
            var dgvPorVencer = new DataGridView
            {
                Dock = DockStyle.Fill,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                AllowUserToAddRows = false,
                ReadOnly = true
            };
            dgvPorVencer.DataSource = reporte.Sections["Productos Por Vencer (Próximos 2 días)"];
            ConfigurarGridProductosVencer(dgvPorVencer);  // Quitamos el cast
            tabPorVencer.Controls.Add(dgvPorVencer);

            // Agregar las pestañas al TabControl
            tabControl.TabPages.Add(tabStockCritico);
            tabControl.TabPages.Add(tabPorVencer);

            // Agregar el TabControl al formulario
            Controls.Add(tabControl);
        }





        private async void BtnExportar_Click(object sender, EventArgs e)
        {
            try
            {
                using (SaveFileDialog saveDialog = new SaveFileDialog())
                {
                    saveDialog.Filter = "Excel Files|*.xlsx";
                    saveDialog.Title = "Guardar Reporte";
                    saveDialog.FileName = $"Reporte_Compras_{DateTime.Now:yyyyMMdd}";

                    if (saveDialog.ShowDialog() == DialogResult.OK)
                    {
                        var periodo = cmbPeriodo.SelectedItem.ToString();
                        var esPorCategoria = cmbTipoReporte.SelectedItem.ToString() == "Por Categoría";

                        Report reporte;
                        if (esPorCategoria)
                        {
                            var categoria = cmbCategoria.SelectedItem.ToString();
                            reporte = await _reportGenerator.GenerateComprasPorCategoriaReport(periodo, categoria);
                        }
                        else
                        {
                            var productoSeleccionado = (ComboBoxItem)cmbProducto.SelectedItem;
                            reporte = await _reportGenerator.GenerateComprasPorProductoReport(periodo, productoSeleccionado.Value);
                        }

                        await reporte.ExportToExcel(saveDialog.FileName);
                        MessageBox.Show("Reporte exportado exitosamente", "Éxito",
                            MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error al exportar el reporte: {ex.Message}", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }

    public class ComboBoxItem
    {
        public string Value { get; set; }
        public string Text { get; set; }

        public override string ToString()
        {
            return Text;
        }
    }
}