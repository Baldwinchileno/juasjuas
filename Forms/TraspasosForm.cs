using System;
using System.Data.SQLite;
using System.Drawing;
using System.Windows.Forms;
using AdminSERMAC.Models;
using AdminSERMAC.Services;
using AdminSERMAC.Services.Database;
using Microsoft.Extensions.Logging;

namespace AdminSERMAC.Forms
{
    public class TraspasosForm : Form
    {
        private readonly ILogger<TraspasosForm> _logger;
        private readonly IInventarioDatabaseService _inventarioService;
        private readonly SQLiteService _sqliteService;

        private DataGridView productosDataGridView;
        private Button confirmarTraspasoButton;
        private Label totalKilosLabel;
        private Label totalUnidadesLabel;

        private CheckBox traspasoCongeladosCheckBox;

        private ComboBox origenComboBox;
        private bool cargandoDatos = false;

        public TraspasosForm(ILogger<TraspasosForm> logger, IInventarioDatabaseService inventarioService, SQLiteService sqliteService)
        {
            _logger = logger;
            _inventarioService = inventarioService;
            _sqliteService = sqliteService;

            InitializeComponents();
            ConfigureEvents();
        }

        private void InitializeComponents()
        {
            this.Text = "Traspaso a Local";
            this.Size = new Size(1000, 600);
            this.StartPosition = FormStartPosition.CenterScreen;

            // DataGridView para productos
            productosDataGridView = new DataGridView
            {
                Location = new Point(20, 20),
                Size = new Size(940, 400),
                AllowUserToAddRows = true,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                MultiSelect = false
            };

            ConfigureGridColumns();

            // Labels para totales
            totalKilosLabel = new Label
            {
                Text = "Total Kilos: 0",
                Location = new Point(20, 440),
                AutoSize = true,
                Font = new Font("Segoe UI", 10F, FontStyle.Bold)
            };

            totalUnidadesLabel = new Label
            {
                Text = "Total Unidades: 0",
                Location = new Point(200, 440),
                AutoSize = true,
                Font = new Font("Segoe UI", 10F, FontStyle.Bold)
            };

            origenComboBox = new ComboBox
            {
                Location = new Point(20, 20),
                Size = new Size(200, 25),
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            origenComboBox.Items.AddRange(new string[] { "Inventario Normal", "Cámara Congelados" });
            origenComboBox.SelectedIndex = 0;

            // Botón confirmar
            confirmarTraspasoButton = new Button
            {
                Text = "Confirmar Traspaso",
                Location = new Point(20, 480),
                Size = new Size(150, 35),
                BackColor = Color.FromArgb(40, 167, 69),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat
            };

            traspasoCongeladosCheckBox = new CheckBox
            {
                Text = "Traspaso a Congelados",
                Location = new Point(200, 487),
                AutoSize = true,
                Font = new Font("Segoe UI", 10F)
            };

            traspasoCongeladosCheckBox.Text = "Traspaso a Congelados";

            this.Controls.Add(origenComboBox);

            productosDataGridView.Location = new Point(20, 55);
            productosDataGridView.Size = new Size(940, 365);

            // Agregar controles al formulario
            this.Controls.AddRange(new Control[] {
    productosDataGridView,
    totalKilosLabel,
    totalUnidadesLabel,
    confirmarTraspasoButton,
    traspasoCongeladosCheckBox
});
        }

        private async Task RegistrarHistorialTraspaso(
    SQLiteConnection connection,
    SQLiteTransaction transaction,
    string codigo,
    string descripcion,
    double kilos,
    int unidades,
    bool esOrigenCongelados,
    bool esTraspasoCongelados)
        {
            string origen = esOrigenCongelados ? "Cámara Congelados" : "Inventario Normal";
            string destino = esTraspasoCongelados ?
                (esOrigenCongelados ? "Inventario Normal" : "Cámara Congelados") :
                "Local";

            var command = new SQLiteCommand(@"
        INSERT INTO HistorialTraspasos (
            FechaTraspaso, Origen, Destino, 
            CodigoProducto, Descripcion, Kilos, Unidades
        ) VALUES (
            @fecha, @origen, @destino,
            @codigo, @descripcion, @kilos, @unidades
        )", connection, transaction);

            command.Parameters.AddWithValue("@fecha", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
            command.Parameters.AddWithValue("@origen", origen);
            command.Parameters.AddWithValue("@destino", destino);
            command.Parameters.AddWithValue("@codigo", codigo);
            command.Parameters.AddWithValue("@descripcion", descripcion);
            command.Parameters.AddWithValue("@kilos", kilos);
            command.Parameters.AddWithValue("@unidades", unidades);

            await command.ExecuteNonQueryAsync();
        }
        private void ConfigureGridColumns()
        {
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
                    ReadOnly = true,
                    Width = 250
                },
                new DataGridViewTextBoxColumn
                {
                    Name = "UnidadesDisponibles",
                    HeaderText = "Unid. Disponibles",
                    ReadOnly = true,
                    Width = 120
                },
                new DataGridViewTextBoxColumn
                {
                    Name = "UnidadesTraspaso",
                    HeaderText = "Unid. a Traspasar",
                    Width = 120
                },
                new DataGridViewTextBoxColumn
                {
                    Name = "KilosDisponibles",
                    HeaderText = "Kilos Disponibles",
                    ReadOnly = true,
                    Width = 120
                },
                new DataGridViewTextBoxColumn
                {
                    Name = "KilosTraspaso",
                    HeaderText = "Kilos a Traspasar",
                    Width = 120
                }
            });
        }

        private void ConfigureEvents()
        {
            confirmarTraspasoButton.Click += ConfirmarTraspasoButton_Click;
            productosDataGridView.CellEndEdit += ProductosDataGridView_CellEndEdit;
            productosDataGridView.CellValidating += ProductosDataGridView_CellValidating;
            origenComboBox.SelectedIndexChanged += OrigenComboBox_SelectedIndexChanged;
        }

        private void OrigenComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cargandoDatos) return;
            bool esCongelados = origenComboBox.SelectedIndex == 1;
            traspasoCongeladosCheckBox.Checked = false; // Siempre inicia desmarcado
            traspasoCongeladosCheckBox.Text = esCongelados ? "Traspaso a Inventario" : "Traspaso a Congelados";
            productosDataGridView.Rows.Clear();
        }

        private void ProductosDataGridView_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == productosDataGridView.Columns["Codigo"].Index)
            {
                string codigo = productosDataGridView.Rows[e.RowIndex].Cells["Codigo"].Value?.ToString();
                if (!string.IsNullOrEmpty(codigo))
                {
                    using (var connection = new SQLiteConnection(_sqliteService.connectionString))
                    {
                        connection.Open();
                        string tableName = origenComboBox.SelectedIndex == 1 ? "InventarioCongelados" : "Inventario";
                        var command = new SQLiteCommand(
                            $@"SELECT p.Nombre, i.Unidades, i.Kilos 
                       FROM Productos p 
                       JOIN {tableName} i ON p.Codigo = i.Codigo 
                       WHERE p.Codigo = @codigo", connection);
                        command.Parameters.AddWithValue("@codigo", codigo);

                        using (var reader = command.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                productosDataGridView.Rows[e.RowIndex].Cells["Descripcion"].Value = reader["Nombre"].ToString();
                                productosDataGridView.Rows[e.RowIndex].Cells["UnidadesDisponibles"].Value = reader["Unidades"].ToString();
                                productosDataGridView.Rows[e.RowIndex].Cells["KilosDisponibles"].Value = reader["Kilos"].ToString();
                            }
                            else
                            {
                                productosDataGridView.Rows[e.RowIndex].Cells["Codigo"].Value = null;
                                productosDataGridView.Rows[e.RowIndex].Cells["Descripcion"].Value = null;
                                productosDataGridView.Rows[e.RowIndex].Cells["UnidadesDisponibles"].Value = null;
                                productosDataGridView.Rows[e.RowIndex].Cells["KilosDisponibles"].Value = null;
                                MessageBox.Show("Código de producto no encontrado", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            }
                        }
                    }
                }
                ActualizarTotales();
            }
        }

        private void ProductosDataGridView_CellValidating(object sender, DataGridViewCellValidatingEventArgs e)
        {
            var row = productosDataGridView.Rows[e.RowIndex];

            if (e.ColumnIndex == productosDataGridView.Columns["UnidadesTraspaso"].Index)
            {
                if (!int.TryParse(e.FormattedValue.ToString(), out int valor))
                {
                    e.Cancel = true;
                    MessageBox.Show("Por favor ingrese un número entero válido.",
                        "Error de validación", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                if (valor < 0)
                {
                    e.Cancel = true;
                    MessageBox.Show("No se permiten valores negativos.",
                        "Error de validación", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                if (int.TryParse(row.Cells["UnidadesDisponibles"].Value?.ToString(), out int disponible))
                {
                    if (valor > disponible)
                    {
                        e.Cancel = true;
                        MessageBox.Show("Las unidades a traspasar no pueden ser mayores que las disponibles.",
                            "Error de validación", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }
            }
            else if (e.ColumnIndex == productosDataGridView.Columns["KilosTraspaso"].Index)
            {
                if (!double.TryParse(e.FormattedValue.ToString(), out double valor))
                {
                    e.Cancel = true;
                    MessageBox.Show("Por favor ingrese un valor numérico válido.",
                        "Error de validación", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                if (valor < 0)
                {
                    e.Cancel = true;
                    MessageBox.Show("No se permiten valores negativos.",
                        "Error de validación", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                if (double.TryParse(row.Cells["KilosDisponibles"].Value?.ToString(), out double disponible))
                {
                    if (valor > disponible)
                    {
                        e.Cancel = true;
                        MessageBox.Show("Los kilos a traspasar no pueden ser mayores que los disponibles.",
                            "Error de validación", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }
            }
        }

        private void ActualizarTotales()
        {
            double totalKilos = 0;
            int totalUnidades = 0;

            foreach (DataGridViewRow row in productosDataGridView.Rows)
            {
                if (row.IsNewRow) continue;

                if (double.TryParse(row.Cells["KilosTraspaso"].Value?.ToString(), out double kilos))
                {
                    totalKilos += kilos;
                }

                if (int.TryParse(row.Cells["UnidadesTraspaso"].Value?.ToString(), out int unidades))
                {
                    totalUnidades += unidades;
                }
            }

            totalKilosLabel.Text = $"Total Kilos: {totalKilos:N2}";
            totalUnidadesLabel.Text = $"Total Unidades: {totalUnidades}";
        }

        private async void ConfirmarTraspasoButton_Click(object sender, EventArgs e)
        {
            try
            {
                if (productosDataGridView.Rows.Count == 1)
                {
                    MessageBox.Show("Agregue al menos un producto para realizar el traspaso.",
                        "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                bool esOrigenCongelados = origenComboBox.SelectedIndex == 1;

                var resultado = MessageBox.Show(
                    traspasoCongeladosCheckBox.Checked ?
                    (esOrigenCongelados ? "¿Está seguro de confirmar el traspaso a inventario? Esta acción descontará los productos de congelados." :
                                        "¿Está seguro de confirmar el traspaso a congelados? Esta acción descontará los productos del inventario.") :
                    $"¿Está seguro de confirmar el traspaso a local? Esta acción descontará los productos de {(esOrigenCongelados ? "congelados" : "inventario")}.",
                    "Confirmar Traspaso",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question
                );

                if (resultado == DialogResult.Yes)
                {
                    using (var connection = new SQLiteConnection(_sqliteService.connectionString))
                    {
                        await connection.OpenAsync();
                        using (var transaction = connection.BeginTransaction())
                        {
                            try
                            {
                                foreach (DataGridViewRow row in productosDataGridView.Rows)
                                {
                                    if (row.IsNewRow) continue;

                                    string codigo = row.Cells["Codigo"].Value?.ToString();
                                    if (string.IsNullOrEmpty(codigo)) continue;

                                    int unidades = Convert.ToInt32(row.Cells["UnidadesTraspaso"].Value);
                                    double kilos = Convert.ToDouble(row.Cells["KilosTraspaso"].Value);

                                    if (unidades > 0 || kilos > 0)
                                    {
                                        if (!traspasoCongeladosCheckBox.Checked)
                                        {
                                            string tableName = esOrigenCongelados ? "InventarioCongelados" : "Inventario";
                                            var updateOrigen = new SQLiteCommand($@"
                                       UPDATE {tableName}
                                       SET Unidades = Unidades - @unidades,
                                           Kilos = Kilos - @kilos
                                       WHERE Codigo = @codigo", connection, transaction);

                                            updateOrigen.Parameters.AddWithValue("@codigo", codigo);
                                            updateOrigen.Parameters.AddWithValue("@unidades", unidades);
                                            updateOrigen.Parameters.AddWithValue("@kilos", kilos);
                                            updateOrigen.ExecuteNonQuery();
                                        }
                                        else
                                        {
                                            if (!esOrigenCongelados)
                                            {
                                                var updateInventario = new SQLiteCommand(@"
                                           UPDATE Inventario 
                                           SET Unidades = Unidades - @unidades,
                                               Kilos = Kilos - @kilos
                                           WHERE Codigo = @codigo", connection, transaction);

                                                updateInventario.Parameters.AddWithValue("@codigo", codigo);
                                                updateInventario.Parameters.AddWithValue("@unidades", unidades);
                                                updateInventario.Parameters.AddWithValue("@kilos", kilos);
                                                updateInventario.ExecuteNonQuery();

                                                var updateCongelados = new SQLiteCommand(@"
                                           UPDATE InventarioCongelados 
                                           SET Unidades = Unidades + @unidades,
                                               Kilos = Kilos + @kilos
                                           WHERE Codigo = @codigo", connection, transaction);

                                                updateCongelados.Parameters.AddWithValue("@codigo", codigo);
                                                updateCongelados.Parameters.AddWithValue("@unidades", unidades);
                                                updateCongelados.Parameters.AddWithValue("@kilos", kilos);
                                                updateCongelados.ExecuteNonQuery();
                                            }
                                            else
                                            {
                                                var updateCongelados = new SQLiteCommand(@"
                                           UPDATE InventarioCongelados 
                                           SET Unidades = Unidades - @unidades,
                                               Kilos = Kilos - @kilos
                                           WHERE Codigo = @codigo", connection, transaction);

                                                updateCongelados.Parameters.AddWithValue("@codigo", codigo);
                                                updateCongelados.Parameters.AddWithValue("@unidades", unidades);
                                                updateCongelados.Parameters.AddWithValue("@kilos", kilos);
                                                updateCongelados.ExecuteNonQuery();

                                                var updateInventario = new SQLiteCommand(@"
                                           UPDATE Inventario 
                                           SET Unidades = Unidades + @unidades,
                                               Kilos = Kilos + @kilos
                                           WHERE Codigo = @codigo", connection, transaction);

                                                updateInventario.Parameters.AddWithValue("@codigo", codigo);
                                                updateInventario.Parameters.AddWithValue("@unidades", unidades);
                                                updateInventario.Parameters.AddWithValue("@kilos", kilos);
                                                updateInventario.ExecuteNonQuery();
                                            }
                                        }

                                        await RegistrarHistorialTraspaso(
                                            connection,
                                            transaction,
                                            codigo,
                                            row.Cells["Descripcion"].Value.ToString(),
                                            kilos,
                                            unidades,
                                            esOrigenCongelados,
                                            traspasoCongeladosCheckBox.Checked
                                        );
                                    }
                                }

                                transaction.Commit();
                                MessageBox.Show("Traspaso realizado exitosamente.",
                                    "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);

                                this.DialogResult = DialogResult.OK;
                                this.Close();
                            }
                            catch
                            {
                                transaction.Rollback();
                                throw;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al confirmar el traspaso");
                MessageBox.Show("Error al procesar el traspaso: " + ex.Message,
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}