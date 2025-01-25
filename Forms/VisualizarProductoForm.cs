using System;
using System.Data;
using System.Data.SQLite;
using System.Drawing;
using System.Windows.Forms;
using AdminSERMAC.Services;
using Microsoft.Extensions.Logging;

namespace AdminSERMAC.Forms
{
    public class VisualizarProductoForm : Form
    {
        private DataGridView productoDataGridView;
        private SQLiteService sqliteService;

        public VisualizarProductoForm(ILogger<SQLiteService> logger)
        {
            this.Text = "Visualizar Productos";
            this.Width = 1000;
            this.Height = 700;

            sqliteService = new SQLiteService(logger);

            InitializeComponents();
            CargarProductos();
        }

        private void InitializeComponents()
        {
            productoDataGridView = new DataGridView
            {
                Dock = DockStyle.Fill,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                ReadOnly = true,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                BackgroundColor = Color.White,
                RowHeadersVisible = false
            };

            this.Controls.Add(productoDataGridView);
        }

        private void CargarProductos()
        {
            try
            {
                using (var connection = new SQLiteConnection(sqliteService.connectionString))
                {
                    connection.Open();

                    string query = @"
                        SELECT 
                            Codigo AS 'Código',
                            Nombre AS 'Nombre',
                            Categoria AS 'Categoría',
                            SubCategoria AS 'SubCategoría',
                            Precio AS 'Precio',
                            UnidadMedida AS 'Unidad de Medida'
                        FROM Productos
                        ORDER BY Codigo";

                    var command = new SQLiteCommand(query, connection);
                    var adapter = new SQLiteDataAdapter(command);

                    var dataTable = new DataTable();
                    adapter.Fill(dataTable);

                    productoDataGridView.DataSource = dataTable;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error al cargar los productos: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}

