using System;
using System.Windows.Forms;
using AdminSERMAC.Forms;
using AdminSERMAC.Core.Interfaces;
using AdminSERMAC.Services;
using System.Drawing;
using AdminSERMAC.Core.Theme;
using System.IO;
using Microsoft.Extensions.Logging;
using AdminSERMAC.Core.Infrastructure;
using AdminSERMAC.Services.Database;

namespace AdminSERMAC.Forms
{
    public class MainForm : Form
    {
        private readonly IClienteService _clienteService;
        private readonly ILogger<MainForm> _logger;
        private readonly ILogger<SQLiteService> _sqliteLogger;
        private readonly SQLiteService _sqliteService;  // Cambiado de _sqliteService
        private readonly ILoggerFactory _loggerFactory;
        private readonly IInventarioDatabaseService _inventarioService;
        private readonly IProductoDatabaseService _productoDatabaseService;



        private Button mostrarClientesButton;
        private Button mostrarVentasButton;
        private Button mostrarInventarioButton;
        private Button visualizarProductoButton;
        private Button importarProductosButton;
        private Button cambiarTemaButton;
        private Button visualizarGuiasButton;

        public MainForm(
    IClienteService clienteService,
    SQLiteService sqliteService,
    ILogger<MainForm> logger,
    ILogger<SQLiteService> sqliteLogger,
    ILoggerFactory loggerFactory,
    IInventarioDatabaseService inventarioService,
    IProductoDatabaseService productoDatabaseService)
        {
            _clienteService = clienteService ?? throw new ArgumentNullException(nameof(clienteService));
            _sqliteService = sqliteService ?? throw new ArgumentNullException(nameof(sqliteService));
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
            _loggerFactory = loggerFactory ?? throw new ArgumentNullException(nameof(loggerFactory));
            _inventarioService = inventarioService ?? throw new ArgumentNullException(nameof(inventarioService));
            _productoDatabaseService = productoDatabaseService ?? throw new ArgumentNullException(nameof(productoDatabaseService));

            InitializeComponents();
            ThemeManager.ApplyTheme(this);
        }

        private void InitializeComponents()
        {
            // Configuración del formulario
            this.Text = "Menú Principal - SERMAC";
            this.Width = 800;
            this.Height = 700; // Aumentado para acomodar el nuevo botón
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;

            // Panel principal
            Panel mainPanel = new Panel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(20)
            };
            this.Controls.Add(mainPanel);

            // Título
            Label titleLabel = new Label
            {
                Text = "Sistema de Gestión SERMAC",
                Font = new Font("Arial", 24, FontStyle.Bold),
                AutoSize = true,
                Location = new Point(50, 20),
                ForeColor = Color.FromArgb(0, 122, 204)
            };
            mainPanel.Controls.Add(titleLabel);

            // Botones principales
            mostrarClientesButton = CreateMenuButton("Gestión de Clientes", 100);
            mostrarVentasButton = CreateMenuButton("Gestión de Ventas", 170);
            mostrarInventarioButton = CreateMenuButton("Gestión de Inventario", 240);
            visualizarProductoButton = CreateMenuButton("Visualizar Producto", 310);
            visualizarGuiasButton = CreateMenuButton("Visualizar Guías", 380);
            importarProductosButton = CreateMenuButton("Importar Productos", 450);
            cambiarTemaButton = CreateMenuButton(ThemeManager.IsDarkMode ? "Tema Claro" : "Tema Oscuro", 520);

            // Configurar colores especiales para ciertos botones
            visualizarGuiasButton.BackColor = Color.FromArgb(70, 130, 180); // Steel Blue
            importarProductosButton.BackColor = Color.FromArgb(51, 122, 183);
            cambiarTemaButton.BackColor = Color.FromArgb(68, 157, 68);
            cambiarTemaButton.Tag = "primary";

            // Agregar botones al panel
            mainPanel.Controls.AddRange(new Control[] {
                mostrarClientesButton,
                mostrarVentasButton,
                mostrarInventarioButton,
                visualizarProductoButton,
                visualizarGuiasButton,
                importarProductosButton,
                cambiarTemaButton
            });

            Button reportesButton = CreateMenuButton("Reportes y Estadísticas", 590);
            reportesButton.BackColor = Color.FromArgb(75, 139, 59);
            reportesButton.Click += (s, e) => {
                try
                {
                    var reportLogger = _loggerFactory.CreateLogger<ReportGenerator>();
                    var reportGenerator = new ReportGenerator(_sqliteService, reportLogger);
                    var reportesForm = new ReportesForm(reportGenerator, _sqliteService);
                    reportesForm.Show();
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Error al abrir el formulario de reportes: {ex.Message}",
                        "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            };
            mainPanel.Controls.Add(reportesButton);

            // Eventos
            mostrarClientesButton.Click += MostrarClientesButton_Click;
            mostrarVentasButton.Click += MostrarVentasButton_Click;
            mostrarInventarioButton.Click += MostrarInventarioButton_Click;
            visualizarProductoButton.Click += VisualizarProductoButton_Click;
            visualizarGuiasButton.Click += VisualizarGuiasButton_Click;
            importarProductosButton.Click += ImportarProductosButton_Click;
            cambiarTemaButton.Click += CambiarTemaButton_Click;
        }

        private Button CreateMenuButton(string text, int top)
        {
            var button = new Button
            {
                Text = text,
                Top = top,
                Left = 50,
                Width = 250,
                Height = 45,
                Font = new Font("Arial", 11, FontStyle.Bold),
                BackColor = Color.FromArgb(0, 122, 204),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand,
                TextAlign = ContentAlignment.MiddleCenter
            };

            button.FlatAppearance.BorderSize = 0;
            button.FlatAppearance.MouseOverBackColor = Color.FromArgb(0, 102, 204);

            return button;
        }

        private void VisualizarGuiasButton_Click(object sender, EventArgs e)
        {
            try
            {
                var sqliteLogger = _loggerFactory.CreateLogger<SQLiteService>();
                var visualizarGuiasForm = new VisualizarGuiasForm(sqliteLogger);
                visualizarGuiasForm.Show();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error al abrir el visualizador de guías: {ex.Message}",
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


        private void CambiarTemaButton_Click(object sender, EventArgs e)
        {
            ThemeManager.IsDarkMode = !ThemeManager.IsDarkMode;
            ThemeManager.ApplyTheme(this);
            cambiarTemaButton.Text = ThemeManager.IsDarkMode ? "Tema Claro" : "Tema Oscuro";
        }

        private void MostrarClientesButton_Click(object sender, EventArgs e)
        {
            try
            {
                var clientesLogger = _loggerFactory.CreateLogger<ClientesForm>();
                var clienteService = _clienteService; // Asegúrate de que MainForm tenga acceso a IClienteService
                var clientesForm = new ClientesForm(_sqliteService, clientesLogger, clienteService);
                clientesForm.Show();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error al abrir el formulario de clientes: {ex.Message}",
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }



        private void MostrarVentasButton_Click(object sender, EventArgs e)
        {
            try
            {
                var sqliteLogger = _loggerFactory.CreateLogger<SQLiteService>();
                var ventasForm = new VentasForm(sqliteLogger);
                ventasForm.Show();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error al abrir el formulario de ventas: {ex.Message}",
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


        private void MostrarInventarioButton_Click(object sender, EventArgs e)
        {
            try
            {
                var inventarioLogger = _loggerFactory.CreateLogger<InventarioForm>();
                var inventarioForm = new InventarioForm(
                    inventarioLogger,
                    _sqliteService,
                    _inventarioService,
                    _productoDatabaseService);
                inventarioForm.Show();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al abrir el formulario de inventario");
                MessageBox.Show($"Error al abrir el formulario de inventario: {ex.Message}",
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void VisualizarProductoButton_Click(object sender, EventArgs e)
        {
            try
            {
                var sqliteLogger = _loggerFactory.CreateLogger<SQLiteService>();
                var visualizarProductoForm = new VisualizarProductoForm(sqliteLogger);
                visualizarProductoForm.Show();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error al abrir el formulario de productos: {ex.Message}",
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ImportarProductosButton_Click(object sender, EventArgs e)
        {
            try
            {
                using (OpenFileDialog openFileDialog = new OpenFileDialog())
                {
                    openFileDialog.Filter = "Archivos de texto (*.txt;*.csv)|*.txt;*.csv|Todos los archivos (*.*)|*.*";
                    openFileDialog.Title = "Seleccionar archivo de productos";

                    if (openFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        // Leemos todas las líneas del archivo
                        string[] allLines = File.ReadAllLines(openFileDialog.FileName, System.Text.Encoding.UTF8);

                        // Verificamos que tengamos contenido
                        if (allLines == null || allLines.Length < 2)
                        {
                            MessageBox.Show("El archivo está vacío o no contiene datos suficientes", "Error",
                                MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }

                        // Verificamos que la primera línea contenga los headers esperados
                        string[] headers = allLines[0].Split(',');
                        if (!headers.Contains("ID") || !headers.Contains("Nombre"))
                        {
                            MessageBox.Show("El archivo no contiene el formato correcto. Se requieren las columnas 'ID' y 'Nombre'",
                                "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }

                        // Construimos el contenido para la importación
                        string csvContent = string.Join(Environment.NewLine, allLines);

                        // Intentamos la importación
                        _sqliteService.ImportarProductosDesdeCSV(csvContent);

                        MessageBox.Show($"Se importaron exitosamente {allLines.Length - 1} productos", "Éxito",
                            MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al importar productos");
                MessageBox.Show($"Error al importar productos: {ex.Message}",
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            base.OnFormClosing(e);
            if (e.CloseReason == CloseReason.UserClosing)
            {
                if (MessageBox.Show("¿Está seguro que desea salir de la aplicación?",
                    "Confirmar salida", MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question) == DialogResult.No)
                {
                    e.Cancel = true;
                }
            }
        }
    }
}
