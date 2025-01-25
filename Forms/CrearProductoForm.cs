using AdminSERMAC.Services;
using System.Data.SQLite;

public class CrearProductoForm : Form
{
    private TextBox codigoTextBox;
    private TextBox nombreTextBox;
    private TextBox marcaTextBox;
    private ComboBox unidadMedidaComboBox;
    private ComboBox categoriaComboBox;
    private ComboBox subcategoriaComboBox;
    private Button guardarButton;
    private readonly SQLiteService sqliteService;
    private readonly ILogger<CrearProductoForm> _logger;

    public CrearProductoForm(ILogger<CrearProductoForm> logger)
    {
        _logger = logger ?? throw new ArgumentNullException(nameof(logger));
        // Creamos directamente una nueva instancia de SQLiteService con un nuevo logger
        var sqliteLogger = LoggerFactory.Create(builder => builder.AddConsole().AddDebug())
                                      .CreateLogger<SQLiteService>();
        sqliteService = new SQLiteService(sqliteLogger);

        this.Text = "Crear Producto";
        this.Width = 400;
        this.Height = 450;

        InitializeComponents();
        CargarCategorias();
    }

    private void InitializeComponents()
    {
        // Código de inicialización de componentes...
        Label codigoLabel = new Label { Text = "Código:", Top = 20, Left = 20, Width = 100 };
        codigoTextBox = new TextBox { Top = 20, Left = 130, Width = 200 };

        Label nombreLabel = new Label { Text = "Nombre:", Top = 60, Left = 20, Width = 100 };
        nombreTextBox = new TextBox { Top = 60, Left = 130, Width = 200 };

        Label marcaLabel = new Label { Text = "Marca:", Top = 100, Left = 20, Width = 100 };
        marcaTextBox = new TextBox { Top = 100, Left = 130, Width = 200 };

        Label unidadMedidaLabel = new Label { Text = "Unidad de Medida:", Top = 140, Left = 20, Width = 120 };
        unidadMedidaComboBox = new ComboBox
        {
            Top = 140,
            Left = 150,
            Width = 180,
            DropDownStyle = ComboBoxStyle.DropDownList
        };
        unidadMedidaComboBox.Items.Add("Kg");
        unidadMedidaComboBox.Items.Add("Unidades");
        unidadMedidaComboBox.SelectedIndex = 0;

        Label categoriaLabel = new Label { Text = "Categoría:", Top = 180, Left = 20, Width = 100 };
        categoriaComboBox = new ComboBox { Top = 180, Left = 130, Width = 200, DropDownStyle = ComboBoxStyle.DropDownList };

        Label subcategoriaLabel = new Label { Text = "Subcategoría:", Top = 220, Left = 20, Width = 100 };
        subcategoriaComboBox = new ComboBox { Top = 220, Left = 130, Width = 200, DropDownStyle = ComboBoxStyle.DropDownList };

        guardarButton = new Button
        {
            Text = "Guardar",
            Top = 280,
            Left = 130,
            Width = 100,
            BackColor = Color.FromArgb(0, 122, 204),
            ForeColor = Color.White
        };
        guardarButton.Click += GuardarButton_Click;

        this.Controls.AddRange(new Control[]
        {
            codigoLabel, codigoTextBox,
            nombreLabel, nombreTextBox,
            marcaLabel, marcaTextBox,
            unidadMedidaLabel, unidadMedidaComboBox,
            categoriaLabel, categoriaComboBox,
            subcategoriaLabel, subcategoriaComboBox,
            guardarButton
        });
    }

    private void CargarCategorias()
    {
        try
        {
            categoriaComboBox.Items.Clear();
            categoriaComboBox.Items.Add("Seleccionar");

            // Obtener categorías existentes
            var categorias = sqliteService.GetCategorias();
            if (categorias != null)
            {
                // Agregar las nuevas categorías si no existen
                if (!categorias.Contains("Embutidos"))
                    categorias.Add("Embutidos");
                if (!categorias.Contains("Congelados"))
                    categorias.Add("Congelados");

                categoriaComboBox.Items.AddRange(categorias.ToArray());
            }

            categoriaComboBox.SelectedIndex = 0;
            categoriaComboBox.SelectedIndexChanged += (s, e) =>
            {
                subcategoriaComboBox.Items.Clear();
                subcategoriaComboBox.Items.Add("Seleccionar");

                if (categoriaComboBox.SelectedIndex > 0)
                {
                    // Primero obtener las subcategorías existentes de la base de datos
                    var subcategoriasExistentes = sqliteService.GetSubCategorias(categoriaComboBox.SelectedItem.ToString());
                    var todasSubcategorias = new List<string>(subcategoriasExistentes);

                    // Agregar las nuevas subcategorías según la categoría seleccionada
                    if (categoriaComboBox.SelectedItem.ToString() == "Embutidos")
                    {
                        var nuevasSubcategorias = new[] { "Longanizas", "Vienesas", "Otro" };
                        foreach (var sub in nuevasSubcategorias)
                        {
                            if (!todasSubcategorias.Contains(sub))
                                todasSubcategorias.Add(sub);
                        }
                    }
                    else if (categoriaComboBox.SelectedItem.ToString() == "Congelados")
                    {
                        var nuevasSubcategorias = new[] { "Verduras", "Pollo", "Vacuno", "Cerdo", "Pavo", "Pescado", "Fritos", "Otro" };
                        foreach (var sub in nuevasSubcategorias)
                        {
                            if (!todasSubcategorias.Contains(sub))
                                todasSubcategorias.Add(sub);
                        }
                    }

                    subcategoriaComboBox.Items.AddRange(todasSubcategorias.ToArray());
                }
                subcategoriaComboBox.SelectedIndex = 0;
            };
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error al cargar categorías");
            MessageBox.Show("Error al cargar las categorías: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private void GuardarButton_Click(object sender, EventArgs e)
    {
        try
        {
            if (!ValidarCampos())
                return;

            using (var connection = new SQLiteConnection(sqliteService.connectionString))
            {
                connection.Open();
                using (var transaction = connection.BeginTransaction())
                {
                    try
                    {
                        // Primero verificamos si el código ya existe
                        using (var checkCommand = new SQLiteCommand(
                            "SELECT COUNT(*) FROM Productos WHERE Codigo = @codigo",
                            connection))
                        {
                            checkCommand.Parameters.AddWithValue("@codigo", codigoTextBox.Text.Trim());
                            int count = Convert.ToInt32(checkCommand.ExecuteScalar());
                            if (count > 0)
                            {
                                MessageBox.Show("Ya existe un producto con este código.",
                                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                return;
                            }
                        }

                        // Si no existe, procedemos a insertar
                        using (var command = new SQLiteCommand(
                            @"INSERT INTO Productos (
                            Codigo, 
                            Nombre, 
                            Marca, 
                            UnidadMedida, 
                            Categoria, 
                            SubCategoria,
                            Precio
                        ) VALUES (
                            @codigo,
                            @nombre,
                            @marca,
                            @unidadMedida,
                            @categoria,
                            @subcategoria,
                            @precio
                        )", connection))
                        {
                            command.Parameters.Add(new SQLiteParameter("@codigo", DbType.String) { Value = codigoTextBox.Text.Trim() });
                            command.Parameters.Add(new SQLiteParameter("@nombre", DbType.String) { Value = nombreTextBox.Text.Trim() });
                            command.Parameters.Add(new SQLiteParameter("@marca", DbType.String) { Value = marcaTextBox.Text.Trim() });
                            command.Parameters.Add(new SQLiteParameter("@unidadMedida", DbType.String) { Value = unidadMedidaComboBox.SelectedItem?.ToString() ?? "" });
                            command.Parameters.Add(new SQLiteParameter("@categoria", DbType.String) { Value = categoriaComboBox.SelectedItem?.ToString() ?? "" });
                            command.Parameters.Add(new SQLiteParameter("@subcategoria", DbType.String) { Value = subcategoriaComboBox.SelectedItem?.ToString() ?? "" });
                            command.Parameters.Add(new SQLiteParameter("@precio", DbType.Double) { Value = 0.0 });

                            command.ExecuteNonQuery();
                        }

                        transaction.Commit();
                        MessageBox.Show("Producto guardado exitosamente.",
                            "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        this.DialogResult = DialogResult.OK;
                        this.Close();
                    }
                    catch (Exception ex)
                    {
                        transaction.Rollback();
                        throw new Exception("Error en la base de datos: " + ex.Message);
                    }
                }
            }
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error al guardar el producto");
            MessageBox.Show($"Error al guardar el producto: {ex.Message}",
                "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private bool ValidarCampos()
    {
        if (string.IsNullOrWhiteSpace(codigoTextBox.Text))
        {
            MessageBox.Show("El código es obligatorio.", "Validación", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            codigoTextBox.Focus();
            return false;
        }

        if (string.IsNullOrWhiteSpace(nombreTextBox.Text))
        {
            MessageBox.Show("El nombre es obligatorio.", "Validación", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            nombreTextBox.Focus();
            return false;
        }

        if (string.IsNullOrWhiteSpace(marcaTextBox.Text))
        {
            MessageBox.Show("La marca es obligatoria.", "Validación", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            marcaTextBox.Focus();
            return false;
        }

        if (unidadMedidaComboBox.SelectedIndex < 0)
        {
            MessageBox.Show("Debe seleccionar una unidad de medida.", "Validación", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            unidadMedidaComboBox.Focus();
            return false;
        }

        if (categoriaComboBox.SelectedIndex <= 0 || categoriaComboBox.SelectedItem?.ToString() == "Seleccionar")
        {
            MessageBox.Show("Debe seleccionar una categoría válida.", "Validación", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            categoriaComboBox.Focus();
            return false;
        }

        if (subcategoriaComboBox.SelectedIndex <= 0 || subcategoriaComboBox.SelectedItem?.ToString() == "Seleccionar")
        {
            MessageBox.Show("Debe seleccionar una subcategoría válida.", "Validación", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            subcategoriaComboBox.Focus();
            return false;
        }

        return true;
    }
}