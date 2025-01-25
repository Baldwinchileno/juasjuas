using AdminSERMAC.Models;
using AdminSERMAC.Services;
using AdminSERMAC.Services.Database;
using System.Drawing;
using Microsoft.Extensions.Logging;

public class ModificarProductoForm : Form
{
    private readonly IProductoDatabaseService _productoDatabaseService;
    private readonly IInventarioDatabaseService _inventarioDatabaseService;
    private readonly string codigoOriginal;

    // Controles del formulario
    private Label titleLabel;
    private TableLayoutPanel mainTableLayout;
    private TextBox codigoTextBox;
    private TextBox nombreTextBox;
    private TextBox marcaTextBox;
    private TextBox categoriaTextBox;
    private TextBox subCategoriaTextBox;
    private TextBox unidadMedidaTextBox;
    private NumericUpDown precioNumeric;
    private Button guardarButton;
    private Button cancelarButton;

    public ModificarProductoForm(
        string codigo,
        ILogger<SQLiteService> logger,
        IProductoDatabaseService productoDatabaseService,
        IInventarioDatabaseService inventarioDatabaseService)
    {
        _productoDatabaseService = productoDatabaseService;
        _inventarioDatabaseService = inventarioDatabaseService;
        codigoOriginal = codigo;

        InitializeComponents();
        CargarDatosProducto().Wait();
    }

    private void InitializeComponents()
    {
        // Configuración del formulario
        this.Text = "Modificar Producto";
        this.Size = new Size(500, 500);
        this.StartPosition = FormStartPosition.CenterScreen;
        this.FormBorderStyle = FormBorderStyle.FixedDialog;
        this.MaximizeBox = false;
        this.MinimizeBox = false;

        // Título
        titleLabel = new Label
        {
            Text = "Modificar Producto",
            Font = new Font("Segoe UI", 16, FontStyle.Bold),
            ForeColor = Color.FromArgb(0, 122, 204),
            Dock = DockStyle.Top,
            TextAlign = ContentAlignment.MiddleCenter,
            Height = 50
        };

        // TableLayoutPanel principal
        mainTableLayout = new TableLayoutPanel
        {
            ColumnCount = 2,
            RowCount = 8,
            Dock = DockStyle.Fill,
            Padding = new Padding(20),
            ColumnStyles = {
                new ColumnStyle(SizeType.Percent, 30),
                new ColumnStyle(SizeType.Percent, 70)
            }
        };

        // Crear y configurar los controles
        CreateFormControls();

        // Configurar los botones
        var buttonPanel = new FlowLayoutPanel
        {
            Dock = DockStyle.Bottom,
            Height = 60,
            FlowDirection = FlowDirection.RightToLeft,
            Padding = new Padding(10)
        };

        guardarButton = new Button
        {
            Text = "Guardar",
            Size = new Size(100, 35),
            BackColor = Color.FromArgb(0, 122, 204),
            ForeColor = Color.White,
            Font = new Font("Segoe UI", 9, FontStyle.Regular)
        };

        cancelarButton = new Button
        {
            Text = "Cancelar",
            Size = new Size(100, 35),
            Font = new Font("Segoe UI", 9, FontStyle.Regular)
        };

        // Agregar los botones al panel
        buttonPanel.Controls.Add(cancelarButton);
        buttonPanel.Controls.Add(new Label { Width = 10 }); // Espaciador
        buttonPanel.Controls.Add(guardarButton);

        // Configurar eventos
        guardarButton.Click += GuardarButton_Click;
        cancelarButton.Click += (s, e) => this.Close();

        // Agregar controles al formulario
        this.Controls.Add(titleLabel);
        this.Controls.Add(mainTableLayout);
        this.Controls.Add(buttonPanel);
    }

    private void CreateFormControls()
    {
        // Código
        AddFormField("Código:", codigoTextBox = new TextBox { ReadOnly = true }, 0);

        // Nombre
        AddFormField("Nombre:", nombreTextBox = new TextBox(), 1);

        // Marca
        AddFormField("Marca:", marcaTextBox = new TextBox(), 2);

        // Categoría
        AddFormField("Categoría:", categoriaTextBox = new TextBox(), 3);

        // SubCategoría
        AddFormField("SubCategoría:", subCategoriaTextBox = new TextBox(), 4);

        // Unidad de Medida
        AddFormField("Unidad de Medida:", unidadMedidaTextBox = new TextBox(), 5);

        // Precio
        precioNumeric = new NumericUpDown
        {
            Minimum = 0,
            Maximum = 999999999,
            DecimalPlaces = 2,
            ThousandsSeparator = true
        };
        AddFormField("Precio:", precioNumeric, 6);
    }

    private void AddFormField(string labelText, Control control, int row)
    {
        var label = new Label
        {
            Text = labelText,
            Anchor = AnchorStyles.Left | AnchorStyles.Right,
            TextAlign = ContentAlignment.MiddleRight
        };

        mainTableLayout.Controls.Add(label, 0, row);
        mainTableLayout.Controls.Add(control, 1, row);

        if (control is TextBox textBox)
        {
            textBox.Anchor = AnchorStyles.Left | AnchorStyles.Right;
            textBox.Width = 250;
        }
        else if (control is NumericUpDown numericUpDown)
        {
            numericUpDown.Anchor = AnchorStyles.Left | AnchorStyles.Right;
            numericUpDown.Width = 250;
        }
    }

    private async Task CargarDatosProducto()
    {
        try
        {
            var producto = await _productoDatabaseService.GetByCodigo(codigoOriginal);
            if (producto != null)
            {
                codigoTextBox.Text = producto.Codigo;
                nombreTextBox.Text = producto.Nombre;
                marcaTextBox.Text = producto.Marca;
                categoriaTextBox.Text = producto.Categoria;
                subCategoriaTextBox.Text = producto.SubCategoria;
                unidadMedidaTextBox.Text = producto.UnidadMedida;
                precioNumeric.Value = (decimal)producto.Precio;
            }
            else
            {
                MessageBox.Show("No se encontró el producto especificado.",
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                this.Close();
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Error al cargar los datos del producto: {ex.Message}",
                "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            this.Close();
        }
    }

    private async void GuardarButton_Click(object sender, EventArgs e)
    {
        try
        {
            if (string.IsNullOrWhiteSpace(nombreTextBox.Text))
            {
                MessageBox.Show("El nombre del producto es obligatorio.",
                    "Validación", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var producto = new Producto
            {
                Codigo = codigoTextBox.Text,
                Nombre = nombreTextBox.Text.Trim(),
                Marca = marcaTextBox.Text.Trim(),
                Categoria = categoriaTextBox.Text.Trim(),
                SubCategoria = subCategoriaTextBox.Text.Trim(),
                UnidadMedida = unidadMedidaTextBox.Text.Trim(),
                Precio = (double)precioNumeric.Value
            };

            await _productoDatabaseService.UpdateAsync(producto);

            MessageBox.Show("Producto actualizado exitosamente",
                "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
            this.DialogResult = DialogResult.OK;
            this.Close();
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Error al actualizar el producto: {ex.Message}",
                "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }
}