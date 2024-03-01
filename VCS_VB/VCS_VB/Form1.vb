Imports Microsoft.Win32
Imports System
Imports System.Collections.Generic
Imports System.ComponentModel
Imports System.Data
Imports System.Diagnostics
Imports System.Drawing
Imports System.IO
Imports System.Linq
Imports System.Security.Policy
Imports System.Text
Imports System.Threading.Tasks
Imports System.Windows.Forms
Imports System.Windows.Forms.VisualStyles.VisualStyleElement
Imports VCS_VB.CSV_VB

Namespace CSV
    Partial Public Class Form1
        Inherits Form

        Private registros As List(Of registros) = New List(Of registros)()
        Private rutaArchivoActual As String = ""
        Private formato As String
        Private Sub aGREGARToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs)
            If String.IsNullOrWhiteSpace(txtNombre.Text) OrElse String.IsNullOrWhiteSpace(txtTelefono.Text) OrElse String.IsNullOrWhiteSpace(txtCorreo.Text) Then
                MessageBox.Show("Por favor, complete todos los campos antes de agregar.", "Error", MessageBoxButtons.OK, MessageBoxIcon.[Error])
                Return
            End If

            Dim nuevoRegistro As registros = New registros With {
                .Nombre = txtNombre.Text,
                .Telefono = txtTelefono.Text,
                .Correo = txtCorreo.Text
            }
            registros.Add(nuevoRegistro)
            dgvDatos.DataSource = Nothing
            dgvDatos.DataSource = registros
            LimpiarCampos()
        End Sub

        Private Sub gUARDARToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs)
            Dim NombreA As String = txtCorreo.Text

            Try
                Dim escritorio As String = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
                Dim rutaArchivo As String = Path.Combine(escritorio, NombreA & "." & formato)

                Using writer As StreamWriter = New StreamWriter(rutaArchivo)
                    writer.WriteLine("Nombre,Telefono,Correo")

                    For Each registro As registros In registros
                        writer.WriteLine($"{registro.Nombre},{registro.Telefono},{registro.Correo}")
                    Next
                End Using

                MessageBox.Show($"Datos guardados exitosamente en el archivo CSV en el escritorio ({rutaArchivo}).", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information)
                LimpiarCampos()
            Catch ex As Exception
                MessageBox.Show($"Error al guardar en el archivo CSV: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.[Error])
            End Try

            dgvDatos.DataSource = Nothing
            dgvDatos.Rows.Clear()
            registros.Clear()
        End Sub

        Private Sub LimpiarCampos()
            txtNombre.Text = ""
            txtTelefono.Text = ""
            txtCorreo.Text = ""
            txtCorreo.Text = ""
        End Sub

        Private Sub aBRIRToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs)
            Dim openFileDialog As OpenFileDialog = New OpenFileDialog With {
                .Filter = "Archivos CSV |*.csv |Archivos Txt|*.txt|Archivos xml|*.xml|Archivos json|*.json",
                .Title = "Archivos Cargados Correctamente."
            }

            If openFileDialog.ShowDialog() = DialogResult.OK Then
                Dim rutaArchivo As String = openFileDialog.FileName

                Using reader As StreamReader = New StreamReader(rutaArchivo)
                    reader.ReadLine()
                    registros.Clear()

                    While Not reader.EndOfStream
                        Dim campos As String() = reader.ReadLine().Split(","c)
                        Dim nuevoRegistro As registros = New registros With {
                            .Nombre = campos(0),
                            .Telefono = campos(1),
                            .Correo = campos(2)
                        }
                        registros.Add(nuevoRegistro)
                    End While
                End Using

                dgvDatos.DataSource = Nothing
                dgvDatos.DataSource = registros
                MessageBox.Show("Datos cargados exitosamente.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
        End Sub

        Private Sub rEMPLACARToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs)
        End Sub

        Private Sub comboBox1_SelectedIndexChanged(ByVal sender As Object, ByVal e As EventArgs)
            Dim valorSeleccionado As Object = comboBox1.SelectedItem

            If valorSeleccionado IsNot Nothing Then
                formato = valorSeleccionado.ToString()
            End If
        End Sub

        Private Sub eDITARToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs)
            If dgvDatos.SelectedRows.Count > 0 Then
                Dim indiceSeleccionado As Integer = dgvDatos.SelectedRows(0).Index
                Dim filaSeleccionada As DataGridViewRow = dgvDatos.Rows(indiceSeleccionado)
                Dim valorCelda0 As String = filaSeleccionada.Cells(0).Value.ToString()
                Dim valorCelda1 As String = filaSeleccionada.Cells(1).Value.ToString()
                Dim valorCelda2 As String = filaSeleccionada.Cells(2).Value.ToString()
                txtNombre.Text = valorCelda0
                txtTelefono.Text = valorCelda1
                txtCorreo.Text = valorCelda2
            End If
        End Sub

        Friend WithEvents ContextMenuStrip1 As ContextMenuStrip
        Private components As IContainer

        Private Sub InitializeComponent()
            Me.components = New System.ComponentModel.Container()
            Me.ContextMenuStrip1 = New System.Windows.Forms.ContextMenuStrip(Me.components)
            Me.txtCorreo = New System.Windows.Forms.TextBox()
            Me.txtNombre = New System.Windows.Forms.TextBox()
            Me.txtTelefono = New System.Windows.Forms.TextBox()
            Me.Label1 = New System.Windows.Forms.Label()
            Me.Label2 = New System.Windows.Forms.Label()
            Me.Label3 = New System.Windows.Forms.Label()
            Me.MenuStrip1 = New System.Windows.Forms.MenuStrip()
            Me.dgvDatos = New System.Windows.Forms.DataGridView()
            Me.ComboBox1 = New System.Windows.Forms.ComboBox()
            Me.Label4 = New System.Windows.Forms.Label()
            Me.TextBox1 = New System.Windows.Forms.TextBox()
            Me.Label5 = New System.Windows.Forms.Label()
            CType(Me.dgvDatos, System.ComponentModel.ISupportInitialize).BeginInit()
            Me.SuspendLayout()
            '
            'ContextMenuStrip1
            '
            Me.ContextMenuStrip1.ImageScalingSize = New System.Drawing.Size(20, 20)
            Me.ContextMenuStrip1.Name = "ContextMenuStrip1"
            Me.ContextMenuStrip1.Size = New System.Drawing.Size(61, 4)
            '
            'txtCorreo
            '
            Me.txtCorreo.Location = New System.Drawing.Point(73, 94)
            Me.txtCorreo.Name = "txtCorreo"
            Me.txtCorreo.Size = New System.Drawing.Size(193, 22)
            Me.txtCorreo.TabIndex = 1
            '
            'txtNombre
            '
            Me.txtNombre.Location = New System.Drawing.Point(73, 38)
            Me.txtNombre.Name = "txtNombre"
            Me.txtNombre.Size = New System.Drawing.Size(193, 22)
            Me.txtNombre.TabIndex = 2
            '
            'txtTelefono
            '
            Me.txtTelefono.Location = New System.Drawing.Point(73, 66)
            Me.txtTelefono.Name = "txtTelefono"
            Me.txtTelefono.Size = New System.Drawing.Size(193, 22)
            Me.txtTelefono.TabIndex = 3
            '
            'Label1
            '
            Me.Label1.AutoSize = True
            Me.Label1.Location = New System.Drawing.Point(8, 41)
            Me.Label1.Name = "Label1"
            Me.Label1.Size = New System.Drawing.Size(59, 16)
            Me.Label1.TabIndex = 4
            Me.Label1.Text = "Nombre:"
            '
            'Label2
            '
            Me.Label2.AutoSize = True
            Me.Label2.Location = New System.Drawing.Point(3, 66)
            Me.Label2.Name = "Label2"
            Me.Label2.Size = New System.Drawing.Size(64, 16)
            Me.Label2.TabIndex = 5
            Me.Label2.Text = "Telefono:"
            '
            'Label3
            '
            Me.Label3.AutoSize = True
            Me.Label3.Location = New System.Drawing.Point(16, 97)
            Me.Label3.Name = "Label3"
            Me.Label3.Size = New System.Drawing.Size(51, 16)
            Me.Label3.TabIndex = 6
            Me.Label3.Text = "Correo:"
            '
            'MenuStrip1
            '
            Me.MenuStrip1.ImageScalingSize = New System.Drawing.Size(20, 20)
            Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
            Me.MenuStrip1.Name = "MenuStrip1"
            Me.MenuStrip1.Size = New System.Drawing.Size(800, 24)
            Me.MenuStrip1.TabIndex = 7
            Me.MenuStrip1.Text = "MenuStrip1"
            '
            'dgvDatos
            '
            Me.dgvDatos.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
            Me.dgvDatos.Location = New System.Drawing.Point(6, 122)
            Me.dgvDatos.Name = "dgvDatos"
            Me.dgvDatos.RowHeadersWidth = 51
            Me.dgvDatos.RowTemplate.Height = 24
            Me.dgvDatos.Size = New System.Drawing.Size(782, 316)
            Me.dgvDatos.TabIndex = 8
            '
            'ComboBox1
            '
            Me.ComboBox1.FormattingEnabled = True
            Me.ComboBox1.Items.AddRange(New Object() {"csv", "txt", "xml", "json"})
            Me.ComboBox1.Location = New System.Drawing.Point(388, 97)
            Me.ComboBox1.Name = "ComboBox1"
            Me.ComboBox1.Size = New System.Drawing.Size(121, 24)
            Me.ComboBox1.TabIndex = 9
            '
            'Label4
            '
            Me.Label4.AutoSize = True
            Me.Label4.Location = New System.Drawing.Point(272, 100)
            Me.Label4.Name = "Label4"
            Me.Label4.Size = New System.Drawing.Size(110, 16)
            Me.Label4.TabIndex = 10
            Me.Label4.Text = "Ruta Del Archivo:"
            '
            'TextBox1
            '
            Me.TextBox1.Location = New System.Drawing.Point(409, 62)
            Me.TextBox1.Name = "TextBox1"
            Me.TextBox1.Size = New System.Drawing.Size(100, 22)
            Me.TextBox1.TabIndex = 11
            '
            'Label5
            '
            Me.Label5.AutoSize = True
            Me.Label5.Location = New System.Drawing.Point(272, 65)
            Me.Label5.Name = "Label5"
            Me.Label5.Size = New System.Drawing.Size(131, 16)
            Me.Label5.TabIndex = 12
            Me.Label5.Text = "Nombre Del Archivo:"
            '
            'Form1
            '
            Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
            Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
            Me.ClientSize = New System.Drawing.Size(800, 450)
            Me.Controls.Add(Me.Label5)
            Me.Controls.Add(Me.TextBox1)
            Me.Controls.Add(Me.Label4)
            Me.Controls.Add(Me.ComboBox1)
            Me.Controls.Add(Me.dgvDatos)
            Me.Controls.Add(Me.Label3)
            Me.Controls.Add(Me.Label2)
            Me.Controls.Add(Me.MenuStrip1)
            Me.Controls.Add(Me.Label1)
            Me.Controls.Add(Me.txtTelefono)
            Me.Controls.Add(Me.txtNombre)
            Me.Controls.Add(Me.txtCorreo)
            Me.MainMenuStrip = Me.MenuStrip1
            Me.Name = "Form1"
            Me.Text = "Form1"
            CType(Me.dgvDatos, System.ComponentModel.ISupportInitialize).EndInit()
            Me.ResumeLayout(False)
            Me.PerformLayout()

        End Sub

        Friend WithEvents txtCorreo As Windows.Forms.TextBox
        Friend WithEvents txtNombre As Windows.Forms.TextBox
        Friend WithEvents txtTelefono As Windows.Forms.TextBox
        Friend WithEvents Label1 As Label
        Friend WithEvents Label2 As Label
        Friend WithEvents Label3 As Label
        Friend WithEvents MenuStrip1 As MenuStrip
        Friend WithEvents dgvDatos As DataGridView
        Friend WithEvents ComboBox1 As Windows.Forms.ComboBox
        Friend WithEvents Label4 As Label
        Friend WithEvents TextBox1 As Windows.Forms.TextBox
        Friend WithEvents Label5 As Label

        Public Sub New()
        End Sub
    End Class
End Namespace

