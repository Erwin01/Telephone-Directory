using System;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace DirectorioTelefonicoCsharp
{
    public partial class frmPrincipal : Form
    {
        // Establezco la conexion Base Datos/SQL SERVER
        string connectionString = @"Data Source=DESKTOP-T9QOPMI\SQLEXPRESS;Initial Catalog=DirectorioTelefonicoDB;Integrated Security=True";
        int DirectorioTelefonicoID = 0;

        public frmPrincipal()
        {
            InitializeComponent();
            // ajustar la tabla ancho y largo
            DgvDirectorioTelefonico.AutoSizeColumnsMode = DgvDirectorioTelefonico.AutoSizeColumnsMode;
            // tooltip botones
            NombreBotones();
            // botones deshabilitados
            TxtNombres.Enabled = false;
            TxtApellidos.Enabled = false;
            TxtContacto.Enabled = false;
            txtCorreoElectronico.Enabled = false;
            TxtDireccionResidencia.Enabled = false;
            TxtBuscar.Enabled = false;
            DgvDirectorioTelefonico.Enabled = false;
            BtnGuardar.Enabled = false;
            BtnLimpiar.Enabled = false;
            cmbCambioIdioma.Enabled = false;
            btnExportarExcel.Enabled = false;
            
            
        }

        #region[Etiqueta Botones]
        /// <summary>
        /// Metodo Etiqueta Botones
        /// </summary>
        public void NombreBotones()
        {
            ToolTip tool = new ToolTip();
            tool.ShowAlways = true;
            tool.SetToolTip(btnNuevo, "Nuevo");
            tool.SetToolTip(BtnGuardar, "Guardar");
            tool.SetToolTip(BtnLimpiar, "Limpiar");
            tool.SetToolTip(BtnEliminar, "Eliminar");
            tool.SetToolTip(btnExportarExcel, "Exportar Excel");
            // Labels
            tool.SetToolTip(lblNombres, "Nombres Completos");
            tool.SetToolTip(lblApellidos, "Apellidos Completos");
            tool.SetToolTip(lblNumeroContacto, "Número Contacto");
            tool.SetToolTip(lblEmail, "Correo Electrónico");
            tool.SetToolTip(lblComentario, "Comentario");
            tool.SetToolTip(lblBuscar, "Buscar");
            tool.SetToolTip(lblIdioma, "Idioma");
        }
        #endregion

        #region[Boton Guardar]
        /// <summary>
        /// Guarda los datos en la base de datos
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnGuardar_Click(object sender, System.EventArgs e)
        {


            if (TxtNombres.Text.Trim() != "" && TxtApellidos.Text.Trim() != "" && TxtContacto.Text.Trim() != "")
            {
                // Expresion regular
                Regex reg = new Regex(@"^([\w\.\-]+)@([\w\-]+)((\.(\w){2,3})+)$");
                Match match = reg.Match(txtCorreoElectronico.Text.Trim());
                if (match.Success)
                {
                    // Instancia la conexion
                    using (SqlConnection sqlCon = new SqlConnection(connectionString))
                    {
                        // Abre conexion
                        sqlCon.Open();
                        // guarda procedimiento
                        SqlCommand sqlcmd = new SqlCommand("AgregarContactoOEditar", sqlCon);
                        sqlcmd.CommandType = CommandType.StoredProcedure;
                        sqlcmd.Parameters.AddWithValue("@DirectorioTelefonicoID", DirectorioTelefonicoID);
                        sqlcmd.Parameters.AddWithValue("@Nombres", TxtNombres.Text.Trim());
                        sqlcmd.Parameters.AddWithValue("@Apellidos", TxtApellidos.Text.Trim());
                        sqlcmd.Parameters.AddWithValue("@Contacto", TxtContacto.Text.Trim());
                        sqlcmd.Parameters.AddWithValue("@CorreoElectronico", txtCorreoElectronico.Text.Trim());
                        sqlcmd.Parameters.AddWithValue("@DireccionResidencia", TxtDireccionResidencia.Text.Trim());
                        sqlcmd.ExecuteNonQuery();
                        MessageBox.Show("Dato Guardado Exitosamente", "Información Importante", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        Limpiar();
                        GridLlenar();
                        DgvDirectorioTelefonico.Refresh();
                    }
                }
                else
                    MessageBox.Show("Dirección de E-Mail no válido.", "Mensaje Importante", MessageBoxButtons.OK, MessageBoxIcon.Stop);
            }
            else
            {
                MessageBox.Show("Por favor ingrese los campos obligatorios.","Información Importante", MessageBoxButtons.OK, MessageBoxIcon.Stop);
            }
        }
        #endregion

        #region[Limpiar Cajas de Texto]
        /// <summary>
        ///  Metodo Limpiar Campos
        /// </summary>
        void Limpiar()
        {
            TxtNombres.Text = TxtApellidos.Text
            = TxtContacto.Text = txtCorreoElectronico.Text 
            = TxtDireccionResidencia.Text = TxtBuscar.Text
            = "";

            DirectorioTelefonicoID = 0;
            BtnGuardar.Text = "Guardar";
            BtnEliminar.Enabled = false;
        }
        #endregion

        #region[Boton Limpiar]
        /// <summary>
        /// Evento Limpiar Cajas de Texto
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnLimpiar_Click(object sender, System.EventArgs e)
        {

            //Limpiar();
            if (TxtNombres.Text == string.Empty && TxtApellidos.Text == string.Empty && 
                TxtContacto.Text == string.Empty && txtCorreoElectronico.Text == string.Empty &&
                TxtDireccionResidencia.Text == string.Empty)
            {
                MessageBox.Show("El ó los Campos de texto se encuentran vacíos", "Información Importante", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            else
            {
                Limpiar();
                TxtNombres.Focus();
            }
        }
        #endregion

        #region[Llenar Tabla]
        /// <summary>
        /// Metodo Para Llenar Tabla
        /// </summary>
        void GridLlenar()
        {
            using (SqlConnection sqlcon = new SqlConnection(connectionString))
            {
                sqlcon.Open();
                SqlDataAdapter sqlDa = new SqlDataAdapter("ContactoVerTodos", sqlcon);
                sqlDa.SelectCommand.CommandType = CommandType.StoredProcedure;
                DataTable DtBl = new DataTable();
                sqlDa.Fill(DtBl);
                DgvDirectorioTelefonico.DataSource = DtBl;
                PintarEstadoColumna();
                DgvDirectorioTelefonico.Refresh();
            }
        }
        #endregion

        #region[Cargar Formulario]
        private void Form1_Load(object sender, System.EventArgs e)
        {
            //GridLlenar();
            BtnEliminar.Enabled = false;
        }
        #endregion

        #region[Seleccionar Dato Tabla]
        /// <summary>
        /// Evento Doble Clic Al Seleccionar Un Dato En La Tabla
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void DgvDirectorioTelefonico_DoubleClick(object sender, System.EventArgs e)
        {
            if (DgvDirectorioTelefonico.CurrentRow.Index != -1)
            {
                TxtNombres.Text = DgvDirectorioTelefonico.CurrentRow.Cells[1].Value.ToString();
                TxtApellidos.Text = DgvDirectorioTelefonico.CurrentRow.Cells[2].Value.ToString();
                TxtContacto.Text = DgvDirectorioTelefonico.CurrentRow.Cells[3].Value.ToString();
                txtCorreoElectronico.Text = DgvDirectorioTelefonico.CurrentRow.Cells[4].Value.ToString();
                TxtDireccionResidencia.Text = DgvDirectorioTelefonico.CurrentRow.Cells[5].Value.ToString();
                DirectorioTelefonicoID = Convert.ToInt32(DgvDirectorioTelefonico.CurrentRow.Cells[0].Value.ToString());

                BtnGuardar.Text = "Actualizar";
                BtnEliminar.Enabled = true;
            }
        }
        #endregion

        #region[Boton Eliminar]
        /// <summary>
        /// Boton Eliminar
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnEliminar_Click(object sender, EventArgs e)
        {
            using (SqlConnection sqlCon = new SqlConnection(connectionString))
            {
                ;
                // Abre conexion
                sqlCon.Open();
                // guarda procedimiento
                SqlCommand sqlcmd = new SqlCommand("ContactoEliminarPoID", sqlCon);
                sqlcmd.CommandType = CommandType.StoredProcedure;
                sqlcmd.Parameters.AddWithValue("@DirectorioTelefonicoID", DirectorioTelefonicoID);
                DialogResult confirmar = MessageBox.Show("Realmente Desea Eliminar a " + TxtNombres.Text + "?", "Información Importante", MessageBoxButtons.YesNo, MessageBoxIcon.Question);            
                if (confirmar == DialogResult.Yes)
                {
                    sqlcmd.ExecuteNonQuery();
                    MessageBox.Show("Dato Eliminado", "Información Importante", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                }
                else if(confirmar == DialogResult.No)
                {
                   
                }
                
                
               
                
                Limpiar();
                GridLlenar();
            }
        }
        #endregion

        #region[Buscar Dato]
        /// <summary>
        /// Buscar Dato En La Tabla Al Digitar Cualquier Letra
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void TxtBuscar_TextChanged(object sender, EventArgs e)
        {
            using (SqlConnection sqlcon = new SqlConnection(connectionString))
            {
                sqlcon.Open();
                SqlDataAdapter sqlDa = new SqlDataAdapter("ContactoBuscarPorValor", sqlcon);
                sqlDa.SelectCommand.CommandType = CommandType.StoredProcedure;
                sqlDa.SelectCommand.Parameters.AddWithValue("@BuscarValor", TxtBuscar.Text.Trim());
                DataTable DtBl = new DataTable();
                sqlDa.Fill(DtBl);
                DgvDirectorioTelefonico.DataSource = DtBl;
            }
        }
        #endregion

        #region[Boton Nuevo/Habilitar Campos]
        /// <summary>
        /// Muestra Los Campos Habilitados Al Presionar Clic En El Boton 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnNuevo_Click(object sender, EventArgs e)
        {
            GridLlenar();
            TxtNombres.Enabled = true;
            TxtApellidos.Enabled = true;
            TxtContacto.Enabled = true;
            txtCorreoElectronico.Enabled = true;
            TxtDireccionResidencia.Enabled = true;
            TxtBuscar.Enabled = true;
            DgvDirectorioTelefonico.Enabled = true;
            BtnGuardar.Enabled = true;
            BtnLimpiar.Enabled = true;
            cmbCambioIdioma.Enabled = true;
        }
        #endregion

        #region[Cambiar Idioma]
        private void CambiarEstadoIdioma_TextChanged(object sender, EventArgs e)
        {
            string comboBox = cmbCambioIdioma.SelectedItem.ToString();

            if (comboBox.Equals("English"))
            {
                // Etiquetas
                lblNombres.Text = "Username:";
                lblApellidos.Text = "Last Name:";
                lblNumeroContacto.Text = "Contact Number:";
                lblEmail.Text = "Email:";
                lblComentario.Text = "Commentary:";
                lblBuscar.Text = "Search:";
                lblIdioma.Text = "Language:";
                // Botones
                btnNuevo.Text = "New (+)";
                BtnGuardar.Text = "Save";
                BtnEliminar.Text = "Delete";
                BtnLimpiar.Text = "Clear";
                btnExportarExcel.Text = "Export to Excel";
            }
            else
            {

                // Etiquetas
                lblNombres.Text = "Nombres Completos:";
                lblApellidos.Text = "Apellidos Completos:";
                lblNumeroContacto.Text = "Número Contacto:";
                lblEmail.Text = "Correo Electrónico:";
                lblComentario.Text = "Comentario:";
                lblBuscar.Text = "Buscar:";
                lblIdioma.Text = "Idioma";
                // Botones
                btnNuevo.Text = "Nuevo (+)";
                BtnGuardar.Text = "Guardar";
                BtnEliminar.Text = "Eliminar";
                BtnLimpiar.Text = "Limpiar";
                btnExportarExcel.Text = "Exportar Excel";
            }
        }
        #endregion

        #region[Pintar Estados de las Columnas]
        public void PintarEstadoColumna()
        {
            // Metodo para recorrer todas las columnas de la tabla
            for (int i = 0; i < DgvDirectorioTelefonico.Rows.Count; i++)
            {
                string valor = DgvDirectorioTelefonico.Rows[i].Cells[5].Value.ToString();

                if (valor == "PENDIENTE POR LLAMAR")
                {
                    DgvDirectorioTelefonico.Rows[i].DefaultCellStyle.BackColor = Color.AntiqueWhite;
                }
                else if(valor == "")
                {
                    DgvDirectorioTelefonico.Rows[i].DefaultCellStyle.BackColor = Color.White;
                }
                else if (valor == "NO ASISTIRA AL EVENTO")
                {
                    DgvDirectorioTelefonico.Rows[i].DefaultCellStyle.BackColor = Color.Red;
                }
                else if (valor == "VOLVER A LLAMAR OTRA VEZ")
                {
                    DgvDirectorioTelefonico.Rows[i].DefaultCellStyle.BackColor = Color.Yellow;
                }
                else if (valor == "ENVIAR DOCUMENTACION")
                {
                    DgvDirectorioTelefonico.Rows[i].DefaultCellStyle.BackColor = Color.OrangeRed;
                }
                else if (valor == "NO CONTESTA")
                {
                    DgvDirectorioTelefonico.Rows[i].DefaultCellStyle.BackColor = Color.Cyan;
                }
                else
                {
                    valor = "ASISTIRA AL EVENTO";
                    DgvDirectorioTelefonico.Rows[i].DefaultCellStyle.BackColor = Color.LimeGreen;
                }
            }
        }
        #endregion


        // Creo el Metodo para Exportar desde el DataGridview a Hoja Excel
        private void BtnExportarExcel_Click(object sender, EventArgs e)
        {
            //SaveFileDialog sfd = new SaveFileDialog();
            //sfd.Filter = "Excel Document (*.xls)|*.xls";
            //sfd.FileName = "export.xls";

            //if (sfd.ShowDialog() == DialogResult.OK)
            //{
            //    ToExcel(DgvDirectorioTelefonico, sfd.FileName);
            //    MessageBox.Show("Archivo Creado");

            //}
        }

        #region[Cerrar y Salir del Formulario]
        private void frmPrincipal_FormClosing(object sender, FormClosingEventArgs e)
        {
            DialogResult confirmar = MessageBox.Show("Desea Salir de la Aplicación?","Salir ", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation);
            if (confirmar == DialogResult.Yes)
            {
                e.Cancel = false;
            }
            else if(confirmar == DialogResult.No)
            {
                e.Cancel = true;
            }
        }
        #endregion
    }
}
