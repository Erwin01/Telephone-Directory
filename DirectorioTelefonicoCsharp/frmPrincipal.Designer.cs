namespace DirectorioTelefonicoCsharp
{
    partial class frmPrincipal
    {
        /// <summary>
        /// Variable del diseñador necesaria.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Limpiar los recursos que se estén usando.
        /// </summary>
        /// <param name="disposing">true si los recursos administrados se deben desechar; false en caso contrario.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Código generado por el Diseñador de Windows Forms

        /// <summary>
        /// Método necesario para admitir el Diseñador. No se puede modificar
        /// el contenido de este método con el editor de código.
        /// </summary>
        private void InitializeComponent()
        {
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmPrincipal));
            this.lblNombres = new System.Windows.Forms.Label();
            this.lblApellidos = new System.Windows.Forms.Label();
            this.lblNumeroContacto = new System.Windows.Forms.Label();
            this.lblEmail = new System.Windows.Forms.Label();
            this.lblComentario = new System.Windows.Forms.Label();
            this.TxtApellidos = new System.Windows.Forms.TextBox();
            this.TxtContacto = new System.Windows.Forms.TextBox();
            this.txtCorreoElectronico = new System.Windows.Forms.TextBox();
            this.TxtNombres = new System.Windows.Forms.TextBox();
            this.lblBuscar = new System.Windows.Forms.Label();
            this.DgvDirectorioTelefonico = new System.Windows.Forms.DataGridView();
            this.ID = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.NombresCompletos = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ApellidosCompletos = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Contactos = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.CorreoElectronico = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.DireccionResidencia = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.TxtDireccionResidencia = new System.Windows.Forms.TextBox();
            this.TxtBuscar = new System.Windows.Forms.TextBox();
            this.BtnLimpiar = new System.Windows.Forms.Button();
            this.BtnEliminar = new System.Windows.Forms.Button();
            this.BtnGuardar = new System.Windows.Forms.Button();
            this.label4 = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            this.label9 = new System.Windows.Forms.Label();
            this.btnNuevo = new System.Windows.Forms.Button();
            this.cmbCambioIdioma = new System.Windows.Forms.ComboBox();
            this.lblIdioma = new System.Windows.Forms.Label();
            this.btnExportarExcel = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.DgvDirectorioTelefonico)).BeginInit();
            this.SuspendLayout();
            // 
            // lblNombres
            // 
            this.lblNombres.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.lblNombres.AutoSize = true;
            this.lblNombres.Font = new System.Drawing.Font("Segoe UI", 11.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblNombres.Location = new System.Drawing.Point(13, 113);
            this.lblNombres.Name = "lblNombres";
            this.lblNombres.Size = new System.Drawing.Size(157, 20);
            this.lblNombres.TabIndex = 0;
            this.lblNombres.Text = "Nombres Completos:";
            // 
            // lblApellidos
            // 
            this.lblApellidos.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.lblApellidos.AutoSize = true;
            this.lblApellidos.Font = new System.Drawing.Font("Segoe UI", 11.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblApellidos.Location = new System.Drawing.Point(13, 148);
            this.lblApellidos.Name = "lblApellidos";
            this.lblApellidos.Size = new System.Drawing.Size(157, 20);
            this.lblApellidos.TabIndex = 0;
            this.lblApellidos.Text = "Apellidos Completos:";
            // 
            // lblNumeroContacto
            // 
            this.lblNumeroContacto.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.lblNumeroContacto.AutoSize = true;
            this.lblNumeroContacto.Font = new System.Drawing.Font("Segoe UI", 11.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblNumeroContacto.Location = new System.Drawing.Point(13, 186);
            this.lblNumeroContacto.Name = "lblNumeroContacto";
            this.lblNumeroContacto.Size = new System.Drawing.Size(138, 20);
            this.lblNumeroContacto.TabIndex = 0;
            this.lblNumeroContacto.Text = "Número Contacto:";
            // 
            // lblEmail
            // 
            this.lblEmail.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.lblEmail.AutoSize = true;
            this.lblEmail.Font = new System.Drawing.Font("Segoe UI", 11.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblEmail.Location = new System.Drawing.Point(13, 218);
            this.lblEmail.Name = "lblEmail";
            this.lblEmail.Size = new System.Drawing.Size(141, 20);
            this.lblEmail.TabIndex = 0;
            this.lblEmail.Text = "Correo Electrónico:";
            // 
            // lblComentario
            // 
            this.lblComentario.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.lblComentario.AutoSize = true;
            this.lblComentario.Font = new System.Drawing.Font("Segoe UI", 11.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblComentario.Location = new System.Drawing.Point(13, 252);
            this.lblComentario.Name = "lblComentario";
            this.lblComentario.Size = new System.Drawing.Size(95, 20);
            this.lblComentario.TabIndex = 0;
            this.lblComentario.Text = "Comentario:";
            // 
            // TxtApellidos
            // 
            this.TxtApellidos.Font = new System.Drawing.Font("Segoe UI", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.TxtApellidos.Location = new System.Drawing.Point(178, 145);
            this.TxtApellidos.Name = "TxtApellidos";
            this.TxtApellidos.Size = new System.Drawing.Size(207, 29);
            this.TxtApellidos.TabIndex = 1;
            // 
            // TxtContacto
            // 
            this.TxtContacto.Font = new System.Drawing.Font("Segoe UI", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.TxtContacto.Location = new System.Drawing.Point(178, 180);
            this.TxtContacto.Name = "TxtContacto";
            this.TxtContacto.Size = new System.Drawing.Size(207, 29);
            this.TxtContacto.TabIndex = 2;
            // 
            // txtCorreoElectronico
            // 
            this.txtCorreoElectronico.Font = new System.Drawing.Font("Segoe UI", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtCorreoElectronico.Location = new System.Drawing.Point(178, 215);
            this.txtCorreoElectronico.Name = "txtCorreoElectronico";
            this.txtCorreoElectronico.Size = new System.Drawing.Size(207, 29);
            this.txtCorreoElectronico.TabIndex = 4;
            // 
            // TxtNombres
            // 
            this.TxtNombres.Font = new System.Drawing.Font("Segoe UI", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.TxtNombres.Location = new System.Drawing.Point(178, 110);
            this.TxtNombres.Name = "TxtNombres";
            this.TxtNombres.Size = new System.Drawing.Size(207, 29);
            this.TxtNombres.TabIndex = 0;
            this.TxtNombres.TabStop = false;
            // 
            // lblBuscar
            // 
            this.lblBuscar.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.lblBuscar.AutoSize = true;
            this.lblBuscar.Font = new System.Drawing.Font("Segoe UI", 11.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblBuscar.Location = new System.Drawing.Point(625, 64);
            this.lblBuscar.Name = "lblBuscar";
            this.lblBuscar.Size = new System.Drawing.Size(61, 20);
            this.lblBuscar.TabIndex = 0;
            this.lblBuscar.Text = "Buscar:";
            // 
            // DgvDirectorioTelefonico
            // 
            this.DgvDirectorioTelefonico.AllowUserToAddRows = false;
            this.DgvDirectorioTelefonico.AllowUserToDeleteRows = false;
            this.DgvDirectorioTelefonico.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.DgvDirectorioTelefonico.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.DgvDirectorioTelefonico.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells;
            this.DgvDirectorioTelefonico.BackgroundColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle1.Font = new System.Drawing.Font("Verdana", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.DgvDirectorioTelefonico.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle1;
            this.DgvDirectorioTelefonico.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.DgvDirectorioTelefonico.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.ID,
            this.NombresCompletos,
            this.ApellidosCompletos,
            this.Contactos,
            this.CorreoElectronico,
            this.DireccionResidencia});
            this.DgvDirectorioTelefonico.Cursor = System.Windows.Forms.Cursors.Hand;
            this.DgvDirectorioTelefonico.Location = new System.Drawing.Point(415, 110);
            this.DgvDirectorioTelefonico.Name = "DgvDirectorioTelefonico";
            this.DgvDirectorioTelefonico.ReadOnly = true;
            this.DgvDirectorioTelefonico.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.DgvDirectorioTelefonico.Size = new System.Drawing.Size(656, 244);
            this.DgvDirectorioTelefonico.TabIndex = 10;
            this.DgvDirectorioTelefonico.DoubleClick += new System.EventHandler(this.DgvDirectorioTelefonico_DoubleClick);
            // 
            // ID
            // 
            this.ID.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.ColumnHeader;
            this.ID.DataPropertyName = "DirectorioTelefonicoID";
            this.ID.HeaderText = "ID";
            this.ID.Name = "ID";
            this.ID.ReadOnly = true;
            this.ID.Width = 47;
            // 
            // NombresCompletos
            // 
            this.NombresCompletos.DataPropertyName = "Nombres";
            this.NombresCompletos.HeaderText = "Nombres Completos";
            this.NombresCompletos.Name = "NombresCompletos";
            this.NombresCompletos.ReadOnly = true;
            // 
            // ApellidosCompletos
            // 
            this.ApellidosCompletos.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.ApellidosCompletos.DataPropertyName = "Apellidos";
            this.ApellidosCompletos.FillWeight = 87.92079F;
            this.ApellidosCompletos.HeaderText = "Apellidos Completos";
            this.ApellidosCompletos.Name = "ApellidosCompletos";
            this.ApellidosCompletos.ReadOnly = true;
            // 
            // Contactos
            // 
            this.Contactos.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.Contactos.DataPropertyName = "Contacto";
            this.Contactos.FillWeight = 96.68622F;
            this.Contactos.HeaderText = "Teléfono";
            this.Contactos.Name = "Contactos";
            this.Contactos.ReadOnly = true;
            // 
            // CorreoElectronico
            // 
            this.CorreoElectronico.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.CorreoElectronico.DataPropertyName = "CorreoElectronico";
            this.CorreoElectronico.FillWeight = 105.2829F;
            this.CorreoElectronico.HeaderText = "E-Mail";
            this.CorreoElectronico.Name = "CorreoElectronico";
            this.CorreoElectronico.ReadOnly = true;
            // 
            // DireccionResidencia
            // 
            this.DireccionResidencia.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.DireccionResidencia.DataPropertyName = "DireccionResidencia";
            this.DireccionResidencia.FillWeight = 110.1101F;
            this.DireccionResidencia.HeaderText = "Comentario";
            this.DireccionResidencia.Name = "DireccionResidencia";
            this.DireccionResidencia.ReadOnly = true;
            // 
            // TxtDireccionResidencia
            // 
            this.TxtDireccionResidencia.Font = new System.Drawing.Font("Segoe UI", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.TxtDireccionResidencia.Location = new System.Drawing.Point(178, 250);
            this.TxtDireccionResidencia.Multiline = true;
            this.TxtDireccionResidencia.Name = "TxtDireccionResidencia";
            this.TxtDireccionResidencia.Size = new System.Drawing.Size(207, 104);
            this.TxtDireccionResidencia.TabIndex = 5;
            // 
            // TxtBuscar
            // 
            this.TxtBuscar.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.TxtBuscar.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.TxtBuscar.Location = new System.Drawing.Point(702, 63);
            this.TxtBuscar.Name = "TxtBuscar";
            this.TxtBuscar.Size = new System.Drawing.Size(205, 25);
            this.TxtBuscar.TabIndex = 9;
            this.TxtBuscar.TextChanged += new System.EventHandler(this.TxtBuscar_TextChanged);
            // 
            // BtnLimpiar
            // 
            this.BtnLimpiar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.BtnLimpiar.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.BtnLimpiar.Cursor = System.Windows.Forms.Cursors.Hand;
            this.BtnLimpiar.Font = new System.Drawing.Font("Segoe UI", 11.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.BtnLimpiar.ForeColor = System.Drawing.SystemColors.WindowText;
            this.BtnLimpiar.Image = global::DirectorioTelefonicoCsharp.Properties.Resources.Clear_37294;
            this.BtnLimpiar.Location = new System.Drawing.Point(277, 395);
            this.BtnLimpiar.Name = "BtnLimpiar";
            this.BtnLimpiar.Size = new System.Drawing.Size(83, 36);
            this.BtnLimpiar.TabIndex = 8;
            this.BtnLimpiar.Text = "Limpiar";
            this.BtnLimpiar.TextAlign = System.Drawing.ContentAlignment.BottomCenter;
            this.BtnLimpiar.UseVisualStyleBackColor = true;
            this.BtnLimpiar.Click += new System.EventHandler(this.BtnLimpiar_Click);
            // 
            // BtnEliminar
            // 
            this.BtnEliminar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.BtnEliminar.Cursor = System.Windows.Forms.Cursors.Hand;
            this.BtnEliminar.Font = new System.Drawing.Font("Segoe UI", 11.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.BtnEliminar.ForeColor = System.Drawing.SystemColors.WindowText;
            this.BtnEliminar.Image = global::DirectorioTelefonicoCsharp.Properties.Resources.delete_delete_exit_1577;
            this.BtnEliminar.Location = new System.Drawing.Point(164, 395);
            this.BtnEliminar.Name = "BtnEliminar";
            this.BtnEliminar.Size = new System.Drawing.Size(82, 36);
            this.BtnEliminar.TabIndex = 7;
            this.BtnEliminar.Text = "Eliminar";
            this.BtnEliminar.TextAlign = System.Drawing.ContentAlignment.BottomCenter;
            this.BtnEliminar.UseVisualStyleBackColor = true;
            this.BtnEliminar.Click += new System.EventHandler(this.BtnEliminar_Click);
            // 
            // BtnGuardar
            // 
            this.BtnGuardar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.BtnGuardar.Cursor = System.Windows.Forms.Cursors.Hand;
            this.BtnGuardar.Font = new System.Drawing.Font("Segoe UI", 11.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.BtnGuardar.ForeColor = System.Drawing.SystemColors.WindowText;
            this.BtnGuardar.Image = global::DirectorioTelefonicoCsharp.Properties.Resources.clean_well_accept_6239;
            this.BtnGuardar.Location = new System.Drawing.Point(43, 395);
            this.BtnGuardar.Name = "BtnGuardar";
            this.BtnGuardar.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.BtnGuardar.Size = new System.Drawing.Size(94, 36);
            this.BtnGuardar.TabIndex = 6;
            this.BtnGuardar.Text = "Guardar";
            this.BtnGuardar.TextAlign = System.Drawing.ContentAlignment.BottomCenter;
            this.BtnGuardar.UseVisualStyleBackColor = true;
            this.BtnGuardar.Click += new System.EventHandler(this.BtnGuardar_Click);
            // 
            // label4
            // 
            this.label4.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.ForeColor = System.Drawing.Color.Red;
            this.label4.Location = new System.Drawing.Point(391, 115);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(18, 24);
            this.label4.TabIndex = 0;
            this.label4.Text = "*";
            // 
            // label8
            // 
            this.label8.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.label8.AutoSize = true;
            this.label8.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label8.ForeColor = System.Drawing.Color.Red;
            this.label8.Location = new System.Drawing.Point(391, 150);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(18, 24);
            this.label8.TabIndex = 0;
            this.label8.Text = "*";
            // 
            // label9
            // 
            this.label9.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.label9.AutoSize = true;
            this.label9.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label9.ForeColor = System.Drawing.Color.Red;
            this.label9.Location = new System.Drawing.Point(391, 184);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(18, 24);
            this.label9.TabIndex = 0;
            this.label9.Text = "*";
            // 
            // btnNuevo
            // 
            this.btnNuevo.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btnNuevo.Font = new System.Drawing.Font("Segoe UI", 11.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnNuevo.ForeColor = System.Drawing.SystemColors.WindowText;
            this.btnNuevo.Location = new System.Drawing.Point(43, 27);
            this.btnNuevo.Name = "btnNuevo";
            this.btnNuevo.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.btnNuevo.Size = new System.Drawing.Size(105, 33);
            this.btnNuevo.TabIndex = 11;
            this.btnNuevo.Text = "Nuevo (+)";
            this.btnNuevo.TextAlign = System.Drawing.ContentAlignment.BottomCenter;
            this.btnNuevo.UseVisualStyleBackColor = true;
            this.btnNuevo.Click += new System.EventHandler(this.btnNuevo_Click);
            // 
            // cmbCambioIdioma
            // 
            this.cmbCambioIdioma.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.cmbCambioIdioma.Cursor = System.Windows.Forms.Cursors.Hand;
            this.cmbCambioIdioma.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.cmbCambioIdioma.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cmbCambioIdioma.Items.AddRange(new object[] {
            "English",
            "Español"});
            this.cmbCambioIdioma.Location = new System.Drawing.Point(922, 12);
            this.cmbCambioIdioma.Name = "cmbCambioIdioma";
            this.cmbCambioIdioma.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.cmbCambioIdioma.Size = new System.Drawing.Size(145, 23);
            this.cmbCambioIdioma.TabIndex = 12;
            this.cmbCambioIdioma.Text = "Seleccione";
            this.cmbCambioIdioma.TextChanged += new System.EventHandler(this.CambiarEstadoIdioma_TextChanged);
            // 
            // lblIdioma
            // 
            this.lblIdioma.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.lblIdioma.AutoSize = true;
            this.lblIdioma.Font = new System.Drawing.Font("Segoe UI", 11.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblIdioma.Location = new System.Drawing.Point(842, 12);
            this.lblIdioma.Name = "lblIdioma";
            this.lblIdioma.Size = new System.Drawing.Size(62, 20);
            this.lblIdioma.TabIndex = 13;
            this.lblIdioma.Text = "Idioma:";
            // 
            // btnExportarExcel
            // 
            this.btnExportarExcel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnExportarExcel.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btnExportarExcel.Font = new System.Drawing.Font("Segoe UI", 11.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnExportarExcel.ForeColor = System.Drawing.SystemColors.WindowText;
            this.btnExportarExcel.Location = new System.Drawing.Point(599, 398);
            this.btnExportarExcel.Name = "btnExportarExcel";
            this.btnExportarExcel.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.btnExportarExcel.Size = new System.Drawing.Size(178, 33);
            this.btnExportarExcel.TabIndex = 14;
            this.btnExportarExcel.Text = "Exportar  Excel";
            this.btnExportarExcel.TextAlign = System.Drawing.ContentAlignment.BottomCenter;
            this.btnExportarExcel.UseVisualStyleBackColor = true;
            this.btnExportarExcel.Click += new System.EventHandler(this.BtnExportarExcel_Click);
            // 
            // frmPrincipal
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoSize = true;
            this.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.ClientSize = new System.Drawing.Size(1079, 447);
            this.Controls.Add(this.btnExportarExcel);
            this.Controls.Add(this.lblIdioma);
            this.Controls.Add(this.cmbCambioIdioma);
            this.Controls.Add(this.btnNuevo);
            this.Controls.Add(this.label9);
            this.Controls.Add(this.label8);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.TxtBuscar);
            this.Controls.Add(this.TxtDireccionResidencia);
            this.Controls.Add(this.DgvDirectorioTelefonico);
            this.Controls.Add(this.lblBuscar);
            this.Controls.Add(this.BtnLimpiar);
            this.Controls.Add(this.BtnEliminar);
            this.Controls.Add(this.BtnGuardar);
            this.Controls.Add(this.TxtNombres);
            this.Controls.Add(this.txtCorreoElectronico);
            this.Controls.Add(this.TxtContacto);
            this.Controls.Add(this.TxtApellidos);
            this.Controls.Add(this.lblComentario);
            this.Controls.Add(this.lblEmail);
            this.Controls.Add(this.lblNumeroContacto);
            this.Controls.Add(this.lblApellidos);
            this.Controls.Add(this.lblNombres);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "frmPrincipal";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "MIS CONTACTOS TELÉFONICOS";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.frmPrincipal_FormClosing);
            this.Load += new System.EventHandler(this.Form1_Load);
            ((System.ComponentModel.ISupportInitialize)(this.DgvDirectorioTelefonico)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label lblNombres;
        private System.Windows.Forms.Label lblApellidos;
        private System.Windows.Forms.Label lblNumeroContacto;
        private System.Windows.Forms.Label lblEmail;
        private System.Windows.Forms.Label lblComentario;
        private System.Windows.Forms.TextBox TxtApellidos;
        private System.Windows.Forms.TextBox TxtContacto;
        private System.Windows.Forms.TextBox txtCorreoElectronico;
        private System.Windows.Forms.TextBox TxtNombres;
        private System.Windows.Forms.Button BtnGuardar;
        private System.Windows.Forms.Button BtnEliminar;
        private System.Windows.Forms.Button BtnLimpiar;
        private System.Windows.Forms.Label lblBuscar;
        private System.Windows.Forms.DataGridView DgvDirectorioTelefonico;
        private System.Windows.Forms.TextBox TxtDireccionResidencia;
        private System.Windows.Forms.TextBox TxtBuscar;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.Button btnNuevo;
        private System.Windows.Forms.ComboBox cmbCambioIdioma;
        private System.Windows.Forms.Label lblIdioma;
        private System.Windows.Forms.Button btnExportarExcel;
        private System.Windows.Forms.DataGridViewTextBoxColumn ID;
        private System.Windows.Forms.DataGridViewTextBoxColumn NombresCompletos;
        private System.Windows.Forms.DataGridViewTextBoxColumn ApellidosCompletos;
        private System.Windows.Forms.DataGridViewTextBoxColumn Contactos;
        private System.Windows.Forms.DataGridViewTextBoxColumn CorreoElectronico;
        private System.Windows.Forms.DataGridViewTextBoxColumn DireccionResidencia;
    }
}

