namespace RenombramientoIcfes
{
    partial class Form1
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
            this.Digitalizadas = new WinFormsControlLibrary.usrSeleccionarCarpeta();
            this.Renombrar = new System.Windows.Forms.Button();
            this.ArchivoCargado = new WinFormsControlLibrary.usrSeleccionarArchivo();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.SuspendLayout();
            // 
            // Digitalizadas
            // 
            this.Digitalizadas.BackColor = System.Drawing.SystemColors.Control;
            this.Digitalizadas.CarpetaSeleccionado = "";
            this.Digitalizadas.ColorBoton = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(255)))), ((int)(((byte)(128)))));
            this.Digitalizadas.Location = new System.Drawing.Point(114, 53);
            this.Digitalizadas.Name = "Digitalizadas";
            this.Digitalizadas.SelectedPath = null;
            this.Digitalizadas.Size = new System.Drawing.Size(458, 34);
            this.Digitalizadas.TabIndex = 8;
            this.Digitalizadas.Titulo = "Seleccionar Ruta Hojas a Renombrar:";
            this.Digitalizadas.OnSeleccionCarpeta += new System.EventHandler<WinFormsControlLibrary.usrSeleccionarCarpeta.SeleccionCarpetaEventArgs>(this.Digitalizadas_OnSeleccionCarpeta);
            // 
            // Renombrar
            // 
            this.Renombrar.BackColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.Renombrar.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Renombrar.ForeColor = System.Drawing.SystemColors.ButtonHighlight;
            this.Renombrar.Location = new System.Drawing.Point(331, 189);
            this.Renombrar.Name = "Renombrar";
            this.Renombrar.Size = new System.Drawing.Size(177, 28);
            this.Renombrar.TabIndex = 6;
            this.Renombrar.Text = "Renombrar";
            this.Renombrar.UseVisualStyleBackColor = false;
            this.Renombrar.Click += new System.EventHandler(this.button1_Click);
            // 
            // ArchivoCargado
            // 
            this.ArchivoCargado.ArchivoSeleccionado = "";
            this.ArchivoCargado.ColorBoton = System.Drawing.SystemColors.ActiveCaption;
            this.ArchivoCargado.ColorTextoBoton = System.Drawing.Color.Black;
            this.ArchivoCargado.FiltroExtensionPermitida = "listado asistencia|*.xlsx";
            this.ArchivoCargado.InitialDirectory = null;
            this.ArchivoCargado.IsEnabled = true;
            this.ArchivoCargado.Location = new System.Drawing.Point(114, 101);
            this.ArchivoCargado.Name = "ArchivoCargado";
            this.ArchivoCargado.Size = new System.Drawing.Size(506, 34);
            this.ArchivoCargado.TabIndex = 7;
            this.ArchivoCargado.Titulo = "Cargue archivo enviado por el cliente:";
            this.ArchivoCargado.OnSeleccionArchivo += new System.EventHandler<WinFormsControlLibrary.usrSeleccionarArchivo.SeleccionArchivoEventArgs>(this.ArchivoCargado_OnSeleccionArchivo);
            // 
            // pictureBox1
            // 
            this.pictureBox1.Image = global::RenombramientoIcfes.Properties.Resources.LogoCadena;
            this.pictureBox1.Location = new System.Drawing.Point(657, 20);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(178, 103);
            this.pictureBox1.TabIndex = 10;
            this.pictureBox1.TabStop = false;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.ButtonHighlight;
            this.ClientSize = new System.Drawing.Size(856, 229);
            this.Controls.Add(this.pictureBox1);
            this.Controls.Add(this.Digitalizadas);
            this.Controls.Add(this.Renombrar);
            this.Controls.Add(this.ArchivoCargado);
            this.Name = "Form1";
            this.Text = "Form1";
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.PictureBox pictureBox1;
        private WinFormsControlLibrary.usrSeleccionarCarpeta Digitalizadas;
        private System.Windows.Forms.Button Renombrar;
        private WinFormsControlLibrary.usrSeleccionarArchivo ArchivoCargado;
    }
}

