namespace Es21_WordCSharpù
{
    partial class frmWordCSharp
    {
        /// <summary>
        /// Variabile di progettazione necessaria.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Pulire le risorse in uso.
        /// </summary>
        /// <param name="disposing">ha valore true se le risorse gestite devono essere eliminate, false in caso contrario.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Codice generato da Progettazione Windows Form

        /// <summary>
        /// Metodo necessario per il supporto della finestra di progettazione. Non modificare
        /// il contenuto del metodo con l'editor di codice.
        /// </summary>
        private void InitializeComponent()
        {
            this.btnCrea = new System.Windows.Forms.Button();
            this.btnSalvaChiudi = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.btnInserisci = new System.Windows.Forms.Button();
            this.txtTesto = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // btnCrea
            // 
            this.btnCrea.Location = new System.Drawing.Point(12, 12);
            this.btnCrea.Name = "btnCrea";
            this.btnCrea.Size = new System.Drawing.Size(118, 36);
            this.btnCrea.TabIndex = 0;
            this.btnCrea.Text = "CREA DOCUMENTO WORD";
            this.btnCrea.UseVisualStyleBackColor = true;
            this.btnCrea.Click += new System.EventHandler(this.btnCrea_Click);
            // 
            // btnSalvaChiudi
            // 
            this.btnSalvaChiudi.Location = new System.Drawing.Point(136, 12);
            this.btnSalvaChiudi.Name = "btnSalvaChiudi";
            this.btnSalvaChiudi.Size = new System.Drawing.Size(118, 36);
            this.btnSalvaChiudi.TabIndex = 1;
            this.btnSalvaChiudi.Text = "SALVA E CHIUDI WORD";
            this.btnSalvaChiudi.UseVisualStyleBackColor = true;
            this.btnSalvaChiudi.Click += new System.EventHandler(this.btnSalvaChiudi_Click);
            // 
            // label1
            // 
            this.label1.Location = new System.Drawing.Point(9, 91);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(100, 23);
            this.label1.TabIndex = 2;
            this.label1.Text = "label1";
            // 
            // btnInserisci
            // 
            this.btnInserisci.Location = new System.Drawing.Point(212, 91);
            this.btnInserisci.Name = "btnInserisci";
            this.btnInserisci.Size = new System.Drawing.Size(75, 23);
            this.btnInserisci.TabIndex = 3;
            this.btnInserisci.Text = "INSERISCI";
            this.btnInserisci.UseVisualStyleBackColor = true;
            this.btnInserisci.Click += new System.EventHandler(this.btnInserisci_Click);
            // 
            // txtTesto
            // 
            this.txtTesto.Location = new System.Drawing.Point(106, 94);
            this.txtTesto.Name = "txtTesto";
            this.txtTesto.Size = new System.Drawing.Size(100, 20);
            this.txtTesto.TabIndex = 4;
            // 
            // frmWordCSharp
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.txtTesto);
            this.Controls.Add(this.btnInserisci);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.btnSalvaChiudi);
            this.Controls.Add(this.btnCrea);
            this.Name = "frmWordCSharp";
            this.Text = "Gestione Word da C#";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnCrea;
        private System.Windows.Forms.Button btnSalvaChiudi;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btnInserisci;
        private System.Windows.Forms.TextBox txtTesto;
    }
}

