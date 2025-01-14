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
            this.grbWordCreated = new System.Windows.Forms.GroupBox();
            this.btnCreaTabella = new System.Windows.Forms.Button();
            this.cmbRighe = new System.Windows.Forms.ComboBox();
            this.cmbColonne = new System.Windows.Forms.ComboBox();
            this.chkItalic = new System.Windows.Forms.CheckBox();
            this.chkBold = new System.Windows.Forms.CheckBox();
            this.cmbColor = new System.Windows.Forms.ComboBox();
            this.cmbAlignment = new System.Windows.Forms.ComboBox();
            this.cmbUnderlined = new System.Windows.Forms.ComboBox();
            this.cmbFont = new System.Windows.Forms.ComboBox();
            this.cmbSize = new System.Windows.Forms.ComboBox();
            this.txtTesto = new System.Windows.Forms.TextBox();
            this.btnInserisci = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.btnSalvaChiudi = new System.Windows.Forms.Button();
            this.txtDa = new System.Windows.Forms.TextBox();
            this.btnSelezionaTesto = new System.Windows.Forms.Button();
            this.txtA = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.txtTestoSostituire = new System.Windows.Forms.TextBox();
            this.btnRicerca = new System.Windows.Forms.Button();
            this.txtTestoRicercare = new System.Windows.Forms.TextBox();
            this.chkSostituisci = new System.Windows.Forms.CheckBox();
            this.btnCreaPDF = new System.Windows.Forms.Button();
            this.btnStampa = new System.Windows.Forms.Button();
            this.grbWordCreated.SuspendLayout();
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
            // grbWordCreated
            // 
            this.grbWordCreated.Controls.Add(this.btnStampa);
            this.grbWordCreated.Controls.Add(this.btnCreaPDF);
            this.grbWordCreated.Controls.Add(this.btnCreaTabella);
            this.grbWordCreated.Controls.Add(this.cmbRighe);
            this.grbWordCreated.Controls.Add(this.cmbColonne);
            this.grbWordCreated.Controls.Add(this.chkItalic);
            this.grbWordCreated.Controls.Add(this.chkBold);
            this.grbWordCreated.Controls.Add(this.cmbColor);
            this.grbWordCreated.Controls.Add(this.cmbAlignment);
            this.grbWordCreated.Controls.Add(this.cmbUnderlined);
            this.grbWordCreated.Controls.Add(this.cmbFont);
            this.grbWordCreated.Controls.Add(this.cmbSize);
            this.grbWordCreated.Controls.Add(this.txtTesto);
            this.grbWordCreated.Controls.Add(this.btnInserisci);
            this.grbWordCreated.Controls.Add(this.label1);
            this.grbWordCreated.Controls.Add(this.btnSalvaChiudi);
            this.grbWordCreated.Enabled = false;
            this.grbWordCreated.Location = new System.Drawing.Point(12, 73);
            this.grbWordCreated.Name = "grbWordCreated";
            this.grbWordCreated.Size = new System.Drawing.Size(736, 316);
            this.grbWordCreated.TabIndex = 1;
            this.grbWordCreated.TabStop = false;
            this.grbWordCreated.Text = "CONTROLLI";
            // 
            // btnCreaTabella
            // 
            this.btnCreaTabella.Location = new System.Drawing.Point(344, 88);
            this.btnCreaTabella.Name = "btnCreaTabella";
            this.btnCreaTabella.Size = new System.Drawing.Size(76, 21);
            this.btnCreaTabella.TabIndex = 25;
            this.btnCreaTabella.Text = "Tabella";
            this.btnCreaTabella.UseVisualStyleBackColor = true;
            this.btnCreaTabella.Click += new System.EventHandler(this.btnCreaTabella_Click);
            // 
            // cmbRighe
            // 
            this.cmbRighe.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbRighe.FormattingEnabled = true;
            this.cmbRighe.Location = new System.Drawing.Point(344, 61);
            this.cmbRighe.Name = "cmbRighe";
            this.cmbRighe.Size = new System.Drawing.Size(76, 21);
            this.cmbRighe.TabIndex = 24;
            // 
            // cmbColonne
            // 
            this.cmbColonne.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbColonne.FormattingEnabled = true;
            this.cmbColonne.Location = new System.Drawing.Point(344, 34);
            this.cmbColonne.Name = "cmbColonne";
            this.cmbColonne.Size = new System.Drawing.Size(76, 21);
            this.cmbColonne.TabIndex = 23;
            // 
            // chkItalic
            // 
            this.chkItalic.AutoSize = true;
            this.chkItalic.Location = new System.Drawing.Point(208, 38);
            this.chkItalic.Name = "chkItalic";
            this.chkItalic.Size = new System.Drawing.Size(48, 17);
            this.chkItalic.TabIndex = 22;
            this.chkItalic.Text = "Italic";
            this.chkItalic.UseVisualStyleBackColor = true;
            // 
            // chkBold
            // 
            this.chkBold.AutoSize = true;
            this.chkBold.Location = new System.Drawing.Point(134, 38);
            this.chkBold.Name = "chkBold";
            this.chkBold.Size = new System.Drawing.Size(47, 17);
            this.chkBold.TabIndex = 21;
            this.chkBold.Text = "Bold";
            this.chkBold.UseVisualStyleBackColor = true;
            // 
            // cmbColor
            // 
            this.cmbColor.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbColor.FormattingEnabled = true;
            this.cmbColor.Location = new System.Drawing.Point(7, 88);
            this.cmbColor.Name = "cmbColor";
            this.cmbColor.Size = new System.Drawing.Size(121, 21);
            this.cmbColor.TabIndex = 20;
            // 
            // cmbAlignment
            // 
            this.cmbAlignment.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbAlignment.FormattingEnabled = true;
            this.cmbAlignment.Location = new System.Drawing.Point(134, 61);
            this.cmbAlignment.Name = "cmbAlignment";
            this.cmbAlignment.Size = new System.Drawing.Size(121, 21);
            this.cmbAlignment.TabIndex = 19;
            // 
            // cmbUnderlined
            // 
            this.cmbUnderlined.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbUnderlined.FormattingEnabled = true;
            this.cmbUnderlined.Location = new System.Drawing.Point(134, 88);
            this.cmbUnderlined.Name = "cmbUnderlined";
            this.cmbUnderlined.Size = new System.Drawing.Size(121, 21);
            this.cmbUnderlined.TabIndex = 18;
            // 
            // cmbFont
            // 
            this.cmbFont.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbFont.FormattingEnabled = true;
            this.cmbFont.Location = new System.Drawing.Point(7, 61);
            this.cmbFont.Name = "cmbFont";
            this.cmbFont.Size = new System.Drawing.Size(121, 21);
            this.cmbFont.TabIndex = 17;
            // 
            // cmbSize
            // 
            this.cmbSize.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbSize.FormattingEnabled = true;
            this.cmbSize.Location = new System.Drawing.Point(7, 34);
            this.cmbSize.Name = "cmbSize";
            this.cmbSize.Size = new System.Drawing.Size(121, 21);
            this.cmbSize.TabIndex = 16;
            // 
            // txtTesto
            // 
            this.txtTesto.Location = new System.Drawing.Point(7, 179);
            this.txtTesto.Name = "txtTesto";
            this.txtTesto.Size = new System.Drawing.Size(415, 20);
            this.txtTesto.TabIndex = 15;
            // 
            // btnInserisci
            // 
            this.btnInserisci.Location = new System.Drawing.Point(7, 205);
            this.btnInserisci.Name = "btnInserisci";
            this.btnInserisci.Size = new System.Drawing.Size(75, 23);
            this.btnInserisci.TabIndex = 14;
            this.btnInserisci.Text = "INSERISCI";
            this.btnInserisci.UseVisualStyleBackColor = true;
            this.btnInserisci.Click += new System.EventHandler(this.btnInserisci_Click);
            // 
            // label1
            // 
            this.label1.Location = new System.Drawing.Point(4, 139);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(100, 23);
            this.label1.TabIndex = 13;
            this.label1.Text = "TESTO";
            // 
            // btnSalvaChiudi
            // 
            this.btnSalvaChiudi.Location = new System.Drawing.Point(259, 205);
            this.btnSalvaChiudi.Name = "btnSalvaChiudi";
            this.btnSalvaChiudi.Size = new System.Drawing.Size(163, 23);
            this.btnSalvaChiudi.TabIndex = 12;
            this.btnSalvaChiudi.Text = "SALVA E CHIUDI WORD";
            this.btnSalvaChiudi.UseVisualStyleBackColor = true;
            this.btnSalvaChiudi.Click += new System.EventHandler(this.btnSalvaChiudi_Click);
            // 
            // txtDa
            // 
            this.txtDa.Location = new System.Drawing.Point(587, 104);
            this.txtDa.Name = "txtDa";
            this.txtDa.Size = new System.Drawing.Size(149, 20);
            this.txtDa.TabIndex = 2;
            // 
            // btnSelezionaTesto
            // 
            this.btnSelezionaTesto.Location = new System.Drawing.Point(661, 156);
            this.btnSelezionaTesto.Name = "btnSelezionaTesto";
            this.btnSelezionaTesto.Size = new System.Drawing.Size(75, 23);
            this.btnSelezionaTesto.TabIndex = 26;
            this.btnSelezionaTesto.Text = "INSERISCI";
            this.btnSelezionaTesto.UseVisualStyleBackColor = true;
            this.btnSelezionaTesto.Click += new System.EventHandler(this.btnSelezionaTesto_Click);
            // 
            // txtA
            // 
            this.txtA.Location = new System.Drawing.Point(587, 130);
            this.txtA.Name = "txtA";
            this.txtA.Size = new System.Drawing.Size(149, 20);
            this.txtA.TabIndex = 27;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(490, 107);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(88, 13);
            this.label2.TabIndex = 28;
            this.label2.Text = "SELEZIONA DA:";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(490, 133);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(80, 13);
            this.label3.TabIndex = 29;
            this.label3.Text = "SELEZIONA A:";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(490, 238);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(103, 13);
            this.label4.TabIndex = 34;
            this.label4.Text = "SOSTITUISCI CON:";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(490, 212);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(49, 13);
            this.label5.TabIndex = 33;
            this.label5.Text = "CERCA: ";
            // 
            // txtTestoSostituire
            // 
            this.txtTestoSostituire.Location = new System.Drawing.Point(599, 235);
            this.txtTestoSostituire.Name = "txtTestoSostituire";
            this.txtTestoSostituire.Size = new System.Drawing.Size(137, 20);
            this.txtTestoSostituire.TabIndex = 32;
            // 
            // btnRicerca
            // 
            this.btnRicerca.Location = new System.Drawing.Point(661, 261);
            this.btnRicerca.Name = "btnRicerca";
            this.btnRicerca.Size = new System.Drawing.Size(75, 23);
            this.btnRicerca.TabIndex = 31;
            this.btnRicerca.Text = "RICERCA";
            this.btnRicerca.UseVisualStyleBackColor = true;
            this.btnRicerca.Click += new System.EventHandler(this.btnRicerca_Click);
            // 
            // txtTestoRicercare
            // 
            this.txtTestoRicercare.Location = new System.Drawing.Point(556, 209);
            this.txtTestoRicercare.Name = "txtTestoRicercare";
            this.txtTestoRicercare.Size = new System.Drawing.Size(180, 20);
            this.txtTestoRicercare.TabIndex = 30;
            // 
            // chkSostituisci
            // 
            this.chkSostituisci.AutoSize = true;
            this.chkSostituisci.Location = new System.Drawing.Point(493, 265);
            this.chkSostituisci.Name = "chkSostituisci";
            this.chkSostituisci.Size = new System.Drawing.Size(93, 17);
            this.chkSostituisci.TabIndex = 35;
            this.chkSostituisci.Text = "SOSTITUISCI";
            this.chkSostituisci.UseVisualStyleBackColor = true;
            // 
            // btnCreaPDF
            // 
            this.btnCreaPDF.Location = new System.Drawing.Point(7, 249);
            this.btnCreaPDF.Name = "btnCreaPDF";
            this.btnCreaPDF.Size = new System.Drawing.Size(75, 23);
            this.btnCreaPDF.TabIndex = 26;
            this.btnCreaPDF.Text = "CREA PDF";
            this.btnCreaPDF.UseVisualStyleBackColor = true;
            this.btnCreaPDF.Click += new System.EventHandler(this.btnCreaPDF_Click);
            // 
            // btnStampa
            // 
            this.btnStampa.Location = new System.Drawing.Point(88, 249);
            this.btnStampa.Name = "btnStampa";
            this.btnStampa.Size = new System.Drawing.Size(75, 23);
            this.btnStampa.TabIndex = 27;
            this.btnStampa.Text = "STAMPA";
            this.btnStampa.UseVisualStyleBackColor = true;
            this.btnStampa.Click += new System.EventHandler(this.btnStampa_Click);
            // 
            // frmWordCSharp
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(755, 452);
            this.Controls.Add(this.chkSostituisci);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.txtTestoSostituire);
            this.Controls.Add(this.btnRicerca);
            this.Controls.Add(this.txtTestoRicercare);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.txtA);
            this.Controls.Add(this.btnSelezionaTesto);
            this.Controls.Add(this.txtDa);
            this.Controls.Add(this.grbWordCreated);
            this.Controls.Add(this.btnCrea);
            this.Name = "frmWordCSharp";
            this.StartPosition = System.Windows.Forms.FormStartPosition.Manual;
            this.Text = "Gestione Word da C#";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.frmWordCSharp_FormClosing);
            this.Load += new System.EventHandler(this.frmWordCSharp_Load);
            this.grbWordCreated.ResumeLayout(false);
            this.grbWordCreated.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnCrea;
        private System.Windows.Forms.GroupBox grbWordCreated;
        private System.Windows.Forms.CheckBox chkItalic;
        private System.Windows.Forms.CheckBox chkBold;
        private System.Windows.Forms.ComboBox cmbColor;
        private System.Windows.Forms.ComboBox cmbAlignment;
        private System.Windows.Forms.ComboBox cmbUnderlined;
        private System.Windows.Forms.ComboBox cmbFont;
        private System.Windows.Forms.ComboBox cmbSize;
        private System.Windows.Forms.TextBox txtTesto;
        private System.Windows.Forms.Button btnInserisci;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btnSalvaChiudi;
        private System.Windows.Forms.Button btnCreaTabella;
        private System.Windows.Forms.ComboBox cmbRighe;
        private System.Windows.Forms.ComboBox cmbColonne;
        private System.Windows.Forms.TextBox txtDa;
        private System.Windows.Forms.Button btnSelezionaTesto;
        private System.Windows.Forms.TextBox txtA;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TextBox txtTestoSostituire;
        private System.Windows.Forms.Button btnRicerca;
        private System.Windows.Forms.TextBox txtTestoRicercare;
        private System.Windows.Forms.CheckBox chkSostituisci;
        private System.Windows.Forms.Button btnCreaPDF;
        private System.Windows.Forms.Button btnStampa;
    }
}

