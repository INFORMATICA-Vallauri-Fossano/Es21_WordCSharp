using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using WordCSharp_ns;

namespace Es21_WordCSharpù
{
    public partial class frmWordCSharp : Form
    {
        clsWord word=new clsWord();
        public frmWordCSharp()
        {
            InitializeComponent();
        }

        private void btnCrea_Click(object sender, EventArgs e)
        {
            word.CreaDocumento(true);
            if(!grbWordCreated.Enabled) grbWordCreated.Enabled = true;
        }

        private void btnSalvaChiudi_Click(object sender, EventArgs e)
        {
            word.SalvaChiudi("CiaoBro\\Weoeoew");
            grbWordCreated.Enabled=false;
        }

        private void btnInserisci_Click(object sender, EventArgs e)
        {
            object start=0;
            object end=2;
            //MessageBox.Show("TO-DO");
            //word.impostaRange(ref start, ref end);
            word.impostaRange(ref start, ref end);
            word.InserisciTesto(start, end,txtTesto.Text,cmbFont.Text,cmbSize.Text,chkBold.Checked,chkItalic.Checked,cmbUnderlined.Text,cmbAlignment.Text,cmbColor.Text);
        }

        private void frmWordCSharp_Load(object sender, EventArgs e)
        {
            StartPosition = 0;
            word.impostaFont(cmbFont);
            word.impostaSize(cmbSize);
            word.impostaUnderline(cmbUnderlined);
            word.impostaAlignment(cmbAlignment);
            word.impostaColor(cmbColor);
            word.impostaRigheColonne(cmbRighe, cmbColonne);
        }

        private void frmWordCSharp_FormClosing(object sender, FormClosingEventArgs e)
        {
            word.SalvaChiudi();
        }

        private void btnCreaTabella_Click(object sender, EventArgs e)
        {
            object start = 0;
            object end = 2;
            //MessageBox.Show("TO-DO");
            //word.impostaRange(ref start, ref end);
            word.impostaRange(ref start, ref end);
            word.InserisciTabella(Convert.ToInt32(cmbColonne.Text),Convert.ToInt32(cmbRighe.Text), ref start, ref end);
        }
    }
}
