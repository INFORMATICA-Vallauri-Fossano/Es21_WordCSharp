using Microsoft.Office.Interop.Word;
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
            int c = Convert.ToInt32(cmbColonne.Text);
            int r = Convert.ToInt32(cmbRighe.Text);
            Table tabella;
            tabella = word.InserisciTabella(r, c, ref start, ref end);
            word.scriviCella(tabella, r, c, "ciao", WdCellVerticalAlignment.wdCellAlignVerticalCenter, WdParagraphAlignment.wdAlignParagraphCenter, true, 15, "arial", WdColor.wdColorBlue);
            Table tabella1;

            word.impostaRange(ref start, ref end);
            word.InserisciTesto(start,end,"\n");
            word.impostaRange(ref start, ref end);

            tabella1 = word.InserisciTabella(r, c, ref start, ref end);
            word.scriviCella(tabella1, r, c, "ciao", WdCellVerticalAlignment.wdCellAlignVerticalCenter, WdParagraphAlignment.wdAlignParagraphCenter, true, 15, "arial", WdColor.wdColorBlue);
        }

        private void btnSelezionaTesto_Click(object sender, EventArgs e)
        {
            object start = 0;
            object end = 0;
            //posiziono il cursore all'inizio del documento
                word.impostaRange(ref start, ref end);
            try
            {
                start = Convert.ToInt16(txtDa.Text);
                end = Convert.ToInt16(txtA.Text);

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            MessageBox.Show(word.Select(start, end));
        }

        private void btnRicerca_Click(object sender, EventArgs e)
        {
            object start=0, end = 0;
            word.impostaRange(ref start, ref end);
            if (word.ricercaSostituisci(txtTestoRicercare.Text, txtTestoSostituire.Text, chkSostituisci.Checked, ref start, ref end)) MessageBox.Show("Trovato");
            else MessageBox.Show("Testop non trovato");
        }
    }
}
