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
        }

        private void btnSalvaChiudi_Click(object sender, EventArgs e)
        {
            word.SalvaChiudi("CiaoBro\\Weoeoew");
        }

        private void btnInserisci_Click(object sender, EventArgs e)
        {
            object start=0;
            object end=0;
            MessageBox.Show("TO-DO");
            //word.impostaRange(ref start, ref end);
            word.InserisciTesto(start, end,txtTesto.Text);
        }
    }
}
