using Microsoft.Office.Interop.Word;
using Microsoft.VisualStudio.TextManager.Interop;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Font = Microsoft.Office.Interop.Word.Font;

namespace WordCSharp_ns
{
    internal class clsWord
    {
        //applicazione word
        _Application myWord;
        //documento word
        Document myDocument;
        internal void CreaDocumento(bool visible=true)
        {
            //istanzio l'applicazione word
            myWord = new Microsoft.Office.Interop.Word.Application();
            myWord.Visible = visible;
            //istanzio il documento
            myDocument = myWord.Documents.Add();
        }

        internal void impostaAlignment(System.Windows.Forms.ComboBox cmbAlignment)
        {
            string[] wdU = (string[])Enum.GetNames(typeof(WdParagraphAlignment));
            foreach (string u in wdU)
            {
                cmbAlignment.Items.Add(u.Substring(16));
            }
            cmbAlignment.SelectedIndex = 0;
        }

        internal void impostaColor(System.Windows.Forms.ComboBox cmbColor)
        {
            string[] wdU = (string[])Enum.GetNames(typeof(WdColor));
            foreach (string u in wdU)
            {
                cmbColor.Items.Add(u.Substring(7));
            }
            cmbColor.SelectedIndex = 0;
        }

        internal void impostaFont(System.Windows.Forms.ComboBox cmbFont)
        {
            foreach(FontFamily family in FontFamily.Families)
            {
                cmbFont.Items.Add(family.Name);
            }
            cmbFont.SelectedIndex = 11;

        }

        internal void impostaRange(ref object start, ref object end)
        {
            start = myDocument.Sentences[myDocument.Sentences.Count].End-1;
                                                                    //-1 per il carattere speciale
            end = myDocument.Sentences[myDocument.Sentences.Count].End;

        }

        internal void impostaRigheColonne(ComboBox cmbRighe, ComboBox cmbColonne)
        {
                for (int i = 1; i < 10; i ++)
                {
                    cmbRighe.Items.Add(i.ToString());
                    cmbColonne.Items.Add(i.ToString());
                }
        }

        internal void impostaSize(System.Windows.Forms.ComboBox cmbSize)
        {
            for(int i = 8; i < 50; i+=2)
            {
                cmbSize.Items.Add(i.ToString());
            }
            cmbSize.SelectedIndex = 4;
        }

        internal void impostaUnderline(System.Windows.Forms.ComboBox cmbUnderlined)
        {
            string[] wdU = (string[])Enum.GetNames(typeof(WdUnderline));
            foreach(string u in wdU)
            {
                cmbUnderlined.Items.Add(u.Substring(11));
            }
            cmbUnderlined.SelectedIndex = 0;

        }

        internal void InserisciTabella(int c, int r,ref object start, ref object end)
        {
            Table tabella;
            tabella = myDocument.Tables.Add(myDocument.Range(ref start, ref end), r, c);
            tabella.Borders.Enable = 1; 

        }


        internal void InserisciTesto(object start, object end, string text,string font="Arial",
            string size="12",bool bold=false,bool italic=false,string sottolineato="None",
            string allinemanto="Left",string colore="Black")
        {
            Range myrange= myDocument.Range(ref start,ref end);
            myrange.Font.Name = font;
            myrange.Font.Size=Convert.ToInt16(size);
            myrange.Bold = Convert.ToInt16(bold);
            myrange.Italic = Convert.ToInt16(italic);
            myrange.Underline = (WdUnderline)Enum.Parse(typeof(WdUnderline), "wdUnderline" + sottolineato);
            myrange.ParagraphFormat.Alignment = (WdParagraphAlignment)Enum.Parse(typeof(WdParagraphAlignment), "wdAlignParagraph" + allinemanto);
            myrange.Font.Color = (WdColor)Enum.Parse(typeof(WdColor), "wdColor" + colore);
            myrange.Text = text;
        }

        internal void SalvaChiudi(string nomefile="")
        {
            if (myWord != null)
            {
                if (nomefile == "")
                    myDocument.Save();
                else
                    myDocument.SaveAs(nomefile);
                chiudi();
            }

        }

        private void chiudi()
        {
            myDocument.Saved = true;
            myDocument.Close();
            myWord.Quit();
        }
    }
}
