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

        internal void creaPDF(string path, bool visualizzaPDF)
        {
            myDocument.ExportAsFixedFormat(path, WdExportFormat.wdExportFormatPDF, visualizzaPDF);
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

        internal Table InserisciTabella(int r, int c,ref object start, ref object end)
        {
            Table tabella;
            tabella = myDocument.Tables.Add(myDocument.Range(ref start, ref end), r, c);
            tabella.Borders.Enable = 1; 
            return tabella;
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

        internal bool ricercaSostituisci(string ricercato, ref object start, ref object end)
        {
            bool trovato = false;
            object findText = ricercato;
            
            myWord.Selection.Find.ClearFormatting();
            myWord.Selection.Start = myDocument.Content.Start;
            myWord.Selection.End = myDocument.Content.End;

            if(myWord.Selection.Find.Execute(ref findText)) trovato = true;
          
            return trovato;
        }

        
        internal bool ricercaSostituisci(string ricercato, string sostituto, bool sostituisci, ref object start, ref object end)
        {
            bool trovato = false;
            object findText = ricercato;
            object replaceText = sostituto;
            object ms = System.Type.Missing;
            myWord.Selection.Find.ClearFormatting();
            myWord.Selection.Start = myDocument.Content.Start;
            myWord.Selection.End = myDocument.Content.End;

            if (myWord.Selection.Find.Execute(ref findText))
            {
                trovato = true;
                myWord.Selection.Find.Execute(ref findText, ref ms, ref ms, ref ms, ref ms, ref ms, ref ms, ref ms, ref ms, ref replaceText, sostituisci ? Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll : Microsoft.Office.Interop.Word.WdReplace.wdReplaceNone);
            }
            return trovato;
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

        /// <summary>
        /// Metodo per scrivere all'interno di una cella specificata da riga, colonna e formattazione
        /// </summary>
        /// <param name="tabella">Puntatore della tabella a cui fare riferimento</param>
        /// <param name="r"></param>
        /// <param name="c"></param>
        /// <param name="testo"></param>
        /// <param name="vAlign"></param>
        /// <param name="oAlign"></param>
        /// <param name="bold"></param>
        /// <param name="size"></param>
        /// <param name="font"></param>
        /// <param name="color"></param>
        internal void scriviCella(Table tabella, int r, int c, string testo, WdCellVerticalAlignment vAlign, WdParagraphAlignment oAlign, bool bold, int size, string font, WdColor color)
        {
          tabella.Cell(r, c).Range.Text=testo;
          tabella.Cell(r, c).Range.Cells.VerticalAlignment=vAlign;
            tabella.Cell(r, c).Range.Paragraphs.Alignment=oAlign;
            tabella.Cell(r, c).Range.Bold = Convert.ToInt32(bold);
            tabella.Cell(r, c).Range.Font.Size= size;
            tabella.Cell(r, c).Range.Font.Name= font;
            tabella.Cell(r, c).Range.Font.Color= color;
        }

        internal string Select(object start, object end)
        {
            Range range = myDocument.Range(ref start,ref end);
            range.Select();
            
            return range.Text;
        }

        internal void Stampa()
        {
            myDocument.PrintOut();
        }

        private void chiudi()
        {
            myDocument.Saved = true;
            myDocument.Close();
            myWord.Quit();
        }
    }
}
