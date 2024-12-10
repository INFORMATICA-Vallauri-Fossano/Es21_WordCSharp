using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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
            myWord = new Application();
            myWord.Visible = visible;
            //istanzio il documento
            myDocument = myWord.Documents.Add();
        }

        internal void impostaRange(ref object start, ref object end)
        {
            throw new NotImplementedException();
        }

        internal void InserisciTesto(object start, object end, string text,string font="Arial",string size="12",bool bold=false,bool italic=false,string sottolineato="None",string allinemanto="Left",string colore="Black")
        {
            Range myrange= myDocument.Range(ref start,ref end);
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
