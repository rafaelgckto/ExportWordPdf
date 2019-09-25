using System;
using Spire.Doc;
using Spire.Doc.Documents;
using System.Drawing;

namespace RepositoryWordPdf
{
    class Program
    {
        static void Main(string[] args)
        {
            // Criando um novo documento com o nome documento
            Document doc = new Document();
            
            // Criando uma seção dentro do documento
            // A cada seção criada uma nova página é adicionada
            Section secaoCapa = doc.AddSection();

            Paragraph titulo = secaoCapa.AddParagraph();

            titulo.AppendText("Título muito bonito\n\n");
            titulo.Format.HorizontalAlignment = HorizontalAlignment.Center;

            ParagraphStyle estilo01 = new ParagraphStyle(doc);
            estilo01.Name = "Cor do título"; //Define o nome da classe estilo01
            estilo01.CharacterFormat.TextColor = Color.DarkBlue;
            estilo01.CharacterFormat.Bold = true;
            
            doc.Styles.Add(estilo01);
            titulo.ApplyStyle(estilo01.Name);

            doc.SaveToFile(@"Saída\exemploWord.docx", FileFormat.Docx);
        }
    }
}
