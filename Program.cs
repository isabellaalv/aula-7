using System;
using System.Drawing;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;

namespace aula_7 {
    class Program {
        static void Main (string[] args) {
            #region Criacao do documento
            Document exemploDoc = new Document ();
            #endregion

            #region Criacao de secao no documento
            //cada secao é tipo uma pagina

            Section secaoCapa = exemploDoc.AddSection ();

            #endregion
            // cria um paragrafo e adc na secao
            //os paragrafos são necesseraio para a insençao de texto, img e outra coisas
            #region Adiciona um paragrafo

            Paragraph Titulo = secaoCapa.AddParagraph ();

            #endregion

            #region Adiciona texto no paragrafo

            //adc o texto ao paragrafo titulo
            Titulo.AppendText ("Exempo de titulo\n\n");

            #endregion

            #region Formata paragrafo

            //atraves da propriedade Hozi..Ali.., é possivel alinhar o paragrafo
            Titulo.Format.HorizontalAlignment = HorizontalAlignment.Center;

            ParagraphStyle estilo01 = new ParagraphStyle (exemploDoc);

            //adc um nome ao estilo01
            estilo01.Name = "Cor do titulo";

            //difinir a cor do titulo
            estilo01.CharacterFormat.TextColor = Color.DarkBlue;

            // define q o texto sera negrito
            estilo01.CharacterFormat.Bold = true;

            // adc estilo01 ao doc exemplodoc
            exemploDoc.Styles.Add (estilo01);

            //Aplica o estilo01 ao paragrafo titulo
            Titulo.ApplyStyle (estilo01.Name);

            #endregion

            #region Trabalhar cm tabulacao

            // adc um parag.. textocapa a secaocapa
            Paragraph TextoCapa = secaoCapa.AddParagraph ();

            TextoCapa.AppendText ("\t este é um exeplo de texto cm tabulaçaõ\n");

            Paragraph TextoCapa2 = secaoCapa.AddParagraph ();

            TextoCapa2.AppendText ("\t pgn doc, paragrafo na mesma seção" + "obviamente na mesma seção");

            #endregion

            #region Inserir Imagem

            Paragraph ImagemCapa = secaoCapa.AddParagraph ();

            ImagemCapa.AppendText ("\n\n\t agrora vamos inserir uma img no doc\n\n");

            //paragrafo com img
            ImagemCapa.Format.HorizontalAlignment = HorizontalAlignment.Center;

            DocPicture ImagemExemplo = ImagemCapa.AppendPicture (Image.FromFile (@"saida\logo_csharp.png"));

            //defini altura e largura img

            ImagemExemplo.Width = 300;
            ImagemExemplo.Height = 300;

            #endregion

            #region Adc nova secao

            //adc nova secao
            Section secaoCorpo = exemploDoc.AddSection ();

            Paragraph paragrafoCorpo1 = secaoCorpo.AddParagraph ();

            paragrafoCorpo1.AppendText ("exemplo de paragrafo" + "\t o texto aparece em uma nova pgn por ser uma nova secao");

            #endregion

            #region Adc uma tabela 

            //criacao de tabela
            Table tabela = secaoCorpo.AddTable (true);

            String[] cabecalho = { "Item", "Descrição", "Qtd", "Preço Unit", "Preço" };

            String[][] dados = {
                new String[] { "Cenoura", "Vegetal Nutritivo", "1", "R$ 4,00", "R$ 4,00" },
                new String[] { "Batata", "Vegetal Nutritivo", "1", "R$ 4,00", "R$ 4,00" },
                new String[] { "Banana", "Vegetal Nutritivo", "1", "R$ 4,00", "R$ 4,00" },
                new String[] { "Tomate", "Vegetal Nutritivo", "1", "R$ 4,00", "R$ 4,00" },
            };

            //adc celular
            tabela.ResetCells (dados.Length + 1, cabecalho.Length);

            // adc uma linha na posicao 0 do vetor
            // define q a linha é um cabecalho
            TableRow Linha1 = tabela.Rows[0];
            Linha1.IsHeader = true;

            //define altura da linha
            Linha1.Height = 23;

            //formata cabecalho
            Linha1.RowFormat.BackColor = Color.AliceBlue;

            for (int i = 0; i < cabecalho.Length; i++) {

                Paragraph p = Linha1.Cells[i].AddParagraph();
                Linha1.Cells[i].CellFormat.VerticalAlignment = VerticalAlignment.Middle;

                p.Format.HorizontalAlignment = HorizontalAlignment.Center;

                //formata dados do cabecalho

                TextRange TR = p.AppendText(cabecalho[i]);
                TR.CharacterFormat.FontName = "Calibre";
                TR.CharacterFormat.FontSize = 14;
                TR.CharacterFormat.TextColor = Color.Teal;
                TR.CharacterFormat.Bold = true;

            }

            //adc  as linhas do corpo da tabela

            for (int r = 0; r < dados.Length; r++) {
                TableRow LinhaDados = tabela.Rows[r + 1];

                //define altura da linha
                LinhaDados.Height = 20;

                for (int c = 0; c < dados[r].Length; c++)
                {
                    //alinha as celulas
                    LinhaDados.Cells[c].CellFormat.VerticalAlignment = VerticalAlignment.Middle;

                    //preenche dados nas linhas
                    Paragraph p2 = LinhaDados.Cells[c].AddParagraph();

                    TextRange TR2 = p2.AppendText(dados[r][c]);

                    //formata celulas

                    p2.Format.HorizontalAlignment = HorizontalAlignment.Center;
                    TR2.CharacterFormat.FontName = "Calibre";
                    TR2.CharacterFormat.FontSize = 12;
                    TR2.CharacterFormat.TextColor = Color.Brown;

                }
            }

            #endregion

            #region Salvar Arquivo
                
                // salva arquivo em docx
                //savefile para salvar como deseja
                exemploDoc.SaveToFile(@"saida\aula-7.docx", FileFormat.Docx);
                
            #endregion
        }
    }
}