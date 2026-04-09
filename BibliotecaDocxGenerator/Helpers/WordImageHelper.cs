using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Drawing;
using System.IO;

using D = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using DP = DocumentFormat.OpenXml.Drawing.Pictures;

namespace Helpers
{
    /// <summary>
    /// Classe auxiliar para manipulação de imagens em documentos Word Open XML.
    /// </summary>
    public class WordImageHelper
    {
        private uint _imgIdCounter = 1U;

        /// <summary>
        /// Adiciona uma imagem ao corpo do documento.
        /// </summary>
        /// <param name="body">Corpo do documento onde a imagem será inserida.</param>
        /// <param name="mainPart">Parte principal do documento (dona da imagem).</param>
        /// <param name="caminhoImagem">Caminho completo do arquivo de imagem.</param>
        /// <param name="larguraPx">Largura desejada em pixels.</param>
        /// <param name="alturaPx">Altura desejada em pixels (0 para manter proporção).</param>
        public void AdicionarImagem(Body body, MainDocumentPart mainPart,
                                           string caminhoImagem, int larguraPx, int alturaPx)
        {
            var p = new Paragraph();
            var pp = new ParagraphProperties(new Justification { Val = JustificationValues.Center });
            p.AppendChild(pp);
            AdicionarImagemNoParagrafo(p, mainPart, caminhoImagem, larguraPx, alturaPx);
            body.AppendChild(p);
        }

        /// <summary>
        /// Adiciona uma imagem a um parágrafo existente, suportando diferentes tipos de partes (MainDocumentPart, HeaderPart, FooterPart).
        /// </summary>
        /// <param name="paragraph">Parágrafo onde a imagem será inserida.</param>
        /// <param name="ownerPart">Parte proprietária da imagem (MainDocumentPart, HeaderPart ou FooterPart).</param>
        /// <param name="caminhoImagem">Caminho completo do arquivo de imagem.</param>
        /// <param name="larguraPx">Largura desejada em pixels.</param>
        /// <param name="alturaPx">Altura desejada em pixels (0 para manter proporção).</param>
        public void AdicionarImagemNoParagrafo(Paragraph paragraph, OpenXmlPart ownerPart,
                                                      string caminhoImagem, int larguraPx, int alturaPx)
        {
            if (string.IsNullOrEmpty(caminhoImagem) || !File.Exists(caminhoImagem))
                return;

            try
            {
                byte[] bytes = File.ReadAllBytes(caminhoImagem);
                if (bytes.Length == 0) return;

                // Ajusta altura proporcionalmente se altura for zero
                if (alturaPx == 0)
                {
                    try
                    {
                        using (var bmp = new Bitmap(new MemoryStream(bytes)))
                        {
                            if (bmp.Width > 0)
                                alturaPx = (int)(larguraPx * ((double)bmp.Height / bmp.Width));
                        }
                    }
                    catch { }
                }

                long cx = larguraPx * 9525L;
                long cy = alturaPx * 9525L;

                string ext = Path.GetExtension(caminhoImagem).ToLowerInvariant();
                // Usar ImagePartType diretamente, sem alias
                //ImagePartType partType = (ext == ".jpg" || ext == ".jpeg")
                //    ? ImagePartType.Jpeg : ImagePartType.Png;

                ImagePart imagePart;
                string relationshipId;

                // Adiciona a ImagePart ao ownerPart apropriado
                switch (ownerPart)
                {
                    case MainDocumentPart mp:
                        imagePart = mp.AddImagePart(ImagePartType.Jpeg);
                        using (var s = new MemoryStream(bytes)) imagePart.FeedData(s);
                        relationshipId = mp.GetIdOfPart(imagePart);
                        break;
                    case HeaderPart hp:
                        imagePart = hp.AddImagePart(ImagePartType.Jpeg);
                        using (var s = new MemoryStream(bytes)) imagePart.FeedData(s);
                        relationshipId = hp.GetIdOfPart(imagePart);
                        break;
                    case FooterPart fp:
                        imagePart = fp.AddImagePart(ImagePartType.Jpeg);
                        using (var s = new MemoryStream(bytes)) imagePart.FeedData(s);
                        relationshipId = fp.GetIdOfPart(imagePart);
                        break;
                    default:
                        return;
                }

                uint id = _imgIdCounter++;

                // Cria o elemento Drawing com a estrutura correta
                var drawing = new Drawing(
                    new DW.Inline(
                        new DW.Extent { Cx = cx, Cy = cy },
                        new DW.EffectExtent
                        {
                            LeftEdge = 0L,
                            TopEdge = 0L,
                            RightEdge = 0L,
                            BottomEdge = 0L
                        },
                        new DW.DocProperties { Id = id, Name = "Imagem" + id },
                        new DW.NonVisualGraphicFrameDrawingProperties(
                            new D.GraphicFrameLocks { NoChangeAspect = true }),
                        new D.Graphic(
                            new D.GraphicData(
                                new DP.Picture(
                                    new DP.NonVisualPictureProperties(
                                        new DP.NonVisualDrawingProperties { Id = 0U, Name = "Imagem" },
                                        new DP.NonVisualPictureDrawingProperties()
                                    ),
                                    new DP.BlipFill(
                                        new D.Blip { Embed = relationshipId, CompressionState = D.BlipCompressionValues.Print },
                                        new D.Stretch(new D.FillRectangle())
                                    ),
                                    new DP.ShapeProperties(
                                        new D.Transform2D(
                                            new D.Offset { X = 0L, Y = 0L },
                                            new D.Extents { Cx = cx, Cy = cy }
                                        ),
                                        new D.PresetGeometry(new D.AdjustValueList())
                                        { Preset = D.ShapeTypeValues.Rectangle }
                                    )
                                )
                            )
                            { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" }
                        )
                    )
                    {
                        DistanceFromTop = 0U,
                        DistanceFromBottom = 0U,
                        DistanceFromLeft = 0U,
                        DistanceFromRight = 0U
                    }
                );

                paragraph.AppendChild(new Run(drawing));
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("Erro ao adicionar imagem: " + ex.Message);
            }
        }

        /// <summary>
        /// Reinicia o contador de IDs de imagens. Útil ao reutilizar o helper em múltiplos documentos.
        /// </summary>
        public void ResetImageCounter()
        {
            _imgIdCounter = 1U;
        }
    }
}