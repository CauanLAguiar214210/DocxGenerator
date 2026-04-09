using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Wordprocessing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Web;
using Color = DocumentFormat.OpenXml.Wordprocessing.Color;
using HorizontalAlignmentValues = DocumentFormat.OpenXml.Wordprocessing.HorizontalAlignmentValues;
using VerticalAlignmentValues = DocumentFormat.OpenXml.Wordprocessing.VerticalAlignmentValues;

namespace Helpers
{
    public class OfficeWordHelper
    {

        #region Constantes semânticas
        internal static class WordFont
        {
            public const string Default = "Times New Roman";
        }

        internal static class WordSize
        {
            public const string Corpo = "24"; // 12pt
            public const string Citacao = "20"; // 10pt
            public const string Assinatura = "22"; // 11pt
            public const string Titulo = "26"; // 13pt
            public const string Rodape = "18"; // 9pt
        }

        internal static class WordSpacing
        {
            public const string Simples = "240";
            public const string Duplo = "360";
            public const string Rodape = "80";
        }

        internal static class WordIndent
        {
            public const string PrimeiraLinhaCorpo = "709";  // 1,25 cm
            public const string RecuoCitacao = "2268"; // 4 cm
        }
        #endregion

        #region Fragmento de texto configurável dentro de um parágrafo misto
        public sealed class Trecho
        {
            public string Texto { get; private set; }
            public string Tamanho { get; private set; }
            public bool Negrito { get; private set; }
            public bool Italico { get; private set; }
            public string Cor { get; private set; }
            public bool Quebra { get; private set; }
            public JustificationValues? Alinhamento { get; private set; }
            public SpaceProcessingModeValues? Espacamento { get; private set; }
            public string RecuoPrimeiraLinha { get; private set; }


            private Trecho() { }

            /// <summary>Inicia um builder para o trecho com o texto fornecido.</summary>
            public static TrechoBuilder T(string texto)
            {
                return new TrechoBuilder(texto);
            }

            internal static Trecho Criar(string texto, string tamanho,
                                         bool negrito, bool italico, string cor, bool quebra, JustificationValues? alinhamento, SpaceProcessingModeValues? espacamento, string recuoPrimeiraLinha)
            {
                return new Trecho
                {
                    Texto = texto ?? string.Empty,
                    Tamanho = tamanho ?? WordSize.Corpo,
                    Negrito = negrito,
                    Italico = italico,
                    Cor = cor,
                    Quebra = quebra,
                    Alinhamento = alinhamento,
                    Espacamento = espacamento,
                    RecuoPrimeiraLinha = recuoPrimeiraLinha
                };
            }
        }

        public sealed class TrechoBuilder
        {
            private string _texto;
            private string _tamanho = WordSize.Corpo;
            private bool _negrito;
            private bool _italico;
            private string _cor;
            private bool _quebra;
            private JustificationValues? _alinhamento;
            private SpaceProcessingModeValues? _espacamento;
            private string _recuoPrimeiraLinha;

            internal TrechoBuilder(string texto) { _texto = texto ?? string.Empty; }

            // ── formatação ───────────────────────────────────────────────────────────
            public TrechoBuilder Negrito() { _negrito = true; return this; }
            public TrechoBuilder Italico() { _italico = true; return this; }
            public TrechoBuilder Cor(string hexSemHash) { _cor = hexSemHash; return this; }
            public TrechoBuilder Vermelho() { _cor = "CC0000"; return this; }
            public TrechoBuilder Azul() { _cor = "0000CC"; return this; }
            public TrechoBuilder Cinza() { _cor = "666666"; return this; }

            public Trecho Quebra() { _quebra = true; return this; }

            // ── tamanho ──────────────────────────────────────────────────────────────
            public TrechoBuilder Tamanho(string tam) { _tamanho = tam; return this; }
            public TrechoBuilder TamCorpo() { _tamanho = WordSize.Corpo; return this; }
            public TrechoBuilder TamAssinatura() { _tamanho = WordSize.Assinatura; return this; }
            public TrechoBuilder TamCitacao() { _tamanho = WordSize.Citacao; return this; }
            public TrechoBuilder TamTitulo() { _tamanho = WordSize.Titulo; return this; }

            // ── alinhamento ──────────────────────────────────────────────────────────
            public TrechoBuilder AlinhadoEsquerda() { _alinhamento = JustificationValues.Left; return this; }
            public TrechoBuilder AlinhadoDireita() { _alinhamento = JustificationValues.Right; return this; }
            public TrechoBuilder AlinhadoCentro() { _alinhamento = JustificationValues.Center; return this; }
            public TrechoBuilder AlinhadoJustify() { _alinhamento = JustificationValues.Both; return this; }
            public TrechoBuilder AlinhadoInicio() { _alinhamento = JustificationValues.Start; return this; }
            public TrechoBuilder AlinhadoFim() { _alinhamento = JustificationValues.End; return this; }
            public TrechoBuilder Alinhamento(JustificationValues val) { _alinhamento = val; return this; }

            public TrechoBuilder EspacamentoSimples() { _espacamento = SpaceProcessingModeValues.Preserve; return this; }
            public TrechoBuilder EspacamentoDefault() { _espacamento = SpaceProcessingModeValues.Default; return this; }

            public TrechoBuilder RecuoCorpo()
            {
                _recuoPrimeiraLinha = WordIndent.PrimeiraLinhaCorpo;
                return this;
            }

            public TrechoBuilder RecuoPrimeiraLinha(string twips)
            {
                _recuoPrimeiraLinha = twips;
                return this;
            }

            public TrechoBuilder SemRecuo()
            {
                _recuoPrimeiraLinha = null;
                return this;
            }

            public static implicit operator Trecho(TrechoBuilder b)
            {
                return b.Build();
            }

            public Trecho Build()
            {
                return Trecho.Criar(_texto, _tamanho, _negrito, _italico, _cor, _quebra, _alinhamento, _espacamento,
                            _recuoPrimeiraLinha);
            }
        }

        #endregion

        #region DTO Templete
        public class TempleteWord
        {
            public string CabecalhoCliente { get; set; }
            public string CaminhoLogoHeader { get; set; }
            public string Titulo { get; set; }
            public string BackgroundImage { get; set; }
            public string CaminhoLogoRodape { get; set; }
            public string RodapeCliente { get; set; }
        }

        #endregion

        public interface IParagraphFactory
        {
            #region Métodos simples (texto único)
            Paragraph ParEspaco();
            Paragraph ParTitulo(string txt);
            Paragraph ParCorpo(string txt);
            Paragraph ParCitacao(string txt);
            Paragraph ParEsquerda(string txt, string tam = WordSize.Corpo, bool negrito = false);
            Paragraph ParDireita(string txt, string tam = WordSize.Corpo, bool negrito = false, bool italico = false);
            Paragraph ParCentro(string txt, string tam = WordSize.Corpo, bool negrito = false, bool italico = false);
            Paragraph ParCentroSimples(string txt, string tam = WordSize.Corpo, bool negrito = false);
            Paragraph ParQuebraPagina();
            #endregion

            #region Métodos mistos (múltiplos trechos formatados)
            Paragraph ParMistoCorpo(params Trecho[] trechos);
            Paragraph ParMistoEsquerda(params Trecho[] trechos);
            Paragraph ParMistoCentro(params Trecho[] trechos);
            Paragraph ParMistoDireita(params Trecho[] trechos);
            #endregion
        }

        public class PropsFactory
        {
            public ParagraphProperties PP(JustificationValues alinhamento,
                                           string espac = WordSpacing.Duplo,
                                           string antes = "0",
                                           string depois = "0")
            {
                var pp = new ParagraphProperties();
                pp.AppendChild(new Justification { Val = alinhamento });
                pp.AppendChild(new SpacingBetweenLines
                {
                    Line = espac,
                    LineRule = LineSpacingRuleValues.Auto,
                    Before = antes,
                    After = depois,
                    AfterAutoSpacing = false,
                    BeforeAutoSpacing = false
                });
                return pp;
            }

            public RunProperties RP(string tam = WordSize.Corpo,
                                    bool negrito = false,
                                    bool italico = false,
                                    string cor = null)
            {
                var rp = new RunProperties();
                rp.AppendChild(new RunFonts
                {
                    Ascii = WordFont.Default,
                    HighAnsi = WordFont.Default,
                    ComplexScript = WordFont.Default
                });
                rp.AppendChild(new FontSize { Val = tam });
                rp.AppendChild(new FontSizeComplexScript { Val = tam });
                if (negrito) rp.AppendChild(new Bold());
                if (italico) rp.AppendChild(new Italic());
                if (!string.IsNullOrEmpty(cor))
                    rp.AppendChild(new Color { Val = cor });
                return rp;
            }

            public RunProperties RPDeTrecho(Trecho trecho)
            {
                return RP(trecho.Tamanho, trecho.Negrito, trecho.Italico, trecho.Cor);
            }
        }

        #region Métodos de extensão para facilitar a criação de parágrafos
        public class ParagraphFactory : IParagraphFactory
        {
            private readonly PropsFactory _props;

            public ParagraphFactory(PropsFactory props) { _props = props; }

            #region Helpers privados
            private Run CriarRun(string texto, RunProperties rp)
            {
                return new Run(rp,
                    new Text(texto) { Space = SpaceProcessingModeValues.Preserve });
            }

            private Run CriarRunDeTrecho(Trecho trecho)
            {
                return CriarRun(trecho.Texto, _props.RPDeTrecho(trecho));
            }

            private Paragraph AdicionarTrechos(Paragraph p, IEnumerable<Trecho> trechos)
            {
                foreach (var trecho in trechos)
                {
                    p.AppendChild(CriarRunDeTrecho(trecho));

                    if (trecho.Texto.Contains("\n"))
                        p.AppendChild(new Break());
                }

                return p;
            }
            #endregion

            #region Métodos uso simples
            public Paragraph ParEspaco()
            {
                var pp = _props.PP(JustificationValues.Left,
                                   WordSpacing.Simples, antes: "0", depois: "0");
                return new Paragraph(pp,
                    new Run(_props.RP(WordSize.Corpo),
                        new Text("") { Space = SpaceProcessingModeValues.Preserve }));
            }

            public Paragraph ParTitulo(string txt)
            {
                var pp = _props.PP(JustificationValues.Center);
                pp.AppendChild(new ParagraphBorders(
                    new TopBorder
                    {
                        Val = BorderValues.Single,
                        Size = 6,
                        Space = 4,
                        Color = "000000"
                    }));
                var p = new Paragraph(pp);
                p.AppendChild(new Run(_props.RP(WordSize.Titulo, negrito: true),
                    new Text(txt) { Space = SpaceProcessingModeValues.Preserve }));
                return p;
            }

            public Paragraph ParCorpo(string txt)
            {
                var pp = _props.PP(JustificationValues.Both,
                                   WordSpacing.Duplo, antes: "0", depois: "80");
                pp.AppendChild(new Indentation { FirstLine = WordIndent.PrimeiraLinhaCorpo });
                var p = new Paragraph(pp);
                p.AppendChild(new Run(_props.RP(WordSize.Corpo),
                    new Text(txt) { Space = SpaceProcessingModeValues.Preserve }));
                return p;
            }

            public Paragraph ParCitacao(string txt)
            {
                var pp = _props.PP(JustificationValues.Both, WordSpacing.Simples);
                pp.AppendChild(new Indentation { Left = WordIndent.RecuoCitacao });
                var p = new Paragraph(pp);
                p.AppendChild(new Run(_props.RP(WordSize.Citacao),
                    new Text(txt) { Space = SpaceProcessingModeValues.Preserve }));
                return p;
            }

            public Paragraph ParEsquerda(string txt, string tam = WordSize.Corpo,
                                         bool negrito = false)
            {
                var p = new Paragraph(_props.PP(JustificationValues.Left, WordSpacing.Simples));
                var run = new Run(_props.RP(tam, negrito));

                var linhas = txt.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None);
                for (int i = 0; i < linhas.Length; i++)
                {
                    run.AppendChild(new Text(linhas[i]) { Space = SpaceProcessingModeValues.Preserve });
                    if (i < linhas.Length - 1)
                        run.AppendChild(new Break());
                }

                p.AppendChild(run);
                return p;
            }

            public Paragraph ParDireita(string txt, string tam = WordSize.Corpo,
                                        bool negrito = false, bool italico = false)
            {
                var p = new Paragraph(_props.PP(JustificationValues.Right, WordSpacing.Simples));
                p.AppendChild(new Run(_props.RP(tam, negrito, italico),
                    new Text(txt) { Space = SpaceProcessingModeValues.Preserve }));
                return p;
            }

            public Paragraph ParCentro(string txt, string tam = WordSize.Corpo,
                                       bool negrito = false, bool italico = false)
            {
                var p = new Paragraph(_props.PP(JustificationValues.Center, WordSpacing.Simples));
                p.AppendChild(new Run(_props.RP(tam, negrito, italico),
                    new Text(txt) { Space = SpaceProcessingModeValues.Preserve }));
                return p;
            }

            public Paragraph ParCentroSimples(string txt, string tam = WordSize.Corpo,
                                              bool negrito = false)
                => ParCentro(txt, tam, negrito);

            public Paragraph ParFimDePagina(string txt, string tam = WordSize.Corpo, bool negrito = false, bool italico = false,
                string largura = null)   // twips como string, ex: "9638" — null = automático
            {
                var pp = _props.PP(JustificationValues.Center,
                                   WordSpacing.Simples,
                                   antes: "0",
                                   depois: "0");

                pp.AppendChild(new FrameProperties
                {
                    XAlign = HorizontalAlignmentValues.Center,
                    HorizontalPosition = HorizontalAnchorValues.Margin,
                    Y = "13000",
                    YAlign = VerticalAlignmentValues.Bottom,
                    VerticalPosition = VerticalAnchorValues.Margin,
                    Width = !string.IsNullOrEmpty(largura)
                                             ? new StringValue(largura)
                                             : null,
                    Wrap = TextWrappingValues.None,
                });

                var p = new Paragraph(pp);
                p.AppendChild(new Run(
                    _props.RP(tam, negrito, italico),
                    new Text(txt) { Space = SpaceProcessingModeValues.Preserve }));

                return p;
            }

            public Paragraph ParQuebraPagina()
            {
                var pp = _props.PP(JustificationValues.Left,
                                   WordSpacing.Simples,
                                   antes: "0",
                                   depois: "0");

                var p = new Paragraph(pp);
                p.AppendChild(new Run(
                    _props.RP(WordSize.Corpo),
                    new Break { Type = BreakValues.Page }));

                return p;
            }
            #endregion

            #region Métodos mistos (múltiplos trechos)

            public Paragraph ParMisto(JustificationValues alinhamento,
                                      string espacamento,
                                      bool recuoPrimeiraLinha,
                                      params Trecho[] trechos)
            {
                var pp = _props.PP(alinhamento, espacamento, "0",
                                   alinhamento == JustificationValues.Both ? "80" : "0");

                if (recuoPrimeiraLinha)
                    pp.AppendChild(new Indentation
                    { FirstLine = WordIndent.PrimeiraLinhaCorpo });

                return AdicionarTrechos(new Paragraph(pp), trechos);
            }

            public Paragraph ParMistoCorpo(params Trecho[] trechos)
            {
                return ParMisto(JustificationValues.Both,
                                WordSpacing.Duplo,
                                recuoPrimeiraLinha: true,
                                trechos);
            }

            public Paragraph ParMistoEsquerda(params Trecho[] trechos)
            {
                return ParMisto(JustificationValues.Left,
                                WordSpacing.Simples,
                                recuoPrimeiraLinha: false,
                                trechos);
            }

            public Paragraph ParMistoCentro(params Trecho[] trechos)
            {
                return ParMisto(JustificationValues.Center,
                                WordSpacing.Simples,
                                recuoPrimeiraLinha: false,
                                trechos);
            }

            public Paragraph ParMistoDireita(params Trecho[] trechos)
            {
                return ParMisto(JustificationValues.Right,
                                WordSpacing.Simples,
                                recuoPrimeiraLinha: false,
                                trechos);
            }

            #endregion           
        }

        #endregion

        #region Métodos de extensão para facilitar a criação de tabelas — exemplo: tabela de informações duplas (rótulo + valor)
        public class TableFactory
        {
            private readonly PropsFactory _props;
            public TableFactory(PropsFactory props) { _props = props; }

            /// <summary>Tabela de duas colunas com texto simples.</summary>
            public Table InfoDuplo(string label, string valor, bool destaque)
            {
                var tbl = new Table();
                tbl.AppendChild(BuildTableProps());

                var row = new TableRow();
                row.AppendChild(BuildCell(
                    new[] { Trecho.Criar(label, WordSize.Assinatura, false, false, null, false, null, null, null) },
                    "3200"));
                row.AppendChild(BuildCell(
                    new[] { Trecho.Criar(valor, WordSize.Assinatura, true, false,
                                         destaque ? "CC0000" : null, false, null, null, null) },
                    "6438"));
                tbl.AppendChild(row);
                return tbl;
            }

            /// <summary>
            /// Tabela de duas colunas onde a célula de valor aceita trechos mistos.
            /// Útil quando parte do valor precisa de cor ou peso diferente.
            /// Exemplo:
            /// <code>
            /// tblFactory.InfoMisto("Paciente:",
            ///     Trecho.T("João Silva").Negrito(),
            ///     Trecho.T("  (em risco)").Vermelho().Italico()
            /// );
            /// </code>
            /// </summary>
            public Table InfoMisto(string label, params Trecho[] trechosValor)
            {
                var tbl = new Table();
                tbl.AppendChild(BuildTableProps());

                var row = new TableRow();
                row.AppendChild(BuildCell(
                    new[] { Trecho.Criar(label, WordSize.Assinatura, false, false, null, false, null, null, null) },
                    "3200"));
                row.AppendChild(BuildCell(trechosValor, "6438"));
                tbl.AppendChild(row);
                return tbl;
            }

            #region Helpers privados
            private TableProperties BuildTableProps()
            {
                var semBorda = new TableBorders(
                    new TopBorder { Val = BorderValues.None },
                    new BottomBorder { Val = BorderValues.None },
                    new LeftBorder { Val = BorderValues.None },
                    new RightBorder { Val = BorderValues.None },
                    new InsideHorizontalBorder { Val = BorderValues.None },
                    new InsideVerticalBorder { Val = BorderValues.None });

                var semMargem = new TableCellMarginDefault(
                    new TopMargin { Width = "0", Type = TableWidthUnitValues.Dxa },
                    new BottomMargin { Width = "0", Type = TableWidthUnitValues.Dxa },
                    new StartMargin { Width = "0", Type = TableWidthUnitValues.Dxa },
                    new EndMargin { Width = "0", Type = TableWidthUnitValues.Dxa });

                return new TableProperties(
                    new TableWidth { Width = "9638", Type = TableWidthUnitValues.Dxa },
                    new TableLook { Val = "0000" },
                    semBorda,
                    semMargem);
            }

            private TableCell BuildCell(IEnumerable<Trecho> trechos, string largura)
            {
                var cell = new TableCell(new TableCellProperties(
                    new TableCellWidth
                    {
                        Width = largura,
                        Type = TableWidthUnitValues.Dxa
                    },
                    new TableCellVerticalAlignment
                    {
                        Val = TableVerticalAlignmentValues.Top
                    },
                    new TableCellMargin(
                        new TopMargin { Width = "0", Type = TableWidthUnitValues.Dxa },
                        new BottomMargin { Width = "0", Type = TableWidthUnitValues.Dxa })));

                var pp = _props.PP(JustificationValues.Left,
                                   WordSpacing.Simples, "0", "0");
                var p = new Paragraph(pp);

                foreach (var trecho in trechos)
                {
                    var rp = _props.RPDeTrecho(trecho);
                    p.AppendChild(new Run(rp,
                        new Text(trecho.Texto ?? string.Empty)
                        { Space = SpaceProcessingModeValues.Preserve }));
                }

                cell.AppendChild(p);
                return cell;
            }
            #endregion
        }

        public class CelulaBuilder
        {
            private readonly PropsFactory _props;
            private readonly int _largura;     // twips
            private int _gridSpan = 1;
            private TableVerticalAlignmentValues _vAlign = TableVerticalAlignmentValues.Top;
            private int _margemH = 80;   // top/bottom twips
            private int _margemV = 120;  // left/right twips
            private readonly List<ParagraphoDef> _pars = new List<ParagraphoDef>();

            internal CelulaBuilder(PropsFactory props, int largura)
            {
                _props = props;
                _largura = largura;
            }

            // ── configuração da célula ───────────────────────────────────────────
            public CelulaBuilder MesclarColunas(int span) { _gridSpan = span; return this; }
            public CelulaBuilder AlinhamentoVertical(TableVerticalAlignmentValues v) { _vAlign = v; return this; }
            public CelulaBuilder Margens(int horizontalTwips, int verticalTwips)
            {
                _margemH = horizontalTwips;
                _margemV = verticalTwips;
                return this;
            }

            // ── adicionar parágrafos ─────────────────────────────────────────────

            /// <summary>
            /// Parágrafo com um ou mais trechos inline.
            /// Cada Trecho que contenha '\n' gera quebras de linha dentro do mesmo parágrafo.
            /// </summary>
            public CelulaBuilder Par(JustificationValues alinhamento, params Trecho[] trechos)
            {
                _pars.Add(new ParagraphoDef(alinhamento, trechos));
                return this;
            }

            /// <summary>
            /// NOVO: Permite adicionar vários parágrafos de uma vez em uma célula
            /// </summary>
            public CelulaBuilder Pars(params ParagraphoDef[] paragrafos)
            {
                _pars.AddRange(paragrafos);
                return this;
            }

            /// <summary>
            /// Atalho útil: adiciona vários parágrafos justificados
            /// </summary>
            public CelulaBuilder ParsJustificados(params Trecho[][] trechosPorParagrafo)
            {
                foreach (var trechos in trechosPorParagrafo)
                    _pars.Add(new ParagraphoDef(JustificationValues.Both, trechos));
                return this;
            }

            /// <summary>Atalho: parágrafo justificado.</summary>
            public CelulaBuilder ParJustificado(params Trecho[] trechos)
                => Par(JustificationValues.Both, trechos);

            /// <summary>Atalho: parágrafo centralizado.</summary>
            public CelulaBuilder ParCentro(params Trecho[] trechos)
                => Par(JustificationValues.Center, trechos);

            /// <summary>Atalho: parágrafo à esquerda.</summary>
            public CelulaBuilder ParEsquerda(params Trecho[] trechos)
                => Par(JustificationValues.Left, trechos);

            /// <summary>Atalho: parágrafo à direita.</summary>
            public CelulaBuilder ParDireita(params Trecho[] trechos)
                => Par(JustificationValues.Right, trechos);

            // ── build ────────────────────────────────────────────────────────────
            internal TableCell Build(string tamanhoPadrao)
            {
                var props = new TableCellProperties(
                    new TableCellWidth { Width = _largura.ToString(), Type = TableWidthUnitValues.Dxa },
                    new TableCellVerticalAlignment { Val = _vAlign },
                    new TableCellMargin(
                        new TopMargin { Width = _margemH.ToString(), Type = TableWidthUnitValues.Dxa },
                        new BottomMargin { Width = _margemH.ToString(), Type = TableWidthUnitValues.Dxa },
                        new StartMargin { Width = _margemV.ToString(), Type = TableWidthUnitValues.Dxa },
                        new EndMargin { Width = _margemV.ToString(), Type = TableWidthUnitValues.Dxa }
                    )
                );

                if (_gridSpan > 1)
                    props.AppendChild(new GridSpan { Val = _gridSpan });

                var cell = new TableCell(props);

                // se não foram adicionados parágrafos, Word exige ao menos um vazio
                if (_pars.Count == 0)
                {
                    cell.AppendChild(new Paragraph(
                        new ParagraphProperties(
                            new SpacingBetweenLines
                            {
                                Line = WordSpacing.Simples,
                                LineRule = LineSpacingRuleValues.Auto,
                                Before = "0",
                                After = "0"
                            })));
                    return cell;
                }

                foreach (var def in _pars)
                    cell.AppendChild(ConstruirParagrafo(def, tamanhoPadrao));

                return cell;
            }

            // ── helpers privados ─────────────────────────────────────────────────
            private Paragraph ConstruirParagrafo(ParagraphoDef def, string tamanhoPadrao)
            {
                var pp = _props.PP(def.Alinhamento, WordSpacing.Simples, "0", "0");
                var p = new Paragraph(pp);

                foreach (var trecho in def.Trechos)
                {
                    var tam = string.IsNullOrEmpty(trecho.Tamanho) ? tamanhoPadrao : trecho.Tamanho;
                    var rp = _props.RP(tam, trecho.Negrito, trecho.Italico, trecho.Cor);

                    // suporte a quebras de linha dentro do trecho (texto multiline)
                    var linhas = (trecho.Texto ?? string.Empty)
                                     .Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None);

                    for (int i = 0; i < linhas.Length; i++)
                    {
                        if (i > 0)
                            p.AppendChild(new Run(new Break()));

                        // clona RunProperties para cada run (OpenXML não permite reusar instâncias)
                        var rpClone = (RunProperties)rp.CloneNode(true);
                        p.AppendChild(new Run(rpClone,
                            new Text(linhas[i]) { Space = SpaceProcessingModeValues.Preserve }));
                    }

                    if (trecho.Quebra)
                        p.AppendChild(new Run(new Break()));
                }

                return p;
            }

            // ── tipos internos ───────────────────────────────────────────────────
            public class ParagraphoDef
            {
                public JustificationValues Alinhamento { get; private set; }
                public Trecho[] Trechos { get; private set; }

                public ParagraphoDef(JustificationValues alinhamento, Trecho[] trechos)
                {
                    Alinhamento = alinhamento;
                    Trechos = trechos;
                }
            }
        }

        public class LinhaBuilder
        {
            private readonly PropsFactory _props;
            private readonly List<CelulaBuilder> _celulas = new List<CelulaBuilder>();
            private string _tamanhoPadrao = WordSize.Titulo;

            internal LinhaBuilder(PropsFactory props) { _props = props; }

            /// <summary>Tamanho de fonte padrão para todas as células desta linha.</summary>
            public LinhaBuilder TamanhoPadrao(string tam) { _tamanhoPadrao = tam; return this; }

            /// <summary>
            /// Adiciona uma célula de <paramref name="largura"/> twips e a configura com o delegate.
            /// </summary>
            public LinhaBuilder Celula(int largura, Action<CelulaBuilder> configurar)
            {
                var builder = new CelulaBuilder(_props, largura);
                configurar(builder);
                _celulas.Add(builder);
                return this;
            }

            /// <summary>
            /// Atalho: célula com um único parágrafo justificado e trechos variáveis.
            /// </summary>
            public LinhaBuilder CelulaTexto(int largura, JustificationValues alinhamento,
                                            params Trecho[] trechos)
                => Celula(largura, c => c.Par(alinhamento, trechos));

            /// <summary>
            /// Atalho: célula com múltiplos parágrafos (o mais importante para você)
            /// </summary>
            public LinhaBuilder CelulaMultiParagrafo(int largura, JustificationValues alinhamentoPadrao,
                                                     params Trecho[][] trechosPorParagrafo)
            {
                return Celula(largura, c =>
                {
                    c.AlinhamentoVertical(TableVerticalAlignmentValues.Top);
                    foreach (var trechos in trechosPorParagrafo)
                    {
                        if (trechos != null && trechos.Length > 0)
                            c.Par(alinhamentoPadrao, trechos);
                    }
                });
            }

            internal TableRow Build()
            {
                var row = new TableRow();
                foreach (var cel in _celulas)
                    row.AppendChild(cel.Build(_tamanhoPadrao));
                return row;
            }
        }

        public class TableBuilder
        {
            private readonly PropsFactory _props;
            private int _larguraTotal = 9638;   // twips (padrão A4 com margens)
            private bool _comBorda = true;
            private BorderValues _tipoBorda = BorderValues.Single;
            private uint _tamBorda = 6;
            private string _corBorda = "000000";
            private bool _bordaInterna = true;
            private string _tamanhoPadrao = WordSize.Titulo;
            private readonly List<LinhaBuilder> _linhas = new List<LinhaBuilder>();

            public TableBuilder(PropsFactory props) { _props = props; }

            // ── configuração da tabela ───────────────────────────────────────────
            public TableBuilder LarguraTotal(int twips) { _larguraTotal = twips; return this; }
            public TableBuilder SemBorda() { _comBorda = false; return this; }
            public TableBuilder SemBordaInterna() { _bordaInterna = false; return this; }
            public TableBuilder TipoBorda(BorderValues v) { _tipoBorda = v; return this; }
            public TableBuilder TamanhoBorda(uint tam) { _tamBorda = tam; return this; }
            public TableBuilder CorBorda(string hex) { _corBorda = hex; return this; }
            public TableBuilder TamanhoPadrao(string tam) { _tamanhoPadrao = tam; return this; }

            // ── adicionar linhas ─────────────────────────────────────────────────

            /// <summary>Adiciona uma linha configurada via delegate.</summary>
            public TableBuilder Linha(Action<LinhaBuilder> configurar)
            {
                var lb = new LinhaBuilder(_props).TamanhoPadrao(_tamanhoPadrao);
                configurar(lb);
                _linhas.Add(lb);
                return this;
            }

            /// <summary>
            /// Atalho: linha de título que ocupa toda a largura (GridSpan calculado
            /// automaticamente com base no número de colunas de larguras fornecidas).
            /// </summary>
            public TableBuilder LinhaTitulo(string texto, int gridSpan = 2)
                => Linha(l => l.Celula(_larguraTotal, c => c
                    .MesclarColunas(gridSpan)
                    .AlinhamentoVertical(TableVerticalAlignmentValues.Center)
                    .ParCentro(Trecho.T(texto).Negrito())));

            /// <summary>
            /// Atalho: linha label + valor (2 colunas), mesma aparência de AdicionarLinhaSetor.
            /// </summary>
            public TableBuilder LinhaLabelValor(string label, string valor,
                                                int largLabel = 2800, int largValor = 6838)
                => Linha(l => l
                    .Celula(largLabel, c => c
                        .AlinhamentoVertical(TableVerticalAlignmentValues.Center)
                        .ParCentro(Trecho.T(label).Negrito()))
                    .Celula(largValor, c => c
                        .AlinhamentoVertical(TableVerticalAlignmentValues.Center)
                        .ParEsquerda(Trecho.T(valor))));

            /// <summary>
            /// Atalho: linha de texto completo (ocupa _larguraTotal, gridSpan = 2).
            /// Aceita vários trechos para formatação mista (negrito/normal).
            /// </summary>
            public TableBuilder LinhaTexto(JustificationValues alinhamento, params Trecho[] trechos)
                => Linha(l => l.Celula(_larguraTotal, c => c
                    .MesclarColunas(2)
                    .AlinhamentoVertical(TableVerticalAlignmentValues.Center)
                    .Par(alinhamento, trechos)));

            public TableBuilder LinhaTextoJustificado(params Trecho[] trechos)
                => LinhaTexto(JustificationValues.Both, trechos);

            public TableBuilder LinhaTextoEsquerda(params Trecho[] trechos)
                => LinhaTexto(JustificationValues.Left, trechos);

            // ── build ────────────────────────────────────────────────────────────
            public Table Build()
            {
                var tbl = new Table();
                tbl.AppendChild(BuildTableProperties());
                foreach (var lb in _linhas)
                    tbl.AppendChild(lb.Build());
                return tbl;
            }

            // ── helpers privados ─────────────────────────────────────────────────
            private TableProperties BuildTableProperties()
            {
                TableBorders bordas;

                if (_comBorda)
                {
                    var bv = new EnumValue<BorderValues>(_tipoBorda);
                    var bordaExterna = new Func<OpenXmlElement, OpenXmlElement>(x => x); // identidade

                    if (_bordaInterna)
                        bordas = new TableBorders(
                            new TopBorder { Val = bv, Size = _tamBorda, Color = _corBorda },
                            new BottomBorder { Val = bv, Size = _tamBorda, Color = _corBorda },
                            new LeftBorder { Val = bv, Size = _tamBorda, Color = _corBorda },
                            new RightBorder { Val = bv, Size = _tamBorda, Color = _corBorda },
                            new InsideHorizontalBorder { Val = bv, Size = _tamBorda, Color = _corBorda },
                            new InsideVerticalBorder { Val = bv, Size = _tamBorda, Color = _corBorda });
                    else
                        bordas = new TableBorders(
                            new TopBorder { Val = bv, Size = _tamBorda, Color = _corBorda },
                            new BottomBorder { Val = bv, Size = _tamBorda, Color = _corBorda },
                            new LeftBorder { Val = bv, Size = _tamBorda, Color = _corBorda },
                            new RightBorder { Val = bv, Size = _tamBorda, Color = _corBorda },
                            new InsideHorizontalBorder { Val = BorderValues.None },
                            new InsideVerticalBorder { Val = BorderValues.None });
                }
                else
                {
                    bordas = new TableBorders(
                        new TopBorder { Val = BorderValues.None },
                        new BottomBorder { Val = BorderValues.None },
                        new LeftBorder { Val = BorderValues.None },
                        new RightBorder { Val = BorderValues.None },
                        new InsideHorizontalBorder { Val = BorderValues.None },
                        new InsideVerticalBorder { Val = BorderValues.None });
                }

                return new TableProperties(
                    new TableWidth { Width = _larguraTotal.ToString(), Type = TableWidthUnitValues.Dxa },
                    new TableLook { Val = "0000" },
                    bordas);
            }
        }
        #endregion

        #region Metodos de extensão para manipulação do documento, como adicionar header, footer, background etc
        public class DocumentPartService
        {
            private readonly ParagraphFactory _parFactory;
            private readonly WordImageHelper _imgHelper;
            private uint _imgIdCounter = 1U;

            public DocumentPartService(ParagraphFactory parFactory,
                                       WordImageHelper imgHelper)
            {
                _parFactory = parFactory;
                _imgHelper = imgHelper;
            }

            public void AdicionarEstilos(MainDocumentPart main)
            {
                var sp = main.AddNewPart<StyleDefinitionsPart>();
                sp.Styles = new Styles(
                    new DocDefaults(
                        new RunPropertiesDefault(new RunPropertiesBaseStyle(
                            new RunFonts
                            {
                                Ascii = WordFont.Default,
                                HighAnsi = WordFont.Default
                            },
                            new FontSize { Val = WordSize.Corpo },
                            new FontSizeComplexScript { Val = WordSize.Corpo })),
                        new ParagraphPropertiesDefault(new ParagraphPropertiesBaseStyle(
                            new Justification { Val = JustificationValues.Both },
                            new SpacingBetweenLines
                            {
                                Line = WordSpacing.Duplo,
                                LineRule = LineSpacingRuleValues.Auto,
                                Before = "0",
                                After = "0"
                            }))));
                sp.Styles.Save();
            }

            public void AdicionarHeader(MainDocumentPart main, TempleteWord dados)
            {
                var headerPart = main.AddNewPart<HeaderPart>();
                var header = new Header();

                var pImg = new Paragraph(
                    CriarPP(JustificationValues.Center,
                            WordSpacing.Simples, "0", "0"));

                _imgHelper.AdicionarImagemNoParagrafo(
                    pImg, headerPart, dados.CaminhoLogoHeader, 80, 80);
                header.AppendChild(pImg);

                foreach (var linha in dados.CabecalhoCliente
                             .Split(new[] { "\r\n", "\r", "\n" },
                                    StringSplitOptions.None))
                {
                    if (!string.IsNullOrEmpty(linha))
                        header.AppendChild(
                            _parFactory.ParCentroSimples(linha, WordSize.Rodape,
                                                         negrito: true));
                }

                headerPart.Header = header;
                headerPart.Header.Save();

                var sectPr = ObteroSectPr(main);
                sectPr.AppendChild(new HeaderReference
                {
                    Type = HeaderFooterValues.Default,
                    Id = main.GetIdOfPart(headerPart)
                });
            }

            public void AdicionarFooter(MainDocumentPart main, TempleteWord dados)
            {
                var footerPart = main.AddNewPart<FooterPart>();
                var footer = new Footer();

                footer.AppendChild(BuildLinhaRodape());
                footer.AppendChild(BuildTabelaRodape(footerPart));

                footerPart.Footer = footer;
                footerPart.Footer.Save();

                var sectPr = ObteroSectPr(main);
                sectPr.AppendChild(new FooterReference
                {
                    Type = HeaderFooterValues.Default,
                    Id = main.GetIdOfPart(footerPart)
                });
            }

            public void AdicionarBackground(MainDocumentPart main, TempleteWord dados)
            {
                if (string.IsNullOrEmpty(dados.BackgroundImage) || !File.Exists(dados.BackgroundImage)) return;

                int largPx = 794, altPx = 1123;

                // Reutiliza o HeaderPart existente ou cria um novo
                HeaderPart headerPart;
                string relId;

                var headerRef = main.Document.Body
                    .Elements<SectionProperties>().FirstOrDefault()
                    ?.Elements<HeaderReference>()
                    .FirstOrDefault(h => h.Type == HeaderFooterValues.Default);

                if (headerRef != null)
                {
                    headerPart = (HeaderPart)main.GetPartById(headerRef.Id);
                }
                else
                {
                    headerPart = main.AddNewPart<HeaderPart>();
                    headerPart.Header = new Header();
                    headerPart.Header.Save();

                    relId = main.GetIdOfPart(headerPart);
                    var sectPr = ObteroSectPr(main);
                    sectPr.AppendChild(new HeaderReference
                    {
                        Type = HeaderFooterValues.Default,
                        Id = relId
                    });
                }

                // Adiciona a imagem como part no HeaderPart
                byte[] bytes = File.ReadAllBytes(dados.BackgroundImage);

                // Ajusta altura proporcionalmente se possível
                try
                {
                    using (var bmp = new Bitmap(new MemoryStream(bytes)))
                        if (bmp.Width > 0)
                            altPx = (int)(largPx * ((double)bmp.Height / bmp.Width));
                }
                catch { }

                long cx = largPx * 9525L;
                long cy = altPx * 9525L;

                var ext = Path.GetExtension(dados.BackgroundImage).ToLowerInvariant();
                var partType = (ext == ".jpg" || ext == ".jpeg") ? ImagePartType.Jpeg : ImagePartType.Png;

                var imgPart = headerPart.AddImagePart(partType);
                using (var ms = new MemoryStream(bytes)) imgPart.FeedData(ms);
                string imgRelId = headerPart.GetIdOfPart(imgPart);

                uint id = _imgIdCounter++;

                // Parágrafo que contém a imagem ancorada atrás do texto
                var pBg = new Paragraph(
                    CriarPP(JustificationValues.Center, WordSpacing.Simples, "0", "0"));

                pBg.AppendChild(new Run(
                    new Drawing(
                        new DocumentFormat.OpenXml.Drawing.Wordprocessing.Anchor(
                            new DocumentFormat.OpenXml.Drawing.Wordprocessing.SimplePosition { X = 0L, Y = 0L },
                            new DocumentFormat.OpenXml.Drawing.Wordprocessing.HorizontalPosition(
                                new DocumentFormat.OpenXml.Drawing.Wordprocessing.HorizontalAlignment
                                { Text = "center" })
                            { RelativeFrom = HorizontalRelativePositionValues.Page },
                            new DocumentFormat.OpenXml.Drawing.Wordprocessing.VerticalPosition(
                                new DocumentFormat.OpenXml.Drawing.Wordprocessing.VerticalAlignment
                                { Text = "center" })
                            { RelativeFrom = VerticalRelativePositionValues.Page },
                            new DocumentFormat.OpenXml.Drawing.Wordprocessing.Extent { Cx = cx, Cy = cy },
                            new DocumentFormat.OpenXml.Drawing.Wordprocessing.EffectExtent
                            { LeftEdge = 0L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 0L },
                            new DocumentFormat.OpenXml.Drawing.Wordprocessing.WrapNone(),
                            new DocumentFormat.OpenXml.Drawing.Wordprocessing.DocProperties
                            { Id = id, Name = "BgImg" + id },
                            new DocumentFormat.OpenXml.Drawing.Wordprocessing
                                .NonVisualGraphicFrameDrawingProperties(
                                new DocumentFormat.OpenXml.Drawing.GraphicFrameLocks
                                { NoChangeAspect = true }),
                            new DocumentFormat.OpenXml.Drawing.Graphic(
                                new DocumentFormat.OpenXml.Drawing.GraphicData(
                                    new DocumentFormat.OpenXml.Drawing.Pictures.Picture(
                                        new DocumentFormat.OpenXml.Drawing.Pictures.NonVisualPictureProperties(
                                            new DocumentFormat.OpenXml.Drawing.Pictures.NonVisualDrawingProperties
                                            { Id = 0U, Name = "bg" },
                                            new DocumentFormat.OpenXml.Drawing.Pictures
                                                .NonVisualPictureDrawingProperties()),
                                        new DocumentFormat.OpenXml.Drawing.Pictures.BlipFill(
                                            new DocumentFormat.OpenXml.Drawing.Blip
                                            {
                                                Embed = imgRelId,
                                                CompressionState =
                                                    DocumentFormat.OpenXml.Drawing.BlipCompressionValues.Print
                                            },
                                            new DocumentFormat.OpenXml.Drawing.Stretch(
                                                new DocumentFormat.OpenXml.Drawing.FillRectangle())),
                                        new DocumentFormat.OpenXml.Drawing.Pictures.ShapeProperties(
                                            new DocumentFormat.OpenXml.Drawing.Transform2D(
                                                new DocumentFormat.OpenXml.Drawing.Offset { X = 0L, Y = 0L },
                                                new DocumentFormat.OpenXml.Drawing.Extents { Cx = cx, Cy = cy }),
                                            new DocumentFormat.OpenXml.Drawing.PresetGeometry(
                                                new DocumentFormat.OpenXml.Drawing.AdjustValueList())
                                            { Preset = DocumentFormat.OpenXml.Drawing.ShapeTypeValues.Rectangle }))
                                )
                                { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })
                        )
                        {
                            DistanceFromTop = 0U,
                            DistanceFromBottom = 0U,
                            DistanceFromLeft = 0U,
                            DistanceFromRight = 0U,
                            SimplePos = false,
                            RelativeHeight = 1U,      // z-order baixo = atrás do texto
                            BehindDoc = true,    // ← chave: imagem atrás do conteúdo
                            Locked = false,
                            LayoutInCell = true,
                            AllowOverlap = true
                        }
                    )
                ));

                headerPart.Header.InsertAt(pBg, 0); // insere antes de qualquer outro conteúdo
                headerPart.Header.Save();
            }

            public void AppendAoBody(MainDocumentPart main, OpenXmlElement elemento)
            {
                var body = main.Document.Body;
                var sectPr = body.Elements<SectionProperties>().FirstOrDefault();

                if (sectPr != null)
                    body.InsertBefore(elemento, sectPr);
                else
                    body.AppendChild(elemento);
            }

            public void AplicarLayoutA4(MainDocumentPart main)
            {
                var body = main.Document.Body;
                var sectPr = body.Elements<SectionProperties>().FirstOrDefault();

                if (sectPr == null)
                {
                    sectPr = new SectionProperties();
                    body.AppendChild(sectPr);
                }
                else
                {
                    sectPr.Remove();
                    body.AppendChild(sectPr);
                }

                foreach (var ps in sectPr.Elements<PageSize>().ToList()) ps.Remove();
                foreach (var pm in sectPr.Elements<PageMargin>().ToList()) pm.Remove();

                sectPr.PrependChild(new PageMargin
                {
                    Top = 1701,
                    Bottom = 1134,
                    Left = 1701U,
                    Right = 1134U,
                    Header = 10U,
                    Footer = 567U
                });
                sectPr.PrependChild(new PageSize
                {
                    Width = 11906U,
                    Height = 16838U
                });
            }

            // helpers privados de rodapé omitidos por brevidade — mesma lógica do original

            private SectionProperties ObteroSectPr(MainDocumentPart main)
            {
                var body = main.Document.Body;
                var sectPr = body.Elements<SectionProperties>().FirstOrDefault();
                if (sectPr != null) return sectPr;
                sectPr = new SectionProperties();
                body.AppendChild(sectPr);
                return sectPr;
            }

            private Paragraph BuildLinhaRodape()
            {
                var pp = new PropsFactory().PP(
                    JustificationValues.Left,
                    WordSpacing.Rodape,
                    antes: "0",
                    depois: "0");

                pp.AppendChild(new Indentation { Left = "0", Right = "0", FirstLine = "0" });
                pp.AppendChild(new ParagraphBorders(
                    new TopBorder
                    {
                        Val = BorderValues.Single,
                        Size = 6,
                        Space = 2,
                        Color = "000000"
                    }));

                return new Paragraph(pp);
            }

            private Table BuildTabelaRodape(FooterPart footerPart)
            {
                var tbl = new Table();
                tbl.AppendChild(new TableProperties(
                    new TableWidth { Width = "5000", Type = TableWidthUnitValues.Pct },
                    new TableBorders(
                        new TopBorder { Val = BorderValues.None },
                        new BottomBorder { Val = BorderValues.None },
                        new LeftBorder { Val = BorderValues.None },
                        new RightBorder { Val = BorderValues.None },
                        new InsideHorizontalBorder { Val = BorderValues.None },
                        new InsideVerticalBorder { Val = BorderValues.None }),
                    new TableCellMarginDefault(
                        new TopMargin { Width = "0", Type = TableWidthUnitValues.Dxa },
                        new BottomMargin { Width = "0", Type = TableWidthUnitValues.Dxa },
                        new StartMargin { Width = "0", Type = TableWidthUnitValues.Dxa },
                        new EndMargin { Width = "0", Type = TableWidthUnitValues.Dxa })));

                var row = new TableRow();

                // ── Célula esquerda: logo (reservada, sem imagem por ora) ──────────
                var cLogo = new TableCell(new TableCellProperties(
                    new TableCellWidth
                    {
                        Width = "1500",
                        Type = TableWidthUnitValues.Dxa
                    },
                    new TableCellVerticalAlignment
                    {
                        Val = TableVerticalAlignmentValues.Center
                    },
                    new TableCellMargin(
                        new LeftMargin { Width = "0", Type = TableWidthUnitValues.Dxa },
                        new RightMargin { Width = "0", Type = TableWidthUnitValues.Dxa })));

                cLogo.AppendChild(
                    new Paragraph(
                        new PropsFactory().PP(
                            JustificationValues.Left,
                            WordSpacing.Simples,
                            antes: "0",
                            depois: "0")));

                row.AppendChild(cLogo);

                // ── Célula central: endereço ───────────────────────────────────────
                var cEnd = new TableCell(new TableCellProperties(
                    new TableCellWidth
                    {
                        Width = "6638",
                        Type = TableWidthUnitValues.Dxa
                    },
                    new TableCellVerticalAlignment
                    {
                        Val = TableVerticalAlignmentValues.Center
                    },
                    new TableCellMargin(
                        new LeftMargin { Width = "0", Type = TableWidthUnitValues.Dxa },
                        new RightMargin { Width = "0", Type = TableWidthUnitValues.Dxa },
                        new TopMargin { Width = "0", Type = TableWidthUnitValues.Dxa },
                        new BottomMargin { Width = "0", Type = TableWidthUnitValues.Dxa })));

                cEnd.AppendChild(_parFactory.ParCentroSimples(
                    "Tv. Lomas Valentina, n° 2190 - Marco - Belém - PA - CEP: 66093-677",
                    WordSize.Rodape));
                cEnd.AppendChild(_parFactory.ParCentroSimples(
                    "Fone: (91) 4006-4398   E-mail: ndj.sespa2@gmail.com",
                    WordSize.Rodape));

                row.AppendChild(cEnd);

                // ── Célula direita: equilíbrio visual ─────────────────────────────
                var cEsp = new TableCell(new TableCellProperties(
                    new TableCellWidth
                    {
                        Width = "1500",
                        Type = TableWidthUnitValues.Dxa
                    }));

                cEsp.AppendChild(
                    new Paragraph(
                        new PropsFactory().PP(
                            JustificationValues.Left,
                            WordSpacing.Simples)));

                row.AppendChild(cEsp);

                tbl.AppendChild(row);
                return tbl;
            }
            public ParagraphProperties CriarPP(JustificationValues j,
                                                 string e, string a, string d)
                => new PropsFactory().PP(j, e, a, d);

            public RunProperties CriarRP(string e,
                                                 bool n, bool i, string c)
                => new PropsFactory().RP(e, n, i, c);
        }

        #endregion
    }
}