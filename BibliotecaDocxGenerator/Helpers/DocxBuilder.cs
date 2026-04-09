using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using static Helpers.OfficeWordHelper;


namespace Helpers
{
    public class DocxBuilder
    {
        private readonly ParagraphFactory _parFactory;
        private readonly TableFactory _tblFactory;
        private readonly DocumentPartService _partService;
        private readonly string _basePath;

        public DocxBuilder(ParagraphFactory parFactory,
                           TableFactory tblFactory,
                           DocumentPartService partService,
                           string basePath = null)
        {
            _parFactory = parFactory;
            _tblFactory = tblFactory;
            _partService = partService;
            _basePath = basePath ?? AppDomain.CurrentDomain.BaseDirectory;
        }

        // Monta tudo que é igual em todo relatório
        // O corpo é passado como Action — compatível com .NET 3.5
        public byte[] Gerar(string titulo, Action<Body, MainDocumentPart, ParagraphFactory, TableFactory> construirCorpo)
        {
            using (var ms = new MemoryStream())
            {
                using (var doc = WordprocessingDocument.Create(
                    ms, WordprocessingDocumentType.Document, true))
                {
                    var templete = new TempleteWord()
                    {
                        Titulo = titulo,
                        CabecalhoCliente = "Governo do Pará DSV DSV - Secretaria de Saúde do Estado do Pará Gabinete do Secretário DSV"
                    };

                    templete.CaminhoLogoHeader = Path.Combine(_basePath, "mj-header.jpg");
                    templete.CaminhoLogoRodape = Path.Combine(_basePath, "estrela-para.png");
                    templete.BackgroundImage = Path.Combine(_basePath, "background-para.png");

                    var main = doc.AddMainDocumentPart();
                    main.Document = new Document(new Body());

                    _partService.AdicionarEstilos(main);
                    _partService.AdicionarHeader(main, templete);
                    _partService.AdicionarFooter(main, templete);
                    _partService.AdicionarBackground(main, templete);

                    var body = main.Document.Body;
                    body.AppendChild(_parFactory.ParTitulo(templete.Titulo));
                    construirCorpo(body, main, _parFactory, _tblFactory);

                    _partService.AplicarLayoutA4(main);

                    Action<OpenXmlElement> appendSeguro =
                        elemento => _partService.AppendAoBody(main, elemento);

                    main.Document.Save();
                }
                return ms.ToArray();
            }
        }
    }
}