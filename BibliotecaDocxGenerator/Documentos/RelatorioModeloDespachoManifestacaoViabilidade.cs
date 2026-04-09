using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Web;

using static Helpers.OfficeWordHelper;
using Helpers;

namespace Documentos
{
    public class RelatorioModeloDespachoManifestacaoViabilidade
    {
        public static byte[] GerarDocx()
        {
            var props = new PropsFactory();
            var parFactory = new ParagraphFactory(props);
            var tblFactory = new TableFactory(props);
            var imgHelper = new WordImageHelper();
            var partSvc = new DocumentPartService(parFactory, imgHelper);
            var builder = new DocxBuilder(parFactory, tblFactory, partSvc);

            return builder.Gerar("FOLHA DE DESPACHO",
                (body, main, pf, tf) => ConstruirCorpo(body, main, pf, tf));
        }

        private static void ConstruirCorpo(Body body, MainDocumentPart main,
                             ParagraphFactory pf, TableFactory tf)
        {
            string cns = "943935451780008";
            string cpf = "25593832412";
            string nomeAutor = "BERNARDO BUENO";
            string paciente = "NATHAN BRUM";

            // ── INFORMAÇÕES DO PROCESSO ───────────────────────────────
            body.AppendChild(tf.InfoDuplo("Processo PAE:", "NÃO INFORMADO", false));
            body.AppendChild(tf.InfoDuplo("Processo Judicial n°:", "1010101-01.0101.0.10.1010", false));
            body.AppendChild(tf.InfoDuplo("Paciente/Beneficiário(a):", paciente, false));
            body.AppendChild(tf.InfoDuplo("Assunto:", "Manifestação de Viabilidade", false));
            body.AppendChild(tf.InfoDuplo("Referência:", "Ação Judicial", true));
            body.AppendChild(tf.InfoDuplo("Objeto:", "MEDICAMENTO - PAMELOR, MUSCULARE 10MG", true));
            body.AppendChild(tf.InfoDuplo("CNS n°:", cns, true));
            body.AppendChild(tf.InfoDuplo("CPF n°:", cpf, true));
            body.AppendChild(pf.ParEspaco());

            // ── DESTINATÁRIO ──────────────────────────────────────────
            body.AppendChild(pf.ParEsquerda("Ao Gabinete do Secretário de Estado de Saúde - GABS,"));
            body.AppendChild(pf.ParEspaco());
            body.AppendChild(pf.ParEsquerda("Senhor Secretário,"));
            body.AppendChild(pf.ParEspaco());

            // ── CORPO DO TEXTO ────────────────────────────────────────
            body.AppendChild(pf.ParMistoCorpo(
             Trecho.T("Com os nossos cordiais cumprimentos, informamos a Vossa Excelência que o Núcleo de Demandas Judiciais - NDJ recebeu mandado de intimação, referente à Ação Judicial ajuizada em prol de "),
             Trecho.T(nomeAutor).Negrito(),
             Trecho.T(" objetivando "),
             Trecho.T("MEDICAMENTO - PAMELOR, MUSCULARE 10MG").Negrito(),
             Trecho.T(" , pedido este que ainda não foi apreciado pelo douto Juízo, uma vez que se vislumbrou a necessidade de manifestação prévia dos Entes Públicos Requeridos ESTADO DO PARA.")
             ));

            body.AppendChild(pf.ParCitacao("Praesent maximus molestie velit sed fermentum. Quisque consequat eget ligula eget varius. Donec in ligula id odio sagittis consequat. Proin sollicitudin sed odio ac dignissim. Morbi nec eros libero. Aenean eget quam iaculis, eleifend massa in, commodo turpis. Cras in porttitor urna, a rhoncus eros. Nam non congue sem. Sed eros mi, tempor et pretium nec, condimentum vitae est."));
            body.AppendChild(pf.ParCorpo("Sendo assim e, em razão do disposto acima, encaminham-se os autos a V. Exa., para conhecimento da decisão proferida e deliberação junto ao setor técnico quanto à elaboração de manifestação técnica no sentido de demonstrar a viabilidade desta Secretaria de Saúde ao pretendido. " +
                "Ademais, a manifestação técnica deverá conter em seu bojo dentre todas as informações técnicas acerca do requerido, a possibilidade de cumprimento da ordem judicial pelo Estado do Pará, por intermédio desta Secretaria de Saúde, bem como as dificuldades em cumpri-las, caso haja condenação em desfavor do Estado do Pará. "));

            body.AppendChild(pf.ParCorpo("Informo que este documento contém dados pessoais/sensíveis, sendo de uso restrito à finalidade institucional, nos termos da Lei Federal de Acesso à Informação – LAI nº 12.527/2011, e Lei Federal de Geral de Proteção de Dados  - LGPD n° 13.709/2018."));

            body.AppendChild(pf.ParCorpo("Sem mais para o momento, colocamo-nos à disposição para eventuais esclarecimentos adicionais que se fizerem necessários."));
            body.AppendChild(pf.ParCorpo("Respeitosamente, "));

            // ── ASSINATURA ────────────────────────────────────────────
            //Dados Usuario
            var data = DateTime.Now;
            body.AppendChild(pf.ParDireita(
                $"Belém-PA, {data.Day.ToString()} de {data.ToString("MMMM", new CultureInfo("pt-BR"))} de {data.Year.ToString()}"
                , "22", negrito: true));

            body.AppendChild(pf.ParDireita("Elaborado por", "22", italico: true));
            body.AppendChild(pf.ParDireita("ADMINISTRADOR"));
            body.AppendChild(pf.ParDireita("Analista Judicial NDJ/SESPA", "22"));
            body.AppendChild(pf.ParEspaco());

            //Assinatura
            body.AppendChild(pf.ParCentro("Adrianne Costa Alves", "22", negrito: true));
            body.AppendChild(pf.ParCentro("Coordenadora do Núcleo de Demandas Judiciais", "22"));
            body.AppendChild(pf.ParCentro("NDJ/SESPA", "22"));
        }
    }
}
