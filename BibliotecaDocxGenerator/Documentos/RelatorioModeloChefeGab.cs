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
    public class RelatorioModeloChefeGab
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

        private static void ConstruirCorpo(Body body, MainDocumentPart main, ParagraphFactory pf, TableFactory tf)
        {
            string cns = "943935451780008";
            string cpf = "25593832412";
            string nomeAutor = "BERNARDO BUENO";
            string paciente = "NATHAN BRUM";

            // ── INFORMAÇÕES DO PROCESSO ───────────────────────────────
            body.AppendChild(tf.InfoDuplo("Processo PAE:", "NÃO INFORMADO", false));
            body.AppendChild(tf.InfoDuplo("Processo Judicial n°:", "1010101-01.0101.0.10.1010", false));
            body.AppendChild(tf.InfoDuplo("Paciente/Beneficiário(a):", paciente, false));
            body.AppendChild(tf.InfoDuplo("Assunto:", "Cumprimento de Decisão Judicial", false));
            body.AppendChild(tf.InfoDuplo("Referência:", "Ação Judicial", true));
            body.AppendChild(tf.InfoDuplo("Objeto:", "MEDICAMENTO - PAMELOR, MUSCULARE 10MG", true));
            body.AppendChild(tf.InfoDuplo("CNS n°:", cns, true));
            body.AppendChild(tf.InfoDuplo("CPF n°:", cpf, true));
            body.AppendChild(pf.ParEspaco());

            // ── DESTINATÁRIO ──────────────────────────────────────────
            body.AppendChild(pf.ParCorpo("À (ao)_______________________________________________________"));
            body.AppendChild(pf.ParEspaco());
            body.AppendChild(pf.ParCorpo("Prezado(a) ___________________________________________________,"));
            body.AppendChild(pf.ParEspaco());

            // ── CORPO DO TEXTO ────────────────────────────────────────
            body.AppendChild(pf.ParMistoCorpo(
              Trecho.T("Trata - se de "),
              Trecho.T("ACAO COMINATORIA DE OBRIGACAO DE FAZER COM PEDIDO EXPRESSO DE TUTELA DE URGENCIA").Negrito(),
              Trecho.T(" ajuizada por "),
              Trecho.T("SOPHIA REZENDE").Negrito(),
              Trecho.T(" em face do ESTADO DO PARÁ")
              ));

            body.AppendChild(pf.ParCorpo("No feito, o MM. Juízo decidiu da seguinte forma:"));

            //body.AppendChild(pf.ParCitacao(_dj.ProcessoJudicial.DemandaJudicialLista.OrderBy(x => x.DataCadastro).FirstOrDefault().DescricaoDecisao));
            body.AppendChild(pf.ParCitacao("Aenean nec libero eget mi ultricies ultrices. Morbi eu velit malesuada, maximus ante nec, congue ante. Aenean viverra lorem ac nisl facilisis, id venenatis nisl posuere. Vivamus ornare lectus a semper facilisis. Nullam sodales vel augue ac commodo. Integer gravida justo felis, eu congue libero condimentum et. Mauris et maximus ex. Proin a ipsum vitae elit commodo faucibus sed non purus. Vestibulum pretium pellentesque lectus sed pellentesque. Maecenas dignissim sed nisi et accumsan. Phasellus viverra elit et leo commodo, nec ornare odio mattis. Phasellus ut orci non erat viverra suscipit. Lorem ipsum dolor sit amet, consectetur adipiscing elit. Donec porttitor convallis felis, tempor volutpat sapien semper ut."));

            body.AppendChild(pf.ParMistoCorpo(
             Trecho.T("No ensejo, informa-se que o processo acima seguiu ao cumprimento no PAE "),
             Trecho.T("1010101-01.0101.0.10.1010").Negrito(),
             Trecho.T(" com o objeto: "),
             Trecho.T("MEDICAMENTO - PAMELOR, MUSCULARE 10MG").Negrito(),
             Trecho.T(" e que encontra-se  no _____________________.")
             ));

            body.AppendChild(pf.ParMistoCorpo(
             Trecho.T("Contudo, no decorrer da instrução foi peticionado que a paciente sofre de "),
             Trecho.T("Não Informado").Negrito(),
             Trecho.T(" precisa de "),
             Trecho.T("MEDICAMENTO - PAMELOR, MUSCULARE 10MG").Negrito(),
             Trecho.T(", neste sentido, foi deferida nova decisão, conforma abaixo:")
             ));

            body.AppendChild(pf.ParCitacao("Proin auctor ornare consequat. Sed iaculis faucibus elit vitae porta. Donec ultrices elit at dolor finibus ullamcorper. Fusce tristique ultrices leo ac consequat. In quis feugiat libero, malesuada lobortis elit. Sed id euismod sapien. Sed posuere, mauris vel efficitur semper, felis magna dignissim massa, ut elementum mauris libero sit amet massa. Vestibulum hendrerit, massa vel sodales tempus, felis augue tincidunt nunc, ac bibendum ante erat sed nulla. Mauris rutrum mi nec elit eleifend sollicitudin. Integer ut auctor ex, vel auctor leo. Suspendisse ut elit ut ligula euismod rutrum a et diam. Duis bibendum turpis eget purus scelerisque, nec sodales odio efficitur. Fusce lobortis egestas magna, vitae bibendum urna pellentesque ut."));

            body.AppendChild(pf.ParCorpo("Sendo assim, e em razão do exposto acima, encaminho-vos os autos para ciência da decisão proferida e providências sequenciais cabíveis no tocante ao cumprimento de Sentença, " +
                                    "anexando aos autos documentos comprobatórios de cumprimento, para fins de prestar informação à Procuradoria Geral do Estado, para viabilizar a defesa do Estado em juízo, " +
                                    "ressalta-se, que o processo principal está no (a)____________ em fase ____________”. "));

            body.AppendChild(pf.ParCorpo("Informo que este documento contém dados pessoais/sensíveis, sendo de uso restrito à finalidade institucional, nos termos da Lei Federal de Acesso à Informação – LAI nº 12.527/2011, e Lei Federal de Geral de Proteção de Dados  - LGPD n° 13.709/2018."));

            body.AppendChild(pf.ParCorpo("Sem mais para o momento, colocando-nos à disposição para eventual esclarecimento adicional que se fizer necessário. "));

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
