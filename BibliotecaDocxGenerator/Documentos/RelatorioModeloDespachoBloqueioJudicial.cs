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
    public class RelatorioModeloDespachoBloqueioJudicial
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
            body.AppendChild(tf.InfoDuplo("Assunto:", "Bloqueio Judicial", false));
            body.AppendChild(tf.InfoDuplo("Referência:", "Ação Judicial", true));
            body.AppendChild(tf.InfoDuplo("Objeto:", "MEDICAMENTO - PAMELOR, MUSCULARE 10MG", true));
            body.AppendChild(tf.InfoDuplo("CNS n°:", cns, true));
            body.AppendChild(tf.InfoDuplo("CPF n°:", cpf, true));
            body.AppendChild(pf.ParEspaco());

            body.AppendChild(pf.ParCorpo("À _____________________________________________________________"));
            body.AppendChild(pf.ParEspaco());
            body.AppendChild(pf.ParCorpo("Prezado(a) ___________________________________________________,"));
            body.AppendChild(pf.ParEspaco());           

            body.AppendChild(pf.ParMistoCorpo(
             Trecho.T("Honrado em cumprimenta-lo, cito informar Vossa Excelência que autos em epigrafe tratam da "),
             Trecho.T("ACAO COMINATORIA DE OBRIGACAO DE FAZER COM PEDIDO EXPRESSO DE TUTELA DE URGENCIA").Negrito(),
             Trecho.T(" ajuizada pelo "),
             Trecho.T("SOPHIA REZENDE").Negrito(),
             Trecho.T(" substituto processual de "),
             Trecho.T("PATRICK RODRIGUES").Negrito(),
             Trecho.T(" em face do ESTADO DO PARA.")
             ));

            body.AppendChild(pf.ParCorpo("Considerando a DECISÃO JUDICIAL proferida nos autos(Seq.________– Fls.______), " +
                "cumpre observar o disposto no dispositivo abaixo transcrito, cuja determinação expressa impõe as providências ora analisadas: "));

            body.AppendChild(pf.ParCitacao("Aenean nec libero eget mi ultricies ultrices. Morbi eu velit malesuada, maximus ante nec, congue ante. Aenean viverra lorem ac nisl facilisis, id venenatis nisl posuere. Vivamus ornare lectus a semper facilisis. Nullam sodales vel augue ac commodo. Integer gravida justo felis, eu congue libero condimentum et. Mauris et maximus ex. Proin a ipsum vitae elit commodo faucibus sed non purus. Vestibulum pretium pellentesque lectus sed pellentesque. Maecenas dignissim sed nisi et accumsan. Phasellus viverra elit et leo commodo, nec ornare odio mattis. Phasellus ut orci non erat viverra suscipit. Lorem ipsum dolor sit amet, consectetur adipiscing elit. Donec porttitor convallis felis, tempor volutpat sapien semper ut."));
            body.AppendChild(pf.ParCorpo("Considerando a Comunicação nº______________/____________-PGE-GAB, exarada pela PROCURADORIA - GERAL DO ESTADO DO PARÁ - PGE(Seq.__________________) onde informa:"));

            string CampoEscrever = "___________________________________________________";
            body.AppendChild(pf.ParDireita(CampoEscrever));
            body.AppendChild(pf.ParDireita(CampoEscrever));
            body.AppendChild(pf.ParDireita(CampoEscrever));

            body.AppendChild(pf.ParCorpo("Insta informar a Vossa Excelência que a demanda de cumprimento tramita no ambito da Secretaria de Estado de Saude Publica - SESPA, por meio do Processo Administrativo "
                                    + $"Eletronico no 1010101-01.0101.0.10.1010. Em consulta avançada, verificou-se que o referido PAE encontra-se registrado no "
                                    + "fluxo __________________ > __________________ > __________________, conforme se observa do documento acostado ao sequencial __________________ ."));

            body.AppendChild(pf.ParCorpo("Diante do exposto, de ordem da Senhora Coordenadora do Núcleo de Demandas Judiciais NDJ/SESPA, submeto os autos à elevada consideração de Vossa Excelência, para ciência e necessárias "
                                    + "deliberações acerca das informações e sequenciais anteriormente mencionados."));

            body.AppendChild(pf.ParCorpo("Informo que este documento contém dados pessoais/sensíveis, sendo de uso restrito à finalidade institucional, nos termos da Lei Federal de Acesso à Informação – LAI nº 12.527/2011, e Lei Federal de Geral de Proteção de Dados  - LGPD n° 13.709/2018."));

            body.AppendChild(pf.ParCorpo("Reiteramos, por fim, nossa inteira disposição para prestar quaisquer informações complementares que se fizerem necessárias."));

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
