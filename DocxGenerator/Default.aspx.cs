using System;
using System.Web;
using Documentos;

public partial class _Default : System.Web.UI.Page 
{
    protected void Page_Load(object sender, EventArgs e)
    {

    }

    protected void btnRelatorioBloqueioJudicial_OnClick(object sender, EventArgs e)
    {
        byte[] bytes = RelatorioModeloDespachoManifestacaoViabilidade.GerarDocx();
        Exportar(bytes, "RelatorioModeloDespachoBloqueioJudicial");

    }

    protected void btnManifestacaoViabilidade_OnClick(object sender, EventArgs e)
    {
        byte[] bytes = RelatorioModeloDespachoManifestacaoViabilidade.GerarDocx();
        Exportar(bytes, "RelatorioModeloDespachoManifestacaoViabilidade");
    }

    protected void btnAcaoJudicial_OnClick(object sender, EventArgs e)
    {
        byte[] bytes = RelatorioModeloDespachoManifestacaoViabilidade.GerarDocx();
        Exportar(bytes, "RelatorioModeloChefeGab");
    }

    private void Exportar(byte[] bytes, string relatorio)
    {
        Response.Clear();
        Response.Buffer = true;
        Response.ContentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
        Response.AddHeader("content-disposition", "attachment; filename=\"RelatorioModeloDespachoManifestacaoViabilidade.docx\"");
        Response.AddHeader("content-length", bytes.Length.ToString());
        Response.BinaryWrite(bytes);
        Response.Flush();

        // Encerra o pipeline sem ThreadAbortException
        HttpContext.Current.ApplicationInstance.CompleteRequest();
    }
}