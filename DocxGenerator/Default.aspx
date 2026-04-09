<%@ Page Language="C#" AutoEventWireup="true"  CodeFile="Default.aspx.cs" Inherits="_Default" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
    <title>Sistema de Documentos Judiciais</title>
    <style type="text/css">
        body { font-family: Arial, sans-serif; margin: 0; padding: 0; background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); min-height: 100vh; }
        .container { max-width: 800px; margin: 0 auto; padding: 40px 20px; }
        .header { text-align: center; color: white; margin-bottom: 40px; }
        .header h1 { font-size: 3em; margin-bottom: 10px; text-shadow: 2px 2px 4px rgba(0,0,0,0.3); }
        .header p { font-size: 1.2em; opacity: 0.9; }
        .card { background: white; border-radius: 10px; padding: 30px; box-shadow: 0 10px 30px rgba(0,0,0,0.2); margin-bottom: 30px; transition: transform 0.3s; }
        .card:hover { transform: translateY(-5px); }
        .card h2 { color: #2c3e50; margin-bottom: 15px; }
        .card p { color: #7f8c8d; line-height: 1.6; margin-bottom: 20px; }
        .btn { display: inline-block; background: #3498db; color: white; padding: 12px 30px; text-decoration: none; border-radius: 5px; transition: background 0.3s; }
        .btn:hover { background: #2980b9; }
        .features { display: grid; grid-template-columns: repeat(auto-fit, minmax(250px, 1fr)); gap: 20px; margin-top: 30px; }
        .feature { background: rgba(255,255,255,0.1); padding: 20px; border-radius: 8px; color: white; text-align: center; }
        .feature h3 { margin-bottom: 10px; }
        .feature-icon { font-size: 2em; margin-bottom: 10px; }
    </style>
</head>
<body>
    <form id="form1" runat="server">
    <div class="container">
        <div class="header">
            <h1>⚖️ Sistema de Documentos Judiciais</h1>
            <p>Gere documentos jurídicos profissionais em formato DOCX ou PDF</p>
        </div>

        <div class="card">
            <h2>📝 Gerador de Documentos</h2>
            <p>Crie documentos judiciais completos com todas as informações necessárias. Preencha o formulário com os dados do processo, partes envolvidas e conteúdo do documento.</p>
            <p><strong>Recursos disponíveis:</strong></p>
            <ul>
                <li>Múltiplos tipos de documentos (Petição, Contestação, Recursos, etc.)</li>
                <li>Formatação profissional padrão jurídico</li>
                <li>Exportação em DOCX (compatível com Word) ou PDF</li>
                <li>Interface intuitiva e fácil de usar</li>
            </ul>
            <asp:Button class="btn" id="btnRelatorioBloqueioJudicial" runat="server" Text="Bloqueio Judicial" 
                onclick="btnRelatorioBloqueioJudicial_OnClick" />
            <asp:Button class="btn" id="btnManifestacaoViabilidade" runat="server" Text="Manifestação de Viabilidade" 
                onclick="btnManifestacaoViabilidade_OnClick" />
            <asp:Button class="btn" id="btnAcaoJudicial" runat="server" Text="Ação Judicial" 
                onclick="btnAcaoJudicial_OnClick" />

        </div>

        <div class="features">
            <div class="feature">
                <div class="feature-icon">📄</div>
                <h3>Documentos Profissionais</h3>
                <p>Formatação padrão para documentos jurídicos</p>
            </div>
            <div class="feature">
                <div class="feature-icon">💾</div>
                <h3>Múltiplos Formatos</h3>
                <p>Exporte em DOCX ou PDF conforme sua necessidade</p>
            </div>
            <div class="feature">
                <div class="feature-icon">⚡</div>
                <h3>Rápido e Eficiente</h3>
                <p>Gere documentos em segundos com poucos cliques</p>
            </div>
            <div class="feature">
                <div class="feature-icon">🔒</div>
                <h3>Seguro</h3>
                <p>Sistema local com seus dados sempre protegidos</p>
            </div>
        </div>
        <br />
        <div class="card">
            <h2>🔧 Tecnologias Utilizadas</h2>
            <p>Este sistema foi desenvolvido com as seguintes tecnologias:</p>
            <ul>
                <li><strong>.NET Framework 3.5</strong> - Framework robusto e confiável</li>
                <li><strong>ASP.NET WebForms</strong> - Framework web maduro e testado</li>
                <li><strong>RTF Generation</strong> - Geração de documentos compatíveis com Word</li>
                <li><strong>HTML to PDF</strong> - Conversão para formato PDF universal</li>
            </ul>
        </div>
    </div>
    </form>
</body>
</html>
