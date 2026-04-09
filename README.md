# DocxGenerator - Sistema de Geração de Documentos DOCX

Sistema para geração de documentos DOCX nativos desenvolvido em .NET Framework 3.5 com interface WebForms, otimizado para criar documentos oficiais com formatação profissional.

## Visão Geral
O DocxGenerator é uma solução completa para geração de documentos Microsoft Word (.docx) nativos, desenvolvida especificamente para ambientes corporativos que utilizam .NET Framework 3.5. O sistema oferece uma interface web intuitiva via ASP.NET WebForms para criação de documentos padronizados com cabeçalhos, rodapés e formatação profissional.

## Características Principais

### Funcionalidades
- **Geração de DOCX Nativos**: Cria documentos Microsoft Word compatíveis
- **Interface WebForms**: Interface amigável via ASP.NET WebForms
- **Múltiplos Modelos**: Suporte para diversos tipos de documentos (Bloqueio Judicial, Manifestação de Viabilidade, Ação Judicial)
- **Formatação Profissional**: Cabeçalhos, rodapés e fundos personalizados
- **Exportação Direta**: Download automático dos documentos gerados

### Modelos Disponíveis
- **Folha de Despacho**: Documentos para bloqueio judicial
- **Manifestação de Viabilidade**: Relatórios de análise processual
- **Ação Judicial**: Documentos para procedimentos judiciais
- **Relatório Chefia**: Documentos administrativos

## Arquitetura

### Estrutura do Projeto
```
DocxGenerator/
|
|-- BibliotecaDocxGenerator/          # Biblioteca principal de geração DOCX
|   |-- Helpers/                       # Classes utilitárias
|   |   |-- DocxBuilder.cs            # Builder principal para documentos
|   |   |-- OfficeWordHelper.cs       # Helper extenso para manipulação Word
|   |   |-- WordImageHelper.cs        # Manipulação de imagens
|   |-- Documentos/                    # Modelos de documentos específicos
|   |   |-- RelatorioModeloDespachoBloqueioJudicial.cs
|   |   |-- RelatorioModeloDespachoManifestacaoViabilidade.cs
|   |   |-- RelatorioModeloChefeGab.cs
|   |-- Recursos/                      # Imagens e assets
|       |-- mj-header.jpg             # Logo do cabeçalho
|       |-- estrela-para.png          # Logo do rodapé
|       |-- background-para.png       # Imagem de fundo
|
|-- DocxGenerator/                     # Aplicação Web WebForms
|   |-- Default.aspx                  # Interface principal
|   |-- Default.aspx.cs               # Code-behind da interface
|   |-- Web.config                    # Configurações da aplicação
|   |-- Bin/                          # Assemblies compilados
|
|-- DocxBuilder.sln                   # Solution Visual Studio
```

## Tecnologias Utilizadas

### Backend
- **.NET Framework 3.5**: Framework principal de desenvolvimento
- **C#**: Linguagem de programação
- **DocumentFormat.OpenXml v2.20.0**: Biblioteca para manipulação de documentos Office
- **System.IO**: Manipulação de arquivos e streams
- **System.Web**: Funcionalidades web ASP.NET

### Frontend
- **ASP.NET WebForms**: Framework de apresentação web
- **HTML5/CSS3**: Estrutura e estilização da interface
- **JavaScript**: Interações client-side (se aplicável)

## Pré-requisitos

### Ambiente de Desenvolvimento
- **Visual Studio 2008+** ou **Visual Studio 2019+** com suporte a .NET Framework 3.5
- **.NET Framework 3.5** instalado
- **IIS** (Internet Information Services) para hospedagem
- **Windows Server** (para produção) ou **Windows 10/11** (para desenvolvimento)

### Dependências
- `DocumentFormat.OpenXml` v2.20.0
- Referências ao `System.Web` e assemblies relacionados

## Instalação e Configuração

### 1. Clonar o Repositório
```bash
git clone [URL-DO-REPOSITORIO]
cd DocxGenerator
```

### 2. Abrir no Visual Studio
1. Abra o arquivo `DocxBuilder.sln`
2. Restaure os pacotes NuGet (se necessário)
3. Compile a solução

### 3. Configurar o IIS
1. Crie um novo site no IIS
2. Aponte para a pasta `DocxGenerator`
3. Configure o Application Pool para .NET Framework 3.5
4. Garanta permissões de leitura/escrita na pasta

### 4. Configurar Recursos
Verifique se as imagens de recursos estão presentes na pasta `BibliotecaDocxGenerator`:
- `mj-header.jpg`
- `estrela-para.png`
- `background-para.png`

## Como Usar

### Interface Web
1. Acesse a aplicação via navegador (ex: `http://localhost/DocxGenerator`)
2. Escolha o tipo de documento desejado:
   - **Bloqueio Judicial**: Gera despachos para bloqueios judiciais
   - **Manifestação de Viabilidade**: Cria relatórios de viabilidade
   - **Ação Judicial**: Produz documentos para ações judiciais
3. Clique no botão correspondente para gerar e baixar o documento

### Programático
```csharp
// Exemplo de geração de documento
byte[] documento = RelatorioModeloDespachoBloqueioJudicial.GerarDocx();
// Salvar ou enviar o documento gerado
```

## Personalização

### Adicionar Novos Modelos
1. Crie uma nova classe em `BibliotecaDocxGenerator/Documentos/`
2. Herde ou siga o padrão das classes existentes
3. Implemente o método `GerarDocx()` estático
4. Adicione o botão correspondente na interface WebForms

### Customizar Layout
- **Cabeçalho/Rodapé**: Modifique a classe `TempleteWord` em `DocxBuilder.cs`
- **Estilos**: Edite os métodos em `OfficeWordHelper.cs`
- **Imagens**: Substitua os arquivos de imagem na pasta raiz

## Exemplos de Uso

### Geração de Documento Judicial
```csharp
// Gera um despacho de bloqueio judicial
var bytes = RelatorioModeloDespachoBloqueioJudicial.GerarDocx();
Response.Clear();
Response.ContentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
Response.AddHeader("content-disposition", "attachment; filename=\"despacho_bloqueio.docx\"");
Response.BinaryWrite(bytes);
```

### Criação de Documento Customizado
```csharp
var builder = new DocxBuilder(parFactory, tblFactory, partService);
var documento = builder.Gerar("Título Customizado",
    (body, main, pf, tf) => {
        // Construção do corpo do documento
        body.AppendChild(pf.ParTexto("Conteúdo do documento"));
    });
```

## Estrutura de Classes Principais

### DocxBuilder
Classe principal que orquestra a geração de documentos:
- Configura estrutura básica do documento
- Aplica cabeçalhos, rodapés e fundos
- Gerencia o fluxo de criação

### OfficeWordHelper
Utilitário extenso com factories para:
- `ParagraphFactory`: Criação de parágrafos estilizados
- `TableFactory`: Geração de tabelas formatadas
- `PropsFactory`: Propriedades de formatação
- `DocumentPartService`: Manipulação de partes do documento

### Modelos de Documento
Classes específicas em `Documentos/` que implementam:
- Lógica de negócio para cada tipo de documento
- Estrutura de conteúdo específica
- Dados e formatação padronizados

## Troubleshooting

### Problemas Comuns

**Erro: "Could not load file or assembly DocumentFormat.OpenXml"**
- Verifique se o pacote NuGet está instalado
- Confirme a versão compatível com .NET 3.5

**Erro: "Access denied" ao gerar documentos**
- Verifique permissões da pasta de execução
- Garanta acesso de escrita para o usuário do IIS

**Documentos gerados em branco**
- Confirme se as imagens de recursos existem
- Verifique os paths no método `DocxBuilder.Gerar()`

### Debug
- Ative o modo debug no Web.config: `<compilation debug="true">`
- Use o Visual Studio para depurar o code-behind
- Verifique os logs do IIS para erros HTTP

## Contribuição

### Para contribuir:
1. Fork o projeto
2. Crie uma branch para sua feature
3. Mantenha a compatibilidade com .NET Framework 3.5
4. Teste em ambiente WebForms
5. Submeta pull request

### Padrões de Código
- Siga as convenções de nomenclatura C#
- Mantenha a estrutura existente de Helpers/Factories
- Documente métodos públicos com XML comments

## Licença

[Adicionar informações de licença aqui]

## Suporte

Para suporte técnico ou dúvidas:
- [Email de contato]
- [Issues no repositório]
- [Documentação adicional]

## Histórico de Versões

### v1.0.0
- Versão inicial com geração básica de DOCX
- Interface WebForms funcional
- Três modelos de documentos implementados
- Suporte a .NET Framework 3.5

---

**Nota**: Este projeto foi desenvolvido especificamente para .NET Framework 3.5 para garantir compatibilidade com sistemas legados corporativos. Para novas implementações, considere atualizar para versões mais recentes do .NET.
