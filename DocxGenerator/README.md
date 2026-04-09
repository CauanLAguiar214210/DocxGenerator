# Sistema de Geração de Documentos Judiciais

Projeto WebForms em .NET Framework 3.5 para geração de documentos judiciais em formato DOCX e PDF.

## 🚀 Funcionalidades

- **Geração de Documentos**: Crie documentos judiciais profissionais com formatação padrão
- **Múltiplos Formatos**: Exporte em DOCX (compatível com Word) ou PDF
- **Tipos de Documentos**: Suporte para Petição Inicial, Contestação, Recursos, Sentenças e mais
- **Interface Intuitiva**: Formulário completo com todos os campos necessários
- **Validação de Dados**: Verificação automática de campos obrigatórios

## 📋 Estrutura do Projeto

```
WebSite1/
├── Default.aspx              # Página inicial com apresentação
├── GeradorDocumentos.aspx    # Formulário de geração de documentos
├── Models/
│   └── DocumentoJudicial.cs  # Modelo de dados para documentos
├── Utils/
│   └── DocumentGenerator.cs  # Utilitário para geração de documentos
└── Web.config               # Configurações do aplicativo
```

## 🛠️ Tecnologias Utilizadas

- **.NET Framework 3.5**: Framework principal
- **ASP.NET WebForms**: Framework de desenvolvimento web
- **RTF Generation**: Geração de documentos compatíveis com Microsoft Word
- **HTML to PDF**: Conversão para formato PDF universal

## 📖 Como Usar

1. **Acessar o Sistema**: Abra o projeto no Visual Studio e execute
2. **Página Inicial**: Acesse `Default.aspx` para ver a apresentação
3. **Gerar Documento**: Clique em "Acessar Gerador de Documentos"
4. **Preencher Formulário**: Complete todos os campos obrigatórios
5. **Escolher Formato**: Selecione DOCX ou PDF
6. **Gerar**: Clique em "Gerar Documento" para download

## 📄 Campos do Formulário

### Dados Básicos
- **Número do Processo**: Identificação única do processo
- **Tipo de Documento**: Selecione entre 8 tipos diferentes
- **Data**: Automática com data atual

### Partes Envolvidas
- **Requerente/Autor**: Nome da parte que inicia o processo
- **Requerido/Réu**: Nome da parte respondente
- **Advogados**: Nome e OAB dos representantes legais

### Dados do Processo
- **Juiz**: Nome do magistrado responsável
- **Vara**: Vara específica
- **Comarca**: Local da jurisdição
- **Valor da Causa**: Valor monetário envolvido

### Conteúdo
- **Objeto**: Descrição do que se busca no processo
- **Conteúdo**: Texto completo do documento judicial

## 🔧 Formatos de Saída

### DOCX (Formato RTF)
- **Compatibilidade**: Microsoft Word, LibreOffice, Google Docs
- **Formatação**: Mantém formatação profissional
- **Edição**: Permite edição posterior

### PDF
- **Universal**: Visualização em qualquer dispositivo
- **Seguro**: Formato não editável
- **Profissional**: Apresentação padronizada

## 📝 Tipos de Documentos Suportados

1. **Petição Inicial**: Documento que inicia um processo
2. **Contestação**: Resposta do réu à petição inicial
3. **Recurso de Apelação**: Recurso contra sentenças
4. **Embargos de Declaração**: Esclarecimento de pontos obscuros
5. **Sentença**: Decisão do juiz sobre o mérito
6. **Despacho**: Decisões interlocutórias simples
7. **Decisão Interlocutória**: Decisões durante o processo
8. **Acórdão**: Decisão colegiada em tribunais

## 🚀 Instalação e Configuração

### Pré-requisitos
- Visual Studio 2008 ou superior
- .NET Framework 3.5 instalado
- IIS (Internet Information Services)

### Passos
1. Clone ou baixe o projeto
2. Abra no Visual Studio
3. Configure o IIS para a aplicação
4. Execute o projeto

## 🔒 Segurança

- **Dados Locais**: Todas as informações ficam no servidor local
- **Sem Dependências Externas**: Funciona offline
- **Validação**: Validação de entrada de dados
- **Controle de Acesso**: Configuração via web.config

## 🎯 Benefícios

- **Produtividade**: Gere documentos em segundos
- **Padronização**: Formatação consistente em todos os documentos
- **Flexibilidade**: Escolha o formato ideal para cada necessidade
- **Economia**: Reduza tempo e recursos na criação de documentos
- **Profissionalismo**: Documentos com aparência profissional

## 📞 Suporte

Para dúvidas ou sugestões de melhoria, verifique a documentação técnica do projeto ou consulte um desenvolvedor .NET especializado.

---

**Desenvolvido com ❤️ usando ASP.NET WebForms e .NET Framework 3.5**
