# Guia Técnico do Projeto LabelIt

## 1. Introdução

O **LabelIt** é uma aplicação de design e impressão de etiquetas desenvolvida em **Lazarus (Free Pascal)**. O objetivo principal é fornecer uma ferramenta intuitiva para a criação de etiquetas personalizadas, permitindo a inclusão de textos, imagens, códigos de barras e tabelas, com suporte a ligação de dados externos (Data Binding).

A filosofia de desenvolvimento segue o princípio **KISS (Keep It Simple and Stable)**, priorizando a simplicidade de uso e a estabilidade do código.

---

## 2. Arquitetura do Sistema

### 2.1. Tecnologias Utilizadas

- **Linguagem:** Object Pascal (Modo Delphi).
- **IDE:** Lazarus.
- **Compilador:** Free Pascal (FPC).
- **Bibliotecas Principais:**
  - `lazbarcodes`: Para geração de códigos de barras (EAN13, Code128, QR Code).
  - `Zipper`: Para compressão e persistência de ficheiros `.lit`.
  - `Graphics`: Para renderização visual no designer.

### 2.2. Estrutura de Ficheiros

- `uMain.pas`: Formulário principal da aplicação e gestão de menus/ficheiros.
- `uEtiqueta.pas`: O "coração" do editor, contendo a lógica do designer visual, gestão de eventos, grelha de propriedades e definições de janelas (`TTabSettings`).
- `uLabelObjects.pas`: Definição das classes de objetos da etiqueta (`TLabelItem` e derivadas).
- `uLabelIt.pas`: Estruturas de dados globais e lógica de persistência (Save/Load).
- `uDadosExternos.pas`: Gestão de ligações a bases de dados (SQL Server, MySQL, SQLite) e ficheiros de dados.

### 2.3. Padrões de Código

O projeto segue rigorosamente as boas práticas de programação em Pascal:

- **Nomenclatura:** Prefixos `T` para tipos, `F` para campos privados, `L` para variáveis locais, `A` para argumentos e `C` para constantes.
- **Encapsulamento:** Uso extensivo de propriedades e métodos para manipulação de estado.
- **Manutenibilidade:** Separação clara entre a interface (`interface`) e a implementação (`implementation`).

---

## 3. Modelo de Objetos (`uLabelObjects.pas`)

Todos os elementos visíveis na etiqueta derivam da classe base `TLabelItem`.

### 3.1. Hierarquia de Classes

- **`TLabelItem`**: Classe abstrata que define propriedades comuns como `Left`, `Top`, `Width`, `Height`, `Visible`, `FieldName` e `RecordOffset`.
  - **`TLabelText`**: Representa blocos de texto com suporte a fontes e alinhamento.
  - **`TLabelImage`**: Permite a inserção de imagens (estáticas ou via Data Binding).
  - **`TLabelBarcode`**: Gera códigos de barras dinâmicos.
  - **`TLabelGrid`**: Uma estrutura de tabela complexa com suporte a:
    - Fusão de células adjacentes (Merge/Span).
    - Cor de fundo individual por célula.
    - Formatação de bordas (Top, Right, Bottom, Left) com visibilidade e cores independentes.

### 3.2. Gestão de Objetos (`TLabelManager`)

A classe `TLabelManager` é responsável por manter a lista de objetos ativos, gerir a seleção múltipla, duplicação e operações de ordenação (Z-Order).

---

## 4. Persistência de Dados

As etiquetas são guardadas no formato proprietário `.lit`, que é um arquivo comprimido (ZIP). A aplicação garante a integridade dos dados através de:

- **Encoding:** Suporte total a UTF-8 para garantir a correta representação de caracteres especiais (ã, ç, á, etc.).
- **Deteção de Corrupção:** Mecanismo inteligente que deteta inconsistências entre `TipoDadosExternos` e os ficheiros configurados (`SQL.r`, `SQLQuery.sql`), procedendo à correção automática quando necessário.

### 4.1. Estrutura Interna do Ficheiro `.lit`

O ficheiro `.lit` é um arquivo ZIP que organiza os dados em níveis: a raiz para dados globais de ligação e subdiretórios para cada modelo/tipo de etiqueta presente no projeto.

| Caminho / Ficheiro        | Descrição                | Conteúdo Principal
|---------------------------|--------------------------|-------------------
| **`E.0/`**, **`E.1/`**... | **Diretórios de Modelo** | Contentores para diferentes definições de etiquetas no mesmo ficheiro.
| `E.N/Main.r`              | Configurações do Modelo  | Impressora, Papel, Dimensões, Margens, Orientação e Layout.
| `E.N/Obj.0.r`...          | Objetos da Etiqueta      | Dados binários serializados de cada componente dentro daquele modelo.
| `SQL.r`                   | Configurações de Ligação | Tipo de dados, Versão (MySQL), Servidor, Base de Dados, Utilizador, Password e Parâmetros.
| `SQLQuery.sql`            | Consulta de Dados        | Comando SQL ou metadados da folha de cálculo.

---

## 5. Funcionalidades Avançadas

### 5.1. Designer Visual

- **Designer Per Tab:** Cada aba aberta possui as suas próprias configurações de zoom, orientação e estado visual (`TTabSettings`).
- **Multi-Seleção:** Edição em lote na grelha de propriedades de múltiplos objetos selecionados.
- **Clipboard:** Sistema de Cut/Copy/Paste que preserva todas as propriedades dos objetos, incluindo ligações a dados.

### 5.2. Data Binding (Ligação de Dados)

Os objetos podem ser ligados dinamicamente:

- **`FieldName`**: Nome do campo na base de dados.
- **`RecordOffset`**: Permite que diferentes objetos apontem para registos subsequentes na mesma página de impressão.
- **Inteligência:** Atualização automática da visualização no designer ao navegar pelos dados externos.

### 5.3. Grelha de Propriedades (`sgProperties`)

A grelha foi otimizada para produtividade:

- **Agrupamento:** Propriedades organizadas em secções expansíveis (➖Fonte, ➖Célula, ➖Bordas).
- **Hierarquia de Prioridade:** Na edição de células de grelha, a prioridade de exibição é: Objeto Filho > Célula > Tabela (Grid).
- **Editores Especializados:** Diálogos integrados para Cores, Fontes e Ficheiros.

### 5.4. Automação de Layout

O sistema deteta automaticamente as dimensões e o layout com base no nome do papel da impressora:

- Identificação de papel contínuo vs. etiquetas fixas.
- Cálculo automático do número de colunas (ex: "54mm x 3" define 3 colunas).
- Ajuste dinâmico das margens e orientação.

---

## 6. Fluxo de Impressão

A impressão utiliza definições precisas em milímetros, convertendo coordenadas lógicas para o DPI da impressora. O motor de renderização garante que o que é visto no designer é exatamente o que é impresso (WYSIWYG).

---

## 7. Notas de Manutenção e Filosofia
>
> *"É simples parecer complicado, o complicado é parecer simples."*

Ao realizar manutenção, observe:

1. **Consistência:** Novos campos devem ser integrados em `SaveToWriter`, `LoadFromReader` e `Clone`.
2. **Alertas do Compilador:** Warnings de conversão de string são monitorizados, mas permitidos em contextos seguros de UTF-8. Para novos hints, utilize `{%H-}` se o parâmetro não for utilizado.
3. **Performance:** Operações de desenho de grelha e códigos de barras devem ser eficientes para manter a fluidez do designer em zoom elevado.

---
*Relatório atualizado em 22 de Janeiro de 2026.*
