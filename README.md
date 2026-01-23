# LabelIt 🏷️

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Lazarus](https://img.shields.io/badge/Lazarus-Free_Pascal-blue.svg)](https://www.lazarus-ide.org/)
[![Antigravity](https://img.shields.io/badge/antigravity-enabled-brightgreen.svg)](https://antigravity.google/)

> *"É simples parecer complicado, o complicado é parecer simples."*

> [!IMPORTANT]
> **Estado do Projeto:** Esta aplicação encontra-se atualmente em fase de desenvolvimento ativo da sua **primeira versão Beta**. Algumas funcionalidades podem sofrer alterações significativas.

O **LabelIt** é uma plataforma de design e impressão de etiquetas robusta e intuitiva, desenvolvida em **Lazarus (Free Pascal)**. Focada na produtividade e na facilidade de uso, permite criar etiquetas complexas com integração dinâmica de dados de forma simplificada.

---

## ✨ Funcionalidades

- **🚀 Designer Visual "Drag-and-Drop":** Interface fluida para posicionamento e redimensionamento de objetos em tempo real.
- **📦 Objetos Suportados:**
  - **Texto:** Suporte total a fontes, alinhamentos e cores.
  - **Imagens:** Inserção de gráficos estáticos ou via ligação de dados.
  - **Códigos de Barras:** Integração com `lazbarcodes` (EAN13, Code128, QR Code, etc.).
  - **Grelhas Dinâmicas:** Edição avançada de tabelas com fusão de células.
- **🔗 Data Binding Abrangente:** Ligue os seus campos de etiquetas a diversas fontes de dados, incluindo:
  - **Bases de Dados SQL:** SQL Server, MySQL, SQLite, entre outras.
  - **Folhas de Cálculo:** Ficheiros Excel e outros formatos de dados estruturados.
- **📄 Formato .lit:** Sistema de ficheiros otimizado (ZIP) que agrupa definições, consultas e metadados.
- **🛠️ Grelha de Propriedades:** Edição precisa de todos os atributos dos objetos, incluindo edição em lote para multi-seleção.

---

## 🛠️ Tecnologias e Dependências

A aplicação foi construída utilizando o ecossistema **Object Pascal**:

- **IDE:** [Lazarus](https://www.lazarus-ide.org/) (recomendado 2.2.0 ou superior)
- **Compilador:** Free Pascal (FPC)
- **Bibliotecas Requeridas:**
  - `lazbarcodes` (disponível via Online Package Manager)

---

## 🚀 Como Compilar

1. **Instale o Lazarus:** Descarregue e instale a versão mais recente em [lazarus-ide.org](https://www.lazarus-ide.org/).
2. **Instale as Dependências:** No Lazarus, vá a `Package` -> `Online Package Manager` e procure/instale por `lazbarcodes`.
3. **Abra o Projeto:** No menu `Project`, selecione `Open Project` e escolha o ficheiro `.lpi` na pasta `Source`.
4. **Compile:** Pressione `F9` para compilar e executar.

---

## 🎯 Filosofia de Design (KISS)

O **LabelIt** segue rigorosamente a filosofia **KISS (Keep It Simple and Stable)**. Acreditamos que uma ferramenta poderosa não precisa de ser complicada. Cada linha de código é escrita a pensar na estabilidade e na clareza para o utilizador final.

---

## 📄 Licença

Este projeto está licenciado sob a licença MIT - veja o ficheiro `LICENSE` para mais detalhes.

---

<p align="center">
  Desenvolvido com ❤️ por José Tomás Rodrigues
</p>
