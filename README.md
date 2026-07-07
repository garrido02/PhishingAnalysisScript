# 🎣 Phishing Campaign Analytics

Uma ferramenta automatizada em Python concebida para extrair, processar e analisar dados de campanhas de simulação de phishing. Esta aplicação agrupa vários ficheiros CSV gerados por plataformas de phishing e compila um relatório global e detalhado em Excel.

## ✨ Funcionalidades

* **Interface Gráfica (GUI):** Interface simples e intuitiva construída com `Tkinter`, incluindo uma barra de progresso em tempo real. Não é necessário interagir com o terminal.
* **Processamento em Lote:** Seleciona uma pasta com dezenas ou centenas de ficheiros CSV, e a ferramenta processa todos automaticamente.
* **Métricas Inteligentes:** Separa campanhas por tipo (Credenciais, Anexos, ou Ambos) e analisa taxas de conversão (Aberturas, Cliques, Inserção de Dados).
* **Análise de Utilizadores (User Analytics):** Mapeia o comportamento individual dos utilizadores, identificando o número de interações e a sua distribuição por localização.
* **Exportação em Excel:** Gera um ficheiro `phishing_results.xlsx` dividido em três páginas claras:
    * `Raw Data`: Todos os dados agregados.
    * `Analytics`: Volumes da campanha e taxas de conversão (%).
    * `Analysis - Users`: Distribuição de cliques e comportamento granular dos utilizadores.
* **Executável Automático (CI/CD):** Graças ao GitHub Actions, a cada atualização de código, a ferramenta é testada e compilada automaticamente num ficheiro `.exe` pronto a usar.

---

## 🚀 Como Utilizar (Método Recomendado)

Não precisas de instalar Python ou dependências para usar a ferramenta no Windows.

1. Navega até ao separador **[Releases](../../releases)** deste repositório, no lado direito da página.
2. Descarrega a versão mais recente do ficheiro `PhishingAnalytics.exe`.
3. Faz duplo clique no executável.
4. Seleciona a pasta onde tens os teus ficheiros CSV da campanha.
5. Aguarda que a barra de progresso conclua. O relatório final será guardado numa nova pasta chamada `[nome-da-tua-pasta]_results`.

---

## 🛠️ Como Correr a Partir do Código Fonte (Para Desenvolvedores)

Se quiseres modificar o código ou correr num ambiente local (Windows, macOS, ou Linux):

### 1. Pré-requisitos
* Python 3.10 ou superior.
* Gerenciador de pacotes `pip`.
* O pacote `tkinter` instalado no sistema (dependendo do SO).

### 2. Instalação
Clona o repositório e instala as bibliotecas necessárias:

```bash
git clone [https://github.com/garrido02/PhihingAnalysisScript.git](https://github.com/garrido02/PhihingAnalysisScript.git)
cd PhihingAnalysisScript
pip install pandas openpyxl pytest flake8
