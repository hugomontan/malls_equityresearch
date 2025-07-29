# Consolidador de Reports - Automatização de Dados

Sistema automatizado para consolidação de dados financeiros das empresas Allos, Iguatemi e Multiplan.

## 📋 Descrição

Este projeto automatiza o processamento e consolidação de planilhas Excel contendo dados financeiros das empresas de shopping centers. O sistema:

- Processa planilhas das empresas Allos, Iguatemi e Multiplan
- Consolida dados em uma única planilha Excel
- Gera gráficos automáticos
- Atualiza dados de inflação (IPCA/IGPM)
- Interface gráfica intuitiva

## 🚀 Instalação

### Pré-requisitos
- Python 3.8 ou superior
- pip (gerenciador de pacotes Python)

### Passos para instalação

1. **Clone o repositório**
   ```bash
   git clone https://github.com/seu-usuario/malls_automated.git
   cd malls_automated
   ```

2. **Instale as dependências**
   ```bash
   pip install -r requirements.txt
   ```

3. **Execute o script de build (opcional)**
   ```bash
   python build_executable.py
   ```

## 📁 Estrutura do Projeto

```
malls_automated/
├── main_optimized.py          # Script principal
├── inflation.py               # Script de atualização de inflação
├── build_executable.py        # Script para gerar executável
├── requirements.txt           # Dependências do projeto
├── main_optimized.spec        # Configuração PyInstaller
├── reports/                   # Pasta para arquivos de entrada
├── dist/                      # Executável gerado
└── build/                     # Arquivos de build
```

## 🎯 Como Usar

### Opção 1: Executável (Recomendado)
1. Execute `Consolidador_Reports.exe`
2. Selecione os 3 arquivos Excel das empresas
3. Aguarde o processamento
4. O arquivo consolidado será aberto automaticamente

### Opção 2: Script Python
1. Coloque os arquivos Excel na pasta `reports/`
2. Execute: `python main_optimized.py`
3. Siga as instruções na interface

### Arquivos Necessários
- `Allos Planilha 1T25.xlsx`
- `Iguatemi Planilha 1T25.xlsx`
- `Multiplan Planilha 1T25.xlsx`

## 🔧 Configuração

### Arquivo .spec
Para gerar executável sem terminal, edite `main_optimized.spec`:
```python
console=False,  # Mude de True para False
```

### Pasta Reports
A pasta `reports/` deve estar na raiz do projeto junto com o executável.

## 📊 Saídas

- **Consolidado.xlsx**: Planilha consolidada na área de trabalho
- **graficos_consolidado.png**: Gráficos gerados automaticamente
- **erro_consolidador.log**: Log de erros (se houver)

## 🛠️ Desenvolvimento

### Dependências Principais
- `openpyxl`: Manipulação de arquivos Excel
- `pandas`: Processamento de dados
- `matplotlib`: Geração de gráficos
- `tkinter`: Interface gráfica

### Gerar Executável
```bash
python build_executable.py
```

## 📝 Logs

O sistema gera logs em:
- `erro_consolidador.log`: Log principal
- Console: Mensagens de debug

## 🔄 Atualizações

### Dados de Inflação
O sistema atualiza automaticamente os dados de IPCA/IGPM via `inflation.py`.

### Versão
v2.0 - Automatização de Reportse

Para problemas ou dúvidas:
1. Verifique os logs em `erro_consolidador.log`
2. Confirme se os arquivos Excel estão na pasta `reports/`
3. Verifique se as dependências estão instaladas

---

**Nota**: Certifique-se de que a pasta `reports/` existe na raiz do projeto antes de executar o sistema. 
