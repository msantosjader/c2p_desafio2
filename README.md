# c2p_desafio2

## Descrição
**Desafio 2 - Processo Seletivo C2P Estágio Backoffice**
Script em python que faça webscrapping das Taxas de Títulos Públicos – ANBIMA;


## Requisitos
- Python 3.9+

## Instalação

```bash
git clone https://github.com/msantosjader/c2p_desafio2.git
cd c2p_desafio2
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## Uso

```bash
# Para obter as taxas dos último dia útil disponível
python -m msec

# Opcionalmente informe a data desejada
python -m msec 06/11/2025
```

> Precisa ser um dia útil e a ANBIMA mantém histórico apenas dos 5 últimos dias úteis

> Serão exibidas mensagens de erro caso estes requisitos não sejam atendidos.


## Estrutura do Projeto

```
c2p_desafio2/
├── msec.python				# Código para execução
├── requirements.txt		# Dependências
├── relatorios/				# Diretórios onde os arquivos serão gerados
│   └── msec_[data].xlsx	# Como as planilhas serão geradas
└── .modelo.xlsx			# Modelo para layout dos relatórios
```
