[README.md](https://github.com/user-attachments/files/35117379/README.md)
# Monitor de Indicadores BCB

Automação em Python que acompanha indicadores econômicos brasileiros. A cada execução, o script consulta a API pública SGS do Banco Central do Brasil e gera um painel visual com cards, gráfico e resumo automático.

## O que o projeto faz

- **Dólar PTAX:** gráfico de linha dos últimos 30 dias (a janela é recalculada a cada execução) e variação diária.
- **Selic:** taxa diária (% a.d.), mensal (% a.m.) e anual (% a.a.), usando as séries oficiais do BCB, sem cálculo aproximado.
- **IPCA:** variação mensal e acumulado em 12 meses, também pelas séries oficiais.
- **Histórico:** grava os valores coletados em um banco SQLite, sem duplicar registros.
- **Relatório:** gera uma planilha Excel com o resumo do PTAX e o histórico.

## Saídas geradas

| Arquivo | Descrição |
|---|---|
| `painel_indicadores.html` | Painel com cards, gráfico PTAX e resumo automático |
| `monitor_financeiro.db` | Banco SQLite com a tabela `historico_indicadores` |
| `relatorio_cotacoes.xlsx` | Abas `Resumo_PTAX_30d` e `Historico` |

No Colab, os arquivos são salvos em `/content`. Fora do Colab, são salvos na pasta atual.

## Séries utilizadas (SGS/BCB)

| Indicador | Série SGS | Unidade |
|---|---|---|
| Dólar (venda) | 1 | R$ |
| Selic diária | 11 | % a.d. |
| Selic acumulada no mês | 4390 | % a.m. |
| Selic anualizada (base 252) | 1178 | % a.a. |
| IPCA variação mensal | 433 | % a.m. |
| IPCA acumulado em 12 meses | 13522 | % 12m |

A série 1 é divulgada pelo BCB como "Taxa de câmbio - Livre - Dólar americano (venda) - diário".

Os códigos de série e a unidade de cada indicador ficam registrados no banco, junto com a data de referência e a data da execução.

## Como executar

### No Google Colab

1. Crie um notebook e cole o conteúdo de `monitor_indicadores_bcb.py` em uma célula.
2. Execute a célula. O painel aparece na própria célula e os arquivos são salvos em `/content`.
3. Baixe os arquivos se quiser guardá-los: o Colab apaga `/content` quando a sessão termina. Para acumular histórico entre execuções, monte o Google Drive e aponte `BASE_DIR` para uma pasta nele.

### Localmente

```bash
pip install requests pandas openpyxl
python monitor_indicadores_bcb.py
```

Depois, abra `painel_indicadores.html` no navegador.

## Dependências

- Python 3.9+
- `requests`
- `pandas`
- `openpyxl`

## Configuração

No início do script:

- `DIAS_GRAFICO`: janela do gráfico do dólar, em dias corridos (padrão: 30).
- `SERIES`: dicionário com os códigos SGS e as unidades de cada indicador.
- `BASE_DIR`: pasta onde os arquivos são gravados.

## Comportamento em caso de falha

- A API é consultada com até 3 tentativas por requisição.
- Se a Selic ou o IPCA estiverem indisponíveis, o card mostra "N/D" e o painel continua sendo gerado. A mensagem de aviso aparece no console.
- Se o dólar não puder ser consultado, o script é interrompido com uma mensagem de erro.

## Limitações

- Selic diária e Selic mensal/anual têm referências diferentes: a diária é de um dia útil e as outras se referem a um mês ou a um ano. O painel exibe a referência de cada valor.
- O IPCA é divulgado mensalmente, então muda no máximo uma vez por mês.
- Dias sem divulgação (fins de semana e feriados) não aparecem no gráfico.

## Fonte dos dados

[Banco Central do Brasil, API SGS](https://dadosabertos.bcb.gov.br/)
