# Projeto – Pesquisa Nacional (Tratamento e Amostragem)

Automação completa em Python para tratamento, consolidação e geração de **três amostras** de grandes bases de dados da Pesquisa Nacional.  
O projeto foi desenvolvido para padronizar seis fontes distintas, totalizando mais de 100 mil registros brutos, aplicando regras de exclusão, limpeza e amostragem controlada.

---

## Contexto e objetivo

O projeto surgiu da necessidade de unificar bases heterogêneas em um formato padronizado, garantindo rastreabilidade e reprodutibilidade em todas as etapas.  
Antes da automação, o processo de tratamento e consolidação era feito manualmente em Excel e demandava mais de 30 dias de trabalho da equipe.  
Com o novo fluxo em Python, o mesmo volume de dados passou a ser tratado e auditado em poucos minutos, com resultados padronizados e totalmente reproduzíveis.

---

## Arquitetura do fluxo

1. Leitura e consolidação das seis planilhas originais  
2. Limpeza e normalização de campos  
3. Aplicação de regras de exclusão por telefone, segmento e siglas  
4. Consolidação da base final (Tabelão_Final.xlsx)  
5. Geração de três amostras segmentadas por origem e região  
6. Controle de não repetição de IDs entre as amostras  

---

## Estrutura de pastas

Pesquisa_Nacional_Tratamento/  
├── src/  
│   ├── tratamento_principal.py  
│   ├── segmentado.py  
│   └── amostra_2.py  
│  
├── data/  
│   ├── Entradas/  
│   └── Saidas/  
│  
└── README.md  

---

## Scripts principais

### tratamento_principal.py  
Responsável por todo o tratamento das planilhas: exclusão de registros inválidos, limpeza de textos, padronização de colunas e consolidação dos arquivos em uma base única (Tabelão_Final.xlsx).

### segmentado.py  
Executa a geração da **primeira amostra proporcional** por origem e região, com metas predefinidas.  
O script aplica alocação proporcional, sorteio aleatório e priorização de compradores, além de gerar planilhas de auditoria automáticas (Amostra_Final_OrigemFixa.xlsx).

### amostra_2.py  
Cria as **amostras subsequentes** (segunda e terceira), sem repetir IDs já utilizados.  
Respeita novamente as metas por origem e aplica a lógica de seleção de contatos principais.  
As saídas são registradas nos arquivos Amostra_2_SemRepetir.xlsx e Amostra_3_SemRepetir.xlsx.

---

## Destaques técnicos

- Processamento e automação de mais de 100 mil linhas brutas  
- Redução de execução manual de 30 dias para poucos minutos  
- Geração de três amostras controladas e auditáveis  
- Uso de Pandas e NumPy para manipulação e sorteio proporcional  
- Exportação multi-aba com controle de auditoria (xlsxwriter)  
- Estrutura modular e reaproveitável para futuras execuções  
- Pipeline totalmente rastreável e reproduzível  

---

## Tecnologias

Python · Pandas · NumPy · OS · Truststore · Excel Automation  

---

## Resultados

- Padronização de seis fontes distintas em um formato único  
- Automação completa do fluxo de tratamento e amostragem  
- Geração de três amostras balanceadas, auditáveis e sem repetição  
- Aumento expressivo de eficiência operacional e qualidade dos dados  

---

**Autora:** Ana Fonseca  
Engenharia de Dados | Automação e Qualidade de Dados


