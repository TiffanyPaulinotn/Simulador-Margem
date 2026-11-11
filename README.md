# Simulador de Margens – versão Python

Réplica da aba **SIMULADOR** da planilha `Simulador_de_Margens_Canais_BR_jun23.xlsb`,
lendo direto as abas de base (Custos Fábrica/CD/DT/Broker, Preço, Impostos, Fretes,
Taxa Filial, Contratos, DE-PARA, Listas, Prêmio DT + Preço Lista).

## Como rodar

```bash
pip install -r requirements.txt
# copie a planilha .xlsb para esta pasta (ou informe o caminho na barra lateral)
streamlit run app.py
```

Abre no navegador em http://localhost:8501. A primeira carga leva alguns segundos;
depois fica em cache (`.cache.pkl`). Para o mês novo é só trocar a planilha (ou enviar
pela barra lateral) — o cache é refeito sozinho.

## Base de resultados (recomendado)

Coloque na pasta também o **export mensal do BI (.xlsx)** — colunas `Ano/Mês`, `Expedição`, `Cidade`,
`Canal`, `Produto`, `Incoterm`, `Volume (t)`, `Receita Bruta (R$)`... (qualquer nome de arquivo; o app
detecta sozinho) ou envie pela barra lateral. Ela passa a ser a **fonte principal do histórico**
(R$/t = soma dos R$ ÷ soma das toneladas no período escolhido). As bases jun/23 continuam sendo usadas
só para o que a base nova não tem (taxa filial, contratos, preço lista, prêmio AP, chapa) ou quando não há
volume para a combinação.

Ordem de busca de cada linha: base de resultados (combinação exata) → bases jun/23 → médias da base de
resultados (expedição/canal/produto → expedição/produto) → médias jun/23 → frete médio da cidade /
frete comercial. A coluna **Origem** mostra qual foi usada.

Observação: a coluna `Capex (R$)` da base de resultados tem valores inválidos em algumas linhas
(ex. 2,9 trilhões); essas linhas são ignoradas no cálculo do Capex.

## Arquivos

| Arquivo | O que faz |
|---|---|
| `resultados.py` | Lê a base de resultados do BI e calcula as linhas em R$/t por combinação. |
| `exportar.py` | Gera o Excel formatado (Resumo, DRE comparativa, Rotas, Premissas e origem). |
| `arquivos.py` | Identifica sozinho cada arquivo (planilha de bases ou base de resultados), em .xlsb, .xlsx, .xlsm ou .csv. |
| `bases.py` | Lê as abas da planilha e monta as tabelas de consulta (equivalente aos PROCV). |
| `motor.py` | Cálculo: cascata Preço Bruto → Receita Líquida → Preço FOB → MCU → MGS → Margem Caixa, para Fábrica, CD, DT e Broker. Pode ser usado sem a tela. |
| `app.py` | Tela (Streamlit): campos "PREENCHER", tabelas editáveis "ALTERAR", comparativo e exportação para Excel. |

## Usar só o motor (sem tela)

```python
from bases import carregar
from motor import Entrada, simular, resumo

b = carregar("Simulador_de_Margens_Canais_BR_jun23.xlsb")
e = Entrada(fabrica="APIAÍ", cidade="PR_Andirá", canal="Mat.Con.", produto=1393,
            incoterm="CIF", sacos=600, preco_negociado=29.0, cd="CD - Londrina",
            ajustes={"CD": {"Custo Fixo": 70.0}})   # opcional: sobrescreve linhas (R$/t)
r = simular(b, e)
print(resumo(r))                 # FOB Líquido, MCU, MGS, MGCX por canal
print(r["canais"]["Fábrica"])    # cascata completa
```

Dá para rodar em lote (vários clientes/preços) num loop e salvar tudo num Excel.

## Regras replicadas da planilha

- Preço Bruto (R$/t) = preço negociado (R$/saco) × 1000 / peso do saco; soma chapa (se "Sim")
  e, no CD, a taxa filial.
- Impostos = Preço Bruto × alíquota do histórico (ou da aba Impostos por UF origem/UF cliente/canal).
- Frete de vendas = Frete Logístico por zona de transporte × (1 + adicional de fracionamento); FOB = 0.
- CD sem histórico usa os valores da fábrica; Movimentação de CDs vem de Contratos (remuneração total).
- DT/AP = histórico do distribuidor da cidade. Broker = histórico da carteira + comissão de Contratos.
- Incentivo é exibido mas não entra na Receita Líquida ("s/ Incentivo"); Frete Devolução simulado = 0.
- Sem histórico → "n/e". Com a opção **Completar dados sem histórico** (padrão ligada), a linha é
  preenchida com a média da base (fábrica/canal/produto → fábrica/produto; CD idem) e o frete com o
  frete comercial cadastrado. A coluna **Origem** mostra de onde veio cada valor. Desligando, fica igual ao Excel.

## Diferenças em relação ao Excel (correções)

Algumas fórmulas do Excel apontam para colunas além do fim da base (índice do PROCV errado)
e por isso devolvem 0 ou "n/e". Aqui as colunas são buscadas **pelo nome do cabeçalho**:

- Descontos (fábrica) → usa "Descontos (Soma com Provisão e Estorno)" (no Excel virava 0).
- Outros Custos Log. → usa "Outros Custos Log. (SOMA)" (no Excel virava 0 em todas as bases).
- CD: Logísticos, Movimentação de CDs e Capex → colunas corretas da Base de Custos CD.
- Broker: impostos recalculados pela alíquota quando o preço do broker é informado.

Se a planilha mudar de layout (colunas renomeadas), ajuste os nomes em `motor.py` / `bases.py`.

## Onde atender (fábrica × CD)

A aba **Onde atender** simula a mesma venda (mesmo preço, produto equivalente de mesma família e peso)
saindo de cada fábrica, direto ou por cada CD dela, e ordena pela margem caixa. Clique numa rota da
tabela e em **Simular** para carregá-la na tela. Rotas sem frete de vendas conhecido para a cidade
aparecem em "sem dados suficientes".

## Preço mínimo

A aba **Preço mínimo** mostra o preço negociado a partir do qual Fábrica e CD atingem a margem caixa
desejada (a margem é linear no preço).

Correção: a tabela Frete Logístico da planilha não tinha o incoterm na chave e o PROCV pegava a 1ª linha
(às vezes a FOB, com frete 0). Agora a busca usa o modelo de frete (CIF/FOB).

## Enviar planilhas

Use **Bases de dados → Enviar / atualizar planilhas** na barra lateral (ou simplesmente copie os arquivos
para a pasta do app). Aceita .xlsb, .xlsx, .xlsm e .csv, com qualquer nome: o simulador reconhece a
planilha de bases pelas abas e a base de resultados pelas colunas, e usa sempre o arquivo mais recente
de cada tipo. Colunas com acento/maiúsculas diferentes e números no formato "1.234,56" também são aceitos.

## Só com a base de resultados

O simulador já vem com as bases jun/23 embutidas (`bases_jun23_embutidas.xlsx`: taxa filial, contratos,
preço lista, DE-PARA...). Basta colocar na pasta a base de resultados do BI (ex. "Sem título - 5 de novembro
de 2026.xlsx"). Se um dia vier uma planilha de bases nova (.xlsb), é só colocá-la na pasta que ela passa a ser
usada no lugar da embutida.

## Por destino (1ª aba)

Escolha o **estado** (ex. SP) ou uma **cidade**, o canal, o cimento e as **fábricas para comparar**
(fábrica 1, 2, 3…). Com o resultado real da base de 2026 (média ponderada pelo volume), a aba mostra:
- o **melhor caminho** (fábrica + formato), a melhor fábrica, o melhor formato, o melhor CD e o melhor cimento;
- **1 · Fábrica × Fábrica**: ranking das fábricas escolhidas, com o resultado de cada uma direto, via CD e distribuidor;
- **2 · Fábrica direto × CD × Distribuidor**: qual formato rende mais, com mapa de calor fábrica × formato;
- **3 · Qual cimento** e **4 · Qual CD**: rankings por margem caixa;
- **5 · CIF × FOB**: margem de cada incoterm, por fábrica e por formato (filtro Incoterm: CIF + FOB, CIF ou FOB).

Na simulação, o quadro **CIF × FOB** calcula os dois incoterms ao mesmo tempo: o escolhido com o preço
informado e o outro com o preço médio praticado (ou estimado pelo frete comercial).

Os filtros mostram todas as opções (fábricas, cidades, canais, cimentos), sem duplicadas por acento ou
espaço; as que têm venda aparecem primeiro, com as toneladas, e as demais como "sem venda".

## Simular rotas (antiga "Comparar fábrica × CD × distribuidor")

Primeira aba. Simula a mesma venda (mesmo preço e produto equivalente) por todos os caminhos e mostra:
1. **Fábrica × Fábrica** – qual fábrica dá mais margem vendendo direto;
2. **CD × Fábrica** – se compensa atender via CD ou direto;
3. **Distribuidor × Fábrica/CD** – se compensa vender para o distribuidor/AP;
e o veredito com a **melhor opção**, a tabela por fábrica e o ranking. A opção "só rotas com histórico de 2026"
(ligada por padrão) deixa de fora as rotas calculadas apenas com as bases jun/23.

Os filtros da barra lateral mostram primeiro as cidades, canais e produtos que têm venda na base (com as
toneladas), e o preço já vem preenchido com o preço médio praticado.

## Conferir com a planilha Excel

No topo da tela, escolha **📗 Igual à planilha Excel** (padrão) ou **✅ Corrigido (dados 2026)**. Nesse modo o simulador usa só as bases
jun/23, não completa lacunas e repete as fórmulas da aba SIMULADOR — inclusive as que buscam colunas que não
existem e por isso zeram linhas (no CD: Movimentação de CDs e Capex; em todos: Outros Custos Log.; Descontos
da fábrica; frete de vendas buscado sem o incoterm, que pega a linha FOB = 0 mesmo em venda CIF; linhas
sem histórico entram como 0). Exemplos conferidos (CIF e FOB): APIAÍ → PR_Abatiá, Mat.Con., 1164, FOB R$ 22,00, CD Londrina →
Fábrica MGCX 47,20 e CD −22,65, iguais ao Excel.

Fora desse modo, o simulador usa o histórico real de 2026 e as colunas corretas, então os valores do CD
são diferentes da planilha (por exemplo, Movimentação de CDs e Capex deixam de ser zero).

## Volume em toneladas

O volume pode ser informado em **toneladas** (padrão) ou em sacos; os cartões mostram a margem em R$/t e o
total no volume em toneladas.

## Login (só e-mail @intercement.com)

O simulador pede login com a **conta Microsoft corporativa** (a mesma do Outlook/Teams) e só libera e-mails
**@intercement.com** — outras contas veem "Acesso não permitido". Não existe senha própria do simulador.

- **No seu computador** (localhost), sem configuração, ele abre direto com o aviso "Modo local".
- **Na internet/rede**, sem configuração, o acesso fica **bloqueado** até a TI configurar o login.

Para a TI configurar (uma vez):
1. Registrar um aplicativo no **Microsoft Entra ID** (Azure AD) da InterCement, com Redirect URI
   `https://<endereço-do-simulador>/oauth2callback`, e gerar um client secret.
2. Copiar `.streamlit/secrets.toml.exemplo` para `.streamlit/secrets.toml` no servidor e preencher
   `redirect_uri`, `cookie_secret`, `client_id`, `client_secret` e o `<TENANT_ID>`.
3. Instalar as dependências (`pip install -r requirements.txt`, inclui Authlib) e rodar
   `streamlit run app.py --server.port 8501 --server.address 0.0.0.0`.
4. Apontar o domínio (ex. `simulador.intercement.com`) para o servidor, com HTTPS.

Para liberar outro domínio, edite `dominios = ["intercement.com"]` no secrets.toml.
O `secrets.toml` nunca deve ir para o GitHub (já está no `.gitignore`).
