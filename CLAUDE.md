# Horas de Professores — BRASAS BI

Webapp em Google Apps Script para diretores de unidade lançarem, mensalmente, as horas dos professores (turmas, sábado, atividades extras, substituições, faltas, etc.). Mesmo padrão dos webapps de VR e VT, mas escopo restrito a **professores** (uma única categoria, sem separação Administrativo/Docente).

Repositório: `adrianeca/horas`. Irmão de `VR webapp` (`adrianeca/valerefeicao`) e `VT webapp` (`adrianeca/valetransporte`) no mesmo Desktop/Claude — arquitetura e boa parte do código são espelhados entre os três.

## Arquivos

- `Code.gs` — todo o backend (Apps Script / `google.script.run`)
- `Index.html` — front-end single-page (HTML+CSS+JS inline, sem framework)

## IDs e configuração (topo do Code.gs)

- `HORAS_SHEET_ID = '1fbBw4ynJqpIkBwQR0dQyIjnGaw3WBh_KY9bRc-nN-Lg'` — planilha central, aba **"HORAS"** (já existia antes deste webapp, com cabeçalhos prontos — não é auto-provisionada como no VT)
- `FUNC_SHEET_ID` — planilha de funcionários, aba "RJ - UNIDADES" (compartilhada com VR e VT)
- `HUB_SS_ID` — planilha do Hub BRASAS BI (aba SESSOES para autenticação)
- `MEU_ACESSO = 'webhoras'` — chave de acesso que precisa estar na coluna ACESSOS da aba SESSOES do Hub para o diretor ver este card **(ainda precisa ser cadastrada lá)**
- `DP_EMAIL = 'dp@brasas.com'`

## Modelo de dados — aba "HORAS"

Colunas A-V já existiam antes do webapp; o app só lê/escreve A-Q e nunca toca em R-V:

| Col | Campo | Quem escreve |
|---|---|---|
| A-C | Unidade, Mês, Ano | app |
| D | Matrícula | app |
| E | Apelido | app (auto, vem de `FUNC_SHEET_ID`) |
| F | Nome | app |
| G | Nível | app (auto, vem de `FUNC_SHEET_ID` col. AH "NÍVEL PROFESSORES") — **não é digitado pelo diretor** |
| H-Q | Horas Turmas, Horas Turmas Sábado, Ativ. Extras, Subs., Faltas Descontadas, Faltas Abonadas, Top Special, Subs. em Outras Unidade, Faltas Descontadas (dias), Faltas Abonadas (dias) | app — 10 campos numéricos digitados livremente pelo diretor, sem cálculo nem dependência entre eles |
| R-V | Chave Matrícula, Apelido Ajustado, Nível Ajustado, Chave Matrícula Unidade, Data | **fórmula da própria planilha (ex. ARRAYFORMULA) — o app NUNCA escreve aqui** |
| W-X | Editado Em, Editado Por | app (sinalização de edição, mesmo padrão do VR/VT) |
| Y-AA | Comentário, Comentado Em, Comentado Por | app (anotação livre por lançamento, ver seção "Comentários" abaixo) |

Constantes `COL` (índices na planilha de funcionários) e `HORAS_COL` (índices na aba HORAS) documentam exatamente essas posições no topo do `Code.gs`.

- "Subs. em Outras Unidade" é só mais um campo numérico na linha do professor, na unidade dele — não gera lançamento nem visibilidade na unidade "receptora".
- "Substitui" (funcao PROFESSOR) é o único critério de quem aparece no app — vem de `getFuncionarios()`, filtrando `COL.FUNCAO === 'PROFESSOR'` e ativo.

## Comentários por lançamento

Cada linha da tabela de Professores tem um ícone de comentário (💬) que abre um modal para anotar uma observação livre sobre aquele professor/mês — não é um campo de horas, então funciona mesmo com o período bloqueado, e não entra no CSV/Planilha Google exportados.

- Um comentário por lançamento (não é thread/multi-mensagem); salvar um novo texto sobrescreve o anterior. Apagar o texto e salvar limpa o comentário e o rastro de autor/data.
- Visível tanto para o diretor quanto para o DP — ambos usam a mesma tabela (`tabProfessores`), então quem tiver acesso ao webapp e à unidade vê o mesmo comentário.
- Só é possível comentar um lançamento já salvo na planilha (`salvarComentarioHoras` busca a linha por unidade+mês+ano+matrícula); linhas recém-adicionadas via "+ Adicionar"/"Copiar mês anterior" mostram o ícone desabilitado até serem salvas.
- Gravado nas colunas Y (Comentário), Z (Comentado Em) e AA (Comentado Por) — `getHorasSheet_()` cria esse cabeçalho automaticamente na primeira execução após o deploy, igual já faz com W/X (Editado Em/Por).

## Período e bloqueio — diferente do VR/VT

**O professor lança o mês CORRENTE, não o seguinte.** Ex.: horas de julho são lançadas em julho mesmo (diferente do VT/VR, que lançam adiantado o mês seguinte). O campo de resposta ainda se chama `previsto` só por reaproveitar o mesmo formato usado no restante do app — `getCurrentPeriod()`/`getPeriodoPrevistoAtual_()` retornam o mês/ano **atuais**, não mês+1.

- Bloqueia automaticamente a partir do dia 12 do próprio mês.
- Liberação temporária e fluxo de solicitação (diretor pede → e-mail pro DP → aprova/reprova → e-mail de resposta) — abas `LIBERACOES` e `SOLICITACOES`, criadas automaticamente na planilha `HORAS_SHEET_ID`.
- `MOTIVOS_LIBERACAO` no topo do Code.gs é uma **lista provisória** — ajustar quando a Adriane mandar a lista real (mesma pendência do VR/VT).

## Associação de códigos de unidade (NS/CH, MRI/MR)

Mesma lógica do VR/VT: o cadastro de permissões do Hub usa "NS" e "MRI" pra unidades que a planilha de funcionários e os lançamentos chamam de "CH" e "MR". Toda unidade crua passa por `canonUnidade_()`:

```js
const UNIDADE_ALIASES_ = { ns: 'CH', mri: 'MR' };
```

CH e MR continuam sendo unidades diferentes entre si; o alias só evita duplicar NS/CH e MRI/MR como se fossem unidades separadas em listas, filtros e lembretes.

## Lembretes automáticos de preenchimento

Gatilhos mensais (dias 1, 5, 9 e 11), instalados manualmente uma vez rodando `instalarGatilhosLembreteHoras()` no editor do Apps Script:

- **Dia 1**: e-mail de abertura do período, pra todas as unidades com professor ativo, incondicional.
- **Dias 5, 9 e 11**: lembrete só pra quem ainda não tem nenhum lançamento salvo no mês corrente (`getUnidadesPreenchidas_`). Dia 11 avisa que é o último dia antes do bloqueio.
- Destinatário por unidade: lista fixa `DIRETORES_UNIDADE` (mesma lista do VR/VT — e-mail do diretor específico de cada unidade, passada manualmente pela Adriane). Unidades sem e-mail cadastrado ali são puladas.
- `getUnidadesAtivas_()` filtra só unidades com **professor** ativo (`COL.FUNCAO === 'PROFESSOR'`) — diferente do VR/VT, que consideram qualquer funcionário ativo.
- Antes de instalar os gatilhos de verdade, rodar `diagnosticoLembretesHoras()` (ajustando a variável `dia` no topo da função) — só loga no console quem receberia o quê, **não envia e-mail nenhum**.
- Template de e-mail: mesma identidade visual do resto do app (faixa colorida no topo — azul na abertura, âmbar/laranja/vermelho conforme a urgência sobe).

## Deploy

1. Criar o projeto no Apps Script vinculado à planilha `HORAS_SHEET_ID` (aba "HORAS"), colar `Code.gs`/`Index.html`.
2. Publicar como Web App (Implantar > Nova implantação).
3. Adicionar `webhoras` nos acessos dos diretores certos na aba SESSOES do Hub, e o card apontando pra esse webapp.
4. Rodar `diagnosticoLembretesHoras()` pra validar, depois `instalarGatilhosLembreteHoras()` uma vez.

## Pendências conhecidas

- `MOTIVOS_LIBERACAO` ainda é lista provisória.
- Confirmar se `DP_EMAIL` é o destinatário certo das solicitações.
- Registrar `webhoras` no Hub (acessos + card) antes de liberar pros diretores.
- Conferir que as colunas Y/Z/AA da aba "HORAS" estão livres na planilha real antes de publicar a versão com comentários — `getHorasSheet_()` escreve o cabeçalho ali na primeira execução e sobrescreveria qualquer coisa manual que já exista nessas colunas.
