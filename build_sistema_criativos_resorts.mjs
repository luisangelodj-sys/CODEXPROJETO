import fs from "node:fs/promises";
import path from "node:path";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const outputDir = path.resolve("outputs", "sistema_criativos_resorts");
await fs.mkdir(outputDir, { recursive: true });

const status = "RASCUNHO PARA APROVACAO";

const coreInputs = [
  ["Produto real", "Ferias recorrentes com clareza de uso, estrutura de resort, destino, servico, previsibilidade e decisao segura.", "Nao reduzir a venda a cota, contrato ou luxo generico."],
  ["Promessa segura", "Entenda como ter ferias mais planejadas em um resort, com estrutura, regras claras de uso e experiencia pensada para sua familia.", "Sem retorno garantido, renda passiva, valorizacao certa ou lucro com ferias."],
  ["Inimigo narrativo", "Diaria cara na alta temporada, viagem improvisada, casa de praia com manutencao, Airbnb imprevisivel, vendedor agressivo, taxa escondida, contrato mal explicado.", "O inimigo nao deve ser um concorrente especifico."],
  ["Desejo central", "Tempo de qualidade com familia, conforto, lazer, previsibilidade, status, pertencimento, conveniencia, destino desejado e servico.", "Desejo precisa ter prova visual ou argumento concreto."],
  ["Filtro central", "Para familias que viajam todos os anos, querem entender antes de decidir e buscam ferias planejadas, nao lucro rapido.", "Filtro evita curioso, investidor especulativo e lead sem fit."],
  ["CTA central", "Receba o guia, faca uma simulacao de uso, compare com hotel/Airbnb/casa de praia, agende conversa consultiva.", "Evitar saiba mais, clique agora, garanta ja e ultima chance sem prova."],
];

const personas = [
  "Familia classe A/B com filhos",
  "Casal que quer ferias recorrentes",
  "Empresario sem tempo",
  "Profissional liberal aspiracional",
  "Comprador de segunda residencia",
  "Cliente que ja viaja para resort",
  "Cliente que compara com hotel",
  "Cliente que compara com Airbnb",
  "Cliente que pensa em casa de praia",
  "Cliente com medo de multipropriedade",
  "Cliente que quer status",
  "Cliente que quer destino seguro para familia",
  "Cliente que precisa convencer conjuge",
  "Investidor cauteloso",
  "Cliente premium que quer privacidade e servico",
];

const objections = [
  "Isso e golpe?",
  "Isso e timeshare?",
  "Vou conseguir usar?",
  "Tem taxa escondida?",
  "E se eu nao viajar?",
  "Posso vender?",
  "Posso alugar?",
  "E investimento?",
  "Vale mais que hotel?",
  "Vale mais que Airbnb?",
  "Por que nao comprar casa de praia?",
  "E caro.",
  "Minha familia vai usar?",
  "Tenho que falar com meu conjuge.",
  "Tenho medo de contrato longo.",
  "Nao entendi o que estou comprando.",
  "Ja ouvi falar mal desse mercado.",
  "E se a obra atrasar?",
  "Quem administra?",
  "Tem escritura ou e direito de uso?",
];

const channels = [
  ["Instagram Reels", "Video curto", "Descoberta, desejo, objecao e WhatsApp"],
  ["Instagram Stories", "Sequencia interativa", "Pesquisa, relacionamento, direct e clique"],
  ["Instagram Carrossel", "Educativo/salvavel", "Educacao, comparacao e FAQ"],
  ["TikTok", "Video curto oral", "Curiosidade, objecao e viralidade qualificada"],
  ["YouTube Shorts", "Video curto com tese", "Topo e remarketing"],
  ["YouTube longo", "Video educativo/tour", "Autoridade e decisao"],
  ["Meta Ads", "Video/carrossel/formulario/WhatsApp", "Lead, remarketing e conversao"],
  ["Google Display", "Banner/HTML5", "Remarketing e lembranca"],
  ["YouTube Ads", "Video pago", "Educacao e remarketing"],
  ["Landing Page", "Blocos de conversao", "Explicar, provar e converter"],
  ["WhatsApp", "Mensagem consultiva", "Qualificar e agendar"],
  ["CRM", "Sequencia", "Nutrir, recuperar e segmentar"],
  ["App/Area cliente", "Tela/notificacao", "Pos-venda, prova e indicacao"],
];

const blocks = [
  ["BLOCO 1", "Criativos de atencao", 20, "Parar rolagem sem prometer demais.", "Retencao 3s, comentarios qualificados, salvamentos", "Gancho forte + tensao + CTA leve"],
  ["BLOCO 2", "Criativos de desejo", 20, "Ativar familia, resort, destino, conforto, status e servico.", "Compartilhamentos, visitas perfil, cliques", "Visual real + desejo especifico + prova"],
  ["BLOCO 3", "Criativos educativos", 20, "Explicar produto complexo de forma simples.", "Salvamentos, tempo de visualizacao, leads guia", "Glossario + comparativo + CTA guia"],
  ["BLOCO 4", "Criativos de objecao", 30, "Quebrar medo, taxa, contrato, uso, investimento e golpe.", "Comentarios qualificados, WhatsApp, FAQ views", "Objecao direta + resposta segura + prova"],
  ["BLOCO 5", "Criativos de qualificacao", 20, "Filtrar curioso e atrair comprador com perfil.", "Leads com fit, taxa de resposta, agendamentos", "Para voce se... / nao e para voce se..."],
  ["BLOCO 6", "Criativos de comparacao", 20, "Comparar resort com hotel, Airbnb, casa de praia e modelos.", "Salvamentos, cliques comparativo, tempo na pagina", "Tabela simples + decisao por perfil"],
  ["BLOCO 7", "Criativos de prova", 20, "Mostrar estrutura, documentos, destino, bastidores e operacao.", "CTR, confianca, pedidos de visita", "Prova visual + detalhe concreto"],
  ["BLOCO 8", "Criativos de remarketing", 20, "Recuperar quem viu, clicou, salvou, sumiu ou foi ao showroom.", "Retorno WhatsApp, agendamento, proposta", "Retomar objecao + proximo passo"],
  ["BLOCO 9", "Criativos pagos", 24, "Criar versoes para Meta, TikTok, YouTube e Display.", "CPL qualificado, CTR, conversa, show rate", "Copy segura + landing coerente"],
  ["BLOCO 10", "Landing page", 8, "Criar dobras que convertem: hero, desejo, explicacao, comparacao, prova, FAQ, CTA.", "Conversao, scroll, clique WhatsApp", "Uma objecao por dobra"],
  ["BLOCO 11", "WhatsApp", 13, "Mensagens consultivas por momento e objecao.", "Taxa de resposta, agendamento, proposta", "Pergunta de perfil antes de proposta"],
  ["BLOCO 12", "App/area cliente", 12, "Telas/notificacoes para pos-venda, prova e indicacao.", "Ativacao, NPS, indicacoes", "Uso real vira prova comercial"],
];

const hookFormulas = [
  ["Alerta", "Antes de comprar [produto], veja isso.", "Antes de comprar uma cota em resort, veja isso."],
  ["Comparacao", "[A] ou [B]: qual faz mais sentido?", "Casa de praia, hotel ou resort compartilhado: qual faz mais sentido?"],
  ["Objecao", "[Objeção]? Depende do que te explicaram.", "Multipropriedade e golpe? Depende do que te explicaram."],
  ["Curiosidade", "O que quase ninguem pergunta antes de [decisao].", "O que quase ninguem pergunta antes de comprar uma casa de ferias em resort."],
  ["Dor", "Todo ano voce [dor]?", "Todo ano voce paga diaria cara e ainda viaja no improviso?"],
  ["Anti-promessa", "Nao compre [produto] pensando primeiro em [promessa ruim].", "Nao compre resort pensando primeiro em lucro. Entenda o uso."],
  ["Selecao", "Esse modelo nao e para todo mundo.", "Esse modelo nao e para todo mundo. E por isso precisa ser explicado antes."],
  ["Autoridade", "[Numero] perguntas que um comprador inteligente faz antes de [acao].", "3 perguntas que um comprador inteligente faz antes de assinar."],
  ["Familia", "Se suas ferias sempre ficam para depois, [consequencia].", "Se suas ferias sempre ficam para depois, sua familia sente isso."],
  ["Manutencao", "[Alternativa] parece sonho ate virar [dor].", "Casa de praia parece sonho ate virar conta, reforma e preocupacao."],
  ["Luxo real", "Luxo nao e so [aparencia]. E [operacao].", "Luxo nao e so piscina bonita. E chegar e tudo funcionar."],
  ["Transparencia", "Se ninguem explicou [pontos], voce ainda nao recebeu proposta completa.", "Se ninguem explicou taxa, calendario e contrato, voce ainda nao recebeu proposta completa."],
];

const filtersAndCtas = [
  ["Filtro", "Para familias que viajam todos os anos.", "Atrai recorrencia e poder de uso"],
  ["Filtro", "Para quem quer entender antes de decidir.", "Atrai lead racional"],
  ["Filtro", "Para quem busca ferias planejadas, nao lucro rapido.", "Remove especulador"],
  ["Filtro", "Para quem compara com hotel, Airbnb e casa de praia.", "Atrai intencao real"],
  ["Filtro", "Para quem deseja apresentacao clara, sem pressao.", "Atrai lead consultivo"],
  ["CTA", "Receba o guia antes de falar com um consultor.", "Topo/meio"],
  ["CTA", "Faca uma simulacao de uso.", "Meio/fundo"],
  ["CTA", "Compare com hotel, Airbnb e casa de praia.", "Comparacao"],
  ["CTA", "Entenda regras, custos e disponibilidade.", "Objecao"],
  ["CTA", "Agende uma conversa consultiva.", "Fundo"],
  ["CTA", "Receba o checklist antes de decidir.", "Meio"],
  ["CTA", "Veja se faz sentido para sua familia.", "Qualificacao"],
];

const validationMatrix = [
  ["Retencao 3s", "Valida gancho", "Baixa retencao inicial", "Reescrever gancho, cortar introducao, abrir com tensao"],
  ["Retencao 50%", "Valida roteiro", "Queda depois do gancho", "Melhorar explicacao e ritmo visual"],
  ["Salvamentos", "Valida conteudo educativo", "Poucos salvamentos em FAQ/comparativo", "Transformar em checklist mais claro"],
  ["Compartilhamentos", "Valida identificacao", "Baixo compartilhamento", "Aumentar situacao especifica de familia/dor"],
  ["Comentarios qualificados", "Valida conversa", "Comentario generico ou piada", "Adicionar filtro e pergunta qualificada"],
  ["Cliques WhatsApp", "Valida intencao", "Muitos cliques e pouca resposta", "Ajustar promessa e primeira mensagem"],
  ["Leads com perfil", "Valida qualificacao", "Leads sem renda/sem viagem recorrente", "Reforcar filtro no criativo"],
  ["Agendamentos", "Valida oferta", "Lead conversa mas nao agenda", "Melhorar prova, FAQ e proximo passo"],
  ["Show rate", "Valida compromisso", "Agenda e nao comparece", "Enviar material pre-call e confirmar decisor"],
  ["Propostas", "Valida consultivo", "Muitos leads sem proposta", "Treinar diagnostico e segmentacao"],
  ["Vendas", "Valida sistema", "Propostas nao fecham", "Revisar oferta, objeções e follow-up"],
];

const decisionMatrix = [
  ["Manter", "Criativo bate benchmark minimo e gera lead com fit", "Continuar rodando e coletar comentarios"],
  ["Editar", "Boa atencao, mas lead ruim ou baixa conversao", "Inserir filtro, ajustar CTA e prova"],
  ["Pausar", "Baixa retencao, baixo clique e lead ruim", "Pausar e recriar a partir da objecao real"],
  ["Escalar", "Boa retencao, leads qualificados, agendamento e baixo risco", "Aumentar verba ou transformar em serie"],
];

const calendar30 = Array.from({ length: 30 }, (_, i) => {
  const day = i + 1;
  const cycle = [
    ["Atencao", "Reels/TikTok", "Antes de comprar uma cota em resort, veja isso.", "Receber guia"],
    ["Desejo", "Reels/Stories", "Ferias planejadas em familia sem improviso.", "Simular perfil"],
    ["Educativo", "Carrossel", "Multipropriedade, clube e residence club: diferencas.", "Salvar checklist"],
    ["Objecao", "Reels/Shorts", "Isso e golpe? O que precisa estar claro.", "Ver FAQ"],
    ["Qualificacao", "Stories/Reels", "Esse modelo e para voce se...", "Falar consultor"],
    ["Comparacao", "Carrossel/YouTube", "Hotel, Airbnb ou resort?", "Ver comparativo"],
    ["Prova", "Tour/Stories", "Tour do resort e regras de uso.", "Agendar visita"],
    ["Remarketing", "Ads/CRM", "Ainda em duvida sobre taxa e contrato?", "Retomar conversa"],
  ][i % 8];
  return [day, cycle[0], cycle[1], cycle[2], cycle[3], "Retencao, clique, lead qualificado", status];
});

const scoreCriteria = [
  "Atencao",
  "Clareza",
  "Desejo",
  "Confianca",
  "Qualificacao do lead",
  "Potencial de compartilhamento",
  "Potencial de salvamento",
  "Potencial de comentario",
  "Potencial de clique",
  "Seguranca de compliance",
  "Forca comercial",
  "Adequacao ao publico premium",
];

const sources = [
  ["Base anterior", "Plano_Operacional_Maquina_Vendas_Resorts.xlsx", "Funil, canais, scripts, compliance, calendario e banco de copy."],
  ["Dossie 3.0", "Dossie_3_0_Resorts_Marketing_Operacional.xlsx", "Segmentacao, auditoria de sites, ofertas, ads e landing."],
  ["Base inteligencia", "Base_Inteligencia_Imobiliario_Turistico_Premium.xlsx", "Modelos de negocio, players, personas, objeções, claims e fontes."],
  ["Meta Advertising Standards", "https://transparency.meta.com/policies/ad-standards/", "Regras de politica e coerencia anuncio/landing."],
  ["TikTok Misleading Content", "https://ads.tiktok.com/help/article/tiktok-ads-policy-misleading-and-false-content", "Evitar claim enganoso, exagerado e omissao de informacao."],
  ["Google Helpful Content", "https://developers.google.com/search/docs/fundamentals/creating-helpful-content", "Conteudo util, confiavel e orientado a pessoa."],
  ["Lei 13.777/2018", "http://www.planalto.gov.br/ccivil_03/_ato2015-2018/2018/lei/L13777.htm", "Referencia de multipropriedade imobiliaria."],
];

const promptMaster = `# PROMPT MESTRE - SISTEMA DE CRIATIVOS PARA IMOBILIARIO TURISTICO PREMIUM

Voce e um diretor criativo senior, estrategista de performance, copywriter high-ticket e especialista em marketing para imobiliario turistico premium.

Voce cria criativos para resorts, multipropriedade, residence club, branded residences, clubes de ferias, segunda residencia, vacation ownership, destination club e empreendimentos imobiliarios de alto padrao ligados a hospitalidade.

## Missao

Criar CRIATIVOS QUE VIRALIZAM COM INTENCAO COMERCIAL.

Nao crie post bonito por vaidade. Crie criativos que:
- chamem atencao nos 3 primeiros segundos;
- despertem desejo real;
- expliquem produto complexo de forma simples;
- eduquem mercado;
- reduzam desconfianca;
- filtrem curioso;
- validem lead;
- levem para WhatsApp, site, formulario, visita ou consultor;
- respeitem compliance;
- nao parecam golpe;
- nao prometam investimento garantido;
- nao usem luxo vazio;
- nao atraiam apenas curioso sem poder de compra.

## Contexto correto do mercado

Nao venda "cota" como ideia central.

O mercado deve ser tratado como:

imobiliario turistico premium + hospitalidade + resort + segunda residencia + casa de ferias + clube de ferias + branded residence + residence club + venda consultiva high-ticket.

O cliente compra ferias planejadas, previsibilidade, acesso a destino desejado, estrutura de resort, lazer para familia, status, pertencimento, conforto, seguranca na decisao, conveniencia, reducao da dor de manter uma casa de praia sozinho, experiencia recorrente, servico e tempo de qualidade.

## Inimigo narrativo

Use como tensao:
- diaria cara na alta temporada;
- viagem improvisada;
- casa de praia que da manutencao;
- Airbnb imprevisivel;
- hotel caro todo ano;
- vendedor agressivo;
- promessa de investimento facil;
- contrato mal explicado;
- taxa escondida;
- medo de golpe;
- falta de clareza sobre uso.

## Promessa segura

Promessa principal:
"Entenda como ter ferias mais planejadas em um resort, com estrutura, regras claras de uso e uma experiencia pensada para sua familia."

Nunca use como promessa principal:
- retorno garantido;
- renda passiva;
- valorizacao certa;
- investimento sem risco;
- lucro com ferias;
- compre agora ou perca;
- ultimas unidades sem validacao;
- melhor investimento do Brasil;
- fique rico com resort.

## Processo antes de criar

Antes de escrever qualquer criativo, defina:
1. objetivo;
2. persona;
3. etapa do funil;
4. canal;
5. dor;
6. desejo;
7. objecao;
8. filtro de lead;
9. prova necessaria;
10. CTA;
11. metrica;
12. risco de compliance.

## Estrutura para videos curtos

Use:
1. Gancho de 0 a 3 segundos.
2. Tensao ou problema.
3. Explicacao simples.
4. Prova visual ou argumento concreto.
5. Filtro de qualificacao.
6. CTA consultivo.

## Filtros obrigatorios

Inclua filtros como:
- Para familias que viajam todos os anos.
- Para quem quer entender antes de decidir.
- Para quem busca ferias planejadas, nao lucro rapido.
- Para quem compara com hotel, Airbnb e casa de praia.
- Para quem tem perfil de viagem recorrente.
- Para quem deseja uma apresentacao clara, sem pressao.

Evite CTAs genericos: saiba mais, clique agora, garanta ja, ultima chance, nao perca.

Prefira CTAs consultivos:
- Veja se faz sentido para sua familia.
- Receba o guia antes de falar com um consultor.
- Faca uma simulacao de uso.
- Entenda regras, custos e disponibilidade.
- Compare com hotel, Airbnb e casa de praia.
- Agende uma conversa consultiva.
- Receba o checklist antes de decidir.

## Saida obrigatoria

Entregue em tabela com estas colunas:
1. Nome do criativo.
2. Canal.
3. Formato.
4. Etapa do funil.
5. Persona.
6. Dor trabalhada.
7. Desejo ativado.
8. Objecao quebrada.
9. Gancho.
10. Roteiro/texto.
11. Texto na tela.
12. Legenda/copy.
13. Ideia visual.
14. Prova necessaria.
15. Filtro de lead.
16. CTA.
17. Metrica principal.
18. Risco de compliance.
19. Versao segura.
20. Nota de forca comercial.

## Blocos a criar

Crie:
- 20 criativos de atencao;
- 20 de desejo;
- 20 educativos;
- 30 de objecao;
- 20 de qualificacao;
- 20 de comparacao;
- 20 de prova;
- 20 de remarketing;
- criativos pagos para Meta, TikTok, YouTube e Google Display;
- blocos de landing page;
- mensagens de WhatsApp;
- telas/notificacoes de app ou area do cliente.

## Score obrigatorio

Avalie cada criativo de 0 a 10 em:
- atencao;
- clareza;
- desejo;
- confianca;
- qualificacao do lead;
- potencial de compartilhamento;
- potencial de salvamento;
- potencial de comentario;
- potencial de clique;
- seguranca de compliance;
- forca comercial;
- adequacao ao publico premium.

Se a nota de forca comercial for menor que 8, reescreva antes de entregar.

## Matriz de validacao

Para cada criativo, informe a metrica principal:
- retencao nos 3 primeiros segundos;
- retencao ate 50%;
- comentarios qualificados;
- salvamentos;
- compartilhamentos;
- cliques no perfil;
- cliques no WhatsApp;
- respostas no direct;
- leads com perfil;
- agendamentos;
- comparecimento;
- propostas;
- vendas.

Tambem diga qual sinal indicaria lead errado:
- comentario generico;
- pergunta so de preco;
- lead sem renda;
- lead sem intencao de viajar;
- lead buscando lucro rapido;
- lead confundindo com diaria, sorteio, emprego ou brinde;
- baixa resposta no WhatsApp.

## Regra final

Nao crie criativo generico.
Nao venda apenas cota.
Nao venda apenas resort bonito.
Nao venda apenas luxo.

Venda clareza, desejo, confianca, uso real, experiencia, familia, destino, servico e decisao segura.

Todo criativo precisa responder:
- Por que essa pessoa pararia para ver?
- Por que ela confiaria?
- Por que ela se identificaria?
- Por que ela clicaria?
- Por que ela seria lead qualificado?
- Por que esse criativo nao atrai apenas curioso?
- Qual proxima acao comercial esse criativo gera?
`;

const workbook = Workbook.create();

function writeSheet(name, rows, widths = []) {
  const sheet = workbook.worksheets.add(name);
  sheet.showGridLines = false;
  const cols = rows[0].length;
  const data = rows.map((row) => {
    if (row.length === cols) return row;
    if (row.length < cols) return [...row, ...Array(cols - row.length).fill("")];
    return [...row.slice(0, cols - 1), row.slice(cols - 1).join(" | ")];
  });
  const range = sheet.getRangeByIndexes(0, 0, data.length, cols);
  range.values = data;
  sheet.freezePanes.freezeRows(1);
  range.format = { wrapText: true, borders: { preset: "all", style: "thin", color: "#D9D9D9" } };
  const header = sheet.getRangeByIndexes(0, 0, 1, cols);
  header.format = { fill: "#183B56", font: { bold: true, color: "#FFFFFF" }, wrapText: true };
  try {
    const lastCol = String.fromCharCode("A".charCodeAt(0) + cols - 1);
    const table = sheet.tables.add(`A1:${lastCol}${data.length}`, true, name.replace(/[^A-Za-z0-9]/g, "_").slice(0, 30));
    table.style = "TableStyleMedium2";
    table.showFilterButton = true;
  } catch {}
  for (let c = 0; c < cols; c++) {
    sheet.getRangeByIndexes(0, c, data.length, 1).format.columnWidthPx = widths[c] ?? (c === 0 ? 180 : 260);
  }
  sheet.getRangeByIndexes(1, 0, Math.max(data.length - 1, 1), cols).format.rowHeightPx = 92;
}

writeSheet("00 Leia Primeiro", [
  ["Campo", "Conteudo"],
  ["Objetivo", "Base de sistema de criativos para transformar inteligencia de mercado em criacao com funil, persona, objecao, filtro, metrica e compliance."],
  ["Como usar", "Use a aba 10 Prompt Mestre para pedir criativos. Use as abas 01-09 como base de dados para briefing, validacao e score."],
  ["Regra", "Criativo so entra em producao se tiver objetivo, persona, objecao, prova, filtro, CTA, metrica e risco de compliance."],
  ["Status", status],
], [180, 900]);
writeSheet("01 Inputs Mercado", [["Input", "Direcao", "Alerta"], ...coreInputs], [220, 620, 520]);
writeSheet("02 Personas Objecoes", [["Persona", "Objecao prioritaria", "Dor", "Desejo", "Filtro recomendado"], ...personas.map((p, i) => [p, objections[i % objections.length], "Medo de errar na decisao e pagar por algo que nao usa.", "Ferias planejadas com familia, servico e clareza.", filtersAndCtas[i % 5][1]])], [300, 300, 420, 420, 360]);
writeSheet("03 Canais", [["Canal", "Formato", "Funcao"], ...channels], [230, 260, 520]);
writeSheet("04 Blocos Criativos", [["Bloco", "Tipo", "Quantidade", "Objetivo", "Metrica", "Estrutura"], ...blocks], [120, 260, 120, 520, 360, 380]);
writeSheet("05 Formulas Ganchos", [["Tipo", "Formula", "Exemplo"], ...hookFormulas], [180, 420, 520]);
writeSheet("06 Filtros CTAs", [["Tipo", "Texto", "Uso"], ...filtersAndCtas], [150, 520, 420]);
writeSheet("07 Score", [["Criterio", "Pergunta de avaliacao", "Nota minima"], ...scoreCriteria.map((c) => [c, `Esse criativo entrega ${c.toLowerCase()} de forma forte e especifica para o nicho?`, 8])], [280, 640, 120]);
writeSheet("08 Validacao", [["Metrica", "O que valida", "Sinal de problema", "Acao"], ...validationMatrix], [260, 360, 420, 420]);
writeSheet("09 Decisao", [["Decisao", "Criterio", "Acao"], ...decisionMatrix], [160, 520, 420]);
writeSheet("10 Prompt Mestre", [["Secao", "Texto"], ...promptMaster.split("\n\n").map((chunk, i) => [`Bloco ${i + 1}`, chunk])], [160, 1000]);
writeSheet("11 Calendario 30d", [["Dia", "Tipo", "Canal/Formato", "Tema", "CTA", "Metrica", "Status"], ...calendar30], [70, 160, 220, 520, 220, 300, 220]);
writeSheet("12 Fontes", [["Fonte", "Referencia", "Uso"], ...sources], [260, 520, 620]);

const errors = await workbook.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
  options: { useRegex: true, maxResults: 100 },
  summary: "formula error scan",
});

for (const sheetName of ["01 Inputs Mercado", "04 Blocos Criativos", "08 Validacao", "10 Prompt Mestre"]) {
  const img = await workbook.render({ sheetName, autoCrop: "all", scale: 1, format: "png" });
  await fs.writeFile(path.join(outputDir, `${sheetName.replaceAll(" ", "_")}.png`), new Uint8Array(await img.arrayBuffer()));
}

const xlsxPath = path.join(outputDir, "Base_Sistema_Criativos_Resorts.xlsx");
const xlsx = await SpreadsheetFile.exportXlsx(workbook);
await xlsx.save(xlsxPath);

const promptPath = path.join(outputDir, "Prompt_Mestre_Criacao_Criativos_Resorts.md");
await fs.writeFile(promptPath, promptMaster, "utf8");

const guide = `# Guia - Sistema de Criativos para Resorts e Imobiliario Turistico Premium

Este pacote cruza as bases anteriores e organiza o processo de criacao de criativos.

## O que mudou

Antes, a base dizia o que vender, para quem, por qual canal e com quais riscos.
Agora, a base orienta como criar pecas com:

- funil;
- persona;
- dor;
- desejo;
- objecao;
- filtro de lead;
- CTA;
- prova;
- metrica;
- compliance;
- score comercial.

## Uso recomendado

1. Abra a planilha Base_Sistema_Criativos_Resorts.xlsx.
2. Leia as abas 01 a 09 para entender inputs, blocos, ganchos, filtros, score e validacao.
3. Use o arquivo Prompt_Mestre_Criacao_Criativos_Resorts.md para pedir a criacao dos criativos.
4. Depois de gerar os criativos, cole os resultados em uma nova aba operacional e aplique o score.

## Regra de qualidade

Criativo bom nao e o mais bonito. E o que:

- prende a pessoa certa;
- afasta curioso errado;
- educa sem esfriar desejo;
- reduz medo;
- leva para uma acao comercial;
- nao viola compliance.
`;

const guidePath = path.join(outputDir, "Guia_Sistema_Criativos_Resorts.md");
await fs.writeFile(guidePath, guide, "utf8");

console.log(errors.ndjson);
console.log(JSON.stringify({ outputDir, xlsxPath, promptPath, guidePath }, null, 2));
