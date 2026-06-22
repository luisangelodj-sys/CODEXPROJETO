import fs from "node:fs/promises";
import path from "node:path";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const outputDir = path.resolve("outputs", "plano_operacional_maquina_vendas_resorts");
await fs.mkdir(outputDir, { recursive: true });

const status = "RASCUNHO PARA APROVACAO";
const certainty = {
  comprovado: "COMPROVADO",
  inferido: "INFERIDO",
  hipotese: "HIPOTESE",
  validacao: "PRECISA VALIDACAO",
};

const sources = [
  ["GAV Resorts", "https://www.gavresorts.com.br/", "Casa de ferias, ferias compartilhadas, multipropriedade, hotelaria, FAQ, blog, WhatsApp e portal.", certainty.comprovado],
  ["Aviva", "https://www.aviva.com.br/", "Hoteis, Residence Clubs, Clube de Ferias, Rio Quente, Costa do Sauipe e Hot Park.", certainty.comprovado],
  ["Beach Park", "https://beachpark.com.br/", "Destino completo, parques, resorts, Vacation Club, Beach Card, blog, redes, WhatsApp e portal do agente.", certainty.comprovado],
  ["Hot Beach", "https://hotbeach.com.br/", "Parque, resorts, ingressos, hospedagem, Residence Club e casa de ferias anexa ao parque.", certainty.comprovado],
  ["Four Seasons Private Residences", "https://www.fourseasons.com/residences/", "Residences, residence clubs, villa rentals, property management, service and peace of mind.", certainty.comprovado],
  ["Ritz-Carlton Residences", "https://www.ritzcarlton.com/en/residences/", "Luxury branded residences, servico, lifestyle, reputacao e beneficios.", certainty.comprovado],
  ["Marriott Vacation Club", "https://www.marriottvacationclubs.com/ownership/vacation-club-brands/marriott-vacation-club.html", "Vacation ownership, request information, club points, villas/city properties e disclosures.", certainty.comprovado],
  ["Hilton Grand Vacations", "https://www.hiltongrandvacations.com/en/discover-hilton-grand-vacations", "Vacation membership, home resort, flexibility, member stories e disclosures.", certainty.comprovado],
  ["Meta Advertising Standards", "https://transparency.meta.com/policies/ad-standards/", "Politicas de anuncios, landing page, praticas enganosas, atributos pessoais, lead ads e Ad Library.", certainty.comprovado],
  ["Google Ads Personalized Advertising", "https://support.google.com/adspolicy/answer/143465", "Regras de publicidade personalizada, categorias sensiveis, dados pessoais e segmentacao.", certainty.comprovado],
  ["TikTok Misleading and False Content", "https://ads.tiktok.com/help/article/tiktok-ads-policy-misleading-and-false-content", "Regras contra conteudo enganoso, claims exagerados, omissao e landing inconsistente.", certainty.comprovado],
  ["Google Search Helpful Content", "https://developers.google.com/search/docs/fundamentals/creating-helpful-content", "Conteudo util, confiavel, people-first e orientado por experiencia real.", certainty.comprovado],
  ["Lei 13.777/2018", "http://www.planalto.gov.br/ccivil_03/_ato2015-2018/2018/lei/L13777.htm", "Marco legal brasileiro da multipropriedade imobiliaria.", certainty.comprovado],
  ["ANPD", "https://www.gov.br/anpd/pt-br", "LGPD, protecao de dados e boas praticas de consentimento/finalidade.", certainty.comprovado],
];

const diagnostics = [
  ["Qual e o verdadeiro produto?", "Ferias recorrentes com clareza de uso, estrutura de resort, servico, destino, previsibilidade e, quando aplicavel, direito imobiliario ou clube bem explicado.", "Cliente nao compra cota; compra um jeito menos improvisado de viajar com familia e menos dor de manter uma segunda casa.", "Desejo: ferias planejadas. Objecao: medo de golpe/taxa. Canal: site + WhatsApp. Ativo: guia comparativo. Metrica: lead qualificado.", certainty.inferido],
  ["O que nao vender como promessa principal?", "Retorno financeiro, renda garantida, valorizacao, economia absoluta, disponibilidade total ou exclusividade vaga.", "Essas promessas aumentam bloqueio, risco juridico e rejeicao de publico premium.", "Ativo: matriz de claims. Metrica: anuncios reprovados = zero.", certainty.comprovado],
  ["Linguagem a evitar", "Cota imperdivel, investimento seguro, renda garantida, oportunidade unica, luxo inesquecivel, sem burocracia, ultimas unidades sem prova.", "Parece pressa, golpe ou venda agressiva.", "Canal: todos. Ativo: guia de compliance. Metrica: taxa de rejeicao/denuncias.", certainty.inferido],
  ["Linguagem que gera confianca", "Modelo claro, contrato, regras de uso, custos recorrentes, disponibilidade, destino, estrutura, servico, documentos e perfil de uso.", "O mercado e desconfiado; clareza vende antes da seducao.", "Ativo: FAQ dificil, comparativo, simulador. Metrica: perguntas repetidas reduzem.", certainty.inferido],
  ["Popular x premium x alto luxo", "Popular compra beneficio claro e condicao; premium compra familia, destino e conveniencia; alto luxo compra raridade, servico, privacidade e reputacao.", "Misturar tudo gera tom errado: desconto em luxo enfraquece; copy sofisticada demais no volume nao converte.", "Ativo: segmentacao. Metrica: lead-fit por campanha.", certainty.inferido],
  ["Objecoes antes do vendedor", "Isso e timeshare? Tem taxa? Posso usar nas ferias? Posso vender? E investimento? E seguro? Meu conjuge vai aceitar?", "Se o vendedor entrar antes da educacao, vira defesa e pressao.", "Ativo: biblioteca de objecoes no site e conteudo. Metrica: show rate.", certainty.inferido],
  ["Etapa que educa", "Descoberta, educacao e comparacao.", "Aqui o publico aprende modelo, custos, alternativas e riscos.", "Ativos: Reels, YouTube, SEO, carrossel, guia. Metrica: salvamentos, watch time, leads guia.", certainty.inferido],
  ["Etapa que vende desejo", "Descoberta e confianca.", "Destino, familia, parque, praia, privacidade, servico e sensacao de chegar com tudo pronto.", "Ativos: tours, bastidores, depoimentos, UGC. Metrica: visualizacao, compartilhamento, clique.", certainty.inferido],
  ["Etapa que prova seguranca", "Comparacao, confianca, qualificacao e proposta.", "Documentos, contratos, regras, FAQ dificil, prova operacional e vendedor consultivo.", "Ativos: dossie, pagina juridica, tabela de custos. Metrica: taxa de proposta e objecoes resolvidas.", certainty.inferido],
  ["Etapa que leva para consultor", "Qualificacao, apresentacao, visita/showroom e proposta.", "Consultor deve entrar quando o lead entende o modelo e tem pergunta real.", "CTA: simular perfil, falar com consultor, agendar apresentacao. Metrica: agendamentos e comparecimento.", certainty.inferido],
];

const positionings = [
  ["Ferias planejadas em familia", "Planejar as ferias da familia com mais previsibilidade, estrutura e menos improviso.", "Familias classe A/B com filhos ou netos", "Instagram, Meta, YouTube, WhatsApp", "Simular perfil de ferias", "Parecer pacote turistico comum", "Sua familia nao precisa decidir ferias todo ano do zero.", "Ferias recorrentes em resort, explicadas com clareza antes de qualquer proposta.", "Me conta como sua familia costuma viajar para eu te mostrar se esse modelo faz sentido.", "Topo de funil, publico familiar, produto com estrutura kids/parque", "Alto luxo individual ou comprador que busca investimento"],
  ["Casa de ferias com estrutura de resort", "Ter um lugar de retorno, com lazer, servico e estrutura de resort.", "Familias que sonham com segunda casa, mas nao querem cuidar de tudo", "Google Search, landing, YouTube, Meta", "Receber guia da casa de ferias em resort", "Confundir com propriedade integral ou omitir custos", "Casa de ferias em resort: antes de comprar, entenda as regras.", "A casa de ferias sem romantizar a manutencao, as taxas e o calendario.", "Vou te mostrar o que e uso, o que e custo recorrente e o que depende do contrato.", "Leads de busca ativa por casa de praia/resort", "Clube de ferias sem direito imobiliario claro"],
  ["Segunda residencia sem a dor de manter uma casa sozinho", "Acessar a logica de segunda casa com gestao, manutencao e uso mais organizado.", "Publico premium, empresarios, profissionais liberais", "LinkedIn, Google, YouTube, eventos, corretor", "Solicitar dossie privado", "Prometer zero preocupacao sem regra contratual", "Comprar uma casa de praia e facil. Manter bem todos os anos e outra historia.", "Uma alternativa para quem quer uso recorrente sem administrar uma casa sozinho.", "A parte mais importante e entender o que a operacao cuida e o que continua sendo custo do proprietario.", "Produto premium com gestao real", "Oferta popular baseada em preco"],
  ["Acesso a destinos e experiencias premium", "Acesso recorrente a destinos, hospitalidade, servico e experiencias selecionadas.", "Casais, viajantes recorrentes, alta renda, membros de clube", "Instagram, PR, YouTube, email, app", "Conhecer beneficios e destinos", "Ficar abstrato e virar luxo vazio", "Nao e so viajar mais. E viajar com padrao, destino e suporte.", "Acesso a destinos e experiencias com curadoria, regras claras e atendimento consultivo.", "Se voce gosta de variar destino, precisamos olhar flexibilidade antes de qualquer proposta.", "Vacation club, destination club, branded residence", "Produto fixo sem rede ou flexibilidade"],
  ["Imobiliario turistico com hospitalidade e servico", "Unir propriedade/uso imobiliario com padrao de hospitalidade, operacao e atendimento.", "Compradores patrimoniais, premium imobiliario, alto luxo", "Site, SEO, broker, evento, PR, Google", "Agendar apresentacao consultiva", "Soar como investimento financeiro", "O valor nao esta so no imovel. Esta na operacao que sustenta a experiencia.", "Imobiliario turistico com servico, destino e transparencia operacional.", "Antes de falar de valores, vamos separar uso, contrato, servico, custos e expectativa patrimonial.", "Branded residence, residence club, condo-hotel com juridico validado", "Campanha fria de massa e produto sem documentos"],
];

const funnel = [
  ["Descoberta", "Parar a rolagem e nomear uma dor", "Reels, TikTok, Shorts, Meta video", "Mitos, erros, comparativos curtos, destino", "Comentar/Salvar/Ver guia", "Ferias nao precisam ser improvisadas todo ano.", "Nao entendo esse mercado", "30 videos curtos", "Comecar por preco/cota", "Educacao"],
  ["Educacao", "Explicar modelos e reduzir confusao", "YouTube, blog, carrossel, landing", "Multipropriedade x clube x resort x casa", "Baixar guia", "Antes de comprar, entenda o que voce esta comprando.", "Isso e timeshare?", "Guia e FAQ", "Falar bonito sem explicar regra", "Comparacao"],
  ["Comparacao", "Ajudar escolha racional", "SEO, YouTube, carrossel, email", "Hotel x Airbnb x casa de praia x resort", "Fazer quiz", "O melhor modelo depende de como sua familia viaja.", "Vale a pena?", "Tabela comparativa", "Comparar so preco", "Confianca"],
  ["Confianca", "Provar seguranca operacional", "Site, remarketing, WhatsApp, YouTube", "Documentos, obra, taxa, reserva, depoimento", "Ver dossie", "Contrato, custo e disponibilidade precisam estar claros.", "Medo de golpe", "Dossie tecnico", "Esconder pergunta dificil", "Qualificacao"],
  ["Qualificacao", "Descobrir fit e temperatura", "WhatsApp, formulario, CRM", "Perguntas de perfil de uso", "Simular perfil", "Nao e para todo mundo; precisa caber na rotina.", "Minha familia usaria?", "Quiz + CRM tags", "Mandar tabela cedo", "Apresentacao"],
  ["Apresentacao", "Mostrar modelo e proposta com contexto", "Call, showroom, video consultivo", "Destino, estrutura, contrato, custos, uso", "Agendar visita", "Primeiro modelo, depois tabela.", "Nao entendi custos", "Deck comercial", "Pressao de fechamento", "Visita/showroom"],
  ["Visita/showroom", "Transformar abstrato em experiencia", "Showroom, tour, open house, video", "Unidade, resort, maquete, documentos", "Receber proposta", "Veja como isso funcionaria na vida real.", "Preciso ver para confiar", "Checklist de visita", "Tour sem diagnostico", "Proposta"],
  ["Proposta", "Personalizar valor e condicoes", "WhatsApp, CRM, email, consultor", "Simulacao, tabela, custos recorrentes, prazos", "Reuniao decisores", "A proposta precisa conversar com seu perfil de uso.", "Preco/conjuge", "Proposta 1-pagina", "PDF confuso", "Fechamento"],
  ["Fechamento", "Reduzir risco percebido e formalizar", "Consultor, juridico, CRM", "Contrato, FAQ final, prazo real", "Assinar/reservar", "Decida sabendo exatamente o que esta assumindo.", "Liquidez, contrato, arrependimento", "Checklist final", "Urgencia falsa", "Pos-venda"],
  ["Pos-venda", "Evitar arrependimento e ativar uso", "App, email, WhatsApp, onboarding", "Calendario, documentos, taxas, reserva", "Ativar conta", "Compra boa e compra usada.", "E agora?", "Onboarding 30 dias", "Sumir apos venda", "Indicacao"],
  ["Indicacao", "Transformar cliente em prova", "App, WhatsApp, NPS, eventos", "Depoimento, convite, bonus, experiencia", "Indicar familia/amigo", "Quem usa bem explica melhor que qualquer anuncio.", "Tenho alguem para indicar?", "Programa indicacao", "Pedir indicacao sem satisfacao", "Novo lead"],
];

const reels = [
  "Se voce acha que multipropriedade e tudo igual, veja isso antes de decidir.",
  "A pergunta certa nao e quanto custa a cota. E como sua familia realmente viaja.",
  "Casa de praia parece sonho ate a manutencao virar rotina.",
  "Antes de comprar qualquer casa de ferias em resort, pergunte isso.",
  "Hotel, Airbnb, casa de praia ou resort: o que faz sentido para sua familia?",
  "O erro de comprar ferias pela emocao do fim de semana.",
  "Taxa de manutencao nao e detalhe. E parte da decisao.",
  "Ferias em familia precisam de uma coisa que quase ninguem planeja.",
  "Quando uma casa de ferias em resort pode fazer sentido?",
  "Quando esse modelo nao faz sentido nenhum?",
  "Por que vendedor serio fala de custo antes de falar de desconto?",
  "O que precisa estar claro antes de falar em investimento turistico.",
  "Multipropriedade explicada sem enrolacao em 45 segundos.",
  "Clube de ferias e multipropriedade nao sao a mesma coisa.",
  "Residence club: o nome bonito que precisa ser explicado direito.",
  "Se o vendedor prometeu retorno garantido, ligue o alerta.",
  "Tres perguntas para nao se arrepender depois da visita.",
  "O luxo de verdade nesse mercado nao e a piscina. E a operacao.",
  "Destino completo vende mais do que apartamento bonito.",
  "Por que parque e resort mudam o valor percebido?",
  "A diferenca entre comprar uma experiencia e comprar uma promessa.",
  "Como saber se sua familia realmente usaria todo ano?",
  "A conversa que voce precisa ter com o conjuge antes da proposta.",
  "O que perguntar sobre revenda antes de assinar.",
  "O que perguntar sobre disponibilidade na alta temporada.",
  "O que perguntar sobre heranca e transferencia.",
  "O que perguntar sobre locacao da semana.",
  "O que perguntar sobre contrato e documentos.",
  "Um resort imobiliario serio precisa responder isso.",
  "Nao compre pressa. Compre clareza.",
];

const carousels = [
  "7 perguntas antes de comprar uma casa de ferias em resort",
  "Multipropriedade, clube de ferias e residence club: diferencas reais",
  "Hotel x Airbnb x casa de praia x resort: compare antes de decidir",
  "O que entra na taxa de manutencao?",
  "Como funciona a reserva de datas?",
  "Quando uma casa de ferias em resort nao faz sentido?",
  "O que um contrato serio precisa explicar",
  "Por que promessa de retorno garantido e sinal de alerta",
  "Como calcular se sua familia usaria esse modelo",
  "Destino, estrutura ou preco: o que deve pesar mais?",
  "Guia de termos: cota, fração, semana, clube, pool, uso",
  "Checklist para visitar showroom",
  "Como conversar com o conjuge sobre a decisao",
  "O que e prova real nesse mercado?",
  "Como comparar custo de ferias em 5 anos",
  "Casa de praia propria: vantagens e dores",
  "Resort com parque: por que muda a percepcao de valor",
  "O que perguntar sobre revenda e liquidez",
  "O que perguntar sobre locacao e renda",
  "Glossario do imobiliario turistico premium",
];

const storySequences = [
  "Enquete: sua familia viaja todo ano para o mesmo destino? > Caixa: qual maior duvida sobre casa de ferias? > CTA guia",
  "Bastidor do resort > detalhe da estrutura kids > pergunta: isso mudaria suas ferias?",
  "Antes/depois: reservar hotel todo ano x ter plano de ferias > CTA quiz",
  "3 perguntas sobre taxa > sticker de duvida > resposta em video",
  "Tour rapido: entrada, lazer, unidade > CTA WhatsApp",
  "Mito ou verdade: multipropriedade e clube sao iguais? > explicacao > guia",
  "Cliente pergunta sobre investimento > resposta segura > aviso de compliance",
  "Quiz de perfil de uso > resultado > consultor",
  "Bastidor de obra/documento > prova > CTA dossie",
  "Comparativo hotel x resort > enquete > carrossel completo",
  "Caixa: o que te daria medo nesse mercado? > respostas > FAQ dificil",
  "Mostre destino > mapa > tempo de deslocamento > CTA visita",
  "Mostre taxa como tema > transparencia > falar com consultor",
  "Showroom day > agenda > convite leve",
  "Depoimento de uso > contexto > sem promessa financeira",
  "Erro comum do comprador > pergunta > link do guia",
  "Conjuge decide junto? > material para decisores > CTA PDF",
  "Alta temporada > regras de reserva > cuidado com disponibilidade",
  "Pos-venda/app > calendario > prova de operacao",
  "Indicacao > familia usando > convite para conhecer",
];

const lives = [
  "Multipropriedade vale a pena? Depende do seu perfil de viagem",
  "Casa de praia, hotel, Airbnb ou resort: como escolher",
  "O que perguntar antes de comprar uma casa de ferias em resort",
  "Taxas, reserva e contrato: a parte que ninguem pode ignorar",
  "Como funciona residence club na pratica",
  "Ferias em familia: planejamento, custo e previsibilidade",
  "Tour comentado do empreendimento e do destino",
  "Conversa com consultor: objecoes reais sem roteiro decorado",
  "O que nao pode ser prometido nesse mercado",
  "Pos-venda: como usar bem depois de comprar",
];

const safeCtas = [
  "Receber guia", "Simular meu perfil de uso", "Falar com consultor", "Ver FAQ dificil", "Agendar apresentacao", "Conhecer o destino", "Ver comparativo", "Tirar duvida no WhatsApp", "Solicitar dossie", "Ver calendario de uso"
];
const riskyPhrases = [
  "Retorno garantido", "Renda mensal garantida", "Valorizacao certa", "Nunca mais pague hotel", "Sem taxa", "Disponibilidade garantida sempre", "Ultima chance sem prova", "Investimento seguro", "O melhor do Brasil sem fonte", "Compra sem risco"
];

const instagramCalendar = Array.from({ length: 30 }, (_, i) => {
  const day = i + 1;
  const format = day % 5 === 0 ? "Carrossel" : day % 4 === 0 ? "Stories" : day % 7 === 0 ? "Live" : "Reels";
  const idea = format === "Reels" ? reels[(day - 1) % reels.length] : format === "Carrossel" ? carousels[(day - 1) % carousels.length] : format === "Stories" ? storySequences[(day - 1) % storySequences.length] : lives[(day - 1) % lives.length];
  const cta = day <= 10 ? "Receber guia" : day <= 20 ? "Simular perfil" : "Falar com consultor";
  return [day, "Instagram", format, idea, cta, "Salvar, responder ou clicar", status];
});

const tiktokHooks = [
  ...reels,
  "Se voce vai comprar ferias planejadas, nao comece pela tabela.",
  "Esse e o detalhe que separa resort serio de venda empurrada.",
  "Voce sabe a diferenca entre usar, possuir e pertencer a um clube?",
  "Antes de entrar em uma apresentacao de vendas, veja isso.",
  "A maior armadilha e comprar destino que sua familia nao repete.",
  "O que um consultor deveria perguntar antes de te oferecer uma cota?",
  "O que acontece depois que voce compra? Essa e a pergunta certa.",
  "Nao existe modelo bom para todo mundo. Existe perfil certo.",
  "Se a proposta depende so de desconto, cuidado.",
  "A taxa anual precisa entrar na conta desde o primeiro dia.",
  "A promessa de economia pode esconder uma conta incompleta.",
  "O que ninguem explica sobre alta temporada.",
  "Por que algumas pessoas odeiam esse mercado?",
  "Como reconhecer uma venda agressiva em 20 segundos.",
  "O que olhar no contrato antes da emocao da visita.",
  "O vendedor falou em renda? Pergunte isso.",
  "Se voce ama viajar, talvez posse nao seja a melhor pergunta.",
  "Servico e manutencao sao a parte invisivel do produto.",
  "O que o resort precisa provar alem das fotos bonitas.",
  "A decisao nao e comprar ferias. E comprar compromisso.",
].slice(0, 50);

function makeRows(items, type, channel, cta = "Receber guia") {
  return items.map((item, idx) => [idx + 1, type, channel, item, cta, "Objecao/desejo real", "Cuidado com promessa exagerada", status]);
}

const tiktokScripts = tiktokHooks.slice(0, 20).map((hook, idx) => [
  idx + 1,
  hook,
  `Cena 1: fale o gancho olhando para a camera. Cena 2: mostre um exemplo concreto. Cena 3: explique a pergunta que o comprador deve fazer. Cena 4: convide para guia/WhatsApp.`,
  "Texto na tela: antes de comprar, entenda o modelo.",
  idx % 2 === 0 ? "Receber guia" : "Comentar DUVIDA",
  "Retencao, comentario, compartilhamento",
]);
const tiktokComparatives = [
  "Hotel x casa de ferias em resort", "Airbnb x resort com servico", "Casa de praia x residence club", "Clube de ferias x multipropriedade", "Pacote anual x uso recorrente", "Comprar por preco x comprar por perfil", "Baixa temporada x alta temporada", "Destino unico x rede de destinos", "Propriedade integral x fracionada", "Parque + resort x hotel isolado",
].map((x, i) => [i + 1, "Comparativo", x, "Mostrar vantagens, limites e pergunta final.", "Ver comparativo completo", "Nao fingir que uma opcao e sempre melhor"]);
const tiktokObjections = [
  "Isso e timeshare?", "Tem taxa escondida?", "Vou conseguir usar nas datas que quero?", "E investimento?", "Posso revender?", "Posso alugar?", "E se minha familia nao usar?", "Por que nao hotel?", "Tenho medo de golpe", "Meu conjuge nao vai aceitar",
].map((x, i) => [i + 1, "Objecao", x, "Responder com clareza e prometer envio de material.", "Ver FAQ", "Nao ridicularizar a duvida"]);
const tiktokMistakes = [
  "Comprar pela emocao da visita", "Nao calcular taxa anual", "Nao ler regra de reserva", "Confundir clube com propriedade", "Acreditar em retorno garantido", "Ignorar conjuge", "Nao pensar no destino em 5 anos", "Nao perguntar sobre revenda", "Nao entender manutencao", "Comparar so com diaria de hotel",
].map((x, i) => [i + 1, "Erro comprador", x, "Abrir alerta e ensinar a corrigir.", "Baixar checklist", "Sem tom apocaliptico"]);

const youtubeLongTitles = [
  "Multipropriedade vale a pena? O guia honesto para decidir sem pressao",
  "Casa de praia, hotel, Airbnb ou resort: qual modelo combina com sua familia?",
  "Como funciona uma casa de ferias em resort na pratica",
  "Residence club explicado: uso, custos, contrato e cuidados",
  "O que perguntar antes de comprar uma cota de ferias",
  "Taxa de manutencao: a pergunta que precisa vir antes da proposta",
  "Por que promessa de retorno garantido e perigosa nesse mercado",
  "Tour completo: como avaliar resort, unidade, estrutura e destino",
  "Como comparar custo de ferias em 5 anos",
  "Clube de ferias x multipropriedade: diferencas reais",
  "O que e branded residence e por que marcas de luxo vendem isso",
  "Como reconhecer uma venda agressiva de multipropriedade",
  "Vale comprar para usar, alugar ou revender?",
  "Como conversar com o conjuge antes de comprar",
  "Alta temporada: como funcionam reservas e disponibilidade",
  "Documentos, contrato e regras: checklist antes de assinar",
  "O que acontece depois da compra: app, reserva, suporte e taxas",
  "Como vender imobiliario turistico sem parecer golpe",
  "O papel do destino na decisao de compra",
  "Luxo real x luxo generico em resort imobiliario",
];
const youtubeShorts = tiktokHooks.slice(0, 30).map((h, i) => [i + 1, "Short", h, "30-45s", "Abrir loop + explicar um ponto", i <= 10 ? "Receber guia" : "Ver video completo"]);
const youtubeScripts = [
  ["Multipropriedade vale a pena?", "Abertura: depende do perfil de uso. Bloco 1: o que e. Bloco 2: custos. Bloco 3: reserva. Bloco 4: riscos. Bloco 5: para quem e/nao e. CTA: checklist."],
  ["Casa de praia, hotel, Airbnb ou resort?", "Abertura: nao existe melhor universal. Compare manutencao, flexibilidade, servico, custo recorrente, previsibilidade e familia. CTA: simulador."],
  ["Tour completo", "Comece pelo destino, depois acesso, estrutura, unidade, lazer, servico, documentos, custos e proximo passo. Feche com FAQ dificil."],
  ["Video de objecoes", "Liste 10 objecoes reais. Responda uma por vez sem debochar. Diga quando precisa validar no contrato. CTA: WhatsApp consultivo."],
  ["SEO YouTube", "Titulo com pergunta real, thumbnail com tensao concreta, descricao com capitulos, links para guia, FAQ e pagina comparativa."],
];

const keywordGroups = [
  ["Modelo", "multipropriedade como funciona; residence club o que e; clube de ferias vale a pena", "Educacao", "Guia/FAQ", "Baixar guia"],
  ["Destino", "resort em [destino]; resort com parque aquatico; casa de ferias em [destino]", "Comercial", "Pagina de destino", "Falar consultor"],
  ["Comparacao", "hotel ou casa de praia; Airbnb ou resort; multipropriedade ou clube de ferias", "Comparacao", "Pagina comparativa", "Simular perfil"],
  ["Objecao", "taxa manutencao multipropriedade; multipropriedade e golpe; posso vender minha cota", "Confianca", "FAQ dificil", "Ver documentos"],
  ["Premium", "branded residences brasil; imobiliario turistico premium; segunda residencia em resort", "Premium", "Dossie privado", "Agendar apresentacao"],
];
const negativeKeywords = ["gratis", "emprego", "reclamacao sem marca", "download contrato", "leilao", "barato demais", "airbnb barato", "cupom", "vagas", "trabalho", "concurso", "financiamento sem entrada"];
const googleAds = Array.from({ length: 30 }, (_, i) => {
  const group = keywordGroups[i % keywordGroups.length];
  return [i + 1, group[0], `Entenda ${group[1].split(";")[0]}`, "Compare custos, regras e modelo antes de decidir.", group[3], group[4], "Sem promessa financeira"];
});

const metaCampaigns = [
  ["Video View - Descoberta", "Familias, viagens, resorts, destinos, engajados turismo", "Video mito/erro", "Voce nao precisa comprar uma promessa de ferias.", "Assistir/Salvar", "CPV, retencao, engajamento", "Nao segmentar atributo sensivel"],
  ["Lead Magnet - Educacao", "Lookalike/semelhante de leads qualificados e engajados", "Carrossel guia", "Entenda modelos antes de conversar com vendedor.", "Baixar guia", "CPL qualificado", "Formulario com consentimento"],
  ["WhatsApp - Qualificacao", "Engajados 30/60/90 dias e visitantes site", "Video consultor", "Simule se faz sentido para sua familia.", "Enviar WhatsApp", "Custo conversa qualificada", "Nao prometer retorno"],
  ["Remarketing - Objecoes", "Visitantes FAQ, video 50%, leads sem agendamento", "FAQ dificil", "Ainda em duvida? Veja taxas, uso e disponibilidade.", "Ver FAQ", "Agendamento", "Landing consistente"],
  ["Destino - Desejo", "Interesses destino, viagem familia, resort", "Tour destino", "Ferias em resort com estrutura para voltar todos os anos.", "Conhecer destino", "CTR e leads", "Sem exclusividade vaga"],
  ["Premium - Dossie", "Alta renda inferida por contexto, base propria, indicacao", "Video arquitetura/servico", "Residencia de ferias com servico e operacao.", "Solicitar dossie", "Reunioes qualificadas", "Cuidado com segmentacao sensivel"],
];
const metaCopies = Array.from({ length: 30 }, (_, i) => {
  const angles = ["familia", "objecao", "comparacao", "destino", "premium", "FAQ"];
  const angle = angles[i % angles.length];
  const text = {
    familia: "Planejar ferias em familia pode ser mais simples quando voce entende modelo, custos e regras antes da proposta.",
    objecao: "Antes de comprar uma casa de ferias em resort, veja as perguntas que todo consultor serio precisa responder.",
    comparacao: "Hotel, Airbnb, casa de praia ou resort? Compare antes de decidir pelo impulso.",
    destino: "Um destino completo muda a experiencia: lazer, estrutura, servico e retorno planejado.",
    premium: "Imobiliario turistico premium nao se vende com pressa. Primeiro, clareza; depois, proposta.",
    FAQ: "Taxa, reserva, revenda, uso e contrato: o FAQ que precisa vir antes da tabela.",
  }[angle];
  return [i + 1, angle, text, i % 3 === 0 ? "Receber guia" : i % 3 === 1 ? "Simular perfil" : "Falar no WhatsApp", "Seguro", "Usar com landing coerente"];
});

const landingSections = [
  ["Hero", "Sua casa de ferias em um resort, com regras claras para usar, planejar e voltar todos os anos.", "Entenda o modelo, os custos, as regras de reserva e quando esse tipo de produto faz sentido para sua familia.", "Receber guia / Simular perfil"],
  ["Promessa principal", "Ferias planejadas com estrutura de resort, sem romantizar custos, contrato ou disponibilidade.", "A promessa e clareza + experiencia, nao retorno financeiro.", "Ver como funciona"],
  ["Explicacao do modelo", "O produto pode ser multipropriedade, residence club, clube de ferias ou hotelaria. Cada um tem regras diferentes.", "Mostre um diagrama simples.", "Comparar modelos"],
  ["Para quem e", "Familias que viajam com frequencia, valorizam estrutura e querem planejar melhor as ferias.", "Incluir criterios de fit.", "Fazer quiz"],
  ["Para quem nao e", "Nao e para quem quer retorno garantido, disponibilidade total sem regra ou comprar sem ler contrato.", "Aumenta confianca.", "Ver checklist"],
  ["Como funciona", "Passo 1 perfil. Passo 2 modelo. Passo 3 regras. Passo 4 proposta. Passo 5 decisao.", "Sequencia consultiva.", "Falar com consultor"],
  ["Destino", "Explique acesso, atracoes, sazonalidade, familia, gastronomia e motivo de retorno.", "Destino precisa justificar recorrencia.", "Conhecer destino"],
  ["Estrutura resort", "Lazer, piscina, parque, kids, restaurantes, servico, seguranca e suporte.", "Provar com fotos reais.", "Ver tour"],
  ["Tipos de uso", "Uso proprio, ferias escolares, baixa temporada, convidados, regras de cessao/locacao quando permitido.", "Validar contrato.", "Ver regras"],
  ["Custos e taxas", "Aquisicao, manutencao, taxas, impostos, transferencia, uso opcional.", "Sem esconder custo recorrente.", "Ver custos"],
  ["Comparativo", "Hotel x Airbnb x casa de praia x multipropriedade x clube.", "Tabela franca.", "Baixar comparativo"],
  ["FAQ dificil", "Taxa, reserva, revenda, heranca, locacao, arrependimento, obra, contrato.", "Responder antes do vendedor.", "Abrir FAQ"],
  ["Provas", "Marca, entregas, obras, documentos, depoimentos, premios, operadora.", "Prova concreta, nao adjetivo.", "Ver provas"],
  ["Tour video", "Video 3-7 min: destino, unidade, lazer, servico, documentos e FAQ.", "Retencao e confianca.", "Assistir tour"],
  ["Documentos", "Memorial, contrato, regulamento, perguntas juridicas, politica de uso.", "Disponibilizar sob solicitacao.", "Solicitar dossie"],
  ["WhatsApp", "Conversa consultiva baseada em perfil de uso.", "Nao mandar tabela antes de diagnosticar.", "Abrir WhatsApp"],
  ["Formulario", "Nome, WhatsApp, cidade, frequencia de viagem, familia, destino, duvida principal.", "Microcopy: vamos indicar o modelo certo.", "Enviar"],
  ["SEO", "Textos sobre modelos, destino, comparativos e FAQ.", "Ajudar Google e usuario.", "Ler guia"],
  ["Compliance", "Sem promessa de retorno, renda, valorizacao ou disponibilidade absoluta.", "Rodape com disclaimer.", "Ver termos"],
  ["Proximo passo", "Receber guia, simular perfil ou agendar apresentacao.", "CTA por temperatura.", "Escolher caminho"],
];

const whatsFlows = [
  ["Lead frio", "Oi, [nome]. Antes de falar de proposta, posso entender como sua familia costuma viajar?", "Viaja todo ano? mesmo destino? filhos? alta temporada?", "frio, origem, perfil", "Frio", "Enviar guia", "Conseguiu ver o guia? A parte de custos e regras costuma esclarecer bastante.", "Ficou alguma duvida sobre modelo ou voce so esta pesquisando por enquanto?", "Quando pergunta regra/custo ou pede simulacao", "Guia comparativo", "Mandar tabela cedo"],
  ["Lead Instagram", "Vi que voce veio pelo Instagram. O conteudo era sobre qual duvida: taxa, modelo, destino ou medo de golpe?", "Qual post te chamou atencao?", "instagram, tema", "Morno", "Enviar conteudo relacionado", "Te mandei o material daquele tema. Quer que eu simule pelo seu perfil?", "Posso te mandar um resumo de 2 minutos por audio?", "Quando responde com contexto familiar", "Reels/carrossel relacionado", "Responder com mensagem robotica"],
  ["Lead TikTok", "Voce veio pelo video do 'antes de comprar'? Boa. Qual parte te preocupou mais?", "O que voce quer evitar nessa compra?", "tiktok, objecao", "Frio/morno", "FAQ dificil", "Aquele video tem uma continuação com checklist. Quer receber?", "Se preferir, te mando so as 7 perguntas principais.", "Quando pede checklist", "Checklist", "Entrar vendendo"],
  ["Lead Google", "Voce provavelmente esta pesquisando o modelo. Quer entender diferenca entre hotel, clube, multipropriedade e residence club?", "Qual termo voce pesquisou?", "google, intencao", "Morno", "Pagina comparativa", "Esse comparativo ajudou ou ficou alguma duvida de contrato/custo?", "Posso resumir o modelo que mais combina com seu caso.", "Quando pede preco ou modelo", "Comparativo SEO", "Ignorar intencao da busca"],
  ["Lead indicacao", "O [indicador] comentou que voce queria entender o modelo. Prefiro te explicar com calma para voce nao comparar coisa diferente.", "Voce conhece o empreendimento ou ainda esta entendendo?", "indicacao", "Morno/quente", "Apresentacao curta", "Conseguiu conversar com [indicador] sobre uso real?", "Tenho um resumo para voce e seu conjuge olharem juntos.", "Quando aceita call", "PDF 2 paginas", "Queimar indicacao com pressao"],
  ["Pediu preco", "Eu te passo valores, claro. Mas antes preciso entender qual modelo voce esta comparando, porque preco sem uso e custo recorrente engana.", "Qual seria seu uso ideal por ano?", "preco", "Morno", "Diagnosticar e enviar faixa", "A faixa fez sentido? Quer ver o que entra no custo recorrente?", "Posso te mostrar uma comparacao com hotel/casa de praia.", "Quando aceita simular", "Tabela + simulador", "Mandar preco seco"],
  ["Perguntou investimento", "Eu prefiro ser direto: nao trabalho com promessa de retorno garantido. Podemos analisar uso e parte patrimonial com documentos.", "Voce pensa em usar, alugar ou revender?", "investimento", "Morno", "Enviar compliance/FAQ", "A parte patrimonial ficou clara ou voce quer olhar documentos?", "Se sua decisao depende de retorno garantido, esse produto talvez nao seja o ideal.", "Quando quer documentos", "FAQ investimento", "Prometer renda"],
  ["Medo de golpe", "Essa preocupacao e justa. O melhor caminho e olhar documentos, regras, custos e historico antes de proposta.", "O que te deixaria mais seguro?", "medo", "Frio/morno", "Dossie de seguranca", "O material respondeu sua principal preocupacao?", "Posso te mostrar a lista de perguntas que voce deve fazer a qualquer vendedor.", "Quando pede prova", "Dossie tecnico", "Desqualificar medo"],
  ["Conjuge", "Faz sentido decidir junto. Posso te mandar um resumo simples para voces dois olharem sem pressa.", "Qual duvida seu conjuge provavelmente faria?", "conjuge", "Morno/quente", "PDF decisores", "Conseguiram olhar juntos?", "Se quiser, marco uma conversa curta para responder os dois.", "Quando marca reuniao", "PDF 2 paginas", "Forcar decisao individual"],
  ["Sumiu", "Oi, [nome]. Nao quero te pressionar. So queria entender se travou em preco, modelo ou momento.", "Qual desses pontos pesou mais?", "sumiu", "Frio", "Diagnosticar perda", "Posso te mandar so o material da duvida principal?", "Se nao for prioridade agora, tudo bem. Te envio um guia para quando retomar.", "Quando responde objecao", "Resumo por objecao", "Cobrar resposta"],
  ["Interessado visita", "Perfeito. Para a visita render, vou te mandar antes um checklist do que observar.", "Vai sozinho ou com decisor?", "visita", "Quente", "Agendar visita", "Confirmando: horario, local e perguntas que vamos responder.", "Quer incluir seu conjuge/filhos na visita?", "Quando confirma data", "Checklist visita", "Visita sem objetivo"],
  ["Pos-visita", "Depois da visita, qual ponto ficou mais forte: destino, estrutura, custo, contrato ou uso?", "O que ainda impede a decisao?", "pos-visita", "Quente", "Enviar proposta personalizada", "Posso organizar proposta focada nos pontos que voce levantou?", "Quer uma reuniao so para duvidas finais?", "Quando pede proposta", "Resumo pos-visita", "Mandar proposta generica"],
  ["Pos-proposta", "Revendo sua proposta, os pontos decisivos sao entrada, custo recorrente e calendario. Qual deles voce quer ajustar/entender melhor?", "Qual parte ficou menos clara?", "pos-proposta", "Quente", "Reuniao decisores", "Conseguiu comparar com as alternativas?", "Posso montar um cenario conservador de uso para voce.", "Quando discute condicao", "Simulacao", "Dar desconto sem diagnostico"],
  ["Lead perdido", "Quero registrar corretamente: voce nao avancou por preco, confianca, momento ou fit do produto?", "Pode me dizer o motivo principal?", "perdido", "Perdido", "Aprender e nutrir", "Obrigado pela sinceridade. Posso te manter em uma lista com conteudo educativo?", "Quando fizer sentido, te mando atualizacoes sem pressao.", "Quando aceita nutricao", "Pesquisa perda", "Insistir demais"],
  ["Cliente comprador", "Parabens pela decisao. Agora o foco e garantir que voce use bem. Vou te mandar onboarding com app, documentos, calendario e suporte.", "Quem mais da familia precisa receber acesso?", "cliente", "Cliente", "Onboarding", "Conseguiu acessar app e documentos?", "Quer ajuda para planejar o primeiro uso?", "Quando ativa app", "Onboarding app", "Sumir apos venda"],
  ["Indicacao", "Voce conhece alguem que tambem pensa em ferias planejadas, mas tem medo desse mercado? Posso enviar um guia educativo, sem abordagem agressiva.", "Quem faria sentido receber?", "indicacao", "Cliente", "Pedir indicacao", "Obrigado pela indicacao. Vou abordar com cuidado e contexto.", "Quer que eu envie primeiro o guia para voce encaminhar?", "Quando NPS alto", "Programa indicacao", "Pedir indicacao antes de satisfacao"],
];

const salesScriptTopics = [
  "Abordagem inicial", "Diagnostico", "Apresentacao do modelo", "Explicacao de custos", "Explicacao de uso", "Explicacao de contrato", "Comparacao com hotel", "Comparacao com Airbnb", "Comparacao com casa de praia", "Objecao de preco", "Objecao medo de golpe", "Objecao de conjuge", "Objecao de investimento", "Objecao de liquidez", "Convite para showroom", "Pos-visita", "Fechamento consultivo", "Recuperacao de lead",
];
const salesScripts = salesScriptTopics.map((topic, i) => [
  i + 1,
  topic,
  `Abertura: vamos tratar ${topic.toLowerCase()} com clareza, sem pressa. Pergunta-chave: o que voce precisa entender para decidir com seguranca? Explicacao: separe desejo, regra, custo e risco. Fechamento: o proximo passo so faz sentido se esse ponto ficou claro.`,
  i < 3 ? "Descobrir fit" : i < 9 ? "Educar/comparar" : i < 14 ? "Resolver objecao" : "Avancar etapa",
  "Nao prometer retorno, nao esconder custo, nao pressionar decisor ausente.",
]);

const appRows = [
  ["Home", "Resumo de proximas datas, documentos pendentes, beneficios e suporte", "Cliente comprador", "Ativacao e reducao de ansiedade", "Login ativo, documentos vistos"],
  ["Calendario de uso", "Datas, regras, disponibilidade e lembretes", "Cliente e consultor", "Provar operacao antes e depois da venda", "Reservas iniciadas"],
  ["Reserva", "Solicitar/acompanhar reserva conforme regras", "Cliente", "Transformar compra em uso", "Taxa de reserva concluida"],
  ["Documentos", "Contrato, regulamento, taxas, comunicados", "Cliente e lead quente", "Transparencia", "Downloads/visualizacoes"],
  ["Taxas", "Historico, vencimentos, explicacao do que compoe", "Cliente", "Reduzir reclamacao", "Pagamentos em dia"],
  ["Beneficios", "Parque, resort, clube, parceiros, upgrades", "Cliente", "Aumentar valor percebido", "Uso de beneficios"],
  ["Concierge", "Dúvidas, suporte, solicitações especiais", "Cliente premium", "Servico e fidelizacao", "Tempo resposta/NPS"],
  ["Indicacao", "Enviar guia para amigo, acompanhar status", "Cliente satisfeito", "Gerar leads com contexto", "Indicacoes qualificadas"],
  ["NPS", "Pesquisa apos uso/reserva/atendimento", "Gestor", "Detectar arrependimento cedo", "NPS, motivos"],
  ["Conteudo pos-compra", "Como usar bem, regras, destino, boas praticas", "Cliente", "Educacao e menos suporte repetitivo", "Conteudo visto"],
  ["Prova comercial", "Demonstração do app no showroom", "Lead quente", "Mostrar pos-venda antes da compra", "Conversao pos-demo"],
];
const notifications = [
  ["D+1 compra", "Seu onboarding esta disponivel: documentos, calendario e proximos passos.", "Ativar cliente"],
  ["30 dias antes uso", "Hora de revisar sua reserva e beneficios para a proxima viagem.", "Aumentar uso"],
  ["Taxa proxima", "Sua taxa recorrente vence em X dias. Veja detalhes do que ela cobre.", "Reduzir inadimplencia"],
  ["NPS pos-uso", "Como foi sua experiencia? Sua resposta ajuda a melhorar o servico.", "Coletar prova e risco"],
  ["Indicacao NPS alto", "Conhece alguem que deveria entender esse modelo sem pressao?", "Gerar indicacao"],
];

const headlineBase = [
  "Sua casa de ferias em resort, explicada sem pressa e sem promessa vazia",
  "Antes de comprar uma cota de ferias, entenda o modelo",
  "Ferias planejadas para familias que nao querem improvisar todo ano",
  "Casa de praia, hotel ou resort? Compare antes de decidir",
  "Multipropriedade sem confusao: custos, uso e regras",
  "O guia claro para entender imobiliario turistico premium",
  "Resort, familia e previsibilidade: o que realmente esta sendo vendido",
  "Nao compre uma promessa de ferias. Compre clareza",
  "Sua segunda residencia pode ter servico, estrutura e regras claras",
  "O luxo de chegar e encontrar tudo pronto",
];
const headlines = Array.from({ length: 50 }, (_, i) => {
  const suffixes = ["para sua familia", "antes da proposta", "sem cair em promessa de retorno", "com comparativo real", "com FAQ dificil"];
  return [i + 1, "Headline", ["Descoberta", "Educacao", "Comparacao", "Confianca", "Desejo"][i % 5], `${headlineBase[i % headlineBase.length]} ${suffixes[Math.floor(i / 10)]}.`, "Site/Ads/Email", "Seguro com contexto"];
});
const copyHooks = tiktokHooks.map((h, i) => [i + 1, "Gancho Reels/TikTok", ["Descoberta", "Educacao", "Objecao", "Comparacao", "Desejo"][i % 5], h, "Reels/TikTok/Shorts", "Sem claim financeiro"]);
const ctas = Array.from({ length: 30 }, (_, i) => [i + 1, "CTA", ["frio", "morno", "quente"][i % 3], safeCtas[i % safeCtas.length], "Site/Ads/WhatsApp", "Ajustar ao funil"]);
const remarketingCopies = Array.from({ length: 20 }, (_, i) => [i + 1, "Remarketing", ["Confianca", "Comparacao", "Simulacao", "Visita"][i % 4], `Ainda em duvida sobre ${["taxa", "reserva", "contrato", "modelo", "destino"][i % 5]}? Veja o material antes de conversar com um consultor.`, "Meta/Google/Email", "Nao pressionar"]);
const objectionAnswers = tiktokObjections.concat(tiktokMistakes).slice(0, 20).map((r, i) => [i + 1, "Resposta objecao", "Confianca", `${r[2]}: resposta curta com documento, regra e proximo passo seguro.`, "WhatsApp/Site", "Validar contrato"]);
const seoThemePool = [
  "Multipropriedade: como funciona, custos e cuidados",
  "Casa de ferias em resort: modelos, taxas e regras",
  "Residence club: o que e e quando faz sentido",
  "Clube de ferias vale a pena? Depende do seu perfil",
  "Taxa de manutencao em multipropriedade: o que perguntar",
  "Hotel, Airbnb ou resort: comparativo para familias",
  "Segunda casa de praia ou resort com servico",
  "Branded residences no Brasil: conceito e cuidados",
  "Como escolher um resort imobiliario com seguranca",
  "FAQ dificil antes de comprar uma cota de ferias",
  ...youtubeLongTitles,
];
const seoThemes = Array.from({ length: 20 }, (_, i) => [i + 1, "Tema SEO", ["Educacao", "Comparacao", "Objecao", "Destino"][i % 4], seoThemePool[i], "Blog/YouTube", "Conteudo util"]);
const copyBank = [
  ...headlines,
  ...copyHooks,
  ...ctas,
  ...metaCopies.map((r) => [r[0], "Copy Meta Ads", r[1], r[2], "Meta", r[4]]),
  ...googleAds.map((r) => [r[0], "Copy Google Ads", r[1], `${r[2]} - ${r[3]}`, "Google", r[6]]),
  ...remarketingCopies,
  ...whatsFlows.slice(0, 20).map((r, i) => [i + 1, "Mensagem WhatsApp", r[0], r[1], "WhatsApp", r[10]]),
  ...objectionAnswers,
  ...carousels.map((c, i) => [i + 1, "Tema carrossel", ["Educacao", "Comparacao", "Confianca"][i % 3], c, "Instagram", "Salvar/compartilhar"]),
  ...youtubeLongTitles.map((t, i) => [i + 1, "Tema YouTube", ["Educacao", "Comparacao", "Confianca"][i % 3], t, "YouTube", "SEO + autoridade"]),
  ...seoThemes,
];

const complianceClaims = [
  ["Investimento", "Invista com retorno garantido", "PROIBIDA", "Promete resultado financeiro sem base e alto risco regulatorio.", "Entenda o modelo, os riscos e o potencial sem garantia de retorno.", "Nenhum sem juridico", "Validar com juridico/CVM se houver pool ou renda"],
  ["Valorizacao", "Valorizacao certa no destino", "PROIBIDA", "Valorizacao nao e garantida.", "O valor pode variar conforme mercado, destino, operacao e liquidez.", "Site com disclaimer", "Precisa documentos e fontes"],
  ["Renda", "Ganhe renda mensal alugando sua semana", "PROIBIDA", "Promessa de renda e locacao depende de regra.", "Quando permitido, consulte regras de locacao, custos e riscos.", "FAQ/consultor", "Validar contrato"],
  ["Locacao", "Voce pode alugar sempre que quiser", "ARRISCADA", "Pode depender de regras.", "A possibilidade de locacao depende do contrato e disponibilidade.", "WhatsApp/FAQ", "Validar empreendimento"],
  ["Escassez", "Ultimas unidades hoje", "MODERADA", "So pode se estoque/prazo real.", "Condicao valida ate X, conforme disponibilidade real.", "Ads/site", "Registrar prova comercial"],
  ["Preco", "Mais barato que hotel", "ARRISCADA", "Comparacao incompleta.", "Compare custo total conforme perfil de uso.", "Comparativo", "Precisa calculadora"],
  ["Desconto", "Desconto imperdivel", "MODERADA", "Pode baixar premium e exigir regra.", "Condicao especial com prazo e regra informada.", "Ads", "Validar campanha"],
  ["Exclusividade", "Exclusivo para poucos", "MODERADA", "Vago sem criterio.", "Acesso limitado por unidades/beneficios/regras do clube.", "Premium", "Definir exclusivo"],
  ["Disponibilidade", "Use quando quiser", "PROIBIDA", "Disponibilidade depende de regras.", "Uso conforme calendario, reserva e disponibilidade contratual.", "Todos", "Validar contrato"],
  ["Economia", "Nunca mais pague hotel", "PROIBIDA", "Promessa absoluta e falsa para muitos perfis.", "Planeje ferias recorrentes comparando alternativas.", "Educacao", "Evitar absoluto"],
  ["Propriedade", "Sua propriedade garantida", "VALIDAR", "Depende do modelo juridico.", "Direitos e forma de uso conforme contrato/escritura aplicavel.", "Site/consultor", "Validar documentacao"],
  ["Direito de uso", "Voce tem direito de uso vitalicio", "VALIDAR", "Prazo pode variar.", "Direito de uso conforme regras e prazo contratual.", "FAQ", "Validar contrato"],
  ["Taxa", "Sem taxa", "ARRISCADA", "Pode ocultar custos.", "Veja custos de aquisicao, recorrentes e opcionais.", "Todos", "Transparencia obrigatoria"],
  ["Manutencao", "Sem preocupacao com manutencao", "MODERADA", "Gestao nao elimina custo.", "Manutencao administrada conforme regras e taxas.", "Site/ads", "Validar operacao"],
  ["Revenda", "Revenda facil", "PROIBIDA", "Liquidez nao garantida.", "Transferencia/revenda conforme regras, custos e demanda.", "FAQ", "Validar contrato"],
  ["Heranca", "Passa automaticamente para herdeiros", "VALIDAR", "Depende do direito e formalizacao.", "Sucessao/heranca devem ser analisadas conforme documentos e lei.", "Consultor/juridico", "Juridico obrigatorio"],
];

const plan30 = Array.from({ length: 30 }, (_, i) => {
  const d = i + 1;
  const week = d <= 7 ? "Semana 1 - Diagnostico e estrutura" : d <= 14 ? "Semana 2 - Conteudo e landing" : d <= 21 ? "Semana 3 - Trafego e CRM" : "Semana 4 - Otimizacao e vendas";
  const actions = [
    "Reunir contratos, taxas, regras de uso, estoque, fotos, videos, provas e fontes.",
    "Separar produto por tipo: hotelaria, clube, multipropriedade, residence, branded, condo.",
    "Validar claims com juridico e criar matriz proibida/segura.",
    "Definir posicionamento principal e publico prioritario.",
    "Mapear CRM: etapas, tags, temperaturas e materiais.",
    "Criar FAQ dificil e comparativo de modelos.",
    "Criar scripts de WhatsApp e venda para equipe.",
    "Escrever landing page com hero, modelo, custos, FAQ e formularios.",
    "Produzir 10 Reels/TikToks de objecao.",
    "Produzir 5 carrosseis educativos.",
    "Produzir 1 video YouTube longo e 5 Shorts.",
    "Criar lead magnet guia dos modelos.",
    "Criar quiz/simulador de perfil de uso.",
    "Configurar eventos, UTMs, WhatsApp tracking e CRM.",
    "Subir campanha Meta video view descoberta.",
    "Subir campanha Meta lead magnet.",
    "Subir Google Search por modelo e destino.",
    "Criar remarketing para visitantes e engajados.",
    "Treinar vendedores com scripts e roleplay.",
    "Publicar landing e testar mobile/velocidade.",
    "Rodar primeira leitura de CPL, CTR e conversas.",
    "Classificar leads por origem e objecao.",
    "Ajustar copies que geram lead ruim.",
    "Publicar videos respondendo objecoes reais.",
    "Criar campanha de WhatsApp para leads mornos.",
    "Criar campanha pos-visita e pos-proposta.",
    "Analisar show rate, proposta e perda por motivo.",
    "Otimizar landing com perguntas mais repetidas.",
    "Escalar criativos vencedores por segmento.",
    "Fechar relatorio executivo e proximos testes.",
  ];
  return [d, week, actions[i], d <= 14 ? "Marketing/Copy" : d <= 21 ? "Trafego/CRM" : "Gestor/Vendas", d <= 14 ? "Conteudo/pagina" : d <= 21 ? "Campanha/CRM" : "Relatorio/otimizacao", d <= 7 ? "Estruturar" : d <= 14 ? "Educar" : d <= 21 ? "Gerar lead" : "Converter", d <= 14 ? "Ativo aprovado" : "Lead qualificado / agendamento", status];
});

const workbook = Workbook.create();

function writeSheet(name, data, widths = []) {
  const sheet = workbook.worksheets.add(name);
  sheet.showGridLines = false;
  const cols = data[0].length;
  const normalized = data.map((row) => {
    if (row.length === cols) return row;
    if (row.length < cols) return [...row, ...Array(cols - row.length).fill("")];
    return [...row.slice(0, cols - 1), row.slice(cols - 1).join(" | ")];
  });
  const rows = normalized.length;
  const range = sheet.getRangeByIndexes(0, 0, rows, cols);
  range.values = normalized;
  sheet.freezePanes.freezeRows(1);
  range.format = { wrapText: true, borders: { preset: "all", style: "thin", color: "#D9D9D9" } };
  const header = sheet.getRangeByIndexes(0, 0, 1, cols);
  header.format = { fill: "#12355B", font: { bold: true, color: "#FFFFFF" }, wrapText: true };
  try {
    const lastCol = String.fromCharCode("A".charCodeAt(0) + cols - 1);
    const table = sheet.tables.add(`A1:${lastCol}${rows}`, true, name.replace(/[^A-Za-z0-9]/g, "_").slice(0, 30));
    table.style = "TableStyleMedium2";
    table.showFilterButton = true;
  } catch {}
  for (let c = 0; c < cols; c++) {
    sheet.getRangeByIndexes(0, c, rows, 1).format.columnWidthPx = widths[c] ?? (c === 0 ? 150 : 260);
  }
  sheet.getRangeByIndexes(1, 0, Math.max(rows - 1, 1), cols).format.rowHeightPx = 92;
}

writeSheet("00 Leia Primeiro", [
  ["Campo", "Conteudo"],
  ["Objetivo", "Plano operacional completo para transformar inteligencia de mercado em maquina de marketing, trafego, CRM, vendedor, app e metricas."],
  ["Regra central", "Nao vender luxo vazio, retorno garantido ou cota por pressa. Vender clareza, desejo, seguranca, uso real, destino, servico e proximo passo comercial."],
  ["Como usar", "Cada aba e feita para um responsavel: gestor, copy, social, trafego, CRM, vendedor, produto/app e juridico/compliance."],
  ["Status", status],
], [170, 900]);
writeSheet("01 Diagnostico", [["Pergunta", "Resposta operacional", "Por que importa", "Aplicacao", "Certeza"], ...diagnostics], [250, 430, 430, 480, 150]);
writeSheet("02 Posicionamentos", [["Posicionamento", "Promessa central", "Publico ideal", "Canal ideal", "Melhor CTA", "Risco", "Frase anuncio", "Frase site", "Frase WhatsApp", "Quando usar", "Quando evitar"], ...positionings], [220, 360, 280, 250, 220, 300, 360, 430, 430, 320, 300]);
writeSheet("03 Funil Completo", [["Etapa", "Objetivo", "Canal", "Conteudo", "CTA", "Metrica", "Mensagem", "Objecao", "Ativo", "Erro comum", "Proximo passo"], ...funnel], [150, 260, 260, 320, 210, 240, 360, 260, 260, 300, 220]);
writeSheet("04 Instagram", [["Item", "Tipo", "Canal", "Ideia/Sequencia", "CTA", "Metrica", "Risco/Status"], ...makeRows(reels, "Reels", "Instagram"), ...makeRows(carousels, "Carrossel", "Instagram"), ...makeRows(storySequences, "Stories", "Instagram"), ...makeRows(lives, "Live", "Instagram"), ...safeCtas.map((x, i) => [i + 1, "CTA seguro", "Instagram", x, "Usar conforme funil", "Clique/resposta", "Seguro"]), ...riskyPhrases.map((x, i) => [i + 1, "Frase proibida/arriscada", "Instagram", x, "Evitar", "Risco juridico/politica", "Arriscado"])], [80, 170, 150, 640, 200, 220, 260]);
writeSheet("05 Calendario IG 30d", [["Dia", "Canal", "Formato", "Tema", "CTA", "Metrica", "Status"], ...instagramCalendar], [70, 150, 150, 560, 190, 250, 220]);
writeSheet("06 TikTok", [["Item", "Tipo", "Formato", "Gancho/Tema", "CTA/Acao", "Metrica/Funcao", "Risco/Status", "Status"], ...makeRows(tiktokHooks, "Gancho", "TikTok"), ...tiktokScripts.map((r) => [r[0], "Roteiro curto", "Video 30-45s", `${r[1]} | ${r[2]} | ${r[3]}`, r[4], r[5], "Seguro se sem promessa financeira", status]), ...tiktokComparatives.map((r) => [r[0], r[1], "Comparativo", r[2], r[4], r[3], r[5], status]), ...tiktokObjections.map((r) => [r[0], r[1], "Objecao", r[2], r[4], r[3], r[5], status]), ...tiktokMistakes.map((r) => [r[0], r[1], "Erro comprador", r[2], r[4], r[3], r[5], status])], [80, 170, 180, 740, 220, 260, 280, 220]);
writeSheet("07 YouTube", [["Item", "Tipo", "Titulo/Tema", "Formato", "Objetivo/Roteiro", "CTA", "SEO/Obs"], ...youtubeLongTitles.map((t, i) => [i + 1, "Video longo", t, "8-18 min", "Autoridade, comparacao e confianca", "Guia/Simulador", "Titulo com pergunta real"]), ...youtubeShorts, ...youtubeScripts.map((r, i) => [i + 1, "Roteiro", r[0], "Video estruturado", r[1], "CTA conforme etapa", "Capitulos + descricao + links"])], [80, 160, 520, 160, 620, 200, 300]);
writeSheet("08 Google SEO Ads", [["Item", "Grupo/Tipo", "Termos/Headline", "Intencao/Descricao", "Pagina", "CTA", "Risco/Obs"], ...keywordGroups.map((r, i) => [i + 1, "Grupo keyword", r[0], r[1], r[3], r[4], "Separar por intencao"]), ...negativeKeywords.map((n, i) => [i + 1, "Negativa", n, "Evitar lead ruim", "Campanha", "N/A", "Adicionar na conta"]), ...googleAds], [80, 170, 430, 420, 240, 200, 280]);
writeSheet("09 Meta Ads", [["Item", "Campanha/Angulo", "Publico", "Criativo/Copy", "CTA", "KPI", "Compliance"], ...metaCampaigns.map((r, i) => [i + 1, r[0], r[1], `${r[2]} | ${r[3]}`, r[4], r[5], r[6]]), ...metaCopies.map((r) => [r[0], `Copy - ${r[1]}`, "Conforme conjunto", r[2], r[3], "CPL/conversa", r[5]])], [80, 230, 360, 560, 200, 220, 360]);
writeSheet("10 Landing Page", [["Secao", "Copy/Conteudo", "Microcopy/Prova", "CTA"], ...landingSections], [190, 600, 450, 220]);
writeSheet("11 WhatsApp CRM", [["Fluxo", "Primeira mensagem", "Pergunta qualificacao", "Tags", "Temperatura", "Proximo passo", "Follow-up", "Recuperacao", "Gatilho consultor", "Material apoio", "Erro proibido"], ...whatsFlows], [170, 520, 360, 220, 160, 260, 430, 430, 300, 220, 260]);
writeSheet("12 Scripts Venda", [["Item", "Momento", "Script base", "Funcao", "Erro a evitar"], ...salesScripts], [80, 240, 780, 230, 380]);
writeSheet("13 App Pos Venda", [["Tela/Funcao", "O que faz", "Usuario", "Como ajuda venda/pos-venda", "Metrica"], ...appRows, ...notifications.map((n) => [`Notificacao - ${n[0]}`, n[1], "Cliente", n[2], "Abertura/clique"])], [220, 560, 180, 420, 220]);
writeSheet("14 Banco Copy", [["Item", "Tipo", "Etapa/Angulo", "Texto", "Canal", "Obs/Risco"], ...copyBank], [80, 190, 190, 760, 180, 260]);
writeSheet("15 Compliance", [["Categoria", "Frase original", "Risco", "Motivo", "Versao segura", "Canal permitido", "Obs juridica"], ...complianceClaims], [170, 330, 170, 430, 520, 220, 330]);
writeSheet("16 Plano 30 Dias", [["Dia", "Semana", "Acao", "Responsavel", "Ativo produzido", "Objetivo", "Metrica", "Status"], ...plan30], [70, 270, 620, 190, 220, 190, 240, 220]);
writeSheet("17 Metricas", [["Area", "Metrica", "Por que importa", "Ferramenta", "Decisao de 48h"], ["Conteudo", "Retencao 3s/50%/100%", "Mostra se gancho segura", "Instagram/TikTok/YouTube", "Regravar gancho se queda inicial"], ["Conteudo", "Salvos/Compartilhamentos", "Mostra utilidade e identificacao", "Instagram", "Transformar em carrossel/lead magnet"], ["Trafego", "CPL qualificado", "Lead barato sem fit nao serve", "Meta/Google/CRM", "Pausar angulo ruim"], ["WhatsApp", "Taxa de resposta", "Mede qualidade do lead e primeira mensagem", "CRM/WhatsApp", "Ajustar abordagem"], ["Vendas", "Show rate", "Mede qualificacao e compromisso", "CRM", "Criar confirmacao e material pre-call"], ["Vendas", "Proposta enviada", "Mede avanco real", "CRM", "Treinar diagnostico"], ["Vendas", "Motivo de perda", "Alimenta copy e produto", "CRM", "Criar conteudo/ajuste oferta"], ["Pos-venda", "Ativacao app", "Reduz arrependimento", "App", "Onboarding humano"], ["Indicacao", "NPS + indicacoes", "Prova recorrencia", "App/CRM", "Pedir indicacao so NPS alto"]], [180, 240, 430, 220, 480]);
writeSheet("18 Fontes", [["Fonte", "URL", "Uso", "Certeza"], ...sources], [260, 520, 620, 160]);

const errors = await workbook.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
  options: { useRegex: true, maxResults: 100 },
  summary: "formula error scan",
});

for (const sheetName of ["02 Posicionamentos", "03 Funil Completo", "11 WhatsApp CRM", "16 Plano 30 Dias"]) {
  const img = await workbook.render({ sheetName, autoCrop: "all", scale: 1, format: "png" });
  await fs.writeFile(path.join(outputDir, `${sheetName.replaceAll(" ", "_")}.png`), new Uint8Array(await img.arrayBuffer()));
}

const xlsxPath = path.join(outputDir, "Plano_Operacional_Maquina_Vendas_Resorts.xlsx");
const xlsx = await SpreadsheetFile.exportXlsx(workbook);
await xlsx.save(xlsxPath);

const md = `# Plano Operacional - Maquina de Marketing e Vendas para Imobiliario Turistico Premium

Status: ${status}  
Data: 2026-06-18  
Objetivo: transformar a base de inteligencia em execucao de marketing, trafego, CRM, venda, pos-venda e metricas.

## Diagnostico central

O produto real nao e "cota". O produto real e ferias recorrentes com clareza de uso, estrutura de resort, servico, destino, previsibilidade e, quando aplicavel, direito imobiliario ou clube bem explicado.

O que nao deve ser vendido como promessa principal:

- retorno financeiro;
- renda garantida;
- valorizacao;
- economia absoluta;
- disponibilidade total;
- exclusividade vaga;
- luxo generico.

O que deve ser vendido:

- clareza;
- uso real;
- familia;
- destino;
- estrutura;
- servico;
- menos improviso;
- menos dor de manter uma segunda casa;
- transparencia de custos, taxas, contrato e regras.

## Como usar a planilha

1. O gestor le as abas 01, 02 e 03 para escolher posicionamento e funil.
2. Social usa as abas 04, 05, 06 e 07 para conteudo.
3. Trafego usa 08 e 09.
4. Copy/site usa 10 e 14.
5. CRM e vendas usam 11 e 12.
6. Produto/pos-venda usa 13.
7. Juridico/compliance usa 15.
8. Operacao acompanha 16 e 17.

## Posicionamento recomendado para iniciar

Comecar com dois angulos em paralelo:

1. **Ferias planejadas em familia** para volume qualificado.
2. **Casa de ferias com estrutura de resort** para leads com desejo mais claro.

Nao iniciar pelo angulo "investimento". Esse angulo deve ser tratado apenas em conteudo explicativo, com disclaimers e validacao juridica.

## Funil recomendado

Descoberta > Educacao > Comparacao > Confianca > Qualificacao > Apresentacao > Visita/showroom > Proposta > Fechamento > Pos-venda > Indicacao.

O erro mais caro e levar o lead para vendedor antes de educar. O vendedor entra melhor quando a pessoa ja entendeu o modelo, sabe que existem custos e tem uma duvida concreta.

## Landing page base

H1:
> Sua casa de ferias em um resort, com regras claras para usar, planejar e voltar todos os anos.

Subtitulo:
> Entenda o modelo, os custos, as regras de reserva e quando esse tipo de produto faz sentido para sua familia.

CTA frio:
> Receber guia dos modelos de ferias em resort

CTA morno:
> Simular meu perfil de uso

CTA quente:
> Falar com consultor

## WhatsApp principal

> Oi, [nome]. Antes de falar de proposta, posso entender como sua familia costuma viajar? Assim eu nao te mostro um modelo que nao faz sentido para voce.

## Regra de compliance

Qualquer frase sobre investimento, renda, valorizacao, disponibilidade, revenda, heranca, taxa, manutencao ou propriedade precisa estar ligada a contrato, regra de uso e validacao juridica.

## Entregaveis incluidos na planilha

- 5 posicionamentos completos.
- Funil com 11 etapas.
- Estrategia de Instagram com 30 Reels, 20 carrosseis, 20 stories, 10 lives, CTAs e frases proibidas.
- Calendario de Instagram 30 dias.
- TikTok com 50 ganchos, 20 roteiros curtos, comparativos, objecoes e erros de comprador.
- YouTube com 20 videos longos, 30 Shorts e roteiros.
- Google com keywords, negativas, anuncios e SEO.
- Meta Ads com campanhas, publicos, copies e compliance.
- Landing page completa.
- 16 fluxos de WhatsApp/CRM.
- 18 scripts de venda.
- Estrategia de app/pos-venda.
- Banco de copy com headlines, ganchos, CTAs, ads, remarketing, WhatsApp, objecoes, carrosseis, YouTube e SEO.
- Matriz de compliance.
- Plano de 30 dias.
- Metricas de decisao.

## Fontes

${sources.map((s) => `- [${s[0]}](${s[1]}) - ${s[2]} (${s[3]})`).join("\n")}
`;

const mdPath = path.join(outputDir, "Plano_Operacional_Maquina_Vendas_Resorts.md");
await fs.writeFile(mdPath, md, "utf8");

console.log(errors.ndjson);
console.log(JSON.stringify({ outputDir, xlsxPath, mdPath }, null, 2));
