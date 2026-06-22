import fs from "node:fs/promises";
import path from "node:path";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const outputDir = path.resolve("outputs", "dossie_3_0_resorts_marketing");
await fs.mkdir(outputDir, { recursive: true });

const C = {
  comprovado: "COMPROVADO",
  inferido: "INFERIDO",
  hipotese: "HIPOTESE",
  validacao: "PRECISA VALIDACAO",
};

const sources = [
  ["GAV Resorts", "https://www.gavresorts.com.br/", "Site oficial auditado: casa de ferias, ferias compartilhadas, hotelaria, FAQ, Select Club, blog, WhatsApp e portal do cliente.", C.comprovado],
  ["Aviva", "https://www.aviva.com.br/", "Site oficial auditado: hoteis, Residence Clubs, Clube de Ferias, Rio Quente, Costa do Sauipe, Hot Park, newsletter e WhatsApp.", C.comprovado],
  ["Beach Park", "https://beachpark.com.br/", "Site oficial auditado: destino completo, parques, resorts, Vacation Club, Beach Card, blog, WhatsApp, redes sociais e portal do agente.", C.comprovado],
  ["Hot Beach", "https://hotbeach.com.br/", "Site oficial auditado: parque aquatico, resorts, ingresso, hospedagem, Residence Club, casa de ferias anexa ao parque e redes sociais.", C.comprovado],
  ["Four Seasons Private Residences", "https://www.fourseasons.com/residences/private_residences/", "Benchmark internacional auditado: private residences, residence clubs, property management, service, peace of mind, portfolio global.", C.comprovado],
  ["Ritz-Carlton Residences", "https://www.ritzcarlton.com/en/residences/", "Benchmark internacional auditado: luxury residences, Gold Standard, lifestyle, servico e reputacao.", C.comprovado],
  ["Marriott Vacation Club", "https://www.marriottvacationclubs.com/ownership/vacation-club-brands/marriott-vacation-club.html", "Benchmark auditado: vacation ownership, Request Information, disclosure de timeshare e oferta regulada.", C.comprovado],
  ["Hilton Grand Vacations", "https://www.hiltongrandvacations.com/en/discover-hilton-grand-vacations", "Benchmark auditado: vacation membership, real member stories, get started, direct sales, disclosures e faixa de precos em USD.", C.comprovado],
  ["Meta Advertising Standards", "https://transparency.meta.com/policies/ad-standards/", "Politica oficial: revisao de anuncios, landing page, praticas enganosas, atributos pessoais, discriminacao e transparencia da Ad Library.", C.comprovado],
  ["Google Ads Transparency Center", "https://adstransparency.google.com/", "Centro oficial de transparencia de anuncios do Google; acesso dinamico, precisa consulta manual por anunciante/dominio.", C.comprovado],
  ["Google Ads Personalized Advertising", "https://support.google.com/adspolicy/answer/143465", "Politica oficial de publicidade personalizada: categorias sensiveis, housing/finance em alguns mercados, dados pessoais e segmentacao.", C.comprovado],
  ["TikTok Misleading and False Content", "https://ads.tiktok.com/help/article/tiktok-ads-policy-misleading-and-false-content", "Politica oficial: claims exagerados, informacao inconsistente entre anuncio e landing, clickbait e comparativos sem prova.", C.comprovado],
  ["Google Helpful Content", "https://developers.google.com/search/docs/fundamentals/creating-helpful-content", "Diretriz oficial para conteudo util, confiavel e people-first.", C.comprovado],
  ["Lei 13.777/2018", "http://www.planalto.gov.br/ccivil_03/_ato2015-2018/2018/lei/L13777.htm", "Marco legal brasileiro da multipropriedade imobiliaria.", C.comprovado],
];

const segmentacao = [
  ["Popular / volume", "diRoma, WAM, parte de GAV, destinos de aguas termais e resorts promocionais", "Preco, parcelamento, beneficio tangivel, parque, crianca, familia, facil acesso", "Direta, promocional, simples, com prova de estrutura", "Baixar demais a percepcao; parecer cota empurrada", "Oferta clara, FAQ, condicao, WhatsApp rapido, video de experiencia", C.inferido],
  ["Premium familiar", "Beach Park, Aviva, Hot Beach, Taua, Club Med", "Destino completo, marca forte, familia, conveniencia, ferias planejadas", "Aspiracional, familiar, confiavel, com ecossistema", "Ficar bonito sem explicar modelo e custos", "Conteudo por destino, prova social, app/CRM, remarketing e lead magnet", C.inferido],
  ["Premium imobiliario", "Residence clubs, fractional premium, casas de praia em resort", "Uso recorrente, servico, menos manutencao, endereco emocional", "Consultiva, clara, elegante, sem pressa", "Prometer investimento sem base; confundir direito de uso e propriedade", "Landing tecnica, dossie privado, consultor, webinar e broker network", C.inferido],
  ["Luxo real brasileiro", "Fasano, JHSF, Txai, Nannai, villas/residencias de alto padrao", "Lugar, arquitetura, reputacao, raridade, privacidade, servico", "Economica, sensorial, autoral, pouca promocao", "Usar linguagem popular de desconto; explicar demais na primeira dobra", "PR, video cinematografico, evento privado, corretor premium e relacionamento", C.inferido],
  ["Global luxury benchmark", "Four Seasons, Ritz, Aman, Rosewood, Six Senses, One&Only", "Servico, paz de espirito, propriedade cuidada, lifestyle, sense of place", "Sofisticada, segura, centrada em servico e identidade", "Copiar palavras sem ter operacao equivalente", "Traduzir principios: property management, concierge, curadoria e prova operacional", C.comprovado],
  ["Vacation ownership internacional", "Marriott Vacation Club, Hilton Grand Vacations, Disney Vacation Club", "Flexibilidade, marca, clube, propriedade/vacation membership, familia", "Educativa e regulada", "Ignorar disclosures; vender como investimento", "Conteudo educativo, disclaimers, comparativos, oferta de apresentacao", C.comprovado],
  ["Destination club ultra-premium", "Exclusive Resorts, Inspirato", "Acesso sem possuir varios imoveis, curadoria, concierge, padrao sem surpresa", "Clube privado, insider, alto toque", "Mercado menor no Brasil; exige renda e base qualificada", "Indicação, eventos, PR, LinkedIn, concierge e conteudo de lifestyle", C.comprovado],
];

const siteAudit = [
  ["GAV Resorts", "Multipropriedade + hotelaria", "Ferias compartilhadas; casa de ferias; destino incrivel; multipropriedade no paraiso", "Adquira/garanta sua casa de ferias; hotelaria; WhatsApp; blog; portal", "FAQ define multipropriedade, depoimentos de multiproprietarios, Select Club, empreendimentos", "Promessa de economia aparece forte; precisa sustentar comparativo e evitar parecer garantia.", "Separar landing de venda imobiliaria da hotelaria; criar calculadora de uso e pagina dura de custos/regras.", C.comprovado],
  ["Aviva", "Hoteis + Residence Clubs + Clube de Ferias", "Momentos unicos; muito alem da estadia; conforto e exclusividade", "Ver pacotes; WhatsApp; clubes; newsletter", "Marcas-ancora Rio Quente, Costa do Sauipe, Hot Park; hoteis e residence clubs", "Home gera desejo, mas modelo residencial/clube precisa mais explicacao para lead frio.", "Landing especifica para Residence Clubs com uso, regras, custos, beneficios e FAQ.", C.comprovado],
  ["Beach Park", "Destino completo + resorts + vacation club", "Destino completo para toda familia no Ceara; Aqua Park; resorts; Vacation Club", "Ingressos, pacotes, WhatsApp, Instagram, portal do agente", "Parque, resorts, Vila Azul, blog, redes, Beach Card, Vacation Club", "Muitos produtos competem na home; venda de club precisa uma jornada propria.", "Criar funil por intencao: ingresso, pacote, resort, vacation club, agente.", C.comprovado],
  ["Hot Beach", "Parque + resorts + residence club", "Experiencia completa; agua quentinha; casa de ferias em Olimpia anexa ao parque", "Reservar hospedagem; comprar ingresso; saber mais; residence club", "4 resorts com acesso ao parque, premios, RCI, Vila Guarani, depoimentos implicitos", "Promocoes de ingresso/hospedagem podem puxar marca para preco; residence precisa elevar linguagem.", "Duas camadas: massa para parque/resort e premium consultivo para casa de ferias.", C.comprovado],
  ["Four Seasons Residences", "Private residences / branded residences", "Come Home; Your home, the centre of our world; tailored service; peace of mind", "Find residence; view portfolio; subscribe", "Property management, cuidados, amenidades, portfolio global, historias", "Benchmark nao deve ser copiado literalmente por marcas sem operacao de servico real.", "Traduzir para Brasil: 'a casa cuidada por uma operacao hoteleira' com provas.", C.comprovado],
  ["Marriott Vacation Club", "Vacation ownership", "Curious about ownership? talk details or stay at destinations", "Let's connect; request information; special offers", "FAQ, ownership, social, legal disclosures de timeshare", "Excelente benchmark de compliance; menos aspiracional que luxo, mais educativo.", "Copiar estrutura: educacao + oferta + disclosure + FAQ.", C.comprovado],
  ["Hilton Grand Vacations", "Vacation membership / timeshare", "Real stories; membership; direct sales; get started", "Get started; schedule virtual presentation; package holder support", "Member stories, Hilton brand, disclosures legais e preco/faixa", "A transparencia legal aumenta confianca; no Brasil, adaptar para LGPD e juridico local.", "Criar rodape de disclosure e pagina de 'termos importantes antes de comprar'.", C.comprovado],
];

const adAudit = [
  ["Meta Ads Library", "GAV Resorts", "https://www.facebook.com/ads/library/?active_status=active&ad_type=all&country=BR&q=GAV%20Resorts&search_type=keyword_unordered", "Acesso dinamico. Meta informa que Ad Library mostra anuncios ativos, mas a coleta automatizada nao abriu criativos nesta sessao.", "PRECISA CAPTURA MANUAL", "Capturar prints de criativos, copy, CTA, pagina de destino, data de inicio e formato.", C.validacao],
  ["Meta Ads Library", "Beach Park", "https://www.facebook.com/ads/library/?active_status=active&ad_type=all&country=BR&q=Beach%20Park&search_type=keyword_unordered", "Acesso dinamico; usar consulta manual.", "PRECISA CAPTURA MANUAL", "Separar anuncios de ingresso, pacote, resort e Vacation Club.", C.validacao],
  ["Meta Ads Library", "Hot Beach", "https://www.facebook.com/ads/library/?active_status=active&ad_type=all&country=BR&q=Hot%20Beach&search_type=keyword_unordered", "Acesso dinamico; usar consulta manual.", "PRECISA CAPTURA MANUAL", "Mapear campanhas de ingresso, hospedagem e Residence Club.", C.validacao],
  ["Meta Ads Library", "Aviva / Rio Quente / Costa do Sauipe / Hot Park", "https://www.facebook.com/ads/library/?active_status=active&ad_type=all&country=BR&q=Aviva%20Rio%20Quente%20Costa%20do%20Sauipe&search_type=keyword_unordered", "Acesso dinamico; marcas podem anunciar por paginas separadas.", "PRECISA CAPTURA MANUAL", "Separar marca guarda-chuva e marcas operacionais.", C.validacao],
  ["Google Ads Transparency Center", "GAV Resorts", "https://adstransparency.google.com/?region=BR&domain=gavresorts.com.br", "Centro oficial abriu, mas busca por dominio e criativos exige interface dinamica.", "PRECISA CAPTURA MANUAL", "Verificar Search/YouTube/Display, dominios e criativos ativos.", C.validacao],
  ["Google Ads Transparency Center", "Beach Park", "https://adstransparency.google.com/?region=BR&domain=beachpark.com.br", "Consulta por dominio deve ser validada manualmente.", "PRECISA CAPTURA MANUAL", "Mapear se usa Search, YouTube ou Display para ingresso/pacotes.", C.validacao],
  ["Google Ads Transparency Center", "Hot Beach", "https://adstransparency.google.com/?region=BR&domain=hotbeach.com.br", "Consulta por dominio deve ser validada manualmente.", "PRECISA CAPTURA MANUAL", "Verificar campanhas de parque, resorts e residence club.", C.validacao],
  ["TikTok Creative Center", "Categoria resort/multipropriedade", "https://ads.tiktok.com/business/creativecenter/", "Creative Center e Top Ads sao dinamicos; usar busca manual por marca, setor e pais.", "PRECISA CAPTURA MANUAL", "Coletar ganchos, duracao, texto na tela, ritmo, CTA e landing.", C.validacao],
  ["YouTube", "GAV/Beach/Hot Beach/Aviva", "https://www.youtube.com/results?search_query=GAV+Resorts+multipropriedade", "Busca organica pode misturar anuncio, institucional e conteudo de terceiros.", "PRECISA VALIDACAO", "Separar canal oficial, shorts, videos pagos identificaveis e reviews de usuarios.", C.validacao],
  ["Site como proxy de campanha", "GAV Resorts", "https://www.gavresorts.com.br/", "Promessas e CTAs ativos no site indicam linguagem comercial provavel: casa de ferias, ferias compartilhadas, economia, blog e WhatsApp.", "AUDITADO", "Usar como base ate capturar Ads Library real.", C.comprovado],
  ["Site como proxy de campanha", "Hot Beach", "https://hotbeach.com.br/", "Home mostra ofertas, compra de ingresso, reserva, casa de ferias anexa ao parque e experience complete.", "AUDITADO", "Distinguir produto de massa e venda consultiva.", C.comprovado],
  ["Site como proxy de campanha", "Beach Park", "https://beachpark.com.br/", "Home mostra destino completo, parque, resorts, pacotes, Vacation Club, Beach Card, WhatsApp e redes.", "AUDITADO", "Criar matriz por produto para nao misturar funis.", C.comprovado],
];

const ofertaPlayer = [
  ["GAV Resorts", "Hotelaria", "Reserva de hospedagem em resorts", "Hotelaria / diaria", "Ver datas / acessar hotelaria", "Baixo", C.comprovado],
  ["GAV Resorts", "Empreendimentos", "Casa de ferias via multipropriedade", "Multipropriedade imobiliaria", "Garanta/adquira sua casa de ferias", "Medio: economia e propriedade precisam explicacao", C.comprovado],
  ["GAV Resorts", "Select Club", "Clube de viagens e beneficios", "Clube / beneficios", "Saiba mais", "Baixo a medio: nao confundir com propriedade", C.comprovado],
  ["Aviva", "Hoteis", "Pacotes e hospedagem em marcas de resort", "Hotelaria", "Ver pacotes", "Baixo", C.comprovado],
  ["Aviva", "Residence Clubs", "InCasa e InCanto", "Residence club", "Conhecer/WhatsApp", "Medio: explicar direito, uso e regras", C.comprovado],
  ["Aviva", "Clube de Ferias", "Programa de ferias/memorias", "Vacation club / clube", "WhatsApp Clube de Ferias", "Medio: diferenciar de residence club", C.comprovado],
  ["Beach Park", "Aqua Park / Parque Arvorar", "Ingresso e experiencia de parque", "Entretenimento", "Comprar ingresso", "Baixo", C.comprovado],
  ["Beach Park", "Resorts", "Hospedagem no destino", "Hotelaria / pacote", "Pacotes / reservas", "Baixo", C.comprovado],
  ["Beach Park", "Vacation Club", "Produto de recorrencia/club", "Vacation club", "Conhecer / lead", "Medio: regras e beneficios precisam clareza", C.comprovado],
  ["Beach Park", "Beach Card", "Cartao/beneficio", "Membership / beneficios", "Conhecer", "Baixo a medio", C.comprovado],
  ["Hot Beach", "Parque", "Ingresso para parque aquatico", "Entretenimento", "Comprar ingresso", "Baixo", C.comprovado],
  ["Hot Beach", "Resorts", "Hospedagem com acesso ao parque", "Hotelaria", "Reservar hospedagem", "Baixo", C.comprovado],
  ["Hot Beach", "Residence Club", "Casa de ferias em Olimpia anexa ao parque", "Residence club / multipropriedade", "Conhecer casa de ferias", "Medio: explicar custos, uso, contrato e disponibilidade", C.comprovado],
  ["Marriott Vacation Club", "Ownership", "Vacation ownership por pontos/club", "Timeshare/vacation ownership", "Request information / Let's connect", "Alto regulatorio: disclosure obrigatorio", C.comprovado],
  ["Hilton Grand Vacations", "Vacation Membership", "Membership/timeshare e apresentacao", "Timeshare/vacation membership", "Get started / Schedule presentation", "Alto regulatorio: disclosure claro", C.comprovado],
];

const seoMatrix = [
  ["multipropriedade como funciona", "Educacional", "Guia completo", "Multipropriedade: como funciona, custos, uso e cuidados antes de comprar", "Lead magnet / WhatsApp", "Alta"],
  ["casa de ferias em resort", "Comercial investigativa", "Landing", "Casa de ferias em resort: entenda modelos, custos e regras", "Simular perfil de uso", "Alta"],
  ["residence club o que e", "Educacional premium", "Glossario + comparativo", "Residence club: o que e e como difere de clube de ferias", "Receber dossie", "Media"],
  ["clube de ferias vale a pena", "Comparativo", "Artigo / YouTube", "Clube de ferias vale a pena? Depende do seu perfil de viagem", "Quiz", "Alta"],
  ["multipropriedade e investimento", "Risco/compliance", "Artigo duro", "Multipropriedade e investimento? O que pode e o que nao pode ser prometido", "Falar com consultor", "Alta"],
  ["resort com parque aquatico familia", "Comercial", "Pagina de destino", "Resort com parque aquatico para ferias em familia", "Ver experiencia", "Media"],
  ["branded residences brasil", "Premium/luxo", "Artigo autoridade", "Branded residences: quando marca, servico e imovel se encontram", "Dossie privado", "Media"],
  ["segunda casa praia ou resort", "Comparativo", "Artigo / carrossel", "Segunda casa de praia ou residencia em resort: vantagens e limites", "Simular", "Media"],
  ["taxa de manutencao multipropriedade", "Objecao", "FAQ", "Taxa de manutencao: o que perguntar antes de comprar", "Baixar checklist", "Alta"],
  ["como vender multipropriedade", "B2B", "Conteudo de autoridade", "Como vender multipropriedade sem parecer venda agressiva de cota", "Treinamento", "Media"],
];

const compliance = [
  ["Voce pode ter sua casa de ferias em um resort.", "PERMITIDA COM CONTEXTO", "Usar se houver propriedade/direito real ou modelo que sustente a frase. Explicar contrato, uso e custos.", "Baixo a medio"],
  ["Multipropriedade permite dividir custos de uma casa de ferias.", "PERMITIDA COM PROVA", "Melhor: pode reduzir o desembolso frente a uma casa integral, conforme perfil de uso e custos.", "Medio"],
  ["Pague menos, aproveite mais.", "USAR COM CUIDADO", "Precisa comparativo real; nao transformar em economia garantida.", "Medio"],
  ["Retorno garantido.", "PROIBIDA", "Trocar por: entenda potencial, riscos e regras, sem garantia de retorno.", "Alto"],
  ["Renda mensal garantida com aluguel.", "PROIBIDA", "Trocar por: quando permitido, consulte regras de locacao e riscos.", "Alto"],
  ["Valorizacao garantida.", "PROIBIDA", "Trocar por: produto imobiliario pode ter variacao de valor, sem garantia.", "Alto"],
  ["Ultimas unidades.", "PERMITIDA SE REAL", "Usar apenas com estoque validado e prazo real.", "Medio"],
  ["Disponibilidade garantida na alta temporada.", "PROIBIDA SE NAO CONTRATUAL", "Trocar por: reservas seguem calendario, regras de antecedencia e disponibilidade.", "Alto"],
  ["Sem preocupacao com manutencao.", "USAR COM CUIDADO", "Trocar por: manutencao administrada conforme regras e custos do empreendimento.", "Medio"],
  ["O melhor resort do Brasil.", "RISCO DE SUPERLATIVO", "Usar so com premio/fonte. Melhor: resort com parque, estrutura X e prova Y.", "Medio"],
  ["Ideal para quem busca investimento seguro.", "PROIBIDA/RISCO ALTO", "Trocar por: ideal para quem busca uso recorrente e quer entender a dimensao patrimonial com transparencia.", "Alto"],
  ["Oferta por tempo limitado.", "PERMITIDA SE REAL", "Informar condicao, data e regra.", "Medio"],
  ["Voce nunca mais vai se preocupar com ferias.", "EXAGERADA", "Trocar por: ajuda a planejar ferias com mais previsibilidade.", "Medio"],
  ["Acesso ao parque incluso.", "PERMITIDA SE CONTRATUAL", "Explicar periodo, pessoas, limites e regras.", "Baixo a medio"],
  ["Clube exclusivo.", "PERMITIDA SE DEFINIDA", "Explicar o que e exclusivo: beneficios, acesso, prioridade ou atendimento.", "Baixo"],
];

const scripts = [
  ["WhatsApp inbound", "Primeira resposta", "Oi, [nome]. Antes de te mandar qualquer proposta, deixa eu entender uma coisa: voce esta buscando hospedagem, clube de ferias ou uma casa de ferias em modelo de propriedade/uso recorrente? Sao produtos diferentes e eu prefiro te orientar certo desde o inicio.", "Qualificar sem empurrar", "Lead novo"],
  ["WhatsApp inbound", "Diagnostico", "Como sua familia costuma viajar: todo ano para o mesmo destino, varia bastante, viaja em ferias escolares ou prefere baixa temporada?", "Descobrir fit real", "Lead morno"],
  ["WhatsApp", "Objecao 'isso e timeshare?'", "Boa pergunta. Existe muita confusao nesse mercado. O que eu posso fazer e te mostrar em uma pagina simples: o que e propriedade, o que e direito de uso, quais sao as regras de reserva e quais custos continuam existindo.", "Reduzir medo", "Lead desconfiado"],
  ["WhatsApp", "Objecao taxa", "Essa duvida e obrigatoria mesmo. Vou te separar os custos em tres blocos: aquisicao, custos recorrentes e custos que dependem de uso. Assim voce nao decide olhando so o valor de entrada.", "Transparencia", "Lead qualificado"],
  ["WhatsApp", "Objecao investimento", "Eu prefiro nao te vender isso como retorno garantido, porque seria irresponsavel. A decisao precisa fazer sentido primeiro pelo uso, pela familia e pelo modelo. A parte patrimonial deve ser analisada com documentos.", "Compliance", "Lead investidor"],
  ["WhatsApp", "Objecao disponibilidade", "A disponibilidade depende das regras do produto. Vou te enviar o calendario e o funcionamento de reservas, porque essa e uma das partes mais importantes antes de qualquer decisao.", "Clareza operacional", "Lead quente"],
  ["Ligacao", "Abertura consultiva", "Meu objetivo aqui nao e te convencer em 10 minutos. E entender se esse tipo de produto combina com a forma como sua familia viaja. Se nao fizer sentido, eu te falo tambem.", "Autoridade e confianca", "Primeira call"],
  ["Ligacao", "Pergunta-chave", "Se voce comprasse uma casa de ferias hoje, o que te daria mais trabalho: comprar, manter, alugar, reservar datas ou garantir que a familia use?", "Mapear dor central", "Call consultiva"],
  ["Showroom", "Inicio da apresentacao", "Vou te mostrar primeiro o destino e a experiencia. Depois o modelo de uso. So no final faz sentido falar de tabela.", "Sequencia correta", "Apresentacao"],
  ["Showroom", "Fechamento sem pressao", "Pelo que voce me disse, o ponto decisivo nao e so preco. E entender se o calendario, os custos e o perfil de uso batem com a rotina da sua familia.", "Fechar com criterio", "Lead quente"],
  ["Pos-visita", "Follow-up 24h", "Depois da visita, qual ponto ficou mais forte para voce: destino, estrutura, modelo de uso, custo ou duvida juridica?", "Recuperar conversa", "24 horas"],
  ["Pos-visita", "Follow-up 3 dias", "Eu revisei suas duvidas e acho que antes de qualquer decisao voce precisa olhar principalmente disponibilidade, custos recorrentes e regras de transferencia. Posso te mandar isso organizado?", "Reabrir com valor", "3 dias"],
  ["Recuperacao", "Lead parado", "Nao quero te pressionar. So quero entender se a duvida foi sobre o produto, o momento financeiro ou a confianca no modelo. Cada uma pede uma resposta diferente.", "Diagnostico de perda", "Lead frio"],
  ["Reativacao", "Novo conteudo", "Gravei um material explicando as diferencas entre hotel, clube de ferias, multipropriedade e casa integral. Acho que resolve boa parte das duvidas que ficam nesse mercado.", "Educar para reativar", "Base antiga"],
];

const campaigns = [
  ["Meta - Lead frio", "Familias que pesquisam ferias, resort e destinos", "Video curto + carrossel", "Casa de ferias em resort: entenda antes de comprar", "Receber guia", "CPL qualificado + taxa WhatsApp", "Evitar renda, investimento, economia garantida"],
  ["Meta - Remarketing", "Visitantes site, engajados video, Instagram", "FAQ em carrossel + depoimento + tour", "Ainda em duvida? Veja as perguntas que travam a decisao", "Simular perfil de uso", "Agendamento", "Landing precisa responder as mesmas perguntas do anuncio"],
  ["Google Search", "Busca ativa por modelo/destino", "Search ads", "Multipropriedade como funciona / casa de ferias em resort", "Falar com consultor", "Lead qualificado", "Enviar para pagina educativa, nao so formulario"],
  ["YouTube", "Consideracao", "Video 6-12 min + Shorts", "Guia honesto do modelo + tour do destino", "Baixar checklist", "Watch time + leads", "Sem promessa de retorno"],
  ["TikTok/Reels organico", "Descoberta", "Objecoes e mitos em 30-45s", "Nao compre uma promessa de ferias. Entenda o modelo.", "Comentar 'guia'", "Retencao, compartilhamento, comentarios", "Evitar clickbait e comparativos sem prova"],
  ["Pos-estadia", "Hospedes recentes", "WhatsApp/email/app", "Transforme sua experiencia em ferias recorrentes", "Conversa consultiva", "Taxa de resposta e proposta", "Usar consentimento LGPD"],
  ["Premium privado", "Alta renda e indicacoes", "Evento, dossie privado, broker", "Residencia de ferias com servico e gestao", "Solicitar dossie", "Reuniao qualificada", "Nao usar linguagem popular/promocional"],
];

const scriptsLanding = [
  ["H1", "Sua casa de ferias em um resort, com regras claras para usar, planejar e voltar todos os anos."],
  ["Subheadline", "Entenda o modelo, os custos, as regras de reserva e quando esse tipo de produto realmente faz sentido para sua familia."],
  ["CTA frio", "Receber guia dos modelos de ferias em resort"],
  ["CTA morno", "Simular meu perfil de uso"],
  ["CTA quente", "Falar com consultor"],
  ["Bloco 1", "Nao comece pela cota. Comece pela forma como sua familia viaja."],
  ["Bloco 2", "Compare hotel, casa de praia, clube de ferias, multipropriedade e residence club."],
  ["Bloco 3", "Veja o que esta incluso, o que e custo recorrente e o que depende das regras do contrato."],
  ["Bloco 4", "Conheca a estrutura: resort, lazer, gastronomia, parque, servico, destino e suporte."],
  ["Bloco 5", "Perguntas obrigatorias antes de comprar: disponibilidade, taxa, revenda, heranca, locacao, reajuste e uso."],
  ["Rodape/disclosure", "As condicoes de uso, disponibilidade, custos, beneficios e eventuais direitos de propriedade dependem do contrato e do empreendimento. Nao ha promessa de renda, valorizacao ou retorno garantido."],
];

const plan30 = Array.from({ length: 30 }, (_, i) => {
  const d = i + 1;
  const phase = d <= 4 ? "Auditoria" : d <= 8 ? "Oferta e site" : d <= 14 ? "Conteudo" : d <= 21 ? "Midia e CRM" : "Otimizacao";
  const actions = [
    "Capturar prints de Ads Library Meta por player e produto.",
    "Capturar Google Ads Transparency por dominio e marca.",
    "Capturar TikTok Creative Center/Top Ads e YouTube oficial por player.",
    "Auditar heroes, CTAs, formularios, WhatsApp, FAQ, SEO e mobile dos sites.",
    "Separar players em popular, premium familiar, premium imobiliario e luxo real.",
    "Criar matriz de oferta: diaria, ingresso, clube, multipropriedade, residence club, branded residence.",
    "Validar juridicamente claims: casa de ferias, investimento, economia, acesso, disponibilidade e exclusividade.",
    "Definir promessa central e inimigo narrativo por segmento.",
    "Escrever landing principal com FAQ, comparativo e CTA por funil.",
    "Criar lead magnet: guia dos modelos de ferias em resort.",
    "Criar quiz de perfil de uso: frequencia, datas, familia, destino, orcamento e objecoes.",
    "Produzir 10 Reels/TikToks de objecoes e mitos.",
    "Produzir 3 videos YouTube de autoridade: guia, comparativo e tour.",
    "Criar carrosseis: perguntas obrigatorias, custos, diferencas entre modelos.",
    "Configurar campanhas Meta: lead frio, remarketing e pos-estadia.",
    "Configurar Google Search por termos de intencao e destino.",
    "Criar campanhas YouTube Shorts/In-stream para educacao e remarketing.",
    "Configurar UTMs, eventos, conversoes, WhatsApp tracking e CRM.",
    "Treinar SDR/vendedor com scripts de diagnostico e objeções.",
    "Criar follow-up 1/3/7/14 dias por objeção.",
    "Criar material para decisor secundario: PDF de 2 paginas.",
    "Analisar CPL, lead qualificado, taxa de resposta e agendamento.",
    "Ler conversas do WhatsApp e classificar objeções reais.",
    "Otimizar landing com as 10 duvidas mais frequentes.",
    "Pausar criativos de curiosidade vazia; escalar criativos que geram pergunta qualificada.",
    "Criar nova rodada de criativos com prova operacional: obra, resort, destino, documentos.",
    "Rodar teste A/B: guia vs simulacao de uso.",
    "Rodar campanha premium privada para lead de indicacao/broker.",
    "Revisar compliance com juridico antes de escalar.",
    "Fechar playbook: criativos vencedores, scripts, objeções, paginas e proximos testes.",
  ];
  return [d, phase, actions[i], "Responsavel definido", "RASCUNHO PARA APROVACAO"];
});

const workbook = Workbook.create();

function writeSheet(name, data, widths = []) {
  const sheet = workbook.worksheets.add(name);
  sheet.showGridLines = false;
  const rows = data.length;
  const cols = data[0].length;
  const range = sheet.getRangeByIndexes(0, 0, rows, cols);
  range.values = data;
  sheet.freezePanes.freezeRows(1);
  range.format = { wrapText: true, borders: { preset: "all", style: "thin", color: "#D9D9D9" } };
  const header = sheet.getRangeByIndexes(0, 0, 1, cols);
  header.format = { fill: "#102A43", font: { bold: true, color: "#FFFFFF" }, wrapText: true };
  try {
    const lastCol = String.fromCharCode("A".charCodeAt(0) + cols - 1);
    const table = sheet.tables.add(`A1:${lastCol}${rows}`, true, name.replace(/[^A-Za-z0-9]/g, "_").slice(0, 30));
    table.style = "TableStyleMedium2";
    table.showFilterButton = true;
  } catch {}
  for (let c = 0; c < cols; c++) {
    sheet.getRangeByIndexes(0, c, rows, 1).format.columnWidthPx = widths[c] ?? (c === 0 ? 190 : 260);
  }
  sheet.getRangeByIndexes(1, 0, Math.max(rows - 1, 1), cols).format.rowHeightPx = 96;
}

writeSheet("00 Leia Primeiro", [
  ["Campo", "Conteudo"],
  ["Objetivo", "Dossie 3.0 para transformar a base de branding em operacao real de marketing, site, anuncios, scripts, CRM e treinamento de venda."],
  ["Tese central", "Nao misturar GAV/WAM/diRoma com Fasano/JHSF/Four Seasons. O mercado deve ser segmentado por popular, premium familiar, premium imobiliario e luxo real."],
  ["Limite da auditoria de anuncios", "Meta Ads Library, Google Ads Transparency Center e TikTok Creative Center sao fontes dinamicas. Os links de consulta foram montados, mas criativos reais precisam captura manual/print para validade plena."],
  ["Regra de uso", "Tudo ainda esta como RASCUNHO PARA APROVACAO. Claims financeiros e juridicos exigem validacao antes de trafego pago."],
], [180, 900]);
writeSheet("01 Segmentacao", [["Segmento", "Players exemplo", "O que vende", "Tom", "Risco", "Estrategia", "Certeza"], ...segmentacao], [180, 320, 300, 280, 300, 380, 160]);
writeSheet("02 Auditoria Sites", [["Player", "Modelo", "Hero/promessa", "CTA", "Provas/ativos", "Risco", "Recomendacao", "Certeza"], ...siteAudit], [180, 220, 360, 240, 360, 340, 420, 150]);
writeSheet("03 Auditoria Ads", [["Plataforma", "Player/consulta", "Link de auditoria", "O que foi possivel verificar", "Status", "Proxima acao", "Certeza"], ...adAudit], [190, 260, 520, 420, 220, 420, 150]);
writeSheet("04 Oferta Player", [["Player", "Produto", "O que e", "Tipo real", "CTA", "Risco", "Certeza"], ...ofertaPlayer], [170, 220, 360, 230, 240, 360, 150]);
writeSheet("05 SEO Conteudo", [["Keyword", "Intencao", "Formato", "Titulo", "CTA", "Prioridade"], ...seoMatrix], [230, 210, 230, 430, 220, 160]);
writeSheet("06 Copy Compliance", [["Claim", "Status", "Como usar/substituir", "Risco"], ...compliance], [330, 210, 540, 180]);
writeSheet("07 Scripts Venda", [["Canal", "Momento", "Script", "Funcao", "Quando usar"], ...scripts], [160, 200, 680, 240, 220]);
writeSheet("08 Campanhas", [["Campanha", "Publico", "Criativo", "Promessa", "CTA", "KPI", "Compliance"], ...campaigns], [200, 280, 260, 340, 220, 220, 360]);
writeSheet("09 Landing Copy", [["Bloco", "Copy pronta"], ...scriptsLanding], [190, 820]);
writeSheet("10 Plano 30 Dias", [["Dia", "Fase", "Acao", "Responsavel", "Status"], ...plan30], [70, 170, 700, 220, 220]);
writeSheet("11 Fontes", [["Fonte", "URL", "Uso", "Certeza"], ...sources], [260, 520, 520, 160]);

const xlsxOut = await SpreadsheetFile.exportXlsx(workbook);
const xlsxPath = path.join(outputDir, "Dossie_3_0_Resorts_Marketing_Operacional.xlsx");
await xlsxOut.save(xlsxPath);

const errors = await workbook.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
  options: { useRegex: true, maxResults: 100 },
  summary: "formula error scan",
});

for (const sheetName of ["01 Segmentacao", "02 Auditoria Sites", "03 Auditoria Ads", "07 Scripts Venda", "08 Campanhas"]) {
  const img = await workbook.render({ sheetName, autoCrop: "all", scale: 1, format: "png" });
  await fs.writeFile(path.join(outputDir, `${sheetName.replaceAll(" ", "_")}.png`), new Uint8Array(await img.arrayBuffer()));
}

const md = `# Dossie 3.0 - Marketing Operacional para Resorts, Multipropriedade e Imobiliario Turistico Premium

Status: RASCUNHO PARA APROVACAO  
Data: 2026-06-18  
Base usada: planilha de branding + dossie de inteligencia anterior + auditoria de fontes abertas atuais.

## Diagnostico duro

A base anterior ficou boa, mas ainda estava em nivel de estrategia. O salto agora e operacional: separar segmentos, auditar ofertas reais, criar scripts, copy segura, paginas, campanhas e plano de execucao.

O erro mais perigoso e misturar tudo:

- GAV, WAM, diRoma e parte do mercado de multipropriedade trabalham com volume, familia, condicao, acesso e explicacao simples.
- Beach Park, Aviva, Hot Beach, Taua e Club Med trabalham melhor como ecossistema de destino: parque, resort, experiencia, familia, recorrencia e marca.
- Fasano, JHSF, Txai, Nannai e benchmarks globais nao podem usar a mesma linguagem de "cota", desconto ou urgencia popular. Luxo real vende lugar, privacidade, servico, arquitetura, reputacao e acesso.
- Marriott Vacation Club e Hilton Grand Vacations sao benchmarks de compliance: eles educam e divulgam disclosures claros quando falam de timeshare/vacation ownership.

## O que foi comprovado

- GAV usa "ferias compartilhadas", "casa de ferias", "multipropriedade" e CTAs como adquirir/garantir casa de ferias.
- Aviva comunica hoteis, Residence Clubs, Clube de Ferias, Rio Quente, Costa do Sauipe e Hot Park.
- Beach Park comunica destino completo com parques, resorts, Vacation Club, Beach Card, blog, WhatsApp, redes e portal do agente.
- Hot Beach comunica parque aquatico, resorts, compra de ingresso, reserva de hospedagem e "casa de ferias em Olimpia anexa ao parque".
- Four Seasons vende residences com foco em casa, servico, property management e peace of mind.
- Marriott e Hilton usam disclosures explicitos para vacation ownership/timeshare.
- Meta, TikTok e Google exigem coerencia entre anuncio e landing, evitam pratica enganosa e restringem claims sensiveis.

## O que precisa validacao manual

As bibliotecas de anuncios sao dinamicas. Foram montados links de auditoria para Meta Ads Library, Google Ads Transparency Center, TikTok Creative Center e YouTube, mas os criativos ativos precisam ser capturados manualmente em print ou export para a auditoria ficar juridicamente e operacionalmente forte.

Isso nao e detalhe. Sem print do anuncio real, qualquer conclusao sobre criativo ativo deve ficar como PRECISA VALIDACAO.

## Segmentacao estrategica

${segmentacao.map((r) => `### ${r[0]}\nPlayers: ${r[1]}\n\nO que vende: ${r[2]}\n\nTom: ${r[3]}\n\nRisco: ${r[4]}\n\nEstrategia: ${r[5]}\n`).join("\n")}

## Landing page recomendada

H1:
> Sua casa de ferias em um resort, com regras claras para usar, planejar e voltar todos os anos.

Subheadline:
> Entenda o modelo, os custos, as regras de reserva e quando esse tipo de produto realmente faz sentido para sua familia.

CTAs:
- Frio: Receber guia dos modelos de ferias em resort.
- Morno: Simular meu perfil de uso.
- Quente: Falar com consultor.

Blocos obrigatorios:
- Nao comece pela cota. Comece pela forma como sua familia viaja.
- Compare hotel, casa de praia, clube de ferias, multipropriedade e residence club.
- Veja o que esta incluso, o que e custo recorrente e o que depende das regras do contrato.
- Conheca estrutura, destino, servico e suporte.
- Perguntas obrigatorias antes de comprar: disponibilidade, taxa, revenda, heranca, locacao, reajuste e uso.

Disclosure recomendado:
> As condicoes de uso, disponibilidade, custos, beneficios e eventuais direitos de propriedade dependem do contrato e do empreendimento. Nao ha promessa de renda, valorizacao ou retorno garantido.

## Campanhas prioritarias

${campaigns.map((r) => `- ${r[0]}: ${r[3]} | CTA: ${r[4]} | KPI: ${r[5]} | cuidado: ${r[6]}`).join("\n")}

## Script principal de WhatsApp

> Oi, [nome]. Antes de te mandar qualquer proposta, deixa eu entender uma coisa: voce esta buscando hospedagem, clube de ferias ou uma casa de ferias em modelo de propriedade/uso recorrente? Sao produtos diferentes e eu prefiro te orientar certo desde o inicio.

## Regra de ouro para vendedores

O vendedor ruim tenta fechar a cota. O vendedor bom diagnostica a forma como a familia viaja.  
Se a familia nao tem fit de uso, calendario, renda, interesse pelo destino ou confianca no modelo, a venda vira arrependimento.

## Fontes

${sources.map((s) => `- [${s[0]}](${s[1]}) - ${s[2]} (${s[3]})`).join("\n")}
`;

const mdPath = path.join(outputDir, "Dossie_3_0_Resorts_Marketing_Operacional.md");
await fs.writeFile(mdPath, md, "utf8");

console.log(errors.ndjson);
console.log(JSON.stringify({ outputDir, xlsxPath, mdPath }, null, 2));
