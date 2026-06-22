import fs from "fs";
import path from "path";
import { fileURLToPath } from "url";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const outDir = path.join(__dirname, "outputs", "dossie_imobiliario_turistico_premium");
fs.mkdirSync(outDir, { recursive: true });

const certainty = {
  comprovado: "COMPROVADO",
  inferido: "INFERIDO",
  hipotese: "HIPOTESE",
  validacao: "PRECISA VALIDACAO",
};

const sources = [
  ["GAV Resorts", "https://www.gavresorts.com.br/", "Site oficial posiciona a empresa em turismo, hotelaria e multipropriedade; usa casa de ferias, ferias compartilhadas, Select Club e venda de projetos.", certainty.comprovado],
  ["Gramado Parks", "https://gramadoparks.com/pt-br/", "Site oficial integra hospedagem, entretenimento e gastronomia, com parques e empreendimentos em Gramado, Foz, Rio e destinos de praia.", certainty.comprovado],
  ["WAM Hoteis", "https://wamhoteis.com.br/", "Site oficial comunica hotelaria, parques, pass, blog, redes sociais e destinos como Caldas Novas, Buzios, Aquiraz e Porto Seguro.", certainty.comprovado],
  ["Aviva", "https://www.aviva.com.br/", "Site oficial integra Rio Quente, Costa do Sauipe, Hot Park, Residence Clubs e Clube de Ferias.", certainty.comprovado],
  ["Beach Park", "https://beachpark.com.br/", "Site oficial comunica destino completo com parque aquatico, resorts, restaurantes, Vacation Club, Beach Card, blog e canais sociais.", certainty.comprovado],
  ["Hot Beach Olimpia", "https://hotbeach.com.br/", "Site oficial combina parque, resorts, Vila Guarani e Residence Club, com promessa de casa de ferias anexa ao parque.", certainty.comprovado],
  ["diRoma", "https://diroma.com.br/", "Site oficial comunica hoteis, Acqua Park, VIP Club, aguas termais, beneficios familiares e condicoes comerciais.", certainty.comprovado],
  ["Taua Resorts", "https://www.tauaresorts.com.br/", "Site oficial comunica resorts para familia, lazer, estrutura, tour virtual, WhatsApp e cultura do sorriso.", certainty.comprovado],
  ["Vila Gale Brasil", "https://www.vilagale.com/pt-br/hoteis-brasil/", "Site oficial lista hoteis e resorts no Brasil e comunica lazer, hospitalidade e destinos.", certainty.comprovado],
  ["Grupo Wish", "https://www.grupowish.com/", "Site oficial do grupo hoteleiro com marcas e resorts no Brasil.", certainty.comprovado],
  ["Bourbon Hoteis e Resorts", "https://www.bourbon.com.br/", "Site oficial do grupo hoteleiro com resorts, eventos e lazer familiar.", certainty.comprovado],
  ["Fasano", "https://fasano.com.br/", "Site oficial comunica hoteis, gastronomia, experiencias, villas e destinos de alto padrao.", certainty.comprovado],
  ["JHSF", "https://jhsf.com.br/", "Site oficial comunica shopping centers, hospitalidade, gastronomia, clubes privados, aviation, real estate e negocios para clientes especiais.", certainty.comprovado],
  ["Club Med Brasil", "https://www.clubmed.com.br/", "Site oficial comunica premium all-inclusive, Exclusive Collection, consultores e resorts no Brasil.", certainty.comprovado],
  ["Nannai", "https://www.nannai.com.br/", "Site oficial comunica experiencia sofisticada, destinos como Muro Alto, Noronha e Milagres, gastronomia e bem-estar.", certainty.comprovado],
  ["Txai Resorts", "https://www.txairesorts.com/", "Site oficial comunica luxo da simplicidade, Relais & Chateaux, bangalos, residencias e vilas.", certainty.comprovado],
  ["Pratagy", "https://pratagy.com.br/", "Site oficial comunica beach all inclusive, acqua park e club de ferias.", certainty.comprovado],
  ["Malai Manso", "https://www.malaimansoresort.com.br/", "Site oficial comunica resort all inclusive, natureza, cotistas, spa e familia.", certainty.comprovado],
  ["Four Seasons Private Residences", "https://www.fourseasons.com/residences/", "Site oficial comunica residencias privadas, residence clubs, villa rentals, property management e paz de espirito.", certainty.comprovado],
  ["Ritz-Carlton Residences", "https://www.ritzcarlton.com/en/residences/", "Site oficial comunica whole ownership, servicos residenciais, concierge, in-residence dining, valet, spa e beneficios Marriott Bonvoy.", certainty.comprovado],
  ["Marriott Vacation Club", "https://www.marriottvacationclubs.com/ownership/vacation-club-brands/marriott-vacation-club.html", "Site oficial comunica vacation ownership, club points, villas/city properties e divulga disclosures de timeshare.", certainty.comprovado],
  ["Hilton Grand Vacations", "https://www.hiltongrandvacations.com/en/discover-hilton-grand-vacations", "Site oficial comunica vacation membership, home resort, weeks, flexibilidade, membros, resorts e disclosures de timeshare.", certainty.comprovado],
  ["Six Senses Residences", "https://www.sixsenses.com/en/residences/", "Site oficial comunica residences, acesso a amenidades de resort e linguagem de natureza, bem-estar e sustentabilidade.", certainty.comprovado],
  ["One&Only Private Homes", "https://www.oneandonlyresorts.com/private-homes", "Site oficial comunica private homes de ultra-luxo para compra e estadia.", certainty.comprovado],
  ["Rosewood Residences", "https://www.rosewoodhotels.com/en/residences", "Site oficial comunica ownership, serviced apartments, rentals, sense of place, servicos e privilegios globais.", certainty.comprovado],
  ["Aman Residences", "https://www.aman.com/residences", "Site oficial comunica imersao permanente no estilo de vida Aman, privacidade e casas para poucos.", certainty.comprovado],
  ["Exclusive Resorts", "https://exclusiveresorts.com/", "Site oficial comunica luxury vacation club com residencias, destinos, membership fee e annual dues.", certainty.comprovado],
  ["Inspirato", "https://www.inspirato.com/", "Site oficial comunica luxury vacation club/travel subscription, residencias, destinos, hoteis e dedicated care.", certainty.comprovado],
  ["The Registry Collection", "https://www.theregistrycollection.com/", "Site oficial comunica exchange vacations e portfolio luxury para membros.", certainty.comprovado],
  ["Leading Hotels of the World", "https://www.lhw.com/", "Site oficial comunica colecao de hoteis independentes de luxo, Leaders Club, ofertas e app.", certainty.comprovado],
  ["Small Luxury Hotels", "https://slh.com/", "Site oficial comunica mais de 700 hoteis independentes em mais de 100 paises e clube de fidelidade.", certainty.comprovado],
  ["Meta Ad Standards", "https://transparency.meta.com/policies/ad-standards/", "Fonte oficial de politicas de anuncios da Meta, incluindo revisao de anuncios, proibicoes de conteudo enganoso e regras de lead ads.", certainty.comprovado],
  ["Google Ads Personalized Advertising", "https://support.google.com/adspolicy/answer/143465", "Fonte oficial de politicas de publicidade personalizada, privacidade, elegibilidade e restricoes para categorias sensiveis em alguns mercados.", certainty.comprovado],
  ["TikTok Advertising Policies", "https://ads.tiktok.com/help/article/tiktok-advertising-policies", "Fonte oficial de politicas de anuncios do TikTok.", certainty.comprovado],
  ["TikTok Misleading and False Content", "https://ads.tiktok.com/help/article/tiktok-ads-policy-misleading-and-false-content", "Fonte oficial sobre conteudo enganoso, promessas exageradas e omissao de informacoes materiais.", certainty.comprovado],
  ["YouTube Recommendations", "https://www.youtube.com/howyoutubeworks/recommendations/", "Fonte oficial sobre recomendacoes, desempenho e satisfacao do usuario.", certainty.comprovado],
  ["Google Search Helpful Content", "https://developers.google.com/search/docs/fundamentals/creating-helpful-content", "Fonte oficial sobre conteudo util, confiavel e voltado para pessoas.", certainty.comprovado],
  ["Lei 13.777/2018", "http://www.planalto.gov.br/ccivil_03/_ato2015-2018/2018/lei/L13777.htm", "Fonte oficial da lei brasileira que regula multipropriedade imobiliaria.", certainty.comprovado],
  ["ANPD", "https://www.gov.br/anpd/pt-br", "Fonte oficial sobre autoridade nacional de protecao de dados e LGPD.", certainty.comprovado],
];

const businessModels = [
  ["Propriedade integral de segunda residencia", "Compra de unidade ou casa em destino turistico para uso proprio, renda eventual e reserva patrimonial.", "Alto patrimonio, familias, empresarios, estrangeiros, investidores patrimoniais.", "Controle, exclusividade, endereco desejado, liberdade de uso.", "Alto ticket, manutencao, vacancia, liquidez.", "Site institucional, broker de luxo, eventos fechados, PR, indicacao, YouTube de destino.", certainty.comprovado],
  ["Multipropriedade imobiliaria", "Fracionamento juridico de unidade por periodos de tempo, com uso exclusivo em semanas ou frações definidas.", "Familias que querem ferias recorrentes com menor desembolso que um imovel inteiro.", "Acesso a casa de ferias, previsibilidade de ferias, infraestrutura de resort.", "Entender regras, custos anuais, disponibilidade, revenda, promessa de valorizacao.", "Google Search, Meta, YouTube educativo, stand, inside sales, WhatsApp consultivo.", certainty.comprovado],
  ["Fractional ownership premium", "Modelo fracionado com linguagem mais patrimonial e sofisticada, geralmente associado a residencia, clube ou destino de alto valor.", "Publico premium que rejeita linguagem de timeshare comum.", "Mais exclusividade, menos compromisso que propriedade integral.", "Confundir com investimento financeiro ou clube comum.", "Conteudo comparativo, consultoria, evento, pagina tecnica e prova de gestao.", certainty.inferido],
  ["Residence club", "Clube residencial com uso recorrente, servicos hoteleiros e acesso a unidades/residencias.", "Familias premium com rotina anual de ferias.", "Pertencimento, padrao de servico, facilidade de uso.", "Contrato, disponibilidade, custos, direito de uso versus propriedade.", "Landing educativa, clube de beneficios, CRM e storytelling de lifestyle.", certainty.comprovado],
  ["Vacation club", "Clube de ferias baseado em pontos, diarias, descontos ou acesso a rede de hospedagem.", "Viajantes recorrentes e familias que querem flexibilidade.", "Flexibilidade, economia percebida, beneficios em rede.", "Regras, disponibilidade, taxas, comparacao com booking direto.", "WhatsApp, remarketing, email, app, pos-estadia e upgrade.", certainty.comprovado],
  ["Destination club", "Clube de viagem premium com residencias em varios destinos, normalmente com taxa de entrada e anuidade.", "Ultra high net worth e executivos que valorizam acesso sem possuir varios imoveis.", "Viajar como insider, residencias prontas, curadoria e servico.", "Custo fixo, disponibilidade, aderencia ao estilo de vida.", "PR, eventos privados, indicacao, concierge, videos de experiencia.", certainty.comprovado],
  ["Branded residences", "Residencias com marca hoteleira ou de luxo, combinando propriedade privada, servico, design e reputacao.", "Compradores globais, investidores patrimoniais, clientes de marca.", "Marca reduz risco percebido, eleva status e padroniza servico.", "Preco premium, promessa de rentabilidade, regras de operacao.", "Relações publicas, broker network, site imersivo, CRM de alto toque.", certainty.comprovado],
  ["Condo-hotel", "Unidades imobiliarias dentro de operacao hoteleira, podendo envolver pool de locacao e regras especificas.", "Investidor patrimonial e comprador que busca uso + renda potencial.", "Gestao profissional, destino turistico, renda potencial.", "Regulacao, promessa de rentabilidade, CVM se caracterizar investimento coletivo.", "Conteudo juridico claro, disclaimers, consultor treinado, pagina de documentos.", certainty.inferido],
  ["Resort membership", "Programa de associacao com beneficios, acesso, descontos, experiencias e status.", "Hospedes frequentes e familias locais/regionais.", "Pertencimento e beneficio recorrente sem compra imobiliaria.", "Valor percebido versus custo recorrente.", "App, email, pos-estadia, social, campanhas sazonais.", certainty.inferido],
  ["All-inclusive resort", "Produto hoteleiro com hospedagem, alimentacao, lazer e experiencia integrada.", "Familias, casais, grupos e empresas.", "Custo previsivel, conforto, entretenimento e conveniencia.", "Preco, localizacao, qualidade real, lotacao.", "SEO, metasearch, Instagram, TikTok, YouTube, agencias.", certainty.comprovado],
  ["Resort imobiliario com parque", "Empreendimento ancora parque aquatico/tematico, hoteis e venda de uso/propriedade/clube.", "Familias com filhos e compradores de ferias planejadas.", "Atracao recorrente aumenta desejo e justificativa de uso.", "Sazonalidade, custos, saturacao do destino, concorrencia.", "Conteudo de experiencia, prova social, YouTube, Google, app.", certainty.comprovado],
  ["Private villas em resort", "Villas/casas dentro de resort ou ecossistema hoteleiro, com servicos e privacidade.", "Publico alto luxo, familias grandes, estrangeiros.", "Privacidade com serviço, espaco, exclusividade.", "Preco, seguranca, gestao e liquidez.", "PR, arquitetura/design, video cinematografico, broker e relacionamento.", certainty.comprovado],
  ["Residencias de praia independentes", "Casas ou apartamentos premium em destino turistico sem marca hoteleira forte.", "Comprador patrimonial, investidor local, familia de alta renda.", "Casa propria, design, localidade, personalizacao.", "Gestao, manutencao, aluguel, seguranca.", "SEO local, Google Ads, Instagram, portais premium e corretores.", certainty.inferido],
  ["Clube de troca/intercambio", "Sistema para trocar semanas/pontos por hospedagens em portfolio nacional ou internacional.", "Proprietarios/cotistas que desejam variar destino.", "Flexibilidade e amplitude de destinos.", "Taxas, disponibilidade e complexidade de regras.", "App, CRM, conteudo educativo, onboarding.", certainty.comprovado],
  ["Resort para eventos e corporativo", "Uso de estrutura de resort para eventos, convencoes, casamentos e incentivos.", "Empresas, agencias MICE, casais, grupos.", "Estrutura completa, hospedagem, lazer e logistica.", "Custo, capacidade, deslocamento, datas.", "LinkedIn, Google, SEO, parceria com agencias e pagina de eventos.", certainty.comprovado],
  ["Hospitality + real estate ecosystem", "Grupo que combina hoteis, restaurantes, clubes, malls, aeroportos/servicos e residencial premium.", "Alta renda urbana e investidores de ecossistema.", "Ecossistema aumenta recorrencia, status e conveniencia.", "Complexidade, brand trust, entrega operacional.", "Brand content, PR, relacao com investidores, eventos e CRM VIP.", certainty.comprovado],
];

const brazilPlayers = [
  ["GAV Resorts", "Multipropriedade + hotelaria", "Familia, ferias compartilhadas, casa de ferias", "Projeto + destino + clube", "Casa de ferias, ferias compartilhadas, Select Club", "Educar juridicamente sem parecer juridico; vender previsibilidade de ferias antes de cota.", certainty.comprovado],
  ["Gramado Parks", "Hospitalidade + parques + multiproduto", "Familia e lazer em destino turistico", "Ecossistema de hospedagem, entretenimento e gastronomia", "Novo padrao de lazer e ferias", "Ecossistema aumenta conteudo e motivo de retorno; bom benchmark para funil de destino.", certainty.comprovado],
  ["WAM", "Hotelaria, parques, pass e multipropriedade associada", "Familias em destinos populares", "Rede de destinos e hospitalidade", "Especialistas em experiencia; WAM Pass", "Usar rede e variedade para reduzir medo de destino unico.", certainty.comprovado],
  ["Aviva", "Resorts, parques, residence clubs, clube de ferias", "Familias, lazer premium, ferias recorrentes", "Rio Quente, Costa do Sauipe e Hot Park", "Muito alem da estadia; conforto e exclusividade", "Ensina a vender ecossistema: resort + parque + clube + marca.", certainty.comprovado],
  ["Beach Park", "Destino integrado + vacation club", "Familias, Nordeste, parque aquatico", "Destino completo no Ceara", "Aqua Park, resorts, Vacation Club, Beach Card", "A ancora de experiencia facilita conteudo organico e remarketing.", certainty.comprovado],
  ["Hot Beach Olimpia", "Parque + resorts + residence club", "Familias que viajam para Olimpia", "Casa de ferias anexa ao parque", "Experiencia completa; acesso ao parque", "Mensagem forte porque junta residencia com acesso imediato a entretenimento.", certainty.comprovado],
  ["diRoma", "Rede hoteleira + parque + VIP Club", "Familias e publico de aguas termais", "Caldas Novas, aguas termais e recreacao", "Aguas termais, kids free, 10x, VIP Club", "Benchmark de oferta clara e promocional; cuidado para nao virar commodity.", certainty.comprovado],
  ["Taua Resorts", "Resorts familiares", "Familias com criancas e eventos", "Estrutura, recreacao, sorriso e conveniencia", "Cultura do sorriso; resort familia", "Benchmark de conteudo de familia e experiencia antes da venda imobiliaria.", certainty.comprovado],
  ["Vila Gale", "Rede de hoteis e resorts", "Familias, lazer, eventos", "Marca internacional lusobrasileira e resorts", "Resorts e hoteis no Brasil", "Bom estudo de SEO destino + oferta hoteleira.", certainty.comprovado],
  ["Grupo Wish", "Grupo hoteleiro premium", "Turismo, eventos e lazer", "Portfolio de marcas e resorts", "Wish, Prodigy, Linx e operacao hoteleira", "Bom benchmark de marca guarda-chuva e segmentacao por unidade.", certainty.comprovado],
  ["Bourbon", "Hoteis, resorts e eventos", "Familia e corporativo", "Tradicao hoteleira, lazer e MICE", "Bourbon Atibaia, Cataratas, eventos", "Mostra que resort premium nao vive so de familia; eventos amortecem sazonalidade.", certainty.comprovado],
  ["Fasano", "Hotelaria, gastronomia e villas de luxo", "Alta renda, lifestyle, gastronomia, design", "Marca aspiracional e experiencia sofisticada", "Dal 1902, gastronomia, hoteis, villas", "Referencia para linguagem de luxo: menos promocional, mais curadoria, lugar e reputacao.", certainty.comprovado],
  ["JHSF", "Ecossistema de luxo e real estate", "High net worth, clientes especiais, investidores", "Shopping, hospitalidade, clubes privados, aviation, real estate", "Unique businesses for special clients", "Ensina que alto luxo vende ecossistema, pertencimento e acesso, nao so metro quadrado.", certainty.comprovado],
  ["Club Med Brasil", "Premium all-inclusive", "Familias, casais e viagens internacionais", "All-inclusive com esportes, gastronomia e kids", "Premium all inclusive, Exclusive Collection", "Bom benchmark para vender previsibilidade e experiencia sem friccao.", certainty.comprovado],
  ["Nannai", "Resort/hotelaria de luxo", "Casais, familias premium, lua de mel", "Luxo sensorial, simplicidade sofisticada, destinos icônicos", "Convite ao sentir; estado de espirito", "Ensina a comunicar desejo sem gritar oferta.", certainty.comprovado],
  ["Txai", "Resort luxo + residencias/vilas", "Alto luxo, natureza, exclusividade", "Relais & Chateaux, luxo da simplicidade", "Arte baiana de bem-receber; residencias e vilas", "Referencia para branded place: identidade local acima de features.", certainty.comprovado],
  ["Pratagy", "All-inclusive + acqua park + club ferias", "Familias e Nordeste", "Resort de praia com parque e clube", "All inclusive 24h, club ferias", "Combina hospedagem imediata com recorrencia de relacionamento.", certainty.comprovado],
  ["Malai Manso", "All-inclusive + cotistas", "Familias, natureza, centro-oeste", "Resort natureza, lake, spa, cotistas", "All nature inclusive; cotistas", "Bom benchmark para vender destino fora do circuito obvio.", certainty.comprovado],
  ["Costao do Santinho", "Resort all-inclusive e eventos", "Familias, eventos, lazer sul", "Praia, estrutura e experiencia completa", "Resort completo em Florianopolis", "PRECISA validar detalhes atuais; fonte oficial deve ser auditada manualmente.", certainty.validacao],
  ["Jurema Aguas Quentes", "Resorts de aguas termais", "Familias e turismo regional", "Aguas termais, natureza, lazer", "Complexo de resorts no Parana", "Bom benchmark regional; precisa captura detalhada do site oficial.", certainty.validacao],
  ["Hard Rock Hotel Brasil / VCI", "Branded hotel + fractional/vacation ownership", "Publico aspiracional e fas de marca", "Marca global de entretenimento aplicada a resort", "Hard Rock como ancora de desejo", "Usar somente apos validacao documental, pois claims comerciais mudam rapido.", certainty.validacao],
  ["Laghetto", "Rede hoteleira e resorts", "Familias, Sul, lazer", "Marca regional forte em Gramado e expansao", "Hoteis e resorts Laghetto", "Bom benchmark para SEO local e recorrencia em Gramado.", certainty.validacao],
  ["Salinas", "All-inclusive resort", "Familias e praia Nordeste", "Maragogi/Maceio, all inclusive", "Resort all inclusive", "Benchmark de experiencia familia/praia; validar detalhes atuais.", certainty.comprovado],
  ["Cana Brava", "Resort all-inclusive", "Familias e praia Bahia", "Resort de praia com lazer completo", "All inclusive, lazer, Ilheus", "Benchmark para conteudo de familia e pacote completo; validar ofertas atuais.", certainty.validacao],
];

const internationalPlayers = [
  ["Four Seasons Private Residences", "Private residences, residence clubs e villa rentals", "Peace of mind, property management, home as center of world", "Servico reduz risco percebido. Alto padrao vende cuidado continuo, nao apenas arquitetura.", certainty.comprovado],
  ["Ritz-Carlton Residences", "Branded residences whole ownership", "Live Here, Always; concierge, in-residence dining, spa, valet", "A promessa e viver o padrao hoteleiro no cotidiano, com beneficios e servicos claros.", certainty.comprovado],
  ["Marriott Vacation Club", "Vacation ownership com club points", "Create Your Best Vacation Life; Marriott Quality; disclosures", "Conteudo educativo e transparencia legal protegem a venda de timeshare.", certainty.comprovado],
  ["Hilton Grand Vacations", "Vacation membership/timeshare", "Home resort, weeks, flexibility, member scale", "Usa escala e flexibilidade como prova; disclosures comerciais devem ser modelo.", certainty.comprovado],
  ["Rosewood Residences", "Ownership, serviced apartments e rentals", "Homes of remarkable character; sense of place", "Luxo sofisticado precisa de territorio cultural e lugar, nao so acabamento.", certainty.comprovado],
  ["Six Senses Residences", "Residences com resort amenities", "Nature, wellness, sustainability, private intimacy", "Wellness e sustentabilidade viram criterio de desejo quando sao operacionais, nao decorativos.", certainty.comprovado],
  ["One&Only Private Homes", "Private homes ultra-luxo", "Ultra-luxury design, comfort and taste", "Fala para quem compra raridade; pouca copy, muita imagem, localidade e acesso.", certainty.comprovado],
  ["Aman Residences", "Residences de altissimo luxo", "Permanent immersion into Aman lifestyle; select few", "Escassez e pertencimento sao mais fortes que desconto no topo da piramide.", certainty.comprovado],
  ["Mandarin Oriental Residences", "Branded residences", "Hospitality oriental, servico, status", "Benchmark de servico e reputacao; requer validacao de pagina especifica por projeto.", certainty.validacao],
  ["Hyatt Vacation Club", "Vacation ownership", "Rede, destinos, clube", "Benchmark de clube de ferias; validar detalhes atuais da oferta.", certainty.validacao],
  ["Exclusive Resorts", "Luxury destination club", "Travel like an Insider; living well over owning more", "Para ultra-premium, vender acesso e curadoria pode superar discurso de posse.", certainty.comprovado],
  ["Inspirato", "Luxury vacation club/subscription", "No Surprises. Just Standards; dedicated care", "Certeza operacional e ausencia de surpresa sao promessas poderosas para alto ticket.", certainty.comprovado],
  ["The Registry Collection", "Luxury exchange", "Luxury without limits; member-only exchange", "Intercambio premium combate objecao de repeticao do mesmo destino.", certainty.comprovado],
  ["Leading Hotels of the World", "Colecao hoteleira independente", "Pursue the Remarkable; Leaders Club", "Curadoria e independencia tambem sao marcas; nao precisa parecer rede massificada.", certainty.comprovado],
  ["Small Luxury Hotels", "Colecao boutique global", "Independently minded hotels; curated annually", "Autenticidade e curadoria podem diferenciar marcas menores de grandes redes.", certainty.comprovado],
  ["Disney Vacation Club", "Vacation ownership tematico", "Pertencimento, magia, resorts Disney", "Forte em comunidade e familia; validar fonte oficial antes de usar dados.", certainty.validacao],
];

const personas = [
  ["Familia planejadora classe A/B", "Quer ferias previsiveis com filhos sem improviso", "Casa de ferias com resort, seguranca e lazer", "Medo de taxa, disponibilidade e compromisso longo", "Comparativo multipropriedade vs hotel anual", "WhatsApp, Instagram, YouTube, Google", certainty.inferido],
  ["Empresario de alta renda", "Compra tempo, status e facilidade", "Nao se preocupar com operacao e manutencao", "Liquidez, reputacao da marca, entrega real", "Branded residence e gestao profissional", "Broker, evento privado, LinkedIn, PR", certainty.inferido],
  ["Casal 45+ filhos saindo de casa", "Busca conforto, gastronomia, praia e bem-estar", "Lugar fixo para descansar com padrao", "Nao quer bagunca de resort infantil", "Residence club com privacidade e servico", "YouTube, Google, Instagram, email", certainty.inferido],
  ["Avos patrocinadores", "Quer reunir filhos e netos todo ano", "Ferias como ritual familiar", "Agenda, espaco, custo de levar todos", "Historia de familia e previsibilidade", "Facebook, WhatsApp, video emocional", certainty.inferido],
  ["Investidor patrimonial cauteloso", "Quer ativo turistico com narrativa de valor", "Patrimonio + uso + renda potencial", "Promessa de rentabilidade, regulacao, liquidez", "Pagina tecnica e juridica sem promessa garantida", "Google Search, webinar, consultor", certainty.inferido],
  ["Profissional liberal aspiracional", "Quer entrar no mercado premium sem comprar imovel inteiro", "Acesso a resort e status familiar", "Tem medo de parecer golpe ou venda agressiva", "Conteudo educativo, simulação de uso, transparência", "Instagram, Meta Ads, WhatsApp", certainty.inferido],
  ["Morador regional", "Ja conhece destino e quer voltar sempre", "Beneficio local e uso recorrente", "Saturacao, custo mensal, validade", "Clube/residence com beneficio concreto", "Radio digital, Meta local, WhatsApp", certainty.inferido],
  ["Hospede encantado pos-estadia", "Acabou de viver a experiencia", "Transformar experiencia em recorrencia", "Medo de decidir por impulso", "Follow-up com comparativo e visita tecnica", "CRM, email, app, WhatsApp", certainty.inferido],
  ["Cliente de alto luxo internacional", "Quer residencia com servico, seguranca e marca", "Endereço raro e gestao sem friccao", "Confianca juridica e operacional no Brasil", "Dossie bilingue, arquitetura, reputacao", "Broker global, PR, landing EN/PT", certainty.inferido],
  ["Amante de parque aquatico/tematico", "Viaja pelo entretenimento dos filhos", "Acesso facil ao parque e estadia integrada", "Fila, lotacao, custo total", "Conteudo de experiencia familiar e beneficios", "TikTok, Reels, YouTube Shorts", certainty.inferido],
  ["Noivo/eventos premium", "Quer hospedar convidados com experiencia", "Casamento/evento com destino e estrutura", "Logistica e custo", "Pagina MICE/casamento com provas", "Google, Instagram, parceiros", certainty.inferido],
  ["Executivo sem tempo", "Delegar tudo e viajar sem planejar", "Concierge, app, padrao previsivel", "Nao quer reuniao longa nem venda insistente", "Fluxo consultivo curto e direto", "LinkedIn, WhatsApp, indicacao", certainty.inferido],
  ["Cliente que rejeita timeshare", "Tem experiencia ruim ou preconceito", "Entender diferenca juridica/operacional", "Venda agressiva, taxas escondidas", "Conteudo anti-confusao e tabela franca", "YouTube, blog, landing FAQ", certainty.inferido],
  ["Comprador de praia premium", "Quer destino bonito, seguro, com liquidez emocional", "Casa de praia sem gestao pesada", "Manutencao, locacao, distancia", "Branded residence/private villa", "Instagram, site visual, broker", certainty.inferido],
];

const objections = [
  ["Isso e timeshare?", "Explique modelo juridico e operacional exato, sem fugir do termo quando aplicavel.", "Artigo/video: multipropriedade, vacation club e residence club: diferencas reais.", "Nao demonizar timeshare nem prometer investimento.", certainty.inferido],
  ["Vou conseguir usar nas datas que quero?", "Mostrar regra de reserva, antecedencia, calendario, prioridade e exemplos.", "Simulador de reserva e FAQ visual.", "Prometer disponibilidade absoluta e arriscado.", certainty.inferido],
  ["Tem taxa escondida?", "Abrir taxa de manutencao, reajuste, impostos, clube, transferencia e regras.", "Landing com tabela de custos e o que esta incluso.", "Omissao de custos gera risco de politica e juridico.", certainty.inferido],
  ["E investimento?", "Separar uso, patrimonio e renda potencial; nao garantir retorno.", "Pagina: o que pode e o que nao pode ser prometido.", "Se houver pool/renda, revisar juridico/CVM.", certainty.inferido],
  ["E se eu quiser vender?", "Explicar regras de cessao/revenda, mercado secundario e custos.", "FAQ de liquidez e revenda.", "Nao prometer liquidez rapida.", certainty.inferido],
  ["Por que nao reservar hotel quando eu quiser?", "Comparar previsibilidade, padrao, custo total e pertencimento, sem esconder flexibilidade do hotel.", "Tabela hotel anual vs propriedade fracionada.", "Nao manipular comparacao com dados irreais.", certainty.inferido],
  ["O destino pode cansar?", "Apresentar intercambios, rede, clube, semanas alternadas ou beneficios em outros destinos quando existirem.", "Conteudo: como variar destino dentro do modelo.", "Se nao houver rede, nao inventar.", certainty.inferido],
  ["Tenho medo de construtora nao entregar.", "Mostrar historico, obras, memorial, garantias, documentos, videos de obra e auditorias.", "Serie de bastidores da obra.", "Nao usar prova que nao existe.", certainty.inferido],
  ["Nao quero reuniao de venda agressiva.", "Criar jornada com diagnostico, material enviado antes e decisao sem pressao.", "Roteiro de WhatsApp consultivo.", "Pressao excessiva prejudica premium.", certainty.inferido],
  ["Minha familia vai usar mesmo?", "Fazer calculo de ritual anual, ferias escolares e perfil de uso.", "Quiz: qual modelo de ferias combina com sua familia?", "Nao vender para quem nao tem fit.", certainty.inferido],
  ["E caro.", "Reposicionar como custo de ferias recorrentes + acesso + servico, comparando com alternativas reais.", "Calculadora de custo de ferias em 5 anos.", "Cuidado com economia garantida.", certainty.inferido],
  ["Ja ouvi falar mal desse mercado.", "Reconhecer a reputacao do segmento e diferenciar por transparência, contrato e entrega.", "Video: 5 perguntas antes de comprar qualquer cota.", "Nao atacar concorrentes.", certainty.inferido],
  ["Nao entendi o que estou comprando.", "Desenhar: direito, uso, custos, regras, prazo, unidade, escritura/contrato.", "Infografico do modelo.", "Clareza vem antes de persuasao.", certainty.inferido],
  ["Prefiro Airbnb.", "Comparar autonomia do Airbnb com padrao, servico e previsibilidade do resort.", "Post comparativo sem desmerecer Airbnb.", "Nao fingir que Airbnb nao tem vantagens.", certainty.inferido],
  ["Minha esposa/marido precisa entender.", "Criar material para decisor secundario, com resumo de riscos e beneficios.", "PDF de 2 paginas para levar para casa.", "Venda high-ticket raramente e individual.", certainty.inferido],
  ["Nao confio em promessa de valorizacao.", "Nao basear venda em valorizacao; usar uso, experiencia, marca e documentos.", "Conteudo: compre pelo uso antes de olhar valorizacao.", "Promessa de ganho e risco alto.", certainty.inferido],
  ["Como funciona heranca/sucessao?", "Encaminhar para explicacao juridica validada e documentos oficiais.", "FAQ juridico revisado.", "Precisa juridico; nao improvisar.", certainty.validacao],
  ["Posso alugar minha semana?", "Explicar regras reais do contrato/operacao.", "FAQ de uso, locacao e cessao.", "Depende do empreendimento.", certainty.validacao],
  ["E se o parque/resort lotar?", "Mostrar beneficios de acesso, horarios, reservas e regras se existirem.", "Conteudo de experiencia real em alta temporada.", "Nao prometer exclusividade inexistente.", certainty.validacao],
  ["Qual a diferenca para clube de ferias?", "Comparar propriedade, uso, pontos, hospedagem e custos.", "Tabela multipropriedade vs clube de ferias.", "Evitar misturar produto.", certainty.inferido],
  ["Por que comprar agora?", "Usar janela real: fase de obra, tabela, disponibilidade, condicao, evento.", "Campanha com motivo concreto.", "Urgencia falsa e ruim para marca premium.", certainty.inferido],
  ["Vai ficar abandonado fora de temporada?", "Mostrar calendario de eventos, ocupacao, plano comercial e manutencao.", "Conteudo destino 12 meses.", "Dados de ocupacao exigem fonte.", certainty.validacao],
];

const channels = [
  ["Instagram", "Desejo, prova social, bastidores, autoridade leve", "Reels de destino, carrossel educativo, stories com enquete, live com consultor", "Salvar, DM, clique no WhatsApp, visita ao perfil", "Bonito demais e claro de menos; excesso de flyer institucional", certainty.inferido],
  ["TikTok", "Descoberta e educacao rapida", "Ganchos de objecao, visita guiada, comparativos, mitos do mercado", "Retencao, compartilhamento, comentario, visita ao perfil", "Promessa exagerada, linguagem financeira agressiva, video muito corporativo", certainty.inferido],
  ["YouTube", "Autoridade, busca e decisao de alto ticket", "Tours longos, explicadores juridicos, series de destino, entrevistas, comparativos", "Watch time, leads qualificados, buscas de marca", "Video bonito sem resposta a objecoes; thumbnail generica", certainty.inferido],
  ["Facebook/Meta Ads", "Geração de demanda regional, remarketing e lead", "Criativos claros por persona, video curto, carrossel, formulario qualificado", "CPL qualificado, taxa de resposta WhatsApp, agendamento", "Personal attributes, promessas de renda, landing divergente", certainty.comprovado],
  ["Google Search", "Capturar intencao ativa", "Campanhas por destino, modelo de negocio, marca e comparativos", "Custo por lead qualificado, taxa de visita, taxa de agendamento", "Comprar palavra ampla sem landing educativa", certainty.inferido],
  ["SEO/Site", "Educar e converter com confianca", "Glossario, FAQ, comparativos, paginas de destino, documentos e prova", "Tráfego orgânico, tempo na pagina, conversao WhatsApp", "Site de luxo sem clareza juridica/comercial", certainty.comprovado],
  ["WhatsApp/CRM", "Converter high-ticket com diagnostico", "Roteiro de qualificacao, materiais por objeção, follow-up 1/3/7/14 dias", "Tempo de resposta, show rate, proposta, venda", "Disparo frio sem contexto; vendedor empurrando cota", certainty.inferido],
  ["App/Area do cliente", "Retencao, uso e prova operacional", "Reserva, calendario, documentos, beneficios, concierge, conteudo do destino", "Ativacao, reservas, NPS, upgrade, indicacao", "App sem utilidade real vira vitrine morta", certainty.inferido],
  ["Eventos e experiencias", "Vender sensorialmente para alto ticket", "Open house, fim de semana convite, jantar, familia test-drive", "Comparecimento, proposta, venda pos-evento", "Evento lotado sem curadoria reduz percepcao premium", certainty.inferido],
  ["PR e imprensa", "Dar legitimidade ao empreendimento", "Matérias de arquitetura, turismo, economia local, hospitalidade", "Menções, backlinks, busca de marca", "Release publicitario sem fato noticiavel", certainty.inferido],
];

const claimCompliance = [
  ["Voce pode ter sua casa de ferias em um resort.", "Seguro se juridicamente verdadeiro e explicado.", "Use com definicao do modelo, custos e regras.", certainty.inferido],
  ["Invista e tenha retorno garantido.", "Alto risco.", "Trocar por: conheca o potencial de uso, locacao e valorizacao, sem garantia de retorno.", certainty.comprovado],
  ["Ferias planejadas todos os anos.", "Seguro se regras de uso sustentarem.", "Adicionar: conforme calendario, contrato e disponibilidade do modelo.", certainty.inferido],
  ["Acesso a estrutura de resort sem comprar um imovel inteiro.", "Forte e seguro se o produto for fracionado.", "Explicar diferenca entre compra integral e fracionada.", certainty.inferido],
  ["O melhor investimento turistico do Brasil.", "Risco por superlativo sem prova.", "Trocar por: uma alternativa para quem busca uso recorrente em destino turistico.", certainty.inferido],
  ["Sem preocupacao com manutencao.", "Pode ser forte se houver gestao profissional.", "Trocar por: manutencao administrada conforme regras do empreendimento.", certainty.inferido],
  ["Ganhe dinheiro alugando sua semana.", "Risco financeiro e regulatorio.", "Trocar por: quando permitido, entenda as regras de locacao/cessao.", certainty.inferido],
  ["Ultimas unidades.", "Seguro apenas se real e documentado.", "Usar so com controle comercial validado.", certainty.validacao],
  ["A familia toda aproveita.", "Seguro como apelo de experiencia.", "Mostrar estrutura concreta: parque, recreacao, quartos, restaurantes.", certainty.inferido],
  ["Seu patrimonio em destino turistico.", "Precisa cuidado.", "Usar se houver propriedade imobiliaria real; nao usar para clube puro.", certainty.validacao],
  ["Mais barato que hotel.", "Risco se sem comparativo.", "Trocar por: compare o custo de ferias recorrentes no seu perfil de uso.", certainty.inferido],
  ["Padrao internacional.", "Generico e arriscado sem prova.", "Trocar por provas: marca, arquitetura, servicos, premios, parceiros.", certainty.inferido],
  ["Exclusividade.", "Usar com criterio.", "Especificar: acesso, unidades, servico, clube, prioridade ou privacidade.", certainty.inferido],
  ["Rentabilidade acima do mercado.", "Alto risco.", "Evitar em anuncio; se existir, precisa oferta regulada e juridico.", certainty.comprovado],
  ["Experiencia completa de ferias.", "Seguro se sustentado por resort/parque/gastronomia.", "Listar o que torna completa.", certainty.inferido],
];

const copyBank = [
  ["Headline", "Topo site", "Sua casa de ferias dentro de um resort, com regras claras para usar, planejar e voltar todos os anos.", "Clareza + desejo", "Baixo"],
  ["Headline", "Topo site", "Imobiliario turistico para quem quer ferias recorrentes sem administrar uma casa sozinho.", "Diferenciacao", "Baixo"],
  ["Headline", "Landing", "Antes de comprar uma cota de ferias, entenda exatamente o que voce esta comprando.", "Educacao", "Baixo"],
  ["Headline", "Landing", "Um resort nao vende apenas hospedagem. Vende rotina de ferias, servico e pertencimento.", "Reposicionamento", "Baixo"],
  ["Headline", "Premium", "Residencias de ferias com servico, privacidade e a conveniencia de uma operacao hoteleira.", "Luxo", "Baixo"],
  ["Headline", "Comparativo", "Multipropriedade, clube de ferias ou segunda casa: qual modelo combina com sua familia?", "Busca ativa", "Baixo"],
  ["Headline", "Google Ads", "Casa de ferias em resort: entenda modelos, custos e regras antes de decidir.", "Intencao", "Baixo"],
  ["Headline", "Facebook Ads", "Planeje as ferias da sua familia em um resort, com orientacao antes da compra.", "Lead seguro", "Baixo"],
  ["Headline", "TikTok/Reels", "Se voce acha que multipropriedade e tudo igual, veja isso antes de decidir.", "Objecao", "Medio"],
  ["Headline", "YouTube", "O guia honesto para entender resort, multipropriedade e vacation club no Brasil.", "Autoridade", "Baixo"],
  ["Gancho video", "Reels", "A pergunta certa nao e quanto custa a cota. E como sua familia realmente viaja.", "Diagnostico", "Baixo"],
  ["Gancho video", "TikTok", "Tem gente que compra ferias como se comprasse impulso. Esse e o erro.", "Contraste", "Medio"],
  ["Gancho video", "YouTube Shorts", "Antes de falar em investimento, voce precisa entender o uso.", "Compliance", "Baixo"],
  ["Gancho video", "Reels", "O que diferencia um resort imobiliario serio de uma venda bonita demais?", "Confianca", "Baixo"],
  ["Gancho video", "TikTok", "Se o vendedor nao explica taxa de manutencao, disponibilidade e revenda, pare a conversa.", "Protecao", "Medio"],
  ["Gancho video", "Reels", "Comprar uma casa de praia parece simples ate a manutencao virar rotina.", "Dor", "Baixo"],
  ["Gancho video", "Shorts", "O luxo aqui nao e ter a chave. E chegar e tudo estar pronto.", "Premium", "Baixo"],
  ["Gancho video", "Reels", "Resort para familia vende uma coisa que hotel comum raramente entrega: repeticao com previsibilidade.", "Desejo", "Baixo"],
  ["Gancho video", "TikTok", "Nao compre uma promessa de ferias. Compre um modelo que voce consegue explicar.", "Clareza", "Baixo"],
  ["Gancho video", "YouTube", "Multipropriedade no Brasil: o que a lei permite, o que o contrato define e o que precisa ser validado.", "Autoridade", "Baixo"],
  ["CTA", "Frio", "Receba o guia dos modelos de ferias em resort antes de falar com um consultor.", "Baixa friccao", "Baixo"],
  ["CTA", "Morno", "Veja uma simulação de uso para o perfil da sua familia.", "Diagnostico", "Baixo"],
  ["CTA", "Quente", "Agende uma conversa consultiva e tire suas duvidas sobre regras, custos e disponibilidade.", "Conversao", "Baixo"],
  ["CTA", "Premium", "Solicite o dossie privado do empreendimento.", "Exclusividade", "Baixo"],
  ["CTA", "Pos-estadia", "Transforme sua experiencia no resort em um plano de ferias recorrente.", "Recorrencia", "Baixo"],
  ["WhatsApp", "Primeira resposta", "Antes de te mandar proposta, preciso entender como sua familia costuma viajar. Assim eu nao te mostro um modelo que nao faz sentido.", "Consultivo", "Baixo"],
  ["WhatsApp", "Objecao taxa", "Perfeito perguntar isso. Vou te enviar a tabela do que e custo de aquisicao, o que e recorrente e o que depende de uso.", "Transparencia", "Baixo"],
  ["WhatsApp", "Objecao investimento", "Eu prefiro separar duas coisas: uso de ferias e expectativa patrimonial. Posso te mostrar o modelo sem tratar retorno como garantia.", "Compliance", "Baixo"],
  ["WhatsApp", "Follow-up 3 dias", "Conseguiu olhar a parte de disponibilidade e custos? Normalmente e ali que a decisao fica clara.", "Follow-up", "Baixo"],
  ["WhatsApp", "Follow-up pos-visita", "Pelo que voce comentou, o ponto central nao e so o resort; e ter um lugar confiavel para reunir a familia todos os anos.", "Personalizacao", "Baixo"],
  ["Email", "Educativo", "O que voce precisa entender antes de comprar qualquer produto de ferias compartilhadas.", "Autoridade", "Baixo"],
  ["Email", "Comparativo", "Hotel, casa de praia, vacation club ou multipropriedade: vantagens e limites de cada caminho.", "Comparativo", "Baixo"],
  ["Carrossel", "Instagram", "7 perguntas que um consultor serio deve responder antes de vender uma cota.", "Salvamento", "Baixo"],
  ["Carrossel", "Instagram", "Casa de ferias nao e so uma unidade. E um sistema de uso, custos, regras e servico.", "Educacao", "Baixo"],
  ["Anuncio", "Meta", "Quer entender se um modelo de ferias em resort faz sentido para sua familia? Receba um comparativo simples antes de conversar com um consultor.", "Lead", "Baixo"],
  ["Anuncio", "Google", "Compare modelos de casa de ferias em resort e entenda custos, uso e disponibilidade.", "Busca", "Baixo"],
  ["Anuncio", "Remarketing", "Ainda em duvida? Veja as respostas para as perguntas que mais travam a decisao.", "Objecoes", "Baixo"],
  ["Story", "Enquete", "O que mais pesa na escolha das ferias: destino, custo total, estrutura para criancas ou conforto?", "Interacao", "Baixo"],
  ["Story", "Caixa", "Qual duvida voce teria antes de comprar uma casa de ferias em resort?", "Pesquisa", "Baixo"],
  ["YouTube title", "Conteudo longo", "Como vender resort imobiliario sem parecer venda agressiva de cota", "B2B", "Baixo"],
  ["YouTube title", "Conteudo longo", "A nova linguagem do imobiliario turistico premium no Brasil", "Branding", "Baixo"],
];

const plan30 = Array.from({ length: 30 }, (_, i) => {
  const day = i + 1;
  const phases = [
    "Diagnostico e base",
    "Arquitetura de oferta",
    "Conteudo e prova",
    "Trafego e CRM",
    "Otimizacao",
  ];
  const phase = day <= 5 ? phases[0] : day <= 10 ? phases[1] : day <= 17 ? phases[2] : day <= 24 ? phases[3] : phases[4];
  const actions = {
    1: "Reunir contratos, regras de uso, tabela de custos, memorial, fotos, plantas, politica de reserva e historico da marca.",
    2: "Mapear concorrentes nacionais e internacionais por modelo: multipropriedade, residence club, vacation club e branded residence.",
    3: "Definir produto exato: o que e propriedade, o que e direito de uso, o que e clube e o que e beneficio.",
    4: "Criar matriz de personas por poder aquisitivo, destino, motivacao, objecao e canal.",
    5: "Auditar promessas proibidas: retorno garantido, valorizacao, disponibilidade absoluta e descontos falsos.",
    6: "Reposicionar oferta em uma frase: para quem e, qual modelo, qual destino, qual beneficio e qual limite.",
    7: "Criar pagina FAQ com custos, uso, reserva, revenda, manutencao e documentos.",
    8: "Criar comparativo entre hotel, segunda casa, multipropriedade e clube de ferias.",
    9: "Definir lead magnet: guia, quiz, simulador de uso ou dossie privado.",
    10: "Montar roteiro de atendimento consultivo para WhatsApp e venda.",
    11: "Produzir 5 Reels/TikToks de objecoes e mitos do mercado.",
    12: "Produzir 3 videos YouTube: guia do modelo, tour do empreendimento e comparativo de alternativas.",
    13: "Produzir carrossel de perguntas essenciais antes da compra.",
    14: "Criar banco de fotos/videos por prova: destino, unidade, servico, familia, bastidor e documentos.",
    15: "Publicar pagina de destino com CTA para guia/WhatsApp qualificado.",
    16: "Criar sequencia de email/WhatsApp para lead frio, morno e pos-visita.",
    17: "Treinar vendedores para responder objecoes sem prometer investimento.",
    18: "Subir campanha Google Search por termos de alta intencao.",
    19: "Subir campanha Meta com lead magnet e criativos por persona.",
    20: "Configurar remarketing para visitantes da pagina e engajados no Instagram/Video.",
    21: "Criar publico de pos-estadia e base de clientes/hospedes se houver consentimento LGPD.",
    22: "Implantar tags, eventos e UTMs para medir origem, lead qualificado e agendamento.",
    23: "Acompanhar conversas de WhatsApp e classificar objecoes reais.",
    24: "Criar relatorio de CPL, taxa de resposta, show rate e proposta enviada.",
    25: "Pausar criativos com leads ruins e reforcar angulos que geram perguntas qualificadas.",
    26: "Transformar as 10 duvidas mais comuns em conteudo organico.",
    27: "Revisar pagina conforme objeções reais: clareza de custo, uso, contrato e prova.",
    28: "Criar campanha de remarketing com FAQ e video de consultor.",
    29: "Treinar time com chamadas gravadas e matriz de respostas.",
    30: "Fechar decisao de escala: canais, verba, criativos vencedores e ajustes de oferta.",
  };
  return [day, phase, actions[day], "Responsavel marketing/vendas", day <= 5 ? "Base validada" : day <= 17 ? "Ativo publicado" : "Metrica e proxima acao", "RASCUNHO PARA APROVACAO"];
});

const siteChecklist = [
  ["Hero", "Explicar em 5 segundos o modelo: residencia/casa de ferias/resort/clube, destino e CTA.", "H1 especifico, subtitulo com modelo, CTA para guia ou consultor, imagem real.", "COMPROVADO/INFERIDO"],
  ["Prova de destino", "Mostrar por que o lugar merece repeticao anual.", "Mapa, distancias, atracoes, sazonalidade, videos reais.", "INFERIDO"],
  ["Modelo de compra", "Evitar confusao entre propriedade, clube e hospedagem.", "Infografico de como funciona, regras, contrato, disponibilidade.", "INFERIDO"],
  ["Custos", "Reduzir medo de taxa escondida.", "Tabela de aquisicao, recorrente, manutencao, impostos e opcionais.", "INFERIDO"],
  ["Comparativo", "Ajudar decisor a comparar alternativas.", "Hotel x Airbnb x segunda casa x multipropriedade x clube.", "INFERIDO"],
  ["FAQ juridico", "Responder objeções antes do vendedor.", "Lei, escritura/contrato, cessao, heranca, revenda, uso.", "PRECISA VALIDACAO"],
  ["Prova operacional", "Mostrar que existe entrega.", "Obras, hoteis existentes, marcas, gestores, NPS se houver.", "PRECISA VALIDACAO"],
  ["Conteudo SEO", "Capturar busca ativa.", "Glossario, paginas por destino, comparativos, guias.", "COMPROVADO"],
  ["CTA", "Adequar ao funil.", "Frio: guia/quiz. Morno: simulacao. Quente: consultor.", "INFERIDO"],
  ["Compliance", "Evitar bloqueio e risco juridico.", "Sem retorno garantido, sem disponibilidade absoluta, sem omitir taxas.", "COMPROVADO"],
];

const workbook = Workbook.create();

function addSheet(name, headers, rows, widths = []) {
  const ws = workbook.worksheets.add(name);
  ws.showGridLines = false;
  const data = [headers, ...rows];
  const rowCount = data.length;
  const colCount = headers.length;
  const range = ws.getRangeByIndexes(0, 0, rowCount, colCount);
  range.values = data;
  ws.freezePanes.freezeRows(1);
  range.format = { wrapText: true, borders: { preset: "all", style: "thin", color: "#D9D9D9" } };
  const header = ws.getRangeByIndexes(0, 0, 1, colCount);
  header.format = { fill: "#16324F", font: { bold: true, color: "#FFFFFF" }, wrapText: true };
  try {
    const lastCol = String.fromCharCode("A".charCodeAt(0) + colCount - 1);
    const table = ws.tables.add(`A1:${lastCol}${rowCount}`, true, name.replace(/[^A-Za-z0-9]/g, "_").slice(0, 30));
    table.style = "TableStyleMedium2";
    table.showFilterButton = true;
  } catch (error) {
    // Table styling is helpful but not required for the workbook to be useful.
  }
  for (let c = 0; c < colCount; c++) {
    ws.getRangeByIndexes(0, c, rowCount, 1).format.columnWidthPx = widths[c] ? widths[c] * 8 : 220;
  }
  ws.getRangeByIndexes(1, 0, Math.max(rowCount - 1, 1), colCount).format.rowHeightPx = 92;
  return ws;
}

addSheet("00 Leia Primeiro", ["Campo", "Conteudo"], [
  ["Objetivo", "Base profunda para estudar, posicionar e vender imobiliario turistico premium: resort, multipropriedade, residence club, branded residences, vacation club e modelos adjacentes."],
  ["Uso recomendado", "Ler o dossie primeiro, depois usar as abas como base operacional de pesquisa, copy, canais, objeções, compliance e plano de 30 dias."],
  ["Regra de certeza", "Cada linha relevante marca COMPROVADO, INFERIDO, HIPOTESE ou PRECISA VALIDACAO. Nao transforme inferencia em promessa comercial."],
  ["Regra comercial", "Vender primeiro clareza, destino, uso, experiencia, servico e confianca. So falar de patrimonio/renda quando houver suporte juridico e documental."],
  ["Status", "RASCUNHO PARA APROVACAO. Base montada com fontes abertas e deve ser complementada com contratos, anuncios reais, analytics, CRM e auditoria juridica."],
], [24, 110]);
addSheet("01 Modelos Negocio", ["Modelo", "Como funciona", "Publico", "Promessa forte", "Objecao central", "Canais prioritarios", "Certeza"], businessModels, [28, 46, 30, 34, 34, 34, 18]);
addSheet("02 Players Brasil", ["Empresa", "Modelo", "Publico/territorio", "Ativo de marca", "Linguagem observada", "Aprendizado aplicavel", "Certeza"], brazilPlayers, [26, 30, 34, 34, 36, 48, 18]);
addSheet("03 Players Internacionais", ["Empresa", "Modelo", "Linguagem/posicionamento", "Aprendizado aplicavel ao Brasil", "Certeza"], internationalPlayers, [30, 32, 44, 55, 18]);
addSheet("04 Personas", ["Persona", "Situacao", "Desejo", "Objecao", "Mensagem principal", "Canais", "Certeza"], personas, [30, 36, 34, 34, 46, 30, 18]);
addSheet("05 Objecoes", ["Objecao", "Resposta estrategica", "Conteudo para quebrar", "Risco", "Certeza"], objections, [34, 50, 46, 42, 18]);
addSheet("06 Canais", ["Canal", "Funcao", "Formatos", "Metrica", "Erro comum", "Certeza"], channels, [24, 34, 46, 36, 42, 18]);
addSheet("07 Site CRO SEO", ["Area", "Funcao", "O que precisa ter", "Certeza"], siteChecklist, [26, 42, 62, 24]);
addSheet("08 Claims Compliance", ["Claim", "Risco/Leitura", "Versao segura", "Certeza"], claimCompliance, [38, 44, 58, 18]);
addSheet("09 Banco Copy", ["Tipo", "Canal/uso", "Texto", "Funcao", "Risco"], copyBank, [20, 24, 78, 24, 18]);
addSheet("10 Plano 30 Dias", ["Dia", "Fase", "Acao", "Responsavel", "Entregavel", "Status"], plan30, [8, 24, 72, 28, 34, 26]);
addSheet("11 Fontes", ["Fonte", "URL", "Uso na pesquisa", "Certeza"], sources, [34, 58, 80, 18]);

function table(headers, rows) {
  const head = `| ${headers.join(" | ")} |`;
  const sep = `| ${headers.map(() => "---").join(" | ")} |`;
  const body = rows.map((row) => `| ${row.map((v) => String(v).replace(/\n/g, "<br>").replace(/\|/g, "\\|")).join(" | ")} |`).join("\n");
  return [head, sep, body].join("\n");
}

const md = `# Dossie Estrategico - Imobiliario Turistico Premium, Resorts e Multipropriedade

Status: RASCUNHO PARA APROVACAO  
Data: 2026-06-18  
Escopo: Brasil com benchmarks internacionais de branded residences, vacation ownership, residence clubs, destination clubs e hotelaria premium.

## 1. Diagnostico objetivo

O erro da pesquisa rasa e tratar isso como "multipropriedade" isolada. O mercado real e uma intersecao entre imobiliario de alto padrao, turismo, hospitalidade, ferias planejadas, clube de beneficios, familia, status, recorrencia e venda consultiva high-ticket.

**COMPROVADO:** players nacionais como GAV, Aviva, Beach Park, Hot Beach, Gramado Parks, WAM, diRoma e Pratagy comunicam ecossistemas que misturam hospedagem, parque, clube, destino e venda recorrente de ferias.  
**COMPROVADO:** players internacionais como Four Seasons, Ritz-Carlton, Rosewood, Six Senses e Aman vendem residencias com linguagem de servico, lugar, privacidade, reputacao e estilo de vida.  
**INFERIDO:** no Brasil, a oportunidade de marketing esta em subir o nivel do discurso: sair de "compre uma cota" para "entenda qual modelo de ferias, propriedade, servico e destino faz sentido para sua familia".  
**HIPOTESE:** a maior vantagem competitiva para uma marca nova ou media sera criar a pagina, o conteudo e o atendimento mais claros do mercado, porque a categoria sofre com confusao, desconfianca e venda agressiva.

## 2. O problema real de marketing

O publico nao compra "multipropriedade". Ele compra uma combinacao de:

- Ferias recorrentes sem improviso.
- Um endereco emocional para a familia.
- Conveniencia de resort sem administrar uma casa sozinho.
- Sensacao de pertencer a um clube ou destino.
- Acesso a estrutura, servico, parque, praia, gastronomia ou natureza.
- Reducao de incerteza sobre viagem, hospedagem e manutencao.
- Em alguns casos, narrativa patrimonial, mas sem promessa de retorno garantido.

O problema comercial e que o segmento costuma vender rapido demais a unidade, a cota ou a condicao. Para alto padrao, isso derruba confianca.

## 3. Glossario essencial

${table(["Termo", "Definicao pratica", "Como vender sem confundir"], [
  ["Multipropriedade", "Propriedade fracionada por tempo de uso, regulada no Brasil pela Lei 13.777/2018.", "Explicar unidade, periodo, calendario, custos e direito real quando aplicavel."],
  ["Fractional ownership", "Propriedade fracionada, geralmente com linguagem premium e patrimonial.", "Mostrar diferenca para clube puro e para segunda casa integral."],
  ["Vacation club", "Clube de ferias por pontos, beneficios, descontos ou acesso a rede.", "Vender flexibilidade, nao propriedade se nao houver propriedade."],
  ["Residence club", "Clube residencial com uso recorrente, servicos e padrao de resort.", "Vender pertencimento, uso e conveniencia."],
  ["Branded residences", "Residencias associadas a marca hoteleira ou de luxo.", "Vender reputacao, servico, design, gestao e lugar."],
  ["Condo-hotel", "Unidade em operacao hoteleira, muitas vezes com pool de locacao.", "Exige cuidado juridico se comunicar renda."],
  ["Destination club", "Clube premium com acesso a portfolio de casas/residencias em destinos.", "Vender acesso, curadoria e tempo, nao posse."],
  ["Private villas", "Villas/casas dentro ou associadas a resort, com servicos.", "Vender privacidade com conveniencia hoteleira."],
  ["Resort membership", "Associacao com beneficios e acessos recorrentes.", "Vender recorrencia e vantagens claras."],
  ["Exchange club", "Clube de troca de semanas/pontos entre destinos.", "Vender variedade para combater objecao de destino repetido."]
])}

## 4. Mapa dos modelos de negocio

${table(["Modelo", "Promessa forte", "Objecao central", "Certeza"], businessModels.map((r) => [r[0], r[3], r[4], r[6]]))}

## 5. O que os players brasileiros ensinam

${table(["Player", "Modelo", "Linguagem observada", "Aprendizado"], brazilPlayers.slice(0, 18).map((r) => [r[0], r[1], r[4], r[5]]))}

Leitura estrategica:

- **GAV, Hot Beach, Aviva e Beach Park** mostram que a venda fica mais forte quando existe uma ancora de experiencia: parque, destino, resort, clube ou ecossistema.
- **Fasano, Txai e Nannai** mostram a logica do luxo: menos pressa, menos promocao, mais lugar, sensacao, reputacao e curadoria.
- **JHSF** mostra a logica de ecossistema premium: luxo nao e so produto, e rede de acesso, servico e pertencimento.
- **diRoma, WAM e redes regionais** mostram que volume depende de oferta clara, beneficios e canais diretos. O risco e comoditizar a marca se tudo virar desconto.

## 6. O que os benchmarks internacionais ensinam

${table(["Benchmark", "Linguagem", "Licao para o Brasil"], internationalPlayers.slice(0, 12).map((r) => [r[0], r[2], r[3]]))}

Resumo critico:

- **Four Seasons e Ritz-Carlton** vendem paz operacional: propriedade com servico, gestao e padrao.
- **Rosewood e Six Senses** vendem identidade do lugar, bem-estar e estilo de vida, nao apenas unidade.
- **Aman e One&Only** vendem raridade. A copy e economica porque a marca, a imagem e a escassez carregam a percepcao.
- **Marriott Vacation Club e Hilton Grand Vacations** sao importantes pelo lado de compliance: educam o modelo e usam disclosures claros.
- **Exclusive Resorts e Inspirato** vendem uma tese poderosa: viver melhor pode ser mais desejavel do que possuir mais.

## 7. Personas prioritarias

${table(["Persona", "Desejo", "Objecao", "Mensagem"], personas.map((r) => [r[0], r[2], r[3], r[4]]))}

## 8. Objecoes que precisam virar conteudo

${table(["Objecao", "Resposta estrategica", "Conteudo"], objections.slice(0, 18).map((r) => [r[0], r[1], r[2]]))}

## 9. Estrategia por canal

### Instagram
Funcao: desejo, prova social, bastidor, lifestyle e educacao leve.  
Fazer: Reels de objecao, carrossel de comparativo, stories de pergunta, bastidores de obra, tour de unidade, destino, familia usando estrutura.  
Evitar: feed bonito sem explicar modelo.

### TikTok
Funcao: descoberta e quebra de preconceito.  
Fazer: mitos, comparativos, "antes de comprar", visitas guiadas, respostas a comentarios.  
Evitar: video institucional e promessa financeira.

### YouTube
Funcao: autoridade e decisao.  
Fazer: guia completo, tour longo, entrevista com consultor, serie sobre destino, comparativos e documentos.  
Evitar: video que parece propaganda sem responder medos reais.

### Facebook/Meta Ads
Funcao: demanda regional, remarketing e lead.  
Fazer: criativos por persona, lead magnet, formularios qualificados e remarketing de FAQ.  
Evitar: atributos pessoais sensiveis, promessa de renda, omissao de custos e landing incoerente.

### Google Search
Funcao: captar intencao ativa.  
Fazer: termos de destino, modelo e comparativo. Ex.: "casa de ferias em resort", "multipropriedade como funciona", "residence club Brasil".  
Evitar: mandar busca tecnica para pagina emocional sem FAQ.

### Site
Funcao: tirar medo e qualificar.  
Fazer: H1 claro, modelo explicado, custos, regras, comparativo, FAQ juridico, prova de entrega, CTA por funil.  
Evitar: site de luxo que nao explica o que a pessoa compra.

### WhatsApp/CRM
Funcao: converter high-ticket.  
Fazer: diagnostico de perfil de viagem, follow-up por objecao, documentos, simulacao e material para decisor secundario.  
Evitar: abordagem agressiva de "vamos fechar hoje".

### App
Funcao: reter, provar entrega e estimular upgrade/indicacao.  
Fazer: reserva, calendario, documentos, beneficios, concierge, conteudo de destino, status da obra, extrato de uso.  
Evitar: app que so replica site.

## 10. Funil recomendado

1. **Descoberta:** videos curtos sobre dores e mitos: "isso e timeshare?", "vale mais hotel ou casa de ferias?", "o que ninguem explica sobre taxa de manutencao?".
2. **Educacao:** guia, comparativo, FAQ e calculadora de perfil de uso.
3. **Prova:** tour, bastidores, destino, depoimentos, estrutura, documentos, marca operadora, status da obra.
4. **Diagnostico:** quiz ou WhatsApp consultivo para entender familia, frequencia de viagem, datas e objecoes.
5. **Proposta:** simulação de uso, custos e regras. Sem prometer retorno garantido.
6. **Decisao:** reunião com decisores, material resumido, respostas juridicas e prazo real.
7. **Pos-venda:** onboarding, app, calendario, checklists e incentivo a indicacao.

## 11. Estrutura ideal do site

${table(["Area", "Funcao", "O que precisa ter"], siteChecklist.map((r) => [r[0], r[1], r[2]]))}

H1 recomendado:

> Sua casa de ferias dentro de um resort, com regras claras para usar, planejar e voltar todos os anos.

Subtitulo:

> Entenda como funciona o modelo, quais custos entram na decisao, como reservar suas datas e quando esse tipo de produto faz sentido para sua familia.

CTA frio:

> Receber o guia dos modelos de ferias em resort

CTA morno:

> Simular meu perfil de uso

CTA quente:

> Falar com um consultor

## 12. Banco inicial de copy e ganchos

${table(["Tipo", "Uso", "Texto"], copyBank.slice(0, 28).map((r) => [r[0], r[1], r[2]]))}

## 13. Regras de compliance e seguranca

${table(["Claim", "Leitura", "Versao segura"], claimCompliance.map((r) => [r[0], r[1], r[2]]))}

Regra pratica:

- Pode vender uso, experiencia, conveniencia, destino, servico, marca e clareza.
- Pode falar de patrimonio quando o modelo juridico permitir e documentos sustentarem.
- Nao deve prometer retorno, valorizacao, renda, liquidez, disponibilidade absoluta ou economia garantida.
- Se houver renda, pool de locacao ou tese financeira, precisa revisao juridica e regulatoria antes de anuncio.
- Lead ads, formularios e CRM precisam respeitar consentimento, finalidade e protecao de dados.

## 14. Plano de ataque de 30 dias

${table(["Dia", "Fase", "Acao"], plan30.map((r) => [r[0], r[1], r[2]]))}

## 15. Prompt poderoso para continuar pesquisando

\`\`\`text
Voce e um pesquisador senior de mercado imobiliario turistico, hotelaria premium, multipropriedade, vacation ownership, branded residences, residence clubs, destination clubs, resorts e vendas high-ticket.

Objetivo: construir uma base profunda e verificavel para posicionar, vender e fazer marketing de empreendimentos de resort imobiliario no Brasil.

Contexto: nao trate o mercado como apenas "multipropriedade". Analise a intersecao entre imovel de alto padrao, segunda residencia, turismo, hospitalidade, clube de ferias, familia, experiencia, servico, patrimonio, compliance e venda consultiva.

Processo:
1. Pesquise fontes oficiais de players brasileiros e internacionais.
2. Separe modelos de negocio: multipropriedade, fractional ownership, residence club, vacation club, branded residences, condo-hotel, destination club, private villas, resort membership e exchange club.
3. Para cada empresa, registre: proposta, publico, linguagem, canais, ativos de prova, CTA, beneficios, objecoes que responde, riscos e aprendizados aplicaveis.
4. Classifique cada conclusao como COMPROVADO, INFERIDO, HIPOTESE ou PRECISA VALIDACAO.
5. Nao invente metricas, faturamento, ocupacao, ROI, valorizacao ou dados de venda.
6. Nao recomende promessa de retorno garantido, renda garantida, valorizacao garantida ou disponibilidade absoluta.
7. Crie recomendacoes praticas para site, Instagram, TikTok, YouTube, Facebook/Meta Ads, Google Ads, WhatsApp/CRM, app e time de vendas.
8. Entregue um banco de ganchos, headlines, CTAs, objeções e respostas.
9. Inclua fontes com links.

Formato:
- Diagnostico do mercado.
- Glossario profundo.
- Matriz de modelos.
- Matriz de players brasileiros.
- Matriz de players internacionais.
- Personas.
- Objecoes.
- Estrategia por canal.
- Funil completo.
- Site ideal.
- Conteudo organico.
- Trafego pago.
- WhatsApp/CRM.
- App.
- Banco de copy.
- Compliance.
- Plano de 30 dias.
- Fontes.

Estilo: consultor senior, direto, especifico, sem frase generica, com foco em decisao e execucao.

Checklist final: a entrega diferencia o que e provado do que e inferido? Evita promessas ilegais? Explica o que vender, para quem, por qual canal e com qual prova? Esta pronta para virar campanha, site e treinamento de vendas?
\`\`\`

## 16. Fontes principais

${sources.map((s) => `- [${s[0]}](${s[1]}) - ${s[2]} (${s[3]})`).join("\n")}
`;

const mdPath = path.join(outDir, "Dossie_Imobiliario_Turistico_Premium.md");
fs.writeFileSync(mdPath, md, "utf8");

const errors = await workbook.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
  options: { useRegex: true, maxResults: 100 },
  summary: "formula error scan",
});

for (const sheetName of ["00 Leia Primeiro", "01 Modelos Negocio", "02 Players Brasil", "03 Players Internacionais", "09 Banco Copy"]) {
  const image = await workbook.render({ sheetName, autoCrop: "all", scale: 1, format: "png" });
  fs.writeFileSync(path.join(outDir, `${sheetName.replaceAll(" ", "_")}.png`), new Uint8Array(await image.arrayBuffer()));
}

const xlsxPath = path.join(outDir, "Base_Inteligencia_Imobiliario_Turistico_Premium.xlsx");
const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(xlsxPath);

console.log(errors.ndjson);
console.log(JSON.stringify({ outDir, xlsxPath, mdPath }, null, 2));
