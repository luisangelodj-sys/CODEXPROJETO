import fs from "node:fs/promises";
import path from "node:path";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const outputDir = path.resolve("outputs", "branding_resorts_multipropriedade");
await fs.mkdir(outputDir, { recursive: true });

const companies = [
  {
    empresa: "GAV Resorts",
    pais: "Brasil",
    tipo: "Multipropriedade + hotelaria + clube de viagens",
    ativos: "Porto de Galinhas, Pirenopolis, Salinopolis e outros empreendimentos",
    promessa: "Casa de ferias, ferias compartilhadas, destinos desejados, conforto e economia",
    branding: "Democratiza a ideia de casa de ferias com linguagem familiar e acessivel",
    voz: "Aspiracional popular-premium, familia, economia, pertencimento",
    publico: "Familias e compradores aspiracionais que querem ferias recorrentes sem comprar uma casa integral",
    provas: "FAQ de multipropriedade, depoimentos de multiproprietarios, portal do cliente, Select Club",
    canais: "Site, blog/GAV News, WhatsApp, Instagram, Facebook, portal do cliente",
    cta: "Garanta/adquira sua casa de ferias; falar com consultor; reservas de hotelaria",
    licao: "Explicar o modelo antes de vender; usar 'casa de ferias' como traducao simples de multipropriedade",
    risco: "Cuidado com economia/valorizacao: tratar como beneficio possivel, nunca garantia",
    evidencia: "COMPROVADO + INFERIDO",
    fonte: "https://www.gavresorts.com.br/"
  },
  {
    empresa: "Gramado Parks",
    pais: "Brasil",
    tipo: "Hospitalidade + entretenimento + gastronomia + multipropriedade",
    ativos: "Buona Vitta, Exclusive, Bella, Aquan, Namareh, Wyndham Gramado Termas, Snowland, Acquamotion, Yup Star",
    promessa: "Hospedagem com entretenimento de qualidade e novo padrao de lazer e ferias",
    branding: "Ecossistema de ferias: hotel + parque + gastronomia + atracoes",
    voz: "Familiar, sonhadora, plural, experiencial",
    publico: "Familias, turistas de Gramado e compradores de ferias que valorizam entretenimento integrado",
    provas: "Portfolio amplo de hoteis, parques, gastronomia e blog educativo",
    canais: "Site institucional, hotelaria, blog, YouTube, redes sociais, turismo/receptivo",
    cta: "Conhecer hoteis, atracoes, reservas e turismo",
    licao: "Construir valor vendendo destino completo, nao apenas unidade ou hospedagem",
    risco: "Pode ficar generico se nao explicar a diferenca entre hotelaria, clube e multipropriedade",
    evidencia: "COMPROVADO + INFERIDO",
    fonte: "https://gramadoparks.com/pt-br/"
  },
  {
    empresa: "WAM Experience / WAM Hoteis",
    pais: "Brasil",
    tipo: "Hotelaria, resorts, parques, pass e experiencias",
    ativos: "Caldas Novas, Aquiraz, Buzios, Porto Seguro; Privé Thermas, Dom Pedro Laguna, Ondas Praia Resort e outros",
    promessa: "Descontos para viajar com a familia; hospitalidade e experiencia",
    branding: "Volume, capilaridade e conveniencia de reserva em destinos populares",
    voz: "Promocional, familiar, direta, orientada a reserva",
    publico: "Familias buscando resort, parque, all inclusive e lazer com boa relacao de custo",
    provas: "Portfolio de 10 hoteis/resorts, blog, canais sociais e chamadas promocionais",
    canais: "Site, blog, Instagram, Facebook, TikTok, YouTube, WhatsApp, newsletter",
    cta: "Reserva agora, quero promos, fazer reserva",
    licao: "Promocao e urgencia funcionam melhor quando conectadas a destino, estrutura e facilidade",
    risco: "Excesso de desconto pode baixar percepcao premium; separar produto popular de alto padrao",
    evidencia: "COMPROVADO + INFERIDO",
    fonte: "https://wamhoteis.com.br/"
  },
  {
    empresa: "Aviva",
    pais: "Brasil",
    tipo: "Resorts, parques aquaticos, clube de ferias e residence clubs",
    ativos: "Rio Quente, Costa do Sauipe, Hot Park, InCasa, InCanto, Clube de Ferias Aviva",
    promessa: "Momentos unicos, conforto, exclusividade e felicidade como destino",
    branding: "Ecossistema nacional de resorts e parques com camada residencial/club",
    voz: "Premium familiar, emocional, natureza, experiencia",
    publico: "Familias e casais com desejo de resort consolidado e destino reconhecido",
    provas: "Marcas conhecidas, Hot Park, Rio Quente, Costa do Sauipe, residence clubs, canais de venda",
    canais: "Site, Instagram das marcas, WhatsApp, blog, newsletter, central de ajuda",
    cta: "Ver pacotes, conhecer residence clubs, falar com central",
    licao: "Usar marcas-ancora fortes para transferir confianca a produtos residenciais/clubes",
    risco: "Residence club precisa de explicacao didatica para nao parecer apenas hospedagem",
    evidencia: "COMPROVADO + INFERIDO",
    fonte: "https://www.aviva.com.br/"
  },
  {
    empresa: "Beach Park",
    pais: "Brasil",
    tipo: "Destino completo: parque, resorts, praia, vila, vacation club",
    ativos: "Aqua Park, Parque Arvorar, Resorts Beach Park, Vila Azul do Mar, Vacation Club, Beach Card",
    promessa: "Destino completo para toda a familia no Ceara; parque, praia e resorts",
    branding: "Entretenimento iconico + familia + recordes + destino completo",
    voz: "Divertida, vibrante, turistica, familiar",
    publico: "Familias, turistas de ferias, compradores de clube/recorrencia e agentes de viagem",
    provas: "Aqua Park, Guinness/recorde, TripAdvisor, quatro resorts, blog, portal do agente",
    canais: "Site, blog, Instagram, TikTok, YouTube, Facebook, Pinterest, portal do agente",
    cta: "Comprar ingresso, reservar pacote, conhecer vacation club",
    licao: "A venda fica forte quando o produto vira destino completo e nao hospedagem isolada",
    risco: "Para multipropriedade/clube, explicar regras e valor alem da diversao imediata",
    evidencia: "COMPROVADO + INFERIDO",
    fonte: "https://beachpark.com.br/"
  },
  {
    empresa: "Hot Beach",
    pais: "Brasil",
    tipo: "Parque aquatico + resorts + residence club",
    ativos: "Hot Beach Olimpia, Hot Beach Resort, Celebration, Raizes, Suites, Vila Guarani, Residence Club",
    promessa: "Experiencia completa com agua quentinha, resort, parque e casa de ferias anexa ao parque",
    branding: "Olimpia como destino de lazer termal, familiar e recorrente",
    voz: "Calorosa, ludica, familiar, relaxamento e diversao",
    publico: "Familias que viajam para Olimpia e compradores de casa de ferias vinculada ao parque",
    provas: "Acesso ilimitado ao parque, premios/selo, depoimentos, parque visitado, RCI Gold Crown em ativos",
    canais: "Site, redes sociais, YouTube, ingresso online, reservas, residence club",
    cta: "Reservar hospedagem, comprar ingresso, conhecer casa de ferias",
    licao: "Ancorar multipropriedade em uso concreto: parque + resort + cidade turistica",
    risco: "Evitar transformar depoimento de lazer em promessa de investimento",
    evidencia: "COMPROVADO + INFERIDO",
    fonte: "https://hotbeach.com.br/"
  },
  {
    empresa: "diRoma",
    pais: "Brasil",
    tipo: "Hoteis, parque aquatico e clube",
    ativos: "Caldas Novas: varios hoteis, diRoma Acqua Park, VIP Club",
    promessa: "Aguas termais, atividades recreativas, conforto, diversao, gastronomia e parcelamento",
    branding: "Tradicao de Caldas Novas com forte apelo promocional e familiar",
    voz: "Direta, promocional, popular-familiar",
    publico: "Familias de turismo domestico, visitantes de aguas termais e compradores de pacote",
    provas: "Ampla lista de hoteis, parque aquatico, canais oficiais, termos/direito de arrependimento",
    canais: "Site, WhatsApp, Instagram, Facebook, YouTube, newsletter",
    cta: "Reserve pelo site, conheca ja, inscreva-se para ofertas",
    licao: "Clareza de beneficios tangiveis vende bem: parque, criancas, parcelamento, aguas termais",
    risco: "Precisa elevar sofisticacao se o objetivo for alto padrao premium",
    evidencia: "COMPROVADO + INFERIDO",
    fonte: "https://diroma.com.br/"
  },
  {
    empresa: "Taua Resorts",
    pais: "Brasil",
    tipo: "Resorts familiares e hotelaria de lazer",
    ativos: "Taua Caete, Atibaia, Joao Pessoa, Alexania, Grande Hotel Termas de Araxa, Alegro",
    promessa: "Pacotes para familia, ferias gigantes, cultura do sorriso, infraestrutura completa",
    branding: "Resort familiar brasileiro com personagem, acolhimento e eventos",
    voz: "Familiar, acolhedora, colorida, promocional",
    publico: "Familias com criancas, eventos corporativos, ferias escolares",
    provas: "Destinos, tour virtual, pensao completa, aquapark, programacao para adultos e criancas",
    canais: "Site, motor de reserva, WhatsApp, CRM/newsletter, redes sociais",
    cta: "Quero reservar, reservar, fale com nossa equipe",
    licao: "Personagem, cultura e recorrencia de programacao ajudam a construir memoria de marca",
    risco: "Nao e modelo de multipropriedade principal; usar como benchmark de resort familiar",
    evidencia: "COMPROVADO + INFERIDO",
    fonte: "https://www.tauaresorts.com.br/"
  },
  {
    empresa: "Vila Gale",
    pais: "Portugal/Brasil",
    tipo: "Hotelaria e resorts internacionais com presenca no Brasil",
    ativos: "Rede de hoteis no Brasil e Portugal, resorts de praia e cidade",
    promessa: "Hospitalidade, lazer, destinos e padrao de rede internacional",
    branding: "Rede hoteleira ampla com confianca operacional e portfolio",
    voz: "Institucional, hoteleira, promocional",
    publico: "Turistas, familias, casais e corporativo que buscam rede reconhecida",
    provas: "Portfolio de hoteis no Brasil, motor de reserva, ofertas e estrutura de rede",
    canais: "Site, motor de reserva, redes sociais, CRM, campanhas de pacote",
    cta: "Reservar hotel, ver ofertas",
    licao: "Marca de rede ajuda a reduzir risco percebido em destino desconhecido",
    risco: "Menos aderente a venda de cota; mais util como benchmark de hospitalidade",
    evidencia: "COMPROVADO + INFERIDO",
    fonte: "https://www.vilagale.com/pt-br/hoteis-brasil/"
  },
  {
    empresa: "Grupo Wish",
    pais: "Brasil",
    tipo: "Hotelaria premium, resorts e experiencias",
    ativos: "Wish Foz, Wish Serrano, Wish Natal, Marupiara, IL Campanario, Jurere Beach Village, Exclusive Guest",
    promessa: "Reserve e viva experiencias inesqueciveis",
    branding: "Rede nacional com apelo premium, lazer e destinos desejados",
    voz: "Premium acessivel, hoteleira, experiencial",
    publico: "Familias, casais e viajantes de lazer premium",
    provas: "Portfolio de resorts e hoteis em destinos turisticos conhecidos",
    canais: "Site, reservas, campanhas de destino, redes sociais",
    cta: "Reservar, ver hoteis e experiencias",
    licao: "Curadoria de portfolio permite vender por destino e nao por produto isolado",
    risco: "Pouca educacao de modelo se usado para produtos imobiliarios complexos",
    evidencia: "COMPROVADO + INFERIDO",
    fonte: "https://www.grupowish.com/"
  },
  {
    empresa: "Bourbon Hospitalidade",
    pais: "Brasil",
    tipo: "Hotelaria, resorts e convencoes",
    ativos: "Rede Bourbon, Rio Hotel by Bourbon, resorts e hoteis urbanos",
    promessa: "Hospitalidade, reserva, lazer e estrutura para familia/eventos",
    branding: "Tradicao hoteleira brasileira e confianca operacional",
    voz: "Institucional, confiavel, hoteleira",
    publico: "Familias, corporativo, eventos e viajantes de rede",
    provas: "Rede consolidada, estrutura de reservas, hoteis e eventos",
    canais: "Site, reservas, redes, CRM, eventos",
    cta: "Reserve, conheca destinos",
    licao: "Operacao hoteleira madura comunica seguranca, nao apenas desejo",
    risco: "Benchmark mais de hotelaria do que de venda imobiliaria",
    evidencia: "COMPROVADO + INFERIDO",
    fonte: "https://www.bourbon.com.br/"
  },
  {
    empresa: "Fasano",
    pais: "Brasil/Internacional",
    tipo: "Hotelaria de luxo, gastronomia, villas e branded lifestyle",
    ativos: "Sao Paulo, Rio, Angra, Trancoso, Boa Vista, Punta del Este, Nova York e outros",
    promessa: "Tradicao, gastronomia, destino, experiencia e sofisticacao discreta",
    branding: "Luxo brasileiro com lastro cultural, gastronomia e selecao de destinos",
    voz: "Sofisticada, minimalista, tradicional, editorial",
    publico: "Alta renda, clientes de luxo, compradores de lifestyle e hospitalidade premium",
    provas: "Historia desde 1902 na gastronomia, hoteis no Brasil/Uruguai/EUA, restaurantes, JHSF",
    canais: "Site, reservas, newsletter, Instagram, Facebook, Vimeo, Spotify, Pinterest",
    cta: "Reserve sua estada, descubra destinos, experiencias",
    licao: "Luxo forte nao explica demais: usa curadoria, historia, silencio visual e prova social implicita",
    risco: "Se copiado por marca sem lastro, vira pose vazia; precisa de produto real",
    evidencia: "COMPROVADO + INFERIDO",
    fonte: "https://fasano.com.br/"
  },
  {
    empresa: "JHSF",
    pais: "Brasil",
    tipo: "Incorporacao de alto padrao + shoppings + hospitality + clubs + aeroportos",
    ativos: "Cidade Jardim, Fazenda Boa Vista, Boa Vista Village, Fasano, Catarina, clubs, retail",
    promessa: "Tradition and Modernity; negocios unicos para clientes especiais",
    branding: "Ecossistema de alta renda: morar, consumir, viajar, socializar e circular",
    voz: "Corporativa, aspiracional, alta renda, institucional",
    publico: "Ultra-alta renda, investidores patrimoniais, clientes de lifestyle e consumo de luxo",
    provas: "Desde 1972, negocios em real estate, hospitality, private member clubs, business aviation e retail",
    canais: "Site institucional, RI, propriedades a venda, Instagram, LinkedIn",
    cta: "Conhecer negocios, propriedades a venda",
    licao: "O alto padrao ganha forca quando o produto vira ecossistema de vida, nao so metragem",
    risco: "Comunicar alto padrao sem ecossistema real fica vazio; precisa de provas estruturais",
    evidencia: "COMPROVADO + INFERIDO",
    fonte: "https://jhsf.com.br/"
  },
  {
    empresa: "Rosewood Residences",
    pais: "Internacional",
    tipo: "Branded residences, serviced apartments e villa rentals",
    ativos: "Residences for sale, serviced apartments e rentals em destinos globais",
    promessa: "Homes of remarkable character; enriched living every day",
    branding: "Conexao profunda com lugar, design, autenticidade e servico Rosewood",
    voz: "Editorial, cultural, sofisticada, sensorial",
    publico: "Ultra-alta renda global, compradores de branded residences e estadias longas",
    provas: "Arquitetos/designers, servico dedicado, beneficios globais, colecao internacional",
    canais: "Site, inquiry form, Instagram/Facebook/WeChat, mailing list",
    cta: "Submit an inquiry, discover the collection",
    licao: "Branded residence premium vende identidade de lugar e servico, nao so luxo material",
    risco: "Evitar copiar linguagem poetica sem entregar arquitetura/servico a altura",
    evidencia: "COMPROVADO + INFERIDO",
    fonte: "https://www.rosewoodhotels.com/en/residences"
  },
  {
    empresa: "Four Seasons Private Residences",
    pais: "Internacional",
    tipo: "Private residences, residence clubs e rentals",
    ativos: "Portfolio global de residences em cidades e resorts",
    promessa: "Come home to Four Seasons; your home, the centre of our world",
    branding: "Casa + servico invisivel + paz de espirito + destinos desejados",
    voz: "Acolhedora, ultra-premium, pessoal, servico antecipatorio",
    publico: "Ultra-alta renda global que quer propriedade com hotel service e gestao",
    provas: "Property management, world-class amenities, global portfolio, equipe dedicada",
    canais: "Site, inquiry, newsletter, social, portfolio map",
    cta: "Find a Private Residence, subscribe, view portfolio",
    licao: "O argumento maximo e paz de espirito: morar/ter casa sem operar a casa sozinho",
    risco: "Nao vender como retorno financeiro; vender como vida, servico e tranquilidade",
    evidencia: "COMPROVADO + INFERIDO",
    fonte: "https://www.fourseasons.com/residences/"
  },
  {
    empresa: "The Ritz-Carlton Residences",
    pais: "Internacional",
    tipo: "Whole-ownership branded residences",
    ativos: "Residences em destinos urbanos e resort globais",
    promessa: "Live Here, Always; legendary attention to detail",
    branding: "Padrao ouro de servico aplicado a vida cotidiana",
    voz: "Luxo classico, servico, pertencimento, legado",
    publico: "Alta e ultra-alta renda que quer propriedade integral com marca hoteleira",
    provas: "24-hour concierge, in-residence dining, valet, pools, golf, spa, Marriott Bonvoy select benefits",
    canais: "Site, Marriott Residences, inquiry, social, Bonvoy/ecossistema Marriott",
    cta: "Explore residences, live in [destination]",
    licao: "Branding forte traduz servico em vida diaria: 'every wish anticipated'",
    risco: "Manter clareza: whole ownership e nao promessa de renda ou clube generico",
    evidencia: "COMPROVADO + INFERIDO",
    fonte: "https://www.ritzcarlton.com/en/residences/"
  },
  {
    empresa: "Six Senses Private Residences",
    pais: "Internacional",
    tipo: "Private residences vinculadas a resorts/wellness",
    ativos: "Residences em Courchevel, Fiji, Ibiza, Dubai, London, Con Dao, Zil Pasyon e outros",
    promessa: "Indulgent living as nature intended",
    branding: "Wellness, natureza, sustentabilidade e luxo de baixa ostentacao",
    voz: "Sensorial, natural, consciente, calma",
    publico: "Ultra-alta renda que busca bem-estar, natureza, privacidade e sustentabilidade",
    provas: "Acesso a amenities Six Senses, materiais renovaveis/locais, eficiencia energetica e hidrica",
    canais: "Site, enquiry form, hoteis/resorts, experiencias e wellness",
    cta: "Enquiry form, book/reserve",
    licao: "Luxo contemporaneo pode ser ancorado em bem-estar e natureza, nao em ostentacao",
    risco: "Sustentabilidade precisa ser especifica; evitar greenwashing",
    evidencia: "COMPROVADO + INFERIDO",
    fonte: "https://www.sixsenses.com/en/residences/"
  },
  {
    empresa: "One&Only Private Homes",
    pais: "Internacional",
    tipo: "Ultra-luxury private homes/villas",
    ativos: "Private homes para compra e estadia em destinos One&Only",
    promessa: "Pinnacle of ultra-luxury design, comfort and taste",
    branding: "Ultra-luxo de design, resort e exclusividade",
    voz: "Minimalista, aspiracional, escassa, premium",
    publico: "Ultra-alta renda global que deseja villa privada com marca de resort de luxo",
    provas: "Marca One&Only, destinos globais, private homes para compra e stay",
    canais: "Site, enquire, brand hospitality channels",
    cta: "Enquire",
    licao: "Em ultra-luxo, menos informacao publica pode reforcar exclusividade, mas exige atendimento consultivo forte",
    risco: "Pouca informacao aberta dificulta educacao do lead; exige follow-up premium",
    evidencia: "COMPROVADO + INFERIDO",
    fonte: "https://www.oneandonlyresorts.com/private-homes"
  },
  {
    empresa: "Marriott Vacation Club",
    pais: "Internacional",
    tipo: "Vacation ownership / timeshare / pontos",
    ativos: "Villa resorts e city properties em destinos premium globais",
    promessa: "Create your best vacation life; villa vacations; Marriott quality",
    branding: "Ferias flexiveis com padrao Marriott e acomodacoes tipo casa",
    voz: "Educativa, familiar, confiavel, compliance forte",
    publico: "Familias e viajantes recorrentes que querem flexibilidade de destinos e acomodacoes maiores",
    provas: "Owner stories, pontos/club, resorts, Abound, Marriott standards, disclaimers legais de timeshare",
    canais: "Site, blog Vacation Life, request information, social, owner portal",
    cta: "Request information, let's connect today",
    licao: "O benchmark de compliance: declara claramente solicitacao de venda de timeshare e offering terms",
    risco: "No Brasil, qualquer promessa financeira deve ser revisada juridicamente antes de anunciar",
    evidencia: "COMPROVADO + INFERIDO",
    fonte: "https://www.marriottvacationclubs.com/ownership/vacation-club-brands/marriott-vacation-club.html"
  },
  {
    empresa: "Hilton Grand Vacations",
    pais: "Internacional",
    tipo: "Vacation membership / timeshare / pontos",
    ativos: "200+ resorts, Home Resort, semanas, pontos, pacotes de apresentacao",
    promessa: "Reimagine the way you travel; feel at home; travel your way",
    branding: "Flexibilidade de viagem com lastro Hilton e comunidade de membros",
    voz: "Comercial, educativa, orientada a oferta, prova social",
    publico: "Familias e viajantes recorrentes acostumados com Hilton e interesse em membership",
    provas: "720k+ members, depoimentos, Hilton quality, detalhes legais, faixa de preco e disclosures",
    canais: "Site, ofertas, Getaway Guide, social, apresentacao virtual, direct sales",
    cta: "Learn more, see how it works, schedule presentation",
    licao: "Ofertas de entrada baratas aquecem para apresentacao, mas precisam de disclosure forte",
    risco: "Promocoes de pacote precisam deixar termos, elegibilidade e natureza de solicitacao claros",
    evidencia: "COMPROVADO + INFERIDO",
    fonte: "https://www.hiltongrandvacations.com/en/discover-hilton-grand-vacations"
  },
  {
    empresa: "Disney Vacation Club",
    pais: "Estados Unidos",
    tipo: "Vacation ownership / timeshare por pontos",
    ativos: "Resorts Disney Vacation Club em Walt Disney World, Disneyland, Aulani, Hilton Head e outros",
    promessa: "Ferias Disney recorrentes com pontos e acesso a villas/resorts",
    branding: "Pertencimento emocional a Disney, ritual familiar e previsibilidade de ferias",
    voz: "Magica, familiar, nostalgica, membership",
    publico: "Familias e fas da Disney que viajam repetidamente para resorts Disney",
    provas: "Marca Disney, resorts tematicos, sistema de pontos, comunidade de membros",
    canais: "Site oficial, conteudos Disney, tours, atendimento, comunidade e canais proprietarios",
    cta: "Conhecer membership e conversar com representante",
    licao: "Marca emocional forte transforma recorrencia em identidade familiar",
    risco: "No Brasil, nao usar 'magia' emocional para esconder custo anual, regras de pontos e contrato",
    evidencia: "INFERIDO + PRECISA VALIDACAO",
    fonte: "https://disneyvacationclub.disney.go.com/"
  }
];

const archetypes = [
  ["Arquetipo", "Empresas exemplo", "Promessa central", "Quando usar", "Risco"],
  ["Casa de ferias acessivel", "GAV, Hot Beach Residence Club", "Ter um lugar de ferias recorrente sem comprar/manter sozinho", "Multipropriedade/residence club aspiracional", "Prometer economia ou valorizacao como certeza"],
  ["Destino completo", "Beach Park, Aviva, Gramado Parks, Hot Beach", "Parque + resort + gastronomia + familia + memoria", "Quando existe ecossistema real de experiencias", "Virar panfleto de lazer sem explicar produto"],
  ["Resort familiar brasileiro", "Taua, diRoma, WAM", "Ferias com estrutura, criancas, programacao e facilidade", "Hotelaria e campanhas promocionais", "Ficar popular demais para oferta premium"],
  ["Hospitalidade premium", "Wish, Bourbon, Vila Gale", "Confianca de rede, reserva, destino e padrao de servico", "Benchmark para operacao e credibilidade", "Nao resolve sozinho a complexidade juridica da cota"],
  ["Luxury lifestyle ecosystem", "JHSF, Fasano", "Morar, consumir, viajar e pertencer a um ecossistema de alto padrao", "Imobiliario de luxo e branded lifestyle", "Luxo sem lastro vira texto vazio"],
  ["Branded residence global", "Four Seasons, Ritz-Carlton, Rosewood, Six Senses, One&Only", "Propriedade + servico hoteleiro + paz de espirito + destino", "Alto padrao e ultra-luxo", "Confundir uso/lifestyle com promessa financeira"],
  ["Vacation ownership compliance", "Marriott Vacation Club, Hilton Grand Vacations, Disney Vacation Club", "Flexibilidade de ferias e membership", "Clubes, pontos, timeshare, apresentacao consultiva", "Falta de disclosure e claims agressivos"]
];

const channels = [
  ["Canal", "Papel no funil", "Conteudo ideal", "O que aprender com os players", "CTA recomendado", "Metrica"],
  ["Instagram", "Desejo e prova visual", "Reels de destino, stories, tour, depoimento, bastidor", "Beach Park/Hot Beach vendem sensacao; GAV/Gramado educam com casa de ferias", "Conhecer proposta / falar no WhatsApp", "Salvos, respostas, cliques, leads"],
  ["TikTok", "Descoberta e quebra de objecao", "Ganchos de dor: diaria cara, casa de praia da trabalho, cota mal explicada", "WAM/Beach Park usam lazer; usar tambem educacao simples", "Ver explicacao completa / simular", "Retencao, compartilhamento, visita ao perfil"],
  ["YouTube", "Educacao e autoridade", "Videos longos de comparacao, contrato, custos, destino, tour completo", "Marriott/Hilton educam com membership, pontos e stories", "Assistir apresentacao / baixar guia", "Watch time, leads assistidos, remarketing"],
  ["Facebook", "Maturacao e remarketing", "Videos explicativos, depoimentos, oferta de apresentacao, lead ads com cuidado", "Familia e publico maduro ainda respondem a clareza", "Falar com consultor", "CPL qualificado, taxa de contato"],
  ["Google Search", "Intencao ativa", "Paginas por modelo: multipropriedade, residence club, resort em destino, duvidas", "SEO precisa responder pergunta dificil", "Entender se faz sentido", "Conversao por termo, qualidade do lead"],
  ["Site/Landing Page", "Confianca e conversao", "Como funciona, para quem e, custos, regras, provas, perguntas dificeis", "GAV explica; Marriott/Hilton usam disclosure", "Solicitar simulacao/apresentacao", "CVR, tempo na pagina, scroll, WhatsApp"],
  ["WhatsApp/CRM", "Qualificacao e venda consultiva", "Roteiro por perfil, duvidas, agendamento, follow-up documentado", "Venda nao fecha no criativo; fecha no processo", "Agendar apresentacao", "Tempo resposta, show rate, fechamento"],
  ["Aplicativo/Portal", "Retencao e prova pos-venda", "Reserva, calendario, documentos, beneficios, suporte, indicacao", "GAV tem portal; clubs globais tratam owner/member area como ativo", "Usar beneficios / indicar amigo", "Uso ativo, NPS, indicacoes"]
];

const lessons = [
  ["Principio", "O que significa na pratica", "Empresas que mostram isso", "Aplicacao para nossa base"],
  ["Nao vender cota; vender traducao mental", "O cliente entende casa de ferias, ferias planejadas, destino e familia; 'cota' vem depois", "GAV, Hot Beach, Marriott", "Criar conteudo que explica antes de pedir lead"],
  ["Produto complexo exige educacao", "Quanto maior o risco percebido, mais perguntas a pagina precisa responder", "GAV, Marriott, Hilton", "Landing page com FAQ duro: taxas, regras, revenda, uso, contrato"],
  ["Destino completo aumenta valor percebido", "Parque, gastronomia, praia, vila, eventos e kids club tornam a oferta menos abstrata", "Beach Park, Aviva, Gramado Parks, Hot Beach", "Mostrar ecossistema e rotina real de uso"],
  ["Luxo precisa de lastro", "Arquitetura, servico, gestao, acesso, curadoria e privacidade valem mais que adjetivo", "Fasano, JHSF, Rosewood, Four Seasons", "Nao usar 'alto padrao' sem prova concreta"],
  ["Compliance tambem e branding", "Disclosures claros reduzem medo e qualificam melhor", "Marriott, Hilton, diRoma", "Criar checklist de claims seguros e proibidos"],
  ["Pos-venda sustenta reputacao", "Portal/app, beneficios e suporte reduzem arrependimento", "GAV, Marriott, Hilton", "Aplicativo deve provar uso e nao apenas vender mais"],
  ["Oferta de entrada pode aquecer a venda", "Pacotes de estadia/apresentacao criam experiencia antes do contrato", "Hilton, Beach Park, WAM", "Usar visita/tour/estadia como camada de conversao, com termos claros"]
];

const sources = companies.map((c) => [c.empresa, c.fonte, c.evidencia, "Fonte principal usada para leitura de posicionamento, produto, canais e linguagem."]);

const workbook = Workbook.create();

function writeSheet(name, data, tableName) {
  const sheet = workbook.worksheets.add(name);
  sheet.showGridLines = false;
  const rows = data.length;
  const cols = data[0].length;
  const range = sheet.getRangeByIndexes(0, 0, rows, cols);
  range.values = data;
  sheet.freezePanes.freezeRows(1);
  const header = sheet.getRangeByIndexes(0, 0, 1, cols);
  header.format = {
    fill: "#1F2937",
    font: { bold: true, color: "#FFFFFF" },
    wrapText: true
  };
  range.format = { wrapText: true, borders: { preset: "all", style: "thin", color: "#D9D9D9" } };
  header.format = { fill: "#1F2937", font: { bold: true, color: "#FFFFFF" }, wrapText: true };
  try {
    const lastCol = String.fromCharCode("A".charCodeAt(0) + cols - 1);
    const t = sheet.tables.add(`A1:${lastCol}${rows}`, true, tableName);
    t.style = "TableStyleMedium2";
    t.showFilterButton = true;
  } catch (e) {
    // Table styling is helpful but not essential.
  }
  for (let c = 0; c < cols; c++) {
    const width = c === 0 ? 180 : c === cols - 1 ? 340 : 240;
    sheet.getRangeByIndexes(0, c, rows, 1).format.columnWidthPx = width;
  }
  sheet.getRangeByIndexes(1, 0, Math.max(rows - 1, 1), cols).format.rowHeightPx = 78;
  return sheet;
}

const summary = [
  ["Campo", "Conteudo"],
  ["Objetivo", "Base de dados de branding, posicionamento, oferta, canais e aprendizados dos principais players de resort, multipropriedade, vacation ownership e branded residences."],
  ["Tese central", "Esse mercado nao deve ser estudado como uma categoria unica. Ele mistura hotelaria, turismo, imobiliario, uso recorrente, clube, luxo, servico e venda consultiva."],
  ["Como usar", "Filtre a aba Base Empresas por tipo, publico, promessa e licao. Use Branding Arquétipos para escolher posicionamento. Use Canais Funil para transformar pesquisa em execucao."],
  ["Leitura critica", "Players populares vendem familia, desconto e ferias. Players de multipropriedade traduzem o produto como casa de ferias. Players globais premium vendem servico, paz de espirito, pertencimento e identidade de lugar."],
  ["Regra anti-chute", "O que vem de site oficial esta marcado como COMPROVADO. A leitura de estrategia, publico e risco esta marcada como INFERIDO quando nao ha dado direto."],
  ["Proxima camada", "Auditar redes sociais e Meta Ads Library/Google Ads/YouTube de cada player para coletar criativos reais, frequencia de postagens, claims e ofertas ativas."]
];

writeSheet("00 Leia Primeiro", summary, "ResumoBase");
writeSheet("01 Base Empresas", [
  ["Empresa", "Pais", "Tipo", "Ativos/Produtos", "Promessa observada", "Leitura de branding", "Tom de voz", "Publico provavel", "Provas usadas", "Canais observados", "CTA dominante", "Licao para copiar", "Risco/atencao", "Evidencia", "Fonte"],
  ...companies.map(c => [c.empresa, c.pais, c.tipo, c.ativos, c.promessa, c.branding, c.voz, c.publico, c.provas, c.canais, c.cta, c.licao, c.risco, c.evidencia, c.fonte])
], "BaseEmpresas");
writeSheet("02 Branding Arquetipos", archetypes, "BrandingArquetipos");
writeSheet("03 Canais Funil", channels, "CanaisFunil");
writeSheet("04 Licoes Aplicaveis", lessons, "LicoesAplicaveis");
writeSheet("05 Fontes", [["Empresa", "URL", "Nivel de evidencia", "Uso na base"], ...sources], "Fontes");

const scoring = [
  ["Empresa", "Forca de marca", "Educacao do mercado", "Ecossistema de experiencia", "Premium/Luxo", "Compliance aparente", "Aprendizado principal"],
  ["GAV Resorts", 4, 5, 4, 3, 3, "Traduzir multipropriedade em casa de ferias"],
  ["Gramado Parks", 4, 3, 5, 3, 3, "Unir hospedagem e entretenimento"],
  ["WAM", 3, 2, 4, 2, 2, "Escala promocional e capilaridade"],
  ["Aviva", 5, 3, 5, 4, 3, "Marca-ancora forte transfere confianca"],
  ["Beach Park", 5, 3, 5, 3, 3, "Destino completo vende melhor que resort isolado"],
  ["Hot Beach", 4, 3, 5, 3, 3, "Produto residencial ancorado em uso real"],
  ["diRoma", 3, 2, 4, 2, 4, "Beneficios tangiveis e clareza promocional"],
  ["Taua", 4, 2, 4, 3, 3, "Cultura familiar/personagem cria memoria"],
  ["Fasano", 5, 2, 4, 5, 4, "Luxo se prova por historia e curadoria"],
  ["JHSF", 5, 2, 5, 5, 4, "Ecossistema de alta renda"],
  ["Four Seasons", 5, 4, 5, 5, 5, "Paz de espirito e gestao da casa"],
  ["Ritz-Carlton", 5, 4, 5, 5, 5, "Servico lendario aplicado a vida diaria"],
  ["Rosewood", 5, 3, 5, 5, 4, "Identidade de lugar e design"],
  ["Six Senses", 5, 3, 5, 5, 4, "Wellness e natureza como luxo"],
  ["Marriott Vacation Club", 5, 5, 4, 4, 5, "Compliance e educacao em vacation ownership"],
  ["Hilton Grand Vacations", 5, 5, 4, 4, 5, "Oferta de entrada + apresentacao + disclosure"],
  ["Disney Vacation Club", 5, 4, 5, 4, 4, "Marca emocional e ritual familiar"]
];
writeSheet("06 Score Comparativo", scoring, "ScoreComparativo");

const errors = await workbook.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
  options: { useRegex: true, maxResults: 100 },
  summary: "formula error scan"
});
console.log(errors.ndjson);

for (const sheetName of ["00 Leia Primeiro", "01 Base Empresas", "02 Branding Arquetipos", "03 Canais Funil", "06 Score Comparativo"]) {
  const preview = await workbook.render({ sheetName, autoCrop: "all", scale: 1, format: "png" });
  await fs.writeFile(path.join(outputDir, `${sheetName.replaceAll(" ", "_")}.png`), new Uint8Array(await preview.arrayBuffer()));
}

const output = await SpreadsheetFile.exportXlsx(workbook);
const xlsxPath = path.join(outputDir, "base_branding_resorts_multipropriedade.xlsx");
await output.save(xlsxPath);
console.log(`saved=${xlsxPath}`);
