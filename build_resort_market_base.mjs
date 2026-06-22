import fs from "node:fs/promises";
import path from "node:path";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const consultationDate = "2026-06-17";
const outputDir = path.resolve("outputs", "resort_market_research");
await fs.mkdir(outputDir, { recursive: true });

const regionByState = {
  AC: "Norte", AL: "Nordeste", AP: "Norte", AM: "Norte", BA: "Nordeste",
  CE: "Nordeste", DF: "Centro-Oeste", ES: "Sudeste", GO: "Centro-Oeste",
  MA: "Nordeste", MT: "Centro-Oeste", MS: "Centro-Oeste", MG: "Sudeste",
  PA: "Norte", PB: "Nordeste", PR: "Sul", PE: "Nordeste", PI: "Nordeste",
  RJ: "Sudeste", RN: "Nordeste", RS: "Sul", RO: "Norte", RR: "Norte",
  SC: "Sul", SP: "Sudeste", SE: "Nordeste", TO: "Norte",
};

const source = {
  gav: "https://www.gavresorts.com.br/",
  gramado: "https://gramadoparks.com/pt-br/",
  wam: "https://wamhoteis.com.br/",
  aviva: "https://www.aviva.com.br/",
  avivaResidence: "https://www.aviva.com.br/residence-clubs",
  beachPark: "https://beachpark.com.br/",
  hotBeach: "https://hotbeach.com.br/",
  hotBeachResidence: "https://hotbeachresidenceclub.com.br/",
  diRoma: "https://diroma.com.br/",
  taua: "https://www.tauaresorts.com.br/",
  vilaGale: "https://www.vilagale.com/pt-br/hoteis-brasil/",
  wish: "https://www.grupowish.com/",
  bourbon: "https://www.bourbon.com.br/",
  cana: "https://canabravaresort.com.br/",
  salinas: "https://www.salinas.com.br/pt",
  pratagy: "https://pratagy.com.br/",
  nannai: "https://www.nannai.com.br/",
  clubMed: "https://www.clubmed.com.br/",
  jurema: "https://juremaaguasquentes.com.br/",
  fasano: "https://fasano.com.br/",
  rosewood: "https://www.rosewoodhotels.com/en/sao-paulo",
  wyndhamGramado: "https://www.wyndhamhotels.com/pt-br/wyndham/gramado-rs-brazil/wyndham-gramado-termas-resort-and-spa/overview",
  briefing: "Briefing inicial fornecido pelo usuario; exige validacao oficial individual.",
};

const masterHeaders = [
  "Nome do empreendimento",
  "Grupo / incorporadora / operadora",
  "Cidade",
  "Estado",
  "Regiao do Brasil",
  "Categoria do produto",
  "Modelo comercial provavel",
  "Publico-alvo principal",
  "Status atual",
  "Ticket ou preco inicial",
  "Tipo de venda",
  "Apelo principal de venda",
  "Principais diferenciais",
  "Estrutura de lazer",
  "Presenca de parque / lago / praia / termas / natureza / entretenimento",
  "Bandeira hoteleira ou marca associada",
  "Canal oficial encontrado",
  "Fonte consultada",
  "Data da consulta",
  "Grau de confianca",
  "Observacoes criticas",
];

const rows = [];
function add({
  name, group, city, state, category, model, audience, status, price = "Nao identificado",
  saleType, appeal, differentials, leisure, experience, brand = "", official, src, confidence, notes,
}) {
  rows.push([
    name, group, city, state, regionByState[state] || "Nao identificado", category, model, audience,
    status, price, saleType, appeal, differentials, leisure, experience, brand, official || src, src,
    consultationDate, confidence, notes,
  ]);
}

function addMany(base, names) {
  for (const item of names) add({ ...base, ...item });
}

addMany({
  group: "GAV Resorts",
  state: "PE",
  category: "Multipropriedade + hotelaria de lazer",
  model: "Casa de ferias compartilhada com uso hoteleiro tradicional",
  audience: "Familias de classe media alta que desejam ferias recorrentes no litoral",
  status: "Operacao hoteleira / produto imobiliario confirmado no grupo",
  saleType: "Fracao / hospedagem",
  appeal: "Casa de ferias em destinos desejados com estrutura de resort",
  differentials: "Marca especializada em multipropriedade; narrativa de ferias compartilhadas; Select Club",
  leisure: "Piscinas, gastronomia, lazer familiar e recreacao, conforme produto",
  experience: "Praia / entretenimento familiar",
  brand: "GAV Resorts",
  src: source.gav,
  confidence: "Alto",
  notes: "A pagina oficial confirma atuacao em turismo, hotelaria e multipropriedade; validar condicoes comerciais por empreendimento.",
}, [
  { name: "Porto 2 Life Resort", city: "Porto de Galinhas", official: "https://www.gavresorts.com.br/" },
  { name: "Porto Alto Resort", city: "Porto de Galinhas", official: "https://www.portoaltoresort.com.br/" },
]);

addMany({
  group: "GAV Resorts", state: "PA", city: "Salinopolis",
  category: "Multipropriedade + resort familiar", model: "Casa de ferias compartilhada com hospedagem",
  audience: "Familias do Norte/Centro-Oeste e clientes recorrentes de ferias",
  status: "Operacao hoteleira / produto imobiliario confirmado no grupo", saleType: "Fracao / hospedagem",
  appeal: "Resort de praia com uso recorrente e conveniencia de propriedade compartilhada",
  differentials: "Ecossistema GAV em Salinopolis; depoimentos de multiproprietarios na pagina oficial",
  leisure: "Piscinas, brinquedos aquaticos, gastronomia e areas infantis",
  experience: "Praia / entretenimento familiar", brand: "GAV Resorts", src: source.gav, confidence: "Alto",
  notes: "Presenca oficial no bloco de hospedagens prontas da GAV.",
}, [
  { name: "Premium GAV Resorts", official: "https://www.premiumgavresorts.com.br/" },
  { name: "Exclusive GAV Resorts", official: "https://www.exclusivegavresorts.com.br/" },
  { name: "Park GAV Resorts", official: "https://www.parkgavresorts.com.br/" },
]);

addMany({
  group: "GAV Resorts", state: "GO", city: "Pirenopolis",
  category: "Residence / multipropriedade", model: "Casa de ferias compartilhada",
  audience: "Familias de Goias/DF em busca de segunda residencia de lazer",
  status: "Operacao hoteleira / produto imobiliario confirmado no grupo", saleType: "Fracao / hospedagem",
  appeal: "Refugio de lazer proximo a Brasilia e Goiania",
  differentials: "Destino de natureza e escapada curta; produto ligado ao portfolio GAV",
  leisure: "Piscinas, lazer familiar e apoio de hotelaria", experience: "Natureza / escapada regional",
  brand: "GAV Resorts", src: source.gav, confidence: "Alto",
  notes: "Canal proprio citado na pagina oficial GAV.",
}, [
  { name: "Pyreneus Residence", official: "https://www.pyreneus.com.br/" },
]);

addMany({
  group: "GAV Resorts", city: "Gramado", state: "RS",
  category: "Multipropriedade / resort em destino turistico",
  model: "Possivel casa de ferias compartilhada; confirmar pagina individual",
  audience: "Familias que desejam ferias recorrentes na Serra Gaucha",
  status: "Lancamento / status nao confirmado individualmente",
  saleType: "Possivel fracao - confirmar", appeal: "Serra Gaucha como destino de ferias de alto apelo emocional",
  differentials: "Gramado, clima europeu, gastronomia e entretenimento",
  leisure: "A validar por empreendimento", experience: "Serra / entretenimento / inverno",
  brand: "GAV Resorts", src: `${source.gav}; ${source.briefing}`, confidence: "Medio",
  notes: "Gran Haus aparece em noticia da GAV em 2026; demais nomes vieram do briefing inicial e exigem validacao por pagina individual.",
}, [
  { name: "Gran Haus Resort" },
  { name: "Gran Garden Resort", confidence: "Baixo", status: "Status nao confirmado" },
  { name: "Gran Valley Resort", confidence: "Baixo", status: "Status nao confirmado" },
]);

addMany({
  group: "GAV Resorts", state: "CE", city: "Jericoacoara",
  category: "Possivel multipropriedade / resort turistico",
  model: "Possivel casa de ferias compartilhada - confirmar",
  audience: "Turistas de praia premium e familias em ferias",
  status: "Status nao confirmado", saleType: "Possivel fracao - confirmar",
  appeal: "Destino aspiracional no litoral cearense",
  differentials: "Jericoacoara, dunas, lagoas e experiencia de natureza",
  leisure: "A validar", experience: "Praia / dunas / lagoa / natureza",
  brand: "GAV Resorts", src: source.briefing, confidence: "Baixo",
  notes: "Incluido por radar do briefing; precisa de confirmacao oficial ativa.",
}, [
  { name: "Jeriquia Dunas Resort" },
  { name: "Jeriquia Lagoa Resort" },
]);

addMany({
  group: "GAV Resorts", category: "Possivel multipropriedade / resort turistico",
  model: "Possivel casa de ferias compartilhada - confirmar", audience: "Familias e turistas de praia",
  status: "Status nao confirmado", saleType: "Possivel fracao - confirmar",
  appeal: "Litoral nordestino como destino de ferias", differentials: "Localizacao turistica forte",
  leisure: "A validar", experience: "Praia / natureza", brand: "GAV Resorts", src: source.briefing,
  confidence: "Baixo", notes: "Incluido por briefing; fonte oficial individual nao validada nesta passada.",
}, [
  { name: "Oikos Maragogi Resort", city: "Maragogi", state: "AL" },
  { name: "Areya Barra Resort", city: "Barra de Sao Miguel", state: "AL" },
  { name: "Beach GAV Resorts", city: "Salinopolis", state: "PA" },
]);

addMany({
  group: "Gramado Parks", city: "Gramado", state: "RS",
  category: "Hospitalidade + entretenimento turistico",
  model: "Hotelaria, turismo e possivel venda relacionada a ferias; confirmar modalidade por empreendimento",
  audience: "Familias que visitam Gramado e compradores de experiencia de ferias",
  status: "Operacao hoteleira / empreendimento listado no canal oficial",
  saleType: "Hospedagem / possivel produto imobiliario conforme unidade",
  appeal: "Hospedagem integrada ao ecossistema de entretenimento de Gramado",
  differentials: "Snowland, Acquamotion, gastronomia e parques do grupo",
  leisure: "Piscinas, spa, recreacao e acesso ao ecossistema de parques",
  experience: "Neve / termas / entretenimento familiar", brand: "Gramado Parks / Wyndham em um ativo",
  src: source.gramado, confidence: "Alto",
  notes: "Pagina oficial lista hospitalidade e entretenimento; validar contrato comercial de multipropriedade em cada produto.",
}, [
  { name: "Buona Vitta" },
  { name: "Exclusive" },
  { name: "Bella" },
  { name: "Wyndham Gramado Termas Resort Spa", src: `${source.gramado}; ${source.wyndhamGramado}`, brand: "Wyndham / Gramado Parks" },
  { name: "Hydros", src: source.briefing, confidence: "Baixo", status: "Status nao confirmado" },
]);

addMany({
  group: "Gramado Parks", category: "Parque / entretenimento integrado a turismo",
  model: "Venda de ingresso, experiencia e hospedagem associada",
  audience: "Familias, turistas e grupos",
  status: "Operacao / listado no canal oficial", saleType: "Ingresso / pacote / hospedagem associada",
  appeal: "Atracao ancora que aumenta recorrencia e desejo pelo destino",
  differentials: "Ativos proprietarios de entretenimento; capacidade de gerar fluxo para hotelaria",
  leisure: "Parques tematicos e experiencias", brand: "Gramado Parks", src: source.gramado,
  confidence: "Alto", notes: "Mais relevante para venda como prova de ecossistema do que como imovel em si.",
}, [
  { name: "Snowland", city: "Gramado", state: "RS", experience: "Neve / entretenimento" },
  { name: "Acquamotion", city: "Gramado", state: "RS", experience: "Termas / parque aquatico indoor" },
  { name: "Acquaventura", city: "Foz do Iguacu", state: "PR", experience: "Parque aquatico / entretenimento" },
  { name: "Yup Star Rio", city: "Rio de Janeiro", state: "RJ", experience: "Roda-gigante / experiencia urbana" },
  { name: "Yup Star Foz", city: "Foz do Iguacu", state: "PR", experience: "Roda-gigante / turismo" },
  { name: "Aquan", city: "Foz do Iguacu", state: "PR", category: "Hospitalidade em destino turistico", model: "Hotelaria / possivel produto imobiliario - confirmar", experience: "Cataratas / lazer familiar" },
  { name: "Namareh", city: "Praia dos Carneiros", state: "PE", category: "Hospitalidade em destino de praia", model: "Hotelaria / possivel produto imobiliario - confirmar", experience: "Praia / resort" },
]);

addMany({
  group: "WAM / WAM Hoteis", category: "Hotelaria de lazer / resort familiar",
  model: "Hospedagem, pacote, operacao hoteleira; multipropriedade a confirmar por ativo",
  audience: "Familias de ferias em destinos de alto volume",
  status: "Operacao hoteleira listada no canal oficial", saleType: "Diaria / pacote / possivel cota conforme ativo",
  appeal: "Resorts com lazer forte, praia, lago ou termas", differentials: "Portfolio multi-destino; WAM Parks e WAM Pass",
  leisure: "Piscinas, parques proximos, recreacao, gastronomia", brand: "WAM Hoteis",
  src: source.wam, confidence: "Alto",
  notes: "A pagina oficial consultada confirma hoteis e resorts; nao confirma em todos a venda imobiliaria.",
}, [
  { name: "Prive Thermas", city: "Caldas Novas", state: "GO", experience: "Termas / parques aquáticos" },
  { name: "Dom Pedro Laguna", city: "Aquiraz", state: "CE", experience: "Praia / alto padrao" },
  { name: "Buzios Beach Resort", city: "Armacao dos Buzios", state: "RJ", experience: "Praia / resort" },
  { name: "Boulevard Suite Hotel", city: "Caldas Novas", state: "GO", experience: "Termas / centro turistico" },
  { name: "Encontro das Aguas Thermas Resort", city: "Caldas Novas", state: "GO", experience: "Termas / lago" },
  { name: "Praias do Lago Eco Resort", city: "Caldas Novas", state: "GO", experience: "Lago Corumba / termas" },
  { name: "Ilhas do Lago Eco Resort", city: "Caldas Novas", state: "GO", experience: "Lago Corumba / termas" },
  { name: "Resort do Lago", city: "Caldas Novas", state: "GO", experience: "Lago Corumba / lazer" },
  { name: "Riviera Park Hotel", city: "Caldas Novas", state: "GO", experience: "Termas / Water Park / Clube Prive" },
  { name: "Ondas Praia Resort", city: "Porto Seguro", state: "BA", experience: "Praia / all inclusive" },
]);

addMany({
  group: "WAM / Enjoy", category: "Possivel multipropriedade / resort em destino de parques",
  model: "Possivel cota ou fracao; confirmar fonte oficial atual",
  audience: "Familias que visitam Olimpia",
  status: "Status nao confirmado nesta passada", saleType: "Possivel cota / diaria",
  appeal: "Olimpia como polo de parque aquatico e ferias familiares",
  differentials: "Destino com grande volume de turismo domestico",
  leisure: "Piscinas e parque aquático na regiao", experience: "Parque aquatico / termas",
  brand: "Enjoy / WAM", src: source.briefing, confidence: "Baixo",
  notes: "Incluido pelo briefing inicial; exige validacao oficial atual e checagem reputacional.",
}, [
  { name: "Enjoy Solar das Aguas Park Resort", city: "Olimpia", state: "SP" },
  { name: "Enjoy Olimpia Park Resort", city: "Olimpia", state: "SP" },
]);

addMany({
  group: "Aviva", category: "Resort + parque aquatico + clube/residence",
  model: "Pacotes, diarias, clube de ferias e residence club",
  audience: "Familias, casais e compradores de experiencia recorrente",
  status: "Operacao / produto listado em canal oficial", saleType: "Diaria / pacote / clube / residence",
  appeal: "Ecossistema de paraisos com parque aquatico e resorts consolidados",
  differentials: "Rio Quente, Costa do Sauipe, Hot Park e residencia de alto padrao",
  leisure: "Parques aquaticos, termas, praia, entretenimento e hotelaria",
  brand: "Aviva / Rio Quente / Costa do Sauipe / Hot Park", src: source.aviva,
  confidence: "Alto", notes: "Aviva separa hoteis, residence clubs e clube de ferias no canal oficial.",
}, [
  { name: "Rio Quente Resorts", city: "Rio Quente", state: "GO", experience: "Termas / Hot Park" },
  { name: "Hot Park Rio Quente", city: "Rio Quente", state: "GO", experience: "Parque aquatico / termas", category: "Parque aquatico integrado a resort", saleType: "Ingresso / pacote" },
  { name: "Costa do Sauipe", city: "Mata de Sao Joao", state: "BA", experience: "Praia / complexo de resorts" },
  { name: "Hot Park Costa do Sauipe", city: "Mata de Sao Joao", state: "BA", experience: "Parque aquatico / praia", category: "Parque aquatico integrado a resort", saleType: "Ingresso / pacote" },
  { name: "Clube de Ferias Aviva", city: "Multi-destino", state: "GO", category: "Clube de ferias", saleType: "Clube / pontos / assinatura", experience: "Rede de destinos" },
  { name: "InCasa Residence Club", city: "Rio Quente", state: "GO", category: "Residence club de alto padrao", saleType: "Unidade / residence club - confirmar", src: source.avivaResidence, experience: "Cerrado / termas / privacidade" },
  { name: "InCanto Residence Club", city: "Mata de Sao Joao", state: "BA", category: "Residence club de alto padrao", saleType: "Unidade / residence club - confirmar", src: source.avivaResidence, experience: "Praia / clube privado" },
]);

addMany({
  group: "Beach Park", city: "Aquiraz", state: "CE",
  category: "Parque + resort + vacation club",
  model: "Pacotes, diarias, clube e beneficios",
  audience: "Familias com criancas, turistas de Fortaleza e clientes recorrentes",
  status: "Operacao / produto listado no canal oficial", saleType: "Diaria / pacote / clube / ingresso",
  appeal: "Destino completo de praia, parque e resorts",
  differentials: "Aqua Park, Parque Arvorar, resorts no entorno e Vacation Club",
  leisure: "Parque aquatico, praia, vila gastronomica, recreacao e hotelaria",
  experience: "Praia / parque aquatico / entretenimento familiar",
  brand: "Beach Park", src: source.beachPark, confidence: "Alto",
  notes: "Beach Park e uma escola forte para venda por ecossistema: parque ancora + hospedagem + recorrencia.",
}, [
  { name: "Acqua Beach Park Resort" },
  { name: "Wellness Beach Park Resort" },
  { name: "Oceani Beach Park Resort" },
  { name: "Suites Beach Park Resort" },
  { name: "Aqua Park", category: "Parque aquatico", saleType: "Ingresso / pacote" },
  { name: "Parque Arvorar", category: "Parque de entretenimento", saleType: "Ingresso / pacote", experience: "Natureza / entretenimento" },
  { name: "Vila Azul do Mar", category: "Entretenimento e gastronomia", saleType: "Consumo / experiencia" },
  { name: "Beach Park Vacation Club", category: "Vacation club", saleType: "Clube / pontos - confirmar" },
  { name: "Beach Card", category: "Clube de beneficios", saleType: "Clube / assinatura - confirmar" },
  { name: "Ohana Beach Park Resort", status: "Status nao confirmado nesta passada", src: source.briefing, confidence: "Baixo" },
]);

addMany({
  group: "Hot Beach / Grupo Ferrasa", city: "Olimpia", state: "SP",
  category: "Parque + resort + multipropriedade",
  model: "Hospedagem, ingresso, residence club e multipropriedade",
  audience: "Familias que buscam parque aquatico e casa de ferias em Olimpia",
  status: "Operacao / produto imobiliario confirmado em canal oficial",
  saleType: "Diaria / pacote / multipropriedade / ingresso",
  appeal: "Casa de ferias anexa a parque aquatico em destino consolidado",
  differentials: "Acesso ao Hot Beach, Vila Guarani, RCI e clube de beneficios",
  leisure: "Parque aquatico, piscinas, vila gastronomica, recreacao, rooftop e brinquedoteca",
  experience: "Parque aquatico / aguas quentes / entretenimento familiar",
  brand: "Hot Beach", src: `${source.hotBeach}; ${source.hotBeachResidence}`, confidence: "Alto",
  notes: "Hot Beach Residence Club explica multipropriedade e lista Hot Beach Suites/You/Raizes.",
}, [
  { name: "Hot Beach Resort", category: "Resort familiar com parque" },
  { name: "Hot Beach Celebration", category: "Resort familiar com parque" },
  { name: "Hot Beach Raizes / Thermas Park", category: "Resort familiar / multipropriedade" },
  { name: "Hot Beach Suites", category: "Multipropriedade / resort", saleType: "Multipropriedade / diaria" },
  { name: "Hot Beach You", category: "Multipropriedade em lancamento", status: "Em desenvolvimento / primeira fase prevista 2027", saleType: "Multipropriedade" },
  { name: "Parque Aquatico Hot Beach Olimpia", category: "Parque aquatico", saleType: "Ingresso / pacote" },
  { name: "Vila Guarani", category: "Entretenimento e gastronomia", saleType: "Consumo / experiencia" },
]);

addMany({
  group: "diRoma", city: "Caldas Novas", state: "GO",
  category: "Hotelaria de lazer + parque aquatico",
  model: "Diarias, pacotes e clube de beneficios; imobiliario nao confirmado por ativo",
  audience: "Familias e turistas de aguas termais",
  status: "Operacao hoteleira listada no canal oficial", saleType: "Diaria / pacote / clube",
  appeal: "Tradicao em Caldas Novas com varios hoteis e acesso a parque",
  differentials: "Grande volume de unidades, marca regional forte e Acqua Park",
  leisure: "Piscinas, parque aquatico, aguas termais, recreacao",
  experience: "Termas / parque aquatico", brand: "diRoma", src: source.diRoma, confidence: "Alto",
  notes: "Fonte oficial lista os hoteis e VIP Club; venda imobiliaria deve ser confirmada caso a caso.",
}, [
  "Hotel Roma", "Imperio Romano", "L'acqua diRoma I", "L'acqua diRoma II", "L'acqua diRoma III",
  "L'acqua diRoma IV", "L'acqua diRoma V", "Piazza diRoma", "Spazzio diRoma", "Thermas diRoma",
  "Villas diRoma", "diRoma Exclusive", "diRoma Fiori", "diRoma Resort", "diRoma Acqua Park", "VIP Club diRoma",
].map((name) => ({ name, category: name.includes("Park") ? "Parque aquatico" : (name.includes("VIP") ? "Clube de beneficios" : "Hotelaria de lazer") })));

addMany({
  group: "Taua Resorts", category: "Resort familiar premium / hotelaria pura",
  model: "Diarias e pacotes de pensao completa; produto imobiliario nao identificado",
  audience: "Familias de classe media alta; eventos corporativos e lazer infantil",
  status: "Operacao hoteleira listada no canal oficial", saleType: "Diaria / pacote",
  appeal: "Ferias em familia com infraestrutura grande e programacao infantil",
  differentials: "Marca familiar forte, pensao completa, aquapark indoor em Atibaia, destinos em expansao",
  leisure: "Piscinas, recreacao, restaurantes, eventos, quadras, spa",
  experience: "Natureza / piscinas / entretenimento familiar", brand: "Taua Resorts", src: source.taua,
  confidence: "Alto", notes: "Excelente benchmark de promessa familiar, mas nao deve ser tratado como multipropriedade sem evidencia.",
}, [
  { name: "Taua Resort Caete", city: "Caete", state: "MG" },
  { name: "Taua Resort Atibaia", city: "Atibaia", state: "SP", experience: "Aquapark indoor / familia" },
  { name: "Taua Resort Joao Pessoa", city: "Joao Pessoa", state: "PB", experience: "Praia / familia" },
  { name: "Taua Resort Alexania", city: "Alexania", state: "GO" },
  { name: "Grande Hotel Termas de Araxa", city: "Araxa", state: "MG", experience: "Termas / patrimonio historico" },
  { name: "Alegro Hotel", city: "Jarinu", state: "SP", category: "Hotel de lazer / eventos" },
]);

addMany({
  group: "Vila Gale Brasil", category: "Rede hoteleira / resort de lazer",
  model: "Diarias, pacotes e programa MyVilaGale; imobiliario nao identificado",
  audience: "Familias, casais, turistas de praia e publico internacional",
  status: "Operacao hoteleira listada no canal oficial", saleType: "Diaria / pacote / clube de fidelidade",
  appeal: "Marca hoteleira portuguesa com presenca nacional em destinos de praia e cultura",
  differentials: "Rede internacional, all inclusive em alguns resorts, escala de marca",
  leisure: "Piscinas, praia, gastronomia, spas, eventos, conforme unidade",
  experience: "Praia / natureza / cidade / cultura", brand: "Vila Gale", src: source.vilaGale,
  confidence: "Alto", notes: "Relevante como benchmark de hotelaria e marca; nao como multipropriedade salvo evidencia especifica.",
}, [
  { name: "Vila Gale Alagoas", city: "Barra de Santo Antonio", state: "AL", experience: "Praia / resort" },
  { name: "Vila Gale Mares", city: "Camacari", state: "BA", experience: "Praia / all inclusive" },
  { name: "Vila Gale Salvador", city: "Salvador", state: "BA", category: "Hotel urbano de lazer" },
  { name: "Vila Gale Collection Sunset Cumbuco", city: "Caucaia", state: "CE", experience: "Praia / collection" },
  { name: "Vila Gale Cumbuco", city: "Caucaia", state: "CE", experience: "Praia / resort" },
  { name: "Vila Gale Fortaleza", city: "Fortaleza", state: "CE", experience: "Praia / urbano" },
  { name: "Vila Gale Collection Ouro Preto", city: "Ouro Preto", state: "MG", experience: "Cultura / patrimonio" },
  { name: "Vila Gale Collection Amazonia", city: "Belem", state: "PA", experience: "Amazonia / cultura" },
  { name: "Vila Gale Cabo", city: "Cabo de Santo Agostinho", state: "PE", experience: "Praia / resort" },
  { name: "Vila Gale Angra", city: "Angra dos Reis", state: "RJ", experience: "Praia / natureza" },
  { name: "Vila Gale Rio de Janeiro", city: "Rio de Janeiro", state: "RJ", category: "Hotel urbano de lazer" },
  { name: "Vila Gale Touros", city: "Touros", state: "RN", experience: "Praia / resort" },
  { name: "Vila Gale Paulista", city: "Sao Paulo", state: "SP", category: "Hotel urbano" },
]);

addMany({
  group: "Grupo Wish", category: "Hotelaria premium / resort familiar",
  model: "Diarias, pacotes e clube de ferias Exclusive Guest",
  audience: "Familias, casais, eventos e turistas premium",
  status: "Operacao listada no canal oficial", saleType: "Diaria / pacote / clube",
  appeal: "Rede brasileira premium com resorts em destinos turisticos fortes",
  differentials: "Portfolio Wish, My Wish e Exclusive Guest",
  leisure: "Piscinas, gastronomia, eventos, lazer infantil, conforme unidade",
  experience: "Praia / natureza / serra / cataratas", brand: "Grupo Wish",
  src: source.wish, confidence: "Alto", notes: "Benchmark forte de hotelaria premium e clube; imobiliario nao confirmado.",
}, [
  { name: "Wish Foz do Iguacu Resort", city: "Foz do Iguacu", state: "PR", experience: "Cataratas / natureza" },
  { name: "Wish Serrano Resort", city: "Gramado", state: "RS", experience: "Serra / inverno" },
  { name: "Wish Natal Resort", city: "Natal", state: "RN", experience: "Praia / resort" },
  { name: "Marupiara All Inclusive Resort by Wish", city: "Ipojuca", state: "PE", experience: "Praia / all inclusive" },
  { name: "IL Campanario Villaggio Resort", city: "Florianopolis", state: "SC", experience: "Jurere / alto padrao" },
  { name: "Jurere Beach Village", city: "Florianopolis", state: "SC", experience: "Praia / segunda residencia" },
  { name: "Exclusive Guest", city: "Multi-destino", state: "SP", category: "Clube de ferias", saleType: "Clube / beneficios" },
]);

addMany({
  group: "Bourbon Hospitalidade", category: "Resort familiar / hotelaria premium",
  model: "Diarias e pacotes; produto imobiliario nao identificado",
  audience: "Familias, eventos e lazer regional premium",
  status: "Operacao hoteleira listada no canal oficial", saleType: "Diaria / pacote",
  appeal: "Tradicao e estrutura familiar com personagens e gastronomia",
  differentials: "Mais de 60 anos, +20 hoteis e resorts, Turma da Monica nos resorts",
  leisure: "Piscinas, recreacao, gastronomia, eventos",
  experience: "Natureza / cataratas / familia", brand: "Bourbon", src: source.bourbon,
  confidence: "Alto", notes: "Hotelaria pura para benchmark de experiencia e programacao familiar.",
}, [
  { name: "Bourbon Resort Atibaia", city: "Atibaia", state: "SP" },
  { name: "Bourbon Cataratas do Iguacu Thermas Eco Resort", city: "Foz do Iguacu", state: "PR", experience: "Cataratas / termas / familia" },
]);

addMany({
  group: "Operadores independentes / redes premium", category: "Resort familiar premium / all inclusive / luxo",
  model: "Diarias e pacotes; produto imobiliario geralmente nao identificado",
  audience: "Familias, casais, alta renda, lua de mel e turismo premium",
  status: "Operacao hoteleira confirmada em fonte oficial", saleType: "Diaria / pacote",
  appeal: "Experiencia completa, destino aspiracional e baixa friccao operacional",
  differentials: "All inclusive, praia, natureza, termas ou luxo discreto conforme marca",
  leisure: "Piscinas, gastronomia, recreacao, spa e experiencias locais",
  brand: "Independente / marca propria", confidence: "Alto",
  notes: "Entram como benchmarks comerciais; nao tratar como multipropriedade sem prova.",
}, [
  { name: "Cana Brava All Inclusive Resort", group: "Cana Brava", city: "Ilheus", state: "BA", src: source.cana, experience: "Praia / all inclusive 24h" },
  { name: "Salinas Maragogi All Inclusive Resort", group: "Amarante / Salinas", city: "Maragogi", state: "AL", src: source.salinas, experience: "Praia / all inclusive" },
  { name: "Salinas Maceio All Inclusive Resort", group: "Amarante / Salinas", city: "Maceio", state: "AL", src: source.salinas, experience: "Praia / all inclusive" },
  { name: "Pratagy Beach All Inclusive", group: "Pratagy", city: "Maceio", state: "AL", src: source.pratagy, experience: "Praia / acqua park" },
  { name: "NANNAI Muro Alto", group: "NANNAI", city: "Ipojuca", state: "PE", src: source.nannai, experience: "Praia / luxo / bangalos" },
  { name: "NANNAI Noronha", group: "NANNAI", city: "Fernando de Noronha", state: "PE", src: source.nannai, category: "Pousada boutique de luxo" },
  { name: "NANNAI Milagres", group: "NANNAI", city: "Sao Miguel dos Milagres", state: "AL", src: source.nannai, category: "Hotelaria boutique de luxo" },
  { name: "Club Med Rio das Pedras", group: "Club Med", city: "Mangaratiba", state: "RJ", src: source.clubMed, experience: "Praia / mata atlantica / premium all inclusive" },
  { name: "Club Med Lake Paradise", group: "Club Med", city: "Mogi das Cruzes", state: "SP", src: source.clubMed, experience: "Lago / premium all inclusive" },
  { name: "Club Med Trancoso", group: "Club Med", city: "Trancoso", state: "BA", src: source.clubMed, experience: "Praia / premium all inclusive" },
  { name: "Lagos de Jurema Termas Resort", group: "Jurema Aguas Quentes", city: "Iretama", state: "PR", src: source.jurema, experience: "Termas / familia" },
  { name: "Jardins de Jurema Convention & Termas Resort", group: "Jurema Aguas Quentes", city: "Iretama", state: "PR", src: source.jurema, experience: "Termas / alto padrao" },
  { name: "Jurema Vacation Club", group: "Jurema Aguas Quentes", city: "Iretama", state: "PR", category: "Vacation club", model: "Pontos de ferias", saleType: "Clube / pontos", src: source.jurema, experience: "Termas / rede de hospedagens" },
  { name: "Fasano Angra dos Reis", group: "Fasano / JHSF", city: "Angra dos Reis", state: "RJ", src: source.fasano, category: "Hotelaria de luxo / complexo imobiliario de alto padrao", audience: "Alta renda", experience: "Praia / marina / luxo" },
  { name: "Fasano Trancoso", group: "Fasano / JHSF", city: "Trancoso", state: "BA", src: source.fasano, category: "Hotelaria de luxo", audience: "Alta renda" },
  { name: "Fasano Boa Vista", group: "Fasano / JHSF", city: "Porto Feliz", state: "SP", src: source.fasano, category: "Hotelaria de luxo integrada a complexo residencial", audience: "Alta renda", experience: "Campo / golfe / segunda residencia" },
  { name: "Boa Vista Surf Lodge", group: "Fasano / JHSF", city: "Porto Feliz", state: "SP", src: source.fasano, category: "Hotelaria de lazer integrada a complexo de alto padrao", audience: "Alta renda", experience: "Piscina de ondas / surf / campo" },
  { name: "Rosewood Sao Paulo", group: "Rosewood / Cidade Matarazzo", city: "Sao Paulo", state: "SP", src: source.rosewood, category: "Hotelaria de luxo urbana / branded hospitality", audience: "Alta renda / luxo internacional", experience: "Cidade / arquitetura / gastronomia" },
]);

add({
  name: "Wyndham Corumba Lake Resort", group: "Wyndham / operador a confirmar", city: "Corumba de Goias / Lago Corumba IV - confirmar",
  state: "GO", category: "Radar comparativo solicitado", model: "Nao confirmado", audience: "A confirmar",
  status: "Fonte oficial nao localizada nesta passada", price: "Nao identificado", saleType: "Nao identificado",
  appeal: "Potencial apelo de lago e proximidade com Brasilia/Goiania, mas nao validado",
  differentials: "A confirmar", leisure: "A confirmar", experience: "Lago / natureza - confirmar",
  brand: "Wyndham - confirmar", official: "Nao localizado", src: "Buscas por 'Wyndham Corumba Lake Resort' sem resultado confiavel nesta consulta",
  confidence: "Baixo", notes: "Nao usar em comparacao comercial sem validacao documental. Prioridade de proxima pesquisa.",
});

const playerHeaders = [
  "Player", "Tipo dominante", "Forca de marca", "Clareza da oferta", "Apelo emocional",
  "Estrutura de lazer", "Estrategia de vendas", "Presenca digital", "Potencial de escala",
  "Inspiracao para vender alto padrao", "Nota media", "Leitura comercial",
];
const players = [
  ["GAV Resorts", "Multipropriedade + hotelaria", 8, 9, 8, 8, 9, 8, 9, 8, "", "Melhor escola para vender casa de ferias compartilhada com linguagem simples e desejo familiar."],
  ["Gramado Parks", "Ecossistema de entretenimento + hospedagem", 8, 7, 8, 9, 8, 8, 8, 7, "", "Excelente para entender parque/atracao como gerador de demanda para hospedagem."],
  ["WAM / Enjoy", "Resorts de alto volume", 7, 7, 7, 8, 8, 7, 8, 6, "", "Bom para estudar escala, destinos populares e oferta de lazer; exige cuidado reputacional."],
  ["Aviva", "Resort + parque + residence club", 9, 8, 9, 10, 8, 8, 8, 8, "", "Benchmark de ecossistema completo: Rio Quente, Hot Park, Sauipe, clube e residences."],
  ["Beach Park", "Destino completo parque + resorts", 9, 9, 9, 10, 8, 9, 9, 8, "", "Um dos melhores exemplos de venda por destino, familia e memoria afetiva."],
  ["Hot Beach", "Parque + multipropriedade", 8, 8, 8, 9, 9, 8, 8, 7, "", "Forte para aprender venda de multipropriedade ancorada em parque aquatico."],
  ["diRoma", "Hotelaria regional + parque", 7, 7, 7, 8, 7, 7, 7, 5, "", "Forte em volume e recorrencia em Caldas; menos sofisticado em narrativa premium."],
  ["Taua Resorts", "Resort familiar premium", 8, 8, 9, 9, 7, 8, 8, 8, "", "Excelente benchmark de promessa familiar sem depender de venda imobiliaria."],
  ["Vila Gale", "Rede hoteleira internacional", 8, 8, 7, 8, 7, 8, 8, 7, "", "Bom para estudar marca, rede e consistencia de produto; menos util para multipropriedade."],
  ["Grupo Wish", "Hotelaria premium + clube", 8, 8, 8, 8, 7, 8, 7, 8, "", "Bom para scripts de hotelaria premium e clube de ferias."],
  ["Club Med", "Premium all inclusive global", 9, 9, 9, 9, 8, 9, 8, 9, "", "A melhor escola para vender all inclusive premium como liberdade e baixa preocupacao."],
  ["Fasano / JHSF", "Luxo + hospitality + real estate", 10, 7, 8, 8, 7, 8, 7, 10, "", "Referencia para alto padrao: marca, status, design e escassez sem promessa popular."],
];

const angleHeaders = ["Angulo comercial", "Quando usar", "Prova que precisa aparecer", "Risco se mal usado", "Frase segura para vendas"];
const angles = [
  ["Ferias em familia", "Lead com filhos e desejo de descanso recorrente", "Estrutura infantil, seguranca, programacao e facilidade", "Virar frase generica de familia feliz", "Um lugar pensado para a familia descansar sem transformar a viagem em logistica."],
  ["Casa de ferias", "Multipropriedade e residence club", "Direito de uso, regras, custos e periodos", "Prometer propriedade plena quando e fracao", "Voce acessa uma casa de ferias com estrutura de resort, sem assumir sozinho o custo de manter tudo o ano inteiro."],
  ["Patrimonio", "Cliente racional ou investidor", "Matricula, contrato, regime juridico e historico do destino", "Prometer valorizacao garantida", "Pode fazer sentido como decisao patrimonial, desde que o cliente entenda custos, regras de uso e liquidez."],
  ["Uso inteligente do dinheiro", "Cliente que compara diaria vs compra", "Simulacao de uso real e custo anual", "Forcar economia sem demonstrar conta", "A conversa boa nao e pagar menos; e pagar por um modelo que combina com sua frequencia de ferias."],
  ["Experiencia premium", "Alta renda e casais", "Design, gastronomia, servico, privacidade e localizacao", "Ficar abstrato demais", "O diferencial esta no conjunto: localizacao, servico, conforto e menos preocupacao na viagem."],
  ["Resort como estilo de vida", "Cliente emocional e aspiracional", "Fotos reais, rotina de uso, depoimentos e clube", "Parecer status vazio", "Nao e so hospedagem. E criar uma rotina de pausa, lazer e encontro em um lugar previsivel."],
  ["Lazer para filhos", "Familias com criancas", "Brinquedoteca, monitoria, parque, seguranca", "Prometer cuidado sem prova operacional", "Enquanto as criancas tem programacao, os adultos conseguem descansar de verdade."],
  ["Marca internacional", "Branded residences e hotelaria premium", "Bandeira, padrao de servico e governanca", "Vender marca como garantia de retorno", "A marca reduz incerteza de experiencia, mas nao elimina risco comercial ou financeiro."],
  ["Valor do destino", "Destinos consolidados ou em expansao", "Fluxo turistico, acesso, atracoes e sazonalidade", "Prometer valorizacao certa", "Destinos com demanda recorrente tendem a sustentar melhor o uso, mas cada projeto precisa ser analisado."],
  ["Facilidade de parcelamento", "Lead sensivel a entrada", "Conditicoes reais e CET", "Ocultar custo total", "A parcela ajuda a entrar, mas a decisao precisa caber no custo total do contrato."],
  ["Rede de resorts / intercambio", "Vacation club e RCI", "Regras de pontos, disponibilidade e taxas", "Vender liberdade total sem restricao", "A rede aumenta possibilidades, desde que o cliente entenda antecedencia, disponibilidade e regras."],
  ["Baixa preocupacao operacional", "Alta renda e familias ocupadas", "Servicos, manutencao, governanca e reservas", "Ignorar taxas e regras", "O valor esta em aproveitar sem administrar uma segunda casa sozinho."],
];

const commercialHeaders = ["Player", "Como vende", "Promessa central", "Desejo ativado", "Gatilhos comerciais", "O que aprender", "O que evitar"];
const commercial = [
  ["GAV Resorts", "Multipropriedade com linguagem de casa de ferias, familia e economia", "Ter ferias recorrentes em destinos desejados sem comprar um imovel sozinho", "Pertencimento, previsibilidade e acesso", "Destino, familia, clube, depoimentos, consultor", "Simplificar o modelo e vender uso antes de investimento", "Prometer rentabilidade, valorizacao ou liquidez sem documento"],
  ["Beach Park", "Destino completo: parque, praia, resorts, vila e clube", "Uma experiencia familiar completa no mesmo lugar", "Memoria afetiva com filhos", "Parque ancora, recorrencia, facilidade, seguranca", "Usar atracao como prova concreta de valor", "Abrir com preco antes de construir desejo"],
  ["Aviva", "Ecossistema de parques, resorts, residence clubs e clube", "Pertencer a paraisos com estrutura consolidada", "Exclusividade com conforto", "Marca, natureza, parque, residence, clube", "Separar bem diaria, clube e residence", "Misturar clube de ferias com propriedade sem explicar contrato"],
  ["Hot Beach", "Casa de ferias anexa a parque aquatico e com RCI", "Ferias garantidas em Olimpia com diversao para todos", "Seguranca e entretenimento previsivel", "Parque, RCI, beneficios, fases do empreendimento", "Vender a logica do uso anual e intercambio", "Falar em investimento sem classificar risco"],
  ["Taua Resorts", "Hotelaria familiar com programacao e pensao completa", "Ferias familiares sem preocupacao", "Descanso dos pais e diversao dos filhos", "Programacao, comida, distancia da capital, eventos", "Copy de familia muito clara e operacional", "Tratar como multipropriedade sem evidencia"],
  ["Club Med", "All inclusive premium global", "Liberdade: tudo resolvido antes da viagem", "Status leve, conforto, esporte, familia", "Premium all inclusive, criancas incluidas, esportes", "Vender ausencia de friccao", "Vender luxo como ostentacao vazia"],
  ["Fasano / JHSF", "Marca de luxo integrada a destinos e complexos de alto padrao", "Pertencer a um estilo de vida raro e bem curado", "Status, privacidade, design e escassez", "Marca, localizacao, arquitetura, gastronomia", "Aprender linguagem de alto padrao sem exagero", "Popularizar demais a promessa ou falar como panfleto"],
];

const learningHeaders = ["Modulo", "O que estudar", "Aplicacao pratica em vendas"];
const learning = [
  ["Produto", "Diferenca entre diaria, clube, vacation club, multipropriedade, residence club, condo-hotel e branded residence", "Nunca apresentar cota como se fosse unidade plena; explicar uso, custos, regras e governanca."],
  ["Mercado", "Destinos consolidados, parques ancora, sazonalidade, acesso aereo/rodoviario e concorrencia", "Argumentar por demanda real do destino, nao por frase de oportunidade."],
  ["Publico", "Familia com filhos, casal, investidor, alta renda, lead frio/morno/quente", "Adaptar abordagem: familia compra previsibilidade; alta renda compra curadoria; investidor exige risco claro."],
  ["Desejo", "Ferias recorrentes, menos logistica, pertencimento, status, memoria familiar", "Vender a vida que o produto facilita, com prova concreta."],
  ["Dor", "Diarias caras, viagem improvisada, manutencao de segunda casa, falta de tempo", "Mostrar o problema operacional que o modelo resolve."],
  ["Objecoes", "Uso limitado, taxa anual, liquidez, disponibilidade, confianca, contrato", "Responder com transparencia e material documental."],
  ["Script de abordagem", "Abrir por destino + uso, nao por preco", "Ex.: 'Voce costuma viajar em familia todo ano ou decide tudo em cima da hora?'"],
  ["Script de qualificacao", "Frequencia de viagem, filhos, destinos preferidos, orcamento, perfil patrimonial", "Separar curioso de comprador real antes de apresentacao longa."],
  ["Script de apresentacao", "Problema -> modelo -> destino -> estrutura -> regras -> proximo passo", "Evitar comecar por planta/tabela se o cliente ainda nao comprou a ideia."],
  ["WhatsApp", "Mensagem curta com pergunta de contexto e prova visual", "Usar audio curto, video do empreendimento e chamada leve para entender perfil."],
  ["Showroom", "Experiencia guiada: destino, unidade, lazer, contrato, simulacao de uso", "Fazer o cliente imaginar a rotina de ferias, depois tratar numeros."],
  ["Familia", "Seguranca, criancas, monitoria, comida, facilidade", "Argumento central: descanso dos pais e diversao dos filhos."],
  ["Investidor", "Contrato, custos, liquidez, demanda, riscos e proibicoes de promessa", "Nao falar em rentabilidade certa; falar em tese e validacao documental."],
  ["Casal", "Privacidade, gastronomia, spa, escapadas curtas", "Mostrar experiencia de reconexao, nao so estrutura."],
  ["Alta renda", "Privacidade, marca, curadoria, acesso, design, baixa friccao", "Menos urgencia artificial; mais criterio e exclusividade real."],
  ["Lead frio", "Educar sobre destino e problema", "CTA leve: entender perfil ou receber material."],
  ["Lead morno", "Comparar modelos e remover objecoes", "Mostrar simulacao de uso anual e regras."],
  ["Lead quente", "Contrato, agenda, condicoes, visita e fechamento", "Reduzir risco percebido e organizar decisao familiar."],
  ["Cuidados juridicos", "CVM, CDC, Lei da Multipropriedade, promessas de rendimento, publicidade", "Nunca usar 'garantido', 'sem risco', 'renda certa', 'valorizacao certa'."],
  ["LGPD / WhatsApp / Meta Ads", "Consentimento, opt-out, atributos sensiveis e politica de promessas", "Evitar abordagem invasiva e claims personalizados negativos."],
];

const workbook = Workbook.create();

function writeSheet(name, headers, data, widths = []) {
  const sheet = workbook.worksheets.add(name);
  sheet.showGridLines = false;
  const matrix = [headers, ...data];
  sheet.getRangeByIndexes(0, 0, matrix.length, headers.length).values = matrix;
  const headerRange = sheet.getRangeByIndexes(0, 0, 1, headers.length);
  headerRange.format = { fill: "#17324D", font: { bold: true, color: "#FFFFFF" }, wrapText: true };
  const bodyRange = sheet.getRangeByIndexes(1, 0, data.length, headers.length);
  bodyRange.format = { wrapText: true, borders: { preset: "all", style: "thin", color: "#D9E2EC" } };
  sheet.getRangeByIndexes(0, 0, matrix.length, headers.length).format.borders = { preset: "all", style: "thin", color: "#D9E2EC" };
  sheet.freezePanes.freezeRows(1);
  for (let i = 0; i < headers.length; i++) {
    sheet.getRangeByIndexes(0, i, matrix.length, 1).format.columnWidthPx = widths[i] || 160;
  }
  try {
    sheet.tables.add(`A1:${colLetter(headers.length)}${matrix.length}`, true, name.replace(/[^A-Za-z0-9]/g, "").slice(0, 20) || "Table");
  } catch {
    // Formatting remains usable if a table name/range is rejected by Excel.
  }
  return sheet;
}

function colLetter(n) {
  let s = "";
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - m) / 26);
  }
  return s;
}

const readme = [
  ["Base", "Pesquisa de mercado imobiliario turistico - resorts, multipropriedade e hotelaria premium"],
  ["Como usar", "Filtre a Lista Master por categoria, modelo comercial, status e grau de confianca."],
  ["Regra de leitura", "Alto = fonte oficial consultada; Medio = fonte oficial do grupo + detalhe individual a validar; Baixo = briefing/radar sem confirmacao suficiente."],
  ["Limite", "Esta e a lista mais ampla encontrada com base nas fontes consultadas nesta passada, nao uma lista de todos os empreendimentos do Brasil."],
  ["Data da consulta", consultationDate],
  ["Proxima validacao", "CVM quando houver discurso de investimento; Reclame Aqui para risco reputacional; Cadastur para operacao hoteleira; matricula/incorporacao para venda imobiliaria."],
];
writeSheet("00 LEIA PRIMEIRO", ["Campo", "Conteudo"], readme, [190, 900]);

writeSheet("01 Lista Master", masterHeaders, rows, [
  220, 210, 150, 70, 120, 220, 250, 240, 180, 140, 170, 260, 280, 260, 290, 180, 250, 320, 110, 120, 360,
]);

for (const p of players) {
  const nums = p.slice(2, 10);
  p[10] = Math.round((nums.reduce((a, b) => a + b, 0) / nums.length) * 10) / 10;
}
writeSheet("02 Ranking Players", playerHeaders, players, [220, 230, 95, 110, 110, 110, 110, 110, 110, 150, 100, 520]);
writeSheet("03 Matriz Vendas", angleHeaders, angles, [220, 260, 260, 260, 460]);
writeSheet("04 Analise Comercial", commercialHeaders, commercial, [180, 280, 250, 220, 260, 300, 300]);
writeSheet("05 Plano Aprendizado", learningHeaders, learning, [180, 420, 520]);

const sourcesRows = Object.entries(source).map(([key, value]) => [key, value]);
writeSheet("06 Fontes", ["Fonte", "Link / observacao"], sourcesRows, [180, 850]);

const csv = [masterHeaders, ...rows]
  .map((r) => r.map((v) => `"${String(v ?? "").replaceAll('"', '""')}"`).join(","))
  .join("\n");
await fs.writeFile(path.join(outputDir, "lista_master_resorts_brasil.csv"), csv, "utf8");

const report = `# Pesquisa de mercado imobiliario turistico no Brasil

Consulta: ${consultationDate}

## Diagnostico objetivo

Esta entrega deve ser lida como **lista ampla encontrada com base nas fontes consultadas**, nao como inventario definitivo de todos os empreendimentos do Brasil.

O mercado se divide em cinco blocos comerciais:

1. **Multipropriedade e casa de ferias compartilhada**: GAV e Hot Beach sao os exemplos mais didaticos porque explicam uso, fracao, ferias e beneficios.
2. **Parque + resort**: Beach Park, Aviva, Hot Beach e Gramado Parks usam atracao ancora para aumentar desejo, recorrencia e prova de lazer.
3. **Hotelaria familiar premium**: Taua, Bourbon, Wish, Salinas, Cana Brava, Club Med e Vila Gale vendem baixa preocupacao operacional, estrutura e memoria familiar.
4. **Residence club / alto padrao**: Aviva Residence Clubs e Fasano/JHSF trabalham exclusividade, privacidade, marca, design e destino.
5. **Hotelaria pura com forca comercial**: importante para benchmarking, mas nao deve ser tratada como produto imobiliario sem evidencia de venda de unidade, cota ou fracao.

## Evidencias usadas

- GAV confirma atuacao em turismo, hotelaria e multipropriedade e lista empreendimentos prontos como Porto 2 Life, Pyreneus, Porto Alto, Premium, Exclusive e Park.
- Gramado Parks lista hospitalidade, gastronomia e entretenimento, incluindo Buona Vitta, Exclusive, Bella, Aquan, Namareh, Wyndham Gramado Termas, Snowland, Acquamotion e Acquaventura.
- WAM Hoteis lista 10 opcoes de hoteis e resorts em Caldas Novas, Aquiraz, Buzios e Porto Seguro.
- Aviva separa hoteis, Residence Clubs e Clube de Ferias, com Rio Quente, Costa do Sauipe, Hot Park, InCasa e InCanto.
- Beach Park mostra resorts, Aqua Park, Parque Arvorar, Vacation Club e Beach Card no mesmo ecossistema.
- Hot Beach Residence Club confirma multipropriedade e RCI, alem de Hot Beach Suites, Hot Beach You e Hot Beach Raizes.
- diRoma lista sua rede de hoteis, Acqua Park e VIP Club.
- Taua, Vila Gale, Grupo Wish, Bourbon, Club Med, Jurema, Nannai, Fasano e outros entraram como benchmarks de hotelaria, luxo, all inclusive ou clube.

## O que esta fraco no mercado

- Muita comunicacao mistura investimento, ferias e patrimonio sem explicar regime juridico.
- Algumas ofertas vendem "casa de ferias" antes de educar o cliente sobre fracao, uso, taxa anual, disponibilidade e liquidez.
- Resorts de hotelaria pura muitas vezes sao confundidos com empreendimento imobiliario. Isso atrapalha treinamento e pode gerar promessa comercial errada.
- O discurso de valorizacao do destino e perigoso quando vira promessa de retorno financeiro.

## Recomendacao comercial

Para aprender a vender em escala, estude primeiro os players por funcao:

- **GAV e Hot Beach** para multipropriedade.
- **Beach Park e Aviva** para ecossistema parque + resort + recorrencia.
- **Taua e Club Med** para promessa familiar e baixa preocupacao operacional.
- **Fasano/JHSF e Aviva Residence Clubs** para alto padrao, status e curadoria.
- **WAM, diRoma e Gramado Parks** para escala, destino popular e operacao de alto volume.

## Matriz de cuidado juridico

Evite: "rentabilidade garantida", "valorizacao certa", "investimento sem risco", "renda passiva garantida", "sempre disponivel", "voce pode vender quando quiser".

Use: "potencial", "historico do destino", "modelo de uso", "sujeito a regras contratuais", "necessita validacao documental", "a disponibilidade depende de regras do programa".

## Proxima rodada ideal

1. Validar cada produto imobiliario em paginas de venda ativas.
2. Checar incorporacao, SPE, memorial e regime juridico.
3. Consultar Cadastur para operacao hoteleira quando necessario.
4. Consultar CVM quando houver captacao com promessa de retorno.
5. Fazer auditoria reputacional por Reclame Aqui apenas como sinal de risco, nunca como unica fonte.
6. Criar scripts por produto: multipropriedade, clube, diaria, residence club e alto padrao.
`;
await fs.writeFile(path.join(outputDir, "relatorio_estrategico_resorts_brasil.md"), report, "utf8");

const previewSheets = ["00 LEIA PRIMEIRO", "01 Lista Master", "02 Ranking Players", "03 Matriz Vendas", "05 Plano Aprendizado"];
for (const sheetName of previewSheets) {
  const png = await workbook.render({ sheetName, autoCrop: "all", scale: 1, format: "png" });
  await fs.writeFile(path.join(outputDir, `${sheetName.replaceAll(" ", "_")}.png`), new Uint8Array(await png.arrayBuffer()));
}

const errors = await workbook.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
  options: { useRegex: true, maxResults: 50 },
  summary: "final formula error scan",
});
console.log(errors.ndjson);

const summary = await workbook.inspect({
  kind: "workbook,sheet,table",
  maxChars: 5000,
  tableMaxRows: 4,
  tableMaxCols: 5,
});
console.log(summary.ndjson);

const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(path.join(outputDir, "base_mercado_resorts_multipropriedade_brasil.xlsx"));

console.log(JSON.stringify({
  xlsx: path.join(outputDir, "base_mercado_resorts_multipropriedade_brasil.xlsx"),
  csv: path.join(outputDir, "lista_master_resorts_brasil.csv"),
  report: path.join(outputDir, "relatorio_estrategico_resorts_brasil.md"),
  rows: rows.length,
}, null, 2));
