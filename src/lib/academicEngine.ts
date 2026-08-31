/**
 * EDUTECH.EMIA - Motor Acadêmico Completo & Normalização ABNT (UNESP & USP)
 * Gera Capa, Contra-Capa (Folha de Rosto), Resumo, Abstract, Sumário, Seções e Referências
 */

export interface GenerateOptions {
  title: string;
  subtitle?: string;
  documentType: string;
  prompt?: string;
  studentName?: string;
  course?: string;
  institution?: string;
  city?: string;
  year?: string;
  advisor?: string;
  universityTemplate?: "unesp" | "usp" | "abnt_padrao";
  includeAbstract?: boolean;
  includeApprovalPage?: boolean;
  customGeminiKey?: string;
  googleToken?: string;
}

export function getActiveGeminiKey(customKey?: string): string {
  // 1. Chave customizada inserida pelo usuário no modal
  if (customKey && customKey.trim().length > 10) return customKey.trim();

  // 2. Chave do ambiente Vite (se configurada)
  const envKey = typeof import.meta !== "undefined" && (import.meta as any).env ? (import.meta as any).env.VITE_GEMINI_API_KEY : "";
  if (envKey && envKey.trim().length > 10) return envKey.trim();

  // 3. Chave do LocalStorage do navegador
  const localKey = typeof localStorage !== "undefined" ? localStorage.getItem("emia_custom_gemini_key") : null;
  if (localKey && localKey.trim().length > 10) return localKey.trim();

  return "";
}

export async function callGeminiDirectly(prompt: string, customKey?: string, model = "gemini-3.6-flash"): Promise<string> {
  const apiKey = getActiveGeminiKey(customKey);

  // 1. Tenta prioritariamente via Backend Seguro (/api/generate e URL de producao)
  const endpoints = ["/api/generate", "https://edutech.emia.workers.dev/api/generate"];
  for (const endpoint of endpoints) {
    try {
      const headers: Record<string, string> = { "Content-Type": "application/json" };
      if (apiKey) {
        headers["x-gemini-api-key"] = apiKey;
      }
      const res = await fetch(endpoint, {
        method: "POST",
        headers,
        body: JSON.stringify({ prompt, model, temperature: 0.7 })
      });
      if (res.ok) {
        const data = await res.json();
        if (data.text && typeof data.text === "string" && data.text.trim().length > 10) {
          return data.text.trim();
        }
        const candText = data?.candidates?.[0]?.content?.parts?.[0]?.text;
        if (candText && candText.trim().length > 10) {
          return candText.trim();
        }
      }
    } catch (backendErr) {
      console.warn(`Proxy ${endpoint} não respondeu:`, backendErr);
    }
  }

  // 2. Chamada direta ao Google com Exponential Backoff
  if (apiKey && apiKey.length > 10) {
    const fallbackModels = [
      model,
      "gemini-3.6-flash",
      "gemini-3-flash-preview",
      "gemini-3.7-flash"
    ].filter((v, i, a) => a.indexOf(v) === i);

    for (const m of fallbackModels) {
      let attempts = 0;
      while (attempts < 2) {
        try {
          const url = `https://generativelanguage.googleapis.com/v1beta/models/${m}:generateContent?key=${apiKey}`;
          const controller = new AbortController();
          const timeoutId = setTimeout(() => controller.abort(), 25000);

          const res = await fetch(url, {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            signal: controller.signal,
            body: JSON.stringify({
              contents: [{ parts: [{ text: prompt }] }],
              generationConfig: {
                temperature: 0.7,
                topP: 0.95,
                maxOutputTokens: 8192
              }
            })
          });
          clearTimeout(timeoutId);

          if (res.status === 429 || res.status === 500 || res.status === 503) {
            attempts++;
            await new Promise(r => setTimeout(r, Math.pow(2, attempts) * 1000 + Math.random() * 400));
            continue;
          }

          if (res.ok) {
            const json = await res.json();
            const text = json?.candidates?.[0]?.content?.parts?.[0]?.text;
            if (text && text.trim().length > 0) {
              return text.trim();
            }
          }
          break;
        } catch (e) {
          attempts++;
          await new Promise(r => setTimeout(r, Math.pow(2, attempts) * 1000));
        }
      }
    }
  }

  throw new Error("Não foi possível obter resposta dos servidores de IA.");
}

export function sanitizeHallucinatedAuthors(text: string): string {
  if (!text) return text;
  
  // Remove blocos de autoria e afiliação alucinados pela IA no início do texto
  let clean = text
    .replace(/^(?:\s*#+\s*.*?\n)?(?:\s*(?:\*\*|__)?(?:Autor|Autores|Autoria|Coautor|Coautores|Coautora|Orientador|Orientadora|Pesquisador|Pesquisadora):(?:\*\*|__)?[^\n]*\n?)+/gmi, '')
    .replace(/(?:^|\n)(?:\*\*|__)?(?:Autor|Autores|Autoria|Coautor|Coautores|Coautora|Orientador|Orientadora|Pesquisador|Pesquisadora):(?:\*\*|__)?[^\n]*/gi, '')
    .replace(/(?:^|\n)(?:\*|_)?(?:Livre-Docente|Doutor|Doutora|Prof\.|Dr\.|Dra\.|Pesquisador|Docente|Mestre em)[^\n]*(?:\*|_)?/gi, '')
    .replace(/(?:^|\n)(?:\*|_)?(?:Universidade|Faculdade|Instituto|Centro Universitário)[^\n]*(?:\*|_)?/gi, '');
  
  return clean.replace(/^\s+/, '');
}

export function normalizeCitationsToABNT2023(text: string): string {
  const sanitized = sanitizeHallucinatedAuthors(text);
  return sanitized.replace(/\(([A-ZÁÉÍÓÚÂÊÔÃÕÇ]{2,})(,\s*\d{4}(?:,\s*p\.\s*\d+)?)\)/g, (_, author, rest) => {
    const titleCaseAuthor = author.charAt(0).toUpperCase() + author.slice(1).toLowerCase();
    return `(${titleCaseAuthor}${rest})`;
  });
}

export async function humanizeTextWithGemini(text: string, customKey?: string): Promise<string> {
  const prompt = `Você é um especialista em escrita acadêmica humana, clareza e originalidade textual (Padrão UNESP/USP).
Reescreva e humanize o texto abaixo, eliminando padrões robóticos, clichês de IA (como "Em suma", "Vale ressaltar", "No cenário atual") e conectivos artificiais.
Mantenha o rigor formal, as citações no formato ABNT NBR 10520:2023 e os conceitos acadêmicos intactos.
NÃO use saudações nem avisos. Retorne apenas o texto reescrito:

${text}`;
  return await callGeminiDirectly(prompt, customKey);
}

export async function correctSpellingWithGemini(text: string, customKey?: string): Promise<string> {
  const prompt = `Você é um revisor filológico da língua portuguesa e normas de publicação científica.
Corrija rigorosamente a ortografia, a concordância verbal e nominal, a regência e a pontuação do texto a seguir.
Mantenha 100% da estrutura, termos técnicos e citações.
Retorne apenas o texto corrigido sem nenhum comentário adicional:

${text}`;
  return await callGeminiDirectly(prompt, customKey);
}

export function buildDynamicTOCBlock(pages: string[]): string {
  const tocEntries: { title: string; page: number }[] = [];
  
  pages.forEach((pageContent, idx) => {
    const pageNum = idx + 1;
    const cleanP = pageContent.trim();
    const cleanUpper = cleanP.toUpperCase().replace(/^#+\s*/, '');
    
    // Elementos pré-textuais NÃO entram no Sumário (ABNT NBR 6027 item 4.2)
    if (cleanUpper.startsWith("CAPA") || cleanP === "CAPA_AUTO" || 
        cleanUpper.startsWith("FOLHA") || cleanP === "FOLHA_ROSTO_AUTO" || 
        cleanUpper.startsWith("SUMÁRIO") || cleanUpper.startsWith("RESUMO") || cleanUpper.startsWith("ABSTRACT") ||
        cleanUpper.startsWith("AGRADECIMENTOS") || cleanUpper.startsWith("DEDICATÓRIA")) {
      return;
    }

    const lines = cleanP.split('\n');
    for (const line of lines) {
      const clean = line.trim().replace(/^#+\s*/, '');
      if (!clean) continue;

      // Seções numeradas (1 INTRODUÇÃO, 2 FUNDAMENTAÇÃO, 2.1 Subseção, etc.)
      const isNumbered = /^\d+(?:\.\d+)*\s+[A-ZÀ-Ú]/.test(clean);
      // Elementos pós-textuais (REFERÊNCIAS, APÊNDICES, ANEXOS)
      const isPostTextual = /^(REFERÊNCIAS(?:\s+BIBLIOGRÁFICAS)?|APÊNDICE|ANEXO)\b/i.test(clean);

      if (isNumbered || isPostTextual) {
        const upperTitle = clean.toUpperCase();
        if (!upperTitle.startsWith("SUMÁRIO") && !upperTitle.startsWith("RESUMO") && !upperTitle.startsWith("ABSTRACT") && !tocEntries.some(e => e.title === upperTitle)) {
          tocEntries.push({
            title: upperTitle,
            page: pageNum
          });
        }
      }
    }
  });

  if (tocEntries.length === 0) {
    tocEntries.push(
      { title: "1 INTRODUÇÃO", page: 6 },
      { title: "2 FUNDAMENTAÇÃO TEÓRICA E METODOLOGIA", page: 7 },
      { title: "3 RESULTADOS E DISCUSSÃO", page: 9 },
      { title: "4 CONSIDERAÇÕES FINAIS", page: 11 },
      { title: "REFERÊNCIAS", page: 12 }
    );
  }

  const formattedTOCLines = tocEntries.map(entry => {
    const dotsCount = Math.max(5, 70 - entry.title.length - String(entry.page).length);
    const dots = ".".repeat(dotsCount);
    return `${entry.title} ${dots} ${entry.page}`;
  });

  return `SUMÁRIO\n\n${formattedTOCLines.join('\n')}`;
}

export function paginateAcademicDocument(rawText: string, documentType = "artigo"): string {
  if (!rawText || !rawText.trim()) return rawText;

  let cleanText = rawText.replace(/\r\n/g, "\n").trim();
  const pageBreakRegex = /\s*---\s*\[(?:QUEBRA DE P[AÁ]GINA|NOVA P[AÁ]GINA)\]\s*---\s*|\s*\[(?:QUEBRA DE P[AÁ]GINA|NOVA P[AÁ]GINA)\]\s*/i;
  const initialBlocks = cleanText.split(pageBreakRegex).map(b => b.trim()).filter(Boolean);

  const rawFinalPages: string[] = [];

  for (const block of initialBlocks) {
    const isCover = block.startsWith("CAPA_AUTO") || block.startsWith("CAPA\n") || block === "CAPA";
    const isTitlePage = block.startsWith("FOLHA_ROSTO_AUTO") || block.startsWith("FOLHA DE ROSTO\n") || block === "FOLHA DE ROSTO";
    const isTOC = block.startsWith("SUMÁRIO") || block.startsWith("# SUMÁRIO") || block.replace(/^#+\s*/, '').startsWith("SUMÁRIO");

    if (isCover || isTitlePage || isTOC) {
      rawFinalPages.push(block);
      continue;
    }

    if (["resumo", "redacao"].includes(documentType)) {
      rawFinalPages.push(block);
      continue;
    }

    // Separa antes de grandes seções ABNT
    const sectionSplitRegex = /\n(?=(?:#+\s*)?(?:RESUMO|ABSTRACT|SUMÁRIO|\d+\s+[A-ZÀ-Ú\s]{3,}|REFERÊNCIAS(?:\s+BIBLIOGRÁFICAS)?|CONSIDERAÇÕES FINAIS|CONCLUSÃO)\b)/gi;
    const sectionChunks = block.split(sectionSplitRegex).map(s => s.trim()).filter(Boolean);

    for (const chunk of sectionChunks) {
      // Se a seção ultrapassar ~1800 caracteres, quebra suavemente entre parágrafos
      const paragraphs = chunk.split(/\n\n+/);
      let curPage = "";

      for (const p of paragraphs) {
        if ((curPage + "\n\n" + p).length > 1800 && curPage.trim().length > 0) {
          rawFinalPages.push(curPage.trim());
          curPage = p;
        } else {
          curPage = curPage ? curPage + "\n\n" + p : p;
        }
      }
      if (curPage.trim()) {
        rawFinalPages.push(curPage.trim());
      }
    }
  }

  // Recalcula o Sumário dinâmico com os números de página reais de cada seção
  const finalPages = [...rawFinalPages];
  const tocIdx = finalPages.findIndex(p => p.startsWith("SUMÁRIO") || p.startsWith("# SUMÁRIO") || p.replace(/^#+\s*/, '').startsWith("SUMÁRIO"));
  
  if (tocIdx >= 0) {
    finalPages[tocIdx] = buildDynamicTOCBlock(finalPages);
  }

  return finalPages.join("\n\n--- [QUEBRA DE PÁGINA] ---\n\n");
}

export async function generateAcademicText(options: GenerateOptions): Promise<string> {
  const {
    title,
    subtitle = "",
    documentType = "artigo",
    prompt = "",
    studentName = "",
    course = "",
    institution = "",
    city = "",
    year = "",
    advisor = "",
    universityTemplate = "abnt_padrao",
    includeAbstract = true,
    includeApprovalPage = false,
    customGeminiKey,
  } = options;

  const cleanTitle = title.trim() || "Trabalho Acadêmico Geral";
  const cleanTopic = cleanTitle;
  const currentYear = new Date().getFullYear();

  const typeMap: Record<string, string> = {
    artigo: "Artigo Acadêmico",
    artigo_cientifico: "Artigo Científico",
    redacao: "Redação Dissertativo-Argumentativa (Padrão ENEM)",
    resenha: "Resenha Crítica",
    resumo: "Resumo / Fichamento",
    estudo_caso: "Estudo de Caso",
    relatorio: "Relatório Técnico-Científico",
    monografia: "Monografia / TCC",
    projeto: "Projeto de Pesquisa",
    artigo_opiniao: "Artigo de Opinião"
  };

  const selectedTypeName = typeMap[documentType] || documentType || "Artigo Acadêmico";
  
  // DOCUMENTOS QUE EXIGEM CAPA E FOLHA DE ROSTO (CONTRA-CAPA) COMPLETAS:
  // Artigo Científico, Artigo Acadêmico, Monografia, TCC, Relatório, Projeto de Pesquisa, Estudo de Caso.
  // APENAS Resumo simples e Redação ENEM não possuem capa separada.
  const requiresCoverAndTitlePage = !["resumo", "redacao", "resenha"].includes(documentType);

  let prefixHeader = "";

  if (requiresCoverAndTitlePage) {
    const inst = institution ? institution.toUpperCase() : "INSTITUIÇÃO DE ENSINO SUPERIOR";
    const crs = course ? course.toUpperCase() : "CURSO DE GRADUAÇÃO / PÓS-GRADUAÇÃO";
    const aut = studentName ? studentName.toUpperCase() : "NOME DO(A) AUTOR(A)";
    const tit = cleanTitle.toUpperCase();
    const sub = subtitle ? `: ${subtitle}` : "";
    const cid = city ? city.toUpperCase() : "CIDADE - UF";
    const an = year || String(currentYear);
    const adv = advisor || "Prof. Dr. Orientador";

    // 1. CAPA OFICIAL (Elemento Pré-Textual 1 - ABNT NBR 14724)
    prefixHeader = `CAPA_AUTO\n${inst}
${crs}



${aut}



${tit}${sub}







${cid}
${an}

--- [QUEBRA DE PÁGINA] ---

`;

    // 2. FOLHA DE ROSTO / CONTRA-CAPA (Elemento Pré-Textual 2 - ABNT NBR 14724)
    const presentationNote = documentType.includes("artigo")
      ? `Artigo científico/acadêmico apresentado ao(à) ${inst}, como requisito de avaliação na área de ${course || "conhecimento acadêmico"}.`
      : documentType === "projeto"
      ? `Projeto de pesquisa apresentado ao(à) ${inst}, como requisito para qualificação e desenvolvimento do trabalho acadêmico.`
      : documentType === "relatorio"
      ? `Relatório técnico-científico apresentado ao(à) ${inst}, para comprovação e avaliação das atividades desenvolvidas.`
      : `Trabalho de Conclusão de Curso apresentado ao(à) ${inst}, como requisito parcial para obtenção de grau em ${course || "Graduação"}.`;

    prefixHeader += `FOLHA_ROSTO_AUTO\n${aut}



${tit}${sub}


${presentationNote}
Orientador(a): ${adv}





${cid}
${an}

--- [QUEBRA DE PÁGINA] ---

`;

    // 3. Folha de Aprovação (se solicitada ou para Monografias)
    if (includeApprovalPage || documentType === "monografia" || documentType === "trabalho_academico") {
      prefixHeader += `${aut}


${tit}${sub}


Trabalho aprovado em ___ de ____________ de ${an}.

BANCA EXAMINADORA:

________________________________________
${adv} (Orientador(a))
${inst}

________________________________________
Prof(a). Dr(a). Avaliador(a) 1
${inst}

________________________________________
Prof(a). Dr(a). Avaliador(a) 2


${cid}
${an}

--- [QUEBRA DE PÁGINA] ---

`;
    }
  } else if (documentType === "resenha") {
    prefixHeader = `RESENHA CRÍTICA: ${cleanTitle.toUpperCase()}${subtitle ? `: ${subtitle}` : ""}
${studentName ? `Resenhista: ${studentName}
` : ""}
`;
  }

  // ESTRUTURA RIGOROSA DE CADA TIPO DE DOCUMENTO CONFORME AS NORMAS ABNT
  let genreInstructions = "";
  
  if (documentType === "resumo") {
    // ABNT NBR 6028: RESUMOS E FICHAMENTOS
    genreInstructions = `ESTRUTURA OBRIGATÓRIA DE RESUMO / FICHAMENTO (ABNT NBR 6028):
- REGRA SUPREMA: PROIBIDO gerar Capa, Folha de Rosto, Sumário ou divisões de capítulos! O resumo é um texto em fluxo contínuo.
- ESPAÇAMENTO MANDATÓRIO ABNT: Texto em entrelinha simples (1,0), alinhamento justificado, sem recuo na primeira linha.
- Título do Trabalho centralizado em negrito no topo.
- Título da seção centralizado: "RESUMO"
- EXATAMENTE 1 linha em branco entre a palavra "RESUMO" e o início do parágrafo.
- Texto em PARÁGRAFO ÚNICO contínuo e justificado (150 a 500 palavras), contendo: objetivo do estudo, metodologia utilizada, principais resultados e conclusão sintética.
- 1 linha em branco após o parágrafo.
- "Palavras-chave: [3 a 5 termos representativos separados por ponto e finalizados por ponto]."
- PROIBIDO incluir citações diretas longas, tabelas ou sumário.`;
  } else if (documentType === "redacao") {
    // REDAÇÃO ENEM NOTA 1000
    genreInstructions = `ESTRUTURA OBRIGATÓRIA DE REDAÇÃO (PADRÃO ENEM NOTA 1000):
- REGRA SUPREMA: PROIBIDO gerar Capa, Folha de Rosto, Sumário ou títulos de seções.
- ESPAÇAMENTO MANDATÓRIO: Texto contínuo com entrelinha 1,5, recuo de parágrafo 1,25 cm e alinhamento justificado.
- Título opcional centralizado.
- 4 parágrafos contínuos e articulados:
  1. Introdução: Apresentação do tema e tese explícita com 2 argumentos centrais.
  2. Desenvolvimento 1: Aprofundamento do argumento 1 com repertório sociocultural legítimo.
  3. Desenvolvimento 2: Aprofundamento do argumento 2 com dados e fundamentação crítica.
  4. Conclusão: Proposta de intervenção social detalhada contendo os 5 elementos obrigatórios (Agente, Ação, Meio/Modo, Efeito e Detalhamento).`;
  } else if (documentType === "artigo" || documentType === "artigo_cientifico" || documentType === "artigo_academico") {
    // ABNT NBR 6022: ARTIGOS CIENTÍFICOS EM PUBLICAÇÃO PERIÓDICA
    genreInstructions = `================================================================================
🚨 REGRA CRÍTICA DE GERAÇÃO COMPLETA - ARTIGO CIENTÍFICO (ABNT NBR 6022):
VOCÊ DEVE GERAR O DOCUMENTO INTEIRO DO COMEÇO AO FIM, SEM OMITIR OU RESUMIR NENHUMA SEÇÃO!
DIRETRIZES DE ESPAÇAMENTO MANDATÓRIAS ABNT:
- CORPO DO TEXTO (Introdução, Fundamentação, Resultados, Considerações): Espaçamento entrelinhas de 1,5 com recuo de primeira linha de 1,25 cm e alinhamento justificado.
- CITAÇÕES DIRETAS LONGAS (> 3 linhas): Recuo de 4,0 cm da margem esquerda, fonte tamanho 10 pt e espaçamento simples (1,0), sem aspas.
- REFERÊNCIAS (NBR 6023:2025): Alinhadas à esquerda, espaçamento simples (1,0) e separadas entre si por 1 linha em branco simples.
- TABELAS E ILUSTRAÇÕES: Título no topo e Fonte na base em fonte 10 pt com espaçamento simples (1,0).

DESENVOLVA CADA UMA DAS SEGUINTES SEÇÕES COM TEXTO ACADÊMICO REAL, EXTENSO E APROFUNDADO:
- RESUMO (NBR 6028): Parágrafo único de 150 a 250 palavras em espaçamento simples + "Palavras-chave: [3 a 5 termos separados por ponto final]."
- ABSTRACT: Versão acadêmica completa em inglês + "Keywords: [termos em inglês]."
- SUMÁRIO (NBR 6027): Inicia obrigatoriamente na seção 1 INTRODUÇÃO, com pontilhados líderes uniformes.
- 1 INTRODUÇÃO (Mínimo 3 a 4 parágrafos densos: Contextualização, Problematização, Objetivos e Justificativa).
- 2 FUNDAMENTAÇÃO TEÓRICA E METODOLOGIA (Mínimo 4 a 6 parágrafos densos com revisões conceituais, citações em padrão ABNT NBR 10520:2023 em caixa mista como (Silva, 2023, p. 15) e instrumentos metodológicos).
- 3 RESULTADOS E DISCUSSÃO (Mínimo 4 parágrafos com 1 Tabela no padrão IBGE de laterais abertas e indicação de Fonte, e 1 Quadro com Fonte).
- 4 CONSIDERAÇÕES FINAIS (Mínimo 3 parágrafos com síntese dos resultados, limites e perspectivas futuras).
- REFERÊNCIAS (Lista completa com todas as fontes citadas no formato ABNT NBR 6023:2025 alinhadas à esquerda com entrelinha simples).
PROIBIDO PARAR NA INTRODUÇÃO! COMPLETE TODAS AS SEÇÕES ATÉ AS REFERÊNCIAS!
================================================================================`;
  } else if (documentType === "projeto" || documentType === "projeto_pesquisa") {
    // ABNT NBR 15287: PROJETO DE PESQUISA
    genreInstructions = `================================================================================
🚨 REGRA CRÍTICA DE GERAÇÃO COMPLETA - PROJETO DE PESQUISA (ABNT NBR 15287):
VOCÊ DEVE GERAR O DOCUMENTO INTEIRO DO COMEÇO AO FIM, SEM OMITIR OU RESUMIR NENHUMA SEÇÃO!
DIRETRIZES DE ESPAÇAMENTO MANDATÓRIAS ABNT:
- CORPO DO TEXTO: Espaçamento entrelinhas 1,5, parágrafos justificados e recuo de 1,25 cm na primeira linha.
- CITAÇÕES LONGAS E NOTAS: Espaçamento simples (1,0) com recuo de 4,0 cm para citações > 3 linhas.
- CRONOGRAMA E ORÇAMENTO: Tabelas com laterais abertas, espaçamento simples e indicação de Fonte.
- REFERÊNCIAS: Alinhadas à esquerda em espaçamento simples (1,0).

DESENVOLVA CADA UMA DAS SEGUINTES SEÇÕES COM TEXTO ACADÊMICO REAL, EXTENSO E APROFUNDADO:
- SUMÁRIO (NBR 6027).
- 1 TEMA E PROBLEMATIZAÇÃO (Mínimo 3 parágrafos densos).
- 2 HIPÓTESES E QUESTÕES NORTEADORAS (Mínimo 2 parágrafos).
- 3 OBJETIVOS (Objetivo Geral e Objetivos Específicos detalhados).
- 4 JUSTIFICATIVA E RELEVÂNCIA CIENTÍFICA (Mínimo 3 parágrafos).
- 5 REVISÃO DE LITERATURA / FUNDAMENTAÇÃO TEÓRICA (Mínimo 4 a 6 parágrafos com citações NBR 10520:2023).
- 6 METODOLOGIA DE PESQUISA (Mínimo 3 parágrafos: Delineamento, sujeitos, instrumentos e procedimentos).
- 7 CRONOGRAMA DE EXECUÇÃO (Tabela em formato IBGE de laterais abertas com Fonte).
- 8 ORÇAMENTO E RECURSOS (Quadro detalhado com Fonte).
- REFERÊNCIAS (Lista completa NBR 6023:2025).
PROIBIDO PARAR NA INTRODUÇÃO! COMPLETE TODAS AS SEÇÕES ATÉ AS REFERÊNCIAS!
================================================================================`;
  } else if (documentType === "relatorio") {
    // ABNT NBR 10719: RELATÓRIO TÉCNICO-CIENTÍFICO
    genreInstructions = `================================================================================
🚨 REGRA CRÍTICA DE GERAÇÃO COMPLETA - RELATÓRIO TÉCNICO-CIENTÍFICO (ABNT NBR 10719):
VOCÊ DEVE GERAR O DOCUMENTO INTEIRO DO COMEÇO AO FIM, SEM OMITIR OU RESUMIR NENHUMA SEÇÃO!
DIRETRIZES DE ESPAÇAMENTO MANDATÓRIAS ABNT:
- CORPO DO TEXTO: Espaçamento entrelinhas 1,5, recuo 1,25 cm e texto justificado.
- ELEMENTOS ESPECIAIS (Citações longas, notas, tabelas e referências): Espaçamento simples (1,0).

DESENVOLVA CADA UMA DAS SEGUINTES SEÇÕES COM TEXTO ACADÊMICO REAL, EXTENSO E APROFUNDADO:
- RESUMO em português + Palavras-chave.
- SUMÁRIO (NBR 6027).
- 1 INTRODUÇÃO (Mínimo 3 parágrafos: Escopo, objetivos do relatório e contextualização).
- 2 METODOLOGIA E PROCEDIMENTOS TÉCNICOS (Mínimo 3 a 4 parágrafos).
- 3 DESENVOLVIMENTO DAS ATIVIDADES E EXECUÇÃO (Mínimo 4 a 5 parágrafos).
- 4 RESULTADOS OBTIDOS E ANÁLISE DE DADOS (Mínimo 4 parágrafos com Tabela IBGE e Ilustrações com Fonte).
- 5 CONCLUSÕES E RECOMENDAÇÕES TÉCNICAS (Mínimo 3 parágrafos).
- REFERÊNCIAS (Lista completa NBR 6023:2025).
PROIBIDO PARAR NA INTRODUÇÃO! COMPLETE TODAS AS SEÇÕES ATÉ AS REFERÊNCIAS!
================================================================================`;
  } else {
    // ABNT NBR 14724: TRABALHOS ACADÊMICOS (TCC, MONOGRAFIA, DISSERTAÇÃO)
    genreInstructions = `================================================================================
🚨 REGRA CRÍTICA DE GERAÇÃO COMPLETA - TRABALHO ACADÊMICO / TCC / MONOGRAFIA (ABNT NBR 14724):
VOCÊ DEVE GERAR O DOCUMENTO INTEIRO DO COMEÇO AO FIM, SEM OMITIR OU RESUMIR NENHUMA SEÇÃO!
DIRETRIZES DE ESPAÇAMENTO MANDATÓRIAS ABNT (NBR 14724):
- CORPO DO TRABALHO: Espaçamento entrelinhas de 1,5 em todo o texto, parágrafos justificados com recuo de 1,25 cm na primeira linha.
- CITAÇÕES DIRETAS LONGAS (> 3 linhas): Recuo de 4,0 cm da margem esquerda, fonte 10 pt e espaçamento simples (1,0).
- NOTA DE APRESENTAÇÃO (Folha de Rosto): Recuo de 7,5 cm à esquerda, fonte 10 pt e espaçamento simples (1,0).
- REFERÊNCIAS (NBR 6023:2025): Alinhadas à esquerda, espaçamento simples (1,0) e separadas entre si por 1 linha em branco simples.
- LEGENDAS E FONTES DE TABELAS/ILUSTRAÇÕES: Fonte 10 pt e espaçamento simples (1,0).

DESENVOLVA CADA UMA DAS SEGUINTES SEÇÕES COM TEXTO ACADÊMICO REAL, EXTENSO E APROFUNDADO:
- RESUMO (NBR 6028): Parágrafo único de 150 a 250 palavras em espaçamento simples + "Palavras-chave: [3 a 5 termos separados por ponto final]."
- ABSTRACT: Versão acadêmica completa em inglês + "Keywords: [termos em inglês]."
- SUMÁRIO completo com pontilhados líderes regulares (NBR 6027).
- 1 INTRODUÇÃO (Mínimo 3 a 4 parágrafos densos: Contextualização, Problema, Hipóteses, Objetivos e Relevância).
- 2 FUNDAMENTAÇÃO TEÓRICA E REVISÃO DE LITERATURA (Mínimo 4 a 6 parágrafos densos com revisão crítica e citações NBR 10520:2023).
- 3 METODOLOGIA DA PESQUISA (Mínimo 3 a 4 parágrafos: Abordagem, universo/amostra, coleta e análise de dados).
- 4 RESULTADOS E DISCUSSÃO (Mínimo 4 parágrafos contendo 1 Tabela no formato tabular IBGE com laterais abertas e indicação de Fonte, e 1 Quadro com Fonte).
- 5 CONSIDERAÇÕES FINAIS (Mínimo 3 parágrafos: Conclusões, contribuições para a área e sugestões para estudos futuros).
- REFERÊNCIAS (Lista completa de fontes consultadas conforme NBR 6023:2025 em ordem alfabética).
PROIBIDO PARAR NA INTRODUÇÃO! COMPLETE TODAS AS SEÇÕES ATÉ AS REFERÊNCIAS!
================================================================================`;
  }

  const systemPrompt = `Você é uma autoridade máxima em Redação e Normalização Acadêmica Brasileira (UNESP, USP e ABNT).
Elabore um(a) ${selectedTypeName} completo, aprofundado e rigorosamente estruturado sobre "${cleanTitle}" ${subtitle ? `com subtítulo "${subtitle}"` : ""}.

================================================================================
🚨 INSTRUÇÕES MANDATÓRIAS E PRIORITÁRIAS DO USUÁRIO:
${prompt ? `O usuário determinou expressamente as seguintes instruções que DEVEM ser integralmente cumpridas no conteúdo:\n"${prompt}"` : "Desenvolva o tema com profundidade científica máxima, dados empíricos e rigor metodológico."}
================================================================================

${genreInstructions}

NORMAS LINGUÍSTICAS E TÉCNICAS INEGOCIÁVEIS:
1. 🛡️ REGRA CRÍTICA DE INEDITISMO & PROIBIÇÃO ABSOLUTA DE AUTORES FICTÍCIOS:
   - O trabalho é uma PRODUÇÃO CIENTÍFICA 100% INÉDITA E ORIGINAL.
   - PROIBIDO ABSOLUTAMENTE inventar, criar ou inserir nomes de autores, coautores, titulações acadêmicas (ex: Dr., Dra., Livre-Docente) ou afiliações institucionais (UNESP, USP, etc.) no topo ou corpo do texto gerado.
   - A autoria do trabalho pertence exclusivamente ao estudante/pesquisador solicitante e é configurada exclusivamente na Capa e Folha de Rosto. O texto gerado DEVE iniciar diretamente no título oficial e no RESUMO!
2. 🛡️ REGRA MESTRE DE VERACIDADE CIENTÍFICA (ZERO ALUCINAÇÃO):
   - Baseie-se apenas em conhecimento científico real, periódicos conceituados (SciELO, Scopus, Google Scholar) e fontes oficiais (IBGE, IPEA, OMS).
3. ESPAÇAMENTO E TIPOGRAFIA RIGOROSA (ABNT NBR 14724 & NBR 6022):
   - Corpo do texto: Espaçamento entrelinhas 1,5, justificado, recuo 1,25 cm.
   - Citações longas (> 3 linhas): Espaçamento simples (1,0), recuo 4,0 cm, fonte 10 pt.
   - Referências (NBR 6023:2025): Espaçamento simples (1,0), alinhamento à esquerda, separadas por 1 linha em branco.
   - Tabelas, Quadros e Legendas: Espaçamento simples (1,0), fonte 10 pt.
4. EXCELÊNCIA GRAMATICAL: Português culto formal, sem desvios de regência, crase ou pontuação.
5. CITAÇÕES (NBR 10520:2023): Sistema autor-data em caixa mista: (Silva, 2023, p. 15) ou Conforme Santos (2022). NUNCA use caixa alta integral como (SILVA, 2023).
6. REFERÊNCIAS (NBR 6023:2025): Alinhadas à esquerda, entrelinha simples, separadas por 1 linha em branco.
7. ZERO CLICHÊS DE IA: Proibido usar "Em suma", "Vale ressaltar", "No cenário atual", "Podemos concluir".`;

  try {
    const generated = await callGeminiDirectly(systemPrompt, customGeminiKey, "gemini-3.6-flash");
    if (generated && generated.trim().length > 50) {
      const normalized = normalizeCitationsToABNT2023(generated);
      const fullDoc = prefixHeader + normalized;
      return paginateAcademicDocument(fullDoc, documentType);
    }
    throw new Error("A IA não retornou um conteúdo válido para este trabalho.");
  } catch (err: any) {
    console.error("[EMIA Engine] Falha na chamada da IA:", err);
    throw new Error(err?.message || "Falha na comunicação com os servidores de IA.");
  }
}

