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

export async function callGeminiDirectly(prompt: string, customKey?: string, model = "gemini-2.5-flash"): Promise<string> {
  const apiKey = getActiveGeminiKey(customKey);

  // 1. Tenta prioritariamente via Backend Seguro (/api/generate)
  try {
    const headers: Record<string, string> = { "Content-Type": "application/json" };
    if (apiKey) {
      headers["x-gemini-api-key"] = apiKey;
    }
    const res = await fetch("/api/generate", {
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
    console.warn("Proxy /api/generate não respondeu, tentando chamada direta:", backendErr);
  }

  // 2. Se houver chave configurada, tenta chamada direta ao Google com Exponential Backoff
  if (apiKey && apiKey.startsWith("AIzaSy")) {
    const fallbackModels = [
      model,
      "gemini-2.5-flash",
      "gemini-2.0-flash",
      "gemini-1.5-flash",
      "gemini-2.5-flash-lite",
      "gemini-1.5-pro"
    ].filter((v, i, a) => a.indexOf(v) === i);

    for (const m of fallbackModels) {
      let attempts = 0;
      while (attempts < 2) {
        try {
          const url = `https://generativelanguage.googleapis.com/v1beta/models/${m}:generateContent?key=${apiKey}`;
          const controller = new AbortController();
          const timeoutId = setTimeout(() => controller.abort(), 12000);

          const res = await fetch(url, {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            signal: controller.signal,
            body: JSON.stringify({
              contents: [{ parts: [{ text: prompt }] }],
              generationConfig: {
                temperature: 0.7,
                topP: 0.9,
                maxOutputTokens: 4096
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

export function normalizeCitationsToABNT2023(text: string): string {
  return text.replace(/\(([A-ZÁÉÍÓÚÂÊÔÃÕÇ]{2,})(,\s*\d{4}(?:,\s*p\.\s*\d+)?)\)/g, (_, author, rest) => {
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
- Título do Trabalho centralizado em negrito no topo.
- Título da seção centralizado: "RESUMO"
- EXATAMENTE 1 linha em branco entre a palavra "RESUMO" e o início do parágrafo.
- Texto em PARÁGRAFO ÚNICO contínuo e justificado (150 a 500 palavras) sem recuo na primeira linha, contendo: objetivo do estudo, metodologia utilizada, principais resultados e conclusão sintética.
- 1 linha em branco após o parágrafo.
- "Palavras-chave: [3 a 5 termos representativos separados por ponto e finalizados por ponto]."
- PROIBIDO incluir citações diretas longas, tabelas ou sumário.`;
  } else if (documentType === "redacao") {
    // REDAÇÃO ENEM NOTA 1000
    genreInstructions = `ESTRUTURA OBRIGATÓRIA DE REDAÇÃO (PADRÃO ENEM NOTA 1000):
- REGRA SUPREMA: PROIBIDO gerar Capa, Folha de Rosto, Sumário ou títulos de seções.
- Título opcional centralizado.
- 4 parágrafos contínuos e articulados:
  1. Introdução: Apresentação do tema e tese explícita com 2 argumentos centrais.
  2. Desenvolvimento 1: Aprofundamento do argumento 1 com repertório sociocultural legítimo.
  3. Desenvolvimento 2: Aprofundamento do argumento 2 com dados e fundamentação crítica.
  4. Conclusão: Proposta de intervenção social detalhada contendo os 5 elementos obrigatórios (Agente, Ação, Meio/Modo, Efeito e Detalhamento).`;
  } else if (documentType === "artigo" || documentType === "artigo_cientifico" || documentType === "artigo_academico") {
    // ABNT NBR 6022: ARTIGOS CIENTÍFICOS EM PUBLICAÇÃO PERIÓDICA
    genreInstructions = `ESTRUTURA OBRIGATÓRIA DE ARTIGO CIENTÍFICO (ABNT NBR 6022):
- REGRA: Artigos científicos NÃO levam Sumário formal (a estrutura é contínua).
- Título e subtítulo centralizados.
- Autor(es) e vinculação institucional.
- RESUMO em português (100 a 250 palavras) em parágrafo único + Palavras-chave.
- ABSTRACT em inglês correspondente + Keywords.
- 1 INTRODUÇÃO (Contextualização, problema, hipótese, objetivos e justificativa).
- 2 DESENVOLVIMENTO / FUNDAMENTAÇÃO TEÓRICA E METODOLOGIA
- 2.1 Análise das Dimensões Teóricas
- 3 RESULTADOS E DISCUSSÃO (com 1 Tabela padrão IBGE e 1 Quadro com Fonte).
- 4 CONSIDERAÇÕES FINAIS
- REFERÊNCIAS (ABNT NBR 6023:2025 alinhadas à esquerda com entrelinha simples).`;
  } else if (documentType === "projeto" || documentType === "projeto_pesquisa") {
    // ABNT NBR 15287: PROJETO DE PESQUISA
    genreInstructions = `ESTRUTURA OBRIGATÓRIA DE PROJETO DE PESQUISA (ABNT NBR 15287):
- SUMÁRIO com pontilhados líderes (NBR 6027).
- 1 TEMA E PROBLEMATIZAÇÃO
- 2 HIPÓTESES
- 3 OBJETIVOS (Geral e Específicos)
- 4 JUSTIFICATIVA E RELEVÂNCIA CIENTÍFICA
- 5 REVISÃO DE LITERATURA / FUNDAMENTAÇÃO TEÓRICA
- 6 METODOLOGIA DE PESQUISA (Tipo, instrumentos de coleta e procedimentos)
- 7 CRONOGRAMA DE EXECUÇÃO (Tabela em formato IBGE)
- 8 ORÇAMENTO (Se aplicável)
- REFERÊNCIAS (NBR 6023).`;
  } else if (documentType === "relatorio") {
    // ABNT NBR 10719: RELATÓRIO TÉCNICO-CIENTÍFICO
    genreInstructions = `ESTRUTURA OBRIGATÓRIA DE RELATÓRIO TÉCNICO-CIENTÍFICO (ABNT NBR 10719):
- RESUMO em português + Palavras-chave.
- SUMÁRIO (NBR 6027).
- 1 INTRODUÇÃO (Objetivo do relatório, escopo e instituições envolvidas).
- 2 DESENVOLVIMENTO DAS ATIVIDADES / METODOLOGIA APLICADA
- 3 RESULTADOS OBTIDOS E ANÁLISE DE DADOS (com Tabela IBGE e Ilustrações).
- 4 CONCLUSÕES E RECOMENDAÇÕES TÉCNICAS
- REFERÊNCIAS (NBR 6023).`;
  } else {
    // ABNT NBR 14724: TRABALHOS ACADÊMICOS (TCC, MONOGRAFIA, DISSERTAÇÃO)
    genreInstructions = `ESTRUTURA OBRIGATÓRIA DE TRABALHO ACADÊMICO / TCC / MONOGRAFIA (ABNT NBR 14724):
- RESUMO em português (150 a 250 palavras) + Palavras-chave (NBR 6028).
- ABSTRACT em inglês + Keywords.
- SUMÁRIO completo com pontilhados líderes regulares (NBR 6027).
- 1 INTRODUÇÃO (Problematização, hipótese, objetivos e estrutura).
- 2 FUNDAMENTAÇÃO TEÓRICA E METODOLOGIA
- 2.1 Análise Sistemática da Literatura
- 3 RESULTADOS E DISCUSSÃO (Obrigatório: 1 Tabela padrão IBGE e 1 Quadro com Fonte).
- 4 CONSIDERAÇÕES FINAIS (Síntese dos achados e sugestões para estudos futuros).
- REFERÊNCIAS (NBR 6023:2025 em ordem alfabética).`;
  }

  const systemPrompt = `Você é uma autoridade máxima em Redação e Normalização Acadêmica Brasileira (UNESP, USP e ABNT).
Elabore um(a) ${selectedTypeName} magistral sobre "${cleanTitle}" ${subtitle ? `com subtítulo "${subtitle}"` : ""}.

================================================================================
🚨 INSTRUÇÕES MANDATÓRIAS E PRIORITÁRIAS DO USUÁRIO:
${prompt ? `O usuário determinou expressamente as seguintes instruções que DEVEM ser integralmente cumpridas no conteúdo:\n"${prompt}"` : "Desenvolva o tema com profundidade científica máxima e dados atualizados."}
================================================================================

${genreInstructions}

NORMAS LINGUÍSTICAS E TÉCNICAS INEGOCIÁVEIS:
1. 🛡️ REGRA MESTRE INVIOLÁVEL DE VERACIDADE CIENTÍFICA (ZERO ALUCINAÇÃO):
   - Busque e utilize SOMENTE informações oriundas de artigos científicos consolidados, periódicos indexados (SciELO, Scopus, Web of Science, Capes, Google Scholar) e fontes oficiais respeitadas (IBGE, OMS, IPEA, Ministérios e Universidades).
   - Proibido inventar dados, autores ou citações falsas.
2. EXCELÊNCIA GRAMATICAL: Redação culta formal impecável, sem desvios de regência, crase ou pontuação.
3. CITAÇÕES (NBR 10520:2023): Sistema autor-data em caixa mista: (Silva, 2023, p. 15).
4. REFERÊNCIAS (NBR 6023:2025): Alinhadas à esquerda, entrelinha simples, separadas por 1 linha em branco.
5. ZERO CLICHÊS DE IA: Proibido usar "Em suma", "Vale ressaltar", "No cenário atual", "Podemos concluir".`;

  try {
    const generated = await callGeminiDirectly(systemPrompt, customGeminiKey, "gemini-2.5-flash");
    if (generated && generated.trim().length > 50) {
      const normalized = normalizeCitationsToABNT2023(generated);
      return prefixHeader + normalized;
    }
    throw new Error("A IA não retornou um conteúdo válido para este trabalho.");
  } catch (err: any) {
    console.error("[EMIA Engine] Falha na chamada da IA:", err);
    throw new Error(err?.message || "Falha na comunicação com os servidores de IA.");
  }
}
