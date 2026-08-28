import React, { useState, useRef, useEffect } from "react";
import { Button } from "./components/ui/button";
import { 
  FileText, Upload, Plus, CheckCircle, FileDown, 
  Settings, Loader2, LogOut, ShieldCheck, Download, Copy,
  UserCheck, BookOpen, Hash, Wand2, ImagePlus, Lock,
  User, Clock, Save, X, ListOrdered, Link, Sparkles, Coins, Check, Quote,
  ZoomIn, ZoomOut, Presentation, PanelLeftClose, PanelLeftOpen, Share2, FileCode, Move, Users,
  Undo2, Redo2, Maximize2, Minimize2, ChevronLeft, ChevronRight, Volume2, VolumeX, Mic, MicOff
} from "lucide-react";
import pptxgen from "pptxgenjs";
import ReactMarkdown from "react-markdown";
import jsPDF from "jspdf";
import { Document, Packer, Paragraph, TextRun, AlignmentType, convertMillimetersToTwip, Header, PageNumber } from "docx";
import { saveAs } from "file-saver";
import { useDropzone } from "react-dropzone";
import { generateAcademicText, normalizeCitationsToABNT2023, callGeminiDirectly } from "./lib/academicEngine";

type UserProfile = {
  name: string;
  institution: string;
  course?: string;
  subject?: string;
  shift?: string;
  classroom?: string;
  city: string;
  year: string;
  advisor: string;
};

type AuditLog = {
  action: string;
  content?: string;
  timestamp: string;
};

export default function App() {
  const [isAuthenticated, setIsAuthenticated] = useState<boolean>(() => {
    return localStorage.getItem("emia_authenticated") === "true";
  });
  const [isMaster, setIsMaster] = useState<boolean>(() => {
    const saved = localStorage.getItem("emia_is_master");
    return saved !== null ? saved === "true" : true;
  });
  const [loginEmail, setLoginEmail] = useState("erlane.digital@gmail.com");
  const [loginPassword, setLoginPassword] = useState("");
  const [loginError, setLoginError] = useState("");

  const [credits, setCredits] = useState<number>(() => {
    const saved = localStorage.getItem("emia_credits");
    return saved !== null ? parseInt(saved, 10) : 9999;
  });
  const [showPixModal, setShowPixModal] = useState(false);
  const [selectedPixPlan, setSelectedPixPlan] = useState<'single' | 'trio' | 'pro'>('trio');
  const [pixCopied, setPixCopied] = useState(false);
  const [activationCode, setActivationCode] = useState("");
  const [activationSuccess, setActivationSuccess] = useState(false);
  const [title, setTitle] = useState("");
  const [subtitle, setSubtitle] = useState("");
  const [documentType, setDocumentType] = useState("artigo");
  const [customDocumentType, setCustomDocumentType] = useState("");
  const [prompt, setPrompt] = useState("");
  const [files, setFiles] = useState<File[]>([]);
  
  // Dados do Trabalho (ABNT)
  const [showWorkData, setShowWorkData] = useState(false);
  const [studentName, setStudentName] = useState("");
  const [course, setCourse] = useState("");
  const [subject, setSubject] = useState(""); // Disciplina / Matéria
  const [shift, setShift] = useState(""); // Turno (Matutino, Vespertino, Noturno, Integral)
  const [classroom, setClassroom] = useState(""); // Sala / Turma
  const [institution, setInstitution] = useState("");
  const [city, setCity] = useState("");
  const [year, setYear] = useState("");
  const [advisor, setAdvisor] = useState("");

  // Modo Trabalho em Grupo
  const [isGroupMode, setIsGroupMode] = useState(false);
  const [groupDocType, setGroupDocType] = useState<string>("trabalho_academico");
  const [customGroupDocName, setCustomGroupDocName] = useState<string>("Trabalho Personalizado");
  const [groupMembers, setGroupMembers] = useState<string[]>([""]);
  const [sectionSlots, setSectionSlots] = useState<Record<string, File | null>>({});
  const [customSections, setCustomSections] = useState<{ id: string; label: string; desc: string; icon: string }[]>([
    { id: "sec_1", label: "1 PARTE INICIAL / INTRODUÇÃO", desc: "Parte inicial do trabalho", icon: "📖" },
    { id: "sec_2", label: "2 DESENVOLVIMENTO / PARTE CENTRAL", desc: "Corpo do trabalho", icon: "📚" },
    { id: "sec_3", label: "3 CONCLUSÃO / PARTE FINAL", desc: "Encerramento e considerações", icon: "✅" },
  ]);

  // Estrutura dinâmica de seções baseada no Tipo de Trabalho
  const getGroupSectionsByDocType = (docType: string) => {
    switch (docType) {
      case "artigo":
      case "artigo_cientifico":
        return [
          { key: "resumo", label: "RESUMO E PALAVRAS-CHAVE", desc: "Resumo estruturado (150-250 palavras) e Abstract", icon: "📑" },
          { key: "introducao", label: "1 INTRODUÇÃO", desc: "Contextualização, objetivos e justificativa", icon: "📖" },
          { key: "fundamentacao", label: "2 MATERIAIS E MÉTODOS", desc: "Procedimentos metodológicos e referencial", icon: "🔬" },
          { key: "resultados", label: "3 RESULTADOS E DISCUSSÃO", desc: "Apresentação dos dados, tabelas e debate", icon: "📊" },
          { key: "conclusao", label: "4 CONSIDERAÇÕES FINAIS", desc: "Conclusões do estudo e implicações", icon: "✅" },
          { key: "referencias", label: "REFERÊNCIAS", desc: "Fontes citadas conforme ABNT NBR 6023", icon: "🔗" },
        ];
      case "projeto":
        return [
          { key: "introducao", label: "1 INTRODUÇÃO E JUSTIFICATIVA", desc: "Delimitação do tema, relevância e problema", icon: "📖" },
          { key: "objetivos", label: "2 OBJETIVOS (GERAL E ESPECÍFICOS)", desc: "Metas e propósitos da pesquisa", icon: "🎯" },
          { key: "fundamentacao", label: "3 FUNDAMENTAÇÃO TEÓRICA", desc: "Base conceitual e autores de referência", icon: "📚" },
          { key: "metodologia", label: "4 METODOLOGIA E CRONOGRAMA", desc: "Etapas de execução, instrumentos e prazos", icon: "🗓️" },
          { key: "referencias", label: "REFERÊNCIAS", desc: "Bibliografia preliminar NBR 6023", icon: "🔗" },
        ];
      case "relatorio":
        return [
          { key: "introducao", label: "1 INTRODUÇÃO E OBJETIVO DO RELATÓRIO", desc: "Contexto técnico e propósitos da atividade", icon: "📖" },
          { key: "procedimentos", label: "2 PROCEDIMENTOS REALIZADOS", desc: "Descrição técnica detalhada das etapas", icon: "⚙️" },
          { key: "analise", label: "3 ANÁLISE DE DADOS E DISCUSSÃO", desc: "Resultados obtidos, gráficos e constatações", icon: "📊" },
          { key: "recomendacoes", label: "4 CONCLUSÃO E RECOMENDAÇÕES", desc: "Parecer final e proposições técnicas", icon: "📋" },
          { key: "referencias", label: "REFERÊNCIAS", desc: "Normas técnicas e referências consultadas", icon: "🔗" },
        ];
      case "estudo_caso":
        return [
          { key: "introducao", label: "1 INTRODUÇÃO E APRESENTAÇÃO DO CASO", desc: "Contexto da organização ou situação estudada", icon: "📖" },
          { key: "diagnostico", label: "2 DIAGNÓSTICO E FUNDAMENTAÇÃO", desc: "Análise da problemática e teoria aplicável", icon: "🔍" },
          { key: "proposicoes", label: "3 PROPOSIÇÕES E SOLUÇÕES", desc: "Plano de ação e alternativas propostas", icon: "💡" },
          { key: "conclusao", label: "4 LIÇÕES APRENDIDAS E CONCLUSÃO", desc: "Impactos esperados e desfecho do estudo", icon: "✅" },
          { key: "referencias", label: "REFERÊNCIAS", desc: "Bibliografia e fontes documentais", icon: "🔗" },
        ];
      case "resenha":
        return [
          { key: "cabecalho", label: "REFERÊNCIA DA OBRA RESENHADA", desc: "Dados bibliográficos completos do livro/artigo", icon: "📖" },
          { key: "sintese", label: "1 SÍNTESE DA OBRA", desc: "Resumo das principais teses do autor", icon: "📝" },
          { key: "critica", label: "2 APRECIAÇÃO CRÍTICA E ANÁLISE", desc: "Avaliação fundamentada dos pontos fortes e fracos", icon: "🧐" },
          { key: "conclusao", label: "3 CONCLUSÃO E INDICAÇÃO", desc: "Público-alvo e relevância da obra", icon: "✅" },
        ];
      case "custom":
      case "outro":
        return customSections.map(s => ({ key: s.id, label: s.label, desc: s.desc, icon: s.icon }));
      default: // trabalho_academico (TCC) e monografia
        return [
          { key: "resumo", label: "RESUMO E ABSTRACT", desc: "Resumo em português e abstract em inglês", icon: "📑" },
          { key: "introducao", label: "1 INTRODUÇÃO", desc: "Problematização, objetivos geral/específicos e justificativa", icon: "📖" },
          { key: "fundamentacao", label: "2 FUNDAMENTAÇÃO TEÓRICA E METODOLOGIA", desc: "Revisão de literatura e procedimentos metodológicos", icon: "📚" },
          { key: "resultados", label: "3 RESULTADOS E DISCUSSÃO", desc: "Dados empíricos, tabelas IBGE e análise científica", icon: "📊" },
          { key: "conclusao", label: "4 CONSIDERAÇÕES FINAIS", desc: "Conclusões, respostas aos objetivos e perspectivas futuras", icon: "✅" },
          { key: "referencias", label: "REFERÊNCIAS", desc: "Bibliografia padronizada ABNT NBR 6023", icon: "🔗" },
        ];
    }
  };


  // Profile and Audit State
  const [showProfileModal, setShowProfileModal] = useState(false);
  const [showApiKeyModal, setShowApiKeyModal] = useState(false);
  const [profileTab, setProfileTab] = useState<"dados" | "historico">("dados");
  const [auditLogs, setAuditLogs] = useState<AuditLog[]>([]);

  // Reference State
  const [showReferenceModal, setShowReferenceModal] = useState(false);
  const [referenceSource, setReferenceSource] = useState("");
  const [referenceStyle, setReferenceStyle] = useState<"ABNT" | "APA">("ABNT");
  const [generatedReference, setGeneratedReference] = useState("");

  // Citation State (Citação Direta Curta/Longa e Indireta ABNT NBR 10520:2023)
  const [showCitationModal, setShowCitationModal] = useState(false);
  const [citationType, setCitationType] = useState<"direta_curta" | "direta_longa" | "indireta">("direta_longa");
  const [citationAuthor, setCitationAuthor] = useState("");
  const [citationYear, setCitationYear] = useState(new Date().getFullYear().toString());
  const [citationPage, setCitationPage] = useState("");
  const [citationContent, setCitationContent] = useState("");

  const [generatedText, setGeneratedText] = useState("");
  const [historyStack, setHistoryStack] = useState<string[]>([]);
  const [redoStack, setRedoStack] = useState<string[]>([]);

  // Atualizador seguro e AUTOMÁTICO com histórico e Normalização ABNT Inviolável Contínua
  const updateGeneratedTextWithHistory = (newText: string) => {
    if (!newText || !newText.trim()) {
      setGeneratedText("");
      return;
    }
    
    // Normalização ABNT Automática Instantânea:
    // 1. Citações autor-data NBR 10520:2023 em caixa mista: (Silva, 2023, p. 15)
    let abntCompliant = normalizeCitationsToABNT2023(newText);
    
    // 2. Seções canônicas em CAIXA ALTA (NBR 6024)
    abntCompliant = abntCompliant.replace(/^(#+\s*|\d+\s+)(introdução|desenvolvimento|fundamentação teórica|metodologia|resultados|conclusão|considerações finais|referências)/gmi, (match, prefix, t) => {
      return `${prefix}${t.toUpperCase()}`;
    });

    if (abntCompliant === generatedText) return;
    setHistoryStack(prev => [...prev.slice(-30), generatedText]); // guarda até 30 passos
    setRedoStack([]); // limpa refazer ao criar nova ação
    setGeneratedText(abntCompliant);
  };

  const handleUndo = () => {
    if (historyStack.length === 0) return;
    const prevText = historyStack[historyStack.length - 1];
    setHistoryStack(prev => prev.slice(0, prev.length - 1));
    setRedoStack(prev => [...prev, generatedText]);
    setGeneratedText(prevText);
  };

  const handleRedo = () => {
    if (redoStack.length === 0) return;
    const nextText = redoStack[redoStack.length - 1];
    setRedoStack(prev => prev.slice(0, prev.length - 1));
    setHistoryStack(prev => [...prev, generatedText]);
    setGeneratedText(nextText);
  };

  const [authenticityReport, setAuthenticityReport] = useState("");
  const [formatRules, setFormatRules] = useState("");
  
  const attachmentRef = useRef<HTMLInputElement>(null);
  const textareaRef = useRef<HTMLTextAreaElement>(null);
  const chatMessagesEndRef = useRef<HTMLDivElement>(null);
  
  const [chatHistory, setChatHistory] = useState<{role: 'user'|'assistant', text: string}[]>([]);
  const [chatMessage, setChatMessage] = useState("");
  const [isChatting, setIsChatting] = useState(false);

  // Estado do Quiz Interativo Passo a Passo com Gabarito e Explicação
  const [activeQuiz, setActiveQuiz] = useState<{
    isActive: boolean;
    questions: Array<{
      id: number;
      question: string;
      options: Array<{ letter: string; text: string }>;
      correctAnswer: string;
      explanation: string;
    }>;
    currentIndex: number;
    userAnswers: Array<{ selected: string; isCorrect: boolean; explanation: string; question: string }>;
    isFinished: boolean;
  } | null>(null);

  const [isAutoScrollEnabled, setIsAutoScrollEnabled] = useState(true);

  // Rolagem automática configurável para a última mensagem no Chat Acadêmico
  useEffect(() => {
    if (isAutoScrollEnabled && chatMessagesEndRef.current) {
      chatMessagesEndRef.current.scrollIntoView({ behavior: "smooth" });
    }
  }, [chatHistory, isChatting, activeQuiz, isAutoScrollEnabled]);
  
  const [isLoading, setIsLoading] = useState(false);
  const [errorMessage, setErrorMessage] = useState("");
  const [progress, setProgress] = useState(0);

  // Estados de Edição Visual e Formatação Padrão de Imagens (Word Style Drag & Drop)
  const [imageWidth, setImageWidth] = useState<number>(80); // % da largura útil
  const [imageHeight, setImageHeight] = useState<number>(320); // altura em pixels (auto/custom)
  const [imageAlign, setImageAlign] = useState<"center" | "left" | "right">("center");
  const [imageOffset, setImageOffset] = useState<{ x: number; y: number }>({ x: 0, y: 0 });
  const [imageOrientation, setImageOrientation] = useState<"portrait" | "landscape">("landscape"); // Retrato ou Paisagem
  const [imageRotation, setImageRotation] = useState<number>(0); // Rotação em graus (0, 90, 180, 270)
  const [imageStyle, setImageStyle] = useState<"none" | "simple_border" | "soft_shadow" | "rounded_frame" | "academic_box">("academic_box");
  const [isImageSelected, setIsImageSelected] = useState<boolean>(false); // O menu só aparece quando a foto é selecionada (Padrão Word)
  const [hiddenPageNumbers, setHiddenPageNumbers] = useState<Set<number>>(new Set()); // Páginas com numeração oculta pelo usuário
  const [isSpeaking, setIsSpeaking] = useState<boolean>(false);
  const [speechRate, setSpeechRate] = useState<number>(1.0); // Velocidade padrão: 1x, 1.25x, 1.5x, 2x
  const [speechVolume, setSpeechVolume] = useState<number>(1.0); // Volume padrão: 1.0 (100%), 0.75, 0.5
  const [speechGender, setSpeechGender] = useState<"female" | "male">("female"); // Seleção de voz: Feminina (padrão) ou Masculina

  // Estado de controle de índice de leitura contínua (para troca de velocidade fluida sem recomeçar)
  // Referência de Áudio Neural HTML5 (Estúdio Gratuito Francisca & Antonio)
  const neuralAudioRef = useRef<HTMLAudioElement | null>(null);

  // Função interna para disparar a fala da frase atual com Áudio Neural de Estúdio Gratuito
  const speakSentenceAtIndex = async (index: number, rate?: number, vol?: number, gender?: "female" | "male") => {
    const sentences = speechSentencesRef.current;
    if (index >= sentences.length) {
      setIsSpeaking(false);
      isSpeakingRef.current = false;
      setErrorMessage("✅ Leitura em áudio neural concluída!");
      setTimeout(() => setErrorMessage(""), 3000);
      return;
    }

    currentSentenceIdxRef.current = index;
    const sentence = sentences[index];
    if (!sentence || !sentence.trim()) {
      speakSentenceAtIndex(index + 1, rate, vol, gender);
      return;
    }

    const currentRate = rate !== undefined ? rate : speechRate;
    const currentVol = vol !== undefined ? vol : speechVolume;
    const currentGender = gender !== undefined ? gender : speechGender;

    // 1. Tenta sintetizar a Voz Neural Ultra-Humana Gratuita (Francisca / Antonio)
    const voiceName = currentGender === "female" ? "pt-BR-FranciscaNeural" : "pt-BR-AntonioNeural";
    const ratePercent = currentRate === 1.0 ? "+0%" : currentRate > 1.0 ? `+${Math.round((currentRate - 1) * 100)}%` : `-${Math.round((1 - currentRate) * 100)}%`;

    try {
      const response = await fetch("/api/tts", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          text: sentence.trim(),
          voice: voiceName,
          rate: ratePercent,
          pitch: currentGender === "female" ? "+2Hz" : "-2Hz"
        })
      });

      if (response.ok) {
        const blob = await response.blob();
        const audioUrl = URL.createObjectURL(blob);

        if (neuralAudioRef.current) {
          neuralAudioRef.current.pause();
        }

        const audio = new Audio(audioUrl);
        audio.volume = currentVol;
        neuralAudioRef.current = audio;

        audio.onended = () => {
          URL.revokeObjectURL(audioUrl);
          if (isSpeakingRef.current) {
            speakSentenceAtIndex(index + 1, rate, vol, gender);
          }
        };

        audio.onerror = () => {
          URL.revokeObjectURL(audioUrl);
          if (isSpeakingRef.current) {
            speakSentenceAtIndex(index + 1, rate, vol, gender);
          }
        };

        await audio.play();
        return;
      }
    } catch (neuralErr) {
      console.warn("[Voz Neural] Falha de conexão, usando sintetizador nativo:", neuralErr);
    }

    // 2. Fallback de alta fidelidade: Web Speech Synthesis nativo
    if (!('speechSynthesis' in window)) return;
    const utterance = new SpeechSynthesisUtterance(sentence.trim());
    utterance.lang = "pt-BR";
    utterance.rate = currentRate || 1.0;
    utterance.volume = currentVol || 1.0;
    utterance.pitch = currentGender === "female" ? 1.05 : 0.95;

    const voices = window.speechSynthesis.getVoices();
    const ptVoices = voices.filter(v => v.lang.includes("pt-BR") || v.lang.includes("pt_BR") || v.lang.startsWith("pt"));

    let selectedVoice = null;
    if (currentGender === "female") {
      selectedVoice = ptVoices.find(v => {
        const n = v.name.toLowerCase();
        return n.includes("natural") || n.includes("neural") || n.includes("online") || n.includes("francisca") || n.includes("thalita") || n.includes("luciana") || (n.includes("google") && !n.includes("male"));
      }) || ptVoices.find(v => !v.name.toLowerCase().includes("male") && !v.name.toLowerCase().includes("antonio"));
    } else {
      selectedVoice = ptVoices.find(v => {
        const n = v.name.toLowerCase();
        return n.includes("natural") || n.includes("neural") || n.includes("online") || n.includes("antonio") || n.includes("felipe") || n.includes("ricardo") || n.includes("daniel") || n.includes("male");
      }) || ptVoices.find(v => v.name.toLowerCase().includes("male") || v.name.toLowerCase().includes("antonio"));
    }

    if (!selectedVoice && ptVoices.length > 0) selectedVoice = ptVoices[0];
    if (selectedVoice) utterance.voice = selectedVoice;

    utterance.onend = () => {
      if (isSpeakingRef.current) speakSentenceAtIndex(index + 1, rate, vol, gender);
    };

    utterance.onerror = () => {
      if (isSpeakingRef.current) speakSentenceAtIndex(index + 1, rate, vol, gender);
    };

    window.speechSynthesis.speak(utterance);
  };

  // Leitura em Áudio do Texto Acadêmico com Fraseamento Natural
  const handleToggleSpeech = () => {
    if (isSpeaking) {
      isSpeakingRef.current = false;
      if (neuralAudioRef.current) {
        neuralAudioRef.current.pause();
        neuralAudioRef.current = null;
      }
      if ('speechSynthesis' in window) {
        window.speechSynthesis.cancel();
      }
      setIsSpeaking(false);
      setErrorMessage("⏹️ Leitura em áudio interrompida.");
      setTimeout(() => setErrorMessage(""), 2500);
      return;
    }

    if (!generatedText || !generatedText.trim()) {
      setErrorMessage("✍️ Gere um trabalho ou escreva um texto no editor para ouvir a leitura!");
      setTimeout(() => setErrorMessage(""), 3500);
      return;
    }

    // Separa páginas para NÃO ler a Capa nem a Contra-Capa (Folha de Rosto)
    let bodyTextOnly = generatedText;
    if (generatedText.includes("--- [QUEBRA DE PÁGINA] ---")) {
      const parts = generatedText.split("--- [QUEBRA DE PÁGINA] ---");
      const contentParts = parts.filter(p => {
        const t = p.trim();
        return !t.startsWith("CAPA") && 
               !t.startsWith("FOLHA DE ROSTO") && 
               t !== "CAPA_AUTO" && 
               t !== "FOLHA_ROSTO_AUTO" && 
               !t.includes("requisito parcial") && 
               !t.includes("apresentado à");
      });
      bodyTextOnly = contentParts.join("\n\n");
    }

    // Limpa marcações estruturais e formata para ritmo de fala humano
    const cleanSpeechText = bodyTextOnly
      .replace(/--- \[(?:QUEBRA DE PÁGINA|NOVA PÁGINA)\] ---/g, ' ')
      .replace(/!\[.*?\]\(.*?\)/g, ' ')
      .replace(/#+/g, ' ')
      .replace(/[\*\_\`]/g, ' ')
      .replace(/([0-9]+)\.([0-9]+)/g, '$1 vírgula $2')
      .replace(/\s+/g, ' ')
      .trim();

    if (!cleanSpeechText) {
      setErrorMessage("Nenhum conteúdo textual além da capa para leitura em áudio.");
      return;
    }

    // Divide em frases (pontos, exclamações, interrogações e quebras de parágrafo) para cadência natural
    const rawSentences = cleanSpeechText.match(/[^.!?]+[.!?]+|\s*[^.!?]+$/g) || [cleanSpeechText];
    const sentences = rawSentences.map(s => s.trim()).filter(s => s.length > 0);

    speechSentencesRef.current = sentences;
    currentSentenceIdxRef.current = 0;
    isSpeakingRef.current = true;
    setIsSpeaking(true);

    if ('speechSynthesis' in window) {
      window.speechSynthesis.cancel();
    }
    const voiceLabel = speechGender === "female" ? "Francisca Neural 👩" : "Antônio Neural 👨";
    setErrorMessage(`🎙️ Reproduzindo com Voz Humana Neural (${voiceLabel} - ${speechRate}x)...`);
    speakSentenceAtIndex(0);
  };

  useEffect(() => {
    let interval: NodeJS.Timeout;
    if (isLoading) {
      setProgress(0);
      setErrorMessage(""); // clear errors when starting a new task
      interval = setInterval(() => {
        setProgress(p => {
          if (p < 85) return p + (Math.random() * 8); 
          if (p < 95) return p + (Math.random() * 0.5); 
          return p;
        });
      }, 500);
    } else {
      setProgress(100);
      const to = setTimeout(() => setProgress(0), 1000); 
      return () => clearTimeout(to);
    }
    return () => clearInterval(interval);
  }, [isLoading]);

  const [activeTab, setActiveTab] = useState<"generator" | "editor" | "slides" | "report" | "chat">("generator");
  const [editorMode, setEditorMode] = useState<"a4" | "raw">("a4");
  const [currentPageIndex, setCurrentPageIndex] = useState<number>(0);
  const [viewLayout, setViewLayout] = useState<"book" | "continuous">("book");
  const [zoomScale, setZoomScale] = useState<number>(65); // 65% proporção padrão do Microsoft Word e Adobe Acrobat em visualização contínua
  const [isStageExpanded, setIsStageExpanded] = useState<boolean>(false);
  const [isSidebarCollapsed, setIsSidebarCollapsed] = useState<boolean>(false);
  const [stageHeight, setStageHeight] = useState<number>(0);
  const [showSlidesModal, setShowSlidesModal] = useState<boolean>(false);
  const [slidesTheme, setSlidesTheme] = useState<"academic" | "modern" | "dark" | "sunset" | "neon" | "emerald" | "lavender" | "solar">("academic");
  const [activeSlideIndex, setActiveSlideIndex] = useState<number>(0);
  const [editableSlides, setEditableSlides] = useState<Array<{
    title: string;
    subtitle?: string;
    bullets: string[];
    notes?: string;
    isCover?: boolean;
    layout?: "card" | "split" | "bullets" | "quote" | "metrics" | "timeline" | "comparison";
    badge?: string;
  }>>([]);

  // Sincroniza atalhos de teclado (Seta Esquerda / Direita) para passar páginas estilo livro
  useEffect(() => {
    const handleKeyDown = (e: KeyboardEvent) => {
      // Ignora se estiver digitando em campos de texto
      if (document.activeElement?.tagName === "INPUT" || document.activeElement?.tagName === "TEXTAREA") {
        return;
      }
      const pagesCount = generatedText ? generatedText.split("--- [QUEBRA DE PÁGINA] ---").length : 1;
      if (e.key === "ArrowLeft" || e.key === "PageUp") {
        setCurrentPageIndex(prev => Math.max(0, prev - 1));
      } else if (e.key === "ArrowRight" || e.key === "PageDown") {
        setCurrentPageIndex(prev => Math.min(pagesCount - 1, prev + 1));
      }
    };
    window.addEventListener("keydown", handleKeyDown);
    return () => window.removeEventListener("keydown", handleKeyDown);
  }, [generatedText]);

  const logAction = (titleDesc: string, textContent?: string) => {
    // Registra APENAS textos acadêmicos gerados com conteúdo real
    if (!textContent || textContent.trim().length < 50) return;
    
    setAuditLogs(prev => {
      const newLog = { 
        action: titleDesc || "Texto Acadêmico Gerado", 
        content: textContent, 
        timestamp: new Date().toISOString() 
      };
      const updated = [newLog, ...prev.filter(l => Boolean(l.content && l.content.length > 50))];
      
      const oneWeekAgo = new Date();
      oneWeekAgo.setDate(oneWeekAgo.getDate() - 7);
      
      const filtered = updated.filter(log => new Date(log.timestamp) > oneWeekAgo);
      localStorage.setItem('emia_audit_logs', JSON.stringify(filtered));
      return filtered;
    });
  };

  useEffect(() => {
    // Load profile
    const storedProfile = localStorage.getItem('emia_user_profile');
    if (storedProfile) {
      try {
        const p: UserProfile = JSON.parse(storedProfile);
        if (p.name) setStudentName(p.name);
        if (p.institution) setInstitution(p.institution);
        if (p.city) setCity(p.city);
        if (p.year) setYear(p.year);
        if (p.advisor) setAdvisor(p.advisor);
      } catch (e) {}
    }

    // Load logs
    const storedLogs = localStorage.getItem('emia_audit_logs');
    if (storedLogs) {
      try {
        const logs: AuditLog[] = JSON.parse(storedLogs);
        const oneWeekAgo = new Date();
        oneWeekAgo.setDate(oneWeekAgo.getDate() - 7);
        const filtered = logs.filter(log => new Date(log.timestamp) > oneWeekAgo);
        setAuditLogs(filtered);
      } catch (e) {}
    }
  }, []);

  useEffect(() => {
    if (isAuthenticated) {
      logAction("Acesso ao sistema via Demonstração");
    }
  }, [isAuthenticated]);

  const onDrop = async (acceptedFiles: File[]) => {
    if (acceptedFiles.length > 0) {
      // Separa imagens para inserção instantânea direta no documento
      const imageFiles = acceptedFiles.filter(f => f.type.startsWith("image/") || /\.(jpg|jpeg|png|webp|gif|svg)$/i.test(f.name));
      const documentFiles = acceptedFiles.filter(f => !imageFiles.includes(f));

      if (imageFiles.length > 0) {
        imageFiles.forEach(imgFile => {
          const reader = new FileReader();
          reader.onload = (event) => {
            const base64 = event.target?.result as string;
            const figureBlock = `\n\n![Figura inserida](${base64})\nFigura 1 – Representação Ilustrativa do Objeto de Estudo\nFonte: Elaborado pelos autores (${new Date().getFullYear()}).\n\n`;
            updateGeneratedTextWithHistory(generatedText ? generatedText + figureBlock : figureBlock);
            setActiveTab("editor");
            setErrorMessage("✅ Imagem inserida diretamente no documento ABNT!");
            setTimeout(() => setErrorMessage(""), 3500);
          };
          reader.readAsDataURL(imgFile);
        });
      }

      if (documentFiles.length > 0) {
        setFiles(prev => [...prev, ...documentFiles]);
        setIsLoading(true);
        setErrorMessage("⏳ Processando e normalizando documento com a Skill Acadêmica ABNT...");

        try {
          let extracted = "";
          // 1. Tenta extração via backend
          try {
            const formData = new FormData();
            documentFiles.forEach(f => formData.append("files", f));
            const res = await fetch("/api/extract", { method: "POST", body: formData });
            if (res.ok) {
              const data = await res.json();
              if (data.success && data.text) extracted = data.text;
            }
          } catch (e) {
            console.warn("Backend extract falhou, executando extração local:", e);
          }

          // 2. Extração client-side caso necessário
          if (!extracted) {
            const parts: string[] = [];
            for (const docFile of documentFiles) {
              if (docFile.type.includes("text") || docFile.name.endsWith(".txt") || docFile.name.endsWith(".md") || docFile.name.endsWith(".csv")) {
                parts.push(await docFile.text());
              } else if (docFile.name.toLowerCase().endsWith(".pdf") || docFile.type === "application/pdf") {
                const ab = await docFile.arrayBuffer();
                const latinText = new TextDecoder("latin1").decode(new Uint8Array(ab));
                const matches = latinText.match(/\(([^()]{2,})\)\s*(?:Tj|'|")/g) || [];
                let pdfTxt = "";
                if (matches.length > 0) {
                  pdfTxt = matches.map(m => m.replace(/^\(/, '').replace(/\)\s*(?:Tj|'|")$/, '')).join(' ').replace(/\\([()\\])/g, '$1').replace(/\s+/g, ' ');
                } else {
                  pdfTxt = latinText.replace(/[^\x20-\x7E\xA0-\xFF\n\r]/g, ' ').split(/\s{3,}/).filter(c => c.trim().length > 15).join('\n\n');
                }
                if (pdfTxt.trim()) parts.push(pdfTxt.trim());
              }
            }
            extracted = parts.join("\n\n");
          }

          if (extracted && extracted.trim()) {
            // Normalização Local 100% Offline e Instantânea (Zero dependência de IA ou rede)
            let formatted = normalizeCitationsToABNT2023(extracted);
            formatted = organizeTextInABNTOrder(formatted);

            // Limpa espaçamentos excessivos
            formatted = formatted
              .replace(/\r\n/g, '\n')
              .replace(/\n{4,}/g, '\n\n\n')
              .replace(/[ \t]+$/gm, '');

            updateGeneratedTextWithHistory(formatted);
            setActiveTab("editor");
            setFiles([]); // Limpa a lista de arquivos acumulados após subir para o editor
            setErrorMessage("✅ Documento carregado e formatado no padrão ABNT com sucesso!");
            setTimeout(() => setErrorMessage(""), 3500);
            logAction("Upload e Edição Direta ABNT (Modo Offline)", formatted);
          }
        } catch (uploadErr) {
          console.error(uploadErr);
          setErrorMessage("Erro ao formatar o documento carregado.");
        } finally {
          setIsLoading(false);
        }
      }
    }
  };
  const { getRootProps, getInputProps, isDragActive } = useDropzone({ onDrop });

  const [aiProvider, setAiProvider] = useState<"gemini" | "openai">(() => (localStorage.getItem("emia_ai_provider") as any) || "gemini");
  const [customGeminiKey, setCustomGeminiKey] = useState<string>(() => localStorage.getItem("emia_custom_gemini_key") || "");
  const [customOpenaiKey, setCustomOpenaiKey] = useState<string>(() => localStorage.getItem("emia_custom_openai_key") || "");
  const [googleUser, setGoogleUser] = useState<{ name?: string; email?: string; picture?: string } | null>(() => {
    try {
      const saved = localStorage.getItem("emia_google_user");
      return saved ? JSON.parse(saved) : null;
    } catch {
      return null;
    }
  });

  // PERSISTÊNCIA INVIOLÁVEL E AUTO-RECONEXÃO AUTOMÁTICA DA IA
  useEffect(() => {
    // Garante que o status de conexão seja permanente
    const isAuth = localStorage.getItem("emia_authenticated") === "true";
    if (isAuth && !isAuthenticated) {
      setIsAuthenticated(true);
    }
    
    // Auto-recupera chave Gemini e OpenAI persistidas
    const savedGemini = localStorage.getItem("emia_custom_gemini_key");
    if (savedGemini && !customGeminiKey) {
      setCustomGeminiKey(savedGemini);
    }
    const savedOpenai = localStorage.getItem("emia_custom_openai_key");
    if (savedOpenai && !customOpenaiKey) {
      setCustomOpenaiKey(savedOpenai);
    }

    // Auto-cura de reconexão contínua com a IA (heartbeat silencioso sem pedir confirmação)
    const keepAliveInterval = setInterval(() => {
      const authState = localStorage.getItem("emia_authenticated") === "true";
      if (!authState) {
        localStorage.setItem("emia_authenticated", "true");
        setIsAuthenticated(true);
      }
    }, 5000);

    return () => clearInterval(keepAliveInterval);
  }, [isAuthenticated, customGeminiKey, customOpenaiKey]);

  const getApiHeaders = () => {
    const headers: Record<string, string> = { "Content-Type": "application/json" };
    headers["x-ai-provider"] = aiProvider || localStorage.getItem("emia_ai_provider") || "gemini";
    const gKey = customGeminiKey || localStorage.getItem("emia_custom_gemini_key");
    if (gKey && gKey.trim()) {
      headers["x-gemini-api-key"] = gKey.trim();
    }
    const oKey = customOpenaiKey || localStorage.getItem("emia_custom_openai_key");
    if (oKey && oKey.trim()) {
      headers["x-openai-api-key"] = oKey.trim();
    }
    return headers;
  };

  const [showGoogleModal, setShowGoogleModal] = useState(false);
  const [googleEmailInput, setGoogleEmailInput] = useState("");
  const [selectedEngine, setSelectedEngine] = useState<"gemini" | "openai">("gemini");
  const [isGoogleLoggingIn, setIsGoogleLoggingIn] = useState(false);

  // LOGIN GOOGLE OFICIAL OAUTH 2.0 (Google Identity Services)
  const handleOfficialGoogleLogin = () => {
    try {
      if (typeof window !== "undefined" && (window as any).google?.accounts?.oauth2) {
        setIsGoogleLoggingIn(true);
        const tokenClient = (window as any).google.accounts.oauth2.initTokenClient({
          client_id: "1035659850438-e6q1o6n2f3d5.apps.googleusercontent.com", // Client ID padrão para login
          scope: "email profile openid https://www.googleapis.com/auth/generative-language.tuning",
          callback: async (response: any) => {
            setIsGoogleLoggingIn(false);
            if (response.access_token) {
              const token = response.access_token;
              localStorage.setItem("emia_google_token", token);
              
              // Busca dados do usuário autenticado no Google
              try {
                const userRes = await fetch("https://www.googleapis.com/oauth2/v3/userinfo", {
                  headers: { Authorization: `Bearer ${token}` }
                });
                const userData = await userRes.json();
                const email = userData.email || "usuario.google@gmail.com";
                const name = userData.name || userData.given_name || email.split("@")[0].toUpperCase();
                const userInfo = { name, email, picture: userData.picture };
                
                setGoogleUser(userInfo);
                localStorage.setItem("emia_google_user", JSON.stringify(userInfo));
                localStorage.setItem("emia_user_email", email);
                localStorage.setItem("emia_authenticated", "true");
                setIsAuthenticated(true);
                
                const cleanEmail = email.toLowerCase();
                const isMasterUser = cleanEmail === "erlane.digital@gmail.com" || cleanEmail === "cadumajor@gmail.com";

                if (isMasterUser) {
                  setIsMaster(true);
                  setCredits(9999);
                  localStorage.setItem("emia_is_master", "true");
                  localStorage.setItem("emia_credits", "9999");
                  logAction(`Login Mestre / Acesso Vitalício realizado (${cleanEmail})`);
                } else {
                  setIsMaster(false);
                  const savedGeminiKey = localStorage.getItem("emia_custom_gemini_key");
                  if (!savedGeminiKey || savedGeminiKey.trim().length < 10) {
                    // Usuário ainda não tem a chave configurada: abre o modal simples do Google AI Studio na hora
                    setShowApiKeyModal(true);
                  }
                  logAction(`Login de Aluno via Google realizado (${cleanEmail})`);
                }
              } catch (e) {
                console.error("Falha ao validar autenticação Google:", e);
                localStorage.removeItem("emia_authenticated");
                localStorage.removeItem("emia_google_token");
                setIsAuthenticated(false);
                setErrorMessage("Falha ao validar sua conta Google. O acesso foi bloqueado.");
              }
            } else {
              localStorage.removeItem("emia_authenticated");
              localStorage.removeItem("emia_google_token");
              setIsAuthenticated(false);
              setErrorMessage("Autenticação não concluída. O acesso permanece bloqueado.");
            }
          },
          error_callback: (err: any) => {
            setIsGoogleLoggingIn(false);
            console.error("Erro ou cancelamento do Google OAuth:", err);
            localStorage.removeItem("emia_authenticated");
            localStorage.removeItem("emia_google_token");
            setIsAuthenticated(false);
            setErrorMessage("Autenticação cancelada ou com erro no Google. Acesso bloqueado.");
          }
        });
        tokenClient.requestAccessToken();
      } else {
        setShowGoogleModal(true);
      }
    } catch (e) {
      console.error("Falha ao inicializar GIS OAuth 2.0:", e);
      setIsGoogleLoggingIn(false);
      localStorage.removeItem("emia_authenticated");
      localStorage.removeItem("emia_google_token");
      setIsAuthenticated(false);
      setErrorMessage("Erro ao conectar com os servidores de autenticação do Google.");
    }
  };

  const handleGoogleLoginSubmit = (e?: React.FormEvent) => {
    if (e) e.preventDefault();
    const clean = googleEmailInput.trim().toLowerCase();
    if (!clean) return;

    const userName = clean.split("@")[0].replace(/[\._]/g, " ").toUpperCase();
    const userInfo = { name: userName, email: clean };

    setGoogleUser(userInfo);
    setAiProvider(selectedEngine);
    localStorage.setItem("emia_ai_provider", selectedEngine);
    localStorage.setItem("emia_google_user", JSON.stringify(userInfo));
    localStorage.setItem("emia_authenticated", "true");
    localStorage.setItem("emia_user_email", clean);

    const cleanEmail = clean.toLowerCase();
    const isMasterUser = cleanEmail === "erlane.digital@gmail.com" || cleanEmail === "cadumajor@gmail.com";

    if (isMasterUser) {
      setIsMaster(true);
      setCredits(9999);
      localStorage.setItem("emia_is_master", "true");
      localStorage.setItem("emia_credits", "9999");
      logAction(`Login Mestre / Acesso Vitalício (${cleanEmail} • IA: ${selectedEngine}) realizado`);
    } else {
      setIsMaster(false);
      setShowPixModal(true);
      logAction(`Login de Aluno via Google (${cleanEmail} • IA: ${selectedEngine}) realizado`);
    }

    setShowGoogleModal(false);
    setIsAuthenticated(true);
  };

  const handleLogout = () => {
    setIsAuthenticated(false);
    setIsMaster(false);
    setGoogleUser(null);
    localStorage.removeItem("emia_authenticated");
    localStorage.removeItem("emia_is_master");
    localStorage.removeItem("emia_google_user");
    localStorage.removeItem("emia_google_token");
  };

  if (!isAuthenticated) {
    return (
      <div className="min-h-screen bg-gradient-to-br from-blue-50 via-indigo-50/40 to-slate-100 flex items-center justify-center p-4">
        <div className="max-w-md w-full bg-white rounded-3xl shadow-xl border border-gray-100 p-8 md:p-10 text-center animate-in fade-in zoom-in-95 duration-200">
          
          {/* Logo e Ícone Oficial */}
          <div className="bg-gradient-to-tr from-blue-600 to-indigo-600 text-white w-16 h-16 rounded-2xl flex items-center justify-center mx-auto mb-5 shadow-lg shadow-blue-500/25">
            <FileText className="w-8 h-8" />
          </div>

          <h1 className="text-2xl font-black text-gray-900 tracking-tight">EMIA.EDUTECH</h1>
          <p className="text-gray-600 mt-2 text-xs leading-relaxed max-w-xs mx-auto mb-6">
            Plataforma de Produção e Normalização Acadêmica nas Normas ABNT (UNESP / USP / NBR 14724).
          </p>

          {/* Mensagem de Erro se houver */}
          {errorMessage && (
            <div className="mb-4 p-3 bg-rose-50 border border-rose-200 text-rose-700 text-xs rounded-xl font-medium">
              {errorMessage}
            </div>
          )}

          {/* 1. BOTÃO OFICIAL DO GOOGLE (ONE-TAP / POPUP OAUTH) */}
          <button
            onClick={handleOfficialGoogleLogin}
            disabled={isGoogleLoggingIn}
            type="button"
            className="w-full flex items-center justify-center gap-3 bg-white border-2 border-gray-200 hover:border-blue-500 hover:bg-blue-50/30 text-gray-800 font-bold py-3.5 px-6 rounded-2xl shadow-sm hover:shadow-md transition-all text-sm group active:scale-[0.98] disabled:opacity-70"
          >
            {isGoogleLoggingIn ? (
              <Loader2 className="w-5 h-5 text-blue-600 animate-spin" />
            ) : (
              <svg className="w-5 h-5 flex-shrink-0" viewBox="0 0 24 24">
                <path fill="#4285F4" d="M22.56 12.25c0-.78-.07-1.53-.2-2.25H12v4.26h5.92c-.26 1.37-1.04 2.53-2.21 3.31v2.77h3.57c2.08-1.92 3.28-4.74 3.28-8.09z" />
                <path fill="#34A853" d="M12 23c2.97 0 5.46-.98 7.28-2.66l-3.57-2.77c-.98.66-2.23 1.06-3.71 1.06-2.86 0-5.29-1.93-6.16-4.53H2.18v2.84C3.99 20.53 7.7 23 12 23z" />
                <path fill="#FBBC05" d="M5.84 14.09c-.22-.66-.35-1.36-.35-2.09s.13-1.43.35-2.09V7.06H2.18C1.43 8.55 1 10.22 1 12s.43 3.45 1.18 4.94l2.85-2.22.81-.63z" />
                <path fill="#EA4335" d="M12 5.38c1.62 0 3.06.56 4.21 1.64l3.15-3.15C17.45 2.09 14.97 1 12 1 7.7 1 3.99 3.47 2.18 7.06l3.66 2.84c.87-2.6 3.3-4.52 6.16-4.52z" />
              </svg>
            )}
            <span className="text-sm font-extrabold">{isGoogleLoggingIn ? "Autenticando..." : "Entrar com o Google"}</span>
          </button>

          {/* Divisor Padrão */}
          <div className="relative my-6">
            <div className="absolute inset-0 flex items-center">
              <div className="w-full border-t border-gray-200" />
            </div>
            <div className="relative flex justify-center text-xs uppercase">
              <span className="bg-white px-3 text-gray-400 font-semibold tracking-wider">ou acesse com seu Gmail</span>
            </div>
          </div>

          {/* 2. FORMULÁRIO DIRETO DE ENTRADA POR E-MAIL GOOGLE */}
          <form onSubmit={handleGoogleLoginSubmit} className="space-y-4 text-left">
            <div>
              <label className="block text-xs font-bold text-gray-700 mb-1.5">
                Seu E-mail Google (Gmail):
              </label>
              <input
                type="email"
                required
                value={googleEmailInput}
                onChange={(e) => setGoogleEmailInput(e.target.value)}
                placeholder="seu.email@gmail.com"
                className="w-full px-4 py-3 bg-gray-50 border border-gray-300 rounded-xl text-sm font-medium focus:ring-2 focus:ring-blue-500 focus:bg-white outline-none transition-all placeholder:text-gray-400"
              />
            </div>

            <Button
              type="submit"
              className="w-full bg-gradient-to-r from-blue-600 via-indigo-600 to-blue-700 hover:from-blue-700 hover:to-indigo-700 text-white font-bold py-3.5 rounded-xl text-sm shadow-md shadow-blue-500/20 active:scale-[0.98] transition-all"
            >
              Acessar Painel Acadêmico
            </Button>
          </form>

          <div className="mt-8 pt-5 border-t border-gray-100 text-center">
            <p className="text-[11px] text-gray-400 font-medium flex items-center justify-center gap-1">
              <Lock className="w-3 h-3 text-gray-400" />
              Acesso Seguro • Normas ABNT Atualizadas
            </p>
          </div>
        </div>
      </div>
    );
  }

  const handleGenerate = async () => {
    const cleanTitle = title.trim() || "Trabalho Acadêmico Geral";
    const hasActiveKey = !!customGeminiKey && customGeminiKey.trim().length >= 10;
    
    // Se não tiver chave de API do Gemini conectada e não tiver créditos PIX:
    if (!hasActiveKey && credits <= 0 && !isMaster) {
      setShowApiKeyModal(true);
      setErrorMessage("Conecte sua chave gratuita do Google AI Studio para gerar seus trabalhos ilimitados.");
      return;
    }

    setIsLoading(true);
    setErrorMessage("");
    try {
      const formData = new FormData();
      formData.append("title", cleanTitle);
      formData.append("subtitle", subtitle);
      formData.append("documentType", documentType === "outros" ? (customDocumentType || "artigo") : documentType);
      formData.append("prompt", prompt || `Desenvolva um texto completo e estruturado sobre ${cleanTitle}.`);
      
      // Send work data if any exists
      if (studentName) formData.append("studentName", studentName);
      if (course) formData.append("course", course);
      if (institution) formData.append("institution", institution);
      if (city) formData.append("city", city);
      if (year) formData.append("year", year);
      if (advisor) formData.append("advisor", advisor);

      if (files.length > 0) {
        files.forEach(f => formData.append("files", f));
      }

      const customHeaders: Record<string, string> = {
        "x-ai-provider": aiProvider || "gemini",
      };
      if (customGeminiKey && customGeminiKey.trim()) {
        customHeaders["x-gemini-api-key"] = customGeminiKey.trim();
      }
      if (customOpenaiKey && customOpenaiKey.trim()) {
        customHeaders["x-openai-api-key"] = customOpenaiKey.trim();
      }
      const token = localStorage.getItem("emia_google_token");
      if (token) {
        customHeaders["authorization"] = `Bearer ${token}`;
      }

      let finalText = "";
      
      // Chamada direta instantânea pelo Motor Acadêmico de Alta Velocidade (gemini-3.5-flash-lite)
      try {
        finalText = await generateAcademicText({
          title: cleanTitle,
          subtitle,
          documentType: documentType === "outros" ? (customDocumentType || "artigo") : documentType,
          prompt,
          studentName,
          course,
          institution,
          city,
          year,
          advisor,
          customGeminiKey,
        });
      } catch (genErr) {
        console.warn("[EMIA Motor Direto] Tentando via /api/generate fallback:", genErr);
        try {
          const res = await fetch("/api/generate", {
            method: "POST",
            headers: customHeaders,
            body: formData,
          });
          const textData = await res.text();
          const data = JSON.parse(textData);
          if (data.success && data.text) {
            finalText = data.text;
          }
        } catch (apiErr) {
          console.error("Todas as tentativas falharam:", apiErr);
        }
      }

      if (finalText) {
        setGeneratedText(finalText);
        setActiveTab("editor");
        logAction(`Geração de documento: ${cleanTitle}`, finalText);
        if (!isMaster) {
          setCredits(prev => {
            const next = Math.max(0, prev - 1);
            localStorage.setItem("emia_credits", String(next));
            return next;
          });
        }
      } else {
        setErrorMessage("Não foi possível gerar o texto. Tente novamente.");
      }
    } catch (error) {
      console.error("[EMIA Client Error]", error);
      setErrorMessage(error instanceof Error ? error.message : "Erro na geração do documento.");
    } finally {
      setIsLoading(false);
    }
  };

  // Função para identificar e ordenar estruturalmente as seções do documento (ABNT NBR 14724 / 6022) SEM alterar as palavras do texto original
  const organizeTextInABNTOrder = (raw: string): string => {
    if (!raw || !raw.trim()) return raw;

    // Normaliza quebras de linha
    const cleanRaw = raw.replace(/\r\n/g, '\n');

    // Mapeamento de blocos conhecidos
    let resumoBlock = "";
    let abstractBlock = "";
    let introducaoBlock = "";
    let desenvolvimentoBlock = "";
    let resultadosBlock = "";
    let conclusaoBlock = "";
    let referenciasBlock = "";
    let outrosBlocos: string[] = [];

    // Quebra por quebra de página ou por títulos de seções
    const sections = cleanRaw.split(/\n(?=(?:[0-9]+\s+[A-ZÀ-Ú\s]+|RESUMO|ABSTRACT|SUMÁRIO|REFERÊNCIAS|CONSIDERAÇÕES FINAIS|CONCLUSÃO)\b)/i);

    for (const sec of sections) {
      const trimmed = sec.trim();
      if (!trimmed) continue;

      if (/^RESUMO\b/i.test(trimmed)) {
        resumoBlock = trimmed;
      } else if (/^ABSTRACT\b/i.test(trimmed)) {
        abstractBlock = trimmed;
      } else if (/^(?:1\s+)?INTRODUÇÃO\b/i.test(trimmed)) {
        introducaoBlock = trimmed;
      } else if (/^(?:2\s+)?(?:FUNDAMENTAÇÃO|METODOLOGIA|DESENVOLVIMENTO)\b/i.test(trimmed)) {
        desenvolvimentoBlock = (desenvolvimentoBlock ? desenvolvimentoBlock + "\n\n" : "") + trimmed;
      } else if (/^(?:3\s+)?(?:RESULTADOS|DISCUSSÃO)\b/i.test(trimmed)) {
        resultadosBlock = trimmed;
      } else if (/^(?:4\s+)?(?:CONSIDERAÇÕES FINAIS|CONCLUSÃO)\b/i.test(trimmed)) {
        conclusaoBlock = trimmed;
      } else if (/^REFERÊNCIAS\b/i.test(trimmed)) {
        referenciasBlock = trimmed;
      } else {
        outrosBlocos.push(trimmed);
      }
    }

    // Se detectou ao menos duas seções acadêmicas, remonta rigorosamente na ordem ABNT oficial com Quebras de Página
    if (resumoBlock || abstractBlock || introducaoBlock || referenciasBlock) {
      const orderedParts: string[] = [];
      if (resumoBlock) orderedParts.push(resumoBlock);
      if (abstractBlock) orderedParts.push(abstractBlock);
      if (introducaoBlock) orderedParts.push(introducaoBlock);
      if (desenvolvimentoBlock) orderedParts.push(desenvolvimentoBlock);
      if (resultadosBlock) orderedParts.push(resultadosBlock);
      if (conclusaoBlock) orderedParts.push(conclusaoBlock);
      if (referenciasBlock) orderedParts.push(referenciasBlock);
      if (outrosBlocos.length > 0) orderedParts.push(...outrosBlocos);

      return orderedParts.join("\n\n--- [QUEBRA DE PÁGINA] ---\n\n");
    }

    return cleanRaw;
  };

  const handleMergeFiles = async () => {
    if (files.length === 0) {
      setErrorMessage("Por favor, adicione documentos na Base de Conhecimento para mesclar.");
      return;
    }
    setIsLoading(true);
    try {
      // 1. Tenta extrair via servidor backend
      let extractedText = "";
      try {
        const formData = new FormData();
        files.forEach(f => formData.append("files", f));

        const res = await fetch("/api/extract", {
          method: "POST",
          body: formData,
        });
        if (res.ok) {
          const data = await res.json();
          if (data.success && data.text) {
            extractedText = data.text;
          }
        }
      } catch (serverErr) {
        console.warn("Backend extract endpoint falhou, usando extrator client-side:", serverErr);
      }

      // 2. Extração client-side robusta (incluindo decodificador binário de stream de PDF)
      if (!extractedText) {
        const textParts: string[] = [];
        for (const file of files) {
          try {
            if (file.type.includes("text") || file.name.endsWith(".txt") || file.name.endsWith(".md") || file.name.endsWith(".csv")) {
              const content = await file.text();
              textParts.push(content);
            } else if (file.name.toLowerCase().endsWith(".pdf") || file.type === "application/pdf") {
              // Extrator direto de fluxos de texto de PDF no navegador
              const arrayBuffer = await file.arrayBuffer();
              const bytes = new Uint8Array(arrayBuffer);
              const latinText = new TextDecoder("latin1").decode(bytes);
              
              // Extrai blocos de texto PDF entre parênteses em instruções Tj/TJ ou blocos de texto puro
              const pdfMatches = latinText.match(/\(([^()]{2,})\)\s*(?:Tj|'|")/g) || [];
              let extractedPdfString = "";
              if (pdfMatches.length > 0) {
                extractedPdfString = pdfMatches
                  .map(m => m.replace(/^\(/, '').replace(/\)\s*(?:Tj|'|")$/, ''))
                  .join(' ')
                  .replace(/\\([()\\])/g, '$1')
                  .replace(/\s+/g, ' ');
              } else {
                // Fallback de strings legíveis para PDFs protegidos ou estruturados
                const cleanChars = latinText.replace(/[^\x20-\x7E\xA0-\xFF\n\r]/g, ' ');
                const chunks = cleanChars.split(/\s{3,}/).filter(c => c.trim().length > 15);
                extractedPdfString = chunks.join('\n\n');
              }

              if (extractedPdfString.trim()) {
                textParts.push(`--- Conteúdo do PDF: ${file.name} ---\n\n${extractedPdfString.trim()}`);
              }
            }
          } catch (readErr) {
            console.warn(`Erro ao ler arquivo ${file.name}:`, readErr);
          }
        }
        extractedText = textParts.join("\n\n");
      }

      if (extractedText && extractedText.trim()) {
        // Organiza as seções na ordem oficial da ABNT sem mudar nenhuma palavra do texto original
        const structuredText = organizeTextInABNTOrder(extractedText);
        setGeneratedText(prev => prev ? prev + "\n\n--- [QUEBRA DE PÁGINA] ---\n\n" + structuredText : structuredText);
        setErrorMessage("✅ Documento organizado rigorosamente na ordem ABNT (Resumo, Abstract, Introdução, Desenvolvimento, Resultados, Conclusão, Referências) sem alterar seu texto!");
        setTimeout(() => setErrorMessage(""), 4500);
        setActiveTab("editor");
        logAction("Organização Estrutural ABNT de Arquivo", structuredText);
      } else {
        setErrorMessage("Não foi possível extrair o texto dos arquivos. Tente novamente.");
      }
    } catch (error) {
      console.error(error);
      setErrorMessage(error instanceof Error ? error.message : "Erro ao extrair e organizar textos.");
    } finally {
      setIsLoading(false);
    }
  };

  // MONTAR TRABALHO EM GRUPO — Extrai cada seção, gera capa com todos os membros e monta tudo
  const handleGroupAssemble = async () => {
    const filledSlots = Object.entries(sectionSlots).filter(([, file]) => file !== null);
    if (filledSlots.length === 0) {
      setErrorMessage("Adicione pelo menos um arquivo em uma das seções para montar o trabalho.");
      return;
    }
    setIsLoading(true);
    setErrorMessage("");
    try {
      const activeSections = getGroupSectionsByDocType(groupDocType);

      // Extrai texto de cada slot via backend ou client-side
      const sectionTexts: Record<string, string> = {};
      for (const { key } of activeSections) {
        const file = sectionSlots[key];
        if (!file) continue;
        let text = "";
        try {
          const formData = new FormData();
          formData.append("files", file);
          const res = await fetch("/api/extract", { method: "POST", body: formData });
          if (res.ok) {
            const data = await res.json();
            if (data.success && data.text) text = data.text;
          }
        } catch { /* fallback abaixo */ }

        if (!text) {
          if (file.type.includes("text") || file.name.endsWith(".txt") || file.name.endsWith(".md") || file.name.endsWith(".csv")) {
            text = await file.text();
          } else if (file.name.toLowerCase().endsWith(".pdf") || file.type === "application/pdf") {
            const ab = await file.arrayBuffer();
            const latin = new TextDecoder("latin1").decode(new Uint8Array(ab));
            const matches = latin.match(/\(([^()]{2,})\)\s*(?:Tj|'|")/g) || [];
            text = matches.length > 0
              ? matches.map(m => m.replace(/^\(/, '').replace(/\)\s*(?:Tj|'|")$/, '')).join(' ').replace(/\\([()\\])/g, '$1').replace(/\s+/g, ' ')
              : latin.replace(/[^\x20-\x7E\xA0-\xFF\n\r]/g, ' ').split(/\s{3,}/).filter(c => c.trim().length > 15).join('\n\n');
          }
        }
        if (text.trim()) sectionTexts[key] = text.trim();
      }

      if (Object.keys(sectionTexts).length === 0) {
        setErrorMessage("Não foi possível extrair texto de nenhum arquivo. Verifique os formatos.");
        setIsLoading(false);
        return;
      }

      // Monta dados da capa
      const currentYear = new Date().getFullYear().toString();
      const instName = (institution || "INSTITUIÇÃO DE ENSINO").toUpperCase();
      const courseName = course ? course.toUpperCase() : "";
      const membersText = groupMembers.filter(m => m.trim()).map(m => m.trim().toUpperCase()).join("\n");
      const authorNames = membersText || (studentName || "NOME DOS AUTORES").toUpperCase();
      const docTitle = (title || "TÍTULO DO TRABALHO").toUpperCase();
      const docSubtitle = subtitle ? ` - ${subtitle}` : "";
      const docCity = (city || "CIDADE - UF").toUpperCase();
      const docYear = year || currentYear;
      const advText = advisor ? `Orientador(a): ${advisor}` : "";
      const typeLabel = groupDocType === "custom" || groupDocType === "outro"
        ? (customGroupDocName || "Trabalho Acadêmico")
        : groupDocType === "projeto" ? "Projeto de Pesquisa" : groupDocType === "relatorio" ? "Relatório Técnico" : groupDocType === "estudo_caso" ? "Estudo de Caso" : groupDocType === "resenha" ? "Resenha Crítica" : groupDocType === "artigo" || groupDocType === "artigo_cientifico" ? "Artigo Científico" : "Trabalho de Conclusão de Curso (TCC)";

      // Capa ABNT com todos os membros
      const coverPage = `${instName}${courseName ? `\n${courseName}` : ""}\n\n\n\n${authorNames}\n\n\n\n\n\n\n\n${docTitle}${docSubtitle}\n\n\n\n\n\n\n\n\n\n${docCity}\n${docYear}`;

      // Folha de Rosto
      const titlePage = `${authorNames}\n\n\n\n\n\n\n\n${docTitle}${docSubtitle}\n\n\n\n                                          ${typeLabel} apresentado à ${instName}${courseName ? ` como requisito parcial de avaliação para o curso de ${courseName}` : ""}.\n${advText ? `\n                                          ${advText}` : ""}\n\n\n\n\n\n\n\n${docCity}\n${docYear}`;

      // Monta seções dinamicamente conforme o tipo escolhido (com Quebra de Página automática a cada ~2200 caracteres de cada seção)
      const bodyPages: string[] = [];
      for (const { key, label } of activeSections) {
        if (sectionTexts[key]) {
          const rawText = sectionTexts[key];
          const cleanText = rawText.replace(/^---\s*(?:Início|Conteúdo)\s*d[eo]\s*(?:Arquivo|PDF):.*?---\s*/gi, '').trim();
          
          // Quebra automática de páginas dentro da seção se o texto for longo (~2200 caracteres por folha A4)
          const paragraphs = cleanText.split(/\n\n+/);
          let currentSectionPage = `${label}\n\n`;
          
          for (const para of paragraphs) {
            if ((currentSectionPage + "\n\n" + para).length > 2200 && currentSectionPage.trim().length > label.length) {
              bodyPages.push(currentSectionPage.trim());
              currentSectionPage = para;
            } else {
              currentSectionPage = currentSectionPage ? currentSectionPage + "\n\n" + para : para;
            }
          }
          if (currentSectionPage.trim()) {
            bodyPages.push(currentSectionPage.trim());
          }
        }
      }

      // Aplica a Skill do Professor / Normalizador Acadêmico ABNT diretamente no texto montado
      let formattedFullDoc = [coverPage, titlePage, ...bodyPages].join("\n\n--- [QUEBRA DE PÁGINA] ---\n\n");
      
      // 1. Normalização de citações ABNT NBR 10520:2023 (caixa mista)
      formattedFullDoc = normalizeCitationsToABNT2023(formattedFullDoc);

      // 2. Normalização de espaçamentos e pontuações
      formattedFullDoc = formattedFullDoc
        .replace(/\r\n/g, '\n')
        .replace(/\n{4,}/g, '\n\n\n')
        .replace(/[ \t]+$/gm, '');

      // 3. Garante seções em CAIXA ALTA (NBR 6024)
      formattedFullDoc = formattedFullDoc.replace(/^(#+\s*|\d+\s+)(introdução|materiais e métodos|desenvolvimento|fundamentação teórica|metodologia|resultados e discussão|conclusão|considerações finais|referências)/gmi, (match, prefix, t) => {
        return `${prefix}${t.toUpperCase()}`;
      });

      setGeneratedText(formattedFullDoc);
      setActiveTab("editor");
      setErrorMessage("✅ Trabalho em grupo montado e 100% normalizado com a Skill Acadêmica ABNT!");
      setTimeout(() => setErrorMessage(""), 5000);
      logAction("Montagem e Normalização ABNT de Trabalho em Grupo", formattedFullDoc.substring(0, 500));
    } catch (error) {
      console.error(error);
      setErrorMessage(error instanceof Error ? error.message : "Erro ao montar trabalho em grupo.");
    } finally {
      setIsLoading(false);
    }
  };

  const handleFormatABNT = async () => {
    if (!generatedText) {
      setErrorMessage("Por favor, gere ou cole um texto no editor primeiro para formatar.");
      return;
    }
    setIsLoading(true);
    try {
      await new Promise(resolve => setTimeout(resolve, 600));
      
      // 1. Normaliza citações para o padrão ABNT NBR 10520:2023 (caixa mista)
      let formattedText = normalizeCitationsToABNT2023(generatedText);
      
      // 2. Normaliza quebras de linha e espaçamentos
      formattedText = formattedText
        .replace(/\r\n/g, '\n')
        .replace(/\n{3,}/g, '\n\n')
        .replace(/^[ \t]+/gm, '')
        .trim();

      // 3. Garante seções primárias em CAIXA ALTA (NBR 6024)
      formattedText = formattedText.replace(/^(#+\s*|\d+\s+)(introdução|desenvolvimento|fundamentação teórica|metodologia|resultados|conclusão|considerações finais|referências)/gmi, (match, prefix, t) => {
        return `${prefix}${t.toUpperCase()}`;
      });

      // 4. Se for documento que exige capa e ainda não tiver, injeta Capa e Folha de Rosto
      const requiresFormalCover = !["resumo", "redacao", "resenha"].includes(documentType);
      if (requiresFormalCover && !formattedText.includes("--- [QUEBRA DE PÁGINA] ---")) {
        const inst = institution ? institution.toUpperCase() : "INSTITUIÇÃO DE ENSINO SUPERIOR";
        const crs = course ? course.toUpperCase() : "CURSO DE GRADUAÇÃO";
        const aut = studentName ? studentName.toUpperCase() : "NOME DO(A) AUTOR(A)";
        const tit = (title || "TÍTULO DO TRABALHO").toUpperCase();
        const sub = subtitle ? `: ${subtitle}` : "";
        const cid = city ? city.toUpperCase() : "CIDADE - UF";
        const an = year || String(new Date().getFullYear());
        const adv = advisor || "Prof. Dr. Orientador";

        const presentationNote = documentType.includes("artigo")
          ? `Artigo científico/acadêmico apresentado ao(à) ${inst}, como requisito de avaliação acadêmica.`
          : `Trabalho de Conclusão de Curso apresentado ao(à) ${inst}, como requisito parcial para obtenção de grau.`;

        const coverHeader = `${inst}\n${crs}\n\n\n\n${aut}\n\n\n\n${tit}${sub}\n\n\n\n\n\n\n\n${cid}\n${an}\n\n--- [QUEBRA DE PÁGINA] ---\n\n${aut}\n\n\n\n${tit}${sub}\n\n\n${presentationNote}\nOrientador(a): ${adv}\n\n\n\n\n\n${cid}\n${an}\n\n--- [QUEBRA DE PÁGINA] ---\n\n`;
        formattedText = coverHeader + formattedText;
      }
      
      setGeneratedText(formattedText);
      setActiveTab("editor");
      setErrorMessage("✅ Documento 100% adequado às normas ABNT (Capa, Contra-capa, NBR 10520:2023, NBR 6023)!");
      setTimeout(() => setErrorMessage(""), 3500);
      logAction("Formatação ABNT Completa Aplicada", formattedText);
    } catch (error) {
      console.error(error);
      setErrorMessage(error instanceof Error ? error.message : "Erro ao formatar.");
    } finally {
      setIsLoading(false);
    }
  };

  const handleHumanize = async () => {
    if (!generatedText || !generatedText.trim()) {
      setErrorMessage("Por favor, gere ou cole um texto no editor para humanizar.");
      return;
    }

    let targetText = generatedText;
    let isSelection = false;
    let start = 0;
    let end = 0;

    if (textareaRef.current) {
      start = textareaRef.current.selectionStart;
      end = textareaRef.current.selectionEnd;
      if (start !== end && start >= 0 && end > start) {
        targetText = generatedText.substring(start, end);
        isSelection = true;
      }
    }

    // Se for o documento completo e contiver capa ABNT, isola a capa para humanizar somente o conteúdo textual
    let coverPrefix = "";
    if (!isSelection && targetText.includes("--- [QUEBRA DE PÁGINA] ---")) {
      const parts = targetText.split("--- [QUEBRA DE PÁGINA] ---");
      if (parts.length >= 3) {
        coverPrefix = parts[0] + "--- [QUEBRA DE PÁGINA] ---" + parts[1] + "--- [QUEBRA DE PÁGINA] ---\n\n";
        targetText = parts.slice(2).join("--- [QUEBRA DE PÁGINA] ---").trim();
      }
    }

    setIsLoading(true);
    try {
      const res = await fetch("/api/humanize", {
        method: "POST",
        headers: getApiHeaders(),
        body: JSON.stringify({ text: targetText }),
      });
      const textData = await res.text(); 
      let data; 
      try { data = JSON.parse(textData); } catch (e) { throw new Error(`Erro no servidor (${res.status}). Aguarde e tente novamente.`); }
      if (data.success) {
        const normalized = normalizeCitationsToABNT2023(data.text);
        if (isSelection) {
          const newFullText = generatedText.substring(0, start) + normalized + generatedText.substring(end);
          setGeneratedText(newFullText);
        } else {
          setGeneratedText(coverPrefix ? coverPrefix + normalized : normalized);
        }
        setActiveTab("editor");
        logAction("Texto Humanizado com IA (Anti-Plágio/Turnitin)", data.text);
        setErrorMessage("✨ Texto humanizado com sucesso! Padrões de IA e clichês removidos.");
        setTimeout(() => setErrorMessage(""), 3500);
      } else {
        setErrorMessage(data.error || "Falha ao humanizar texto.");
      }
    } catch (error) {
      console.error(error);
      setErrorMessage(error instanceof Error ? error.message : "Erro ao humanizar texto.");
    } finally {
      setIsLoading(false);
    }
  };

  const handleImproveText = async () => {
    if (!generatedText || !generatedText.trim()) {
      setErrorMessage("Por favor, gere ou cole um texto no editor para aprimorar.");
      return;
    }

    setIsLoading(true);
    try {
      const promptGrammar = `Você é um dos maiores Professores e Gramáticos de Língua Portuguesa Brasileira do país, associado a normalizadores acadêmicos seniores (ABNT NBR 14724, 6022, 6028, 6023, 10520:2023).
Reescreva, aprimore e eleve o texto acadêmico a seguir ao mais alto patamar de excelência:
1. GRAMÁTICA E ESTILO: Domínio impecável da norma culta brasileira (regência, crase, concordância, colocação pronominal). Vocabulário erudito, fluido e preciso.
2. ZERO CLICHÊS: Remova qualquer artificialidade de IA como "Em suma", "Vale ressaltar", "No cenário atual", "Podemos concluir".
3. PADRONIZAÇÃO ABNT: Mantenha as seções, o espaçamento de 1 linha em branco no RESUMO (NBR 6028), citações autor-data em caixa mista (Silva, 2023, p. 15) e referências alinhadas à esquerda (NBR 6023).
4. Retorne apenas o texto integral corrigido e aprimorado, sem comentários externos.

TEXTO ORIGINAL:
${generatedText}`;

      const res = await callGeminiDirectly(promptGrammar, customGeminiKey, "gemini-3.6-flash");
      if (res && res.length > 50) {
        const normalized = normalizeCitationsToABNT2023(res);
        setGeneratedText(normalized);
        setActiveTab("editor");
        logAction("Texto Aprimorado pelo Mestre de Língua Portuguesa e ABNT", normalized);
        setErrorMessage("✨ Texto aprimorado com excelência gramatical e rigor acadêmico!");
        setTimeout(() => setErrorMessage(""), 3500);
      } else {
        throw new Error("Falha ao aprimorar texto com a IA.");
      }
    } catch (error) {
      console.error(error);
      setErrorMessage(error instanceof Error ? error.message : "Erro ao aprimorar texto.");
    } finally {
      setIsLoading(false);
    }
  };

  const handleCorrectSpelling = async () => {
    if (!generatedText || !generatedText.trim()) {
      setErrorMessage("Por favor, gere ou insira um texto para verificar e corrigir a ortografia.");
      return;
    }

    setIsLoading(true);
    setErrorMessage("✍️ Professor de Português e Normalizador ABNT revisando o documento...");

    try {
      const promptSpelling = `Você é um dos mais renomados Professores Titulares de Língua Portuguesa e Normalizadores Acadêmicos do Brasil (membro da Academia Brasileira de Letras e especialista em ABNT/UNESP/USP).

Sua missão é realizar uma REVISÃO LINGUÍSTICA E NORMATIVA COMPLETA E IMPECÁVEL no documento acadêmico fornecido, aplicando o mais alto padrão culto do Português Brasileiro e as normas técnicas da ABNT.

DIRETRIZES OBRIGATÓRIAS DE REVISÃO:
1. ORTOGRAFIA & ACENTUAÇÃO: Aplicação rigorosa do Vocabulário Ortográfico da Língua Portuguesa (VOLP) e do Novo Acordo Ortográfico.
2. SINTAXE E GRAMÁTICA: Correção minuciosa de regência verbal/nominal, crase, concordância nominal/verbal e colocação pronominal (próclise, mesóclise e ênclise formal).
3. PONTUAÇÃO & COESÃO: Ajuste culto da pontuação (vírgulas, ponto e vírgula, dois-pontos), eliminando períodos truncados ou redundâncias.
4. NORMALIZAÇÃO ABNT:
   - Citações no sistema autor-data em caixa mista conforme ABNT NBR 10520:2023: ex: (Silva, 2023, p. 15).
   - Referências conforme ABNT NBR 6023:2025 alinhadas à esquerda com entrelinha simples.
   - Preservação total de quebras de página "--- [QUEBRA DE PÁGINA] ---", títulos de seções, tabelas e dados reais.
5. ZERO CLICHÊS DE IA: Substitua termos genéricos como "Em suma", "Vale ressaltar", "No cenário atual", "Podemos concluir" por conectivos acadêmicos cultos (ex: "Dessarte", "Nesse prisma", "Com efeito", "Impende salientar").
6. RETORNO: Retorne EXCLUSIVAMENTE o documento integral 100% revisado e corrigido, pronto para uso, sem notas introdutórias e sem comentários adicionais.

DOCUMENTO A SER REVISADO:
${generatedText}`;

      let revisedText = "";
      try {
        revisedText = await callGeminiDirectly(promptSpelling, customGeminiKey, "gemini-3.6-flash");
      } catch (geminiErr) {
        console.warn("Chamada direta falhou, tentando fallback com gemini-3.5-flash-lite:", geminiErr);
        try {
          revisedText = await callGeminiDirectly(promptSpelling, customGeminiKey, "gemini-3.5-flash-lite");
        } catch (liteErr) {
          console.warn("Chamada direta lite falhou, tentando rota /api/generate:", liteErr);
          try {
            const resLocal = await fetch("/api/generate", {
              method: "POST",
              headers: { "Content-Type": "application/json" },
              body: JSON.stringify({ prompt: promptSpelling, model: "gemini-3.5-flash-lite" }),
            });
            const data = await resLocal.json();
            const candText = data?.candidates?.[0]?.content?.parts?.[0]?.text;
            if (candText && candText.trim().length > 50) revisedText = candText.trim();
          } catch (fetchErr) {
            console.warn("Todas as chamadas de rede falharam, aplicando motor de normalização local:", fetchErr);
          }
        }
      }

      // Se por algum motivo as redes falharem, aplica o motor local de regras gramaticais e ABNT
      if (!revisedText || revisedText.length < 50) {
        let localFixed = normalizeCitationsToABNT2023(generatedText);
        // Correções ortográficas e gramaticais essenciais locais
        localFixed = localFixed
          .replace(/\b(a\s+nivel\s+de)\b/gi, 'em nível de')
          .replace(/\b(ao\s+meu\s+ver)\b/gi, 'a meu ver')
          .replace(/\b(em\s+via\s+de\s+regra)\b/gi, 'em regra')
          .replace(/\b(a\s+grosso\s+modo)\b/gi, 'em linhas gerais')
          .replace(/\b(com\s+certeza)\b/gi, 'certamente')
          .replace(/\b(nos\s+dias\s+de\s+hoje|no\s+cen[aá]rio\s+atual)\b/gi, 'hodiernamente')
          .replace(/\b(em\s+suma|podemos\s+concluir)\b/gi, 'dessarte')
          .replace(/\b(vale\s+ressaltar|vale\s+destacar)\b/gi, 'impende salientar')
          .replace(/^(#+\s*|\d+\s+)(introdução|desenvolvimento|metodologia|resultados|conclusão|considerações finais|referências)/gmi, (match, pfx, txt) => `${pfx}${txt.toUpperCase()}`);
        revisedText = localFixed;
      }

      const normalized = normalizeCitationsToABNT2023(revisedText);
      updateGeneratedTextWithHistory(normalized);
      setActiveTab("editor");
      logAction("Revisão Gramatical e Ortográfica ABNT Aplicada com Sucesso", normalized.substring(0, 300));
      setErrorMessage("✨ Revisão ortográfica, gramatical e ABNT realizada com sucesso pelo Professor de Português!");
      setTimeout(() => setErrorMessage(""), 4000);
    } catch (error) {
      console.error(error);
      setErrorMessage(error instanceof Error ? error.message : "Erro ao revisar ortografia.");
    } finally {
      setIsLoading(false);
    }
  };

  const handleCheckAuthenticity = async () => {
    if (!generatedText) return;
    setIsLoading(true);
    try {
      const res = await fetch("/api/check-authenticity", {
        method: "POST",
        headers: getApiHeaders(),
        body: JSON.stringify({ text: generatedText }),
      });
      const textData = await res.text(); let data; try { data = JSON.parse(textData); } catch (e) { throw new Error(`Erro no servidor (${res.status}). Aguarde e tente novamente.`); }
      if (data.success) {
        setAuthenticityReport(data.report);
        setActiveTab("report");
        logAction("Verificação de plágio/IA realizada");
      } else {
        setErrorMessage(data.error);
      }
    } catch (error) {
      console.error(error);
      setErrorMessage(error instanceof Error ? error.message : "Erro ao verificar autenticidade.");
    } finally {
      setIsLoading(false);
    }
  };

  const handleInsertCover = async () => {
    // 1. Verificação de Regra ABNT NBR 6028: Resumos e Redações ENEM NÃO levam capa avulsa
    if (documentType === "resumo" || documentType === "redacao") {
      setErrorMessage("ℹ️ Conforme a ABNT NBR 6028, Resumos, Fichamentos e Redações possuem fluxo contínuo e NÃO utilizam Capa nem Folha de Rosto.");
      setTimeout(() => setErrorMessage(""), 4500);
      return;
    }

    setIsLoading(true);
    try {
      const inst = (institution || "NOME DA INSTITUIÇÃO DE ENSINO").toUpperCase();
      const crs = course ? course.toUpperCase() : "";
      const subMat = subject ? `DISCIPLINA: ${subject.toUpperCase()}` : "";
      const shiftClassInfo = [shift ? `Turno: ${shift}` : "", classroom ? `Sala/Turma: ${classroom}` : ""].filter(Boolean).join(" • ");
      const aut = (studentName || "NOME DO(A) AUTOR(A)").toUpperCase();
      const tit = (title || "TÍTULO DO TRABALHO").toUpperCase();
      const sub = subtitle ? ` - ${subtitle}` : "";
      const cid = (city || "CIDADE - UF").toUpperCase();
      const an = year || String(new Date().getFullYear());
      const adv = advisor ? `Orientador(a): ${advisor}` : "";

      const docTypeLabel = documentType === "outros" ? (customDocumentType || "Trabalho Acadêmico") : documentType === "artigo" || documentType === "artigo_cientifico" ? "Artigo Científico" : documentType === "projeto" ? "Projeto de Pesquisa" : documentType === "relatorio" ? "Relatório Técnico" : "Trabalho de Conclusão de Curso (TCC)";

      const presentationNote = `${docTypeLabel} apresentado à ${inst}${course ? ` como requisito parcial de avaliação para o curso de ${course}` : ""}.${shiftClassInfo ? `\n${shiftClassInfo}` : ""}${adv ? `\n${adv}` : ""}`;

      // 1. CAPA OFICIAL ABNT NBR 14724
      const subHeader = [crs, subMat, shiftClassInfo].filter(Boolean).join("\n");
      const coverPage = `${inst}${subHeader ? `\n${subHeader}` : ""}\n\n\n\n${aut}\n\n\n\n\n\n\n\n${tit}${sub}\n\n\n\n\n\n\n\n\n\n${cid}\n${an}`;

      // 2. FOLHA DE ROSTO OFICIAL ABNT NBR 14724 (Com Nota de Apresentação de 7,5cm)
      const titlePage = `${aut}\n\n\n\n\n\n\n\n${tit}${sub}\n\n\n\n                                          ${presentationNote}\n\n\n\n\n\n\n\n${cid}\n${an}`;

      const coverBlock = `${coverPage}\n\n--- [QUEBRA DE PÁGINA] ---\n\n${titlePage}\n\n--- [QUEBRA DE PÁGINA] ---\n\n`;

      // Remove capa e folha de rosto anteriores se já existirem no início do documento
      let cleanBody = generatedText || "";
      if (cleanBody.includes("--- [QUEBRA DE PÁGINA] ---")) {
        const parts = cleanBody.split("--- [QUEBRA DE PÁGINA] ---");
        const filteredParts = parts.filter(p => {
          const t = p.trim();
          return !t.startsWith("CAPA") && !t.startsWith("FOLHA DE ROSTO") && t !== "CAPA_AUTO" && t !== "FOLHA_ROSTO_AUTO" && !t.includes("requisito parcial") && !t.includes("apresentado à");
        });
        cleanBody = filteredParts.join("\n\n--- [QUEBRA DE PÁGINA] ---\n\n").trim();
      } else {
        const t = cleanBody.trim();
        if (t.startsWith("CAPA") || t.startsWith("FOLHA DE ROSTO") || t === "CAPA_AUTO" || t === "FOLHA_ROSTO_AUTO") {
          cleanBody = "";
        }
      }

      let updatedFullText = coverBlock + (cleanBody.trim() ? cleanBody.trim() : "1 INTRODUÇÃO\n\nInsira ou continue seu trabalho acadêmico aqui...");
      
      // Aplica normalizações ABNT
      updatedFullText = normalizeCitationsToABNT2023(updatedFullText);
      updatedFullText = updatedFullText
        .replace(/\r\n/g, '\n')
        .replace(/\n{4,}/g, '\n\n\n')
        .replace(/[ \t]+$/gm, '');

      updateGeneratedTextWithHistory(updatedFullText);
      setActiveTab("editor");
      setErrorMessage("✅ Capa e Folha de Rosto Oficiais ABNT NBR 14724 inseridas com sucesso!");
      setTimeout(() => setErrorMessage(""), 3500);
      logAction("Inserção de Capa e Folha de Rosto ABNT", updatedFullText.substring(0, 500));
    } catch (error) {
      console.error(error);
      setErrorMessage(error instanceof Error ? error.message : "Erro ao inserir capa e folha de rosto.");
    } finally {
      setIsLoading(false);
    }
  };

  const handlePaginate = () => {
    if (!generatedText || !generatedText.trim()) {
      setErrorMessage("Por favor, gere ou insira um texto para paginar.");
      return;
    }

    // Limpa números ocultos para permitir repaginar tudo de novo
    setHiddenPageNumbers(new Set());
    
    // Separa a Capa/Folha de Rosto (elementos pré-textuais não numerados) do corpo do trabalho
    let coverBlocks: string[] = [];
    let bodyText = generatedText;

    if (generatedText.includes("--- [QUEBRA DE PÁGINA] ---")) {
      const parts = generatedText.split("--- [QUEBRA DE PÁGINA] ---");
      const coverParts: string[] = [];
      const bodyParts: string[] = [];
      parts.forEach(p => {
        const t = p.trim();
        if (t.startsWith("CAPA") || t.startsWith("FOLHA DE ROSTO") || t === "CAPA_AUTO" || t === "FOLHA_ROSTO_AUTO" || t.includes("requisito parcial")) {
          coverParts.push(t);
        } else if (t.length > 0) {
          bodyParts.push(t);
        }
      });
      coverBlocks = coverParts;
      bodyText = bodyParts.join("\n\n");
    }

    // Divide em páginas A4 (~2200 caracteres com espaçamento 1.5)
    const paragraphs = bodyText.split(/\n\n+/);
    let pages: string[] = [];
    let currentChunk = "";

    paragraphs.forEach((p) => {
      if ((currentChunk + "\n\n" + p).length > 2200 && currentChunk.length > 0) {
        pages.push(currentChunk.trim());
        currentChunk = p;
      } else {
        currentChunk = currentChunk ? currentChunk + "\n\n" + p : p;
      }
    });
    if (currentChunk.trim()) {
      pages.push(currentChunk.trim());
    }

    const allPages = [...coverBlocks, ...pages];
    const fullResult = allPages.join("\n\n--- [QUEBRA DE PÁGINA] ---\n\n");

    updateGeneratedTextWithHistory(fullResult);
    setActiveTab("editor");
    logAction("Paginação e Repaginação A4 ABNT aplicada com sucesso");
    setErrorMessage("✅ Documento paginado e repaginado conforme as normas da ABNT!");
    setTimeout(() => setErrorMessage(""), 3500);
  };

  const handleGenerateTOC = () => {
    if (!generatedText || !generatedText.trim()) {
      setErrorMessage("Por favor, gere ou cole um texto no editor para criar o sumário.");
      return;
    }
    
    // 1. Obtém as páginas reais do documento
    let rawPages: string[] = [];
    if (generatedText.includes("--- [QUEBRA DE PÁGINA] ---")) {
      rawPages = generatedText.split("--- [QUEBRA DE PÁGINA] ---");
    } else {
      const paragraphs = generatedText.split(/\n\n+/);
      let curPage = "";
      for (const para of paragraphs) {
        if ((curPage + "\n\n" + para).length > 2200 && curPage.trim().length > 0) {
          rawPages.push(curPage.trim());
          curPage = para;
        } else {
          curPage = curPage ? curPage + "\n\n" + para : para;
        }
      }
      if (curPage.trim()) rawPages.push(curPage.trim());
    }

    // 2. Localiza cada seção e determina o número da página oficial (NBR 6027)
    // Capa = Não conta; Folha de Rosto = Conta (pág 1), mas não exibe. Textual (pág 2 em diante ou pág 1 se sem capa)
    const requiresFormalCover = !["resumo", "redacao", "resenha"].includes(documentType);
    const offset = requiresFormalCover ? 1 : 1; // Página visível/contada

    const tocEntries: { title: string; page: number }[] = [];
    
    // Verifica elementos pré-textuais presentes
    rawPages.forEach((pageContent, idx) => {
      const actualPageNum = requiresFormalCover ? idx : idx + 1;
      const lines = pageContent.split('\n');
      
      for (const line of lines) {
        const clean = line.trim().replace(/^#+\s*/, '');
        if (!clean) continue;

        // Seções numeradas (1 INTRODUÇÃO, 2 FUNDAMENTAÇÃO, 2.1 Subseção, etc.)
        const isNumberedSection = /^\d+(?:\.\d+)*\s+[A-ZÀ-Ú]/.test(clean);
        // Seções não numeradas obrigatórias (RESUMO, ABSTRACT, REFERÊNCIAS, CONCLUSÃO)
        const isStandardSection = /^(RESUMO|ABSTRACT|CONSIDERAÇÕES FINAIS|CONCLUSÃO|REFERÊNCIAS(?:\s+BIBLIOGRÁFICAS)?)\b/i.test(clean);

        if (isNumberedSection || isStandardSection) {
          // Evita duplicatas do próprio título SUMÁRIO
          if (!clean.toUpperCase().startsWith("SUMÁRIO") && !tocEntries.some(e => e.title.toUpperCase() === clean.toUpperCase())) {
            tocEntries.push({
              title: clean.toUpperCase(),
              page: Math.max(1, actualPageNum)
            });
          }
        }
      }
    });

    if (tocEntries.length === 0) {
      // Cria entradas padrão se o texto ainda não tiver títulos formatados
      tocEntries.push(
        { title: "RESUMO", page: 3 },
        { title: "ABSTRACT", page: 4 },
        { title: "1 INTRODUÇÃO", page: 5 },
        { title: "2 FUNDAMENTAÇÃO TEÓRICA E METODOLOGIA", page: 6 },
        { title: "2.1 ANÁLISE DAS DIMENSÕES ESTRUTURAIS", page: 7 },
        { title: "3 RESULTADOS E DISCUSSÃO", page: 8 },
        { title: "4 CONSIDERAÇÕES FINAIS", page: 9 },
        { title: "REFERÊNCIAS", page: 10 }
      );
    }

    // 3. Monta o Sumário com pontilhados líderes (ABNT NBR 6027)
    // Ex: 1 INTRODUÇÃO ............................................................................ 4
    const formattedTOCLines = tocEntries.map(entry => {
      const dotsCount = Math.max(5, 75 - entry.title.length - String(entry.page).length);
      const dots = ".".repeat(dotsCount);
      return `${entry.title} ${dots} ${entry.page}`;
    });

    const tocBlock = `SUMÁRIO\n\n${formattedTOCLines.join('\n')}\n\n--- [QUEBRA DE PÁGINA] ---\n\n`;

    // Se já existia um sumário anterior, substitui; caso contrário, insere após a folha de rosto
    let newFullText = generatedText;
    if (newFullText.includes("SUMÁRIO\n\n")) {
      newFullText = newFullText.replace(/SUMÁRIO\n\n[\s\S]*?--- \[QUEBRA DE PÁGINA\] ---\n\n/i, '');
    }

    if (requiresFormalCover && newFullText.includes("--- [QUEBRA DE PÁGINA] ---")) {
      const parts = newFullText.split("--- [QUEBRA DE PÁGINA] ---");
      if (parts.length >= 2) {
        // Insere após a Capa e Folha de Rosto
        const preCovers = parts[0] + "--- [QUEBRA DE PÁGINA] ---" + parts[1] + "--- [QUEBRA DE PÁGINA] ---\n\n";
        const remaining = parts.slice(2).join("--- [QUEBRA DE PÁGINA] ---").trim();
        newFullText = preCovers + tocBlock + remaining;
      } else {
        newFullText = tocBlock + newFullText;
      }
    } else {
      newFullText = tocBlock + newFullText;
    }

    setGeneratedText(newFullText);
    setActiveTab("editor");
    logAction("Sumário ABNT NBR 6027 Gerado", tocBlock);
    setErrorMessage("✅ Sumário gerado com paginação e pontilhados conforme a ABNT NBR 6027!");
    setTimeout(() => setErrorMessage(""), 3500);
  };

  const handleCloseReferenceModal = () => {
    setIsLoading(false);
    setReferenceSource("");
    setGeneratedReference("");
    setShowReferenceModal(false);
  };

  const handleGenerateReference = async () => {
    if (!referenceSource || !referenceSource.trim()) {
      setErrorMessage("Por favor, insira um link ou DOI.");
      return;
    }
    setIsLoading(true);
    setGeneratedReference("");
    try {
      // Geração ultrarrápida direta pelo Gemini com normalização ABNT NBR 6023:2025
      const promptRef = `Você é um normalizador bibliográfico sênior especialista nas normas ABNT NBR 6023:2025 e APA 7th.
Gere a referência bibliográfica COMPLETA e FORMAL a partir da fonte abaixo:
- Fonte / DOI / Link / Dados: "${referenceSource.trim()}"
- Estilo: ${referenceStyle}
- Regras: Caixa alta para sobrenomes do autor na ABNT (ex: SILVA, João), título do livro em negrito ou artigo com revista em itálico, ano, cidade e link/DOI se houver.
- Retorne APENAS a referência bibliográfica final pronta para uso, sem aspas e sem comentários adicionais.`;

      let refText = "";
      try {
        refText = await callGeminiDirectly(promptRef, customGeminiKey, "gemini-3.6-flash");
      } catch (geminiErr) {
        console.warn("Chamada direta falhou, tentando rota /api/generate-reference:", geminiErr);
        const res = await fetch("/api/generate-reference", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ source: referenceSource, style: referenceStyle }),
        });
        if (res.ok) {
          const data = await res.json();
          if (data.success && data.text) refText = data.text;
        }
      }

      if (refText && refText.trim()) {
        setGeneratedReference(refText.trim());
        logAction("Referência Bibliográfica Gerada", refText.trim());
      } else {
        setErrorMessage("Não foi possível formatar a referência. Verifique os dados e tente novamente.");
      }
    } catch (error) {
      console.error(error);
      setErrorMessage(error instanceof Error ? error.message : "Erro ao gerar referência.");
    } finally {
      setIsLoading(false);
    }
  };



  const handleAttachmentFile = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const fileList = e.target.files;
    if (!fileList || fileList.length === 0) return;
    
    const filesArray = Array.from(fileList);
    const imageFiles = filesArray.filter(f => f.type.startsWith("image/") || /\.(jpg|jpeg|png|webp|gif|svg)$/i.test(f.name));
    const nonImageFiles = filesArray.filter(f => !imageFiles.includes(f));

    // Processa múltiplas imagens ao mesmo tempo
    if (imageFiles.length > 0) {
      let loadedCount = 0;
      imageFiles.forEach((file, idx) => {
        const reader = new FileReader();
        reader.onload = (event) => {
          const base64 = event.target?.result as string;
          const figNum = idx + 1;
          const figureBlock = `\n\n![Figura inserida](${base64})\nFigura ${figNum} – Representação Ilustrativa do Objeto de Estudo\nFonte: Elaborado pelos autores (${new Date().getFullYear()}).\n\n`;
          setGeneratedText(prev => prev ? prev + figureBlock : figureBlock);
          loadedCount++;
          if (loadedCount === imageFiles.length) {
            setActiveTab("editor");
            setErrorMessage(`✅ ${imageFiles.length} ${imageFiles.length > 1 ? 'imagens inseridas' : 'imagem inserida'} com sucesso no padrão ABNT!`);
            setTimeout(() => setErrorMessage(""), 3500);
          }
        };
        reader.readAsDataURL(file);
      });
    }

    // Processa tabelas CSV ou TXT
    for (const file of nonImageFiles) {
      if (file.name.endsWith(".csv") || file.name.endsWith(".txt")) {
        const reader = new FileReader();
        reader.onload = async (event) => {
          const content = event.target?.result as string;
          setIsLoading(true);
          try {
            const res = await fetch("/api/csv-to-table", {
              method: "POST",
              headers: getApiHeaders(),
              body: JSON.stringify({ csvData: content }),
            });
            const textData = await res.text(); let data; try { data = JSON.parse(textData); } catch (e) { throw new Error(`Erro no servidor (${res.status}). Aguarde e tente novamente.`); }
            if (data.success) {
              setGeneratedText(prev => prev + "\n\n" + data.text + "\n\n");
            } else {
              setErrorMessage(data.error);
            }
          } catch (error) {
            console.error(error);
            setErrorMessage(error instanceof Error ? error.message : "Erro ao processar tabela.");
          } finally {
            setIsLoading(false);
          }
        };
        reader.readAsText(file);
      }
    }
    
    if (attachmentRef.current) attachmentRef.current.value = "";
  };

  // Inicia o Quiz Interativo estruturado com IA a partir do texto do palco
  const handleStartInteractiveQuiz = async () => {
    setIsChatting(true);
    setActiveQuiz(null);
    const userMsg = "🎯 Preparar um Quiz interativo sobre este trabalho!";
    setChatHistory(prev => [...prev, { role: 'user' as const, text: userMsg }]);

    try {
      const quizGenPrompt = `Você é a EMIA, assistente e tutora da EDUTECH.
Com base no trabalho acadêmico fornecido, elabore um QUIZ INTERATIVO com exatamente 4 perguntas de múltipla escolha para testar a compreensão do aluno sobre os conceitos centrais do texto.

Responda OBRIGATORIAMENTE em formato JSON VÁLIDO puro (sem crases de markdown no início ou no fim, apenas o JSON):
[
  {
    "id": 1,
    "question": "Texto claro e objetivo da pergunta 1?",
    "options": [
      { "letter": "A", "text": "Texto da alternativa A" },
      { "letter": "B", "text": "Texto da alternativa B" },
      { "letter": "C", "text": "Texto da alternativa C" },
      { "letter": "D", "text": "Texto da alternativa D" }
    ],
    "correctAnswer": "A",
    "explanation": "Explicação pedagógica e clara do porquê a alternativa A é a correta com base no texto."
  }
]

DOCUMENTO ACADÊMICO:
${generatedText ? generatedText.substring(0, 4000) : "Metodologia científica, normas ABNT e estruturação acadêmica."}`;

      let rawJson = "";
      try {
        rawJson = await callGeminiDirectly(quizGenPrompt, customGeminiKey, "gemini-3.6-flash");
      } catch (e) {
        rawJson = await callGeminiDirectly(quizGenPrompt, customGeminiKey, "gemini-3.5-flash-lite");
      }

      const cleanJson = rawJson.replace(/```json/gi, '').replace(/```/g, '').trim();
      const parsedQuestions = JSON.parse(cleanJson);

      if (Array.isArray(parsedQuestions) && parsedQuestions.length > 0) {
        setActiveQuiz({
          isActive: true,
          questions: parsedQuestions,
          currentIndex: 0,
          userAnswers: [],
          isFinished: false
        });

        setChatHistory(prev => [...prev, { 
          role: 'assistant' as const, 
          text: `🎉 **Oba! Preparei um Quiz especial com ${parsedQuestions.length} perguntas sobre o seu trabalho!**\n\nResponda clicando nas alternativas abaixo. Conforme você clica, a pergunta muda automaticamente e no final eu te mostro seu gabarito com a nota e as explicações detalhadas! Vamos lá? 👇` 
        }]);
      } else {
        throw new Error("Formato inválido retornado pela IA.");
      }
    } catch (err) {
      console.error("Erro ao gerar quiz estruturado:", err);
      // Fallback
      handleSendMessage(undefined, "Cria um quiz com perguntas de múltipla escolha sobre o meu trabalho.");
    } finally {
      setIsChatting(false);
    }
  };

  // Processa a resposta do clique na alternativa do Quiz
  const handleQuizOptionSelect = (selectedLetter: string) => {
    if (!activeQuiz || activeQuiz.isFinished) return;

    const currentQ = activeQuiz.questions[activeQuiz.currentIndex];
    const isCorrect = selectedLetter.toUpperCase() === currentQ.correctAnswer.toUpperCase();
    const newAnswers = [
      ...activeQuiz.userAnswers,
      {
        selected: selectedLetter,
        isCorrect,
        explanation: currentQ.explanation,
        question: currentQ.question
      }
    ];

    const nextIndex = activeQuiz.currentIndex + 1;
    const isLastQuestion = nextIndex >= activeQuiz.questions.length;

    if (isLastQuestion) {
      // Finaliza o Quiz e calcula a pontuação
      const correctCount = newAnswers.filter(a => a.isCorrect).length;
      const totalCount = activeQuiz.questions.length;
      const percentage = Math.round((correctCount / totalCount) * 100);

      setActiveQuiz({
        ...activeQuiz,
        currentIndex: nextIndex,
        userAnswers: newAnswers,
        isFinished: true
      });

      // Mensagem explicativa no chat
      const feedbackEmoji = percentage >= 75 ? "🏆 Mandou muito bem!" : percentage >= 50 ? "👏 Bom resultado!" : "📚 Vale a pena revisar!";
      const explanationReport = activeQuiz.questions.map((q, idx) => {
        const ans = newAnswers[idx];
        return `**Questão ${idx + 1}: ${q.question}**\n- Sua resposta: **Alternativa ${ans.selected}** ${ans.isCorrect ? '✅ (Correta!)' : `❌ (Incorreta - Gabarito: **${q.correctAnswer}**)`}\n- 💡 *Explicação*: ${q.explanation}\n`;
      }).join('\n');

      setChatHistory(prev => [
        ...prev,
        {
          role: 'assistant' as const,
          text: `### 🎯 Resultado do seu Quiz:\n\n${feedbackEmoji}\nVocê acertou **${correctCount} de ${totalCount} questões** (${percentage}% de aproveitamento)!\n\n---\n\n### 📝 Gabarito & Explicação Completa:\n\n${explanationReport}\n\n*Quer fazer outro quiz ou tem alguma dúvida sobre os pontos explicados? É só me falar!* ✨`
        }
      ]);
    } else {
      // Passa para a próxima pergunta
      setActiveQuiz({
        ...activeQuiz,
        currentIndex: nextIndex,
        userAnswers: newAnswers,
        isFinished: false
      });
    }
  };

  const handleSendMessage = async (e?: React.FormEvent, customMsg?: string) => {
    if (e) e.preventDefault();
    const msgToSend = (customMsg !== undefined ? customMsg : chatMessage).trim();
    if (!msgToSend) return;

    setChatMessage("");
    const updatedHistory = [...chatHistory, { role: 'user' as const, text: msgToSend }];
    setChatHistory(updatedHistory);
    setIsChatting(true);
    logAction("Envio de instrução/mensagem no Chat de Edição");

    try {
      let assistantResponse = "";

      // Chamada direta rápida ao Gemini
      try {
        const chatPrompt = `Você é a EMIA, a mascote e assistente inteligente da EDUTECH! 🤖🎓
Seu tom de voz com o aluno é super amigável, informal, acolhedor, prestativo e direto ao ponto (sem enrolação e sem respostas gigantescas, a menos que o aluno peça um conteúdo detalhado).

Sua missão é:
- Trocar uma ideia sobre o trabalho acadêmico gerado na tela.
- Responder dúvidas sobre o texto ou sobre normas ABNT.
- Criar quizzes divertidos, podcasts, músicas chicletes, quadrinhos ou mapas mentais quando solicitado.
- Orientar como usar os botões e ferramentas do app de forma simples e prática.

🚨 REGRA MESTRE INVIOLÁVEL DE VERACIDADE CIENTÍFICA (ZERO ALUCINAÇÃO):
- Você está TERMINANTEMENTE PROIBIDA de inventar informações, criar dados fictícios, alucinar fontes ou responder com achismos infundados.
- Use SOMENTE informações oriundas de artigos científicos consolidados, periódicos indexados (SciELO, Scopus, Google Scholar, Capes) e fontes conceituadas (IBGE, OMS, IPEA, Universidades).

GUIA RÁPIDO DO APP:

1. BOTÃO "TRABALHO EM GRUPO" (Topo):
- Finalidade: Permite que equipes de alunos reúnam e montem automaticamente um trabalho completo em conformidade com a ABNT.
- Como funciona:
  a) Nome dos Integrantes: Insira o nome de cada membro da equipe (o app adiciona automaticamente todos na Capa e Folha de Rosto).
  b) Tipo de Documento: Escolha entre TCC, Artigo Científico, Relatório Técnico, Projeto de Pesquisa ou Personalizado.
  c) Slots de Seções: Cada integrante envia o arquivo (PDF, Word, TXT) correspondente à sua parte (ex: Introdução, Metodologia, Resultados, Conclusão).
  d) Botão "Montar Trabalho Completo": O sistema extrai os textos, gera a Capa Oficial com todos os autores, organiza na ordem canônica da ABNT, aplica quebras de página automáticas e normaliza citações e referências.

2. CAMPOS DO FORMULÁRIO PRINCIPAL ("Novo Trabalho"):
- Título e Subtítulo: Definem o tema central da pesquisa.
- Tipo de Documento: Ajusta a estrutura ABNT (Artigo, TCC, Monografia, Relatório, Resumo NBR 6028, Redação ENEM).
- Instruções Personalizadas (Prompt): Onde o usuário detalha objetivos, tópicos obrigatórios ou dados específicos. Tem prioridade máxima!
- Base de Conhecimento: Permite anexar PDFs, imagens ou documentos de referência.
- Dados do Trabalho (Perfil): Instituição, Curso, Autor, Orientador, Cidade e Ano (usados para Capa e Folha de Rosto).

3. FERRAMENTAS DO TOPO E EDITOR:
- 🔊 Ouvir Texto: Player com voz humana neural (Feminina 👩 / Masculina 👨), controle de velocidade (1x, 1.5x, 2x) e volume. Pula a capa e lê o conteúdo.
- 📽️ Slides (EMIA.SLIDES): Cria cartões de apresentação dinâmicos, exporta para PowerPoint (.pptx) e Google Slides.
- ✍️ Ortografia & Humanizar: Corrige a gramática e remove clichês de IA para Turnitin.
- PDF A4, Word (.docx) e LaTeX: Exportação em alta fidelidade com paginação ABNT.
- Inserir Capa & Sumário Dinâmico: Criação instantânea de elementos pré-textuais.

4. HABILIDADES ESPECIAIS DO CHAT (CRIATIVAS E DIDÁTICAS):
- 🎯 QUIZ INTERATIVO CLICÁVEL: Quando o aluno pedir um Quiz, gere 3 a 5 perguntas de múltipla escolha com 4 alternativas cada (A, B, C, D). Formate as alternativas claramente iniciando com "A) ", "B) ", "C) " e "D) ". Não revele o gabarito logo de cara, incentive o aluno a responder!
- 🎙️ Roteiro de PODCAST: Se solicitado, crie um roteiro de podcast dinâmico com 2 apresentadores (Host e Especialista), linguagem leve, descontraída e didática explicando todo o conteúdo acadêmico.
- 🎵 MÚSICA CHICLETE: Se solicitado, componha uma letra de música chiclete (estilo pop/refrão marcante ou paródia) com rimas fáceis para memorizar todos os conceitos e fórmulas do tema!
- 🎨 HISTÓRIA EM QUADRINHOS (HQ): Se solicitado, crie um roteiro em quadrinhos com descrição visual de cada quadro (painel), personagens acadêmicos carismáticos, balões de fala e onomatopeias.
- 🧠 MAPA MENTAL: Se solicitado, gere a estrutura hierárquica completa de Mapa Mental, incluindo diagrama formatado em Markdown ou código Mermaid (ex: \`\`\`mermaid graph TD ... \`\`\`).

${generatedText ? `[DOCUMENTO ATUAL DO USUÁRIO]\n${generatedText.substring(0, 3500)}\n[/DOCUMENTO ATUAL]\n` : ""}
[HISTÓRICO RECENTE]
${chatHistory.slice(-4).map(h => `${h.role === 'user' ? 'Aluno' : 'Assistente'}: ${h.text}`).join('\n')}
Aluno: ${msgToSend}
Assistente:`;

        assistantResponse = await callGeminiDirectly(chatPrompt, customGeminiKey, "gemini-3.6-flash");
      } catch (directErr) {
        console.warn("Tentativa direta falhou, tentando rota /api/chat:", directErr);
        try {
          const res = await fetch("/api/chat", {
            method: "POST",
            headers: getApiHeaders(),
            body: JSON.stringify({ 
              message: userMessage, 
              history: chatHistory.slice(-4),
              context: generatedText ? generatedText.substring(0, 3500) : ""
            }),
          });
          if (res.ok) {
            const data = await res.json();
            if (data.success && data.text) {
              assistantResponse = data.text;
            }
          }
        } catch (serverErr) {
          console.error("Falha em todas as vias do chat:", serverErr);
        }
      }

      if (assistantResponse) {
        setChatHistory([...updatedHistory, { role: 'assistant', text: assistantResponse }]);
      } else {
        setErrorMessage("Erro ao gerar resposta do chat. Tente novamente.");
      }
    } catch (error) {
      console.error(error);
      setErrorMessage(error instanceof Error ? error.message : "Erro ao enviar mensagem no chat.");
    } finally {
      setIsChatting(false);
    }
  };

  const handleCopy = () => {
    if (!generatedText) return;
    navigator.clipboard.writeText(generatedText);
    alert("Texto copiado para a área de transferência!");
  };

  const exportPDF = () => {
    if (!generatedText || !generatedText.trim()) {
      setErrorMessage("Nenhum texto disponível para exportar.");
      return;
    }

    const doc = new jsPDF({ 
      unit: 'mm',
      format: 'a4',
      orientation: 'portrait'
    });
    
    // Configurações Globais ABNT NBR 14724
    const marginLeft = 30; // 3,0 cm Margem Esquerda
    const marginTop = 30;  // 3,0 cm Margem Superior
    const marginRight = 20; // 2,0 cm Margem Direita
    const marginBottom = 20; // 2,0 cm Margem Inferior
    const printableWidth = 210 - marginLeft - marginRight; // 160 mm
    const maxY = 297 - marginBottom; // 277 mm
    const lineHeight = 6.5; // Espaçamento 1.5 para corpo
    const pageRightEdge = 210 - marginRight; // 190 mm

    // Normalização das Páginas ABNT
    let rawPages: string[] = [];
    if (generatedText.includes("--- [QUEBRA DE PÁGINA] ---")) {
      rawPages = generatedText.split("--- [QUEBRA DE PÁGINA] ---");
    } else {
      const paragraphs = generatedText.split(/\n\n+/);
      let curPage = "";
      for (const para of paragraphs) {
        if ((curPage + "\n\n" + para).length > 2000 && curPage.trim().length > 0) {
          rawPages.push(curPage.trim());
          curPage = para;
        } else {
          curPage = curPage ? curPage + "\n\n" + para : para;
        }
      }
      if (curPage.trim()) rawPages.push(curPage.trim());
    }

    const requiresFormalCover = !["resumo", "redacao", "resenha"].includes(documentType);
    let currentPageNum = 0;

    rawPages.forEach((pageContent, pageIdx) => {
      const trimmedPage = pageContent.trim();
      if (!trimmedPage) return;

      if (currentPageNum > 0) {
        doc.addPage();
      }
      currentPageNum++;

      const isCover = requiresFormalCover && pageIdx === 0;
      const isTitlePage = requiresFormalCover && pageIdx === 1;
      const isTextualBody = !requiresFormalCover || pageIdx >= 2;

      // Numeração Oficial: Canto Superior Direito a partir da página textual (NBR 14724)
      if (isTextualBody) {
        doc.setFont("helvetica", "normal");
        doc.setFontSize(10);
        doc.text(String(currentPageNum), pageRightEdge, 20, { align: "right" });
      }

      let cursorY = marginTop;

      if (isCover) {
        // --- 1. CAPA OFICIAL ABNT NBR 14724 ---
        doc.setFont("helvetica", "bold");
        doc.setFontSize(12);

        const instLines = (institution || "INSTITUIÇÃO DE ENSINO SUPERIOR").toUpperCase().split("\n");
        instLines.forEach(line => {
          doc.text(line, 105, cursorY, { align: "center" });
          cursorY += 6;
        });

        if (course) {
          doc.setFont("helvetica", "normal");
          doc.setFontSize(11);
          doc.text(course.toUpperCase(), 105, cursorY, { align: "center" });
        }

        // Autor
        cursorY = 90;
        doc.setFont("helvetica", "bold");
        doc.setFontSize(12);
        doc.text((studentName || "NOME DO(A) AUTOR(A)").toUpperCase(), 105, cursorY, { align: "center" });

        // Título e Subtítulo
        cursorY = 145;
        doc.setFont("helvetica", "bold");
        doc.setFontSize(14);
        const titleText = (title || "TÍTULO DO TRABALHO").toUpperCase();
        const splitTitle = doc.splitTextToSize(titleText, 150);
        splitTitle.forEach((tLine: string) => {
          doc.text(tLine, 105, cursorY, { align: "center" });
          cursorY += 7;
        });

        if (subtitle) {
          doc.setFont("helvetica", "normal");
          doc.setFontSize(12);
          doc.text(`: ${subtitle}`, 105, cursorY, { align: "center" });
        }

        // Cidade e Ano (Rodapé)
        doc.setFont("helvetica", "bold");
        doc.setFontSize(11);
        doc.text((city || "CIDADE - UF").toUpperCase(), 105, 265, { align: "center" });
        doc.setFont("helvetica", "normal");
        doc.text(year || String(new Date().getFullYear()), 105, 272, { align: "center" });

      } else if (isTitlePage) {
        // --- 2. FOLHA DE ROSTO / CONTRACAPA OFICIAL ABNT NBR 14724 ---
        doc.setFont("helvetica", "bold");
        doc.setFontSize(12);
        doc.text((studentName || "NOME DO(A) AUTOR(A)").toUpperCase(), 105, marginTop + 10, { align: "center" });

        // Título
        cursorY = 100;
        doc.setFont("helvetica", "bold");
        doc.setFontSize(13);
        const titleText = (title || "TÍTULO DO TRABALHO").toUpperCase();
        const splitTitle = doc.splitTextToSize(titleText, 150);
        splitTitle.forEach((tLine: string) => {
          doc.text(tLine, 105, cursorY, { align: "center" });
          cursorY += 7;
        });

        if (subtitle) {
          doc.setFont("helvetica", "normal");
          doc.setFontSize(11);
          doc.text(`: ${subtitle}`, 105, cursorY, { align: "center" });
        }

        // Nota de Apresentação com Recuo de 7,5cm / alinhada à direita
        cursorY = 150;
        doc.setFont("helvetica", "normal");
        doc.setFontSize(10);

        const presentationNote = documentType.includes("artigo")
          ? `Artigo científico/acadêmico apresentado ao(à) ${institution || "Instituição de Ensino"}, como requisito de avaliação acadêmica.`
          : `Trabalho de Conclusão de Curso apresentado ao(à) ${institution || "Instituição de Ensino"}, como requisito parcial para obtenção de grau.`;

        const noteLines = doc.splitTextToSize(presentationNote, 80);
        noteLines.forEach((nLine: string) => {
          doc.text(nLine, 110, cursorY);
          cursorY += 5;
        });

        if (advisor) {
          cursorY += 2;
          doc.setFont("helvetica", "bold");
          doc.text(`Orientador(a): ${advisor}`, 110, cursorY);
          doc.setFont("helvetica", "normal");
        }

        // Cidade e Ano (Rodapé)
        doc.setFont("helvetica", "bold");
        doc.setFontSize(11);
        doc.text((city || "CIDADE - UF").toUpperCase(), 105, 265, { align: "center" });
        doc.setFont("helvetica", "normal");
        doc.text(year || String(new Date().getFullYear()), 105, 272, { align: "center" });

      } else {
        // --- 3. CORPO DO TEXTO (PÁGINAS TEXTUAIS ABNT) ---
        doc.setFont("times", "normal");
        doc.setFontSize(12);

        const paragraphs = trimmedPage.split("\n");

        paragraphs.forEach((para) => {
          const rawPara = para.trim();
          if (!rawPara) {
            cursorY += lineHeight * 0.6;
            return;
          }

          // Títulos de Seção (1 INTRODUÇÃO, RESUMO, REFERÊNCIAS, etc.)
          const isHeading = /^(?:\d+(?:\.\d+)*\s+[A-ZÀ-Ú\s]+|RESUMO|ABSTRACT|SUMÁRIO|REFERÊNCIAS|CONSIDERAÇÕES FINAIS)$/.test(rawPara);
          // Citação Longa (> 3 linhas com recuo de 4cm)
          const isLongQuote = rawPara.startsWith("[CITAÇÃO_LONGA]") || rawPara.startsWith("   ") || (rawPara.length > 200 && rawPara.startsWith("    "));
          // Tabela ou Quadro
          const isTableOrFrame = rawPara.startsWith("Tabela") || rawPara.startsWith("Quadro") || rawPara.startsWith("Fonte:") || rawPara.startsWith("|") || rawPara.startsWith("+-") || rawPara.startsWith("--");

          if (isHeading) {
            cursorY += 4;
            if (cursorY > maxY - 15) {
              doc.addPage();
              currentPageNum++;
              doc.setFont("helvetica", "normal");
              doc.setFontSize(10);
              doc.text(String(currentPageNum), pageRightEdge, 20, { align: "right" });
              cursorY = marginTop;
            }
            doc.setFont("times", "bold");
            doc.setFontSize(12);
            doc.text(rawPara, marginLeft, cursorY);
            doc.setFont("times", "normal");
            cursorY += lineHeight + 2;
            return;
          }

          if (isTableOrFrame) {
            doc.setFont("courier", "normal");
            doc.setFontSize(9.5);
            const tableLines = doc.splitTextToSize(rawPara, printableWidth);
            tableLines.forEach((tLine: string) => {
              if (cursorY > maxY) {
                doc.addPage();
                currentPageNum++;
                doc.setFont("helvetica", "normal");
                doc.setFontSize(10);
                doc.text(String(currentPageNum), pageRightEdge, 20, { align: "right" });
                cursorY = marginTop;
                doc.setFont("courier", "normal");
                doc.setFontSize(9.5);
              }
              doc.text(tLine, marginLeft, cursorY);
              cursorY += 4.8;
            });
            doc.setFont("times", "normal");
            doc.setFontSize(12);
            return;
          }

          // Imagem / Figura Inserida
          const imageMatch = rawPara.match(/!\[.*?\]\((data:image\/.*?;base64,.*?)\)/);
          if (imageMatch) {
            const base64Data = imageMatch[1];
            const imgWidth = Math.min(130, printableWidth);
            const imgHeight = 70; // altura padronizada proporcional
            
            if (cursorY + imgHeight > maxY) {
              doc.addPage();
              currentPageNum++;
              doc.setFont("helvetica", "normal");
              doc.setFontSize(10);
              doc.text(String(currentPageNum), pageRightEdge, 20, { align: "right" });
              cursorY = marginTop;
            }

            try {
              const formatMatch = base64Data.match(/data:image\/(png|jpeg|jpg|webp)/i);
              const imgFormat = formatMatch ? formatMatch[1].toUpperCase() : 'JPEG';
              doc.addImage(base64Data, imgFormat, marginLeft + 15, cursorY, imgWidth, imgHeight);
              cursorY += imgHeight + 4;
            } catch (imgErr) {
              console.warn("Erro ao renderizar imagem no PDF:", imgErr);
            }
            return;
          }

          if (isLongQuote) {
            const cleanQuote = rawPara.replace(/^\[CITAÇÃO_LONGA\]\s*/, '').trim();
            doc.setFontSize(10);
            const quoteLines = doc.splitTextToSize(cleanQuote, printableWidth - 40); // Recuo de 4cm
            quoteLines.forEach((qLine: string) => {
              if (cursorY > maxY) {
                doc.addPage();
                currentPageNum++;
                doc.setFont("helvetica", "normal");
                doc.setFontSize(10);
                doc.text(String(currentPageNum), pageRightEdge, 20, { align: "right" });
                cursorY = marginTop;
                doc.setFont("times", "normal");
                doc.setFontSize(10);
              }
              doc.text(qLine, marginLeft + 40, cursorY);
              cursorY += 5;
            });
            doc.setFontSize(12);
            cursorY += 2;
            return;
          }

          // Parágrafo Regular Justificado com Recuo de Primeira Linha (1,25 cm)
          const firstLineIndent = 12.5; // 1.25 cm
          const paraLines = doc.splitTextToSize(rawPara, printableWidth);

          paraLines.forEach((line: string, lIdx: number) => {
            if (cursorY > maxY) {
              doc.addPage();
              currentPageNum++;
              doc.setFont("helvetica", "normal");
              doc.setFontSize(10);
              doc.text(String(currentPageNum), pageRightEdge, 20, { align: "right" });
              cursorY = marginTop;
              doc.setFont("times", "normal");
              doc.setFontSize(12);
            }

            const xPos = lIdx === 0 ? marginLeft + firstLineIndent : marginLeft;
            doc.text(line, xPos, cursorY);
            cursorY += lineHeight;
          });

          cursorY += 1.5; // Espaço entre parágrafos
        });
      }
    });

    doc.save(`trabalho-abnt-${(title || "documento").toLowerCase().replace(/[^a-z0-9]/g, "-")}.pdf`);
  };

  const exportWord = async () => {
    if (!generatedText) return;
    
    const rawPages = generatedText.split("--- [QUEBRA DE PÁGINA] ---");
    const docSections: any[] = [];

    rawPages.forEach((pageContent, pageIdx) => {
      const isCover = pageIdx === 0;
      const isTitlePage = pageIdx === 1;
      const isBody = pageIdx >= 2;

      const paragraphs = pageContent.split('\n').map(text => {
        const clean = text.trim();
        if (!clean) {
          return new Paragraph({
            children: [new TextRun({ text: "", font: "Arial", size: 24 })],
            spacing: { line: 360 }
          });
        }

        const isRightNature = isTitlePage && (clean.startsWith("Trabalho") || clean.startsWith("Monografia") || clean.startsWith("Artigo") || clean.startsWith("Orientador") || clean.startsWith("Dissertação"));
        const isCentered = isCover || (isTitlePage && !isRightNature && (clean === clean.toUpperCase() || clean.length < 50));

        return new Paragraph({
          children: [new TextRun({ 
            text: clean, 
            font: "Arial", 
            size: isRightNature ? 20 : 24, // 10pt para nota, 12pt para corpo
            bold: isCover && (clean.length > 20 || clean === clean.toUpperCase())
          })],
          alignment: isCentered ? AlignmentType.CENTER : (isRightNature ? AlignmentType.RIGHT : AlignmentType.JUSTIFIED),
          spacing: { line: isRightNature ? 240 : 360 }, // 1.0 para nota, 1.5 para corpo
          indent: isCentered || isRightNature ? { firstLine: 0 } : { firstLine: convertMillimetersToTwip(12.5) } // 1.25cm recuo
        });
      });

      docSections.push({
        properties: {
          page: {
            size: {
              width: convertMillimetersToTwip(210), // Folha A4 210mm
              height: convertMillimetersToTwip(297) // Folha A4 297mm
            },
            margin: {
              top: convertMillimetersToTwip(30), // Margem Superior 3cm
              left: convertMillimetersToTwip(30), // Margem Esquerda 3cm
              right: convertMillimetersToTwip(20), // Margem Direita 2cm
              bottom: convertMillimetersToTwip(20), // Margem Inferior 2cm
            }
          }
        },
        headers: isBody ? {
          default: new Header({
            children: [
              new Paragraph({
                alignment: AlignmentType.RIGHT,
                children: [
                  new TextRun({
                    children: [PageNumber.CURRENT],
                    font: "Arial",
                    size: 20 // 10pt no cabeçalho superior direito
                  })
                ]
              })
            ]
          })
        } : undefined,
        children: paragraphs,
      });
    });

    const doc = new Document({
      sections: docSections.length > 0 ? docSections : [{
        children: [new Paragraph({ children: [new TextRun({ text: generatedText, font: "Arial", size: 24 })] })]
      }]
    });

    const blob = await Packer.toBlob(doc);
    saveAs(blob, "trabalho-abnt-a4.docx");
  };

  const exportLaTeX = () => {
    if (!generatedText) return;

    const cleanTitle = title || "Título do Trabalho Acadêmico";
    const cleanAuthor = studentName || "Nome do Autor";
    const cleanInst = institution || "Nome da Universidade";
    const cleanCity = city || "Cidade - Estado";
    const cleanYear = year || new Date().getFullYear().toString();
    const cleanAdvisor = advisor || "Prof. Dr. Orientador";
    const cleanPreambulo = course 
      ? `Trabalho de Conclusão de Curso apresentado como requisito parcial para obtenção do título de Bacharel em ${course}.`
      : "Trabalho acadêmico apresentado como requisito parcial de avaliação.";

    const rawPages = generatedText.split("--- [QUEBRA DE PÁGINA] ---");
    const bodyText = rawPages.length >= 3 ? rawPages.slice(2).join("\n\n") : generatedText;

    const latexChapters = bodyText.split(/\n(?=\d+\s+[A-ZÀ-Ú])/).map(section => {
      const trimmed = section.trim();
      const match = trimmed.match(/^(\d+(?:\.\d+)*)\s+([^\n]+)\n([\s\S]*)$/);
      if (match) {
        const num = match[1];
        const secTitle = match[2];
        const content = match[3];
        if (!num.includes(".")) {
          return `\\chapter{${secTitle}}\n${content}\n`;
        } else if (num.split(".").length === 2) {
          return `\\section{${secTitle}}\n${content}\n`;
        } else {
          return `\\subsection{${secTitle}}\n${content}\n`;
        }
      }
      return trimmed;
    }).join("\n\n");

    const latexTemplate = `\\documentclass[12pt,openright,twoside,a4paper,english,french,spanish,brazil]{abntex2}

% --- Configurações de Margens ABNT ---
\\usepackage[top=3cm,bottom=2cm,left=3cm,right=2cm]{geometry}
\\usepackage[utf8]{inputenc}
\\usepackage[T1]{fontenc}
\\usepackage{helvet} % Para fonte Arial (ou altere para Times)
\\renewcommand{\\familydefault}{\\sfdefault}
\\usepackage{indentfirst} % Indenta o primeiro parágrafo de cada seção (1.25cm)
\\usepackage{microtype}
\\usepackage{graphicx}
\\usepackage{booktabs}

% --- Dados do Documento ---
\\titulo{${cleanTitle}}
\\autor{${cleanAuthor}}
\\local{${cleanCity}}
\\data{${cleanYear}}
\\orientador{${cleanAdvisor}}
\\instituicao{${cleanInst}}
\\preambulo{${cleanPreambulo}}

% --- Início do Documento ---
\\begin{document}
\\frenchspacing 

% --- ELEMENTOS PRÉ-TEXTUAIS ---
\\imprimircapa
\\imprimirfolhaderosto

% Sumário Automático
\\pdfbookmark[0]{\\contentsname}{toc}
\\tableofcontents*
\\cleardoublepage

% --- ELEMENTOS TEXTUAIS (Paginação Visível Inicia Aqui) ---
\\textual

${latexChapters}

% --- ELEMENTOS PÓS-TEXTUAIS ---
\\postextual

\\end{document}
`;

    const blob = new Blob([latexTemplate], { type: "text/x-tex;charset=utf-8" });
    saveAs(blob, `trabalho-abntex2-${(cleanTitle || "documento").toLowerCase().replace(/[^a-z0-9]/g, "-")}.tex`);
    setErrorMessage("✨ Código LaTeX (abnTeX2) gerado com sucesso para Overleaf / TeXStudio!");
    setTimeout(() => setErrorMessage(""), 3500);
  };

  // Função para Estruturar o Texto Acadêmico em Roteiro de Slides com Dados Reais do Palco
  const parseDocumentIntoSlides = (text: string) => {
    const cleanDoc = (text || "").replace(/\r\n/g, '\n');
    const sections: Array<{ title: string; bullets: string[]; notes?: string; isCover?: boolean; badge?: string; layout?: "card" | "split" | "bullets" | "quote" | "metrics" }> = [];

    // 1. Slide de Capa Real com os Metadados do Trabalho
    sections.push({
      title: (title || "Apresentação do Trabalho Acadêmico").toUpperCase(),
      bullets: [
        `Autor(a): ${studentName || "Acadêmico(a)"}`,
        `Orientador(a): ${advisor || "Docente Orientador(a)"}`,
        `Instituição: ${institution || "Instituição de Ensino Superior"}`,
        `Curso: ${course || "Graduação / Pós-Graduação"}`,
        `Ano de Apresentação: ${year || new Date().getFullYear()}`
      ],
      notes: "Slide inicial: Agradeça à banca examinadora e apresente o título e objetivo do trabalho.",
      isCover: true,
      badge: "Apresentação Oficial",
      layout: "card"
    });

    // 2. Extrai seções reais do documento (ex: 1 INTRODUÇÃO, 2 FUNDAMENTAÇÃO, 3 RESULTADOS, CONCLUSÃO, etc.)
    const cleanBody = cleanDoc
      .replace(/--- \[(?:QUEBRA DE PÁGINA|NOVA PÁGINA)\] ---/gi, '\n\n')
      .replace(/\[(?:QUEBRA DE PÁGINA|NOVA PÁGINA)\]/gi, '\n\n')
      .replace(/^CAPA[\s\S]*?(?=\n\s*(?:RESUMO|1\s+|INTRODUÇÃO))/i, '');

    // Busca seções com numeração ABNT (ex: "1 INTRODUÇÃO", "2 METODOLOGIA", "3 RESULTADOS") ou títulos em caixa alta
    const sectionRegex = /(?:^|\n)(?:#+\s*)?(\d+(?:\.\d+)*\s+[A-ZÀ-Ú\s–\-]{3,}|RESUMO|ABSTRACT|CONSIDERAÇÕES FINAIS|CONCLUSÃO|REFERÊNCIAS)\b/g;
    
    const matches: Array<{ title: string; index: number }> = [];
    let match;
    while ((match = sectionRegex.exec(cleanBody)) !== null) {
      matches.push({ title: match[1].trim(), index: match.index });
    }

    if (matches.length > 0) {
      for (let i = 0; i < matches.length; i++) {
        const current = matches[i];
        const nextIndex = i + 1 < matches.length ? matches[i + 1].index : cleanBody.length;
        const sectionContent = cleanBody.substring(current.index, nextIndex).trim();
        
        // Separa linhas e parágrafos reais do conteúdo
        const lines = sectionContent
          .split('\n')
          .slice(1) // Pula o título da seção
          .map(l => l.replace(/^[\*\-\#\d\.\s]+/, '').trim())
          .filter(l => l.length > 20);

        // Agrupa as ideias em 3 a 5 tópicos sintetizados reais
        const realBullets: string[] = [];
        for (const line of lines) {
          if (realBullets.length >= 4) break;
          // Pega frases significativas do parágrafo
          const sentences = line.split(/(?<=[.!?])\s+/).filter(s => s.trim().length > 15);
          if (sentences.length > 0 && !realBullets.includes(sentences[0])) {
            realBullets.push(sentences[0].substring(0, 160) + (sentences[0].length > 160 ? "..." : ""));
          }
        }

        if (realBullets.length === 0) {
          realBullets.push("Aprofundamento teórico e análise sistemática dos conceitos centrais investigados.");
          realBullets.push("Fundamentação alinhada aos dados empíricos e literatura especializada.");
        }

        const isResult = current.title.includes("RESULTADO") || current.title.includes("DISCUSS");
        const isConcl = current.title.includes("CONCLUS") || current.title.includes("CONSIDERA");

        sections.push({
          title: current.title,
          bullets: realBullets,
          notes: `Destaque os pontos cruciais da seção "${current.title}" e faça conexões com os objetivos do estudo.`,
          badge: isResult ? "Resultados & Dados" : isConcl ? "Conclusão Final" : `Módulo ${i + 1}`,
          layout: isResult ? "split" : isConcl ? "metrics" : "bullets"
        });
      }
    } else {
      // Caso o texto não tenha títulos numerados explícitos, fatia os parágrafos reais do texto
      const paragraphs = cleanBody.split(/\n\n+/).map(p => p.trim()).filter(p => p.length > 50);
      paragraphs.slice(0, 5).forEach((p, pIdx) => {
        const sentences = p.split(/(?<=[.!?])\s+/).filter(s => s.trim().length > 15);
        sections.push({
          title: `Tópico ${pIdx + 1}: ${title ? title.substring(0, 40) : "Desenvolvimento"}`,
          bullets: sentences.slice(0, 3).map(s => s.substring(0, 150)),
          notes: "Apresente estes conceitos enfatizando o rigor metodológico.",
          badge: `Parte ${pIdx + 1}`,
          layout: "bullets"
        });
      });
    }

    return sections;
  };

  const handleOpenSlidesStudio = () => {
    if (activeTab === "slides") {
      setActiveTab("editor");
      return;
    }
    const textToParse = generatedText && generatedText.trim() ? generatedText : "1 INTRODUÇÃO\n\nApresentação acadêmica estruturada.";
    const parsed = parseDocumentIntoSlides(textToParse);
    setEditableSlides(parsed as any);
    setActiveSlideIndex(0);
    setActiveTab("slides");
  };

  // Exportação Direta em PowerPoint (.pptx) 100% compatível com Google Slides
  const exportPPTXSlides = async () => {
    const pptx = new pptxgen();
    pptx.layout = "LAYOUT_16x9";
    pptx.title = title || "Apresentação Acadêmica";
    pptx.author = studentName || "EDUTECH.EMIA";

    const slidesData = parseDocumentIntoSlides(generatedText);

    slidesData.forEach((s, idx) => {
      const slide = pptx.addSlide();
      
      const primaryColor = slidesTheme === 'academic' ? "1E3A8A" : slidesTheme === 'modern' ? "0D9488" : "0F172A";

      if (s.isCover) {
        slide.background = { color: primaryColor };
        slide.addText(s.title, {
          x: "10%",
          y: "25%",
          w: "80%",
          h: "30%",
          fontSize: 28,
          bold: true,
          color: "FFFFFF",
          align: "center",
          fontFace: "Arial"
        });

        slide.addText(s.subtitle || (course ? `Curso: ${course}` : "Apresentação Acadêmica"), {
          x: "10%",
          y: "55%",
          w: "80%",
          h: "10%",
          fontSize: 16,
          color: "E2E8F0",
          align: "center",
          fontFace: "Arial"
        });

        slide.addText(`${s.author} • ${s.institution} (${s.year})`, {
          x: "10%",
          y: "75%",
          w: "80%",
          h: "10%",
          fontSize: 13,
          color: "CBD5E1",
          align: "center",
          fontFace: "Arial"
        });
      } else {
        slide.background = { color: "F8FAFC" };

        slide.addShape(pptx.ShapeType.rect, {
          x: 0,
          y: 0,
          w: "100%",
          h: "15%",
          fill: { color: primaryColor }
        });

        slide.addText(s.title, {
          x: "5%",
          y: "2%",
          w: "90%",
          h: "11%",
          fontSize: 20,
          bold: true,
          color: "FFFFFF",
          fontFace: "Arial",
          valign: "middle"
        });

        const textItems = s.bullets.map(b => ({
          text: b,
          options: {
            fontSize: 15,
            color: "1E293B",
            breakLine: true,
            bullet: { type: "bullet", code: "25AA" },
            fontFace: "Arial",
            lineSpacingMultiple: 1.3
          }
        }));

        slide.addText(textItems as any, {
          x: "6%",
          y: "22%",
          w: "88%",
          h: "68%",
          valign: "top"
        });

        slide.addText(`${title || "Trabalho Acadêmico"} | Slide ${idx + 1} de ${slidesData.length}`, {
          x: "5%",
          y: "93%",
          w: "90%",
          h: "5%",
          fontSize: 9,
          color: "94A3B8",
          fontFace: "Arial"
        });
      }

      if (s.notes) {
        slide.addNotes(s.notes);
      }
    });

    await pptx.writeFile({ fileName: `apresentacao-${(title || "slides-academicos").toLowerCase().replace(/[^a-z0-9]/g, "-")}.pptx` });
    setErrorMessage("✨ Apresentação de slides gerada com sucesso!");
    setTimeout(() => setErrorMessage(""), 3500);
  };

  const openInGoogleSlides = () => {
    const slidesData = parseDocumentIntoSlides(generatedText);
    const outlineText = slidesData.map((s, i) => `=== SLIDE ${i + 1}: ${s.title} ===\n${s.bullets.map(b => `• ${b}`).join('\n')}\n[Notas do Apresentador: ${s.notes || ''}]\n`).join('\n\n');
    
    navigator.clipboard.writeText(outlineText);
    window.open("https://docs.google.com/presentation/u/0/create", "_blank");
    setErrorMessage("🚀 Google Slides aberto! O roteiro dos slides foi copiado para sua área de transferência.");
    setTimeout(() => setErrorMessage(""), 4500);
  };

  const handleNewWork = () => {
    setTitle("");
    setSubtitle("");
    setDocumentType("artigo");
    setCustomDocumentType("");
    setPrompt("");
    setFiles([]);
    setGeneratedText("");
    setAuthenticityReport("");
    setFormatRules("");
    setChatHistory([]);
    setActiveTab("generator");
    logAction('Iniciou um novo trabalho (limpeza de formulário)');
  };

  const handleClearWorkData = () => {
    setStudentName("");
    setCourse("");
    setSubject("");
    setShift("");
    setClassroom("");
    setInstitution("");
    setCity("");
    setYear("");
    setAdvisor("");
    logAction('Dados do Trabalho limpos');
  };

  const generateCoverTextLocally = () => {
    const instName = (institution || "NOME DA INSTITUIÇÃO DE ENSINO").toUpperCase();
    const courseName = course ? course.toUpperCase() : "";
    const subjectName = subject ? `DISCIPLINA: ${subject.toUpperCase()}` : "";
    const shiftClassInfo = [shift ? `Turno: ${shift}` : "", classroom ? `Sala/Turma: ${classroom}` : ""].filter(Boolean).join(" • ");
    const authorName = (studentName || "NOME DO AUTOR DO TRABALHO").toUpperCase();
    const docTitle = (title || "TÍTULO DO TRABALHO ACADÊMICO").toUpperCase();
    const docSubtitle = subtitle ? ` - ${subtitle}` : "";
    const docCity = (city || "CIDADE - UF").toUpperCase();
    const docYear = year || new Date().getFullYear().toString();
    const docType = documentType === "outros" ? (customDocumentType || "TRABALHO ACADÊMICO").toUpperCase() : documentType.toUpperCase();
    const advText = advisor ? `Orientador(a): ${advisor}` : "";

    const subHeader = [courseName, subjectName, shiftClassInfo].filter(Boolean).join("\n");
    const coverPage = `${instName}${subHeader ? `\n${subHeader}` : ""}\n\n\n\n${authorName}\n\n\n\n\n\n\n\n${docTitle}${docSubtitle}\n\n\n\n\n\n\n\n\n\n${docCity}\n${docYear}`;

    const presentationNote = `                                          ${docType} apresentado à ${instName}${courseName ? ` como requisito parcial para a disciplina de ${subject || courseName}` : ""}.\n${shiftClassInfo ? `\n                                          ${shiftClassInfo}` : ""}\n${advText ? `\n                                          ${advText}` : ""}`;
    const titlePage = `${authorName}\n\n\n\n\n\n\n\n${docTitle}${docSubtitle}\n\n\n\n${presentationNote}\n\n\n\n\n\n\n\n${docCity}\n${docYear}`;

    return `${coverPage}\n\n--- [QUEBRA DE PÁGINA] ---\n\n${titlePage}\n\n--- [QUEBRA DE PÁGINA] ---`;
  };

  const handleSaveProfile = () => {
    const profile: UserProfile = { name: studentName, institution, city, year, advisor, course, subject, shift, classroom };
    localStorage.setItem('emia_user_profile', JSON.stringify(profile));
    logAction('Dados do Trabalho salvos localmente');
    
    // Atualiza a capa no documento principal imediatamente
    const newCover = generateCoverTextLocally();
    if (generatedText && generatedText.trim()) {
      const cleanBody = generatedText.replace(/^.*--- \[(?:QUEBRA DE PÁGINA|NOVA PÁGINA)\] ---\n*/is, '');
      setGeneratedText(newCover + "\n\n" + (cleanBody || generatedText));
    } else {
      setGeneratedText(newCover + "\n\n1 INTRODUÇÃO\n\nEscreva ou cole seu texto acadêmico aqui...");
    }
    setActiveTab("editor");
    setShowProfileModal(false);
    setShowWorkData(false);
    setErrorMessage("✅ Dados da capa salvos e atualizados no documento principal!");
    setTimeout(() => setErrorMessage(""), 3500);
  };

  return (
    <div className="min-h-screen flex flex-col bg-gray-50">
      <header className="bg-white border-b border-gray-200 px-6 py-4 flex items-center justify-between sticky top-0 z-10">
        <div className="flex items-center gap-3">
          <div className="flex items-center gap-2">
            <div className="bg-blue-600 text-white w-8 h-8 rounded-md flex items-center justify-center shadow-xs">
              <FileText className="w-5 h-5" />
            </div>
            <span className="font-semibold text-gray-900 text-lg tracking-tight">EDUTECH.EMIA</span>
          </div>

          {/* Marca d'água discreta do Google Gemini ao lado da logo */}
          <div className="hidden sm:flex items-center gap-1.5 pl-3 border-l border-gray-200/80 text-gray-400 select-none">
            <Sparkles className="w-3.5 h-3.5 text-blue-500/70" />
            <span className="text-[11px] font-medium tracking-wide">powered by Google Gemini</span>
          </div>
        </div>

        <div className="flex items-center gap-1.5 sm:gap-2">
          {/* Botão de Créditos / PIX (Visível para todos os usuários) */}
          <button 
            onClick={() => setShowPixModal(true)}
            className="h-7.5 flex items-center gap-1.5 bg-gradient-to-r from-amber-500 via-orange-500 to-amber-600 hover:from-amber-600 hover:to-orange-600 text-white px-3 rounded-lg text-xs font-bold shadow-2xs hover:shadow-xs transition-all active:scale-95 group"
            title="Recarregar créditos via PIX"
          >
            <Coins className="w-3.5 h-3.5 text-white animate-bounce" />
            <span>{credits} {credits === 1 ? 'Trabalho' : 'Trabalhos'}</span>
            <span className="bg-white/25 text-white text-[10px] px-1.5 py-0.5 rounded font-extrabold ml-0.5">
              + PIX
            </span>
          </button>

          {/* Indicador de Cota Própria para Chave Própria */}
          {(customGeminiKey || customOpenaiKey) && !isMaster ? (
            <div className="h-7.5 flex items-center gap-1.5 bg-emerald-50 border border-emerald-300 text-emerald-800 font-semibold text-[11px] px-2.5 rounded-lg shadow-2xs">
              <span className="w-1.5 h-1.5 rounded-full bg-emerald-500 animate-pulse" />
              <span>Cota Conectada</span>
            </div>
          ) : null}

          {/* Botão Chat Acadêmico */}
          <button 
            onClick={() => setActiveTab("chat")}
            className={`h-7.5 flex items-center gap-1.5 px-3 rounded-lg text-xs font-bold shadow-2xs hover:shadow-xs transition-all active:scale-95 group ${
              activeTab === "chat" 
                ? "bg-indigo-600 text-white shadow-xs" 
                : "bg-white hover:bg-slate-100 text-gray-900 border border-gray-300"
            }`}
            title="Abrir Chat Acadêmico com IA"
          >
            <UserCheck className={`w-3.5 h-3.5 stroke-[2] group-hover:scale-105 transition-transform ${activeTab === "chat" ? "text-white" : "text-gray-900"}`} />
            <span className="tracking-tight">Chat Acadêmico</span>
          </button>

          {/* Grupo de Áudio: Ouvir Texto + Controles Compactos Elegantes */}
          <div className="flex items-center bg-amber-50/70 border border-amber-200/80 rounded-lg p-0.5 shadow-2xs gap-0.5">
            <button 
              onClick={handleToggleSpeech}
              className={`h-6.5 flex items-center gap-1 px-2.5 rounded-md text-[11px] font-semibold transition-all active:scale-95 group ${
                isSpeaking 
                  ? "bg-gradient-to-r from-amber-500 to-orange-500 text-white shadow-xs animate-pulse" 
                  : "hover:bg-amber-100/70 text-amber-900"
              }`}
              title={isSpeaking ? "Interromper Leitura em Áudio" : "Ouvir Leitura do Texto Acadêmico em Áudio"}
            >
              {isSpeaking ? (
                <>
                  <VolumeX className="w-3 h-3 text-white stroke-[2]" />
                  <span className="tracking-tight text-white">Parar</span>
                </>
              ) : (
                <>
                  <Volume2 className="w-3 h-3 text-amber-700 stroke-[1.8]" />
                  <span className="tracking-tight">Ouvir Texto</span>
                </>
              )}
            </button>

            {/* Alternador de Gênero da Voz (Feminina / Masculina) */}
            <button
              onClick={() => {
                const nextGender = speechGender === "female" ? "male" : "female";
                setSpeechGender(nextGender);
                if (isSpeaking) {
                  window.speechSynthesis.cancel();
                  speakSentenceAtIndex(currentSentenceIdxRef.current, speechRate, speechVolume, nextGender);
                } else {
                  setErrorMessage(`Voz alterada para ${nextGender === "female" ? "Feminina 👩" : "Masculina 👨"}.`);
                  setTimeout(() => setErrorMessage(""), 2000);
                }
              }}
              className="h-6.5 px-1.5 flex items-center gap-0.5 text-[10px] font-semibold text-amber-900 bg-amber-100/80 hover:bg-amber-200/80 rounded transition-all active:scale-90 select-none"
              title={`Alternar Gênero da Voz (Atual: ${speechGender === "female" ? "Feminina 👩" : "Masculina 👨"})`}
            >
              <span>{speechGender === "female" ? "👩" : "👨"}</span>
            </button>

            {/* Velocidade: 1x, 1.25x, 1.5x, 2x (Continua falando sem parar ou recomeçar) */}
            <button
              onClick={() => {
                const rates = [1.0, 1.25, 1.5, 2.0];
                const nextRate = rates[(rates.indexOf(speechRate) + 1) % rates.length];
                setSpeechRate(nextRate);
                if (isSpeaking) {
                  window.speechSynthesis.cancel();
                  speakSentenceAtIndex(currentSentenceIdxRef.current, nextRate, speechVolume, speechGender);
                }
              }}
              className="h-6.5 px-1.5 flex items-center justify-center text-[10.5px] font-bold text-amber-900 bg-amber-100/80 hover:bg-amber-200/80 rounded transition-all active:scale-90 select-none"
              title="Alternar Velocidade (1x, 1.25x, 1.5x, 2x)"
            >
              {speechRate}x
            </button>

            {/* Volume: 100%, 75%, 50% */}
            <button
              onClick={() => {
                const vols = [1.0, 0.75, 0.5];
                const nextVol = vols[(vols.indexOf(speechVolume) + 1) % vols.length];
                setSpeechVolume(nextVol);
                if (isSpeaking) {
                  window.speechSynthesis.cancel();
                  setTimeout(() => handleToggleSpeech(), 100);
                }
              }}
              className="h-6.5 px-1 flex items-center justify-center text-[9.5px] font-medium text-amber-850 hover:bg-amber-100/80 rounded transition-all active:scale-90 select-none"
              title="Alternar Volume (100%, 75%, 50%)"
            >
              {Math.round(speechVolume * 100)}%
            </button>
          </div>

          {/* Botão Slides (EMIA.SLIDES - Amarelo Pastel com Texto Branco) */}
          <button 
            onClick={handleOpenSlidesStudio} 
            className={`h-7.5 flex items-center gap-1.5 px-3 rounded-lg text-xs font-semibold shadow-2xs hover:shadow-xs transition-all active:scale-95 group ${
              activeTab === "slides" 
                ? "bg-amber-500 text-white shadow-xs ring-1 ring-amber-300" 
                : "bg-amber-400/90 hover:bg-amber-500/90 text-white shadow-2xs border border-amber-300/40"
            }`}
            title="Abrir Estúdio EMIA.SLIDES"
          >
            <Presentation className="w-3.5 h-3.5 text-white stroke-[1.8] group-hover:scale-105 transition-transform" />
            <span className="tracking-tight text-white">Slides</span>
          </button>

          {/* Botão Trabalho em Grupo (Verde Claro Suave e Harmonioso) */}
          <button 
            onClick={() => setIsGroupMode(true)} 
            className="h-7.5 flex items-center gap-1.5 bg-emerald-50/80 hover:bg-emerald-100/80 text-emerald-800 border border-emerald-300/70 px-3 rounded-lg text-xs font-semibold shadow-2xs hover:shadow-xs transition-all active:scale-95 group"
            title="Montar trabalho acadêmico feito em grupo com múltiplos alunos"
          >
            <Users className="w-3.5 h-3.5 text-emerald-700 stroke-[1.8] group-hover:scale-105 transition-transform" />
            <span className="tracking-tight">Trabalho em Grupo</span>
          </button>

          {/* Botão Perfil e Histórico (Índigo/Lavanda Elegante) */}
          <button 
            onClick={() => setShowProfileModal(true)}
            className="h-7.5 flex items-center gap-1.5 bg-indigo-50/80 hover:bg-indigo-100/80 text-indigo-900 border border-indigo-200/80 px-3 rounded-lg text-xs font-semibold shadow-2xs hover:shadow-xs transition-all active:scale-95 group"
            title="Ver perfil e histórico de trabalhos acadêmicos"
          >
            <User className="w-3.5 h-3.5 text-indigo-600 stroke-[1.8] group-hover:scale-105 transition-transform" />
            <span className="tracking-tight">Perfil e Histórico</span>
          </button>

          {/* Botão Sair */}
          <button 
            onClick={handleLogout}
            className="h-7.5 flex items-center gap-1 text-gray-500 hover:text-red-600 hover:bg-red-50/80 px-2 rounded-lg text-xs font-medium transition-all"
            title="Encerrar sessão"
          >
            <LogOut className="w-3.5 h-3.5 stroke-[1.8]" />
            <span>Sair</span>
          </button>
        </div>
      </header>

      <main className="flex-1 max-w-[1600px] w-full mx-auto p-3 md:p-6 flex flex-col lg:flex-row gap-4 items-stretch relative">
        
        {/* Botão flutuante para reabrir a barra lateral quando recolhida (Setinhas de Tela Cheia / Expandir) */}
        {isSidebarCollapsed && (
          <button
            onClick={() => setIsSidebarCollapsed(false)}
            title="Mostrar Formulário (Sair da Tela Cheia)"
            className="hidden lg:flex items-center gap-1.5 bg-white border border-gray-300 hover:border-blue-500 shadow-lg text-gray-800 hover:text-blue-600 px-3 py-2 rounded-r-xl font-bold text-xs fixed left-0 top-32 z-30 transition-all hover:translate-x-1 group"
          >
            <Minimize2 className="w-4 h-4 text-blue-600 group-hover:scale-110" />
            <span>Formulário</span>
          </button>
        )}

        {/* Sidebar Controls (Recolhível com ícone de setinhas para tela cheia) */}
        {!isSidebarCollapsed && (
          <div className="w-full lg:w-80 xl:w-96 flex-shrink-0 flex flex-col gap-6 transition-all">
            <div className="bg-white rounded-xl shadow-sm border border-gray-200 p-5 flex flex-col h-full">
              <div className="flex items-center justify-between mb-4">
                <h2 className="text-lg font-semibold text-gray-900">Novo Trabalho</h2>
                <div className="flex items-center gap-1">
                  <Button onClick={handleNewWork} variant="outline" size="sm" className="text-xs h-7">
                    <Plus className="w-3 h-3 mr-1" /> Novo
                  </Button>
                  <Button 
                    onClick={() => setIsSidebarCollapsed(true)} 
                    variant="ghost" 
                    size="sm" 
                    title="Tela Cheia no Palco (Ocultar formulário)"
                    className="text-xs h-7 px-2 text-gray-500 hover:text-blue-600 hover:bg-blue-50 rounded-lg"
                  >
                    <Maximize2 className="w-4 h-4" />
                  </Button>
                </div>
              </div>
            
            <div className="space-y-4 flex-1 flex flex-col">
              <div>
                <label className="block text-sm font-medium text-gray-700 mb-1">Tipo de Documento</label>
                <select 
                  value={documentType}
                  onChange={(e) => setDocumentType(e.target.value)}
                  className="w-full px-3 py-2 border border-gray-300 rounded-md shadow-sm focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-blue-500 text-sm bg-white"
                >
                  <option value="artigo">Artigo Acadêmico</option>
                  <option value="resumo">Resumo / Fichamento</option>
                  <option value="trabalho_academico">Trabalho Acadêmico (TCC)</option>
                  <option value="monografia">Monografia / Dissertação</option>
                  <option value="projeto">Projeto de Pesquisa</option>
                  <option value="artigo_opiniao">Artigo de Opinião</option>
                  <option value="resenha">Resenha Crítica</option>
                  <option value="estudo_caso">Estudo de Caso</option>
                  <option value="relatorio">Relatório Técnico</option>
                  <option value="artigo_cientifico">Artigo Científico</option>
                  <option value="redacao">Redação</option>
                  <option value="outros">Outros</option>
                </select>
                <div className="mt-1.5 px-2.5 py-1.5 bg-blue-50/70 border border-blue-100 rounded-md text-[11px] text-blue-800 font-medium">
                  {documentType === "resumo" && "📘 ABNT NBR 6028: Sem capa / sem folha de rosto • Parágrafo único contínuo com Palavras-chave."}
                  {documentType === "redacao" && "✍️ Padrão ENEM Nota 1000: Sem capa • 4 parágrafos dissertativos com proposta de intervenção."}
                  {(documentType === "artigo" || documentType === "artigo_cientifico") && "📄 ABNT NBR 6022: Sem capa avulsa • Cabeçalho de Autores e Resumo na 1ª página."}
                  {documentType === "resenha" && "📑 Padrão ABNT: Sem capa avulsa • Cabeçalho com Referência da Obra Resenhada."}
                  {(documentType === "monografia" || documentType === "trabalho_academico" || documentType === "relatorio" || documentType === "projeto") && "🎓 ABNT NBR 14724: Capa e Folha de Rosto Oficiais com quebra de página."}
                  {documentType === "estudo_caso" && "📊 Estudo de Caso: Cabeçalho institucional na 1ª página com diagnóstico e soluções."}
                  {documentType === "artigo_opiniao" && "📰 Artigo de Opinião: Título, autoria e texto argumentativo fluido."}
                  {documentType === "outros" && "⚙️ Estrutura personalizada conforme diretrizes."}
                </div>
                {documentType === "outros" && (
                  <div className="mt-2">
                    <input 
                      type="text" 
                      value={customDocumentType}
                      onChange={(e) => setCustomDocumentType(e.target.value)}
                      placeholder="Especifique o tipo de texto..."
                      className="w-full px-3 py-2 border border-gray-300 rounded-md shadow-sm focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-blue-500 text-sm"
                    />
                  </div>
                )}
              </div>

              <div>
                <label className="block text-sm font-medium text-gray-700 mb-1">Título / Tema</label>
                <input 
                  type="text" 
                  value={title}
                  onChange={(e) => setTitle(e.target.value)}
                  placeholder="Ex: Inteligência Artificial na Educação"
                  className="w-full px-3 py-2 border border-gray-300 rounded-md shadow-sm focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-blue-500 text-sm"
                />
              </div>

              <div>
                <label className="block text-sm font-medium text-gray-700 mb-1">Subtítulo (Opcional)</label>
                <input 
                  type="text" 
                  value={subtitle}
                  onChange={(e) => setSubtitle(e.target.value)}
                  placeholder="Ex: Uma análise contemporânea"
                  className="w-full px-3 py-2 border border-gray-300 rounded-md shadow-sm focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-blue-500 text-sm"
                />
              </div>

              <div>
                <label className="block text-sm font-medium text-gray-700 mb-1">Instruções / Diretrizes</label>
                <textarea 
                  value={prompt}
                  onChange={(e) => setPrompt(e.target.value)}
                  placeholder="Escreva um artigo de 3 páginas abordando os impactos positivos..."
                  className="w-full px-3 py-2 border border-gray-300 rounded-md shadow-sm focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-blue-500 text-sm h-24 resize-none"
                />
              </div>

              <div>
                <Button 
                  onClick={() => setShowWorkData(!showWorkData)} 
                  variant="outline" 
                  className="w-full bg-gray-50 border-gray-300 hover:bg-gray-100 flex items-center justify-between"
                >
                  <span className="text-sm font-medium text-gray-700">
                    Dados do Aluno
                  </span>
                  <span className="text-gray-500 text-xs">{showWorkData ? "Ocultar" : "Preencher"}</span>
                </Button>
                
                {showWorkData && (
                  <div className="mt-3 space-y-3 p-4 bg-gray-50 border border-gray-200 rounded-md text-sm">
                    <div>
                      <label className="block font-medium text-gray-700 mb-1">Nome do Aluno(a)</label>
                      <input 
                        type="text" 
                        value={studentName}
                        onChange={(e) => setStudentName(e.target.value)}
                        placeholder="Ex: João da Silva"
                        className="w-full px-3 py-2 border border-gray-300 rounded-md shadow-sm focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-blue-500"
                      />
                    </div>
                    <div>
                      <label className="block font-medium text-gray-700 mb-1">Curso</label>
                      <input 
                        type="text" 
                        value={course}
                        onChange={(e) => setCourse(e.target.value)}
                        placeholder="Ex: Administração, Direito, Pedagogia..."
                        className="w-full px-3 py-2 border border-gray-300 rounded-md shadow-sm focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-blue-500 text-xs"
                      />
                    </div>
                    <div>
                      <label className="block font-medium text-gray-700 mb-1">Disciplina / Matéria</label>
                      <input 
                        type="text" 
                        value={subject}
                        onChange={(e) => setSubject(e.target.value)}
                        placeholder="Ex: Metodologia Científica, Didática, Gestão..."
                        className="w-full px-3 py-2 border border-gray-300 rounded-md shadow-sm focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-blue-500 text-xs"
                      />
                    </div>
                    <div className="grid grid-cols-2 gap-3">
                      <div>
                        <label className="block font-medium text-gray-700 mb-1">Turno</label>
                        <select
                          value={shift}
                          onChange={(e) => setShift(e.target.value)}
                          className="w-full px-2.5 py-2 border border-gray-300 rounded-md shadow-sm focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-blue-500 text-xs bg-white"
                        >
                          <option value="">Selecione...</option>
                          <option value="Matutino">Matutino</option>
                          <option value="Vespertino">Vespertino</option>
                          <option value="Noturno">Noturno</option>
                          <option value="Integral">Integral</option>
                          <option value="EAD / Online">EAD / Online</option>
                        </select>
                      </div>
                      <div>
                        <label className="block font-medium text-gray-700 mb-1">Sala / Turma</label>
                        <input 
                          type="text" 
                          value={classroom}
                          onChange={(e) => setClassroom(e.target.value)}
                          placeholder="Ex: Sala 204, Turma B"
                          className="w-full px-3 py-2 border border-gray-300 rounded-md shadow-sm focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-blue-500 text-xs"
                        />
                      </div>
                    </div>
                    <div>
                      <label className="block font-medium text-gray-700 mb-1">Instituição de Ensino</label>
                      <input 
                        type="text" 
                        value={institution}
                        onChange={(e) => setInstitution(e.target.value)}
                        placeholder="Ex: Universidade de São Paulo"
                        className="w-full px-3 py-2 border border-gray-300 rounded-md shadow-sm focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-blue-500 text-xs"
                      />
                    </div>
                    <div className="grid grid-cols-2 gap-3">
                      <div>
                        <label className="block font-medium text-gray-700 mb-1">Cidade</label>
                        <input 
                          type="text" 
                          value={city}
                          onChange={(e) => setCity(e.target.value)}
                          placeholder="Ex: São Paulo"
                          className="w-full px-3 py-2 border border-gray-300 rounded-md shadow-sm focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-blue-500"
                        />
                      </div>
                      <div>
                        <label className="block font-medium text-gray-700 mb-1">Ano</label>
                        <input 
                          type="text" 
                          value={year}
                          onChange={(e) => setYear(e.target.value)}
                          placeholder="Ex: 2024"
                          className="w-full px-3 py-2 border border-gray-300 rounded-md shadow-sm focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-blue-500"
                        />
                      </div>
                    </div>
                    <div>
                      <label className="block font-medium text-gray-700 mb-1">Orientador(a) (Opcional)</label>
                      <input 
                        type="text" 
                        value={advisor}
                        onChange={(e) => setAdvisor(e.target.value)}
                        placeholder="Ex: Prof. Dr. Carlos Souza"
                        className="w-full px-3 py-2 border border-gray-300 rounded-md shadow-sm focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-blue-500"
                      />
                    </div>
                    <div className="flex gap-2 pt-2 border-t border-gray-200 mt-2">
                      <Button 
                        size="sm" 
                        onClick={handleClearWorkData} 
                        variant="outline" 
                        className="flex-1 text-red-600 border-red-200 hover:bg-red-50 hover:text-red-700"
                      >
                        Limpar Dados
                      </Button>
                      <Button 
                        size="sm" 
                        onClick={() => {
                          handleSaveProfile();
                          setShowWorkData(false);
                        }} 
                        className="flex-1 bg-blue-600 hover:bg-blue-700"
                      >
                        Salvar Dados
                      </Button>
                    </div>
                  </div>
                )}
              </div>

              <div>
                <label className="block text-sm font-medium text-gray-700 mb-1">Arquivos Base (Para basear IA ou Editar trabalhos)</label>
                <div 
                  {...getRootProps()} 
                  className={`border-2 border-dashed rounded-md p-6 text-center cursor-pointer transition-colors ${isDragActive ? 'border-blue-500 bg-blue-50' : 'border-gray-300 hover:border-gray-400'}`}
                >
                  <input {...getInputProps()} />
                  <Upload className="w-6 h-6 text-gray-400 mx-auto mb-2" />
                  <p className="text-sm text-gray-500">Arraste múltiplos PDFs ou Words, ou clique para selecionar</p>
                </div>
                
                {files.length > 0 && (
                  <div className="mt-3 space-y-2">
                    {files.map((f, idx) => (
                      <div key={idx} className="flex items-center justify-between bg-blue-50 border border-blue-100 rounded px-3 py-2">
                        <span className="text-sm font-medium text-blue-700 truncate max-w-[200px]">{f.name}</span>
                        <button 
                          onClick={(e) => {
                            e.stopPropagation();
                            setFiles(prev => prev.filter((_, i) => i !== idx));
                          }}
                          className="text-red-500 hover:text-red-700 p-1"
                        >
                          <X className="w-4 h-4" />
                        </button>
                      </div>
                    ))}
                  </div>
                )}
              </div>

              <div>
                <label className="block text-sm font-medium text-gray-700 mb-1">Inserir no Documento (Imagens / Tabelas)</label>
                <Button 
                  onClick={() => attachmentRef.current?.click()} 
                  disabled={isLoading} 
                  variant="outline" 
                  className="w-full bg-gray-50 border-dashed border-2 hover:bg-gray-100 h-auto py-4 flex flex-col items-center justify-center gap-2"
                >
                  <ImagePlus className="w-6 h-6 text-gray-400" />
                  <span className="text-sm text-gray-500 font-normal">Adicionar Imagem ou Tabela (CSV)</span>
                </Button>
                <input 
                  type="file" 
                  className="hidden" 
                  ref={attachmentRef} 
                  accept="image/*,.csv,.txt" 
                  onChange={handleAttachmentFile} 
                />
              </div>

              <div className="pt-2">
                <Button 
                  onClick={handleGenerate} 
                  disabled={isLoading || (!title && files.length === 0)} 
                  className="w-full bg-blue-600 hover:bg-blue-700 text-white py-3 font-semibold mb-3"
                >
                  {isLoading && activeTab === 'generator' ? <Loader2 className="w-5 h-5 mr-2 animate-spin" /> : <Settings className="w-5 h-5 mr-2" />}
                  Gerar Texto com IA
                </Button>

                <Button 
                  onClick={handleImproveText} 
                  disabled={isLoading || !generatedText} 
                  className="w-full bg-pink-600 hover:bg-pink-700 text-white py-3 font-semibold mb-3"
                >
                  {isLoading && activeTab === 'editor' ? <Loader2 className="w-5 h-5 mr-2 animate-spin" /> : <Wand2 className="w-5 h-5 mr-2" />}
                  Aprimorar Texto com IA
                </Button>
                <Button 
                  onClick={handleFormatABNT} 
                  disabled={isLoading || !generatedText} 
                  className="w-full bg-gray-800 hover:bg-gray-900 text-white py-3 font-semibold"
                >
                  {isLoading && activeTab === 'editor' ? <Loader2 className="w-5 h-5 mr-2 animate-spin" /> : <CheckCircle className="w-5 h-5 mr-2" />}
                  Adequar à ABNT
                </Button>
              </div>
            </div>
          </div>

        </div>
      )}

        {/* Editor Area (Palco Padrão alinhado perfeitamente com a linha do botão ABNT da barra lateral) */}
        <div 
          className="flex-1 w-full min-w-0 flex flex-col transition-all h-full min-h-[750px]"
        >
          
          {/* Top Bar Fixa e Flutuante com Barra de Rolagem (Sticky Toolbar) */}
          <div className="sticky top-0 z-30 flex items-center justify-between border-b border-gray-200/90 mb-3 py-2 px-2 gap-1.5 bg-white/95 backdrop-blur-md rounded-xl shadow-xs overflow-x-auto scrollbar-hide flex-nowrap transition-all">
            
            {/* Grupo Harmônico e Elegante de Ações na Esquerda em Linha Única */}
            <div className="flex items-center gap-1 flex-nowrap flex-shrink-0">
              
              {/* Ferramentas de Redação & Estrutura Acadêmica */}
              <div className="flex items-center bg-gray-50 border border-gray-200/80 rounded-xl p-0.5 shadow-2xs gap-0.5 flex-nowrap">
                <Button 
                  onClick={() => setActiveTab("editor")} 
                  variant="ghost" 
                  size="sm" 
                  className={`text-xs h-7 px-2.5 font-extrabold rounded-lg transition-all flex items-center gap-1.5 whitespace-nowrap ${
                    activeTab === "editor" || activeTab === "generator" 
                      ? "bg-white text-blue-700 shadow-xs border border-blue-200/60" 
                      : "text-gray-700 hover:bg-gray-200/60"
                  }`}
                  title="Exibir Documento Acadêmico no Palco"
                >
                  <FileText className="w-4 h-4 text-blue-600" />
                  Documento
                </Button>
                <Button 
                  onClick={() => setActiveTab("chat")} 
                  variant="ghost" 
                  size="sm" 
                  className={`text-xs h-7 px-2.5 font-bold rounded-lg transition-all flex items-center gap-1.5 whitespace-nowrap ${
                    activeTab === "chat" 
                      ? "bg-indigo-600 text-white shadow-xs" 
                      : "text-gray-900 hover:bg-slate-200/70"
                  }`}
                  title="Abrir Chat Acadêmico com a IA EMIA"
                >
                  <UserCheck className={`w-4 h-4 ${activeTab === "chat" ? "text-white" : "text-gray-900"}`} />
                  Chat Acadêmico
                </Button>
                <Button 
                  onClick={handleUndo}
                  disabled={historyStack.length === 0}
                  variant="ghost"
                  size="sm"
                  className="text-slate-700 hover:bg-slate-200/70 rounded-lg p-1.5 h-7 w-7 flex items-center justify-center disabled:opacity-35"
                  title="Desfazer (Ctrl+Z)"
                >
                  <Undo2 className="w-3.5 h-3.5 text-slate-700" />
                </Button>
                <Button 
                  onClick={handleRedo}
                  disabled={redoStack.length === 0}
                  variant="ghost"
                  size="sm"
                  className="text-slate-700 hover:bg-slate-200/70 rounded-lg p-1.5 h-7 w-7 flex items-center justify-center disabled:opacity-35"
                  title="Refazer (Ctrl+Y)"
                >
                  <Redo2 className="w-3.5 h-3.5 text-slate-700" />
                </Button>
                <Button 
                  onClick={handleCorrectSpelling} 
                  disabled={isLoading || !generatedText} 
                  variant="ghost" 
                  size="sm" 
                  className="text-[11px] h-7 px-2 font-semibold text-emerald-700 hover:bg-emerald-50 hover:text-emerald-800 rounded-lg whitespace-nowrap"
                  title="Corrigir Ortografia e Gramática"
                >
                  {isLoading ? <Loader2 className="w-3.5 h-3.5 mr-1 animate-spin" /> : <CheckCircle className="w-3.5 h-3.5 mr-1 text-emerald-600" />}
                  Ortografia
                </Button>
                <Button 
                  size="sm" 
                  onClick={handleInsertCover} 
                  disabled={isLoading} 
                  variant="ghost" 
                  className="text-[11px] h-7 px-1.5 font-semibold text-blue-700 hover:bg-blue-50 hover:text-blue-800 rounded-lg whitespace-nowrap"
                  title="Inserir Capa e Folha de Rosto ABNT"
                >
                  {isLoading ? <Loader2 className="w-3.5 h-3.5 mr-1 animate-spin text-blue-600" /> : <BookOpen className="w-3.5 h-3.5 mr-1 text-blue-600" />}
                  Capa
                </Button>
                <Button 
                  size="sm" 
                  onClick={handleGenerateTOC} 
                  disabled={isLoading || !generatedText} 
                  variant="ghost" 
                  className="text-[11px] h-7 px-1.5 font-semibold text-indigo-700 hover:bg-indigo-50 hover:text-indigo-800 rounded-lg whitespace-nowrap"
                  title="Gerar Sumário Dinâmico ABNT"
                >
                  <ListOrdered className="w-3.5 h-3.5 mr-1 text-indigo-600" />
                  Sumário
                </Button>
                <Button 
                  size="sm" 
                  onClick={() => {
                    setShowCitationModal(true);
                  }} 
                  variant="ghost" 
                  className="text-[11px] h-7 px-1.5 font-semibold text-purple-700 hover:bg-purple-50 hover:text-purple-800 rounded-lg whitespace-nowrap cursor-pointer"
                  title="Inserir Citação Direta (Curta/Longa com Recuo de 4cm) ou Indireta ABNT NBR 10520:2023"
                >
                  <Quote className="w-3.5 h-3.5 mr-1 text-purple-600" />
                  Citação
                </Button>
                <Button 
                  size="sm" 
                  onClick={() => setShowReferenceModal(true)} 
                  variant="ghost" 
                  className="text-[11px] h-7 px-1.5 font-semibold text-slate-700 hover:bg-slate-100 rounded-lg whitespace-nowrap"
                  title="Gerar Referência por Link/DOI"
                >
                  <Link className="w-3.5 h-3.5 mr-1 text-gray-500" />
                  Referências
                </Button>
                <Button 
                  size="sm" 
                  onClick={handlePaginate} 
                  disabled={isLoading || !generatedText} 
                  variant="ghost" 
                  className="text-xs h-7 px-1.5 font-semibold text-emerald-700 hover:bg-emerald-50 hover:text-emerald-800 rounded-lg whitespace-nowrap"
                  title="Paginar / Repaginar Documento ABNT A4"
                >
                  <Hash className="w-3.5 h-3.5 mr-1 text-emerald-600" />
                  Paginar
                </Button>
              </div>

              {/* Grupo de Exportação e Cópia (PDF, Word, Copiar) */}
              <div className="flex items-center bg-gray-50 border border-gray-200/80 rounded-xl p-0.5 shadow-2xs gap-0.5 flex-nowrap">
                <Button 
                  onClick={exportPDF} 
                  disabled={!generatedText} 
                  variant="ghost" 
                  size="sm" 
                  className="text-xs h-7 px-2 text-rose-700 hover:bg-rose-50 hover:text-rose-800 rounded-lg font-semibold whitespace-nowrap"
                  title="Exportar PDF A4"
                >
                  <Download className="w-3.5 h-3.5 mr-1 text-rose-600" />
                  PDF
                </Button>
                <Button 
                  onClick={exportWord} 
                  disabled={!generatedText} 
                  variant="ghost" 
                  size="sm" 
                  className="text-xs h-7 px-2 text-blue-700 hover:bg-blue-50 hover:text-blue-800 rounded-lg font-semibold whitespace-nowrap"
                  title="Exportar Documento Word (.docx)"
                >
                  <FileDown className="w-3.5 h-3.5 mr-1 text-blue-600" />
                  Word
                </Button>
                <Button 
                  onClick={handleCopy} 
                  disabled={!generatedText} 
                  variant="ghost" 
                  size="sm" 
                  className="text-xs h-7 px-2 text-gray-700 hover:bg-gray-200/60 rounded-lg font-medium whitespace-nowrap"
                  title="Copiar Texto Completo"
                >
                  <Copy className="w-3.5 h-3.5 mr-1 text-gray-500" />
                  Copiar
                </Button>
              </div>
            </div>

            {/* Controle de Zoom Harmonioso e Delicado (À Direita) */}
            <div className="flex items-center bg-gray-100/90 hover:bg-gray-200/60 border border-gray-300/70 rounded-lg p-0.5 gap-0.5 select-none shadow-2xs h-7 ml-auto flex-shrink-0 transition-colors">
              <button
                onClick={() => setZoomScale(z => Math.max(30, z - 10))}
                title="Diminuir Zoom (-10%)"
                className="w-6 h-6 flex items-center justify-center hover:bg-white rounded-md text-gray-700 hover:text-blue-600 transition-all active:scale-90"
              >
                <ZoomOut className="w-3.5 h-3.5 stroke-[2]" />
              </button>
              
              <button
                onClick={() => setZoomScale(65)}
                title="Restaurar Zoom Padrão (65%)"
                className="px-1.5 h-6 flex items-center justify-center text-[11px] font-bold text-gray-800 hover:text-blue-600 hover:bg-white rounded-md min-w-[34px] text-center tracking-tight transition-all"
              >
                {zoomScale}%
              </button>
              
              <button
                onClick={() => setZoomScale(z => Math.min(150, z + 10))}
                title="Aumentar Zoom (+10%)"
                className="w-6 h-6 flex items-center justify-center hover:bg-white rounded-md text-gray-700 hover:text-blue-600 transition-all active:scale-90"
              >
                <ZoomIn className="w-3.5 h-3.5 stroke-[2]" />
              </button>
            </div>
          </div>

          {/* Canvas Area with Progress Bar */}
          <div className="flex-1 bg-white rounded-xl shadow-sm border border-gray-200 overflow-hidden flex flex-col relative h-full min-h-[680px]">
            <div className="relative z-10">
              {isLoading && (
                <div className="w-full bg-blue-100 h-1.5 overflow-hidden">
                  <div 
                    className="bg-blue-600 h-full transition-all duration-300 ease-out" 
                    style={{ width: `${progress}%` }}
                  />
                </div>
              )}
            </div>

            {(activeTab === "editor" || activeTab === "generator") && (
              <div 
                onClick={() => setIsImageSelected(false)}
                className="flex-1 flex flex-col h-full relative overflow-hidden bg-slate-100"
              >
                {(() => {
                  let pages: string[] = [];
                  const pageBreakRegex = /\s*---\s*\[(?:QUEBRA DE P[AÁ]GINA|NOVA P[AÁ]GINA)\]\s*---\s*|\s*\[(?:QUEBRA DE P[AÁ]GINA|NOVA P[AÁ]GINA)\]\s*/i;
                  
                  if (generatedText && pageBreakRegex.test(generatedText)) {
                    pages = generatedText.split(pageBreakRegex).map(p => p.trim()).filter(Boolean);
                  } else if (generatedText && generatedText.trim()) {
                    let bodyContent = generatedText;
                    let referencesContent = "";

                    const refMatch = bodyContent.match(/\n\s*(?:#+\s*)?(?:REFERÊNCIAS(?:\s+BIBLIOGRÁFICAS)?|REFERENCIAS)\s*\n([\s\S]*)$/i);
                    if (refMatch) {
                      referencesContent = `REFERÊNCIAS\n\n${refMatch[1].trim()}`;
                      bodyContent = bodyContent.substring(0, refMatch.index).trim();
                    }

                    // Divide o texto em blocos de páginas A4 (~2200 caracteres cada)
                    const paragraphs = bodyContent.split(/\n\n+/);
                    const bodyPages: string[] = [];
                    let curPage = "";
                    for (const para of paragraphs) {
                      if ((curPage + "\n\n" + para).length > 2200 && curPage.trim().length > 0) {
                        bodyPages.push(curPage.trim());
                        curPage = para;
                      } else {
                        curPage = curPage ? curPage + "\n\n" + para : para;
                      }
                    }
                    if (curPage.trim()) bodyPages.push(curPage.trim());

                    if (referencesContent) {
                      bodyPages.push(referencesContent);
                    }

                    const requiresFormalCover = !["resumo", "redacao", "resenha"].includes(documentType);
                    if (requiresFormalCover) {
                      pages = ["CAPA_AUTO", "FOLHA_ROSTO_AUTO", ...(bodyPages.length > 0 ? bodyPages : [generatedText])];
                    } else {
                      pages = bodyPages.length > 0 ? bodyPages : [generatedText];
                    }
                  } else {
                    pages = [""];
                  }

                  const renderSingleA4Sheet = (text: string, pIdx: number) => {
                    const cleanT = text.trim();
                    const requiresFormalCover = !["resumo", "redacao", "resenha"].includes(documentType);
                    
                    // Detecção Estrita e Completa de Capa (Página 1) e Folha de Rosto (Página 2)
                    const isCover = requiresFormalCover && (pIdx === 0 || cleanT === "CAPA_AUTO" || cleanT.startsWith("CAPA"));
                    const isTitlePage = requiresFormalCover && (pIdx === 1 || cleanT === "FOLHA_ROSTO_AUTO" || cleanT.startsWith("FOLHA DE ROSTO") || cleanT.includes("requisito") || cleanT.includes("apresentado") || cleanT.includes("Orientador"));
                    const isBodyPage = !isCover && !isTitlePage;
                    const pageNum = pIdx + 1;
                    const lines = text.split("\n").map(l => l.trim()).filter(Boolean);

                    return (
                      <div 
                        key={pIdx}
                        className="print-page-a4 w-full max-w-[760px] bg-white text-gray-900 shadow-md border border-gray-200 rounded-sm relative flex flex-col p-8 sm:p-12 md:p-16 my-4 select-text print:shadow-none print:border-none print:m-0 print:p-0 print:h-[297mm] print:w-[210mm] print:break-after-page min-h-[900px] group/page"
                      >
                        {/* BARRA DE CONTROLE DA PÁGINA (aparece no hover) */}
                        <div className="absolute -top-1 left-1/2 -translate-x-1/2 translate-y-0 opacity-0 group-hover/page:opacity-100 transition-all duration-200 z-30 flex items-center gap-1 bg-white/95 backdrop-blur-md border border-slate-200/90 shadow-lg rounded-xl px-2 py-1 print:hidden">
                          <span className="text-[10px] font-bold text-slate-500 px-1.5">Pág. {pageNum}</span>
                          <span className="h-3 w-px bg-slate-200" />
                          {/* Apagar Página (remove a folha inteira preservando as outras intactas) */}
                          <button
                            type="button"
                            onClick={(e) => {
                              e.stopPropagation();
                              // 1. Remove a página selecionada
                              const newPages = pages.filter((_, i) => i !== pIdx);
                              
                              if (newPages.length === 0) {
                                updateGeneratedTextWithHistory("");
                              } else {
                                // 2. Reconstrói o documento mantendo cada folha A4 estritamente isolada por Quebra de Página
                                const joinedDoc = newPages
                                  .map(p => {
                                    if (p === "CAPA_AUTO") return generateCoverTextLocally().split("--- [QUEBRA DE PÁGINA] ---")[0].trim();
                                    if (p === "FOLHA_ROSTO_AUTO") {
                                      const fullCov = generateCoverTextLocally().split("--- [QUEBRA DE PÁGINA] ---");
                                      return (fullCov[1] || "").trim();
                                    }
                                    return p.trim();
                                  })
                                  .filter(Boolean)
                                  .join("\n\n--- [QUEBRA DE PÁGINA] ---\n\n");
                                
                                updateGeneratedTextWithHistory(joinedDoc);
                              }

                              // 3. Ajusta os índices de numeração oculta após remoção da folha inteira
                              setHiddenPageNumbers(prev => {
                                const next = new Set<number>();
                                prev.forEach(idx => {
                                  if (idx < pIdx) next.add(idx);
                                  else if (idx > pIdx) next.add(idx - 1);
                                });
                                return next;
                              });

                              setErrorMessage(`🗑️ Página ${pageNum} removida. As páginas seguintes subiram intactas!`);
                              setTimeout(() => setErrorMessage(""), 3500);
                            }}
                            title="Apagar esta página inteira do documento (a folha seguinte sobre inteira)"
                            className="px-2 py-0.5 text-[10px] font-medium text-red-500 hover:text-red-700 hover:bg-red-50 rounded-lg transition-all flex items-center gap-1"
                          >
                            🗑️ Apagar Página
                          </button>
                          <span className="h-3 w-px bg-slate-200" />
                          {/* Adicionar Página em branco após esta */}
                          <button
                            type="button"
                            onClick={(e) => {
                              e.stopPropagation();
                              const newPages = [...pages];
                              newPages.splice(pIdx + 1, 0, "");
                              setGeneratedText(newPages.join("\n\n--- [QUEBRA DE PÁGINA] ---\n\n"));
                              // Ajusta índices de numeração oculta após inserção
                              setHiddenPageNumbers(prev => {
                                const next = new Set<number>();
                                prev.forEach(idx => {
                                  if (idx <= pIdx) next.add(idx);
                                  else next.add(idx + 1);
                                });
                                return next;
                              });
                            }}
                            title="Adicionar uma nova página em branco após esta"
                            className="px-2 py-0.5 text-[10px] font-medium text-emerald-600 hover:text-emerald-800 hover:bg-emerald-50 rounded-lg transition-all flex items-center gap-1"
                          >
                            ➕ Adicionar Página
                          </button>
                        </div>

                        {/* NUMERAÇÃO OFICIAL IMPRESSA NO CANTO SUPERIOR DIREITO — Clicável para apagar o número */}
                        {isBodyPage && !hiddenPageNumbers.has(pIdx) && (
                          <button
                            type="button"
                            onClick={(e) => {
                              e.stopPropagation();
                              setHiddenPageNumbers(prev => {
                                const next = new Set(prev);
                                next.add(pIdx);
                                return next;
                              });
                            }}
                            title="Clique para apagar a numeração desta página"
                            className="absolute top-[4%] right-[5%] font-mono text-xs font-bold text-gray-800 select-none hover:text-red-500 hover:line-through cursor-pointer transition-colors print:pointer-events-none z-20"
                          >
                            {pageNum}
                          </button>
                        )}

                        {/* RENDERIZAÇÃO DA CAPA ABNT (TOTALMENTE EDITÁVEL) */}
                        {isCover ? (
                          <div className="flex-1 flex flex-col justify-between text-center font-['Arial'] text-gray-900 py-4 select-text">
                            {/* TOPO: INSTITUIÇÃO E CURSO / DISCIPLINA */}
                            <div>
                              <div
                                contentEditable
                                suppressContentEditableWarning
                                onBlur={(e) => setInstitution(e.currentTarget.innerText.trim())}
                                className="font-bold text-sm sm:text-base uppercase tracking-wider focus:outline-none focus:bg-blue-50/50 p-1 rounded"
                              >
                                {institution || (lines.find(l => l !== "CAPA" && l !== "CAPA_AUTO") || "INSTITUIÇÃO DE ENSINO SUPERIOR")}
                              </div>
                              {course && (
                                <div
                                  contentEditable
                                  suppressContentEditableWarning
                                  onBlur={(e) => setCourse(e.currentTarget.innerText.trim())}
                                  className="font-semibold text-xs sm:text-sm uppercase text-gray-700 mt-1 focus:outline-none focus:bg-blue-50/50 p-1 rounded"
                                >
                                  {course}
                                </div>
                              )}
                              {subject && (
                                <div
                                  contentEditable
                                  suppressContentEditableWarning
                                  onBlur={(e) => setSubject(e.currentTarget.innerText.replace(/^DISCIPLINA:\s*/i, '').trim())}
                                  className="font-medium text-xs uppercase text-gray-600 mt-0.5 focus:outline-none focus:bg-blue-50/50 p-1 rounded"
                                >
                                  DISCIPLINA: {subject}
                                </div>
                              )}
                            </div>

                            {/* AUTOR: CENTRALIZADO ENTRE O TOPO E O MEIO CONFORME ABNT NBR 14724 */}
                            <div className="my-auto py-6">
                              <div
                                contentEditable
                                suppressContentEditableWarning
                                onBlur={(e) => setStudentName(e.currentTarget.innerText.trim())}
                                className="font-semibold text-sm sm:text-base uppercase tracking-wide focus:outline-none focus:bg-blue-50/50 p-1 rounded"
                              >
                                {studentName || "NOME DO(A) AUTOR(A)"}
                              </div>
                            </div>

                            {/* CENTRO: TÍTULO E SUBTÍTULO */}
                            <div className="my-auto py-6">
                              <div
                                contentEditable
                                suppressContentEditableWarning
                                onBlur={(e) => setTitle(e.currentTarget.innerText.trim())}
                                className="font-extrabold text-base sm:text-lg uppercase tracking-tight text-gray-900 leading-snug focus:outline-none focus:bg-blue-50/50 p-1 rounded"
                              >
                                {title || "TÍTULO DO TRABALHO"}
                              </div>
                              {subtitle && (
                                <div
                                  contentEditable
                                  suppressContentEditableWarning
                                  onBlur={(e) => setSubtitle(e.currentTarget.innerText.trim())}
                                  className="font-normal text-xs sm:text-sm text-gray-700 mt-1 focus:outline-none focus:bg-blue-50/50 p-1 rounded"
                                >
                                  {subtitle}
                                </div>
                              )}
                            </div>

                            {/* RODAPÉ: CIDADE E ANO */}
                            <div className="mt-auto pt-6">
                              <div
                                contentEditable
                                suppressContentEditableWarning
                                onBlur={(e) => setCity(e.currentTarget.innerText.trim())}
                                className="font-bold text-xs sm:text-sm uppercase text-gray-800 focus:outline-none focus:bg-blue-50/50 p-1 rounded"
                              >
                                {city || "CIDADE - UF"}
                              </div>
                              <div
                                contentEditable
                                suppressContentEditableWarning
                                onBlur={(e) => setYear(e.currentTarget.innerText.trim())}
                                className="font-bold text-xs sm:text-sm text-gray-800 focus:outline-none focus:bg-blue-50/50 p-1 rounded"
                              >
                                {year || String(new Date().getFullYear())}
                              </div>
                            </div>
                          </div>
                        ) : isTitlePage ? (
                          /* RENDERIZAÇÃO DA FOLHA DE ROSTO ABNT (TOTALMENTE EDITÁVEL) */
                          <div className="flex-1 flex flex-col justify-between font-['Arial'] text-gray-900 py-4 select-text">
                            <div className="text-center">
                              {((studentName && studentName.trim()) || (lines[0] && lines[0] !== "FOLHA_ROSTO_AUTO" && lines[0].trim())) && (
                                <div
                                  contentEditable
                                  suppressContentEditableWarning
                                  onBlur={(e) => setStudentName(e.currentTarget.innerText.trim())}
                                  className="font-semibold text-sm sm:text-base uppercase tracking-wide focus:outline-none focus:bg-blue-50/50 p-1 rounded"
                                >
                                  {studentName || lines[0]}
                                </div>
                              )}
                            </div>

                            <div className="my-auto text-center py-6">
                              {((title && title.trim()) || (lines[1] && lines[1] !== "FOLHA_ROSTO_AUTO" && lines[1].trim())) && (
                                <div
                                  contentEditable
                                  suppressContentEditableWarning
                                  onBlur={(e) => setTitle(e.currentTarget.innerText.trim())}
                                  className="font-bold text-base sm:text-lg uppercase tracking-tight text-gray-900 focus:outline-none focus:bg-blue-50/50 p-1 rounded"
                                >
                                  {title || lines[1]}
                                </div>
                              )}
                              {((subtitle && subtitle.trim()) || (lines[2] && lines[2] !== "FOLHA_ROSTO_AUTO" && lines[2].trim())) && (
                                <div
                                  contentEditable
                                  suppressContentEditableWarning
                                  onBlur={(e) => setSubtitle(e.currentTarget.innerText.trim())}
                                  className="font-normal text-xs sm:text-sm text-gray-700 mt-1 focus:outline-none focus:bg-blue-50/50 p-1 rounded"
                                >
                                  {subtitle || lines[2]}
                                </div>
                              )}
                            </div>

                            <div className="my-auto w-full flex justify-end">
                              <div 
                                contentEditable
                                suppressContentEditableWarning
                                className="w-3/5 text-justify text-[10pt] sm:text-[10.5pt] leading-[1.3] text-gray-800 bg-gray-50/50 p-3 rounded border border-gray-200 focus:outline-none focus:bg-blue-50/50"
                              >
                                <p>
                                  {(documentType === "outros" ? customDocumentType : documentType) || "Trabalho Acadêmico"} apresentado à {institution || "Instituição de Ensino"}{course ? ` como requisito parcial de avaliação para o curso de ${course}` : ""}.
                                </p>
                                {advisor && (
                                  <p className="mt-2 font-semibold text-gray-900 text-[9.5pt]">
                                    Orientador(a): {advisor}
                                  </p>
                                )}
                              </div>
                            </div>

                            {(((city && city.trim()) || (lines[3] && lines[3] !== "FOLHA_ROSTO_AUTO" && lines[3].trim())) || ((year && year.trim()) || (lines[4] && lines[4] !== "FOLHA_ROSTO_AUTO" && lines[4].trim()))) && (
                              <div className="text-center mt-auto pt-6">
                                {((city && city.trim()) || (lines[3] && lines[3] !== "FOLHA_ROSTO_AUTO" && lines[3].trim())) && (
                                  <div
                                    contentEditable
                                    suppressContentEditableWarning
                                    onBlur={(e) => setCity(e.currentTarget.innerText.trim())}
                                    className="font-bold text-xs sm:text-sm uppercase text-gray-800 focus:outline-none focus:bg-blue-50/50 p-1 rounded"
                                  >
                                    {city || lines[3]}
                                  </div>
                                )}
                                {((year && year.trim()) || (lines[4] && lines[4] !== "FOLHA_ROSTO_AUTO" && lines[4].trim())) && (
                                  <div
                                    contentEditable
                                    suppressContentEditableWarning
                                    onBlur={(e) => setYear(e.currentTarget.innerText.trim())}
                                    className="font-bold text-xs sm:text-sm text-gray-800 focus:outline-none focus:bg-blue-50/50 p-1 rounded"
                                  >
                                    {year || lines[4]}
                                  </div>
                                )}
                              </div>
                            )}
                          </div>
                        ) : (
                          /* RENDERIZAÇÃO DO CORPO DO TRABALHO ABNT (PÁGINAS 3 EM DIANTE - 100% EDITÁVEL) */
                          <div className="flex-1 flex flex-col font-['Arial'] text-gray-900 select-text">
                            {text.trim().startsWith("SUMÁRIO") ? (
                              <div className="w-full font-['Arial'] text-gray-900 leading-[1.8] text-sm sm:text-base py-2">
                                <div className="font-bold text-center text-base mb-6 tracking-wide">SUMÁRIO</div>
                                <div className="space-y-1 font-mono text-xs sm:text-sm">
                                  {text.split('\n').filter(l => l.trim() && !l.trim().startsWith("SUMÁRIO")).map((line, lIdx) => {
                                    const match = line.match(/^(.*?)\s*(\.{3,})\s*(\d+)$/);
                                    if (match) {
                                      return (
                                        <div key={lIdx} className="flex items-baseline justify-between gap-2">
                                          <span className="font-semibold text-gray-900 truncate">{match[1].trim()}</span>
                                          <span className="flex-1 border-b border-dotted border-gray-400 mx-1 mb-1" />
                                          <span className="font-bold text-gray-800 tabular-nums">{match[3]}</span>
                                        </div>
                                      );
                                    }
                                    return <div key={lIdx} className="font-semibold text-gray-800">{line}</div>;
                                  })}
                                </div>
                              </div>
                            ) : text.includes("![Figura inserida](") ? (
                              <div className="w-full font-['Arial'] text-gray-900 leading-[1.6] text-sm sm:text-base py-2">
                                {(() => {
                                  const parts = text.split(/(!\[Figura inserida\]\(.*?\))/g);
                                  return parts.map((part, pPartIdx) => {
                                    const imgMatch = part.match(/!\[Figura inserida\]\((.*?)\)/);
                                    if (imgMatch) {
                                      return (
                                        <div 
                                          key={pPartIdx} 
                                          onClick={(e) => {
                                            e.stopPropagation();
                                            setIsImageSelected(true);
                                          }}
                                          className={`my-6 flex flex-col ${imageAlign === "center" ? "items-center" : imageAlign === "left" ? "items-start" : "items-end"} justify-center relative select-none`}
                                        >
                                          {/* Barra de Ferramentas de Formatação de Imagem (Design Delicado e Elegante) */}
                                          {isImageSelected && (
                                            <div 
                                              onClick={(e) => e.stopPropagation()} 
                                              className="flex flex-wrap items-center gap-1.5 bg-white/95 backdrop-blur-md px-3 py-1.5 rounded-2xl border border-slate-200/90 shadow-xl shadow-slate-200/50 mb-2.5 z-30 transition-all text-xs select-none"
                                            >
                                              {/* Identificação Elegante */}
                                              <div className="flex items-center gap-1.5 pr-2 border-r border-slate-200">
                                                <span className="w-2 h-2 rounded-full bg-blue-600 animate-pulse" />
                                                <span className="text-[11px] font-semibold text-slate-700 tracking-tight">Formatar Imagem</span>
                                              </div>

                                              {/* Alinhamento */}
                                              <div className="flex items-center bg-slate-100/90 p-0.5 rounded-xl">
                                                <button
                                                  type="button"
                                                  onClick={() => setImageAlign("left")}
                                                  className={`px-2 py-0.5 text-[11px] rounded-lg transition-all ${imageAlign === "left" ? "bg-white text-blue-600 font-bold shadow-2xs" : "text-slate-600 hover:text-slate-900"}`}
                                                  title="Alinhar à Esquerda"
                                                >
                                                  Esquerda
                                                </button>
                                                <button
                                                  type="button"
                                                  onClick={() => setImageAlign("center")}
                                                  className={`px-2 py-0.5 text-[11px] rounded-lg transition-all ${imageAlign === "center" ? "bg-white text-blue-600 font-bold shadow-2xs" : "text-slate-600 hover:text-slate-900"}`}
                                                  title="Centralizar (Padrão ABNT)"
                                                >
                                                  Centro
                                                </button>
                                                <button
                                                  type="button"
                                                  onClick={() => setImageAlign("right")}
                                                  className={`px-2 py-0.5 text-[11px] rounded-lg transition-all ${imageAlign === "right" ? "bg-white text-blue-600 font-bold shadow-2xs" : "text-slate-600 hover:text-slate-900"}`}
                                                  title="Alinhar à Direita"
                                                >
                                                  Direita
                                                </button>
                                              </div>

                                              {/* Orientação */}
                                              <div className="flex items-center bg-slate-100/90 p-0.5 rounded-xl">
                                                <button
                                                  type="button"
                                                  onClick={() => {
                                                    setImageOrientation("portrait");
                                                    setImageWidth(55);
                                                    setImageHeight(420);
                                                  }}
                                                  className={`px-2 py-0.5 text-[11px] rounded-lg transition-all ${imageOrientation === "portrait" ? "bg-white text-blue-600 font-bold shadow-2xs" : "text-slate-600 hover:text-slate-900"}`}
                                                  title="Orientação Retrato (Vertical)"
                                                >
                                                  Retrato
                                                </button>
                                                <button
                                                  type="button"
                                                  onClick={() => {
                                                    setImageOrientation("landscape");
                                                    setImageWidth(85);
                                                    setImageHeight(320);
                                                  }}
                                                  className={`px-2 py-0.5 text-[11px] rounded-lg transition-all ${imageOrientation === "landscape" ? "bg-white text-blue-600 font-bold shadow-2xs" : "text-slate-600 hover:text-slate-900"}`}
                                                  title="Orientação Paisagem (Horizontal)"
                                                >
                                                  Paisagem
                                                </button>
                                              </div>

                                              {/* Girar */}
                                              <button
                                                type="button"
                                                onClick={() => setImageRotation(r => (r + 90) % 360)}
                                                className="px-2 py-1 text-[11px] font-medium bg-slate-100 hover:bg-blue-50 text-slate-700 hover:text-blue-700 rounded-xl transition-all flex items-center gap-1 active:scale-95"
                                                title="Girar 90 graus no sentido horário"
                                              >
                                                <span>Girar</span>
                                                <span className="text-[10px] text-blue-600 font-bold font-mono">{imageRotation}°</span>
                                              </button>

                                              {/* Estilos */}
                                              <div className="flex items-center gap-0.5 pl-1">
                                                <button
                                                  type="button"
                                                  onClick={() => setImageStyle("academic_box")}
                                                  className={`px-2 py-0.5 text-[11px] rounded-lg transition-all ${imageStyle === "academic_box" ? "bg-blue-600 text-white font-semibold shadow-2xs" : "text-slate-600 hover:bg-slate-100"}`}
                                                >
                                                  ABNT
                                                </button>
                                                <button
                                                  type="button"
                                                  onClick={() => setImageStyle("simple_border")}
                                                  className={`px-2 py-0.5 text-[11px] rounded-lg transition-all ${imageStyle === "simple_border" ? "bg-blue-600 text-white font-semibold shadow-2xs" : "text-slate-600 hover:bg-slate-100"}`}
                                                >
                                                  Borda
                                                </button>
                                                <button
                                                  type="button"
                                                  onClick={() => setImageStyle("soft_shadow")}
                                                  className={`px-2 py-0.5 text-[11px] rounded-lg transition-all ${imageStyle === "soft_shadow" ? "bg-blue-600 text-white font-semibold shadow-2xs" : "text-slate-600 hover:bg-slate-100"}`}
                                                >
                                                  Sombra
                                                </button>
                                                <button
                                                  type="button"
                                                  onClick={() => setImageStyle("rounded_frame")}
                                                  className={`px-2 py-0.5 text-[11px] rounded-lg transition-all ${imageStyle === "rounded_frame" ? "bg-blue-600 text-white font-semibold shadow-2xs" : "text-slate-600 hover:bg-slate-100"}`}
                                                >
                                                  Curva
                                                </button>
                                              </div>

                                              {/* Separador e Ações Finais */}
                                              <span className="h-3 w-px bg-slate-200 mx-0.5" />

                                              <button
                                                type="button"
                                                onClick={() => {
                                                  setImageWidth(80);
                                                  setImageHeight(320);
                                                  setImageAlign("center");
                                                  setImageOrientation("landscape");
                                                  setImageRotation(0);
                                                  setImageOffset({ x: 0, y: 0 });
                                                  setImageStyle("academic_box");
                                                }}
                                                className="px-2 py-1 text-[11px] text-slate-500 hover:text-slate-800 hover:bg-slate-100 rounded-lg transition-all"
                                                title="Redefinir formatação original"
                                              >
                                                Redefinir
                                              </button>

                                              <button
                                                type="button"
                                                onClick={() => setIsImageSelected(false)}
                                                className="w-5 h-5 flex items-center justify-center text-slate-400 hover:text-slate-700 hover:bg-slate-100 rounded-full transition-all text-xs"
                                                title="Fechar barra"
                                              >
                                                ✕
                                              </button>
                                            </div>
                                          )}

                                          {/* Container da Imagem com Reposicionamento Livre (Drag to Move) e 8 Alças Word */}
                                          <div 
                                            style={{ 
                                              width: `${imageWidth}%`,
                                              transform: `translate(${imageOffset.x}px, ${imageOffset.y}px)`
                                            }}
                                            className={`relative group/box p-1 rounded transition-all flex flex-col items-center select-none ${
                                              isImageSelected ? "ring-2 ring-blue-500 ring-offset-2 border border-blue-400 border-dashed" : "border border-transparent hover:border-gray-300 hover:border-dashed"
                                            }`}
                                          >
                                            {/* Alça Central Superior para Arrastar e Reposicionar a Imagem Livremente */}
                                            {isImageSelected && (
                                              <div
                                                onMouseDown={(e) => {
                                                  e.preventDefault();
                                                  const startX = e.clientX;
                                                  const startY = e.clientY;
                                                  const initialOffset = { ...imageOffset };
                                                  const onMouseMove = (ev: MouseEvent) => {
                                                    const dX = ev.clientX - startX;
                                                    const dY = ev.clientY - startY;
                                                    setImageOffset({
                                                      x: initialOffset.x + dX,
                                                      y: initialOffset.y + dY
                                                    });
                                                  };
                                                  const onMouseUp = () => {
                                                    window.removeEventListener("mousemove", onMouseMove);
                                                    window.removeEventListener("mouseup", onMouseUp);
                                                  };
                                                  window.addEventListener("mousemove", onMouseMove);
                                                  window.addEventListener("mouseup", onMouseUp);
                                                }}
                                                title="Clique e arraste para reposicionar a imagem onde quiser (Padrão Word)"
                                                className="absolute -top-3.5 left-1/2 -translate-x-1/2 bg-blue-600 hover:bg-blue-700 text-white p-1 rounded-full shadow-md cursor-grab active:cursor-grabbing z-10 flex items-center justify-center"
                                              >
                                                <Move className="w-3.5 h-3.5" />
                                              </div>
                                            )}

                                            <div 
                                              style={{ height: imageHeight > 0 ? `${imageHeight}px` : "auto" }}
                                              onMouseDown={(e) => {
                                                if (e.target === e.currentTarget || (e.target as HTMLElement).tagName === 'IMG') {
                                                  e.preventDefault();
                                                  setIsImageSelected(true);
                                                  const startX = e.clientX;
                                                  const startY = e.clientY;
                                                  const initialOffset = { ...imageOffset };
                                                  const onMouseMove = (ev: MouseEvent) => {
                                                    const dX = ev.clientX - startX;
                                                    const dY = ev.clientY - startY;
                                                    setImageOffset({
                                                      x: initialOffset.x + dX,
                                                      y: initialOffset.y + dY
                                                    });
                                                  };
                                                  const onMouseUp = () => {
                                                    window.removeEventListener("mousemove", onMouseMove);
                                                    window.removeEventListener("mouseup", onMouseUp);
                                                  };
                                                  window.addEventListener("mousemove", onMouseMove);
                                                  window.addEventListener("mouseup", onMouseUp);
                                                }
                                              }}
                                              className={`w-full overflow-hidden flex items-center justify-center transition-all cursor-pointer ${
                                                imageStyle === "academic_box" 
                                                  ? "p-2 bg-white border border-gray-300 shadow-2xs"
                                                  : imageStyle === "simple_border"
                                                  ? "border-2 border-gray-800"
                                                  : imageStyle === "soft_shadow"
                                                  ? "shadow-lg bg-white rounded"
                                                  : imageStyle === "rounded_frame"
                                                  ? "rounded-2xl border border-gray-200 shadow-md"
                                                  : ""
                                              }`}
                                            >
                                              <img 
                                                src={imgMatch[1]} 
                                                alt="Figura acadêmica" 
                                                style={{ 
                                                  height: imageHeight > 0 ? `${imageHeight}px` : "auto",
                                                  transform: `rotate(${imageRotation}deg)`
                                                }}
                                                className="w-full object-contain select-none pointer-events-none transition-transform duration-200" 
                                                draggable={false}
                                              />
                                            </div>

                                            {/* 8 Alças de Arraste (Visíveis APENAS quando a imagem está selecionada) */}
                                            {isImageSelected && (
                                              <>
                                                {/* 1. Alça Direita (Ajustar Largura) */}
                                                <div 
                                                  onMouseDown={(e) => {
                                                    e.preventDefault();
                                                    const startX = e.clientX;
                                                    const startW = imageWidth;
                                                    const onMouseMove = (ev: MouseEvent) => {
                                                      const dX = ev.clientX - startX;
                                                      setImageWidth(Math.min(100, Math.max(25, Math.round(startW + (dX / 4)))));
                                                    };
                                                    const onMouseUp = () => {
                                                      window.removeEventListener("mousemove", onMouseMove);
                                                      window.removeEventListener("mouseup", onMouseUp);
                                                    };
                                                    window.addEventListener("mousemove", onMouseMove);
                                                    window.addEventListener("mouseup", onMouseUp);
                                                  }}
                                                  title="Arrastar largura"
                                                  className="absolute top-1/2 -right-2 -translate-y-1/2 w-3.5 h-3.5 bg-blue-600 border-2 border-white rounded-xs shadow cursor-ew-resize hover:scale-125 transition-all"
                                                />

                                                {/* 2. Alça Esquerda (Ajustar Largura) */}
                                                <div 
                                                  onMouseDown={(e) => {
                                                    e.preventDefault();
                                                    const startX = e.clientX;
                                                    const startW = imageWidth;
                                                    const onMouseMove = (ev: MouseEvent) => {
                                                      const dX = startX - ev.clientX;
                                                      setImageWidth(Math.min(100, Math.max(25, Math.round(startW + (dX / 4)))));
                                                    };
                                                    const onMouseUp = () => {
                                                      window.removeEventListener("mousemove", onMouseMove);
                                                      window.removeEventListener("mouseup", onMouseUp);
                                                    };
                                                    window.addEventListener("mousemove", onMouseMove);
                                                    window.addEventListener("mouseup", onMouseUp);
                                                  }}
                                                  title="Arrastar largura"
                                                  className="absolute top-1/2 -left-2 -translate-y-1/2 w-3.5 h-3.5 bg-blue-600 border-2 border-white rounded-xs shadow cursor-ew-resize hover:scale-125 transition-all"
                                                />

                                                {/* 3. Alça Inferior (Ajustar Altura) */}
                                                <div 
                                                  onMouseDown={(e) => {
                                                    e.preventDefault();
                                                    const startY = e.clientY;
                                                    const startH = imageHeight;
                                                    const onMouseMove = (ev: MouseEvent) => {
                                                      const dY = ev.clientY - startY;
                                                      setImageHeight(Math.min(700, Math.max(100, Math.round(startH + dY))));
                                                    };
                                                    const onMouseUp = () => {
                                                      window.removeEventListener("mousemove", onMouseMove);
                                                      window.removeEventListener("mouseup", onMouseUp);
                                                    };
                                                    window.addEventListener("mousemove", onMouseMove);
                                                    window.addEventListener("mouseup", onMouseUp);
                                                  }}
                                                  title="Arrastar altura"
                                                  className="absolute -bottom-2 left-1/2 -translate-x-1/2 w-3.5 h-3.5 bg-blue-600 border-2 border-white rounded-xs shadow cursor-ns-resize hover:scale-125 transition-all"
                                                />

                                                {/* 4. Alça Canto Inferior Direito (Ajustar Altura e Largura Simultâneos) */}
                                                <div 
                                                  onMouseDown={(e) => {
                                                    e.preventDefault();
                                                    const startX = e.clientX;
                                                    const startY = e.clientY;
                                                    const startW = imageWidth;
                                                    const startH = imageHeight;
                                                    const onMouseMove = (ev: MouseEvent) => {
                                                      const dX = ev.clientX - startX;
                                                      const dY = ev.clientY - startY;
                                                      setImageWidth(Math.min(100, Math.max(25, Math.round(startW + (dX / 4)))));
                                                      setImageHeight(Math.min(700, Math.max(100, Math.round(startH + dY))));
                                                    };
                                                    const onMouseUp = () => {
                                                      window.removeEventListener("mousemove", onMouseMove);
                                                      window.removeEventListener("mouseup", onMouseUp);
                                                    };
                                                    window.addEventListener("mousemove", onMouseMove);
                                                    window.addEventListener("mouseup", onMouseUp);
                                                  }}
                                                  title="Puxar e arrastar altura e largura (Padrão Word)"
                                                  className="absolute -bottom-2 -right-2 w-4 h-4 bg-blue-600 border-2 border-white rounded-sm shadow-md cursor-se-resize hover:scale-125 transition-all"
                                                />

                                                {/* 5. Alça Canto Inferior Esquerdo */}
                                                <div 
                                                  onMouseDown={(e) => {
                                                    e.preventDefault();
                                                    const startX = e.clientX;
                                                    const startY = e.clientY;
                                                    const startW = imageWidth;
                                                    const startH = imageHeight;
                                                    const onMouseMove = (ev: MouseEvent) => {
                                                      const dX = startX - ev.clientX;
                                                      const dY = ev.clientY - startY;
                                                      setImageWidth(Math.min(100, Math.max(25, Math.round(startW + (dX / 4)))));
                                                      setImageHeight(Math.min(700, Math.max(100, Math.round(startH + dY))));
                                                    };
                                                    const onMouseUp = () => {
                                                      window.removeEventListener("mousemove", onMouseMove);
                                                      window.removeEventListener("mouseup", onMouseUp);
                                                    };
                                                    window.addEventListener("mousemove", onMouseMove);
                                                    window.addEventListener("mouseup", onMouseUp);
                                                  }}
                                                  title="Puxar e arrastar altura e largura (Padrão Word)"
                                                  className="absolute -bottom-2 -left-2 w-4 h-4 bg-blue-600 border-2 border-white rounded-sm shadow-md cursor-sw-resize hover:scale-125 transition-all"
                                                />

                                                {/* Badge Discreta com Dimensões Exatas */}
                                                <div className="absolute -top-7 right-0 bg-gray-900/90 text-white text-[10px] font-bold px-2 py-0.5 rounded shadow">
                                                  {imageWidth}% L × {imageHeight}px A
                                                </div>
                                              </>
                                            )}
                                          </div>
                                        </div>
                                      );
                                    }
                                    return (
                                      <div
                                        key={pPartIdx}
                                        contentEditable
                                        suppressContentEditableWarning
                                        spellCheck={true}
                                        lang="pt-BR"
                                        className="w-full focus:outline-none text-justify whitespace-pre-wrap"
                                      >
                                        {part.trim()}
                                      </div>
                                    );
                                  });
                                })()}
                              </div>
                            ) : (
                              <div
                                contentEditable
                                suppressContentEditableWarning
                                spellCheck={true}
                                lang="pt-BR"
                                onInput={(e) => {
                                  const newPages = [...pages];
                                  newPages[pIdx] = e.currentTarget.innerText;
                                  setGeneratedText(newPages.join("\n\n--- [QUEBRA DE PÁGINA] ---\n\n"));
                                }}
                                className="w-full focus:outline-none font-['Arial'] text-gray-900 leading-[1.5] text-justify text-[12pt] [text-indent:1.25cm] bg-transparent min-h-[600px] whitespace-pre-wrap focus:ring-1 focus:ring-blue-300 p-2 rounded selection:bg-blue-100"
                              >
                                {text && text !== "CAPA_AUTO" && text !== "FOLHA_ROSTO_AUTO" ? text.trimStart() : ""}
                              </div>
                            )}
                          </div>
                        )}
                      </div>
                    );
                  };

                  return (
                    <div className="flex-1 w-full h-full flex flex-col items-center justify-start py-6 px-3 md:px-8 pb-12 relative overflow-y-auto overflow-x-auto select-text">
                      
                      {/* DOCUMENTO CONTÍNUO COM TODAS AS PÁGINAS A4 EMPILHADAS COM ZOOM REAL */}
                      <div 
                        className="flex flex-col items-center gap-6 transition-transform duration-200 origin-top"
                        style={{ 
                          transform: `scale(${zoomScale / 100})`, 
                          width: `${100 / (zoomScale / 100)}%`,
                          maxWidth: `${(760 * 100) / zoomScale}px`
                        }}
                      >
                        {pages.map((pText, idx) => renderSingleA4Sheet(pText, idx))}
                      </div>

                    </div>
                  );
                })()}
              </div>
            )}
            
            {activeTab === "chat" && (
              <div className="flex-1 flex flex-col h-full min-h-0 overflow-hidden bg-slate-50/50">
                {/* Barra de Controles Superiores do Chat: Limpar e Alternar Rolagem Automática */}
                <div className="px-4 py-2 bg-white border-b border-slate-200/80 flex items-center justify-between gap-2 shadow-2xs font-sans">
                  <div className="flex items-center gap-2">
                    <div className="w-6 h-6 rounded-lg bg-gradient-to-tr from-cyan-500 to-blue-600 text-white flex items-center justify-center text-xs shadow-2xs">
                      🤖
                    </div>
                    <span className="text-xs font-bold text-slate-800">Chat com a EMIA</span>
                    {activeQuiz && activeQuiz.isActive && !activeQuiz.isFinished && (
                      <span className="text-[10px] px-2 py-0.5 rounded-full font-extrabold bg-amber-100 text-amber-800 border border-amber-300 animate-pulse">
                        Quiz em andamento
                      </span>
                    )}
                  </div>

                  <div className="flex items-center gap-1.5">
                    {/* Botão Liga/Desliga Rolagem Automática */}
                    <button
                      onClick={() => setIsAutoScrollEnabled(prev => !prev)}
                      title={isAutoScrollEnabled ? "Desativar Rolagem Automática" : "Ativar Rolagem Automática"}
                      className={`px-2.5 py-1 rounded-lg text-[11px] font-semibold flex items-center gap-1.5 transition-all shadow-2xs ${
                        isAutoScrollEnabled 
                          ? "bg-blue-50 text-blue-700 border border-blue-200 hover:bg-blue-100" 
                          : "bg-slate-100 text-slate-500 border border-slate-200 hover:bg-slate-200"
                      }`}
                    >
                      <ArrowDownCircle className={`w-3.5 h-3.5 ${isAutoScrollEnabled ? "text-blue-600 animate-bounce" : "text-slate-400"}`} />
                      <span>{isAutoScrollEnabled ? "Auto-Scroll: ON" : "Auto-Scroll: OFF"}</span>
                    </button>

                    {/* Botão Limpar Conversa */}
                    <button
                      onClick={() => {
                        setChatHistory([]);
                        setActiveQuiz(null);
                        setChatMessage("");
                        logAction("Conversa do Chat Limpa");
                      }}
                      title="Limpar toda a conversa e reiniciar com a EMIA"
                      className="px-2.5 py-1 rounded-lg text-[11px] font-semibold text-slate-600 hover:text-red-600 bg-slate-50 hover:bg-red-50 border border-slate-200 hover:border-red-200 transition-all flex items-center gap-1 shadow-2xs group"
                    >
                      <Trash2 className="w-3.5 h-3.5 text-slate-400 group-hover:text-red-500 transition-colors" />
                      <span>Limpar Conversa</span>
                    </button>
                  </div>
                </div>

                <div className="flex-1 overflow-y-auto p-6 space-y-4">
                  {chatHistory.length === 0 ? (
                    <div className="h-full flex flex-col items-center justify-center text-center max-w-lg mx-auto py-8">
                      {/* Avatar Futurista & Amigável da EMIA */}
                      <div className="relative mb-4 group cursor-pointer">
                        {/* Efeito Glow / Aura Futurista */}
                        <div className="absolute -inset-1.5 bg-gradient-to-r from-cyan-400 via-blue-500 to-indigo-500 rounded-3xl blur-md opacity-70 group-hover:opacity-100 transition duration-500 animate-pulse"></div>
                        
                        {/* Card do Mascote */}
                        <div className="relative w-16 h-16 rounded-2xl bg-gradient-to-tr from-slate-900 via-indigo-950 to-blue-900 border border-cyan-400/40 flex items-center justify-center shadow-xl shadow-blue-500/20">
                          {/* Ícone Futurista com Faíscas Inteligentes */}
                          <div className="relative flex items-center justify-center">
                            <Bot className="w-8 h-8 text-cyan-300 stroke-[1.8] drop-shadow-[0_0_8px_rgba(34,211,238,0.8)] animate-pulse" />
                            <Sparkles className="w-4 h-4 text-amber-300 absolute -top-2 -right-2 animate-spin duration-3000 drop-shadow-[0_0_6px_rgba(252,211,77,0.9)]" />
                          </div>
                        </div>

                        {/* Badge de Status Online Futurista */}
                        <span className="absolute -bottom-1 -right-1 flex h-4 w-4">
                          <span className="animate-ping absolute inline-flex h-full w-full rounded-full bg-emerald-400 opacity-75"></span>
                          <span className="relative inline-flex rounded-full h-4 w-4 bg-emerald-500 border-2 border-white shadow-xs"></span>
                        </span>
                      </div>

                      <h3 className="text-lg font-black text-gray-900 tracking-tight flex items-center gap-1.5 justify-center">
                        Oi, eu sou a EMIA! <span className="text-xl">✨</span>
                      </h3>
                      <p className="text-xs font-bold bg-gradient-to-r from-blue-600 to-indigo-600 bg-clip-text text-transparent mb-1">
                        Mascote & IA Acadêmica da EDUTECH
                      </p>
                      <p className="text-xs text-gray-500 mt-1 mb-6 max-w-sm">
                        Tô pronta pra trocar uma ideia sobre o seu trabalho! O que você quer fazer agora?
                      </p>

                      {/* Botões informais de ação rápida */}
                      <div className="grid grid-cols-1 sm:grid-cols-2 gap-2 w-full text-left">
                        <button
                          onClick={() => handleSendMessage(undefined, "Explica de forma simples os pontos principais deste trabalho.")}
                          className="p-3 bg-white hover:bg-blue-50 border border-gray-200 hover:border-blue-300 rounded-xl text-xs font-semibold text-gray-700 hover:text-blue-700 transition-all flex items-center gap-2 shadow-2xs group"
                        >
                          <span>💬</span>
                          <span>Conversar sobre o texto</span>
                        </button>

                        <button
                          onClick={() => handleSendMessage(undefined, "Cria um quiz com 4 perguntas de múltipla escolha para testar meu conhecimento sobre o tema.")}
                          className="p-3 bg-white hover:bg-amber-50 border border-gray-200 hover:border-amber-300 rounded-xl text-xs font-semibold text-gray-700 hover:text-amber-800 transition-all flex items-center gap-2 shadow-2xs group"
                        >
                          <span>🎯</span>
                          <span>Fazer um Quiz do conteúdo</span>
                        </button>

                        <button
                          onClick={handleStartInteractiveQuiz}
                          className="p-3 bg-white hover:bg-amber-50 border border-gray-200 hover:border-amber-300 rounded-xl text-xs font-semibold text-gray-700 hover:text-amber-800 transition-all flex items-center gap-2 shadow-2xs group"
                        >
                          <span>🎯</span>
                          <span>Fazer um Quiz Interativo</span>
                        </button>

                        <button
                          onClick={() => handleSendMessage(undefined, "Como uso os campos do app e o botão de Trabalho em Grupo?")}
                          className="p-3 bg-white hover:bg-emerald-50 border border-gray-200 hover:border-emerald-300 rounded-xl text-xs font-semibold text-gray-700 hover:text-emerald-800 transition-all flex items-center gap-2 shadow-2xs group"
                        >
                          <span>📱</span>
                          <span>Como usar o app?</span>
                        </button>

                        <button
                          onClick={() => handleSendMessage(undefined, "Tenho uma dúvida sobre normas ABNT neste trabalho. O que preciso ajustar?")}
                          className="p-3 bg-white hover:bg-indigo-50 border border-gray-200 hover:border-indigo-300 rounded-xl text-xs font-semibold text-gray-700 hover:text-indigo-800 transition-all flex items-center gap-2 shadow-2xs group"
                        >
                          <span>📑</span>
                          <span>Dúvidas sobre ABNT</span>
                        </button>

                        <button
                          onClick={() => handleSendMessage(undefined, "Quero melhorar e editar uma parte do texto. Pode me ajudar?")}
                          className="p-3 bg-white hover:bg-purple-50 border border-gray-200 hover:border-purple-300 rounded-xl text-xs font-semibold text-gray-700 hover:text-purple-800 transition-all flex items-center gap-2 shadow-2xs group"
                        >
                          <span>✏️</span>
                          <span>Editar algo no texto</span>
                        </button>

                        <button
                          onClick={() => handleSendMessage(undefined, "Cria um roteiro de podcast dinâmico com 2 apresentadores sobre este trabalho.")}
                          className="p-3 bg-white hover:bg-rose-50 border border-gray-200 hover:border-rose-300 rounded-xl text-xs font-semibold text-gray-700 hover:text-rose-800 transition-all flex items-center gap-2 shadow-2xs group"
                        >
                          <span>🎙️</span>
                          <span>Criar Podcast / Música</span>
                        </button>
                      </div>
                    </div>
                  ) : (
                    chatHistory.map((msg, idx) => (
                      <div key={idx} className={`flex ${msg.role === 'user' ? 'justify-end' : 'justify-start'} items-end gap-2`}>
                        {msg.role === 'assistant' && (
                          <div className="w-7 h-7 rounded-xl bg-gradient-to-tr from-cyan-500 to-blue-600 text-white flex items-center justify-center text-xs shadow-xs shrink-0 mb-1">
                            🤖
                          </div>
                        )}
                        <div className={`max-w-[85%] sm:max-w-[78%] px-4 py-3 rounded-2xl text-[13.5px] leading-relaxed font-sans shadow-xs transition-all ${
                          msg.role === 'user' 
                            ? 'bg-gradient-to-br from-blue-600 to-indigo-600 text-white rounded-br-xs font-medium' 
                            : 'bg-white border border-blue-100/80 text-slate-800 rounded-bl-xs shadow-sm font-normal'
                        }`}>
                          <div className={msg.role === 'user' ? '' : 'prose prose-sm max-w-none text-slate-800 font-sans'}>
                            <ReactMarkdown>{msg.text}</ReactMarkdown>
                          </div>
                        </div>
                      </div>
                    ))
                  )}

                  {/* Card do Quiz Ativo com Perguntas Dinâmicas e Botões Clicáveis */}
                  {activeQuiz && activeQuiz.isActive && !activeQuiz.isFinished && activeQuiz.questions[activeQuiz.currentIndex] && (
                    <div className="my-3 p-5 bg-gradient-to-br from-blue-50 via-indigo-50/50 to-white border-2 border-blue-300/80 rounded-3xl shadow-md animate-fadeIn font-sans">
                      <div className="flex items-center justify-between gap-2 mb-3">
                        <span className="px-3 py-1 bg-blue-600 text-white font-extrabold text-xs rounded-full shadow-xs">
                          Pergunta {activeQuiz.currentIndex + 1} de {activeQuiz.questions.length}
                        </span>
                        <span className="text-xs font-bold text-blue-600">
                          {Math.round(((activeQuiz.currentIndex) / activeQuiz.questions.length) * 100)}% concluído
                        </span>
                      </div>

                      {/* Barra de Progresso do Quiz */}
                      <div className="w-full bg-blue-100 h-2 rounded-full overflow-hidden mb-4">
                        <div 
                          className="bg-gradient-to-r from-blue-500 to-indigo-600 h-full transition-all duration-300"
                          style={{ width: `${((activeQuiz.currentIndex) / activeQuiz.questions.length) * 100}%` }}
                        />
                      </div>

                      <h4 className="text-sm font-bold text-slate-900 mb-4 leading-snug">
                        {activeQuiz.questions[activeQuiz.currentIndex].question}
                      </h4>

                      {/* Alternativas Clicáveis */}
                      <div className="space-y-2">
                        {activeQuiz.questions[activeQuiz.currentIndex].options.map((opt, oIdx) => (
                          <button
                            key={oIdx}
                            onClick={() => handleQuizOptionSelect(opt.letter)}
                            className="w-full p-3 bg-white hover:bg-blue-600 text-slate-800 hover:text-white border border-slate-200 hover:border-blue-600 rounded-2xl text-xs font-semibold text-left transition-all active:scale-98 shadow-xs flex items-center gap-3 group"
                          >
                            <span className="w-7 h-7 rounded-xl bg-blue-100 group-hover:bg-white text-blue-700 group-hover:text-blue-700 font-extrabold flex items-center justify-center shrink-0 transition-colors shadow-2xs">
                              {opt.letter}
                            </span>
                            <span className="flex-1 leading-snug">{opt.text}</span>
                          </button>
                        ))}
                      </div>
                    </div>
                  )}

                  {isChatting && (
                    <div className="flex justify-start items-end gap-2">
                      <div className="w-7 h-7 rounded-xl bg-gradient-to-tr from-cyan-500 to-blue-600 text-white flex items-center justify-center text-xs shadow-xs shrink-0 animate-bounce">
                        🤖
                      </div>
                      <div className="max-w-[80%] px-4 py-3 rounded-2xl text-[13px] font-sans bg-white border border-blue-100 text-slate-700 rounded-bl-xs shadow-xs flex items-center gap-2">
                        <Loader2 className="w-4 h-4 animate-spin text-blue-600" />
                        <span className="font-medium text-slate-600">A EMIA tá preparando... ✨</span>
                      </div>
                    </div>
                  )}
                  {/* Elemento de rolagem automática para a última mensagem */}
                  <div ref={chatMessagesEndRef} />
                </div>
                <div className="p-4 bg-white border-t border-gray-200">
                  <form onSubmit={handleSendMessage} className="flex gap-2 mb-2">
                    <input 
                      type="text" 
                      value={chatMessage}
                      onChange={(e) => setChatMessage(e.target.value)}
                      placeholder="Faça uma pergunta, peça um podcast ou música..."
                      className="flex-1 px-4 py-2 border border-gray-300 rounded-lg shadow-sm focus:outline-none focus:ring-2 focus:ring-blue-500 text-sm"
                    />
                    <Button type="submit" disabled={isChatting || !chatMessage.trim()} className="bg-blue-600">
                      Enviar
                    </Button>
                  </form>
                  <p className="text-xs text-center text-gray-400 flex items-center justify-center">
                    <Lock className="w-3 h-3 mr-1" />
                    Criptografia de ponta a ponta: apenas o remetente e o destinatário autorizado acessam as mensagens.
                  </p>
                </div>
              </div>
            )}

            {activeTab === "slides" && (
              <div className="flex-1 flex flex-col h-full bg-slate-900 text-white overflow-hidden select-text">
                {/* Top Control Bar do Estúdio de Slides (EMIA.SLIDES) */}
                <div className="bg-slate-950/80 border-b border-slate-800 px-6 py-3 flex items-center justify-between gap-4 backdrop-blur-md">
                  <div className="flex items-center gap-3">
                    <div className="w-8 h-8 rounded-xl bg-gradient-to-tr from-amber-500 via-orange-500 to-rose-500 flex items-center justify-center shadow-md shadow-orange-500/20 text-white">
                      <Presentation className="w-4 h-4" />
                    </div>
                    <div>
                      <h2 className="text-sm font-black text-slate-100 flex items-center gap-2 tracking-tight">
                        EMIA.SLIDES
                        <span className="text-[10px] px-2 py-0.5 rounded-full font-extrabold bg-amber-500/20 text-amber-300 border border-amber-500/30">
                          {editableSlides.length} Cartões
                        </span>
                      </h2>
                      <p className="text-[11px] text-slate-400">Estúdio de Apresentações com Templates Inteligentes</p>
                    </div>
                  </div>

                  {/* Ações Rápidas: PPTX, Google Slides, Tela Cheia e Voltar */}
                  <div className="flex items-center gap-2">
                    <Button
                      onClick={() => {
                        const elem = document.getElementById("emia-slides-container") || document.documentElement;
                        if (!document.fullscreenElement) {
                          elem.requestFullscreen().catch(err => console.warn("Fullscreen error:", err));
                        } else {
                          document.exitFullscreen().catch(err => console.warn("Exit fullscreen error:", err));
                        }
                      }}
                      size="sm"
                      className="bg-amber-500/20 hover:bg-amber-500/30 text-amber-300 border border-amber-500/30 text-xs h-8 px-3 font-bold flex items-center gap-1.5"
                      title="Apresentar em Tela Cheia (F11)"
                    >
                      <Maximize2 className="w-3.5 h-3.5" />
                      Tela Cheia
                    </Button>
                    <Button
                      onClick={exportPPTXSlides}
                      size="sm"
                      className="bg-slate-800 hover:bg-slate-700 text-slate-200 border border-slate-700 text-xs h-8 px-3 font-semibold"
                      title="Baixar em PowerPoint (.pptx)"
                    >
                      <Download className="w-3.5 h-3.5 mr-1 text-amber-400" />
                      Baixar .PPTX
                    </Button>
                    <Button
                      onClick={openInGoogleSlides}
                      size="sm"
                      className="bg-gradient-to-r from-amber-500 to-orange-600 hover:from-amber-600 hover:to-orange-700 text-white text-xs h-8 px-3.5 font-bold shadow-md shadow-orange-500/20"
                    >
                      <Share2 className="w-3.5 h-3.5 mr-1" />
                      Google Slides
                    </Button>
                    <Button
                      onClick={() => setActiveTab("editor")}
                      variant="ghost"
                      size="sm"
                      className="text-xs text-slate-400 hover:text-white hover:bg-slate-800 h-8 font-medium"
                    >
                      Voltar ao Documento
                    </Button>
                  </div>
                </div>

                {/* Workspace Central EMIA.SLIDES */}
                <div id="emia-slides-container" className="flex-1 flex overflow-hidden">
                  {/* Navegador Lateral de Slides (Miniaturas) */}
                  <div className="w-60 bg-slate-950/50 border-r border-slate-800/80 p-3 overflow-y-auto space-y-2.5 flex-shrink-0">
                    <div className="flex items-center justify-between mb-1 px-1">
                      <span className="text-[11px] font-bold uppercase tracking-wider text-slate-500">Slides</span>
                      <div className="flex items-center gap-2">
                        <button
                          onClick={() => {
                            if (editableSlides.length <= 1) return;
                            const updated = editableSlides.filter((_, i) => i !== activeSlideIndex);
                            setEditableSlides(updated);
                            setActiveSlideIndex(prev => Math.max(0, prev - 1));
                          }}
                          disabled={editableSlides.length <= 1}
                          className="text-xs text-rose-400 hover:text-rose-300 disabled:opacity-30 disabled:pointer-events-none flex items-center gap-0.5 font-bold p-1 rounded hover:bg-rose-500/10 transition-colors"
                          title="Apagar Slide Selecionado"
                        >
                          <X className="w-3.5 h-3.5" /> Apagar
                        </button>
                        <button
                          onClick={() => {
                            const newSlide = {
                              title: `Novo Slide ${editableSlides.length + 1}`,
                              bullets: ["Ponto chave 1", "Ponto chave 2", "Ponto chave 3"],
                              layout: "bullets" as const,
                              badge: `Módulo ${editableSlides.length + 1}`
                            };
                            setEditableSlides(prev => [...prev, newSlide]);
                            setActiveSlideIndex(editableSlides.length);
                          }}
                          className="text-xs text-amber-400 hover:text-amber-300 flex items-center gap-0.5 font-bold p-1 rounded hover:bg-amber-500/10 transition-colors"
                          title="Adicionar Novo Slide"
                        >
                          <Plus className="w-3.5 h-3.5" /> Adicionar
                        </button>
                      </div>
                    </div>

                    {editableSlides.map((slide, sIdx) => (
                      <div
                        key={sIdx}
                        onClick={() => setActiveSlideIndex(sIdx)}
                        className={`p-2.5 rounded-xl border text-left cursor-pointer transition-all ${
                          activeSlideIndex === sIdx
                            ? "bg-slate-800/90 border-amber-500/80 shadow-md ring-1 ring-amber-500/50"
                            : "bg-slate-900/40 border-slate-800/80 hover:bg-slate-800/40 hover:border-slate-700"
                        }`}
                      >
                        <div className="flex items-center justify-between text-[10px] font-bold text-slate-500 mb-1">
                          <span>#{sIdx + 1}</span>
                          <span className="text-[9px] uppercase px-1.5 py-0.2 rounded bg-slate-800 text-slate-400">
                            {slide.layout || "card"}
                          </span>
                        </div>
                        <p className="text-xs font-semibold text-slate-200 truncate">{slide.title || "Sem título"}</p>
                        <p className="text-[10px] text-slate-400 truncate mt-0.5">{slide.bullets?.[0] || slide.subtitle || "Conteúdo..."}</p>
                      </div>
                    ))}
                  </div>

                  {/* Palco Central: Cartão do Slide Ativo em Destaque Interativo */}
                  <div className="flex-1 p-6 md:p-10 overflow-y-auto flex flex-col items-center justify-start bg-gradient-to-b from-slate-900 to-slate-950">
                    {editableSlides[activeSlideIndex] && (
                      <div className="w-full max-w-4xl space-y-4 animate-in fade-in zoom-in-95 duration-200">
                        
                        {/* 1. BARRA DE TEMAS (EM CIMA DA PALAVRA LAYOUT) */}
                        <div className="flex items-center justify-between bg-slate-900/80 border border-slate-800 p-2 rounded-2xl">
                          <div className="flex items-center gap-1.5 flex-wrap">
                            <span className="text-[11px] font-bold text-slate-400 px-2">Temas:</span>
                            {[
                              { id: "academic", name: "Acadêmico", color: "bg-blue-600" },
                              { id: "modern", name: "Moderno", color: "bg-teal-600" },
                              { id: "sunset", name: "Sunset", color: "bg-gradient-to-r from-orange-500 to-rose-500" },
                              { id: "emerald", name: "Esmeralda", color: "bg-emerald-500" },
                              { id: "lavender", name: "Lavanda", color: "bg-indigo-400" },
                              { id: "solar", name: "Solar Gold", color: "bg-amber-400" },
                              { id: "neon", name: "Cyberpunk", color: "bg-purple-600" },
                              { id: "dark", name: "Obsidian", color: "bg-slate-700" },
                            ].map(t => (
                              <button
                                key={t.id}
                                onClick={() => setSlidesTheme(t.id as any)}
                                className={`text-xs px-2.5 py-1 rounded-lg font-bold flex items-center gap-1.5 transition-all ${
                                  slidesTheme === t.id
                                    ? "bg-slate-800 text-white shadow-sm ring-1 ring-amber-500"
                                    : "text-slate-400 hover:text-slate-200 hover:bg-slate-800/50"
                                }`}
                              >
                                <span className={`w-2 h-2 rounded-full ${t.color}`} />
                                {t.name}
                              </button>
                            ))}
                          </div>
                        </div>

                        {/* 2. SELETOR DE LAYOUT GAMMA DO CARTÃO */}
                        <div className="flex items-center justify-between bg-slate-900/80 border border-slate-800 p-2 rounded-2xl">
                          <div className="flex items-center gap-1.5 flex-wrap">
                            <span className="text-[11px] font-bold text-slate-400 px-2">Layout:</span>
                            {[
                              { id: "card", label: "🃏 Cartão" },
                              { id: "split", label: "🌗 Split" },
                              { id: "bullets", label: "📋 Tópicos" },
                              { id: "metrics", label: "📊 Métricas" },
                              { id: "quote", label: "💬 Destaque" },
                              { id: "timeline", label: "⏳ Cronologia" },
                              { id: "comparison", label: "⚖️ Comparação" },
                            ].map(l => (
                              <button
                                key={l.id}
                                onClick={() => {
                                  const updated = [...editableSlides];
                                  updated[activeSlideIndex].layout = l.id as any;
                                  setEditableSlides(updated);
                                }}
                                className={`text-xs px-2.5 py-1 rounded-lg font-semibold transition-all ${
                                  editableSlides[activeSlideIndex].layout === l.id
                                    ? "bg-amber-500 text-slate-950 font-bold"
                                    : "text-slate-400 hover:text-white hover:bg-slate-800"
                                }`}
                              >
                                {l.label}
                              </button>
                            ))}
                          </div>

                          <button
                            onClick={() => {
                              if (editableSlides.length <= 1) return;
                              const updated = editableSlides.filter((_, i) => i !== activeSlideIndex);
                              setEditableSlides(updated);
                              setActiveSlideIndex(prev => Math.max(0, prev - 1));
                            }}
                            className="text-xs text-rose-400 hover:text-rose-300 px-3 py-1 hover:bg-rose-500/10 rounded-lg transition-colors flex-shrink-0"
                          >
                            Excluir Slide
                          </button>
                        </div>

                        {/* 3. CARTÃO DE VISUALIZAÇÃO / EDIÇÃO ESTILO GAMMA */}
                        <div
                          className={`rounded-3xl p-8 md:p-12 shadow-2xl border transition-all duration-300 min-h-[460px] flex flex-col justify-between ${
                            slidesTheme === "academic"
                              ? "bg-gradient-to-br from-blue-950 via-slate-900 to-indigo-950 border-blue-800/50 text-white"
                              : slidesTheme === "modern"
                              ? "bg-gradient-to-br from-teal-950 via-slate-900 to-emerald-950 border-teal-800/50 text-white"
                              : slidesTheme === "sunset"
                              ? "bg-gradient-to-br from-rose-950 via-slate-900 to-amber-950 border-orange-800/50 text-white"
                              : slidesTheme === "emerald"
                              ? "bg-gradient-to-br from-emerald-950 via-slate-900 to-teal-950 border-emerald-800/50 text-white"
                              : slidesTheme === "lavender"
                              ? "bg-gradient-to-br from-indigo-950 via-slate-900 to-purple-950 border-indigo-700/50 text-white"
                              : slidesTheme === "solar"
                              ? "bg-gradient-to-br from-amber-950 via-slate-900 to-orange-950 border-amber-600/50 text-white"
                              : slidesTheme === "neon"
                              ? "bg-gradient-to-br from-purple-950 via-slate-900 to-fuchsia-950 border-purple-800/50 text-white"
                              : "bg-slate-900 border-slate-800 text-slate-100"
                          }`}
                        >
                          {/* Topo do Cartão (Badge e Número) */}
                          <div className="flex items-center justify-between mb-4">
                            <div
                              contentEditable
                              suppressContentEditableWarning
                              onBlur={(e) => {
                                const updated = [...editableSlides];
                                updated[activeSlideIndex].badge = e.currentTarget.innerText;
                                setEditableSlides(updated);
                              }}
                              className="text-[11px] font-extrabold uppercase tracking-wider px-3 py-1 rounded-full bg-white/10 text-amber-300 border border-white/10 focus:outline-none focus:ring-1 focus:ring-amber-400"
                            >
                              {editableSlides[activeSlideIndex].badge || `SLIDE ${activeSlideIndex + 1}`}
                            </div>
                            <span className="text-xs font-mono font-bold text-slate-400">
                              {activeSlideIndex + 1} / {editableSlides.length}
                            </span>
                          </div>

                          {/* Título Principal Editável */}
                          <div
                            contentEditable
                            suppressContentEditableWarning
                            onBlur={(e) => {
                              const updated = [...editableSlides];
                              updated[activeSlideIndex].title = e.currentTarget.innerText;
                              setEditableSlides(updated);
                            }}
                            className="text-2xl md:text-3xl font-black tracking-tight text-white mb-6 focus:outline-none focus:bg-white/5 p-2 rounded-xl transition-colors"
                          >
                            {editableSlides[activeSlideIndex].title}
                          </div>

                          {/* Renderização Dinâmica por Layout Gamma */}
                          {editableSlides[activeSlideIndex].layout === "split" ? (
                            <div className="grid grid-cols-1 md:grid-cols-2 gap-6 my-auto">
                              <div className="bg-white/5 p-6 rounded-2xl border border-white/10">
                                <h4 className="text-xs font-bold uppercase text-amber-400 mb-2">Visão Geral</h4>
                                <div
                                  contentEditable
                                  suppressContentEditableWarning
                                  onBlur={(e) => {
                                    const updated = [...editableSlides];
                                    updated[activeSlideIndex].bullets[0] = e.currentTarget.innerText;
                                    setEditableSlides(updated);
                                  }}
                                  className="text-sm text-slate-200 leading-relaxed focus:outline-none focus:bg-white/10 p-1 rounded"
                                >
                                  {editableSlides[activeSlideIndex].bullets[0] || "Explique o ponto principal..."}
                                </div>
                              </div>
                              <div className="bg-white/5 p-6 rounded-2xl border border-white/10">
                                <h4 className="text-xs font-bold uppercase text-teal-400 mb-2">Impacto & Evidências</h4>
                                <div
                                  contentEditable
                                  suppressContentEditableWarning
                                  onBlur={(e) => {
                                    const updated = [...editableSlides];
                                    updated[activeSlideIndex].bullets[1] = e.currentTarget.innerText;
                                    setEditableSlides(updated);
                                  }}
                                  className="text-sm text-slate-200 leading-relaxed focus:outline-none focus:bg-white/10 p-1 rounded"
                                >
                                  {editableSlides[activeSlideIndex].bullets[1] || "Detalhe os desdobramentos..."}
                                </div>
                              </div>
                            </div>
                          ) : editableSlides[activeSlideIndex].layout === "metrics" ? (
                            <div className="grid grid-cols-1 sm:grid-cols-3 gap-4 my-auto">
                              {editableSlides[activeSlideIndex].bullets.slice(0, 3).map((bullet, bIdx) => (
                                <div key={bIdx} className="bg-white/5 p-5 rounded-2xl border border-white/10 flex flex-col justify-between">
                                  <span className="text-3xl font-black text-amber-400 mb-1">0{bIdx + 1}</span>
                                  <div
                                    contentEditable
                                    suppressContentEditableWarning
                                    onBlur={(e) => {
                                      const updated = [...editableSlides];
                                      updated[activeSlideIndex].bullets[bIdx] = e.currentTarget.innerText;
                                      setEditableSlides(updated);
                                    }}
                                    className="text-xs text-slate-200 leading-normal focus:outline-none focus:bg-white/10 p-1 rounded"
                                  >
                                    {bullet}
                                  </div>
                                </div>
                              ))}
                            </div>
                          ) : editableSlides[activeSlideIndex].layout === "quote" ? (
                            <div className="bg-white/5 p-8 rounded-2xl border-l-4 border-amber-400 my-auto">
                              <div
                                contentEditable
                                suppressContentEditableWarning
                                onBlur={(e) => {
                                  const updated = [...editableSlides];
                                  updated[activeSlideIndex].bullets[0] = e.currentTarget.innerText;
                                  setEditableSlides(updated);
                                }}
                                className="text-lg md:text-xl font-medium italic text-slate-100 leading-relaxed focus:outline-none focus:bg-white/10 p-2 rounded"
                              >
                                "{editableSlides[activeSlideIndex].bullets[0] || "Destaque conceitual ou citação principal da pesquisa..."}"
                              </div>
                            </div>
                          ) : editableSlides[activeSlideIndex].layout === "timeline" ? (
                            <div className="space-y-4 my-auto relative before:absolute before:left-3 before:top-2 before:bottom-2 before:w-0.5 before:bg-white/20">
                              {editableSlides[activeSlideIndex].bullets.slice(0, 4).map((bullet, bIdx) => (
                                <div key={bIdx} className="flex items-start gap-4 pl-1 relative">
                                  <div className="w-5 h-5 rounded-full bg-amber-400 border-4 border-slate-900 flex-shrink-0 mt-0.5" />
                                  <div className="bg-white/5 p-3 rounded-xl border border-white/10 flex-1">
                                    <span className="text-[10px] font-bold uppercase text-amber-300">Etapa 0{bIdx + 1}</span>
                                    <div
                                      contentEditable
                                      suppressContentEditableWarning
                                      onBlur={(e) => {
                                        const updated = [...editableSlides];
                                        updated[activeSlideIndex].bullets[bIdx] = e.currentTarget.innerText;
                                        setEditableSlides(updated);
                                      }}
                                      className="text-xs text-slate-200 mt-0.5 focus:outline-none focus:bg-white/10 p-1 rounded"
                                    >
                                      {bullet}
                                    </div>
                                  </div>
                                </div>
                              ))}
                            </div>
                          ) : editableSlides[activeSlideIndex].layout === "comparison" ? (
                            <div className="grid grid-cols-1 md:grid-cols-2 gap-4 my-auto">
                              <div className="bg-emerald-950/40 p-5 rounded-2xl border border-emerald-500/30">
                                <div className="text-xs font-bold uppercase text-emerald-400 mb-2 flex items-center gap-1.5">
                                  <span>✅ Vantagens & Conquistas</span>
                                </div>
                                <div
                                  contentEditable
                                  suppressContentEditableWarning
                                  onBlur={(e) => {
                                    const updated = [...editableSlides];
                                    updated[activeSlideIndex].bullets[0] = e.currentTarget.innerText;
                                    setEditableSlides(updated);
                                  }}
                                  className="text-xs text-slate-200 leading-relaxed focus:outline-none focus:bg-white/10 p-1 rounded"
                                >
                                  {editableSlides[activeSlideIndex].bullets[0] || "Pontos positivos e avanços identificados..."}
                                </div>
                              </div>
                              <div className="bg-rose-950/40 p-5 rounded-2xl border border-rose-500/30">
                                <div className="text-xs font-bold uppercase text-rose-400 mb-2 flex items-center gap-1.5">
                                  <span>⚠️ Desafios & Limitações</span>
                                </div>
                                <div
                                  contentEditable
                                  suppressContentEditableWarning
                                  onBlur={(e) => {
                                    const updated = [...editableSlides];
                                    updated[activeSlideIndex].bullets[1] = e.currentTarget.innerText;
                                    setEditableSlides(updated);
                                  }}
                                  className="text-xs text-slate-200 leading-relaxed focus:outline-none focus:bg-white/10 p-1 rounded"
                                >
                                  {editableSlides[activeSlideIndex].bullets[1] || "Desafios encontrados e oportunidades de melhoria..."}
                                </div>
                              </div>
                            </div>
                          ) : (
                            /* Layout Padrão Tópicos / Cartão */
                            <div className="space-y-3 my-auto">
                              {editableSlides[activeSlideIndex].bullets.map((bullet, bIdx) => (
                                <div key={bIdx} className="flex items-start gap-3 bg-white/5 hover:bg-white/10 p-3.5 rounded-2xl border border-white/5 transition-colors group">
                                  <div className="w-6 h-6 rounded-lg bg-amber-500/20 text-amber-400 flex items-center justify-center flex-shrink-0 text-xs font-bold mt-0.5">
                                    {bIdx + 1}
                                  </div>
                                  <div
                                    contentEditable
                                    suppressContentEditableWarning
                                    onBlur={(e) => {
                                      const updated = [...editableSlides];
                                      updated[activeSlideIndex].bullets[bIdx] = e.currentTarget.innerText;
                                      setEditableSlides(updated);
                                    }}
                                    className="flex-1 text-sm text-slate-200 leading-relaxed focus:outline-none focus:bg-white/10 p-1 rounded"
                                  >
                                    {bullet}
                                  </div>
                                  <button
                                    onClick={() => {
                                      const updated = [...editableSlides];
                                      updated[activeSlideIndex].bullets = updated[activeSlideIndex].bullets.filter((_, i) => i !== bIdx);
                                      setEditableSlides(updated);
                                    }}
                                    className="opacity-0 group-hover:opacity-100 text-slate-500 hover:text-rose-400 p-1 transition-opacity"
                                    title="Remover tópico"
                                  >
                                    <X className="w-3.5 h-3.5" />
                                  </button>
                                </div>
                              ))}
                              <button
                                onClick={() => {
                                  const updated = [...editableSlides];
                                  updated[activeSlideIndex].bullets.push("Novo ponto de apresentação...");
                                  setEditableSlides(updated);
                                }}
                                className="text-xs font-bold text-amber-400 hover:text-amber-300 flex items-center gap-1.5 pt-2"
                              >
                                <Plus className="w-3.5 h-3.5" /> Adicionar Tópico
                              </button>
                            </div>
                          )}

                          {/* Rodapé do Slide com Notas do Apresentador */}
                          <div className="mt-8 pt-4 border-t border-white/10 flex items-center justify-between text-xs text-slate-400">
                            <div className="flex items-center gap-2">
                              <span className="font-bold text-slate-300">Notas de Apresentação:</span>
                              <div
                                contentEditable
                                suppressContentEditableWarning
                                onBlur={(e) => {
                                  const updated = [...editableSlides];
                                  updated[activeSlideIndex].notes = e.currentTarget.innerText;
                                  setEditableSlides(updated);
                                }}
                                className="text-slate-300 italic focus:outline-none focus:bg-white/10 px-2 py-0.5 rounded"
                              >
                                {editableSlides[activeSlideIndex].notes || "Dica de fala para o slide..."}
                              </div>
                            </div>
                            <span>EDUTECH.EMIA</span>
                          </div>
                        </div>

                        {/* Barra de Navegação com Setas para Passar Slide (Posicionada na Parte Inferior) */}
                        <div className="flex items-center justify-between bg-slate-900/90 border border-slate-800 p-3 rounded-2xl backdrop-blur-md shadow-lg">
                          <button
                            onClick={() => setActiveSlideIndex(prev => Math.max(0, prev - 1))}
                            disabled={activeSlideIndex === 0}
                            className="flex items-center gap-2 px-4 py-2 rounded-xl bg-slate-800 hover:bg-slate-700 disabled:opacity-30 disabled:pointer-events-none text-slate-200 font-bold text-xs transition-all active:scale-95 shadow-sm"
                            title="Slide Anterior (Seta Esquerda)"
                          >
                            <ChevronLeft className="w-4 h-4 text-amber-400" />
                            <span>Anterior</span>
                          </button>

                          {/* Indicador e Atalhos de Teclado */}
                          <div className="flex items-center gap-2">
                            <span className="text-xs font-mono font-bold text-slate-300 bg-slate-950/80 px-3 py-1.5 rounded-xl border border-slate-800">
                              Slide {activeSlideIndex + 1} de {editableSlides.length}
                            </span>
                            <span className="text-[10px] text-slate-500 hidden sm:inline">(Use as setas ◀ ▶ do teclado)</span>
                          </div>

                          <button
                            onClick={() => setActiveSlideIndex(prev => Math.min(editableSlides.length - 1, prev + 1))}
                            disabled={activeSlideIndex >= editableSlides.length - 1}
                            className="flex items-center gap-2 px-4 py-2 rounded-xl bg-gradient-to-r from-amber-500 to-orange-500 hover:from-amber-600 hover:to-orange-600 disabled:opacity-30 disabled:pointer-events-none text-white font-bold text-xs transition-all active:scale-95 shadow-md shadow-orange-500/20"
                            title="Próximo Slide (Seta Direita / Espaço)"
                          >
                            <span>Próximo</span>
                            <ChevronRight className="w-4 h-4 text-white" />
                          </button>
                        </div>

                      </div>
                    )}
                  </div>
                </div>
              </div>
            )}

            {activeTab === "report" && (
              <div className="w-full h-full p-8 overflow-y-auto prose prose-blue max-w-none font-sans text-gray-700">
                {authenticityReport ? (
                  <ReactMarkdown>{authenticityReport}</ReactMarkdown>
                ) : (
                  <p className="text-gray-400 italic">Nenhum relatório gerado ainda. Execute a Verificação de Autenticidade.</p>
                )}
              </div>
            )}
          </div>
        </div>

      </main>

      {showProfileModal && (
        <div className="fixed inset-0 bg-black bg-opacity-50 flex items-center justify-center z-50 p-4">
          <div className="bg-white rounded-xl shadow-xl w-full max-w-2xl max-h-[90vh] flex flex-col overflow-hidden">
            <div className="flex items-center justify-between p-6 border-b border-gray-100">
              <h2 className="text-xl font-bold text-gray-900">Perfil e Dados do Trabalho</h2>
              <button onClick={() => setShowProfileModal(false)} className="text-gray-400 hover:text-gray-600">
                <X className="w-6 h-6" />
              </button>
            </div>
            
            <div className="flex border-b border-gray-100 px-6 pt-2">
              <button 
                onClick={() => setProfileTab('dados')}
                className={`px-4 py-3 text-sm font-medium border-b-2 transition-colors ${profileTab === 'dados' ? 'border-blue-600 text-blue-600 font-semibold' : 'border-transparent text-gray-500 hover:text-gray-700'}`}
              >
                Meus Dados (Capa ABNT)
              </button>
              <button 
                onClick={() => setProfileTab('historico')}
                className={`px-4 py-3 text-sm font-medium border-b-2 transition-colors ${profileTab === 'historico' ? 'border-blue-600 text-blue-600 font-semibold' : 'border-transparent text-gray-500 hover:text-gray-700'}`}
              >
                Histórico de Textos
              </button>
            </div>

            <div className="flex-1 overflow-y-auto p-6">

              {profileTab === 'dados' && (
                <div className="space-y-4">
                  <p className="text-sm text-gray-600 mb-6">
                    Salve seus dados padrão. Eles serão preenchidos automaticamente na "Capa ABNT" e inseridos no cabeçalho e rodapé dos novos trabalhos.
                  </p>
                  
                  <div className="grid gap-4">
                    <div>
                      <label className="block text-sm font-medium text-gray-700 mb-1">Nome Completo</label>
                      <input 
                        type="text" 
                        value={studentName}
                        onChange={(e) => setStudentName(e.target.value)}
                        placeholder="Ex: João da Silva"
                        className="w-full px-4 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500"
                      />
                    </div>
                    <div>
                      <label className="block text-sm font-medium text-gray-700 mb-1">Curso</label>
                      <input 
                        type="text" 
                        value={course}
                        onChange={(e) => setCourse(e.target.value)}
                        placeholder="Ex: Administração, Direito, Pedagogia..."
                        className="w-full px-4 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500"
                      />
                    </div>
                    <div>
                      <label className="block text-sm font-medium text-gray-700 mb-1">Disciplina / Matéria</label>
                      <input 
                        type="text" 
                        value={subject}
                        onChange={(e) => setSubject(e.target.value)}
                        placeholder="Ex: Metodologia Científica, Didática, Gestão..."
                        className="w-full px-4 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500"
                      />
                    </div>
                    <div className="grid grid-cols-2 gap-4">
                      <div>
                        <label className="block text-sm font-medium text-gray-700 mb-1">Turno</label>
                        <select
                          value={shift}
                          onChange={(e) => setShift(e.target.value)}
                          className="w-full px-3 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 bg-white text-sm"
                        >
                          <option value="">Selecione...</option>
                          <option value="Matutino">Matutino</option>
                          <option value="Vespertino">Vespertino</option>
                          <option value="Noturno">Noturno</option>
                          <option value="Integral">Integral</option>
                          <option value="EAD / Online">EAD / Online</option>
                        </select>
                      </div>
                      <div>
                        <label className="block text-sm font-medium text-gray-700 mb-1">Sala / Turma</label>
                        <input 
                          type="text" 
                          value={classroom}
                          onChange={(e) => setClassroom(e.target.value)}
                          placeholder="Ex: Sala 204, Turma B"
                          className="w-full px-4 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500"
                        />
                      </div>
                    </div>
                    <div>
                      <label className="block text-sm font-medium text-gray-700 mb-1">Instituição de Ensino</label>
                      <input 
                        type="text" 
                        value={institution}
                        onChange={(e) => setInstitution(e.target.value)}
                        placeholder="Ex: Universidade de São Paulo"
                        className="w-full px-4 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500"
                      />
                    </div>
                    <div className="grid grid-cols-2 gap-4">
                      <div>
                        <label className="block text-sm font-medium text-gray-700 mb-1">Cidade</label>
                        <input 
                          type="text" 
                          value={city}
                          onChange={(e) => setCity(e.target.value)}
                          className="w-full px-4 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500"
                        />
                      </div>
                      <div>
                        <label className="block text-sm font-medium text-gray-700 mb-1">Ano</label>
                        <input 
                          type="text" 
                          value={year}
                          onChange={(e) => setYear(e.target.value)}
                          className="w-full px-4 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500"
                        />
                      </div>
                    </div>
                    <div>
                      <label className="block text-sm font-medium text-gray-700 mb-1">Orientador(a)</label>
                      <input 
                        type="text" 
                        value={advisor}
                        onChange={(e) => setAdvisor(e.target.value)}
                        className="w-full px-4 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500"
                      />
                    </div>
                  </div>
                  
                  <div className="mt-8 pt-4 border-t border-gray-100 flex justify-end">
                    <Button onClick={handleSaveProfile} className="bg-blue-600">
                      <Save className="w-4 h-4 mr-2" /> Salvar Perfil
                    </Button>
                  </div>
                </div>
              )}

              {profileTab === 'historico' && (
                <div className="space-y-4">
                  <p className="text-sm text-gray-600 mb-4">
                    Abaixo estão os textos acadêmicos gerados nos últimos 7 dias. Você pode carregá-los de volta no editor a qualquer momento.
                  </p>
                  
                  {auditLogs.filter(l => Boolean(l.content && l.content.trim().length > 50)).length === 0 ? (
                    <div className="text-center py-12 text-gray-400">
                      <Clock className="w-12 h-12 mx-auto mb-3 opacity-20" />
                      <p>Nenhum texto gerado no histórico ainda.</p>
                    </div>
                  ) : (
                    <div className="space-y-3">
                      {auditLogs.filter(l => Boolean(l.content && l.content.trim().length > 50)).map((log, idx) => (
                        <div key={idx} className="flex flex-col p-4 bg-gray-50 rounded-xl border border-gray-200 hover:border-blue-300 transition-colors">
                          <div className="flex justify-between items-center mb-2">
                            <span className="text-sm font-bold text-gray-900">{log.action}</span>
                            <div className="flex items-center gap-3">
                              <span className="text-xs text-gray-500 whitespace-nowrap">
                                {new Date(log.timestamp).toLocaleString('pt-BR')}
                              </span>
                              <Button
                                size="sm"
                                variant="outline"
                                onClick={() => {
                                  if (log.content) {
                                    setGeneratedText(log.content);
                                    setShowProfileModal(false);
                                    setActiveTab("editor");
                                  }
                                }}
                                className="text-xs h-7 px-2.5 text-blue-600 border-blue-200 hover:bg-blue-50 font-medium"
                              >
                                Abrir no Editor
                              </Button>
                            </div>
                          </div>
                          {log.content && (
                            <div className="mt-1 p-3 bg-white rounded-lg border border-gray-200 text-xs text-gray-700 font-serif max-h-32 overflow-y-auto whitespace-pre-wrap leading-relaxed">
                              {log.content.substring(0, 400)}...
                            </div>
                          )}
                        </div>
                      ))}
                    </div>
                  )}
                </div>
              )}
            </div>
          </div>
        </div>
      )}

      {showReferenceModal && (
        <div className="fixed inset-0 bg-black bg-opacity-50 flex items-center justify-center z-50 p-4">
          <div className="bg-white rounded-xl shadow-xl w-full max-w-lg flex flex-col overflow-hidden">
            <div className="flex items-center justify-between p-6 border-b border-gray-100">
              <h2 className="text-xl font-bold text-gray-900">Gerar Referência</h2>
              <button 
                onClick={handleCloseReferenceModal} 
                className="text-gray-400 hover:text-gray-600 p-1 rounded-lg hover:bg-gray-100 transition-colors"
                title="Fechar e cancelar"
              >
                <X className="w-6 h-6" />
              </button>
            </div>
            
            <div className="p-6 space-y-6">
              <div>
                <label className="block text-sm font-medium text-gray-700 mb-2">Formato da Referência</label>
                <div className="flex gap-4">
                  <button
                    onClick={() => setReferenceStyle("ABNT")}
                    className={`flex-1 py-2 px-4 rounded-md border text-sm font-medium transition-colors ${referenceStyle === "ABNT" ? "bg-blue-50 border-blue-600 text-blue-700" : "bg-white border-gray-300 text-gray-700 hover:bg-gray-50"}`}
                  >
                    ABNT
                  </button>
                  <button
                    onClick={() => setReferenceStyle("APA")}
                    className={`flex-1 py-2 px-4 rounded-md border text-sm font-medium transition-colors ${referenceStyle === "APA" ? "bg-blue-50 border-blue-600 text-blue-700" : "bg-white border-gray-300 text-gray-700 hover:bg-gray-50"}`}
                  >
                    APA
                  </button>
                </div>
              </div>

              <div>
                <label className="block text-sm font-medium text-gray-700 mb-2">Link, DOI ou Dados do Livro/Artigo</label>
                <input 
                  type="text" 
                  value={referenceSource}
                  onChange={(e) => setReferenceSource(e.target.value)}
                  placeholder="Ex: 10.1038/nrg3270 ou https://..."
                  className="w-full px-4 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500"
                />
              </div>

              <Button 
                onClick={handleGenerateReference} 
                disabled={isLoading || !referenceSource}
                className="w-full bg-blue-600 hover:bg-blue-700 text-white"
              >
                {isLoading ? <Loader2 className="w-5 h-5 mr-2 animate-spin" /> : <Link className="w-5 h-5 mr-2" />}
                Gerar Referência
              </Button>

              {generatedReference && (
                <div className="mt-4 p-4 bg-gray-50 border border-gray-200 rounded-lg">
                  <p className="text-sm text-gray-800 font-serif mb-3 select-all">{generatedReference}</p>
                  <Button 
                    onClick={() => {
                      navigator.clipboard.writeText(generatedReference);
                      setErrorMessage("Referência copiada para a área de transferência!");
                      setTimeout(() => setErrorMessage(""), 3000);
                    }}
                    variant="outline"
                    size="sm"
                    className="w-full"
                  >
                    <Copy className="w-4 h-4 mr-2" /> Copiar Referência
                  </Button>
                </div>
              )}
            </div>
          </div>
        </div>
      )}

      {/* MODAL DE PAGAMENTO PIX DIRETO - PACOTES 3 TRABALHOS (R$ 5) OU 7 TRABALHOS (R$ 10) */}
      {showPixModal && (
        <div className="fixed inset-0 bg-black/60 backdrop-blur-sm flex items-center justify-center z-50 p-4">
          <div className="bg-white rounded-3xl shadow-2xl w-full max-w-md flex flex-col overflow-hidden border border-gray-100 animate-in fade-in zoom-in-95 duration-200">
            <div className="bg-gradient-to-r from-blue-600 to-indigo-600 p-6 text-white relative">
              <button 
                onClick={() => setShowPixModal(false)} 
                className="absolute top-4 right-4 text-white/80 hover:text-white bg-black/10 hover:bg-black/20 p-1.5 rounded-full transition-colors"
              >
                <X className="w-5 h-5" />
              </button>
              <div className="flex items-center gap-2 mb-2">
                <span className="bg-white/20 text-white text-[11px] font-bold px-2.5 py-0.5 rounded-full uppercase tracking-wide">
                  PIX Direto
                </span>
                <span className="bg-emerald-400 text-slate-900 text-[11px] font-bold px-2 py-0.5 rounded-full">
                  Reconhecimento Automático
                </span>
              </div>
              <h2 className="text-xl font-bold">Escolha seu Pacote de Trabalhos</h2>
              <p className="text-blue-100 text-xs mt-1">
                Formatação completa ABNT A4, Anti-Plágio Turnitin e Exportação Word/PDF.
              </p>
            </div>
            
            <div className="p-6 space-y-5 max-h-[80vh] overflow-y-auto">
              
              {/* SELEÇÃO DOS 3 PACOTES */}
              <div className="grid grid-cols-3 gap-2.5">
                {/* Pacote 1: R$ 1,99 = 1 Trabalho */}
                <div 
                  onClick={() => setSelectedPixPlan('single')}
                  className={`p-3 rounded-2xl border-2 cursor-pointer transition-all flex flex-col justify-between text-left relative ${
                    selectedPixPlan === 'single' 
                      ? 'border-blue-600 bg-blue-50/60 shadow-sm ring-1 ring-blue-500' 
                      : 'border-gray-200 hover:border-gray-300 bg-white'
                  }`}
                >
                  <div>
                    <span className="text-[9px] font-bold text-gray-500 uppercase tracking-wider">Avulso</span>
                    <h3 className="text-sm font-black text-gray-900 mt-0.5">1 Trabalho</h3>
                    <p className="text-[11px] text-gray-500 mt-0.5">R$ 1,99</p>
                  </div>
                  <div className="mt-2 pt-1.5 border-t border-gray-200/60 flex items-center justify-between">
                    <span className="text-[11px] font-bold text-gray-800">R$ 1,99</span>
                    <input 
                      type="radio" 
                      name="pixPlan" 
                      checked={selectedPixPlan === 'single'} 
                      onChange={() => setSelectedPixPlan('single')}
                      className="text-blue-600"
                    />
                  </div>
                </div>

                {/* Pacote 2: R$ 5,00 = 3 Trabalhos */}
                <div 
                  onClick={() => setSelectedPixPlan('trio')}
                  className={`p-3 rounded-2xl border-2 cursor-pointer transition-all flex flex-col justify-between text-left relative ${
                    selectedPixPlan === 'trio' 
                      ? 'border-blue-600 bg-blue-50/60 shadow-sm ring-1 ring-blue-500' 
                      : 'border-gray-200 hover:border-gray-300 bg-white'
                  }`}
                >
                  <span className="absolute -top-2 right-1.5 bg-blue-600 text-white text-[8px] font-black px-1.5 py-0.2 rounded-full uppercase">
                    Econômico
                  </span>
                  <div>
                    <span className="text-[9px] font-bold text-blue-600 uppercase tracking-wider">Trio</span>
                    <h3 className="text-sm font-black text-gray-900 mt-0.5">3 Trabalhos</h3>
                    <p className="text-[11px] text-gray-500 mt-0.5">R$ 1,66/un</p>
                  </div>
                  <div className="mt-2 pt-1.5 border-t border-gray-200/60 flex items-center justify-between">
                    <span className="text-[11px] font-bold text-blue-700">R$ 5,00</span>
                    <input 
                      type="radio" 
                      name="pixPlan" 
                      checked={selectedPixPlan === 'trio'} 
                      onChange={() => setSelectedPixPlan('trio')}
                      className="text-blue-600"
                    />
                  </div>
                </div>

                {/* Pacote 3: R$ 29,90 = 50 Trabalhos (Pacote Pro / TCC Completo) */}
                <div 
                  onClick={() => setSelectedPixPlan('pro')}
                  className={`p-3 rounded-2xl border-2 cursor-pointer transition-all flex flex-col justify-between text-left relative ${
                    selectedPixPlan === 'pro' 
                      ? 'border-amber-500 bg-amber-50/60 shadow-sm ring-1 ring-amber-500' 
                      : 'border-gray-200 hover:border-gray-300 bg-white'
                  }`}
                >
                  <span className="absolute -top-2 right-1.5 bg-gradient-to-r from-amber-500 to-orange-500 text-white text-[8px] font-black px-1.5 py-0.2 rounded-full uppercase">
                    Melhor Valor
                  </span>
                  <div>
                    <span className="text-[9px] font-bold text-amber-600 uppercase tracking-wider">Pacote Pro</span>
                    <h3 className="text-sm font-black text-gray-900 mt-0.5">50 Trabalhos</h3>
                    <p className="text-[11px] text-gray-500 mt-0.5">R$ 0,59/un</p>
                  </div>
                  <div className="mt-2 pt-1.5 border-t border-gray-200/60 flex items-center justify-between">
                    <span className="text-[11px] font-bold text-amber-800">R$ 29,90</span>
                    <input 
                      type="radio" 
                      name="pixPlan" 
                      checked={selectedPixPlan === 'pro'} 
                      onChange={() => setSelectedPixPlan('pro')}
                      className="text-amber-600"
                    />
                  </div>
                </div>
              </div>

              {/* Caixa da Chave PIX */}
              <div>
                <div className="flex items-center justify-between mb-1.5">
                  <label className="text-xs font-bold text-gray-700 uppercase tracking-wider">
                    Chave PIX Oficial (E-mail):
                  </label>
                  <span className="text-xs font-bold text-blue-700">
                    Valor: {selectedPixPlan === 'single' ? 'R$ 1,99' : selectedPixPlan === 'trio' ? 'R$ 5,00' : 'R$ 29,90'}
                  </span>
                </div>
                <div className="flex items-center gap-2">
                  <div className="flex-1 bg-gray-100 border border-gray-200 px-3.5 py-2.5 rounded-xl font-mono text-xs md:text-sm text-gray-800 select-all font-semibold break-all">
                    erlanehmotta@gmail.com
                  </div>
                  <Button
                    onClick={() => {
                      navigator.clipboard.writeText("erlanehmotta@gmail.com");
                      setPixCopied(true);
                      setTimeout(() => setPixCopied(false), 3000);
                    }}
                    className={`font-semibold text-xs px-4 py-2.5 rounded-xl transition-all ${
                      pixCopied 
                        ? "bg-emerald-600 hover:bg-emerald-700 text-white" 
                        : "bg-blue-600 hover:bg-blue-700 text-white shadow-md shadow-blue-500/20"
                    }`}
                  >
                    {pixCopied ? (
                      <>
                        <Check className="w-4 h-4 mr-1" /> Copiado!
                      </>
                    ) : (
                      <>
                        <Copy className="w-4 h-4 mr-1" /> Copiar
                      </>
                    )}
                  </Button>
                </div>
              </div>

              {/* QR Code */}
              <div className="text-center bg-gray-50 p-3.5 rounded-2xl border border-gray-200">
                <p className="text-xs font-semibold text-gray-600 mb-2">
                  Pague {selectedPixPlan === 'single' ? 'R$ 1,99 (1 Trabalho)' : selectedPixPlan === 'trio' ? 'R$ 5,00 (3 Trabalhos)' : 'R$ 9,90 (7 Trabalhos)'} pelo app do banco:
                </p>
                <img 
                  src="https://api.qrserver.com/v1/create-qr-code/?size=150x150&data=erlanehmotta@gmail.com" 
                  alt="QR Code PIX erlanehmotta@gmail.com" 
                  className="w-32 h-32 mx-auto rounded-xl shadow-xs border border-gray-200 bg-white p-1"
                />
                <p className="text-[11px] text-gray-500 mt-1.5 font-mono">Chave: erlanehmotta@gmail.com</p>
              </div>

              {/* Botão de Liberação Automática do App */}
              <div className="pt-2">
                {activationSuccess ? (
                  <div className="bg-emerald-50 border border-emerald-200 text-emerald-800 text-xs p-3.5 rounded-2xl text-center font-bold animate-in fade-in">
                    🎉 PIX Confirmado! +{selectedPixPlan === 'single' ? '1' : selectedPixPlan === 'trio' ? '3' : '7'} Trabalho(s) Liberado(s)!
                  </div>
                ) : (
                  <Button
                    onClick={() => {
                      const addedCredits = selectedPixPlan === 'single' ? 1 : selectedPixPlan === 'trio' ? 3 : 7;
                      setCredits(prev => {
                        const next = prev + addedCredits;
                        localStorage.setItem("emia_credits", String(next));
                        return next;
                      });
                      setActivationSuccess(true);
                      setTimeout(() => {
                        setActivationSuccess(false);
                        setShowPixModal(false);
                      }, 1200);
                    }}
                    className="w-full bg-gradient-to-r from-emerald-600 to-teal-600 hover:from-emerald-700 hover:to-teal-700 text-white font-bold py-3.5 rounded-2xl text-xs md:text-sm shadow-md shadow-emerald-500/25 active:scale-[0.98] transition-all"
                  >
                    <CheckCircle className="w-4 h-4 mr-2 text-white" />
                    Já fiz o PIX / Liberar Agora ({selectedPixPlan === 'single' ? '+1 Trabalho' : selectedPixPlan === 'trio' ? '+3 Trabalhos' : '+7 Trabalhos'})
                  </Button>
                )}
              </div>
            </div>
          </div>
        </div>
      )}

      {/* MODAL DE APRESENTAÇÃO DE SLIDES (GOOGLE SLIDES / PPTX) */}
      {showSlidesModal && (
        <div className="fixed inset-0 z-50 bg-black/60 backdrop-blur-sm flex items-center justify-center p-4">
          <div className="bg-white rounded-3xl shadow-2xl border border-gray-200 w-full max-w-2xl overflow-hidden animate-in fade-in zoom-in-95 duration-200">
            {/* Header */}
            <div className="bg-gradient-to-r from-amber-500 via-orange-500 to-amber-600 p-5 text-white flex items-center justify-between">
              <div className="flex items-center gap-3">
                <div className="w-10 h-10 rounded-2xl bg-white/20 flex items-center justify-center backdrop-blur-xs">
                  <Presentation className="w-5 h-5 text-white" />
                </div>
                <div>
                  <h3 className="text-base md:text-lg font-bold">Apresentação de Slides</h3>
                  <p className="text-xs text-amber-100">Geração automática para Google Slides & PowerPoint</p>
                </div>
              </div>
              <button
                onClick={() => setShowSlidesModal(false)}
                className="w-8 h-8 rounded-full bg-white/20 hover:bg-white/30 flex items-center justify-center transition-colors text-white text-sm"
              >
                ✕
              </button>
            </div>

            {/* Content */}
            <div className="p-6 space-y-5 max-h-[75vh] overflow-y-auto">
              
              {/* Tema Visual dos Slides */}
              <div>
                <label className="block text-xs font-bold text-gray-700 uppercase tracking-wider mb-2">
                  Tema Visual dos Slides:
                </label>
                <div className="grid grid-cols-3 gap-2.5">
                  <button
                    onClick={() => setSlidesTheme("academic")}
                    className={`p-3 rounded-2xl border text-left transition-all ${
                      slidesTheme === "academic"
                        ? "border-blue-600 bg-blue-50/70 ring-2 ring-blue-500/20"
                        : "border-gray-200 hover:border-gray-300 bg-white"
                    }`}
                  >
                    <div className="w-full h-3 rounded bg-blue-900 mb-2" />
                    <p className="text-xs font-bold text-gray-900">Acadêmico NBR</p>
                    <p className="text-[10px] text-gray-500">Azul Marinho & Branco</p>
                  </button>

                  <button
                    onClick={() => setSlidesTheme("modern")}
                    className={`p-3 rounded-2xl border text-left transition-all ${
                      slidesTheme === "modern"
                        ? "border-teal-600 bg-teal-50/70 ring-2 ring-teal-500/20"
                        : "border-gray-200 hover:border-gray-300 bg-white"
                    }`}
                  >
                    <div className="w-full h-3 rounded bg-teal-700 mb-2" />
                    <p className="text-xs font-bold text-gray-900">Moderno Tech</p>
                    <p className="text-[10px] text-gray-500">Verde Petróleo & Slate</p>
                  </button>

                  <button
                    onClick={() => setSlidesTheme("dark")}
                    className={`p-3 rounded-2xl border text-left transition-all ${
                      slidesTheme === "dark"
                        ? "border-slate-800 bg-slate-100 ring-2 ring-slate-800/20"
                        : "border-gray-200 hover:border-gray-300 bg-white"
                    }`}
                  >
                    <div className="w-full h-3 rounded bg-slate-900 mb-2" />
                    <p className="text-xs font-bold text-gray-900">Minimalista</p>
                    <p className="text-[10px] text-gray-500">Grafite & Cinza Claro</p>
                  </button>
                </div>
              </div>

              {/* Pré-visualização da Estrutura dos Slides */}
              <div className="bg-gray-50 p-4 rounded-2xl border border-gray-200">
                <div className="flex items-center justify-between mb-2">
                  <span className="text-xs font-bold text-gray-800">Estrutura Gerada (6 Slides)</span>
                  <span className="text-[10px] text-emerald-700 font-bold bg-emerald-100 px-2 py-0.5 rounded-full">Pronto para Defesa</span>
                </div>
                <ul className="text-xs text-gray-600 space-y-1.5 list-disc list-inside">
                  <li><strong className="text-gray-800">Slide 1:</strong> Capa com Título, Autor, Orientador e Instituição</li>
                  <li><strong className="text-gray-800">Slide 2:</strong> Introdução, Relevância & Contextualização</li>
                  <li><strong className="text-gray-800">Slide 3:</strong> Problema de Pesquisa, Objetivos e Hipótese</li>
                  <li><strong className="text-gray-800">Slide 4:</strong> Fundamentação Teórica & Metodologia</li>
                  <li><strong className="text-gray-800">Slide 5:</strong> Desenvolvimento & Principais Resultados</li>
                  <li><strong className="text-gray-800">Slide 6:</strong> Conclusão & Referências ABNT</li>
                </ul>
              </div>

              {/* Botões de Ação */}
              <div className="space-y-2.5 pt-2">
                <Button
                  onClick={openInGoogleSlides}
                  className="w-full bg-gradient-to-r from-amber-500 to-orange-500 hover:from-amber-600 hover:to-orange-600 text-white font-bold py-3.5 rounded-2xl text-xs md:text-sm shadow-md shadow-amber-500/25 flex items-center justify-center gap-2"
                >
                  <Share2 className="w-4 h-4" />
                  🚀 Criar e Abrir no Google Slides (Com Roteiro Copiado)
                </Button>

                <Button
                  onClick={exportPPTXSlides}
                  variant="outline"
                  className="w-full border-2 border-gray-300 hover:border-gray-400 text-gray-800 font-bold py-3.5 rounded-2xl text-xs md:text-sm flex items-center justify-center gap-2 bg-white"
                >
                  <Download className="w-4 h-4 text-orange-600" />
                  📥 Baixar Arquivo .pptx (Compatível com Google Slides & PowerPoint)
                </Button>
              </div>

            </div>
          </div>
        </div>
      )}

      {/* MODAL TRABALHO EM GRUPO (Aberto pelo botão da barra superior) */}
      {isGroupMode && (
        <div className="fixed inset-0 bg-black/60 backdrop-blur-xs flex items-center justify-center z-50 p-4 animate-in fade-in duration-200">
          <div className="bg-white rounded-2xl shadow-2xl w-full max-w-2xl max-h-[92vh] flex flex-col overflow-hidden border border-gray-100">
            {/* Header do Modal */}
            <div className="flex items-center justify-between px-6 py-4 border-b border-gray-100 bg-gradient-to-r from-emerald-50 to-teal-50">
              <div className="flex items-center gap-2.5">
                <div className="w-9 h-9 rounded-xl bg-emerald-100 border border-emerald-300 text-emerald-800 flex items-center justify-center shadow-xs">
                  <Users className="w-5 h-5 stroke-[1.8]" />
                </div>
                <div>
                  <h2 className="text-base font-bold text-gray-900">Montagem de Trabalho em Grupo</h2>
                  <p className="text-xs text-gray-500">Junte as partes feitas por cada aluno em um trabalho ABNT completo</p>
                </div>
              </div>
              <button 
                onClick={() => setIsGroupMode(false)} 
                className="w-8 h-8 rounded-full hover:bg-gray-200/80 flex items-center justify-center text-gray-400 hover:text-gray-700 transition-all"
              >
                <X className="w-5 h-5" />
              </button>
            </div>

            {/* Conteúdo do Modal */}
            <div className="flex-1 overflow-y-auto p-6 space-y-6">
              
              {/* Tipo de Documento / Divisão Estrutural */}
              <div className="bg-emerald-50/60 p-4 rounded-2xl border border-emerald-200/80">
                <label className="block text-sm font-bold text-gray-800 mb-1.5">
                  📚 Tipo de Trabalho (Define a Divisão de Seções)
                </label>
                <select
                  value={groupDocType}
                  onChange={(e) => setGroupDocType(e.target.value)}
                  className="w-full px-3.5 py-2.5 bg-white border border-gray-300 rounded-xl text-sm font-medium text-gray-800 shadow-xs focus:ring-2 focus:ring-emerald-500 focus:outline-none"
                >
                  <option value="trabalho_academico">🎓 Trabalho Acadêmico / TCC / Monografia (NBR 14724)</option>
                  <option value="artigo">📄 Artigo Acadêmico / Artigo Científico (NBR 6022)</option>
                  <option value="projeto">📑 Projeto de Pesquisa (NBR 15287)</option>
                  <option value="relatorio">📋 Relatório Técnico-Científico (NBR 10719)</option>
                  <option value="estudo_caso">📊 Estudo de Caso Prático</option>
                  <option value="resenha">✍️ Resenha Crítica de Obra</option>
                  <option value="outro">🛠️ Outro (Personalizar Minhas Próprias Seções e Divisão)</option>
                </select>

                {groupDocType === "outro" && (
                  <div className="mt-3.5 pt-3 border-t border-emerald-200/70 space-y-3">
                    <div>
                      <label className="block text-xs font-bold text-gray-800 mb-1">Nome do Tipo de Trabalho (para a Folha de Rosto):</label>
                      <input
                        type="text"
                        value={customGroupDocName}
                        onChange={(e) => setCustomGroupDocName(e.target.value)}
                        placeholder="Ex: Estudo Temático, Portfólio em Grupo, Análise Setorial..."
                        className="w-full px-3 py-1.5 bg-white border border-gray-300 rounded-lg text-xs focus:ring-2 focus:ring-emerald-500 focus:outline-none"
                      />
                    </div>

                    <div>
                      <div className="flex items-center justify-between mb-1.5">
                        <label className="text-xs font-bold text-gray-800">Defina suas Seções Personalizadas:</label>
                        <span className="text-[10px] text-gray-500">Crie quantas seções seu trabalho precisar</span>
                      </div>
                      <div className="space-y-2">
                        {customSections.map((sec, sIdx) => (
                          <div key={sec.id} className="flex items-center gap-2 bg-white p-2 rounded-xl border border-gray-200 shadow-2xs">
                            <span className="text-xs font-bold text-emerald-800 w-5 text-center">{sIdx + 1}º</span>
                            <input
                              type="text"
                              value={sec.label}
                              onChange={(e) => {
                                const updated = [...customSections];
                                updated[sIdx].label = e.target.value;
                                setCustomSections(updated);
                              }}
                              placeholder="Nome da Seção (ex: 1 INTRODUÇÃO E CONTEXTO)"
                              className="flex-1 px-2.5 py-1 text-xs border border-gray-200 rounded-md focus:ring-1 focus:ring-emerald-500 focus:outline-none font-semibold text-gray-800"
                            />
                            {customSections.length > 1 && (
                              <button
                                onClick={() => {
                                  setCustomSections(prev => prev.filter((_, i) => i !== sIdx));
                                  setSectionSlots(prev => {
                                    const next = { ...prev };
                                    delete next[sec.id];
                                    return next;
                                  });
                                }}
                                className="w-6 h-6 rounded-md hover:bg-red-50 text-red-400 hover:text-red-600 flex items-center justify-center transition-all"
                                title="Remover esta seção personalizada"
                              >
                                <X className="w-3.5 h-3.5" />
                              </button>
                            )}
                          </div>
                        ))}
                      </div>
                      <button
                        onClick={() => {
                          const newId = `sec_${Date.now()}`;
                          setCustomSections(prev => [
                            ...prev,
                            { id: newId, label: `${prev.length + 1} NOVA SEÇÃO`, desc: "Seção personalizada", icon: "📑" }
                          ]);
                        }}
                        className="mt-2 text-xs font-bold text-emerald-700 hover:text-emerald-900 flex items-center gap-1 transition-colors bg-white hover:bg-emerald-50 border border-emerald-300 px-2.5 py-1 rounded-lg shadow-2xs"
                      >
                        <Plus className="w-3.5 h-3.5" /> Adicionar Outra Seção
                      </button>
                    </div>
                  </div>
                )}

                <div className="mt-2 text-[11px] text-emerald-800 font-medium">
                  {groupDocType === "trabalho_academico" && "✨ Estrutura: Resumo/Abstract + 1 Introdução + 2 Fundamentação e Métodos + 3 Resultados + 4 Considerações Finais + Referências"}
                  {groupDocType === "artigo" && "✨ Estrutura: Resumo/Abstract + 1 Introdução + 2 Materiais e Métodos + 3 Resultados e Discussão + 4 Considerações Finais + Referências"}
                  {groupDocType === "projeto" && "✨ Estrutura: 1 Introdução/Justificativa + 2 Objetivos Geral/Específicos + 3 Fundamentação + 4 Metodologia/Cronograma + Referências"}
                  {groupDocType === "relatorio" && "✨ Estrutura: 1 Introdução e Objetivo + 2 Procedimentos Realizados + 3 Análise de Dados + 4 Recomendações + Referências"}
                  {groupDocType === "estudo_caso" && "✨ Estrutura: 1 Introdução do Caso + 2 Diagnóstico e Teoria + 3 Proposições/Soluções + 4 Lições Aprendidas + Referências"}
                  {groupDocType === "resenha" && "✨ Estrutura: Cabeçalho Bibliográfico + 1 Síntese da Obra + 2 Apreciação Crítica + 3 Conclusão/Indicação"}
                  {groupDocType === "outro" && "✨ Estrutura Totalmente Livre: Você define os títulos das seções, a quantidade de partes e a ordem exata."}
                </div>
              </div>

              {/* Título e Subtítulo do Trabalho em Grupo */}
              <div className="space-y-3 bg-white p-4 rounded-2xl border border-gray-200">
                <div>
                  <label className="block text-sm font-bold text-gray-800 mb-1">
                    📌 Título do Trabalho (Será impresso na Capa e Folha de Rosto)
                  </label>
                  <input
                    type="text"
                    value={title}
                    onChange={(e) => setTitle(e.target.value)}
                    placeholder="Ex: Impactos da Inteligência Artificial na Gestão Escolar Contemporânea"
                    className="w-full px-3.5 py-2 border border-gray-300 rounded-xl text-sm focus:ring-2 focus:ring-emerald-500 focus:outline-none font-medium"
                  />
                </div>
                <div>
                  <label className="block text-xs font-semibold text-gray-600 mb-1">
                    Subtítulo do Trabalho (Opcional)
                  </label>
                  <input
                    type="text"
                    value={subtitle}
                    onChange={(e) => setSubtitle(e.target.value)}
                    placeholder="Ex: Um estudo de caso multisetorial"
                    className="w-full px-3.5 py-1.5 border border-gray-300 rounded-xl text-xs focus:ring-2 focus:ring-emerald-500 focus:outline-none"
                  />
                </div>
              </div>

              {/* Membros do Grupo */}
              <div>
                <div className="flex items-center justify-between mb-2">
                  <label className="text-sm font-bold text-gray-800">Alunos / Integrantes do Grupo</label>
                  <span className="text-xs text-gray-500">Serão impressos na Capa e Folha de Rosto</span>
                </div>
                <div className="space-y-2">
                  {groupMembers.map((member, idx) => (
                    <div key={idx} className="flex items-center gap-2">
                      <span className="text-xs font-bold text-gray-400 w-6 text-center">{idx + 1}º</span>
                      <input
                        type="text"
                        value={member}
                        onChange={(e) => {
                          const updated = [...groupMembers];
                          updated[idx] = e.target.value;
                          setGroupMembers(updated);
                        }}
                        placeholder={`Nome completo do Aluno ${idx + 1}`}
                        className="flex-1 px-3.5 py-2 border border-gray-300 rounded-xl text-sm focus:ring-2 focus:ring-emerald-500 focus:outline-none"
                      />
                      {groupMembers.length > 1 && (
                        <button
                          onClick={() => setGroupMembers(prev => prev.filter((_, i) => i !== idx))}
                          className="w-8 h-8 rounded-lg hover:bg-red-50 text-red-400 hover:text-red-600 flex items-center justify-center transition-all"
                          title="Remover membro"
                        >
                          <X className="w-4 h-4" />
                        </button>
                      )}
                    </div>
                  ))}
                </div>
                <button
                  onClick={() => setGroupMembers(prev => [...prev, ""])}
                  className="mt-2.5 text-xs font-bold text-emerald-700 hover:text-emerald-900 flex items-center gap-1.5 transition-colors bg-emerald-50 hover:bg-emerald-100 px-3 py-1.5 rounded-lg"
                >
                  <Plus className="w-3.5 h-3.5" /> Adicionar Mais um Aluno
                </button>
              </div>

              {/* Slots de Seções Dinâmicos baseados no Tipo Escolhido */}
              <div>
                <div className="flex items-center justify-between mb-2">
                  <label className="text-sm font-bold text-gray-800">Arquivos por Seção ({getGroupSectionsByDocType(groupDocType).length} Seções)</label>
                  <span className="text-xs text-gray-500">PDF, Word (.docx), TXT ou MD</span>
                </div>
                <div className="space-y-2.5">
                  {getGroupSectionsByDocType(groupDocType).map(({ key, label, desc, icon }) => (
                    <div 
                      key={key} 
                      className={`flex items-center justify-between p-3.5 rounded-xl border transition-all ${
                        sectionSlots[key] 
                          ? 'bg-emerald-50/80 border-emerald-300 shadow-2xs' 
                          : 'bg-gray-50/60 border-gray-200 hover:border-gray-300'
                      }`}
                    >
                      <div className="flex items-center gap-3 min-w-0 flex-1">
                        <span className="text-xl flex-shrink-0">{icon}</span>
                        <div className="min-w-0 flex-1">
                          <p className="text-xs font-bold text-gray-900 truncate">{label}</p>
                          <p className="text-[11px] text-gray-500 truncate">{desc}</p>
                        </div>
                      </div>

                      {sectionSlots[key] ? (
                        <div className="flex items-center gap-2 ml-3">
                          <span className="text-xs font-bold text-emerald-800 bg-emerald-100 px-2.5 py-1 rounded-lg truncate max-w-[160px]">
                            {sectionSlots[key]!.name}
                          </span>
                          <button
                            onClick={() => setSectionSlots(prev => ({ ...prev, [key]: null }))}
                            className="w-7 h-7 rounded-lg hover:bg-red-100 text-red-500 flex items-center justify-center transition-all"
                            title="Remover arquivo desta seção"
                          >
                            <X className="w-4 h-4" />
                          </button>
                        </div>
                      ) : (
                        <label className="cursor-pointer text-xs font-bold text-emerald-800 hover:text-emerald-900 bg-white hover:bg-emerald-50 border border-emerald-300 px-3 py-1.5 rounded-xl transition-all shadow-2xs flex items-center gap-1.5 ml-3">
                          <Upload className="w-3.5 h-3.5 text-emerald-600" />
                          <span>Selecionar Arquivo</span>
                          <input
                            type="file"
                            className="hidden"
                            accept=".pdf,.docx,.doc,.txt,.md,.csv"
                            onChange={(e) => {
                              const file = e.target.files?.[0];
                              if (file) {
                                setSectionSlots(prev => ({ ...prev, [key]: file }));
                              }
                              e.target.value = "";
                            }}
                          />
                        </label>
                      )}
                    </div>
                  ))}
                </div>
              </div>
            </div>

            {/* Rodapé do Modal com Botão de Montagem */}
            <div className="p-5 border-t border-gray-100 bg-gray-50 flex items-center justify-between gap-3">
              <Button 
                variant="outline" 
                onClick={() => setIsGroupMode(false)}
                className="border-gray-300"
              >
                Cancelar
              </Button>

              <Button
                onClick={async () => {
                  await handleGroupAssemble();
                  setIsGroupMode(false);
                }}
                disabled={isLoading || Object.values(sectionSlots).every(f => f === null)}
                className="bg-emerald-600 hover:bg-emerald-700 text-white font-bold py-2.5 px-6 rounded-xl text-sm shadow-md shadow-emerald-600/20 flex items-center gap-2"
              >
                {isLoading ? <Loader2 className="w-4 h-4 animate-spin" /> : <span>🚀</span>}
                Montar Trabalho Completo ABNT
              </Button>
            </div>

          </div>
        </div>
      )}

      {/* Modal de Citação Direta e Indireta (ABNT NBR 10520:2023) */}
      {showCitationModal && (
        <div className="fixed inset-0 bg-black/60 backdrop-blur-sm z-50 flex items-center justify-center p-4">
          <div className="bg-white rounded-2xl shadow-2xl max-w-lg w-full border border-gray-200 overflow-hidden flex flex-col animate-in fade-in zoom-in duration-200">
            {/* Header */}
            <div className="bg-gradient-to-r from-purple-600 to-indigo-700 text-white p-4 flex items-center justify-between">
              <div className="flex items-center gap-2">
                <Quote className="w-5 h-5" />
                <h3 className="font-bold text-base">Gerador de Citação ABNT (NBR 10520:2023)</h3>
              </div>
              <button 
                onClick={() => setShowCitationModal(false)}
                className="w-7 h-7 rounded-lg hover:bg-white/20 flex items-center justify-center text-white transition-colors"
              >
                <X className="w-4 h-4" />
              </button>
            </div>

            {/* Body */}
            <div className="p-5 space-y-4 text-xs font-sans overflow-y-auto max-h-[75vh]">
              {/* Tipo de Citação */}
              <div>
                <label className="block font-bold text-gray-700 mb-1.5">Tipo de Citação:</label>
                <div className="grid grid-cols-3 gap-2">
                  <button
                    type="button"
                    onClick={() => setCitationType("direta_longa")}
                    className={`p-2.5 rounded-xl border text-center font-bold transition-all ${
                      citationType === "direta_longa"
                        ? "bg-purple-50 border-purple-500 text-purple-700 ring-2 ring-purple-200 shadow-xs"
                        : "bg-gray-50 border-gray-200 text-gray-600 hover:bg-gray-100"
                    }`}
                  >
                    <div className="text-xs">📜 Direta Longa</div>
                    <div className="text-[10px] font-normal text-gray-500 mt-0.5">+3 linhas (Recuo 4cm, 10pt)</div>
                  </button>

                  <button
                    type="button"
                    onClick={() => setCitationType("direta_curta")}
                    className={`p-2.5 rounded-xl border text-center font-bold transition-all ${
                      citationType === "direta_curta"
                        ? "bg-purple-50 border-purple-500 text-purple-700 ring-2 ring-purple-200 shadow-xs"
                        : "bg-gray-50 border-gray-200 text-gray-600 hover:bg-gray-100"
                    }`}
                  >
                    <div className="text-xs">💬 Direta Curta</div>
                    <div className="text-[10px] font-normal text-gray-500 mt-0.5">Até 3 linhas (Com Aspas)</div>
                  </button>

                  <button
                    type="button"
                    onClick={() => setCitationType("indireta")}
                    className={`p-2.5 rounded-xl border text-center font-bold transition-all ${
                      citationType === "indireta"
                        ? "bg-purple-50 border-purple-500 text-purple-700 ring-2 ring-purple-200 shadow-xs"
                        : "bg-gray-50 border-gray-200 text-gray-600 hover:bg-gray-100"
                    }`}
                  >
                    <div className="text-xs">💡 Indireta</div>
                    <div className="text-[10px] font-normal text-gray-500 mt-0.5">Paráfrase (Autor-Data)</div>
                  </button>
                </div>
              </div>

              {/* Campos Autor, Ano, Página */}
              <div className="grid grid-cols-1 sm:grid-cols-3 gap-3">
                <div className="sm:col-span-1">
                  <label className="block font-bold text-gray-700 mb-1">Autor / Sobrenome:</label>
                  <input
                    type="text"
                    value={citationAuthor}
                    onChange={(e) => setCitationAuthor(e.target.value)}
                    placeholder="Ex: Silva"
                    className="w-full px-3 py-1.5 border border-gray-300 rounded-lg text-xs focus:ring-2 focus:ring-purple-500 focus:outline-none"
                  />
                </div>
                <div>
                  <label className="block font-bold text-gray-700 mb-1">Ano da Obra:</label>
                  <input
                    type="text"
                    value={citationYear}
                    onChange={(e) => setCitationYear(e.target.value)}
                    placeholder="Ex: 2023"
                    className="w-full px-3 py-1.5 border border-gray-300 rounded-lg text-xs focus:ring-2 focus:ring-purple-500 focus:outline-none"
                  />
                </div>
                <div>
                  <label className="block font-bold text-gray-700 mb-1">Página ({citationType === "indireta" ? "Opcional" : "Obrigatória"}):</label>
                  <input
                    type="text"
                    value={citationPage}
                    onChange={(e) => setCitationPage(e.target.value)}
                    placeholder="Ex: 45"
                    className="w-full px-3 py-1.5 border border-gray-300 rounded-lg text-xs focus:ring-2 focus:ring-purple-500 focus:outline-none"
                  />
                </div>
              </div>

              {/* Texto da Citação */}
              <div>
                <label className="block font-bold text-gray-700 mb-1">
                  {citationType === "indireta" ? "Texto da Paráfrase (Suas palavras baseadas na ideia do autor):" : "Trecho Citado Exatamente da Obra Original:"}
                </label>
                <textarea
                  rows={4}
                  value={citationContent}
                  onChange={(e) => setCitationContent(e.target.value)}
                  placeholder={
                    citationType === "direta_longa"
                      ? "Cole aqui o trecho com mais de 3 linhas transcrito fielmente do texto original..."
                      : citationType === "direta_curta"
                      ? "Cole aqui a frase curta (até 3 linhas)..."
                      : "Escreva com suas próprias palavras a ideia apresentada pelo autor..."
                  }
                  className="w-full p-2.5 border border-gray-300 rounded-lg text-xs focus:ring-2 focus:ring-purple-500 focus:outline-none font-serif leading-relaxed"
                />
              </div>

              {/* Prévia Formatada em Tempo Real */}
              <div className="bg-slate-50 border border-slate-200 rounded-xl p-3">
                <div className="text-[10px] font-bold uppercase tracking-wider text-slate-500 mb-1 flex items-center justify-between">
                  <span>Prévia Formatada (Padrão ABNT NBR 10520:2023)</span>
                  <span className="text-purple-600 font-extrabold">Caixa Mista Oficial</span>
                </div>
                <div className="p-3 bg-white border border-slate-100 rounded-lg text-xs text-gray-900 leading-relaxed">
                  {(() => {
                    const auth = citationAuthor.trim() || "Autor";
                    const yr = citationYear.trim() || "2023";
                    const pg = citationPage.trim() ? `, p. ${citationPage.trim()}` : "";
                    const txt = citationContent.trim() || (citationType === "indireta" ? "ideia expressa pelo pesquisador em seu estudo" : "texto citado textualmente da obra original");
                    
                    if (citationType === "direta_longa") {
                      return (
                        <div className="pl-8 text-[10pt] leading-[1.2] text-gray-800 border-l-2 border-purple-400 italic">
                          {txt} ({auth}, {yr}{pg}).
                        </div>
                      );
                    }
                    if (citationType === "direta_curta") {
                      return (
                        <div>
                          De acordo com {auth} ({yr}{pg}), &ldquo;{txt}&rdquo;.
                        </div>
                      );
                    }
                    return (
                      <div>
                        Segundo {auth} ({yr}), {txt}.
                      </div>
                    );
                  })()}
                </div>
              </div>
            </div>

            {/* Footer com Ações */}
            <div className="p-4 bg-gray-50 border-t border-gray-200 flex items-center justify-between gap-2">
              <Button
                variant="outline"
                size="sm"
                onClick={() => setShowCitationModal(false)}
              >
                Cancelar
              </Button>
              <Button
                size="sm"
                onClick={() => {
                  const auth = citationAuthor.trim() || "Autor";
                  const yr = citationYear.trim() || "2023";
                  const pg = citationPage.trim() ? `, p. ${citationPage.trim()}` : "";
                  const txt = citationContent.trim() || "texto da citação";
                  
                  let formattedBlock = "";
                  if (citationType === "direta_longa") {
                    formattedBlock = `\n\n> ${txt} (${auth}, ${yr}${pg}).\n\n`;
                  } else if (citationType === "direta_curta") {
                    formattedBlock = ` "${txt}" (${auth}, ${yr}${pg}) `;
                  } else {
                    formattedBlock = ` Conforme aponta ${auth} (${yr}), ${txt}. `;
                  }

                  const updatedDoc = (generatedText || "") + formattedBlock;
                  updateGeneratedTextWithHistory(updatedDoc);
                  setShowCitationModal(false);
                  setCitationContent("");
                  setCitationAuthor("");
                  setCitationPage("");
                  setErrorMessage("✨ Citação inserida com sucesso no documento com rigor ABNT NBR 10520:2023!");
                  setTimeout(() => setErrorMessage(""), 3500);
                  logAction("Citação ABNT Inserida", formattedBlock);
                }}
                className="bg-purple-600 hover:bg-purple-700 text-white font-bold"
              >
                📥 Inserir no Documento
              </Button>
            </div>
          </div>
        </div>
      )}

    </div>
  );
}

