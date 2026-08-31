/**
 * Cloudflare Worker for EDUTECH.EMIA
 * Handles API requests and serves static assets from Vite dist/
 */

export default {
  async fetch(request, env, ctx) {
    const url = new URL(request.url);

    // Standard CORS Headers
    const corsHeaders = {
      "Access-Control-Allow-Origin": "*",
      "Access-Control-Allow-Methods": "GET, HEAD, POST, OPTIONS",
      "Access-Control-Allow-Headers": "Content-Type, Authorization, x-gemini-api-key, x-google-api-key, x-ai-provider",
    };

    // Handle preflight CORS requests
    if (request.method === "OPTIONS") {
      return new Response(null, { headers: corsHeaders });
    }

    // Health check endpoint
    if (url.pathname === "/api/health") {
      return new Response(JSON.stringify({ status: "ok", app: "emia-edutech", timestamp: new Date().toISOString() }), {
        headers: { "Content-Type": "application/json", ...corsHeaders }
      });
    }

    // List available Gemini models for the configured secret
    if (url.pathname === "/api/models") {
      const apiKey = env.GEMINI_API_KEY || env.GOOGLE_API_KEY;
      if (!apiKey) {
        return new Response(JSON.stringify({ error: "No API key configured" }), { headers: { "Content-Type": "application/json", ...corsHeaders } });
      }
      try {
        const res = await fetch(`https://generativelanguage.googleapis.com/v1beta/models?key=${apiKey}`);
        const data = await res.json();
        return new Response(JSON.stringify(data), { headers: { "Content-Type": "application/json", ...corsHeaders } });
      } catch (err) {
        return new Response(JSON.stringify({ error: err.message }), { status: 500, headers: { "Content-Type": "application/json", ...corsHeaders } });
      }
    }

    // Endpoint: Verificar Créditos do Aluno por E-mail (Hotmart / Local)
    if (url.pathname === "/api/user-credits" && request.method === "POST") {
      try {
        const body = await request.json();
        const email = (body.email || "").trim().toLowerCase();

        if (!email) {
          return new Response(JSON.stringify({ credits: 0 }), {
            headers: { "Content-Type": "application/json", ...corsHeaders }
          });
        }

        // Admin Geral & Acessos Vitalícios
        if (email === "erlane.digital@gmail.com" || email === "cadumajor@gmail.com") {
          return new Response(JSON.stringify({ credits: 9999, isMaster: true }), {
            headers: { "Content-Type": "application/json", ...corsHeaders }
          });
        }

        let credits = 3; // Créditos padrão de cortesia / início
        if (env.EMIA_KV) {
          const stored = await env.EMIA_KV.get(`credits_${email}`);
          if (stored !== null) {
            credits = parseInt(stored, 10);
          }
        }

        return new Response(JSON.stringify({ credits, isMaster: false }), {
          headers: { "Content-Type": "application/json", ...corsHeaders }
        });
      } catch (err) {
        return new Response(JSON.stringify({ credits: 3 }), {
          headers: { "Content-Type": "application/json", ...corsHeaders }
        });
      }
    }

    // Endpoint: Webhook Universal Multi-Plataformas (Hotmart, Kiwify, Eduzz, Monetizze, Cakto)
    if ((url.pathname === "/api/hotmart-webhook" || url.pathname === "/api/webhook") && request.method === "POST") {
      try {
        const data = await request.json();

        // 1. Extração inteligente do E-mail do Comprador (compatível com todas as plataformas)
        const buyerEmail = (
          data?.data?.buyer?.email || // Hotmart 2.0
          data?.buyer?.email ||       // Hotmart 1.0
          data?.Customer?.email ||    // Kiwify
          data?.customer?.email ||    // Kiwify v2
          data?.cus_email ||          // Eduzz
          data?.buyer_email ||        // Monetizze
          data?.email ||              // Padrão Geral / Cakto
          ""
        ).trim().toLowerCase();

        // 2. Extração do Status de Pagamento Aprovado
        const event = (
          data?.event || 
          data?.status || 
          data?.order_status || 
          data?.trans_status || 
          ""
        ).toLowerCase();

        const isApproved = 
          event.includes("approved") || 
          event.includes("paid") || 
          event.includes("completed") || 
          event.includes("pago") || 
          event.includes("autorizado") || 
          event.includes("purchase_approved") ||
          event === "3"; // Monetizze pago

        // 3. Extração do Nome do Produto / Pacote
        const productName = (
          data?.data?.product?.name || 
          data?.Product?.name || 
          data?.product_name || 
          data?.prod_name || 
          ""
        ).toLowerCase();

        if (buyerEmail && isApproved) {
          // Determina a quantidade de créditos com base no pacote comprado
          let creditsToAdd = 3; // Padrão
          if (productName.includes("50") || productName.includes("pro") || productName.includes("tcc")) {
            creditsToAdd = 50;
          } else if (productName.includes("7") || productName.includes("semestre")) {
            creditsToAdd = 7;
          } else if (productName.includes("unitario") || productName.includes("1")) {
            creditsToAdd = 1;
          }

          if (env.EMIA_KV) {
            const current = await env.EMIA_KV.get(`credits_${buyerEmail}`);
            const total = (current ? parseInt(current, 10) : 0) + creditsToAdd;
            await env.EMIA_KV.put(`credits_${buyerEmail}`, String(total));
          }

          return new Response(JSON.stringify({ 
            success: true, 
            platform: "Universal Webhook", 
            email: buyerEmail, 
            added: creditsToAdd 
          }), {
            headers: { "Content-Type": "application/json", ...corsHeaders }
          });
        }

        return new Response(JSON.stringify({ success: true, message: "Evento registrado" }), {
          headers: { "Content-Type": "application/json", ...corsHeaders }
        });
      } catch (err) {
        return new Response(JSON.stringify({ error: err.message }), {
          status: 500,
          headers: { "Content-Type": "application/json", ...corsHeaders }
        });
      }
    }

    // AI Generation — Proxy Seguro de IA com Retentativas Inteligentes (Exponential Backoff) e Fallback Multi-Modelo
    if (url.pathname === "/api/generate" && request.method === "POST") {
      try {
        let prompt = "";
        let model = "gemini-2.5-flash";
        let temperature = 0.7;

        const contentType = request.headers.get("content-type") || "";
        if (contentType.includes("multipart/form-data")) {
          const form = await request.formData();
          prompt = form.get("prompt") || "";
          if (!prompt && form.get("title")) {
            prompt = `Elabore um trabalho acadêmico sobre ${form.get("title")}`;
          }
          if (form.get("model")) model = String(form.get("model"));
        } else {
          const body = await request.json();
          prompt = body.prompt || body.text || "";
          if (body.model) model = body.model;
          if (body.temperature) temperature = body.temperature;
        }

        const apiKeysToTry = [
          request.headers.get("x-gemini-api-key"),
          request.headers.get("x-google-api-key"),
          env.GEMINI_API_KEY,
          env.GOOGLE_API_KEY
        ].filter(Boolean).filter((v, i, a) => a.indexOf(v) === i);

        if (apiKeysToTry.length === 0) {
          return new Response(JSON.stringify({ 
            success: false, 
            error: "Chave Gemini não configurada no servidor. Configure a variável GEMINI_API_KEY ou insira sua chave no app." 
          }), {
            status: 400,
            headers: { "Content-Type": "application/json", ...corsHeaders }
          });
        }

        const candidateModels = [
          model,
          "gemini-3.6-flash",
          "gemini-3.5-flash",
          "gemini-3-flash-preview",
          "gemini-3.7-flash",
          "gemini-2.5-flash",
          "gemini-2.0-flash",
          "gemini-1.5-flash"
        ].filter(Boolean).filter((v, i, a) => a.indexOf(v) === i);

        let generatedText = "";
        let lastError = null;

        for (const currentKey of apiKeysToTry) {
          for (const m of candidateModels) {
            let attempts = 0;
            const maxAttempts = 2;

            while (attempts < maxAttempts) {
              try {
                const geminiRes = await fetch(
                  `https://generativelanguage.googleapis.com/v1beta/models/${m}:generateContent?key=${currentKey}`,
                  {
                    method: "POST",
                    headers: { "Content-Type": "application/json" },
                    body: JSON.stringify({
                      contents: [{ parts: [{ text: prompt }] }],
                      generationConfig: {
                        temperature: temperature,
                        topP: 0.95,
                        maxOutputTokens: 8192
                      }
                    })
                  }
                );

                if (geminiRes.status === 429 || geminiRes.status === 500 || geminiRes.status === 503) {
                  attempts++;
                  const delay = Math.pow(2, attempts) * 1000 + Math.random() * 500;
                  await new Promise(resolve => setTimeout(resolve, delay));
                  continue;
                }

                if (geminiRes.ok) {
                  const data = await geminiRes.json();
                  const text = data?.candidates?.[0]?.content?.parts?.[0]?.text;
                  if (text && text.trim().length > 0) {
                    generatedText = text.trim();
                    return new Response(JSON.stringify({
                      success: true,
                      text: generatedText,
                      candidates: data?.candidates || [{ content: { parts: [{ text: generatedText }] } }]
                    }), {
                      headers: { "Content-Type": "application/json", ...corsHeaders }
                    });
                  }
                } else {
                  const errData = await geminiRes.json().catch(() => ({}));
                  lastError = errData?.error?.message || `HTTP ${geminiRes.status}`;
                  break; // Tenta o próximo modelo
                }
              } catch (fetchErr) {
                lastError = fetchErr.message;
                attempts++;
                const delay = Math.pow(2, attempts) * 1000 + Math.random() * 500;
                await new Promise(resolve => setTimeout(resolve, delay));
              }
            }
          }
        }

        return new Response(JSON.stringify({ 
          success: false, 
          error: lastError || "Todos os modelos Gemini e tentativas esgotaram." 
        }), {
          status: 502,
          headers: { "Content-Type": "application/json", ...corsHeaders }
        });
      } catch (err) {
        return new Response(JSON.stringify({ 
          success: false, 
          error: err.message || "Erro no Worker" 
        }), {
          status: 500,
          headers: { "Content-Type": "application/json", ...corsHeaders }
        });
      }
    }

    // Endpoint 100% Gratuito de Voz Neural de Estúdio (Edge Neural TTS: Francisca e Antonio)
    if (url.pathname === "/api/tts" && request.method === "POST") {
      try {
        const body = await request.json();
        const text = (body.text || "").trim();
        const voice = body.voice || "pt-BR-FranciscaNeural"; // pt-BR-FranciscaNeural ou pt-BR-AntonioNeural
        const rate = body.rate || "+0%";
        const pitch = body.pitch || "+0Hz";

        if (!text) {
          return new Response(JSON.stringify({ error: "Texto vazio" }), {
            status: 400,
            headers: { "Content-Type": "application/json", ...corsHeaders }
          });
        }

        // Sanitização de SSML para áudio neural
        const cleanText = text
          .replace(/&/g, '&amp;')
          .replace(/</g, '&lt;')
          .replace(/>/g, '&gt;');

        const ssml = `<speak version="1.0" xmlns="http://www.w3.org/2001/10/synthesis" xml:lang="pt-BR"><voice name="${voice}"><prosody rate="${rate}" pitch="${pitch}">${cleanText}</prosody></voice></speak>`;

        // Requisição para a API do Edge TTS Neural (100% Gratuita e Sem Chave Paga)
        const ttsUrl = `https://speech.platform.bing.com/consumer/speech/synthesize/readaloud/edge/v1?TrustedClientToken=6A5AA1D4EAFF4E9FB37E23D68491D6F4`;
        const ttsRes = await fetch(ttsUrl, {
          method: "POST",
          headers: {
            "Content-Type": "application/ssml+xml",
            "X-Timestamp": new Date().toISOString(),
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36 Edg/122.0.0.0"
          },
          body: ssml
        });

        if (ttsRes.ok) {
          const audioBuffer = await ttsRes.arrayBuffer();
          return new Response(audioBuffer, {
            headers: {
              "Content-Type": "audio/mpeg",
              "Cache-Control": "public, max-age=86400",
              ...corsHeaders
            }
          });
        } else {
          return new Response(JSON.stringify({ error: "Falha ao sintetizar áudio neural" }), {
            status: 502,
            headers: { "Content-Type": "application/json", ...corsHeaders }
          });
        }
      } catch (ttsErr) {
        return new Response(JSON.stringify({ error: ttsErr.message || "Erro TTS" }), {
          status: 500,
          headers: { "Content-Type": "application/json", ...corsHeaders }
        });
      }
    }

    // Serve Static Assets (Vite React app in dist/)
    if (env.ASSETS) {
      try {
        const res = await env.ASSETS.fetch(request);
        if (res.status === 404 && request.method === "GET") {
          return await env.ASSETS.fetch(new Request(new URL("/index.html", request.url)));
        }
        return res;
      } catch (assetErr) {
        console.error("Asset fetch error:", assetErr);
      }
    }

    return new Response("EDUTECH.EMIA Worker Online", {
      headers: { "Content-Type": "text/plain", ...corsHeaders }
    });
  }
};
