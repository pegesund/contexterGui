//! LLM grammar correction actor — dedicated thread for API requests.
//!
//! Sends batches of Norwegian sentences to an LLM (via requesty.ai)
//! for grammar correction. Returns corrections asynchronously.

use std::sync::mpsc;

const API_URL: &str = "https://router.requesty.ai/v1/chat/completions";
const API_KEY: &str = "rqsty-sk-ITUDt2zDS9Clb8OtNi8xJUZPkXxGErbOFD4chcu0qPjQr4QfW0Zg/1gdMLeC2A6myVqvckRD5Xd25DqHL4OLb46EKssNfZDGc26RiYn0QA4=";
const MODEL: &str = "deepseek/deepseek-chat";

const SYSTEM_PROMPT: &str = "Du er en norsk korrekturleser. Korriger grammatikk og stavefeil. Svar KUN med JSON-array. For hver setning: {\"corrected\": \"den korrigerte setningen\", \"explanation\": \"kort forklaring\"} eller {\"ok\": true} hvis korrekt.";

pub struct LlmRequest {
    pub request_id: u64,
    pub sentences: Vec<(String, String)>, // (sentence, paragraph_id)
    pub hashes: Vec<u64>,
}

pub struct LlmCorrection {
    pub original: String,
    pub corrected: String,
    pub explanation: String,
    pub paragraph_id: String,
}

pub struct LlmResponse {
    pub request_id: u64,
    pub corrections: Vec<LlmCorrection>,
    pub checked_hashes: Vec<u64>,
}

pub struct LlmActorHandle {
    sender: mpsc::Sender<LlmRequest>,
    receiver: mpsc::Receiver<LlmResponse>,
    next_id: u64,
}

impl LlmActorHandle {
    pub fn send(&mut self, sentences: Vec<(String, String)>, hashes: Vec<u64>) -> u64 {
        let id = self.next_id;
        self.next_id += 1;
        let _ = self.sender.send(LlmRequest {
            request_id: id,
            sentences,
            hashes,
        });
        id
    }

    pub fn try_recv(&self) -> Option<LlmResponse> {
        self.receiver.try_recv().ok()
    }
}

pub fn spawn_llm_actor(repaint_ctx: egui::Context) -> LlmActorHandle {
    let (req_tx, req_rx) = mpsc::channel::<LlmRequest>();
    let (resp_tx, resp_rx) = mpsc::channel::<LlmResponse>();

    std::thread::Builder::new()
        .name("llm-actor".into())
        .spawn(move || {
            while let Ok(req) = req_rx.recv() {
                let corrections = process_batch(&req);
                let _ = resp_tx.send(LlmResponse {
                    request_id: req.request_id,
                    corrections,
                    checked_hashes: req.hashes,
                });
                repaint_ctx.request_repaint();
            }
        })
        .expect("Failed to spawn LLM actor");

    LlmActorHandle {
        sender: req_tx,
        receiver: resp_rx,
        next_id: 0,
    }
}

fn process_batch(req: &LlmRequest) -> Vec<LlmCorrection> {
    if req.sentences.is_empty() {
        return Vec::new();
    }

    // Build numbered sentence list
    let user_msg: String = req.sentences.iter().enumerate()
        .map(|(i, (s, _))| format!("{}. \"{}\"", i + 1, s))
        .collect::<Vec<_>>()
        .join("\n");

    // Build API request
    let body = serde_json::json!({
        "model": MODEL,
        "messages": [
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user", "content": user_msg}
        ],
        "temperature": 0
    });

    {
        use std::io::Write;
        if let Ok(mut f) = std::fs::OpenOptions::new().create(true).append(true)
            .open(std::env::temp_dir().join("acatts-rust.log")) {
            let _ = writeln!(f, "LLM SEND: {} sentences", req.sentences.len());
        }
    }

    // Send HTTP request
    let response = match ureq::post(API_URL)
        .set("Authorization", &format!("Bearer {}", API_KEY))
        .set("Content-Type", "application/json")
        .send_string(&body.to_string())
    {
        Ok(resp) => resp,
        Err(e) => {
            eprintln!("LLM API error: {}", e);
            return Vec::new();
        }
    };

    // Read response body
    let resp_body = match response.into_string() {
        Ok(b) => b,
        Err(e) => {
            eprintln!("LLM response read error: {}", e);
            return Vec::new();
        }
    };

    // Parse outer JSON (OpenAI format)
    let outer: serde_json::Value = match serde_json::from_str(&resp_body) {
        Ok(v) => v,
        Err(e) => {
            eprintln!("LLM JSON parse error: {}", e);
            return Vec::new();
        }
    };

    let content = match outer["choices"][0]["message"]["content"].as_str() {
        Some(c) => c,
        None => {
            eprintln!("LLM: no content in response");
            return Vec::new();
        }
    };

    // Parse inner JSON array (our correction format)
    let results: Vec<serde_json::Value> = match serde_json::from_str(content) {
        Ok(v) => v,
        Err(_) => {
            // Try stripping markdown code fences
            let stripped = content.trim()
                .trim_start_matches("```json").trim_start_matches("```")
                .trim_end_matches("```").trim();
            match serde_json::from_str(stripped) {
                Ok(v) => v,
                Err(e) => {
                    eprintln!("LLM inner JSON parse error: {} content='{}'", e, &content[..content.len().min(200)]);
                    return Vec::new();
                }
            }
        }
    };

    // Build corrections
    let mut corrections = Vec::new();
    for (i, result) in results.iter().enumerate() {
        if i >= req.sentences.len() { break; }
        let (ref original, ref para_id) = req.sentences[i];

        let ok = result["ok"].as_bool().unwrap_or(false);
        if ok { continue; }

        let corrected = result["corrected"].as_str().unwrap_or("").to_string();
        if corrected.is_empty() || corrected == *original { continue; }

        let explanation = result["explanation"].as_str()
            .map(|e| e.to_string())
            .unwrap_or_else(|| format!("AI: «{}» → «{}»", original, corrected));

        {
            use std::io::Write;
            if let Ok(mut f) = std::fs::OpenOptions::new().create(true).append(true)
                .open(std::env::temp_dir().join("acatts-rust.log")) {
                let _ = writeln!(f, "LLM correction: '{}' → '{}'", original, corrected);
            }
        }

        corrections.push(LlmCorrection {
            original: original.clone(),
            corrected,
            explanation,
            paragraph_id: para_id.clone(),
        });
    }

    corrections
}
