//! Math formula recognition from images.
//! Uses pix2text ONNX models (encoder + decoder) to produce LaTeX,
//! then latext_no to convert to Norwegian readable text.

use super::latext_no;

use std::sync::mpsc;
use ort::session::{builder::GraphOptimizationLevel, Session};
use ort::value::TensorRef;

/// Lazy-loaded math OCR engine.
pub struct MathOcr {
    encoder: Session,
    decoder: Session,
    tokenizer: tokenizers::Tokenizer,
    bos_token_id: i64,
    eos_token_id: i64,
}

impl MathOcr {
    pub fn load() -> anyhow::Result<Self> {
        let model_dir = dirs::home_dir()
            .unwrap_or_default()
            .join(".pix2text/1.1/mfr-1.5-onnx");

        if !model_dir.exists() {
            return Err(anyhow::anyhow!("Math OCR models not found at {:?}", model_dir));
        }

        let encoder = Session::builder()?
            .with_optimization_level(GraphOptimizationLevel::Level3)?
            .with_intra_threads(2)?
            .commit_from_file(model_dir.join("encoder_model.onnx"))?;

        let decoder = Session::builder()?
            .with_optimization_level(GraphOptimizationLevel::Level3)?
            .with_intra_threads(2)?
            .commit_from_file(model_dir.join("decoder_model.onnx"))?;

        let tokenizer = tokenizers::Tokenizer::from_file(model_dir.join("tokenizer.json"))
            .map_err(|e| anyhow::anyhow!("Failed to load math tokenizer: {}", e))?;

        Ok(Self {
            encoder,
            decoder,
            tokenizer,
            bos_token_id: 1,
            eos_token_id: 2,
        })
    }

    /// Recognize math formula from image, return Norwegian text.
    pub fn recognize_file(&mut self, image_path: &str) -> anyhow::Result<String> {
        let latex = self.image_to_latex(image_path)?;
        eprintln!("Math LaTeX: {}", latex);
        let norwegian = latext_no::latex_math_to_text(&latex, false);
        Ok(norwegian)
    }

    /// Image → LaTeX string via encoder-decoder.
    fn image_to_latex(&mut self, image_path: &str) -> anyhow::Result<String> {
        let pixels = self.preprocess_image(image_path)?;

        // Encode: pixel_values [1, 3, 384, 384] → encoder_hidden_states [1, 578, 384]
        let pv = TensorRef::from_array_view(
            (vec![1i64, 3, 384, 384], pixels.as_slice()),
        )?;
        let enc_out = self.encoder.run(ort::inputs![pv])?;
        let (enc_dims, enc_data) = enc_out[0].try_extract_tensor::<f32>()?;
        let enc_hidden: Vec<f32> = enc_data.to_vec();
        let enc_seq_len = enc_dims[1] as i64;
        let enc_hidden_dim = enc_dims[2] as i64;

        // Autoregressive decode
        let mut token_ids: Vec<i64> = vec![self.bos_token_id];

        for _ in 0..512 {
            let seq_len = token_ids.len();
            let ids = TensorRef::from_array_view(
                (vec![1i64, seq_len as i64], token_ids.as_slice()),
            )?;
            let hidden = TensorRef::from_array_view(
                (vec![1i64, enc_seq_len, enc_hidden_dim], enc_hidden.as_slice()),
            )?;

            let dec_out = self.decoder.run(ort::inputs![ids, hidden])?;
            let (dec_dims, dec_data) = dec_out[0].try_extract_tensor::<f32>()?;

            // logits shape: [1, seq_len, vocab_size]
            let vocab_size = dec_dims[2] as usize;
            let last_pos_offset = (seq_len - 1) * vocab_size;
            let last_logits = &dec_data[last_pos_offset..last_pos_offset + vocab_size];

            let next_token = last_logits.iter()
                .enumerate()
                .max_by(|(_, a), (_, b)| a.partial_cmp(b).unwrap_or(std::cmp::Ordering::Equal))
                .map(|(idx, _)| idx as i64)
                .unwrap_or(self.eos_token_id);

            if next_token == self.eos_token_id {
                break;
            }
            token_ids.push(next_token);
        }

        // Decode tokens (skip bos)
        let ids_u32: Vec<u32> = token_ids[1..].iter().map(|&id| id as u32).collect();
        let latex = self.tokenizer.decode(&ids_u32, true)
            .map_err(|e| anyhow::anyhow!("Tokenizer decode: {}", e))?;

        Ok(latex)
    }

    /// Resize to 384x384, normalize: (pixel/255 - 0.5) / 0.5
    fn preprocess_image(&self, image_path: &str) -> anyhow::Result<Vec<f32>> {
        let img = image::open(image_path)?
            .resize_exact(384, 384, image::imageops::FilterType::Lanczos3)
            .to_rgb8();

        // Channel-first: [3, 384, 384] flattened
        let mut pixels = vec![0.0f32; 3 * 384 * 384];
        for y in 0..384usize {
            for x in 0..384usize {
                let p = img.get_pixel(x as u32, y as u32);
                for c in 0..3usize {
                    pixels[c * 384 * 384 + y * 384 + x] = (p[c] as f32 / 255.0 - 0.5) / 0.5;
                }
            }
        }
        Ok(pixels)
    }
}

/// Start math OCR in background thread with lazy model loading.
pub fn start_math_ocr(image_path: String) -> mpsc::Receiver<Result<String, String>> {
    let (tx, rx) = mpsc::channel();

    std::thread::Builder::new()
        .name("math-ocr".into())
        .spawn(move || {
            use std::sync::{OnceLock, Mutex};
            static MODEL: OnceLock<Mutex<Option<MathOcr>>> = OnceLock::new();

            let mtx = MODEL.get_or_init(|| {
                eprintln!("Loading math OCR models...");
                match MathOcr::load() {
                    Ok(m) => Mutex::new(Some(m)),
                    Err(e) => {
                        eprintln!("Math OCR load error: {}", e);
                        Mutex::new(None)
                    }
                }
            });

            let result = if let Ok(mut guard) = mtx.lock() {
                if let Some(ref mut m) = *guard {
                    m.recognize_file(&image_path).map_err(|e| format!("{}", e))
                } else {
                    Err("Math OCR models not loaded".into())
                }
            } else {
                Err("Math OCR lock failed".into())
            };

            let _ = tx.send(result);
        })
        .ok();

    rx
}
