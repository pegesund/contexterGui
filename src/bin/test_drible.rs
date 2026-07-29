// Scratch: why does "…god til å dr" suggest drikke but never drible?
// Runs complete_word exactly as bert_worker does (same params) and prints
// the full candidate list at different max_steps.
use std::path::PathBuf;

fn main() -> anyhow::Result<()> {
    let ort_path = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../onnxruntime/onnxruntime-win-x64-1.23.0/lib/onnxruntime.dll");
    unsafe { std::env::set_var("ORT_DYLIB_PATH", &ort_path); }

    let data = PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../../contexter-repo/training-data");
    let onnx = data.join("onnx/norbert4_base_int8.onnx");
    let tok = data.join("onnx/tokenizer.json");
    let wf_path = data.join("wordfreq.tsv");
    let dict_path = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../rustSpell/mtag-rs/data/fullform_bm.mfst");

    let mut model = nostos_cognio::model::Model::load(onnx.to_str().unwrap(), tok.to_str().unwrap())?;
    let pi = nostos_cognio::prefix_index::build_prefix_index(&model.tokenizer);
    let baselines = nostos_cognio::baseline::compute_baseline(&mut model)?;
    let wf = nostos_cognio::wordfreq::load_wordfreq(wf_path.as_path(), 10);
    let analyzer = mtag::Analyzer::new(dict_path.to_str().unwrap())
        .map_err(|e| anyhow::anyhow!("{}", e))?;

    println!("drible in mtag: {}", analyzer.has_word("drible"));
    println!("drikke in mtag: {}", analyzer.has_word("drikke"));
    println!("drible in wordfreq: {:?}", wf.get("drible"));
    println!("drikke in wordfreq: {:?}", wf.get("drikke"));

    let fallback_dict = |w: &str| -> bool { analyzer.has_word(w) };
    let fallback_prefix = |p: &str, limit: usize| -> Vec<String> { analyzer.prefix_lookup(p, limit) };

    let _ = (&fallback_dict, &fallback_prefix);

    // How does the tokenizer split 'drible' (word-initial)?
    for w in [" drible", " drikke", " dribler"] {
        if let Ok(enc) = model.tokenizer.encode(w, false) {
            let ids = enc.get_ids().to_vec();
            let toks: Vec<String> = ids.iter()
                .map(|&id| model.id_to_token[id as usize].clone()).collect();
            println!("'{}' → {:?}", w, toks);
        }
    }

    // Initial mask forward — same shape as complete_word's first pass
    // (glued mask, per NorBERT4 rule). What logit do the relevant start
    // tokens get, and where's the extension threshold (best - 15)?
    let ctx = "Jeg liker å spille fotball og er god til å";
    let masked = format!("{}{} .", ctx, " <mask>");
    let (logits, _) = model.single_forward(&masked)?;
    let mut best = f32::NEG_INFINITY;
    if let Some(entries) = pi.get("dr") {
        for (id, _) in entries {
            best = best.max(logits[*id as usize]);
        }
        println!("best 'dr' start-token logit: {:.2}, extend-threshold: {:.2}", best, best - 15.0);
        for want in ["dri", "drikke", "dra", "dr", "drive"] {
            if let Some((id, s)) = entries.iter().find(|(_, s)| s.to_lowercase() == want) {
                println!("  '{}' (id={}) logit={:.2}", s, id, logits[*id as usize]);
            }
        }
    }

    // If 'dri' gets extended, what is BERT's argmax continuation?
    let ext = format!("{} dri<mask> .", ctx);
    let (ext_logits, _) = model.single_forward(&ext)?;
    let mut idx: Vec<usize> = (0..ext_logits.len()).collect();
    idx.sort_by(|&a, &b| ext_logits[b].partial_cmp(&ext_logits[a]).unwrap());
    println!("top continuations after 'dri':");
    for &i in idx.iter().take(8) {
        println!("  '{}' logit={:.2}", model.id_to_token[i], ext_logits[i]);
    }
    Ok(())
}
