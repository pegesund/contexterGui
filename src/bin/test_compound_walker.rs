/// Test the compound FST walker.
/// Loads the Norwegian FST dictionary and tests compound word decomposition
/// with fuzzy matching per part.

use acatts_rust::compound_walker::{compound_fuzzy_walk, load_fst_from_mfst};
use std::path::PathBuf;
use std::time::Instant;

fn main() {
    let base = PathBuf::from(env!("CARGO_MANIFEST_DIR"));
    let mfst_path = base.join("../rustSpell/mtag-rs/data/fullform_bm.mfst");
    let dict_path = base.join("../rustSpell/mtag-rs/data/fullform_bm.mfst");
    let wf_path = base.join("../contexter-repo/training-data/wordfreq.tsv");

    println!("Loading FST from {}...", mfst_path.display());
    let t = Instant::now();
    let fst = load_fst_from_mfst(mfst_path.to_str().unwrap())
        .expect("Failed to load FST");
    println!("Loaded in {:?}", t.elapsed());

    let analyzer = mtag::Analyzer::new(dict_path.to_str().unwrap())
        .expect("Failed to load analyzer");

    let wordfreq = nostos_cognio::wordfreq::load_wordfreq(wf_path.as_path(), 10);
    println!("Loaded wordfreq: {} words", wordfreq.len());

    // Validate: word must be in dictionary AND not ONLY a proper noun
    let word_check = |w: &str| -> bool {
        if let Some(readings) = analyzer.dict_lookup(w) {
            readings.iter().any(|r| r.pos != mtag::types::Pos::Prop)
        } else {
            false
        }
    };

    // Last part of compound must be a noun (not verb, adverb, name, etc.)
    let noun_check = |w: &str| -> bool {
        if let Some(readings) = analyzer.dict_lookup(w) {
            readings.iter().any(|r| r.pos == mtag::types::Pos::Subst)
        } else {
            false
        }
    };

    // Check 200 compound words — find which are NOT in dictionary
    println!("\n=== Dictionary check (finding productive compounds) ===");
    let check_words = vec![
        // Nature combinations
        "fjellvann", "skogbrann", "elvemunning", "fjordbunn", "innsjøfisk",
        "havørn", "bregnevekst", "moseklump", "steinrøys", "sanddyne",
        "snøfonn", "isflak", "vindkast", "regnbyge", "tåkebank",
        "soloppgang", "måneskinnstur", "nordlyskveld", "frostgrense", "flomfare",
        // Kitchen/food combinations
        "kyllingfilet", "laksebit", "potetstappe", "løksuppe", "brøddeig",
        "eplemos", "jordbærgrøt", "bringebærsaft", "blåbærpai", "havresmuler",
        "kaffekanne", "tekopp", "smørkniv", "osteskive", "syltetøyglass",
        "krydderhylle", "kjøkkenbenk", "oppvaskkum", "stekepanne", "bakebrett",
        // Work/school combinations
        "mattelekse", "engelsklærer", "skolekantine", "eksamensperiode", "karakterbok",
        "kontorlandskap", "møteinnkalling", "prosjektrapport", "kundemøte", "fakturagebyr",
        "jobbsøknad", "ferieplan", "lønnsøkning", "overtidsbetaling", "pensjonssparing",
        "stillingsbeskrivelse", "kompetanseheving", "medarbeidersamtale", "sykefravær", "arbeidsmiljølov",
        // Home/daily life
        "sengeteppe", "putevår", "gardintrapp", "lampeskjerm", "bordplate",
        "skapshylle", "gulvmatte", "veggklokke", "taklysekrone", "dørvrider",
        "vaskemaskinslange", "tørketrommelen", "støvsugerpose", "oppvaskmiddelet", "rengjøringsklut",
        "søppelkasse", "papirkurv", "resirkuleringsstasjonen", "kompostbinge", "avfallssortering",
        // Sports/hobbies
        "fotballtrening", "håndballbane", "skiløype", "sykkelritt", "svømmestevne",
        "treningsprogram", "styrkeøvelse", "kondisjonstrening", "yogamatte", "kampresultat",
        "musikkinstrument", "fiolinstreng", "trommeundervisning", "gitarakkord", "pianostemmer",
        "malerkost", "tegnestift", "leirfigur", "garnhespe", "symaskinsnål",
        // Tech/modern
        "nettmøte", "skjermtid", "passordbytte", "programvareoppdatering", "datakrasj",
        "ladekabel", "batterilevetid", "skriverdriveren", "musematte", "tastaturlayout",
        "strømmetjeneste", "podkastepisode", "nyhetsvarsel", "innboksen", "spamfilter",
        "videomøte", "filopplasting", "skylagringstjeneste", "krypteringsalgoritme", "autentiseringsmetode",
        // Weather/seasons
        "vårrengjøring", "sommeravslutning", "høstferie", "vintersolverv",
        "regnværsdag", "solskinnshelg", "snøværsvarsel", "vindstyrke",
        "temperaturfall", "gradestokken", "fuktighetsmåler", "barometerstanden",
        // Transport
        "togbillett", "busskort", "sykkelhjul", "bilverkstedet", "fergekai",
        "flyplasshotell", "taxiregning", "trikkeholdeplass", "tunnelåpning", "motorveiavkjørsel",
        "parkeringsbøter", "fartskontroll", "trafikklysregulering", "fotgjengerfelt", "sykkelfeltet",
        // Body/health
        "hodepine", "magesmerte", "ryggsmerter", "kneoperasjon", "tannlegebesøk",
        "blodprøve", "røntgenundersøkelse", "allergitest", "vaksinering", "reseptbelagt",
        "treningsøkt", "søvnkvalitet", "stressnivå", "blodtrykksmåling", "kolesterolnivå",
    ];
    let mut not_in_dict = Vec::new();
    let mut in_count = 0;
    for w in &check_words {
        let in_fst = analyzer.has_word(w);
        if !in_fst {
            not_in_dict.push(*w);
            println!("  {:30} ✗ NOT in dictionary", w);
        } else {
            in_count += 1;
        }
    }
    println!("  {} in dictionary, {} NOT in dictionary\n", in_count, not_in_dict.len());

    // (input_misspelled, expected_correct, description)
    // Focus on productive compounds NOT in dictionary, with dyslexic errors.
    // Expected word must appear in top 50 results.
    let tests: Vec<(&str, Vec<&str>, &str)> = vec![
        // === Baseline: single words + known compounds ===
        ("sjøkken", vec!["kjøkken"], "single: s→k"),
        ("kjøkkenbort", vec!["kjøkkenbord"], "known compound: t→d"),
        ("sjøkkenbord", vec!["kjøkkenbord"], "known compound: sj→kj"),
        ("sjøkkenbort", vec!["kjøkkenbord"], "known compound: both parts"),

        // === Productive compounds NOT in dictionary — error in part 1 ===
        ("innsjefisk", vec!["innsjøfisk"], "productive: e→ø in innsjø"),
        ("kyllingsfilet", vec!["kyllingfilet"], "productive: extra s"),
        ("lakzebit", vec!["laksebit"], "productive: z→s"),
        ("jordbergrøt", vec!["jordbærgrøt"], "productive: e→æ in bær"),
        ("skollekantine", vec!["skolekantine"], "productive: ll→l"),
        ("eksamennsperiode", vec!["eksamensperiode"], "productive: nn→n"),
        ("prosjektrapport", vec!["prosjektrapport"], "productive: exact (in result)"),
        ("fakturaggebyr", vec!["fakturagebyr"], "productive: gg→g"),
        ("netbuttikk", vec!["nettbutikk"], "productive: t→tt, tt→t"),
        ("lekssehjlep", vec!["leksehjelp"], "productive: hj swap"),

        // === Productive compounds NOT in dictionary — error in part 2 ===
        ("innsjøfissk", vec!["innsjøfisk"], "productive: ss→s in fisk"),
        ("kyllingfilét", vec!["kyllingfilet"], "productive: é→e"),
        ("osteskivve", vec!["osteskive"], "productive: vv→v"),
        ("svømmestevnne", vec!["svømmestevne"], "productive: nn→n"),
        ("kampressultat", vec!["kampresultat"], "productive: ss→s"),
        ("passordbytte", vec!["passordbytte"], "productive: exact"),
        ("ladekabbel", vec!["ladekabel"], "productive: bb→b"),
        ("nyhettsvarsel", vec!["nyhetsvarsel"], "productive: tt→t"),
        ("busskort", vec!["busskort"], "productive: exact compound"),

        // === Productive compounds — errors in BOTH parts ===
        ("innsjefissk", vec!["innsjøfisk"], "both: e→ø + ss→s"),
        ("kyllingsfilét", vec!["kyllingfilet", "kyllingsfilet"], "both: extra s + é→e"),
        ("taklysekronne", vec!["taklysekrone"], "both: nn→n"),

        // === Phonetic errors (å↔o, ø↔e, æ↔a) in productive compounds ===
        ("frostgrennse", vec!["frostgrense"], "phonetic: nn→n"),
        ("nordlyskvell", vec!["nordlyskveld"], "phonetic: l→ld"),
        ("temperaturfel", vec!["temperaturfall"], "phonetic: e→a in fall"),
        ("sommeravsluttning", vec!["sommeravslutning"], "phonetic: tt→t"),
        ("solvskinnshellg", vec!["solskinnshelg"], "phonetic: v→ø, ll→l"),

        // === Binding letter errors ===
        ("møteinknalling", vec!["møteinnkalling"], "binding: n→nn"),
        ("rengjøringklut", vec!["rengjøringsklut"], "binding: missing s"),
        ("treningsstuddo", vec!["treningsstudio"], "binding: dd→d, o→io"),

        // === Modern/tech compounds ===
        ("skjermtitt", vec!["skjermtid"], "tech: tt→d"),
        ("datakrasj", vec!["datakrasj"], "tech: exact"),
        ("videomette", vec!["videomøte"], "tech: e→ø"),
        ("strømetjeneste", vec!["strømmetjeneste"], "tech: missing m"),

        // === Three-part productive compounds ===
        ("vaskemaskinslannge", vec!["vaskemaskinslange"], "3-part: nn→n"),
        ("flyplashotell", vec!["flyplasshotell"], "3-part: missing s"),

        // === Double consonant confusion ===
        ("jogamatte", vec!["yogamatte"], "double: j→y"),
        ("leerfigur", vec!["leirfigur"], "double: ee→ei"),
        ("allergittes", vec!["allergitest"], "double: tt→t, extra s"),
        ("sevnkvalitet", vec!["søvnkvalitet"], "phonetic: e→ø"),
    ];

    // Detailed dump for 3 hard cases
    for word in &["netbuttikk", "lekssehjlep", "allergittes", "temperaturfel"] {
        let r = compound_fuzzy_walk(&fst, word, &language::BokmalLanguage, Some(&wordfreq), Some(&word_check), Some(&noun_check));
        println!("\n  === {} ({} results) ===", word, r.len());
        for (i, x) in r.iter().take(10).enumerate() {
            let parts: Vec<String> = x.parts.iter()
                .map(|p| format!("{}({})", p.matched_word, p.edits)).collect();
            println!("    #{}: {} [{}] e={}", i+1, x.compound_word, parts.join("+"), x.total_edits);
        }
    }

    let mut pass = 0;
    let mut fail = 0;

    for (input, expected, desc) in &tests {
        let t = Instant::now();
        let results = compound_fuzzy_walk(&fst, &input.to_lowercase(), &language::BokmalLanguage, Some(&wordfreq), Some(&word_check), Some(&noun_check));
        let elapsed = t.elapsed();

        let result_words: Vec<&str> = results.iter().take(50).map(|r| r.compound_word.as_str()).collect();
        let found_rank = expected.iter().find_map(|exp| {
            result_words.iter().position(|r| r == exp).map(|pos| (exp, pos + 1))
        });

        let top3: Vec<String> = results.iter().take(3)
            .map(|r| {
                let parts: Vec<String> = r.parts.iter()
                    .map(|p| format!("{}({})", p.matched_word, p.edits))
                    .collect();
                format!("{}[{}] e={}", r.compound_word, parts.join("+"), r.total_edits)
            })
            .collect();

        if let Some((word, rank)) = found_rank {
            println!("  PASS #{:<3} ({:>5?}): {} → {} | top3: [{}]",
                rank, elapsed, desc, word, top3.join(", "));
            pass += 1;
        } else {
            // Check if expected is anywhere in full results
            let all_words: Vec<&str> = results.iter().map(|r| r.compound_word.as_str()).collect();
            let full_rank = expected.iter().find_map(|exp| {
                all_words.iter().position(|r| r == exp).map(|pos| (exp, pos + 1))
            });
            let rank_info = if let Some((w, r)) = full_rank {
                format!("found at #{} of {} total", r, results.len())
            } else {
                format!("NOT FOUND in {} results", results.len())
            };
            println!("  FAIL      ({:>5?}): {} — input='{}' got [{}] expected {:?} ({})",
                elapsed, desc, input, top3.join(", "), expected, rank_info);
            fail += 1;
        }
    }

    println!("\nResults: {}/{} passed", pass, pass + fail);
    if fail > 0 { std::process::exit(1); }
}
