/// Ordered substitution tables. Order matters: longer patterns before shorter
/// substrings (e.g. `\pm` before `+`).

pub fn math_subs() -> Vec<(&'static str, &'static str)> {
    vec![
        ("\\pm", " pluss eller minus "),
        ("\\times", " ganger "),
        ("\\cdot", " ganger "),
        ("\\div", " delt på "),
        ("\\neq", " er ikke lik "),
        ("+", " pluss "),
        ("-", " minus "),
        ("_", " indeks "),
        ("^", " opphøyd i "),
        ("!", " fakultet "),
    ]
}

pub fn relation_subs() -> Vec<(&'static str, &'static str)> {
    vec![
        ("\\leq", " mindre enn eller lik "),
        ("\\geq", " større enn eller lik "),
        ("\\approx", " omtrent lik "),
        ("\\equiv", " ekvivalent med "),
        ("\\sim", " tilnærmet "),
        ("\\propto", " proporsjonal med "),
        ("\\in", " tilhører "),
        ("\\notin", " tilhører ikke "),
        ("\\subset", " delmengde av "),
        ("\\supset", " overmengde av "),
        ("=", " er lik "),
        ("<", " mindre enn "),
        (">", " større enn "),
    ]
}

pub fn greek_subs() -> Vec<(&'static str, &'static str)> {
    vec![
        // Lowercase
        ("\\alpha", " alfa "),
        ("\\beta", " beta "),
        ("\\gamma", " gamma "),
        ("\\delta", " delta "),
        ("\\epsilon", " epsilon "),
        ("\\varepsilon", " epsilon "),
        ("\\zeta", " zeta "),
        ("\\eta", " eta "),
        ("\\theta", " theta "),
        ("\\vartheta", " theta "),
        ("\\iota", " iota "),
        ("\\kappa", " kappa "),
        ("\\lambda", " lambda "),
        ("\\mu", " my "),
        ("\\nu", " ny "),
        ("\\xi", " ksi "),
        ("\\omicron", " omikron "),
        ("\\pi", " pi "),
        ("\\rho", " rho "),
        ("\\sigma", " sigma "),
        ("\\varsigma", " sigma "),
        ("\\tau", " tau "),
        ("\\upsilon", " ypsilon "),
        ("\\phi", " fi "),
        ("\\varphi", " fi "),
        ("\\chi", " khi "),
        ("\\psi", " psi "),
        ("\\omega", " omega "),
        // Uppercase
        ("\\Alpha", " stor alfa "),
        ("\\Beta", " stor beta "),
        ("\\Gamma", " stor gamma "),
        ("\\Delta", " stor delta "),
        ("\\Epsilon", " stor epsilon "),
        ("\\Zeta", " stor zeta "),
        ("\\Eta", " stor eta "),
        ("\\Theta", " stor theta "),
        ("\\Iota", " stor iota "),
        ("\\Kappa", " stor kappa "),
        ("\\Lambda", " stor lambda "),
        ("\\Mu", " stor my "),
        ("\\Nu", " stor ny "),
        ("\\Xi", " stor ksi "),
        ("\\Omicron", " stor omikron "),
        ("\\Pi", " stor pi "),
        ("\\Rho", " stor rho "),
        ("\\Sigma", " stor sigma "),
        ("\\Tau", " stor tau "),
        ("\\Upsilon", " stor ypsilon "),
        ("\\Phi", " stor fi "),
        ("\\Chi", " stor khi "),
        ("\\Psi", " stor psi "),
        ("\\Omega", " stor omega "),
    ]
}

pub fn trig_subs() -> Vec<(&'static str, &'static str)> {
    vec![
        ("\\sin", " sin "),
        ("\\cos", " cos "),
        ("\\tan", " tan "),
        ("\\cot", " cot "),
        ("\\sec", " sec "),
        ("\\csc", " csc "),
        ("\\arcsin", " arcsin "),
        ("\\arccos", " arccos "),
        ("\\arctan", " arctan "),
        ("\\log", " log "),
        ("\\ln", " ln "),
        ("\\exp", " exp "),
    ]
}

pub fn arrow_and_misc_subs() -> Vec<(&'static str, &'static str)> {
    vec![
        ("\\infty", " uendelig "),
        ("\\partial", " partiell "),
        ("\\nabla", " nabla "),
        ("\\rightarrow", " går mot "),
        ("\\leftarrow", " pil venstre "),
        ("\\Rightarrow", " medfører "),
        ("\\Leftarrow", " følger av "),
        ("\\iff", " hvis og bare hvis "),
        ("\\implies", " medfører "),
        ("\\to", " går mot "),
        ("\\forall", " for alle "),
        ("\\exists", " det eksisterer "),
        ("\\emptyset", " tom mengde "),
        ("\\ldots", " ... "),
        ("\\cdots", " ... "),
        ("\\dots", " ... "),
    ]
}

pub fn misc_sym_subs() -> Vec<(&'static str, &'static str)> {
    vec![
        ("\\odot", " sirkelprikk "),
        ("\\oplus", " sirkelpluss "),
        ("\\ominus", " sirkelminus "),
    ]
}

pub fn astro_subs() -> Vec<(&'static str, &'static str)> {
    vec![
        ("M_\\odot", " solmasser "),
        ("M_{\\odot}", " solmasser "),
        ("R_\\odot", " solradier "),
        ("R_{\\odot}", " solradier "),
    ]
}
