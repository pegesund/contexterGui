mod parser;
mod substitutions;

use regex::Regex;

/// Convert LaTeX text to Norwegian readable text.
///
/// Finds all `$...$` math regions, processes structural commands,
/// then applies flat symbol substitutions. Text outside `$...$` is
/// passed through unchanged.
pub fn latex_to_text(input: &str, astro: bool) -> String {
    let re = Regex::new(r"\$([^$]*)\$").unwrap();

    let result = re.replace_all(input, |caps: &regex::Captures| {
        let math_content = &caps[1];
        process_math(math_content, astro)
    });

    result.trim().to_string()
}

/// Convert a raw LaTeX math string (without $ delimiters) to Norwegian.
pub fn latex_math_to_text(input: &str, astro: bool) -> String {
    process_math(input, astro).trim().to_string()
}

fn process_math(input: &str, astro: bool) -> String {
    // Phase 1: structural commands (recursive parser)
    let mut text = parser::process_structural(input);

    // Phase 2: flat substitutions in order
    if astro {
        apply_subs(&mut text, &substitutions::astro_subs());
    }
    apply_subs(&mut text, &substitutions::trig_subs());
    apply_subs(&mut text, &substitutions::arrow_and_misc_subs());
    apply_subs(&mut text, &substitutions::greek_subs());
    apply_subs(&mut text, &substitutions::relation_subs());
    apply_subs(&mut text, &substitutions::misc_sym_subs());
    apply_subs(&mut text, &substitutions::math_subs());

    // Clean up: remove leftover braces, collapse whitespace
    text = text.replace('{', "").replace('}', "");
    collapse_whitespace(&text)
}

fn apply_subs(text: &mut String, subs: &[(&str, &str)]) {
    for &(pattern, replacement) in subs {
        *text = text.replace(pattern, replacement);
    }
}

fn collapse_whitespace(s: &str) -> String {
    let mut result = String::with_capacity(s.len());
    let mut prev_space = false;
    for c in s.chars() {
        if c.is_whitespace() {
            if !prev_space {
                result.push(' ');
            }
            prev_space = true;
        } else {
            result.push(c);
            prev_space = false;
        }
    }
    result
}

#[cfg(test)]
mod tests {
    use super::*;

    // ---- The 10 pix2text example outputs ----

    #[test]
    fn test_euler() {
        // e^{i \pi}+1=0
        let result = latex_math_to_text("e^{i \\pi}+1=0", false);
        assert!(result.contains("opphøyd i"));
        assert!(result.contains("pi"));
        assert!(result.contains("pluss"));
        assert!(result.contains("er lik"));
        println!("Euler: {}", result);
    }

    #[test]
    fn test_integral() {
        // \begin{aligned} {\int_{0}^{\infty} e^{-x^{2}} d x=\frac{\sqrt{\pi}} {2}} \\ \end{aligned}
        let input = "\\begin{aligned} {\\int_{0}^{\\infty} e^{-x^{2}} d x=\\frac{\\sqrt{\\pi}} {2}} \\\\ \\end{aligned}";
        let result = latex_math_to_text(input, false);
        assert!(result.contains("integralet"));
        assert!(result.contains("fra 0"));
        assert!(result.contains("uendelig"));
        assert!(result.contains("kvadratroten av"));
        assert!(result.contains("over"));
        println!("Integral: {}", result);
    }

    #[test]
    fn test_matrix() {
        // A=\left( \begin{matrix} {1} & {2} \\ {3} & {4} \\ \end{matrix} \right)
        let input = "A=\\left( \\begin{matrix} {1} & {2} \\\\ {3} & {4} \\\\ \\end{matrix} \\right)";
        let result = latex_math_to_text(input, false);
        assert!(result.contains("matrise"));
        assert!(result.contains("1"));
        assert!(result.contains("2"));
        assert!(result.contains("3"));
        assert!(result.contains("4"));
        println!("Matrix: {}", result);
    }

    #[test]
    fn test_taylor() {
        // e^{x}=\sum_{n=0}^{\infty} \frac{x^{n}} {n !}
        let input = "e^{x}=\\sum_{n=0}^{\\infty} \\frac{x^{n}} {n !}";
        let result = latex_math_to_text(input, false);
        assert!(result.contains("opphøyd i"));
        assert!(result.contains("summen"));
        assert!(result.contains("fra"));
        assert!(result.contains("uendelig"));
        assert!(result.contains("over"));
        assert!(result.contains("fakultet"));
        println!("Taylor: {}", result);
    }

    #[test]
    fn test_pythagoras() {
        // a^{2}+b^{2}=c^{2}
        let result = latex_math_to_text("a^{2}+b^{2}=c^{2}", false);
        assert!(result.contains("opphøyd i"));
        assert!(result.contains("pluss"));
        assert!(result.contains("er lik"));
        println!("Pythagoras: {}", result);
    }

    #[test]
    fn test_limit() {
        // \underset{n \to\infty} {\operatorname* {l i m}} \left( 1+\frac{1} {n} \right)^{n}=e
        let input = "\\underset{n \\to\\infty} {\\operatorname* {l i m}} \\left( 1+\\frac{1} {n} \\right)^{n}=e";
        let result = latex_math_to_text(input, false);
        assert!(result.contains("lim"));
        assert!(result.contains("uendelig"));
        assert!(result.contains("over"));
        assert!(result.contains("er lik"));
        println!("Limit: {}", result);
    }

    #[test]
    fn test_derivative() {
        // \frac{d} {d x} \operatorname{s i n} ( x )=\operatorname{c o s} ( x )
        let input = "\\frac{d} {d x} \\operatorname{s i n} ( x )=\\operatorname{c o s} ( x )";
        let result = latex_math_to_text(input, false);
        assert!(result.contains("over"));
        assert!(result.contains("sin"));
        assert!(result.contains("cos"));
        assert!(result.contains("er lik"));
        println!("Derivative: {}", result);
    }

    #[test]
    fn test_binomial() {
        // ( a+b )^{n}=\sum_{k=0}^{n} \binom{n} {k} a^{k} b^{n-k}
        let input = "( a+b )^{n}=\\sum_{k=0}^{n} \\binom{n} {k} a^{k} b^{n-k}";
        let result = latex_math_to_text(input, false);
        assert!(result.contains("opphøyd i"));
        assert!(result.contains("summen"));
        assert!(result.contains("over"));
        println!("Binomial: {}", result);
    }

    #[test]
    fn test_quadratic() {
        // x=\frac{-b \pm\sqrt{b^{2}-4 a c}} {2 a}
        let input = "x=\\frac{-b \\pm\\sqrt{b^{2}-4 a c}} {2 a}";
        let result = latex_math_to_text(input, false);
        assert!(result.contains("over"));
        assert!(result.contains("pluss eller minus"));
        assert!(result.contains("kvadratroten av"));
        println!("Quadratic: {}", result);
    }

    #[test]
    fn test_basel() {
        // \sum_{n=1}^{\infty} \cfrac{1} {n^{2}}=\cfrac{\pi^{2}} {6}
        let input = "\\sum_{n=1}^{\\infty} \\cfrac{1} {n^{2}}=\\cfrac{\\pi^{2}} {6}";
        let result = latex_math_to_text(input, false);
        assert!(result.contains("summen"));
        assert!(result.contains("uendelig"));
        assert!(result.contains("over"));
        assert!(result.contains("pi"));
        println!("Basel: {}", result);
    }

    // ---- Additional tests ----

    #[test]
    fn test_dollar_delimiters() {
        let result = latex_to_text("Gitt $x + y$ er svaret", false);
        assert!(result.contains("Gitt"));
        assert!(result.contains("pluss"));
        assert!(result.contains("er svaret"));
    }

    #[test]
    fn test_nested_frac_in_sqrt() {
        let result = latex_math_to_text("\\sqrt{\\frac{a}{b}}", false);
        assert!(result.contains("kvadratroten av"));
        assert!(result.contains("a over b"));
    }

    #[test]
    fn test_frac_of_frac() {
        let result = latex_math_to_text("\\frac{\\frac{1}{2}}{\\frac{3}{4}}", false);
        assert!(result.contains("1 over 2"));
        assert!(result.contains("3 over 4"));
        println!("Frac of frac: {}", result);
    }

    #[test]
    fn test_astro() {
        let result = latex_to_text("$M_\\odot$", true);
        assert!(result.contains("solmasser"));
    }
}
