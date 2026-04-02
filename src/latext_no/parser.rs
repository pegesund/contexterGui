/// Recursive parser for structural LaTeX commands.
///
/// Handles: \frac, \cfrac, \dfrac, \sqrt, \binom, \sum, \int, \prod, \lim,
/// \underset, \overset, \operatorname, \operatorname*, \left/\right,
/// \begin{matrix|pmatrix|bmatrix|aligned|cases}...\end{...},
/// &, \\, and nested brace groups.

/// Process all structural commands in a LaTeX string, recursively.
pub fn process_structural(input: &str) -> String {
    let mut result = String::new();
    let chars: Vec<char> = input.chars().collect();
    let mut i = 0;

    while i < chars.len() {
        if chars[i] == '\\' {
            let (replacement, new_i) = parse_command(&chars, i);
            result.push_str(&replacement);
            i = new_i;
        } else if chars[i] == '{' {
            let (content, new_i) = extract_brace_group(&chars, i);
            let processed = process_structural(&content);
            result.push_str(&processed);
            i = new_i;
        } else if chars[i] == '&' {
            result.push_str(", ");
            i += 1;
        } else {
            result.push(chars[i]);
            i += 1;
        }
    }

    result
}

/// Extract content inside a brace group starting at pos (which must be '{').
/// Returns (content_inside_braces, position_after_closing_brace).
fn extract_brace_group(chars: &[char], pos: usize) -> (String, usize) {
    debug_assert!(pos < chars.len() && chars[pos] == '{');
    let mut depth = 1;
    let mut i = pos + 1;
    let mut content = String::new();

    while i < chars.len() && depth > 0 {
        if chars[i] == '{' {
            depth += 1;
            if depth > 1 {
                content.push('{');
            }
        } else if chars[i] == '}' {
            depth -= 1;
            if depth > 0 {
                content.push('}');
            }
        } else {
            content.push(chars[i]);
        }
        i += 1;
    }

    (content, i)
}

/// Extract a single argument: either a brace group or a single token
/// (a command like \pi, or a single character).
fn extract_arg(chars: &[char], pos: usize) -> (String, usize) {
    let i = skip_whitespace(chars, pos);
    if i >= chars.len() {
        return (String::new(), i);
    }

    if chars[i] == '{' {
        extract_brace_group(chars, i)
    } else if chars[i] == '\\' {
        // Single command as argument
        let (cmd, new_i) = read_command_name(chars, i);
        (cmd, new_i)
    } else {
        (chars[i].to_string(), i + 1)
    }
}

/// Extract optional bracket argument: [content]. Returns None if no bracket.
fn extract_bracket_arg(chars: &[char], pos: usize) -> (Option<String>, usize) {
    let i = skip_whitespace(chars, pos);
    if i >= chars.len() || chars[i] != '[' {
        return (None, pos);
    }

    let mut depth = 1;
    let mut j = i + 1;
    let mut content = String::new();

    while j < chars.len() && depth > 0 {
        if chars[j] == '[' {
            depth += 1;
        } else if chars[j] == ']' {
            depth -= 1;
            if depth == 0 {
                break;
            }
        }
        content.push(chars[j]);
        j += 1;
    }

    (Some(content), j + 1)
}

/// Read a command name starting at a backslash. Returns the full command
/// string (including backslash) and the position after.
fn read_command_name(chars: &[char], pos: usize) -> (String, usize) {
    debug_assert!(chars[pos] == '\\');
    let mut i = pos + 1;
    let mut name = String::from("\\");

    if i < chars.len() && !chars[i].is_alphabetic() {
        // Single-char command like \\ or \{
        name.push(chars[i]);
        return (name, i + 1);
    }

    while i < chars.len() && chars[i].is_alphabetic() {
        name.push(chars[i]);
        i += 1;
    }

    // Handle \operatorname*
    if name == "\\operatorname" && i < chars.len() && chars[i] == '*' {
        name.push('*');
        i += 1;
    }

    (name, i)
}

fn skip_whitespace(chars: &[char], pos: usize) -> usize {
    let mut i = pos;
    while i < chars.len() && chars[i].is_whitespace() {
        i += 1;
    }
    i
}

/// Try to extract optional sub/superscript bounds after a big operator.
/// Returns (sub, sup, new_pos). Either or both can be None.
fn extract_bounds(chars: &[char], pos: usize) -> (Option<String>, Option<String>, usize) {
    let mut sub = None;
    let mut sup = None;
    let mut i = skip_whitespace(chars, pos);

    // Could be _{}^{} or ^{}_{} (either order), or just one, or neither
    for _ in 0..2 {
        i = skip_whitespace(chars, i);
        if i >= chars.len() {
            break;
        }
        if chars[i] == '_' {
            let (arg, new_i) = extract_arg(chars, i + 1);
            sub = Some(process_structural(&arg));
            i = new_i;
        } else if chars[i] == '^' {
            let (arg, new_i) = extract_arg(chars, i + 1);
            sup = Some(process_structural(&arg));
            i = new_i;
        } else {
            break;
        }
    }

    (sub, sup, i)
}

/// Parse a command starting at position pos (which is '\').
/// Returns (replacement_text, new_position).
fn parse_command(chars: &[char], pos: usize) -> (String, usize) {
    let (cmd, after_cmd) = read_command_name(chars, pos);

    match cmd.as_str() {
        // --- Structural commands with arguments ---

        "\\frac" | "\\cfrac" | "\\dfrac" | "\\tfrac" => {
            let (num, after_num) = extract_arg(chars, after_cmd);
            let (den, after_den) = extract_arg(chars, after_num);
            let num_text = process_structural(&num);
            let den_text = process_structural(&den);
            (format!(" {} over {} ", num_text.trim(), den_text.trim()), after_den)
        }

        "\\sqrt" => {
            let (bracket, after_bracket) = extract_bracket_arg(chars, after_cmd);
            let (body, after_body) = extract_arg(chars, after_bracket);
            let body_text = process_structural(&body);
            match bracket {
                Some(n) => {
                    let n_text = process_structural(&n);
                    (format!(" {}-te roten av {} ", n_text.trim(), body_text.trim()), after_body)
                }
                None => {
                    (format!(" kvadratroten av {} ", body_text.trim()), after_body)
                }
            }
        }

        "\\binom" => {
            let (n, after_n) = extract_arg(chars, after_cmd);
            let (k, after_k) = extract_arg(chars, after_n);
            let n_text = process_structural(&n);
            let k_text = process_structural(&k);
            (format!(" {} over {} ", n_text.trim(), k_text.trim()), after_k)
        }

        // --- Big operators with optional bounds ---

        "\\sum" => {
            let (sub, sup, after_bounds) = extract_bounds(chars, after_cmd);
            let mut text = String::from(" summen");
            if let Some(s) = sub {
                text.push_str(&format!(" fra {}", s.trim()));
            }
            if let Some(s) = sup {
                text.push_str(&format!(" til {}", s.trim()));
            }
            text.push_str(" av ");
            (text, after_bounds)
        }

        "\\prod" => {
            let (sub, sup, after_bounds) = extract_bounds(chars, after_cmd);
            let mut text = String::from(" produktet");
            if let Some(s) = sub {
                text.push_str(&format!(" fra {}", s.trim()));
            }
            if let Some(s) = sup {
                text.push_str(&format!(" til {}", s.trim()));
            }
            text.push_str(" av ");
            (text, after_bounds)
        }

        "\\int" | "\\iint" | "\\iiint" | "\\oint" => {
            let prefix = match cmd.as_str() {
                "\\iint" => " dobbeltintegralet",
                "\\iiint" => " trippelintegralet",
                "\\oint" => " kurveintegralet",
                _ => " integralet",
            };
            let (sub, sup, after_bounds) = extract_bounds(chars, after_cmd);
            let mut text = String::from(prefix);
            if let Some(s) = sub {
                text.push_str(&format!(" fra {}", s.trim()));
            }
            if let Some(s) = sup {
                text.push_str(&format!(" til {}", s.trim()));
            }
            text.push_str(" av ");
            (text, after_bounds)
        }

        "\\lim" => {
            let (sub, _sup, after_bounds) = extract_bounds(chars, after_cmd);
            let mut text = String::from(" grenseverdien");
            if let Some(s) = sub {
                text.push_str(&format!(" når {}", s.trim()));
            }
            text.push(' ');
            (text, after_bounds)
        }

        // --- \underset{below}{expr} and \overset{above}{expr} ---

        "\\underset" => {
            let (below, after_below) = extract_arg(chars, after_cmd);
            let (expr, after_expr) = extract_arg(chars, after_below);
            let below_text = process_structural(&below);
            let expr_text = process_structural(&expr);
            // Commonly used for \underset{n\to\infty}{\lim}
            // We try to produce natural Norwegian
            let expr_trimmed = expr_text.trim();
            if expr_trimmed.contains("grenseverdien") {
                // Already handled by inner \lim, just add the subscript
                (format!(" {} når {} ", expr_trimmed, below_text.trim()), after_expr)
            } else {
                (format!(" {} der {} ", expr_trimmed, below_text.trim()), after_expr)
            }
        }

        "\\overset" => {
            let (above, after_above) = extract_arg(chars, after_cmd);
            let (expr, after_expr) = extract_arg(chars, after_above);
            let above_text = process_structural(&above);
            let expr_text = process_structural(&expr);
            (format!(" {} med {} over ", expr_text.trim(), above_text.trim()), after_expr)
        }

        // --- \operatorname{text} / \operatorname*{text} ---

        "\\operatorname" | "\\operatorname*" => {
            let (body, after_body) = extract_arg(chars, after_cmd);
            // MFR sometimes outputs spaced letters like "s i n" — collapse them
            let collapsed: String = body.chars().filter(|c| !c.is_whitespace()).collect();
            (format!(" {} ", collapsed), after_body)
        }

        // --- \left and \right — just strip the command, keep delimiter ---

        "\\left" => {
            let i = skip_whitespace(chars, after_cmd);
            if i < chars.len() {
                if chars[i] == '\\' {
                    // \left\{ etc
                    let (_, after_delim) = read_command_name(chars, i);
                    (String::new(), after_delim)
                } else {
                    // \left( or \left[ — skip the delimiter char, pass through
                    let delim = match chars[i] {
                        '(' => "(",
                        '[' => "[",
                        '|' => "|",
                        '.' => "",
                        _ => "",
                    };
                    (delim.to_string(), i + 1)
                }
            } else {
                (String::new(), after_cmd)
            }
        }

        "\\right" => {
            let i = skip_whitespace(chars, after_cmd);
            if i < chars.len() {
                if chars[i] == '\\' {
                    let (_, after_delim) = read_command_name(chars, i);
                    (String::new(), after_delim)
                } else {
                    let delim = match chars[i] {
                        ')' => ")",
                        ']' => "]",
                        '|' => "|",
                        '.' => "",
                        _ => "",
                    };
                    (delim.to_string(), i + 1)
                }
            } else {
                (String::new(), after_cmd)
            }
        }

        // --- \begin{env}...\end{env} ---

        "\\begin" => {
            let (env_name, after_name) = extract_arg(chars, after_cmd);
            let (body, after_end) = extract_until_end(chars, after_name, &env_name);
            let processed = process_environment(&env_name, &body);
            (processed, after_end)
        }

        // \end should not appear standalone (consumed by extract_until_end)
        "\\end" => {
            let (_env_name, after_name) = extract_arg(chars, after_cmd);
            (String::new(), after_name)
        }

        // --- \text, \mathrm, \mathbf etc. — just output content ---

        "\\text" | "\\mathrm" | "\\mathbf" | "\\mathit" | "\\mathsf"
        | "\\textbf" | "\\textit" | "\\textrm" => {
            let (body, after_body) = extract_arg(chars, after_cmd);
            (format!(" {} ", body.trim()), after_body)
        }

        // --- \\ (line break / row separator) ---
        "\\\\" => {
            ("; ".to_string(), after_cmd)
        }

        // --- Spacing commands — just emit a space ---
        "\\," | "\\;" | "\\:" | "\\!" | "\\quad" | "\\qquad" | "\\ " => {
            (" ".to_string(), after_cmd)
        }

        // --- \overline, \underline, \hat, \bar, \vec, \dot, \tilde ---

        "\\overline" => {
            let (body, after_body) = extract_arg(chars, after_cmd);
            let text = process_structural(&body);
            (format!(" {} med strek over ", text.trim()), after_body)
        }

        "\\underline" => {
            let (body, after_body) = extract_arg(chars, after_cmd);
            let text = process_structural(&body);
            (format!(" {} med strek under ", text.trim()), after_body)
        }

        "\\hat" => {
            let (body, after_body) = extract_arg(chars, after_cmd);
            let text = process_structural(&body);
            (format!(" {} hatt ", text.trim()), after_body)
        }

        "\\bar" => {
            let (body, after_body) = extract_arg(chars, after_cmd);
            let text = process_structural(&body);
            (format!(" {} strek ", text.trim()), after_body)
        }

        "\\vec" => {
            let (body, after_body) = extract_arg(chars, after_cmd);
            let text = process_structural(&body);
            (format!(" vektor {} ", text.trim()), after_body)
        }

        "\\dot" => {
            let (body, after_body) = extract_arg(chars, after_cmd);
            let text = process_structural(&body);
            (format!(" {} prikk ", text.trim()), after_body)
        }

        "\\tilde" => {
            let (body, after_body) = extract_arg(chars, after_cmd);
            let text = process_structural(&body);
            (format!(" {} tilde ", text.trim()), after_body)
        }

        // --- Unknown command: pass through for flat substitution phase ---
        _ => {
            (cmd, after_cmd)
        }
    }
}

/// Extract everything between \begin{env} and \end{env}, handling nesting.
fn extract_until_end(chars: &[char], pos: usize, env_name: &str) -> (String, usize) {
    let mut depth = 1;
    let mut i = pos;
    let mut content = String::new();

    while i < chars.len() && depth > 0 {
        if chars[i] == '\\' {
            let (cmd, after_cmd) = read_command_name(chars, i);
            if cmd == "\\begin" {
                let (name, after_name) = extract_arg(chars, after_cmd);
                if name == env_name {
                    depth += 1;
                }
                content.push_str(&format!("\\begin{{{}}}", name));
                i = after_name;
            } else if cmd == "\\end" {
                let (name, after_name) = extract_arg(chars, after_cmd);
                if name == env_name {
                    depth -= 1;
                    if depth == 0 {
                        return (content, after_name);
                    }
                }
                content.push_str(&format!("\\end{{{}}}", name));
                i = after_name;
            } else {
                content.push_str(&cmd);
                i = after_cmd;
            }
        } else {
            content.push(chars[i]);
            i += 1;
        }
    }

    (content, i)
}

/// Process the body of an environment into Norwegian text.
fn process_environment(env_name: &str, body: &str) -> String {
    match env_name {
        "matrix" | "pmatrix" | "bmatrix" | "vmatrix" | "Bmatrix" | "Vmatrix" => {
            process_matrix(env_name, body)
        }
        "aligned" => {
            // aligned is like a set of equations — process each line
            let lines: Vec<&str> = body.split("\\\\").collect();
            let processed: Vec<String> = lines
                .iter()
                .map(|line| {
                    let line = line.replace('&', " ");
                    process_structural(line.trim())
                })
                .filter(|s| !s.trim().is_empty())
                .collect();
            processed.join("; ")
        }
        "cases" => {
            let lines: Vec<&str> = body.split("\\\\").collect();
            let processed: Vec<String> = lines
                .iter()
                .map(|line| {
                    let parts: Vec<&str> = line.split('&').collect();
                    if parts.len() >= 2 {
                        let value = process_structural(parts[0].trim());
                        let condition = process_structural(parts[1].trim());
                        format!("{} når {}", value.trim(), condition.trim())
                    } else {
                        process_structural(line.trim())
                    }
                })
                .filter(|s| !s.trim().is_empty())
                .collect();
            format!(" tilfellene: {} ", processed.join("; "))
        }
        _ => {
            // Unknown environment — just process the body
            process_structural(body)
        }
    }
}

/// Process a matrix environment into Norwegian text.
fn process_matrix(env_name: &str, body: &str) -> String {
    let (open, close) = match env_name {
        "pmatrix" => ("(", ")"),
        "bmatrix" => ("[", "]"),
        "vmatrix" => ("|", "|"),
        _ => ("", ""),
    };

    let rows: Vec<&str> = body.split("\\\\").collect();
    let mut row_texts = Vec::new();

    for row in &rows {
        let row = row.trim();
        if row.is_empty() {
            continue;
        }
        let cols: Vec<&str> = row.split('&').collect();
        let col_texts: Vec<String> = cols
            .iter()
            .map(|c| {
                let processed = process_structural(c.trim());
                processed.trim().to_string()
            })
            .filter(|s| !s.is_empty())
            .collect();
        if !col_texts.is_empty() {
            row_texts.push(col_texts.join(", "));
        }
    }

    let matrix_text = row_texts.join("; ");
    format!(" {}matrise {}{}{} ", open, matrix_text, close, "")
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_extract_brace_group() {
        let chars: Vec<char> = "{hello}".chars().collect();
        let (content, pos) = extract_brace_group(&chars, 0);
        assert_eq!(content, "hello");
        assert_eq!(pos, 7);
    }

    #[test]
    fn test_nested_braces() {
        let chars: Vec<char> = "{a{b}c}".chars().collect();
        let (content, pos) = extract_brace_group(&chars, 0);
        assert_eq!(content, "a{b}c");
        assert_eq!(pos, 7);
    }

    #[test]
    fn test_frac() {
        let result = process_structural("\\frac{3}{4}");
        assert!(result.contains("3 over 4"));
    }

    #[test]
    fn test_nested_frac() {
        let result = process_structural("\\frac{\\frac{1}{2}}{3}");
        assert!(result.contains("1 over 2"));
        assert!(result.contains("over 3"));
    }

    #[test]
    fn test_sqrt() {
        let result = process_structural("\\sqrt{x}");
        assert!(result.contains("kvadratroten av x"));
    }

    #[test]
    fn test_sqrt_nth() {
        let result = process_structural("\\sqrt[3]{x}");
        assert!(result.contains("3-te roten av x"));
    }

    #[test]
    fn test_sum_with_bounds() {
        let result = process_structural("\\sum_{i=1}^{n}");
        assert!(result.contains("summen"));
        assert!(result.contains("fra"));
        assert!(result.contains("til"));
    }

    #[test]
    fn test_int_with_bounds() {
        let result = process_structural("\\int_{0}^{1}");
        assert!(result.contains("integralet"));
        assert!(result.contains("fra 0"));
        assert!(result.contains("til 1"));
    }

    #[test]
    fn test_operatorname() {
        let result = process_structural("\\operatorname{s i n}");
        assert!(result.contains("sin"));
    }

    #[test]
    fn test_matrix() {
        let result = process_structural("\\begin{matrix} 1 & 2 \\\\ 3 & 4 \\end{matrix}");
        assert!(result.contains("1, 2"));
        assert!(result.contains("3, 4"));
    }

    #[test]
    fn test_binom() {
        let result = process_structural("\\binom{n}{k}");
        assert!(result.contains("n over k"));
    }

    #[test]
    fn test_lim() {
        let result = process_structural("\\lim_{x \\to 0}");
        assert!(result.contains("grenseverdien"));
    }

    #[test]
    fn test_underset_lim() {
        let result = process_structural("\\underset{n \\to\\infty}{\\operatorname*{l i m}}");
        assert!(result.contains("lim"));
    }

    #[test]
    fn test_left_right() {
        let result = process_structural("\\left(x\\right)");
        assert!(result.contains("("));
        assert!(result.contains(")"));
        assert!(result.contains("x"));
    }

    #[test]
    fn test_aligned() {
        let result = process_structural("\\begin{aligned} a &= b \\\\ \\end{aligned}");
        assert!(!result.trim().is_empty());
    }
}
