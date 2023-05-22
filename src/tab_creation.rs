use crate::AlignTab;

use latex::{Document, Element};
/// Define how shoud be aligne the columns
pub fn define_column(nb_col: usize, align: &AlignTab) -> String {
    let mut column_definition: String = String::new();
    let align = match align {
        AlignTab::C => String::from("c"),
        AlignTab::L => String::from("l"),
        AlignTab::R => String::from("r"),
    };
    column_definition.push('X');
    for _ in 0..nb_col - 1 {
        column_definition.push_str(&format!(" {} ", align)[..]);
    }
    column_definition
}

/// Add an latex end on line.
pub fn end_line_tab(line: &mut String) -> String {
    line.push_str(" \\\\\n");
    line.to_string()
}

/// Add empty & to a row to artificially increase the size of the row. not used
/// anymore.
pub fn add_empty_rows(content: &mut String, nb_rows: usize, nb_params: usize) -> String {
    if nb_params >= nb_rows {
        return String::new();
    }
    content.push_str(&" & ".to_string().repeat(nb_rows - nb_params));
    content.to_string()
}

/// Create the title of a tabular, very specific
pub fn create_title_tabularx(title: String, nb_col: usize) -> String {
    let mut title = String::from(format!("\\rowcolor{{color_title}}{}", title));
    add_empty_rows(&mut title, nb_col, 1);
    end_line_tab(&mut title)
}

/// Function that add a colored line to a tab, very specific.
pub fn add_colored_line() -> String {
    String::from("\\arrayrulecolor{line_color}\\hline\n")
}

/// reshape a 1D vector of string into a 2D vector
/// vec!["a1", "a2", "a3", "b1", "b2", "b3"]
///     -> vec![vec!["a1", "a2", "a3"], vec!["b1", "b2", "b3"]]
pub fn reshape_vector_by_col(single_shape_vec: Vec<String>, nb_param: usize) -> Vec<Vec<String>> {
    let mut reshaped_param: Vec<Vec<String>> = Vec::new();
    for i in 0..(single_shape_vec.len() / (single_shape_vec.len() / nb_param)) {
        let mut buff_vec: Vec<String> = Vec::new();
        for j in (i..single_shape_vec.len()).step_by(nb_param) {
            buff_vec.push(single_shape_vec.iter().nth(j).unwrap().to_string());
        }
        reshaped_param.push(buff_vec);
    }
    reshaped_param
}

/// Transpose a 2D vec of String
/// vec![vec!["a1", "a2", "a3"],vec!["b1", "b2", "b3"]]
///     -> vec![["a1", "b1"], vec!["a2", "b2"], vec!["a3", "b3"]]
pub fn transpose2dvec(col_vec: Vec<Vec<String>>) -> Vec<Vec<String>> {
    let size_col = col_vec.len();
    let size_row = col_vec.first().expect("error, no values are here").len();
    // let test = &col_vec[..];
    let mut output: Vec<Vec<String>> = Vec::new();
    for i in 0..size_row {
        let mut buff_vec: Vec<String> = Vec::new();
        for j in 0..size_col {
            buff_vec.push(col_vec[..][j][i].to_string());
        }
        if buff_vec.iter().nth(1) != Some(&String::from("")) {
            output.push(buff_vec);
        }
    }
    output
}

/// Function to clean a 2D vector of String when all values of one of the vector are
/// equal to ""
/// # Exampes
/// vec![vec!["a","b","c"],vec!["", "", ""], vec!["d", "e", "f"]]
///     -> vec![vec!["a","b","c"], vec!["d", "e", "f"]]
pub fn clean_vector(double_shape_vec: Vec<Vec<String>>) -> (Vec<Vec<String>>, Vec<usize>) {
    let mut resized_vec: Vec<Vec<String>> = Vec::new();
    let mut useless_col: Vec<usize> = Vec::new();
    println!("{double_shape_vec:?}");
    for (i, arr) in double_shape_vec.iter().enumerate() {
        if arr.iter().any(|f| !f.is_empty()) {
            resized_vec.push(arr.to_vec());
        } else {
            useless_col.push(i);
        }
    }
    (resized_vec, useless_col)
}

/// Function that call all the cleaning function of the data
/// The main goal is not to have any empty row
/// n/a is not considered as an empty row
pub fn clean_content(
    parameters: &Vec<String>,
    content: &Vec<String>,
    nb_param: usize,
) -> (Vec<Vec<String>>, Vec<usize>) {
    assert_eq!(parameters.len() % nb_param, 0);
    // Cleaning and re organizing the data
    let parameters = reshape_vector_by_col(parameters.to_vec(), nb_param);
    // println!("{parameters:?}");
    let (mut clean_param, useless_col) = clean_vector(parameters);
    clean_param.insert(1, content.to_vec());
    (transpose2dvec(clean_param), useless_col)
}

/// Function to create the content of the tab
/// Must take a clean content to properly work
pub fn create_content(clean_content: Vec<Vec<String>>, nb_col: usize) -> String {
    let mut content: String = String::new();
    for line in clean_content.iter() {
        content.push_str(&line.join(" & "));
        add_empty_rows(&mut content, nb_col, line.len());
        end_line_tab(&mut content);
        content.push_str(&add_colored_line());
    }
    // just for now, to be improved later
    content = content.replace("%", "\\%");
    content = content.replace("μ", "\\(\\mu\\)");
    content = content.replace("µ", "\\(\\mu\\)");
    content = content.replace("<", "\\(<\\) ");
    content = content.replace(">", "\\(>\\) ");
    content = content.replace("m2", "\\(m^2\\)");
    content
}

/// Function that return a line of a tabular with every element in bold (used
/// here for the parameters names).
pub fn create_parameters_tabularx(
    parameters: &mut Vec<String>,
    useless_col: &Vec<usize>,
    nb_col: usize,
) -> String {
    parameters
        .iter_mut()
        .for_each(|f| *f = format!("\\textbf{{{}}}", f));
    let mut j = 0;
    for i in useless_col.iter() {
        parameters.remove(*i - j);
        j += 1;
    }

    parameters.insert(1, String::from("\\textbf{Target Value}"));
    let param_size = parameters.len();
    let mut parameters = parameters.join(" & ");
    add_empty_rows(&mut parameters, nb_col, param_size);
    end_line_tab(&mut parameters);
    parameters
}

/// Funtion to add some content in a string into an Environment
/// some parameters can be added. It is very simple thought, only one parameter
/// can be added.
pub fn define_environment(name: String, parameters: String, content: String) -> String {
    if parameters.is_empty() {
        return format!("\\begin{{{name}}}\n{content}\n\\end{{{name}}}");
    } else {
        return format!("\\begin{{{name}}}{{{parameters}}}\n{content}\n\\end{{{name}}}");
    }
}

pub fn find_larger_rows(content: &Vec<Vec<String>>) -> Vec<usize> {
    let mut indices_bigger_row: Vec<usize> = Vec::new();
    content.iter().enumerate().for_each(|(i, e)| {
        if e.get(0).unwrap().len() > 26 {
            indices_bigger_row.push(i)
        }
    });
    indices_bigger_row
}

pub fn add_rule_row(content: &mut Vec<Vec<String>>, indices: Vec<usize>) {
    for (i, value) in content.iter_mut().enumerate() {
        if indices.iter().find(|v| **v == i).is_some() {
            value
                .iter_mut()
                .nth(0)
                .unwrap()
                .push_str(" \\rule{80pt}{0pt}");
        }
    }
}

/// Function that reunite all the tabular creation functions
/// add to the page one centered tabular
pub fn create_tabularx(
    page: &mut Document,
    nb_col: usize,
    title: &String,
    parameters: &mut Vec<String>,
    general_content: &Vec<String>,
    product_values: &Vec<String>,
    nb_param: usize,
    align: &AlignTab,
) {
    // println!("{general_content:?}");
    let (mut cleaned_content, useless_col) =
        clean_content(general_content, product_values, nb_param);
    // println!("{cleaned_content:?}");
    let two_col_tab: bool = match cleaned_content.len() {
        0..=13 => false,
        _ => true,
    };
    // textwidth change
    let mut tabular_content: Vec<String> = Vec::new();
    let title = vec![
        "{\\textwidth}".to_string(),
        format!("{{{}}}", define_column(nb_col, &align)),
        create_title_tabularx(title.to_string(), nb_col),
    ];

    let params = create_parameters_tabularx(parameters, &useless_col, nb_col);
    let size_content = cleaned_content.len() / 2;
    if two_col_tab {
        let mut first_half = cleaned_content.split_off(size_content);

        let indices_first_hal = find_larger_rows(&first_half);
        let indices_sec_half = find_larger_rows(&cleaned_content);

        add_rule_row(&mut first_half, indices_sec_half);
        add_rule_row(&mut cleaned_content, indices_first_hal);

        let title_in_env =
            define_environment("tabularx".to_string(), "".to_string(), title.join(""));

        let content_1st_half = vec![
            format!("{{{}}}", define_column(nb_col, &align)),
            params.clone(),
            create_content(first_half, nb_col),
        ];
        let content_2nd_half = vec![
            format!("{{{}}}", define_column(nb_col, &align)),
            params.clone(),
            create_content(cleaned_content, nb_col),
        ];

        let left_tab = define_environment(
            "tabularx".to_string(),
            "0.5\\textwidth".to_string(),
            content_1st_half.join(""),
        );
        let right_tab = define_environment(
            "tabularx".to_string(),
            "0.5\\textwidth".to_string(),
            content_2nd_half.join(""),
        );

        let switch_col = String::from("\\switchcolumn");
        tabular_content.push(
            vec![
                title_in_env,
                define_environment(
                    "paracol".to_string(),
                    "2".to_string(),
                    vec![left_tab, switch_col, right_tab].join(""),
                ),
            ]
            .join(""),
        );
    } else {
        let mut single_tab: Vec<String> = vec![title.join("")];
        single_tab.push(params.clone());
        // println!("ok : {cleaned_content:?}");
        single_tab.push(create_content(cleaned_content, nb_col));
        tabular_content.push(define_environment(
            "tabularx".to_string(),
            "".to_string(),
            single_tab.join(""),
        ));
    }

    let tab = Element::Environment(String::from("center"), tabular_content);

    page.push(tab);
}
