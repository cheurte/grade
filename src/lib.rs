use latex::{Document, Element, PreambleElement};
use std::fs::File;
use std::io::{BufReader, Write};
use std::path::Path;
// use std::process::Output;

use calamine::{open_workbook, DataType, Reader, Xlsx};

use serde::Deserialize;

#[derive(Debug, Deserialize, Clone)]
#[serde(rename_all = "camelCase")]
pub struct ConfigXlsx {
    pub pdf_file: Vec<PdfFile>,
    pub color_text: Vec<i32>,
    pub color_tab_title: Vec<i32>,
    pub color_tab_line: Vec<i32>,
    pub margin_size: f32,
}

#[derive(Debug, Deserialize, Clone)]
#[serde(rename_all = "camelCase")]
pub struct PdfFile {
    _pdf_name: String,
    sources: Vec<String>,
    sheets_name: Vec<String>,
    products: Vec<String>,
    categories: Vec<String>,
    parameters: Vec<String>,
}

#[derive(Debug, Clone)]
pub enum AlignTab {
    C, // Center align
    R, // Right align
    L, // left align
}

#[derive(Debug)]
pub enum TabParameters {
    Parameter,
    Category,
    Product,
}

fn define_column(nb_col: usize, align: AlignTab) -> String {
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

pub fn end_line_tab(line: &mut String) -> String {
    line.push_str(" \\\\\n");
    line.to_string()
}

pub fn add_empty_rows(content: &mut String, nb_rows: usize, nb_params: usize) -> String {
    if nb_params >= nb_rows {
        return String::new();
    }
    content.push_str(&" & ".to_string().repeat(nb_rows - nb_params));
    content.to_string()
}

pub fn create_title_tabularx(title: String, nb_col: usize) -> String {
    let mut title = String::from(format!("\\rowcolor{{color_title}}{}", title));
    add_empty_rows(&mut title, nb_col, 1);
    end_line_tab(&mut title)
}

fn add_colored_line() -> String {
    String::from(
        "\\arrayrulecolor{line_color}\\hline
",
    )
}

/// reshape a 1D vector of string into a 2D vector
/// vec!["a1", "a2", "a3", "b1", "b2", "b3"]
///     -> vec![vec!["a1", "a2", "a3"], vec!["b1", "b2", "b3"]]
fn reshape_vector_by_col(single_shape_vec: Vec<String>, nb_param: usize) -> Vec<Vec<String>> {
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
fn transpose2dvec(col_vec: Vec<Vec<String>>) -> Vec<Vec<String>> {
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

/// Function to clean a 2D vector of String all values of one of the vector are
/// equal to ""
/// # Exampes
/// vec![vec!["a","b","c"],vec!["", "", ""], vec!["d", "e", "f"]]
///     -> vec![vec!["a","b","c"], vec!["d", "e", "f"]]
fn clean_vector(double_shape_vec: Vec<Vec<String>>) -> (Vec<Vec<String>>, Vec<usize>) {
    let mut resized_vec: Vec<Vec<String>> = Vec::new();
    let mut useless_col: Vec<usize> = Vec::new();
    for (i, arr) in double_shape_vec.iter().enumerate() {
        if arr.iter().all(|f| f != "") {
            resized_vec.push(arr.to_vec());
        } else {
            useless_col.push(i);
        }
    }
    (resized_vec, useless_col)
}

fn clean_content(
    parameters: &Vec<String>,
    content: &Vec<String>,
    nb_param: usize,
) -> (Vec<Vec<String>>, Vec<usize>) {
    assert_eq!(parameters.len() % nb_param, 0);
    // Cleaning and re organizing the data
    let parameters = reshape_vector_by_col(parameters.to_vec(), nb_param);
    let (mut clean_param, useless_col) = clean_vector(parameters);
    clean_param.insert(1, content.to_vec());
    (transpose2dvec(clean_param), useless_col)
}
/// Function to create the content of the tab
/// Must take a clean content to properly work
fn create_content(clean_content: Vec<Vec<String>>, nb_col: usize) -> String {
    let mut content: String = String::new();
    for line in clean_content.iter() {
        content.push_str(&line.join(" & "));
        add_empty_rows(&mut content, nb_col, line.len());
        end_line_tab(&mut content);
        content.push_str(&add_colored_line());
    }
    content = content.replace("%", "\\%");
    content = content.replace("μ", "\\(\\mu\\)");
    content = content.replace("µ", "\\(\\mu\\)");
    content = content.replace("<", "\\(<\\) ");
    content = content.replace(">", "\\(>\\) ");
    content = content.replace("m2", "\\(m^2\\)");
    content
}
fn create_parameters_tabularx(
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
    println!("{nb_col}, {param_size}");
    add_empty_rows(&mut parameters, nb_col, param_size);
    end_line_tab(&mut parameters);
    parameters
}

fn define_environment(name: String, parameters: String, content: String) -> String {
    if parameters.is_empty() {
        return format!("\\begin{{{name}}}\n{content}\n\\end{{{name}}}");
    } else {
        return format!("\\begin{{{name}}}{{{parameters}}}\n{content}\n\\end{{{name}}}");
    }
}

/// Function that reunite all the tabular creation functions
/// add to the page one centered tabular
fn create_tabularx(
    page: &mut Document,
    nb_col: usize,
    title: &String,
    parameters: &mut Vec<String>,
    general_content: &Vec<String>,
    product_values: &Vec<String>,
    nb_param: usize,
) {
    let (mut cleaned_content, useless_col) =
        clean_content(general_content, product_values, nb_param);
    let two_col_tab: bool = match cleaned_content.len() {
        0..=13 => false,
        _ => true,
    };
    // textwidth change
    let mut tabular_content: Vec<String> = Vec::new();
    let title = vec![
        "{\\textwidth}".to_string(),
        format!("{{{}}}", define_column(nb_col, AlignTab::L)),
        create_title_tabularx(title.to_string(), nb_col),
    ];

    let params = create_parameters_tabularx(parameters, &useless_col, nb_col);
    let size_content = cleaned_content.len() / 2;
    if two_col_tab {
        let first_half = cleaned_content.split_off(size_content);
        let title_in_env =
            define_environment("tabularx".to_string(), "".to_string(), title.join(""));
        let content_1st_half = vec![
            format!("{{{}}}", define_column(nb_col, AlignTab::L)),
            params.clone(),
            create_content(first_half, nb_col),
        ];
        let content_2nd_half = vec![
            format!("{{{}}}", define_column(nb_col, AlignTab::L)),
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

    // println!("{tabular_content:?}");
    let tab = Element::Environment(String::from("center"), tabular_content);

    page.push(tab);
}

/// Create the string that will be compiled.
/// This function will be depending on json files later. -> Todo
///
pub fn page_blue_print(
    page: &mut Document,
    product_name: String,
    titles: &Vec<String>,
    parameters: Vec<String>,
    general_contents: &Vec<Vec<String>>,
    product_contents: &Vec<Vec<String>>,
    nb_param: usize,
) {
    // we iterate over tabulars
    let mut general_content = general_contents.iter();
    let mut title = titles.iter();
    let mut product_content = product_contents.iter();
    let image =
        String::from("\\includegraphics[scale=0.20]{biotec}\n\\hfill\\tiny Last Updated \\today");
    let image = define_environment("flushleft".to_string(), "".to_string(), image);
    page.push(Element::UserDefined(image));

    let intro = String::from(&format!(
        "\\hspace{{1cm}}\\\\\n\\textbf{{Preliminary Data Sheed}}\\\\\n{}\\\\\n\\hspace{{1cm}}\\\\",
        product_name
    ));
    page.push(Element::UserDefined(define_environment(
        "flushleft".to_string(),
        "".to_string(),
        intro,
    )));
    for _ in 0..titles.len() {
        let mut params = parameters.clone();
        let title = title.next();
        let general_content = general_content.next();
        let product_content = product_content.next();
        let _tab = create_tabularx(
            page,
            params.len(),
            title.unwrap(),
            &mut params,
            &general_content.unwrap(),
            &product_content.unwrap(),
            nb_param,
        );
        // break;
    }

    page.push(Element::UserDefined(String::from("\\vspace*{\\fill}")));
    page.push(Element::UserDefined(String::from("{\\scriptsize
    \\textbf{Disclaimer} This information and our technical advice - whether verbal, in writing or by way of trials - are given in good faith but without warranty, and this also applies where proprietary rights of third parties are involved. Our advice does not release you from the obligation to check its validity and to test our products as to their suitability for the intended processes and uses. The application, use and processing of our products and the products manufactured by you on the basis of our technical advice are beyond our control and, therefore, entirely your own responsibility. Our products are sold in accordance with our General Conditions of Sale and Delivery \\\\ \n BIOTEC Biologische Naturverpackungen GmbH \\& Co. KG · Werner-Heisenberg-Str. 32 · D.46446 Emmerich \\hfill \\textbf{T} +49 2822 92510\\qquad \\textbf{W} biotec.de}")));
    page.push(Element::ClearPage);
}

pub fn first_page(page: &mut Document, product_names: &Vec<String>) {
    let image =
        String::from("\\includegraphics[scale=0.20]{biotec}\n\\hfill\\tiny Last Updated \\today");
    let image = define_environment("flushleft".to_string(), "".to_string(), image);
    page.push(Element::UserDefined(image));

    let mut table_of_content =
        String::from("\\hspace{1cm}\\\\\n\\textbf{Contents}\\\\\n\\hspace{5in}\\\\\n");
    let mut item_product: Vec<String> = Vec::new();
    for (i, product_name) in product_names.iter().enumerate() {
        item_product.push(format!("\\setItemnumber{{{}}}\n", i + 2));
        item_product.push(format!("\\item {}\\\\\n", product_name))
    }
    let item_product = define_environment(
        "enumerate".to_string(),
        "".to_string(),
        item_product.join(""),
    );

    table_of_content.push_str(item_product.as_str());

    page.push(Element::UserDefined(define_environment(
        "flushleft".to_string(),
        "".to_string(),
        table_of_content,
    )));
    page.push(Element::UserDefined(String::from("\\vspace*{\\fill}")));
    page.push(Element::UserDefined(String::from("{\\footnotesize
    \\textbf{Disclaimer} This information and our technical advice - whether verbal, in writing or by way of trials - are given in good faith but without warranty, and this also applies where proprietary rights of third parties are involved. Our advice does not release you from the obligation to check its validity and to test our products as to their suitability for the intended processes and uses. The application, use and processing of our products and the products manufactured by you on the basis of our technical advice are beyond our control and, therefore, entirely your own responsibility. Our products are sold in accordance with our General Conditions of Sale and Delivery.\\\\ \n BIOTEC Biologische Naturverpackungen GmbH \\& Co. KG · Werner-Heisenberg-Str. 32 · D.46446 Emmerich \\hfill \\textbf{T} +49 2822 92510\\qquad \\textbf{W} biotec.de}")));
    page.push(Element::ClearPage);
}

pub fn starting_pdf(page: &mut Document, config: &ConfigXlsx) {
    page.preamble.use_package("tabularx");
    page.preamble.use_package("xcolor");
    page.preamble.use_package("colortbl");
    page.preamble.use_package("geometry");
    page.preamble.use_package("paracol");
    page.preamble.use_package("graphicx");
    let margin: PreambleElement = PreambleElement::UserDefined(String::from(&format!(
        "\\geometry{{margin={}in}}",
        config.margin_size
    )));
    let def_color_title: PreambleElement = PreambleElement::UserDefined(String::from(&format!(
        "\\definecolor{{color_title}}{{RGB}}{{{}}}",
        config
            .color_tab_title
            .iter()
            .enumerate()
            .map(|(i, val)| {
                if i != config.color_tab_title.len() - 1 {
                    val.to_string() + ","
                } else {
                    val.to_string()
                }
            })
            .collect::<String>()
    )));
    let def_color_line: PreambleElement = PreambleElement::UserDefined(String::from(&format!(
        "\\definecolor{{line_color}}{{RGB}}{{{}}}",
        config
            .color_tab_line
            .iter()
            .enumerate()
            .map(|(i, val)| {
                if i != config.color_tab_line.len() - 1 {
                    val.to_string() + ","
                } else {
                    val.to_string()
                }
            })
            .collect::<String>()
    )));
    let def_color_font: PreambleElement = PreambleElement::UserDefined(String::from(&format!(
        "\\definecolor{{font_color}}{{RGB}}{{{}}}",
        config
            .color_text
            .iter()
            .enumerate()
            .map(|(i, val)| {
                if i != config.color_text.len() - 1 {
                    val.to_string() + ","
                } else {
                    val.to_string()
                }
            })
            .collect::<String>()
    )));
    page.preamble.author("Biotec");
    page.preamble.title("Template");
    page.preamble
        .push(margin)
        .push(def_color_title)
        .push(def_color_font)
        .push(def_color_line);

    page.preamble
        .push(PreambleElement::UserDefined(String::from(
            "\\color{font_color}",
        )));
    page.preamble
        .push(PreambleElement::UserDefined(String::from(
            "\\pagenumbering{gobble}",
        )));
    page.preamble
        .push(PreambleElement::UserDefined(String::from(
            "\\renewcommand{\\familydefault}{\\sfdefault}",
        )));
    page.preamble
        .push(PreambleElement::UserDefined(String::from(
            "\\renewcommand{\\arraystretch}{1.25}",
        )));
    page.preamble
        .push(PreambleElement::UserDefined(String::from(
            "\\graphicspath{{./resources/}}",
        )));
    page.preamble
        .push(PreambleElement::UserDefined(String::from(
            "\\newcommand\\setItemnumber[1]{\\setcounter{enumi}{\\numexpr#1-1\\relax}}",
        )));

    // page.push(Element::ClearPage);
}

/// Handle the tex creation.
/// Seperate function to handle the windows server lately -> ToDo
///
pub fn render_tex_file(rendered: String) -> std::io::Result<()> {
    let mut f = File::create("output/report.tex")?;
    write!(f, "{}", rendered)?;
    Ok(())
}

impl ConfigXlsx {
    pub fn new() -> Self {
        Self {
            pdf_file: vec![],
            color_text: vec![13, 64, 47],
            color_tab_title: vec![237, 233, 230],
            color_tab_line: vec![215, 212, 210],
            margin_size: 0.84,
        }
    }

    pub fn from(path: &str) -> Self {
        let path = Path::new(path);
        let file = File::open(path).expect("Problem with the file");
        let config: ConfigXlsx = serde_json::from_reader(file).expect("error in the reading");
        config
    }
}

impl PdfFile {
    /// Function to find and return a workbook by id
    pub fn get_workbook(&self, id_woorkbook: usize) -> Xlsx<BufReader<File>> {
        let path = self.sources.iter().nth(id_woorkbook).unwrap();
        let workbook: Xlsx<_> = open_workbook(path)
            .expect("Cannot open workbook. It is probably an error in the path of the workbook.");
        workbook
    }

    /// Function to find and return a worksheet by id
    pub fn get_worksheet(&self, id_worksheet: usize) -> &String {
        self.sheets_name.iter().nth(id_worksheet).expect(
            "Please make sur the name between the worksheet and
                the name given in the configuration file are the same.",
        )
    }
    /// Function to search the key elements of the file.
    /// Suppose to depend on a json file to know wich line or row can be read.
    pub fn search_cells_coordinates(&self, field: TabParameters) -> Vec<(usize, usize)> {
        let mut workbook = self.get_workbook(0);

        let mut output: Vec<(usize, usize)> = Vec::new();
        let field = match field {
            TabParameters::Product => &self.products,
            TabParameters::Parameter => &self.parameters,
            TabParameters::Category => &self.categories,
        };
        match workbook.worksheet_range(self.get_worksheet(0)) {
            Some(Ok(range)) => {
                for row in 0..range.get_size().0 {
                    for col in 0..range.get_size().1 {
                        let value = range.get_value((row as u32, col as u32));
                        if value != Some(&DataType::Empty) {
                            for category in field {
                                if &value.unwrap().to_string() == category {
                                    output.push((row, col));
                                }
                            }
                        }
                    }
                }
            }
            Some(Err(e)) => println!("{e:?}"),
            None => println!("Sheets name unknown. Maybe check the name in the config file"),
        }
        println!("****************");
        println!("{}", output.len());
        println!("{}", field.len());
        println!("****************");
        assert_eq!(output.len(), field.len());
        output
    }

    /// Search the range of categories, to know where to stop
    /// Take the beginning corrdinates of categories, and return the end coordinates
    pub fn get_parameters_range(
        &self,
        categories_coord: &Vec<(usize, usize)>,
    ) -> Vec<(usize, usize)> {
        // let sheet = self.sheets_name.iter().nth(0).unwrap();
        let mut end_categories: Vec<(usize, usize)> = vec![];
        let mut workbook = self.get_workbook(0);

        match workbook.worksheet_range(&self.get_worksheet(0)) {
            Some(Ok(range)) => {
                for (category_row, category_col) in categories_coord.into_iter() {
                    let mut col = category_col + 1;
                    loop {
                        if range.get_value((*category_row as u32, col as u32))
                            != Some(&DataType::Empty)
                            || col > range.get_size().1
                        {
                            break;
                        }
                        col += 1;
                    }
                    end_categories.push((*category_row, col - 1));
                }
            }
            Some(Err(e)) => println!("{e:?}"),
            None => println!("Sheets name unknown. Maybe check the name in the config file"),
        }
        end_categories
    }
    /// Return the tab titles
    pub fn get_values_at(&self, begin_categories: &Vec<(usize, usize)>) -> Vec<String> {
        let mut workbook = self.get_workbook(0);
        let mut output: Vec<String> = vec![];

        match workbook.worksheet_range(self.get_worksheet(0)) {
            Some(Ok(range)) => {
                for category in begin_categories {
                    let (a, b) = category;
                    // println!("{:?}", range.get_value((*a as u32, *b as u32)));
                    output.push(range.get_value((*a as u32, *b as u32)).unwrap().to_string())
                }
            }
            Some(Err(e)) => println!("{e:?}"),
            None => println!("Sheets name unknown. Maybe check the name in the config file"),
        }
        output
    }

    pub fn get_parameters_by_id(
        &self,
        start_categ_coord: &Vec<(usize, usize)>,
        end_categ_coord: &Vec<(usize, usize)>,
        id_line: Vec<usize>,
    ) -> Vec<Vec<String>> {
        assert_eq!(start_categ_coord.len(), end_categ_coord.len());

        let mut workbook = self.get_workbook(0);
        let mut output: Vec<Vec<String>> = vec![];

        match workbook.worksheet_range(self.get_worksheet(0)) {
            Some(Ok(range)) => {
                let it = start_categ_coord.iter().zip(end_categ_coord.iter());
                for (_, (start_coord, end_coord)) in it.enumerate() {
                    let mut parameters: Vec<String> = vec![];
                    for col in start_coord.1..end_coord.1 + 1 {
                        for line in id_line.iter() {
                            parameters.push(
                                range
                                    .get_value(((start_coord.0 + line) as u32, col as u32))
                                    .unwrap()
                                    .to_string(),
                            )
                        }
                    }
                    output.push(parameters);
                }
            }
            Some(Err(e)) => println!("{e:?}"),
            None => println!("Sheets name unknown. Maybe check the name in the config file"),
        }
        output
    }

    pub fn get_values_from_parameters(
        &self,
        product_coordinates: (usize, usize),
        start_categ_coord: &Vec<(usize, usize)>,
        end_categ_coord: &Vec<(usize, usize)>,
    ) -> Vec<Vec<String>> {
        let mut workbook = self.get_workbook(0);
        let mut out: Vec<Vec<String>> = Vec::new();
        match workbook.worksheet_range(self.get_worksheet(0)) {
            Some(Ok(range)) => {
                for param in 0..start_categ_coord.len() {
                    let mut parameters: Vec<String> = vec![];
                    for y in start_categ_coord.iter().nth(param).unwrap().1
                        ..end_categ_coord.iter().nth(param).unwrap().1 + 1
                    {
                        let x = product_coordinates.0;
                        parameters.push(range.get_value((x as u32, y as u32)).unwrap().to_string());
                    }
                    out.push(parameters);
                }
            }
            Some(Err(e)) => println!("{e:?}"),
            None => println!("Sheets name unknown. Maybe check the name in the config file"),
        }
        out
    }
}
