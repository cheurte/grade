use latex::{print, Document, DocumentClass, Element, PreambleElement};
// use serde::de::value;
use std::fs::File;
use std::io::{BufReader, Write};
use std::path::Path;

// use latexcompile::{LatexCompiler, LatexInput};
// use std::collections::HashMap;
// use std::fs::write;

use calamine::{open_workbook, DataType, Reader, Xlsx};

use serde::Deserialize;

#[derive(Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ConfigXlsx {
    pub pdf_file: Vec<PdfFile>,
    pub color_text: Vec<i32>,
    pub color_tab_title: Vec<i32>,
    pub color_tab_line: Vec<i32>,
}

#[derive(Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct PdfFile {
    _pdf_name: String,
    sources: Vec<String>,
    sheets_name: Vec<String>,
    products: Vec<String>,
    categories: Vec<String>,
}

#[derive(Debug, Clone)]
pub enum AlignTab {
    C,
    R,
    L,
}

fn define_column(nb_col: usize, align: AlignTab) -> String {
    let mut column_definition: String = String::new();
    let align = match align {
        AlignTab::C => String::from("c"),
        AlignTab::L => String::from("l"),
        AlignTab::R => String::from("r"),
    };
    for _ in 0..nb_col {
        column_definition.push_str(&format!("X {} ", align)[..]);
    }
    column_definition
}

pub fn end_line_tab(line: &mut String) -> String {
    line.push_str("\\\\");
    line.to_string()
}

pub fn create_title_tabularx(color: String, title: String) -> String {
    let mut title = String::from(format!("\\rowcolor{{{}}}{}", color, title));
    end_line_tab(&mut title)
}

/// Function to create the sub title
/// missing the &
pub fn create_sub_title_tabularx(sub_title: Vec<String>) -> String {
    let mut buff_sub_title = String::new();
    for (i, param) in sub_title.iter().enumerate() {
        buff_sub_title.push_str(&format!(" \\textbf{{{}}} ", param));
        if i != sub_title.len() - 1 {
            buff_sub_title.push('&')
        }
    }
    end_line_tab(&mut buff_sub_title)
}

/// Create the content of the tab and return it as the form of String.
/// To add :
///     - The line and the color of the line.
///     - The & missing
/// Much later, the posibility to put everything into two columns
///
pub fn create_content_tabularx(content: Vec<Vec<String>>) -> String {
    let mut buff_sub_content = String::new();
    for line in content.iter() {
        buff_sub_content.push_str(&format!("{}", line.join("&")));
        end_line_tab(&mut buff_sub_content);
    }
    buff_sub_content
}

pub fn create_tabularx() -> Element {
    // let title: String = String:
    let tab = Element::Environment(
        String::from("tabularx"),
        vec![
            "{\\textwidth}".to_string(),
            format!("{{{}}}", define_column(6, AlignTab::L)),
            create_title_tabularx(String::from("color_title"), String::from("This is a title")),
            create_sub_title_tabularx(vec![String::from("param 1"), String::from("param 2")]),
            create_content_tabularx(vec![
                vec![String::from("value 1"), String::from("Value 2")],
                vec![String::from("value 1"), String::from("value 2")],
            ]),
        ],
    );
    tab
}

/// Create the string that will be compiled.
/// This function will be depending on json files later. -> Todo
///
pub fn page_blue_print(config: ConfigXlsx) -> String {
    let mut doc = Document::new(DocumentClass::Article);
    doc.preamble.use_package("tabularx");
    doc.preamble.use_package("xcolor");
    doc.preamble.use_package("colortbl");
    doc.preamble.use_package("geometry");
    let margin: PreambleElement =
        PreambleElement::UserDefined(String::from("\\geometry{margin=0.84in}"));
    // let _t =

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
    doc.preamble.author("corentin");
    doc.preamble.title("Template document");
    doc.preamble
        .push(margin)
        .push(def_color_title)
        .push(def_color_font)
        .push(def_color_line);

    doc.push(Element::TitlePage).push(Element::ClearPage);

    doc.push(create_tabularx());
    // for tab in tabs.into_iter() {
    //     doc.push(&begin_section_wo_param(String::from("center"), tab)[..]);
    // }
    // doc.push(&begin_section_wo_param("center".to_string())[..]);

    println!("{doc:?}");

    print(&doc).unwrap()
}

/// Handle the tex creation.
/// Seperate function to handle the windows server lately -> ToDo
///
pub fn render_tex_file(rendered: String) -> std::io::Result<()> {
    let mut f = File::create("output/report.tex")?;
    write!(f, "{}", rendered)?;
    Ok(())
}

/// Create the pdf file by creating a clean environment and executing the
/// compiler in the assets directory. Tested only on linux distribution yet
/// with latexmk compiler.
///
/// Again, to be tested on a windows server -> Todo
///
// pub fn create_pdf() -> std::io::Result<()> {
//     // let mut dict = HashMap::new();
//     // dict.insert("test".into(), "Minimal".into());
//     // // provide the folder where the file for latex compiler are found
//     // let input = LatexInput::from("assets");
//     // // create a new clean compiler enviroment and the compiler wrapper
//     // let compiler = LatexCompiler::new(dict).unwrap();
//     // // run the underlying pdflatex or whatever
//     // let result = compiler.run("assets/test.tex", &input).unwrap();
//
//     // // copy the file into the working directory
//     // let output = ::std::env::current_dir()
//     //     .expect("problem with current dir ??")
//     //     .join("out.pdf");
//     //
//     // assert!(write(output, result).is_ok());
//     Ok(())
// }

impl ConfigXlsx {
    pub fn new() -> Self {
        Self {
            pdf_file: vec![],
            color_text: vec![13, 64, 47],
            color_tab_title: vec![237, 233, 230],
            color_tab_line: vec![215, 212, 210],
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
    pub fn get_workbook(&self, id_woorkbook: usize) -> Xlsx<BufReader<File>> {
        let path = self.sources.iter().nth(id_woorkbook).unwrap();
        let workbook: Xlsx<_> = open_workbook(path)
            .expect("Cannot open workbook. It is probably an error in the path of the workbook.");
        workbook
    }

    pub fn get_worksheet(&self, id_worksheet: usize) -> &String {
        self.sheets_name.iter().nth(id_worksheet).expect(
            "Please make sur the name between the worksheet and
                the name given in the configuration file are the same.",
        )
    }
    /// Function to search the key elements of the file.
    /// Suppose to depend on a json file to know wich line or row can be reed.
    pub fn get_prod_categ_coordinates(&self) -> (Vec<(usize, usize)>, Vec<(usize, usize)>) {
        // To be changed to handle errors
        let mut workbook = self.get_workbook(0);

        let mut categories_coord: Vec<(usize, usize)> = vec![];
        let mut products_coord: Vec<(usize, usize)> = vec![];

        match workbook.worksheet_range(self.get_worksheet(0)) {
            Some(Ok(range)) => {
                for row in 0..range.get_size().0 {
                    for col in 0..range.get_size().1 {
                        let value = range.get_value((row as u32, col as u32));
                        if value != Some(&DataType::Empty) {
                            for category in &self.categories {
                                if &value.unwrap().to_string() == category {
                                    categories_coord.push((row, col));
                                }
                            }
                            for product in &self.products {
                                if &value.unwrap().to_string() == product {
                                    products_coord.push((row, col));
                                }
                            }
                        }
                    }
                }
            }
            Some(Err(e)) => println!("{e:?}"),
            None => println!("Sheets name unknown. Maybe check the name in the config file"),
        }
        (categories_coord, products_coord)
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
    /// Get parameters names
    /// Take the indices of values to get them.
    pub fn get_parameters_name(
        &self,
        start_categ_coord: &Vec<(usize, usize)>,
        end_categ_coord: &Vec<(usize, usize)>,
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
                        parameters.push(
                            range
                                .get_value(((start_coord.0 + 1) as u32, col as u32))
                                .unwrap()
                                .to_string(),
                        )
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
        // return vec![vec![]];
        out
    }
}
