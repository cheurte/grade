use latex::{print, Document, Element, PreambleElement};
use std::default::Default;
use std::error::Error;
use std::fs::File;
use std::io::{BufReader, Write};
use std::path::Path;

use calamine::{open_workbook, DataType, Reader, Xlsx};

use serde::Deserialize;

mod tab_creation;

#[derive(Debug, Deserialize, Clone)]
#[serde(rename_all = "camelCase")]
pub struct ConfigXlsx {
    pub pdf_file: Vec<PdfFile>,
    pub color_text: Vec<i32>,
    pub color_tab_title: Vec<i32>,
    pub color_tab_line: Vec<i32>,
    pub margin_size: f32,
    pub alignment_tabular: String,
}

#[derive(Debug, Deserialize, Clone)]
#[serde(rename_all = "camelCase")]
pub struct PdfFile {
    pdf_name: String,
    pub output: String,
    source: String,
    worksheet: String,
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

#[derive(Debug, Clone)]
pub struct Config {
    config_path: String,
}

impl Config {
    pub fn new(args: &[String]) -> Result<Config, &'static str> {
        println!("{:?}", args);
        if args.len() < 2 {
            return Err("Not enough argument");
        }

        Ok(Config {
            config_path: args[1].clone(),
        })
    }

    pub fn get_config_path(&self) -> &String {
        &self.config_path
    }
}

impl Default for Config {
    fn default() -> Self {
        Self {
            config_path: String::from("config/config_source.json"),
        }
    }
}

/// Handle the tex creation.
/// Seperate function to handle the windows server lately -> ToDo
///
pub fn render_tex_file(rendered: String, pdf_name: String) -> std::io::Result<()> {
    let mut f = File::create(&format!("output/{}.tex", pdf_name))?;
    write!(f, "{}", rendered)?;
    Ok(())
}

impl Default for ConfigXlsx {
    fn default() -> Self {
        Self {
            pdf_file: vec![PdfFile::default()],
            color_text: Vec::from([13, 64, 47]),
            color_tab_title: Vec::from([237, 233, 230]),
            color_tab_line: Vec::from([215, 212, 210]),
            margin_size: 0.80,
            alignment_tabular: String::from("left"),
        }
    }
}

/// Implementation of a config file.
impl ConfigXlsx {
    pub fn new() -> Self {
        Self {
            pdf_file: Vec::new(),
            color_text: Vec::new(),
            color_tab_title: Vec::new(),
            color_tab_line: Vec::new(),
            margin_size: 0.84,
            alignment_tabular: String::from("left"),
        }
    }

    pub fn is_empty(self) -> bool {
        self.alignment_tabular.is_empty()
            && self.color_tab_line.is_empty()
            && self.color_tab_title.is_empty()
            && self.color_text.is_empty()
            && self.pdf_file.is_empty()
    }

    /// from a path
    pub fn from(path: &str) -> Result<Self, Box<dyn std::error::Error>> {
        let path = Path::new(path);
        let file = File::open(path)?;
        let config: ConfigXlsx = serde_json::from_reader(file)?;
        Ok(config)
    }

    /// To define all the preamble element of the page.
    /// All the key element are in the config file
    ///
    pub fn starting_pdf(&self, page: &mut Document) {
        page.preamble.use_package("tabularx");
        page.preamble.use_package("xcolor");
        page.preamble.use_package("colortbl");
        page.preamble.use_package("geometry");
        page.preamble.use_package("paracol");
        page.preamble.use_package("graphicx");
        let margin: PreambleElement = PreambleElement::UserDefined(String::from(&format!(
            "\\geometry{{margin={}in}}",
            self.margin_size
        )));
        let def_color_title: PreambleElement =
            PreambleElement::UserDefined(String::from(&format!(
                "\\definecolor{{color_title}}{{RGB}}{{{}}}",
                self.color_tab_title
                    .iter()
                    .enumerate()
                    .map(|(i, val)| {
                        if i != self.color_tab_title.len() - 1 {
                            val.to_string() + ","
                        } else {
                            val.to_string()
                        }
                    })
                    .collect::<String>()
            )));
        let def_color_line: PreambleElement = PreambleElement::UserDefined(String::from(&format!(
            "\\definecolor{{line_color}}{{RGB}}{{{}}}",
            self.color_tab_line
                .iter()
                .enumerate()
                .map(|(i, val)| {
                    if i != self.color_tab_line.len() - 1 {
                        val.to_string() + ","
                    } else {
                        val.to_string()
                    }
                })
                .collect::<String>()
        )));
        let def_color_font: PreambleElement = PreambleElement::UserDefined(String::from(&format!(
            "\\definecolor{{font_color}}{{RGB}}{{{}}}",
            self.color_text
                .iter()
                .enumerate()
                .map(|(i, val)| {
                    if i != self.color_text.len() - 1 {
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
    }

    /// Define the first page of the document
    /// We find on it only the names of the products
    pub fn first_page(&self, page: &mut Document, product_names: &Vec<String>) {
        let image = String::from(
            "\\includegraphics[scale=0.20]{biotec}\n\\hfill\\tiny Last Updated \\today",
        );
        let image =
            tab_creation::define_environment("flushleft".to_string(), "".to_string(), image);
        page.push(Element::UserDefined(image));

        let mut table_of_content =
            String::from("\\hspace{1cm}\\\\\n\\textbf{Contents}\\\\\n\\hspace{5in}\\\\\n");
        let mut item_product: Vec<String> = Vec::new();
        for (i, product_name) in product_names.iter().enumerate() {
            item_product.push(format!("\\setItemnumber{{{}}}\n", i + 2));
            item_product.push(format!("\\item {}\\\\\n", product_name))
        }
        let item_product = tab_creation::define_environment(
            "enumerate".to_string(),
            "".to_string(),
            item_product.join(""),
        );

        table_of_content.push_str(item_product.as_str());

        page.push(Element::UserDefined(tab_creation::define_environment(
            "flushleft".to_string(),
            "".to_string(),
            table_of_content,
        )));
        page.push(Element::UserDefined(String::from("\\vspace*{\\fill}")));
        page.push(Element::UserDefined(String::from("{\\scriptsize
        \\textbf{Disclaimer} This information and our technical advice - whether verbal, in writing or by way of trials - are given in good faith but without warranty, and this also applies where proprietary rights of third parties are involved. Our advice does not release you from the obligation to check its validity and to test our products as to their suitability for the intended processes and uses. The application, use and processing of our products and the products manufactured by you on the basis of our technical advice are beyond our control and, therefore, entirely your own responsibility. Our products are sold in accordance with our General Conditions of Sale and Delivery.\\\\ \n BIOTEC Biologische Naturverpackungen GmbH \\& Co. KG · Werner-Heisenberg-Str. 32 · D.46446 Emmerich \\hfill \\textbf{T} +49 2822 92510\\qquad \\textbf{W} biotec.de}")));
        page.push(Element::ClearPage);
    }

    /// Create the string that will be compiled.
    /// This function will be depending on json files later. -> Todo
    ///
    pub fn page_blue_print(
        &self,
        page: &mut Document,
        product_name: String,
        titles: &Vec<String>,
        parameters: Vec<String>,
        general_contents: &Vec<Vec<String>>,
        product_contents: &Vec<Vec<String>>,
        nb_param: usize,
    ) -> Option<()> {
        // we iterate over tabulars
        let mut general_content = general_contents.iter();
        let mut title = titles.iter();
        let mut product_content = product_contents.iter();
        let image = String::from(
            "\\includegraphics[scale=0.20]{biotec}\n\\hfill\\tiny Last Updated \\today",
        );
        let image =
            tab_creation::define_environment("flushleft".to_string(), "".to_string(), image);
        page.push(Element::UserDefined(image));

        let intro = String::from(&format!(
        "\\hspace{{1cm}}\\\\\n\\textbf{{Preliminary Data Sheed}}\\\\\n{}\\\\\n\\hspace{{1cm}}\\\\",
        product_name
        ));
        page.push(Element::UserDefined(tab_creation::define_environment(
            "flushleft".to_string(),
            "".to_string(),
            intro,
        )));
        let align = match self.alignment_tabular.as_str() {
            "left" => AlignTab::L,
            "right" => AlignTab::R,
            "center" => AlignTab::C,
            _ => AlignTab::L,
        };

        for _ in 0..titles.len() {
            let mut params = parameters.clone();
            let title = title.next();
            let general_content = general_content.next();
            let product_content = product_content.next();
            let _tab = tab_creation::create_tabularx(
                page,
                params.len(),
                title?,
                &mut params,
                &general_content?.to_vec(),
                &product_content?.to_vec(),
                nb_param,
                &align,
            );
            // break;
        }

        page.push(Element::UserDefined(String::from("\\vspace*{\\fill}")));
        page.push(Element::UserDefined(String::from("{\\scriptsize
        \\textbf{Disclaimer} This information and our technical advice - whether verbal, in writing or by way of trials - are given in good faith but without warranty, and this also applies where proprietary rights of third parties are involved. Our advice does not release you from the obligation to check its validity and to test our products as to their suitability for the intended processes and uses. The application, use and processing of our products and the products manufactured by you on the basis of our technical advice are beyond our control and, therefore, entirely your own responsibility. Our products are sold in accordance with our General Conditions of Sale and Delivery \\\\ \n BIOTEC Biologische Naturverpackungen GmbH \\& Co. KG · Werner-Heisenberg-Str. 32 · D.46446 Emmerich \\hfill \\textbf{T} +49 2822 92510\\qquad \\textbf{W} biotec.de}")));
        page.push(Element::ClearPage);
        Some(())
    }
}

impl Default for PdfFile {
    fn default() -> Self {
        Self {
            pdf_name: String::from("Default File"),
            output: String::from("output/"),
            source: Config::default().get_config_path().to_string(),
            worksheet: String::from("Master - Rigid Overview "),
            products: Vec::new(),
            categories: Vec::new(),
            parameters: Vec::new(),
        }
    }
}

impl PdfFile {
    pub fn new() -> Self {
        Self {
            pdf_name: String::new(),
            output: String::new(),
            source: String::new(),
            worksheet: String::new(),
            products: Vec::new(),
            categories: Vec::new(),
            parameters: Vec::new(),
        }
    }
    pub fn is_empty(self) -> bool {
        self.pdf_name.is_empty()
            && self.output.is_empty()
            && self.source.is_empty()
            && self.worksheet.is_empty()
            && self.products.is_empty()
            && self.categories.is_empty()
            && self.parameters.is_empty()
    }

    // Function to return a workbook
    pub fn get_workbook(&self) -> Result<Xlsx<BufReader<File>>, Box<dyn Error>> {
        let workbook: Xlsx<_> = open_workbook(&self.source)?;
        Ok(workbook)
    }

    pub fn search_cells_coordinates(&self, field: TabParameters) -> Option<Vec<(usize, usize)>> {
        let mut workbook = self.get_workbook().ok()?;
        let mut output: Vec<(usize, usize)> = Vec::new();
        let field = match field {
            TabParameters::Product => &self.products,
            TabParameters::Parameter => &self.parameters,
            TabParameters::Category => &self.categories,
        };
        match workbook.worksheet_range(&self.worksheet) {
            Some(Ok(range)) => {
                for row in 0..range.get_size().0 {
                    for col in 0..range.get_size().1 {
                        let value = range.get_value((row as u32, col as u32));
                        if value != Some(&DataType::Empty) {
                            for category in field {
                                if &value?.to_string() == category {
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
        assert_eq!(output.len(), field.len());
        Some(output)
    }

    /// Search the range of categories, to know where to stop
    /// Take the beginning corrdinates of categories, and return the end coordinates
    pub fn get_parameters_range(
        &self,
        categories_coord: &Vec<(usize, usize)>,
    ) -> Option<Vec<(usize, usize)>> {
        let mut end_categories: Vec<(usize, usize)> = vec![];
        let mut workbook = self.get_workbook().ok()?;

        match workbook.worksheet_range(&self.worksheet) {
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
        Some(end_categories)
    }
    /// Return the values at a given coordinates
    pub fn get_values_at(&self, begin_categories: &Vec<(usize, usize)>) -> Option<Vec<String>> {
        let mut workbook = self.get_workbook().ok()?;
        let mut output: Vec<String> = vec![];

        match workbook.worksheet_range(&self.worksheet) {
            Some(Ok(range)) => {
                for category in begin_categories {
                    let (a, b) = category;
                    // println!("{:?}", range.get_value((*a as u32, *b as u32)));
                    output.push(range.get_value((*a as u32, *b as u32))?.to_string())
                }
            }
            Some(Err(e)) => println!("{e:?}"),
            None => println!("Sheets name unknown. Maybe check the name in the config file"),
        }
        Some(output)
    }

    pub fn get_parameters_by_id(
        &self,
        start_categ_coord: &Vec<(usize, usize)>,
        end_categ_coord: &Vec<(usize, usize)>,
        id_line: Vec<usize>,
    ) -> Option<Vec<Vec<String>>> {
        assert_eq!(start_categ_coord.len(), end_categ_coord.len());

        let mut workbook = self.get_workbook().ok()?;
        let mut output: Vec<Vec<String>> = vec![];

        match workbook.worksheet_range(&self.worksheet) {
            Some(Ok(range)) => {
                let it = start_categ_coord.iter().zip(end_categ_coord.iter());
                for (_, (start_coord, end_coord)) in it.enumerate() {
                    let mut parameters: Vec<String> = vec![];
                    for col in start_coord.1..end_coord.1 + 1 {
                        for line in id_line.iter() {
                            parameters.push(
                                range
                                    .get_value(((start_coord.0 + line) as u32, col as u32))?
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
        Some(output)
    }

    pub fn get_values_from_parameters(
        &self,
        product_coordinates: (usize, usize),
        start_categ_coord: &Vec<(usize, usize)>,
        end_categ_coord: &Vec<(usize, usize)>,
    ) -> Option<Vec<Vec<String>>> {
        let mut workbook = self.get_workbook().ok()?;
        let mut out: Vec<Vec<String>> = Vec::new();
        match workbook.worksheet_range(&self.worksheet) {
            Some(Ok(range)) => {
                for param in 0..start_categ_coord.len() {
                    let mut parameters: Vec<String> = vec![];
                    for y in start_categ_coord.iter().nth(param)?.1
                        ..end_categ_coord.iter().nth(param)?.1 + 1
                    {
                        let x = product_coordinates.0;
                        parameters.push(range.get_value((x as u32, y as u32))?.to_string());
                    }
                    out.push(parameters);
                }
            }
            Some(Err(e)) => println!("{e:?}"),
            None => println!("Sheets name unknown. Maybe check the name in the config file"),
        }
        Some(out)
    }

    /// create and render pdf
    pub fn create_and_render(&self, page: Document) -> Result<(), Box<dyn std::error::Error>> {
        let render = print(&page)?;
        match render_tex_file(render, self.pdf_name.clone()) {
            Ok(_) => println!("rendered completed"),
            Err(e) => println!("{e:?}"),
        }
        let out_path = String::from(&format!("output/{}.tex", self.pdf_name));
        std::process::Command::new("latexmk")
            .arg(out_path)
            .arg("-pdf")
            .arg("--output-directory=output/")
            .status()?;
        Ok(())
    }
}
