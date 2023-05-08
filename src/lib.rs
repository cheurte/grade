use std::path::Path;
use latex::{DocumentClass, Element, Document, Section, Align, print};
use std::fs::File;
use std::io::Write;

use std::collections::HashMap;
use std::fs::write;
use latexcompile::{LatexCompiler, LatexInput};

use calamine::{Reader, open_workbook, Xlsx, DataType};

use serde::{Deserialize, de::value};

#[derive(Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ConfigXlsx {
    pub pdf_file: Vec<PdfFile>, 
}

#[derive(Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct PdfFile {
    pdf_name: String,
    sources: Vec<String>,
    sheets_name: Vec<String>,
    products: Vec<String>,
    categories: Vec<String>,
}

fn create_tabular(columns: String, values: Vec<String>)->String{
    let values = values.join(" & ");
    
    let tabular = &format!(
        "\\begin{{tabular}}{{{}}}\n\
       {} \n\
        \\end{{tabular}}"
        , columns, values);
    tabular.to_string()
}

/// Create the string that will be compiled. 
/// This function will be depending on json files later. -> Todo
///
pub fn page_blue_print()->String{
    let mut doc = Document::new(DocumentClass::Report);
    doc.preamble.title("Template document");
    doc.push(Element::TitlePage);

    let tab = &create_tabular("c c".to_string(), vec!["a".to_string(), "b".to_string()])[..];

    doc.push(tab);
    print(&doc).unwrap()
}

/// Handle the tex creation. 
/// Seperate function to handle the windows server lately -> ToDo
/// 
pub fn create_tex_file(rendered: String)->std::io::Result<()> {
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
pub fn create_pdf()->std::io::Result<()>{
    let mut dict:HashMap<String,String> = HashMap::new();
    dict.insert("test".into(), "Minimal".into());
    // provide the folder where the file for latex compiler are found
    let input = LatexInput::from("assets");
    // create a new clean compiler enviroment and the compiler wrapper
    let compiler = LatexCompiler::new(dict).unwrap();
    // run the underlying pdflatex or whatever
    let result = compiler.run("/home/cheurte/Documents/grade/output/report.tex", &input).unwrap();
    // // copy the file into the working directory
    let output = ::std::env::current_dir().unwrap().join("output/test.pdf");
    assert!(write(output, result).is_ok());
    Ok(())
}

impl ConfigXlsx{
    pub fn new()->Self {
       Self { pdf_file: vec![] }
    }

    pub fn from(path: &str)->Self{
        let path = Path::new(path);
        let file = File::open(path).expect("Problem with the file");
        let config: ConfigXlsx = serde_json::from_reader(file).expect("error in the reading");
        config
    }
}

impl PdfFile {
    /// Function to search the key elements of the file.
    /// Suppose to depend on a json file to know wich line or row can be red.   
    pub fn get_coordinates(&self)->(Vec<(usize, usize)>, Vec<(usize, usize)>){
        // To be changed to handle errors
        let path = self.sources.iter().nth(0).unwrap();
        let sheet = self.sheets_name.iter().nth(0).unwrap();
        let products = &self.products;
        let categories = &self.categories;
        let mut workbook: Xlsx<_> = open_workbook(path)
                                        .expect("Problem by reading file");
        let mut categories_coord: Vec<(usize, usize)>=vec![];
        let mut products_coord: Vec<(usize, usize)>=vec![];
        match workbook.worksheet_range(&sheet) {
            Some(Ok(range)) => {
                for row in 0..range.get_size().0{
                    for col in 0..range.get_size().1{
                        let value = range.get_value((row as u32, col as u32));
                        if value != Some(&DataType::Empty){
                            for category in categories{
                                if &value.unwrap().to_string() == category{
                                    categories_coord.push((row, col));
                                }
                            }
                            for product in products{
                                if &value.unwrap().to_string() == product{
                                    products_coord.push((row, col));
                                }
                            }
                        }
                    }
                }
            },
            Some(Err(e)) => println!("{e:?}"),
            None => println!("Sheets name unknown. Maybe check the name in the config file")
        }
        (categories_coord, products_coord)
    }
    /// Search the range of categories, to know where to stop
    /// Take the beginning corrdinates of categories, and return the end coordinates
    pub fn get_parameters_range(&self, categories_coord: &Vec<(usize, usize)>) 
        -> Vec<(usize, usize)> 
    {
        let sheet = self.sheets_name.iter().nth(0).unwrap();
        let path = self.sources.iter().nth(0).unwrap();

        let mut end_categories: Vec<(usize, usize)> = vec![];
        let mut workbook: Xlsx<_> = open_workbook(path)
                                        .expect("Problem by reading file");
        match workbook.worksheet_range(&sheet) {
            Some(Ok(range)) => {
                for (category_row, category_col) in categories_coord.into_iter(){
                    let mut col = category_col+1;
                    loop {
                        if range.get_value((*category_row as u32, col as u32)) != Some(&DataType::Empty) || col > range.get_size().1{
                            break
                        }
                        col +=1;
                    }
                    end_categories.push((*category_row, col-1));
                }                
            },
            Some(Err(e)) => println!("{e:?}"),
            None => println!("Sheets name unknown. Maybe check the name in the config file")
        }
        end_categories
    }
    /// Get parameters names
    /// Take the indices of values to get them.
    pub fn get_parameters_name(&self, start_categ_coord: &Vec<(usize, usize)>, end_categ_coord: &Vec<(usize, usize)>)->Vec<Vec<String>>{
        assert_eq!(start_categ_coord.len(), end_categ_coord.len()); 
        let sheet = self.sheets_name.iter().nth(0).unwrap();
        let path = self.sources.iter().nth(0).unwrap();
        let mut workbook: Xlsx<_> = open_workbook(path)
                                        .expect("Problem by reading file");
        let mut output: Vec<Vec<String>> = vec![];
        match workbook.worksheet_range(&sheet) {
            Some(Ok(range)) => {
                let it = start_categ_coord.iter().zip(end_categ_coord.iter());
                for (_, (start_coord, end_coord)) in it.enumerate() {
                    let mut parameters: Vec<String> = vec![];
                    for col in start_coord.1+2..end_coord.1{
                        parameters
                            .push(range
                                .get_value(((start_coord.0+1) as u32, col as u32))
                                    .unwrap()
                                    .to_string())
                    }
                    output.push(parameters);
                }
            },
            Some(Err(e)) => println!("{e:?}"),
            None => println!("Sheets name unknown. Maybe check the name in the config file")
        }
        output
    }
}
