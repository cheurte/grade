use std::path::Path;
use latex::{DocumentClass, Element, Document, Section, Align};
use std::fs::File;
use std::io::{Write, Read};

use std::collections::HashMap;
use std::fs::{self, write};
use latexcompile::{LatexCompiler, LatexInput};

use calamine::{Reader, open_workbook, Xlsx, DataType};

use serde::Deserialize;

#[derive(Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
struct ConfigXlsx {
    file_name: String,
    sheet_name: String,
    allowed_changes: Vec<String>,
}

pub fn read_source_config()->std::io::Result<()>{
    let path = Path::new("config/config_source.json");
    let file = File::open(path)?;

    let config:ConfigXlsx= serde_json::from_reader(file)?;
    println!("{:?}", config);
    Ok(())
}

/// Create the string that will be compiled. 
/// This function will be depending on json files later. -> Todo
///
pub fn create_tex_structure()->String{
    let mut doc = Document::new(DocumentClass::Article);

    // / Set some metadata for the document
    doc.preamble.title("My Fancy Document");
    doc.preamble.author("Michael-F-Bryan");

    doc.push(Element::TitlePage)
        .push(Element::ClearPage)
        .push(Element::TableOfContents)
        .push(Element::ClearPage);

    let mut section_1 = Section::new("Section 1");
    section_1.push("Here is some text which will be put in paragraph 1.")
             .push("And here is some more text for paragraph 2.");
    doc.push(section_1);

    let mut section_2 = Section::new("Section 2");

    section_2.push("More text...")
             .push(Align::from("y &= mx + c"));

    doc.push(section_2);

    latex::print(&doc).expect("error at rendering");
    String::new()

}

/// Handle the tex creation. 
/// Seperate function to handle the windows server lately -> ToDo
/// 
fn create_tex_file(rendered: String)->std::io::Result<()> {
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
fn create_pdf()->std::io::Result<()>{
    let mut dict:HashMap<String,String> = HashMap::new();
    dict.insert("test".into(), "Minimal".into());
    // // provide the folder where the file for latex compiler are found
    let input = LatexInput::from("assets");
    // // create a new clean compiler enviroment and the compiler wrapper
    let compiler = LatexCompiler::new(dict).unwrap();
    // // run the underlying pdflatex or whatever
    let result = compiler.run("/home/cheurte/Documents/grade/latex_creation/output/report.tex", &input).unwrap();
    //
    // // copy the file into the working directory
    let output = ::std::env::current_dir().unwrap().join("output/out.pdf");
    assert!(write(output, result).is_ok());
    Ok(())
}

/// Function to read the content of the file. Not finished yet. 
/// Suppose to depend on a json file to know wich line or row can be red.
///
fn read_sources(){
    let path = "sources/BIOTEC.xlsx";
    let mut workbook: Xlsx<_> = open_workbook(path).expect("cannot open file");
    
    if let Some(Ok(range)) = workbook.worksheet_range("Master - Flexible Overview") {
        // let total_cells = range.get_size().0 * range
        println!("{:?}", range.get_size());
    }
}

