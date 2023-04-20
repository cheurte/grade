use latex::{DocumentClass, Element, Document, Section, Align};
use std::fs::File;
use std::io::Write;

use std::collections::HashMap;
use std::fs::write;
use latexcompile::{LatexCompiler, LatexInput};

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

fn create_tex_file(rendered: String)->std::io::Result<()> {
    let mut f = File::create("output/report.tex")?;
    write!(f, "{}", rendered)?;
    Ok(())
}

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
