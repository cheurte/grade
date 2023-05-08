// use grade::{page_blue_print, create_pdf, create_tex_file};
use grade::ConfigXlsx;

fn main() {
    // let content = page_blue_print();
    // println!("{content:?}");
    // match create_tex_file(content){
    //     Ok(_) => println!("ok"),
    //     Err(e) => println!("{e:?}"),
    // }
    // create_pdf().expect("error creating pdf");
    let configs = ConfigXlsx::from("config/config_source.json");
    for pdf_file in configs.pdf_file.iter(){
        let (cat_coord, _): (Vec<(usize, usize)>, Vec<(usize, usize)>) = pdf_file.get_coordinates();
        // println!("{cat_coord:?}");
        let end_categories = pdf_file.get_parameters_range(&cat_coord);
        // println!("{:?}", end_catego as u32ries);
        let param = pdf_file.get_parameters_name(&cat_coord, &end_categories);
        println!("{param:?}");
        
        break;
    }
}
