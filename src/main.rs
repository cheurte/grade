use grade::read_source_config;

fn main() {
    match read_source_config(){
        Ok(_) => println!("No problemo"),
        Err(e) => println!("{e:?}"),
    };
}
