use grade::{read_config, read_sources};

fn main() {
    let configs = read_config().expect("error on reading json");
    for config in configs.pdf_file.iter(){
        read_sources(config);
    }
    // read_sources()
}
