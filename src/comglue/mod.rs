pub mod dispatch;
pub mod glue;
pub mod interface;
pub mod variant;

use once_cell::sync::Lazy;
use simplelog::{self, LevelFilter};
use std::fs::File;

static LOGGER: Lazy<()> = Lazy::new(|| {
    let f = File::create("C:\\Users\\eric\\proj\\netidx-excel\\log.txt")
        .expect("couldn't open log file");
    simplelog::WriteLogger::init(LevelFilter::Warn, simplelog::Config::default(), f)
        .expect("couldn't init log")
});

pub fn maybe_init_logger() {
    *LOGGER
}
