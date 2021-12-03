#[macro_use] extern crate netidx_core;
mod comglue;
mod server;
use comglue::interface::CLSID;
use comglue::glue::NetidxRTD;

com::inproc_dll_module![(CLSID, NetidxRTD),];
