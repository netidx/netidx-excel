#[macro_use] extern crate netidx_core;
mod comglue;
mod server;
use comglue::interface::CLSID;
use comglue::glue::NetidxRTD;

// invoke the deep evil magic of the elder gods of win16
com::inproc_dll_module![(CLSID, NetidxRTD),];
