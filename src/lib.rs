mod comglue;
mod server;
use comglue::glue::NetidxRTD;
use comglue::interface::CLSID;

com::inproc_dll_module![(CLSID, NetidxRTD),];
