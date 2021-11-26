mod interface;
mod server;
use interface::CLSID;
use server::NetidxRTD;

// invoke the deep evil magic of the elder gods of win16
com::inproc_dll_module![(CLSID, NetidxRTD),];
