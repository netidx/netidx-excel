use crate::comglue::glue::IRTDUpdateEventWrap;
use std::collections::HashMap;

struct TopicId(i32);

struct Server {
    update: IRTDUpdateEventWrap,
    by_name: HashMap<String, i32>,
    by_id: HashMap<i32, ()>,
}
