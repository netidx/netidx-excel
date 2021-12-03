use crate::comglue::glue::IRTDUpdateEventWrap;
use tokio::{ runtime::Runtime, sync::mpsc };
use fxhash::{FxHashMap, FxBuildHasher};
use std::{collections::HashMap, sync::Arc};
use netidx::{
    pool::Pooled, 
    subscriber::{Subscriber, Event, Value, Dval, SubId}, 
    config::Config, resolver::Auth
};
use anyhow::Result;
use parking_lot::Mutex;

#[derive(Debug, Clone, Copy, Hash, PartialEq, Eq)]
pub(crate) struct TopicId(i32);

struct ServerInner {
    runtime: Runtime,
    update: IRTDUpdateEventWrap,
    subscriber: Subscriber,
    updates: mpsc::Sender<Pooled<Vec<(SubId, Event)>>>,
    by_id: FxHashMap<SubId, TopicId>,
    by_topic: FxHashMap<TopicId, Dval>,
    pending: FxHashMap<TopicId, Event>,
}

#[derive(Clone)]
pub(crate) struct Server(Arc<Mutex<ServerInner>>);

impl Server {
    async fn updates_loop(self, mut up: mpsc::Receiver<Pooled<Vec<(SubId, Event)>>>) {
        while let Some(updates) = up.recv().await {
            ()
        }
    }

    pub(crate) fn new(update: IRTDUpdateEventWrap) -> Result<Server> {
        let runtime = Runtime::new()?;
        let subscriber = runtime.block_on(async {
            let config = Config::load_default()?;
            Subscriber::new(config, Auth::Anonymous)
        })?;
        let (tx, rx) = runtime.block_on(async { mpsc::channel(3) });
        let t = Server(Arc::new(Mutex::new(ServerInner {
            runtime,
            update,
            subscriber,
            updates: tx,
            by_id: HashMap::with_hasher(FxBuildHasher::default()),
            by_topic: HashMap::with_hasher(FxBuildHasher::default()),
            pending: HashMap::with_hasher(FxBuildHasher::default()),
        })));
        let t_ = t.clone();
        t.0.lock().runtime.spawn(t_.updates_loop(rx));
        Ok(t)
    }
}