use crate::comglue::glue::IRTDUpdateEventWrap;
use anyhow::Result;
use futures::{channel::mpsc, prelude::*};
use fxhash::{FxBuildHasher, FxHashMap, FxHashSet};
use netidx::{
    config::Config,
    path::Path,
    pool::{Pool, Pooled},
    resolver::Auth,
    subscriber::{Dval, Event, SubId, Subscriber, UpdatesFlags},
};
use once_cell::sync::Lazy;
use parking_lot::Mutex;
use std::{
    collections::{HashMap, HashSet},
    mem,
    sync::Arc,
    fmt
};
use tokio::runtime::Runtime;

#[derive(Debug, Clone, Copy, Hash, PartialEq, Eq)]
pub(crate) struct TopicId(i32);

static PENDING: Lazy<Pool<FxHashMap<TopicId, Event>>> =
    Lazy::new(|| Pool::new(3, 1_000_000));

struct ServerInner {
    runtime: Runtime,
    update: IRTDUpdateEventWrap,
    subscriber: Subscriber,
    updates: mpsc::Sender<Pooled<Vec<(SubId, Event)>>>,
    by_id: FxHashMap<SubId, FxHashSet<TopicId>>,
    by_topic: FxHashMap<TopicId, Dval>,
    pending: Pooled<FxHashMap<TopicId, Event>>,
}

#[derive(Clone)]
pub struct Server(Arc<Mutex<ServerInner>>);

impl fmt::Debug for Server {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(f, "Server")
    }
}

impl Server {
    async fn updates_loop(self, mut up: mpsc::Receiver<Pooled<Vec<(SubId, Event)>>>) {
        while let Some(mut updates) = up.next().await {
            let mut inner = self.0.lock();
            let inner = &mut *inner;
            let call_update = inner.pending.is_empty();
            for (id, ev) in updates.drain(..) {
                if let Some(tids) = inner.by_id.get(&id) {
                    let mut iter = tids.iter();
                    for _ in 0..tids.len() - 1 {
                        inner.pending.insert(*iter.next().unwrap(), ev.clone());
                    }
                    inner.pending.insert(*iter.next().unwrap(), ev);
                }
            }
            if call_update {
                inner.update.update_notify();
            }
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
            pending: PENDING.take(),
        })));
        let t_ = t.clone();
        t.0.lock().runtime.spawn(t_.updates_loop(rx));
        Ok(t)
    }

    pub(crate) fn connect_data(&self, tid: TopicId, path: Path) -> Result<()> {
        let mut inner = self.0.lock();
        let dv = inner.subscriber.durable_subscribe(path);
        dv.updates(UpdatesFlags::BEGIN_WITH_LAST, inner.updates.clone());
        inner
            .by_id
            .entry(dv.id())
            .or_insert_with(|| HashSet::with_hasher(FxBuildHasher::default()))
            .insert(tid);
        inner.by_topic.insert(tid, dv);
        Ok(())
    }

    pub(crate) fn disconnect_data(&self, tid: TopicId) {
        let mut inner = self.0.lock();
        if let Some(dv) = inner.by_topic.remove(&tid) {
            if let Some(tids) = inner.by_id.get_mut(&dv.id()) {
                tids.remove(&tid);
                if tids.is_empty() {
                    inner.by_id.remove(&dv.id());
                }
            }
        }
    }

    pub(crate) fn refresh_data(&self) -> Pooled<FxHashMap<TopicId, Event>> {
        let mut inner = self.0.lock();
        mem::replace(&mut inner.pending, PENDING.take())
    }
}
