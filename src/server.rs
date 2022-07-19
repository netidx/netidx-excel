use crate::comglue::{self, dispatch::IRTDUpdateEventWrap};
use anyhow::Result;
use futures::{channel::mpsc, prelude::*};
use fxhash::{FxBuildHasher, FxHashMap, FxHashSet};
use log::{debug, error};
use netidx::{
    config::Config,
    path::Path,
    pool::{Pool, Pooled},
    subscriber::{DesiredAuth, Dval, Event, SubId, Subscriber, UpdatesFlags},
};
use once_cell::sync::Lazy;
use parking_lot::Mutex;
use std::{
    collections::{HashMap, HashSet},
    default::Default,
    fmt, mem,
    sync::Arc,
};
use tokio::runtime::Runtime;

#[derive(Debug, Clone, Copy, Hash, PartialEq, Eq)]
pub(crate) struct TopicId(pub i32);

static PENDING: Lazy<Pool<FxHashMap<TopicId, Event>>> =
    Lazy::new(|| Pool::new(3, 1_000_000));

struct ServerInner {
    runtime: Runtime,
    update: Option<IRTDUpdateEventWrap>,
    subscriber: Subscriber,
    updates: mpsc::Sender<Pooled<Vec<(SubId, Event)>>>,
    by_id: FxHashMap<SubId, FxHashSet<TopicId>>,
    by_topic: FxHashMap<TopicId, Dval>,
    pending: Pooled<FxHashMap<TopicId, Event>>,
}

impl ServerInner {
    fn clear(&mut self) {
        self.update = None;
        self.by_id.clear();
        self.by_topic.clear();
        self.pending.clear();
    }
}

#[derive(Clone)]
pub struct Server(Arc<Mutex<Option<ServerInner>>>);

impl Default for Server {
    fn default() -> Self {
        let cfg = *comglue::CONFIG;
        Self::new(cfg)
    }
}

impl fmt::Debug for Server {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(f, "Server")
    }
}

impl Server {
    async fn updates_loop(self, mut up: mpsc::Receiver<Pooled<Vec<(SubId, Event)>>>) {
        debug!("updates loop started");
        while let Some(mut updates) = up.next().await {
            let mut inner = self.0.lock();
            if let Some(inner) = &mut *inner {
                if let Some(update) = &mut inner.update {
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
                        debug!("calling update_notify");
                        update.update_notify();
                    }
                }
            }
        }
        debug!("updates loop terminated")
    }

    pub(crate) fn new(cfg: comglue::Config) -> Server {
        macro_rules! or_err {
            ($e:expr, $msg:expr) => {
                match $e {
                    Ok(r) => r,
                    Err(e) => {
                        error!($msg, e);
                        return Server(Arc::new(Mutex::new(None)))
                    }
                }
            }
        }
        debug!("init runtime");
        let runtime = or_err!(Runtime::new(), "could not init async runtime {}");
        debug!("entering async to init subscriber");
        let subscriber: Result<Subscriber> = runtime.block_on(async {
            debug!("running in async context");
            let auth = match cfg.auth_mechanism {
                comglue::Auth::Anonymous => DesiredAuth::Anonymous,
                comglue::Auth::Kerberos => DesiredAuth::Krb5 { spn: None, upn: None },
            };
            debug!("loading config file");
            let config = Config::load_default()?;
            debug!("starting subscriber");
            Ok(Subscriber::new(config, auth)?)
        });
        let subscriber = or_err!(subscriber, "could not init subscriber {}");
        debug!("init updates channel");
        let (tx, rx) = runtime.block_on(async { mpsc::channel(3) });
        let t = Server(Arc::new(Mutex::new(Some(ServerInner {
            runtime,
            update: None,
            subscriber,
            updates: tx,
            by_id: HashMap::with_hasher(FxBuildHasher::default()),
            by_topic: HashMap::with_hasher(FxBuildHasher::default()),
            pending: PENDING.take(),
        }))));
        let t_ = t.clone();
        if let Some(inner) = &mut *t.0.lock() {
            debug!("starting updates loop");
            inner.runtime.spawn(t_.updates_loop(rx));
        }
        t
    }

    pub(crate) fn server_start(&self, update: IRTDUpdateEventWrap) {
        if let Some(inner) = &mut *self.0.lock() {
            inner.clear();
            inner.update = Some(update);
            debug!("server_start");
        }
    }

    pub(crate) fn server_terminate(&self) {
        if let Some(inner) = &mut *self.0.lock() {
            inner.clear();
            debug!("server_terminate");
        }
    }

    pub(crate) fn connect_data(&self, tid: TopicId, path: Path) -> Result<()> {
        debug!("connect_data");
        if let Some(inner) = &mut *self.0.lock() {
            let dv = inner.subscriber.durable_subscribe(path);
            inner.pending.insert(tid, dv.last());
            if let Some(update) = inner.update.as_ref() {
                update.update_notify()
            }
            dv.updates(UpdatesFlags::BEGIN_WITH_LAST, inner.updates.clone());
            inner
                .by_id
                .entry(dv.id())
                .or_insert_with(|| HashSet::with_hasher(FxBuildHasher::default()))
                .insert(tid);
            inner.by_topic.insert(tid, dv);
        }
        Ok(())
    }

    pub(crate) fn disconnect_data(&self, tid: TopicId) {
        debug!("disconnect_data");
        if let Some(inner) = &mut *self.0.lock() {
            inner.pending.remove(&tid);
            if let Some(dv) = inner.by_topic.remove(&tid) {
                if let Some(tids) = inner.by_id.get_mut(&dv.id()) {
                    tids.remove(&tid);
                    if tids.is_empty() {
                        inner.by_id.remove(&dv.id());
                    }
                }
            }
        }
    }

    pub(crate) fn refresh_data(&self) -> Pooled<FxHashMap<TopicId, Event>> {
        match &mut *self.0.lock() {
            Some(inner) => {
                debug!("refresh_data");
                mem::replace(&mut inner.pending, PENDING.take())
            }
            None => Pooled::orphan(HashMap::default())
        }
    }
}
