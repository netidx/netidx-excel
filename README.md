This is an Excel COM add-in that allows the `=RTD()` formula to pull in real time data from Netidx.

E.G. You have published some data from the shell publisher
```bash
> netidx publisher --bind 192.168.0.0/24 --spn publish/ryouko.ryu-oh.org@RYU-OH.ORG
/test/foo|array|[42, -42]
/test/bar|i32|42
/test/baz|i32|-42
```

You can pull that data into a cell in an Excel spreadsheet

![Example](./example.png)

The key point is if a value in netidx updates, Excel will update almost immediatly. Even if you subscribe to a lot of data, or you subscribe to data that updates quickly, Excel should remain responsive because RTDs are throttled, and all the netidx processing is happening on a background thread pool.

For example here Excel is maxing out my wifi network by subscribing to the stress publisher, however it remains completely responsive.

![Performance](./perf.png)

# Building

You need a windows machine (obviously) probably windows 10 or above, but 8 might work. You need to install rust, the easiest way is to use rustup. You need to install git. Once you have those two things installed,

```bash
> git clone https://github.com/estokes/netidx-excel
> cd netidx-excel
> cargo build --release
```

The dll should be built in `target/release/netidx_excel.dll`

# Installing

To install you need to decide where you want the dll to live, it really doesn't matter, but I put it in `C:\Program Files\netidx-excel` on my machine. Then you need to run `regsvr32` on the dll as Administrator, that will set up the registry entries to register it as a proper COM server. So in an admin powershell,

```powershell
> mkdir 'C:\Program Files\netidx-excel'
> cp target\release\netidx_excel.dll 'C:\Program Files\netidx-excel'
> regsvr32 'C:\Program Files\netidx-excel\netidx_excel.dll'
```

The most common errors are `regsvr32` isn't in your path, and/or your shell is not running with admin rights.
