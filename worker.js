import * as XLSX from "xlsx";
/**
 * Cloudflare Worker + D1: Broker Summary dashboard.
 * User requests only read D1. ZAPI is accessed by the scheduled ingestion job.
 */
const API_BASE = "https://api.zapi.ink/v1";
const API_BROKER = API_BASE + "/finance:pluang/broker-summary";
const API_SECURITIES = API_BASE + "/finance:idx/securities";
const API_BROKERS = API_BASE + "/finance:pluang/brokers";
const API_KSEI = API_BASE + "/finance:ksei/ownership";
const API_STOCKBIT_SYMBOL = API_BASE + "/finance:stockbit/symbol";
const RETAIL_BROKERS = ["XC","XL","YP","CP","PD","SQ","KK","GR","IF","NI","CC"];

export default {
  async fetch(request, env, ctx) {
    const url = new URL(request.url);
    if (url.pathname === "/health") return json({ok:true,service:"broker-summary"});
    if (url.pathname === "/api/stocks") return handleStocks(url, env);
    if (url.pathname === "/api/dashboard") return handleDashboard(url, env, ctx);
    if (url.pathname === "/api/flow-source") return handleFlowSource(url, env, ctx);
    if (url.pathname === "/api/flow-summary") return handleFlowSummary(url, env, ctx);
    if (url.pathname === "/api/retail-brokers") return handleRetailBrokers(request, env);
    if (url.pathname === "/api/stock-profile") return handleStockProfile(url, env);
    if (url.pathname === "/api/retail-ownership") return handleRetailOwnership(url, env);
    if (url.pathname === "/api/idx-ownership") return handleIdxOwnership(url, env);
    if (url.pathname === "/admin/idx-ownership-refresh") return handleIdxOwnershipRefresh(request, env);
    if (url.pathname === "/admin/idx-broker-master-refresh") return handleIdxBrokerMasterRefresh(request, env);
    if (url.pathname === "/api/idx-broker-master") return handleIdxBrokerMaster(url, env);
    if (url.pathname === "/api/broker-history") return handleBrokerHistory(url, env);
    if (url.pathname === "/api/stockbit-quote") return handleStockbitQuote(url, env);
    if (url.pathname === "/api/pluang-quote") return handlePluangQuote(url, env);
    if (url.pathname === "/admin/backfill") return handleBackfill(request, url, env);
    if (url.pathname === "/") return new Response(await env.ASSETS.fetch(new Request(new URL("/index.html", request.url))), {headers:{"Content-Type":"text/html; charset=utf-8"}});
    return new Response("Not found", {status:404});
  },
  // Tidak ada scheduled bulk-download kepemilikan >5%. Data hanya diambil
  // saat ticker diminta dan hanya baris ticker itu yang disimpan ke D1.
};

function json(data, status=200) {
  return new Response(JSON.stringify(data), {status, headers:{
    "Content-Type":"application/json; charset=utf-8",
    "Cache-Control":"no-store",
    "Access-Control-Allow-Origin":"*"
  }});
}
function cleanStock(v){
  const code=String(v||"").trim().toUpperCase().replace(/\.JK$/i,"");
  // Nama yang sering diketik pengguna; ticker resmi Antam di BEI adalah ANTM.
  return code==="ANTAM" ? "ANTM" : code;
}
function num(v){if(v===null||v===undefined||v==="")return 0;const n=Number(String(v).replace(/,/g,"").trim());return Number.isFinite(n)?n:0;}
function isoTodayWIB(){return new Intl.DateTimeFormat("en-CA",{timeZone:"Asia/Jakarta",year:"numeric",month:"2-digit",day:"2-digit"}).format(new Date());}
function latestCompletedDateWIB(){
  // Data harian biasanya tersedia sekitar pukul 18.00 WIB. Sebelum itu gunakan hari sebelumnya.
  const parts=new Intl.DateTimeFormat("en-CA",{timeZone:"Asia/Jakarta",year:"numeric",month:"2-digit",day:"2-digit",hour:"2-digit",hourCycle:"h23"}).formatToParts(new Date());
  const get=k=>parts.find(p=>p.type===k)?.value||"00";
  let d=new Date(Date.UTC(Number(get("year")),Number(get("month"))-1,Number(get("day"))));
  if(Number(get("hour"))<18)d.setUTCDate(d.getUTCDate()-1);
  // Akhir pekan tidak memiliki pembaruan perdagangan; mundur ke Jumat.
  while(d.getUTCDay()===0 || d.getUTCDay()===6)d.setUTCDate(d.getUTCDate()-1);
  return d.toISOString().slice(0,10);
}
function parseDate(s){return /^\d{4}-\d{2}-\d{2}$/.test(s||"") ? s : "";}
function dateRange(start,end){if(start>end)[start,end]=[end,start];const out=[];let d=new Date(start+"T00:00:00Z"),last=new Date(end+"T00:00:00Z");while(d<=last){out.push(d.toISOString().slice(0,10));d.setUTCDate(d.getUTCDate()+1);}return out;}
function weekdaysOnly(date){const d=new Date(date+"T00:00:00Z").getUTCDay();return d!==0&&d!==6;}
function unwrap(j){
  let out=j;
  // ZAPI kadang membungkus payload lebih dari satu tingkat.
  for(let i=0;i<4;i++){
    if(!out || typeof out!=="object" || Array.isArray(out)) break;
    const next=out.data ?? out.result ?? out.response;
    if(!next || next===out || typeof next!=="object") break;
    out=next;
  }
  return out;
}
function normalizeBrokerSummary(raw, code, date, useNet=false){
  let d=raw;
  // Cari struktur buyers/sellers meskipun dibungkus objek response/summary.
  for(let i=0;i<5;i++){
    if(d && typeof d==="object" && (Array.isArray(d.buyers)||Array.isArray(d.sellers))) break;
    if(!d || typeof d!=="object") break;
    const next=d.data ?? d.result ?? d.summary ?? d.brokerSummary ?? d.items;
    if(!next || next===d) break;
    d=next;
  }
  if(Array.isArray(d)) {
    const buyers=[], sellers=[];
    for(const row of d){
      const side=String(row.side||row.type||row.action||row.transactionType||"").toLowerCase();
      const value = num(row.value??row.netValue??row.amount??row.totalValue??row.total_value??0);
      
      // Mode NET: broker dengan value negatif tidak boleh masuk daftar BUY.
      // Pada mode gross, pertahankan data sumber apa adanya.
      if (useNet && /buy|buyer|beli/.test(side) && value < 0) {
        continue;
      }

      const item={
        broker:row.broker||row.brokerCode||row.code||row.broker_code,
        lots:row.lots??row.netLots??row.lot??row.volume??row.totalLot??row.total_lot??0,
        value:value,
        averagePrice:row.averagePrice??row.avgPrice??row.average_price??row.avg??0
      };
      if(/sell|seller|jual/.test(side)) sellers.push(item);
      else if(/buy|buyer|beli/.test(side)) buyers.push(item);
    }
    d={buyers,sellers};
  }
  if(!d || typeof d!=="object") throw new Error("Format Broker Summary tidak dikenali");
  
  let buyers=Array.isArray(d.buyers)?d.buyers:(Array.isArray(d.buy)?d.buy:[]);
  const sellers=Array.isArray(d.sellers)?d.sellers:(Array.isArray(d.sell)?d.sell:[]);
  if(!Array.isArray(buyers)||!Array.isArray(sellers)) throw new Error("Respons Broker Summary tidak memiliki daftar BUY/SELL");
  
  const normalizeSide=items=>items.map(x=>({
    broker:String(x.broker||x.brokerCode||x.code||x.broker_code||"-").toUpperCase(),
    lots:num(x.lots??x.netLots??x.lot??x.volume??x.totalLot??x.total_lot),
    value:num(x.value??x.netValue??x.amount??x.totalValue??x.total_value),
    averagePrice:num(x.averagePrice??x.avgPrice??x.average_price??x.avg??x.price)
  }));

  // Filter defensif khusus NET BUY: value harus >= 0 sebelum disimpan ke D1.
  const normalizedBuyers = normalizeSide(buyers).filter(x => !useNet || x.value >= 0);
  const normalizedSellers = normalizeSide(sellers);

  return {code:d.code||code,startDate:d.startDate||date,endDate:d.endDate||date,
    capped:d.capped===true,count:num(d.count),buyers:normalizedBuyers,sellers:normalizedSellers};
}
//``[cite: 11]
async function zapi(env,path){
  if(!env.ZAPI_API_KEY) throw new Error("Secret ZAPI_API_KEY belum disetel.");
  const r=await fetch(path,{headers:{"x-api-key":env.ZAPI_API_KEY,"Accept":"application/json"}});
  const t=await r.text();let j;try{j=JSON.parse(t)}catch{throw new Error("Respons ZAPI bukan JSON: "+t.slice(0,180))}
  if(!r.ok)throw new Error("ZAPI HTTP "+r.status+": "+(j.message||"request gagal"));
  return unwrap(j);
}
function hasBrokerRows_(item){
  return !!item && ((Array.isArray(item.buyers) && item.buyers.length > 0) ||
    (Array.isArray(item.sellers) && item.sellers.length > 0));
}
async function getDailyFromDb(env,stock,date,useNet=false){
  const column=useNet?"net_payload":"payload";
  const row=await env.DB.prepare(`SELECT ${column} AS selected_payload FROM broker_daily WHERE stock=? AND trade_date=?`).bind(stock,date).first();
  if(!row || !row.selected_payload)return null;
  try{
    const data=JSON.parse(row.selected_payload);
    // Bersihkan data NET lama yang tersimpan sebelum filter diterapkan.
    if(useNet && data && Array.isArray(data.buyers)){
      data.buyers=data.buyers.filter(x=>num(x.value??x.netValue??x.amount??x.totalValue??x.total_value)>=0);
    }
    return hasBrokerRows_(data)?data:null;
  }catch{return null}
}
async function saveStockMaster_(env, s){
  const code=cleanStock(s.Code||s.code||s.symbol);
  if(!/^[A-Z0-9]{1,4}$/.test(code)) return false;
  const name=String(s.Name||s.name||s.companyName||"");
  const board=String(s.ListingBoard||s.listingBoard||s.board||"");
  const logo=String(s.Logo||s.logo||s.logoUrl||s.iconUrl||s.image||"");
  await env.DB.prepare("INSERT INTO stocks(code,name,board,updated_at) VALUES(?,?,?,?) ON CONFLICT(code) DO UPDATE SET name=CASE WHEN excluded.name<>'' THEN excluded.name ELSE stocks.name END,board=CASE WHEN excluded.board<>'' THEN excluded.board ELSE stocks.board END,updated_at=excluded.updated_at")
    .bind(code,name,board,new Date().toISOString()).run();
  if(logo){try{await env.DB.prepare("UPDATE stocks SET logo=? WHERE code=?").bind(logo,code).run();}catch(e){console.error("Kolom stocks.logo belum tersedia:",String(e.message||e));}}
  return true;
}
async function lookupSecurity_(env, code){
  const attempts=[API_SECURITIES+"?code="+encodeURIComponent(code),API_SECURITIES+"?search="+encodeURIComponent(code),API_SECURITIES+"?q="+encodeURIComponent(code)+"&length=30&start=0"];
  for(const path of attempts){
    try{
      let d=await zapi(env,path);
      for(let i=0;i<3;i++){const nested=d&&(d.data||d.result);if(!nested||nested===d)break;d=nested;}
      const arr=Array.isArray(d)?d:(Array.isArray(d.items)?d.items:(Array.isArray(d.results)?d.results:[]));
      const found=arr.find(x=>cleanStock(x.Code||x.code||x.symbol)===code);
      if(found)return found;
      if(d && !Array.isArray(d) && cleanStock(d.Code||d.code||d.symbol)===code)return d;
    }catch(e){console.error("Security lookup failed:",path,String(e.message||e));}
  }
  return null;
}
function firstOfNextMonth(date){
  const d=new Date(date+"T00:00:00Z");
  d.setUTCDate(1); d.setUTCMonth(d.getUTCMonth()+1);
  return d.toISOString().slice(0,10);
}
async function fetchDailyOnDemand(env,stock,date,useNet=false){
  const today=isoTodayWIB();
  const blockedUntil=await metaGet(env,"zapi_quota_blocked_until");
  if(blockedUntil && today < blockedUntil){
    return {ok:false,blocked:true,error:"Kuota bulanan ZAPI telah habis. Endpoint dicoba kembali mulai "+blockedUntil+". Data yang sudah ada tetap dibaca dari database."};
  }
  if(blockedUntil && today >= blockedUntil) await metaSet(env,"zapi_quota_blocked_until","");
  let lastError="";
  // Retry kegagalan sementara; jeda mencegah lonjakan request serentak ke ZAPI.
  for(let attempt=0;attempt<3;attempt++){
    try{
      const data=await zapi(env,API_BROKER+"?code="+encodeURIComponent(stock)+"&startDate="+date+"&endDate="+date+"&net="+(useNet?"true":"false"));
      await saveBrokerDaily(env,stock,date,data,useNet);
      return {ok:true,date,attempts:attempt+1};
    }catch(e){
      lastError=String(e.message||e).slice(0,500);
      if(/\b429\b|quota exceeded|monthly quota/i.test(lastError)){
        const until=firstOfNextMonth(today);
        await metaSet(env,"zapi_quota_blocked_until",until);
        return {ok:false,date,blocked:true,error:"Kuota bulanan ZAPI telah habis (HTTP 429). Kuota diperkirakan reset pada "+until+". Data yang sudah tersimpan tetap dapat ditampilkan.",detail:lastError};
      }
      // Kesalahan format / autentikasi biasanya tidak pulih dengan retry.
      if(/HTTP 4\d\d/.test(lastError) && !/HTTP 408|HTTP 429/.test(lastError)) break;
      if(attempt<2) await new Promise(resolve=>setTimeout(resolve,300*(attempt+1)));
    }
  }
  return {ok:false,date,blocked:false,error:lastError||"Request gagal setelah 3 percobaan"};
}

function pickProfileText_(obj, keys) {
  if (!obj || typeof obj !== "object") return "";
  const normalized = Object.fromEntries(Object.entries(obj).map(([k,v]) => [String(k).toLowerCase().replace(/[^a-z0-9]/g,""),v]));
  for (const key of keys) {
    const value = normalized[String(key).toLowerCase().replace(/[^a-z0-9]/g,"")];
    if (value !== undefined && value !== null && typeof value !== "object") {
      const textValue = String(value).trim();
      if (textValue && textValue !== "-" && textValue.toLowerCase() !== "null") return textValue;
    }
  }
  return "";
}

async function handleStockProfile(url, env) {
  const code = cleanStock(url.searchParams.get("code"));
  if (!/^[A-Z0-9]{1,4}$/.test(code)) return json({error:"Kode saham tidak valid."},400);
  await env.DB.prepare("CREATE TABLE IF NOT EXISTS stock_profile_cache (code TEXT PRIMARY KEY, payload TEXT NOT NULL, updated_at TEXT NOT NULL)").run();
  const force = url.searchParams.get("refresh") === "1";
  const cachedRow = await env.DB.prepare("SELECT payload,updated_at FROM stock_profile_cache WHERE code=?").bind(code).first();
  if (cachedRow && !force) {
    try {
      const cached = JSON.parse(cachedRow.payload);
      const age = Date.now() - Date.parse(cachedRow.updated_at || "");
      const ttl = cached.sector ? 30*24*60*60*1000 : 6*60*60*1000;
      if (Number.isFinite(age) && age >= 0 && age < ttl) return json(cached);
    } catch {}
  }
  let stockRow = null;
  try { stockRow = await env.DB.prepare("SELECT code,name,board,logo FROM stocks WHERE code=?").bind(code).first(); } catch {}
  let security = null;
  try { security = await lookupSecurity_(env, code); } catch (e) { console.warn("Stock profile lookup:", code, String(e.message || e)); }
  const name = pickProfileText_(security, ["Name","companyName","issuerName","company","issuer"]) || String(stockRow?.name || "");
  let sector = pickProfileText_(security, ["sector","idxSector","sectorName","sector_name","industrySector","businessSector","sectorCategory","classification"]);
  if (!sector) sector = pickProfileText_(security, ["industryGroup","industry_group","industry","subSector","subsector","subSectorName"]);
  const profile = {code, name: name || (code === "SINI" ? "PT Sinar Industri Tbk" : ""), sector, industry: sector, board: pickProfileText_(security, ["ListingBoard","listingBoard","board"]) || String(stockRow?.board || ""), updatedAt:new Date().toISOString()};
  try { await env.DB.prepare("INSERT INTO stock_profile_cache(code,payload,updated_at) VALUES(?,?,?) ON CONFLICT(code) DO UPDATE SET payload=excluded.payload,updated_at=excluded.updated_at").bind(code,JSON.stringify(profile),profile.updatedAt).run(); } catch (e) { console.warn("Stock profile cache save:", String(e.message || e)); }
  return json(profile);
}

async function handleFlowSummary(url, env, ctx) {
  const stock = cleanStock(url.searchParams.get("stock"));
  if (!/^[A-Z0-9]{1,4}$/.test(stock)) return json({error:"Kode saham tidak valid."},400);
  let end = parseDate(url.searchParams.get("endDate")) || latestCompletedDateWIB();
  const requested = Math.max(1, Math.min(5, Math.floor(num(url.searchParams.get("days")) || 5)));
  const force = url.searchParams.get("refresh") === "1";
  await env.DB.prepare("CREATE TABLE IF NOT EXISTS flow_summary_cache (stock TEXT NOT NULL,end_date TEXT NOT NULL,days INTEGER NOT NULL,payload TEXT NOT NULL,updated_at TEXT NOT NULL,PRIMARY KEY(stock,end_date,days))").run();
  const cachedRow = await env.DB.prepare("SELECT payload,updated_at FROM flow_summary_cache WHERE stock=? AND end_date=? AND days=?").bind(stock,end,requested).first();
  if (cachedRow?.payload && !force) {
    try {
      const cached = JSON.parse(cachedRow.payload);
      const age = Date.now() - Date.parse(cachedRow.updated_at || "");
      // Lewati cache schema lama yang menghitung total legend menjadi 0,00%.
      if (cached.schemaVersion === 5 && Array.isArray(cached.days) && Number.isFinite(age) && age >= 0 && age < 30*60*1000) return json(cached);
    } catch {}
  }

  let masterRows = [];
  try { masterRows = (await env.DB.prepare("SELECT code,type FROM broker_master").all()).results || []; } catch {}
  if (!masterRows.length) {
    try { await refreshBrokerMaster(env); masterRows = (await env.DB.prepare("SELECT code,type FROM broker_master").all()).results || []; }
    catch (e) { console.warn("Flow Summary broker master:", String(e.message || e)); }
  }
  const master = Object.fromEntries(masterRows.map(x => [String(x.code || "").toUpperCase(), String(x.type || "LOCAL").toUpperCase()]));
  const configuredRetailBrokers = await getRetailBrokers_(env);
  const retailSet = new Set(configuredRetailBrokers.map(x => String(x).toUpperCase()));
  const categoryFor = code => {
    const b = String(code || "").toUpperCase();
    const t = String(master[b] || "").toUpperCase();
    if (["CC","NI","OD","DX"].includes(b) || ["INSTITUTION","BUMN","BUMD"].includes(t)) return "bumn";
    if (retailSet.has(b)) return "retail";
    if (t === "FOREIGN" || t === "ASING" || t.includes("FOREIGN") || t.includes("ASING")) return "foreign";
    if (!t && ["AK","BK","DP","FS","HD","KZ","LG","MS","PD","SQ","YP","ZP"].includes(b)) return "foreign";
    return "local";
  };

  const dailyRows = [];
  const missingDates = [];
  const fetchErrors = [];
  let cursor = new Date(end + "T00:00:00Z");
  let scanned = 0;
  // Cari 5 hari perdagangan yang benar-benar memiliki data; akhir pekan/libur kosong dilewati.
  while (dailyRows.length < requested && scanned < 35) {
    const date = cursor.toISOString().slice(0,10);
    const day = cursor.getUTCDay();
    scanned++;
    cursor.setUTCDate(cursor.getUTCDate() - 1);
    if (day === 0 || day === 6) continue;

    let snapshot = null;
    try { snapshot = await getDailyFromDb(env, stock, date, true); }
    catch (e) { console.warn("Flow Summary DB read:", date, String(e.message || e)); }
    if (!hasBrokerRows_(snapshot)) {
      const fetched = await fetchDailyOnDemand(env, stock, date, true);
      if (!fetched.ok) {
        fetchErrors.push({date,error:fetched.error || "Data NET tidak berhasil diambil",blocked:!!fetched.blocked});
        if (fetched.blocked) break;
      }
      try { snapshot = await getDailyFromDb(env, stock, date, true); } catch {}
    }
    if (!hasBrokerRows_(snapshot)) {
      missingDates.push(date);
      continue;
    }

    const buyers = Array.isArray(snapshot.buyers) ? snapshot.buyers : [];
    const sellers = Array.isArray(snapshot.sellers) ? snapshot.sellers : [];
    const totalBuyLots = buyers.reduce((sum,x) => sum + Math.abs(num(x.lots)),0);
    const totalSellLots = sellers.reduce((sum,x) => sum + Math.abs(num(x.lots)),0);
    if (totalBuyLots <= 0 && totalSellLots <= 0) { missingDates.push(date); continue; }

    const netMap = signedNetMap_(snapshot);
    const union = new Set([...buyers.map(x=>String(x.broker||"").toUpperCase()),...sellers.map(x=>String(x.broker||"").toUpperCase())]);
    const brokerRows = [...union].filter(Boolean).map(code => {
      const buyRows = buyers.filter(x=>String(x.broker||"").toUpperCase()===code);
      const sellRows = sellers.filter(x=>String(x.broker||"").toUpperCase()===code);
      const buyLots = buyRows.reduce((sum,x)=>sum+Math.abs(num(x.lots)),0);
      const sellLots = sellRows.reduce((sum,x)=>sum+Math.abs(num(x.lots)),0);
      const buyValue = buyRows.reduce((sum,x)=>sum+Math.abs(num(x.value??x.netValue??x.amount??x.totalValue??x.total_value??0)),0);
      const sellValue = sellRows.reduce((sum,x)=>sum+Math.abs(num(x.value??x.netValue??x.amount??x.totalValue??x.total_value??0)),0);
      return {
        code,
        category:categoryFor(code),
        hold:num(netMap[code]?.lots),
        buyLots,
        sellLots,
        buyValue,
        sellValue,
        // Gunakan NET resmi per broker bila tersedia, termasuk nilai NET SELL negatif.
        netValue: netMap[code] && Number.isFinite(Number(netMap[code].value))
          ? num(netMap[code].value)
          : buyValue-sellValue
      };
    });

    function categoryPct(type) {
      // Ambang NET > Rp20 juta hanya berlaku untuk Bandar Flow (broker lokal).
      // Retail dan Foreign tetap dihitung tanpa filter nilai NET.
      const isAllowed = row => {
        if (type === "retail") return row.category === "retail";
        if (type === "foreign") return row.category === "foreign";
        // Ambang berlaku dua arah: NET BUY atau NET SELL dengan nilai absolut > Rp20 juta.
        return (row.category === "local" || row.category === "bumn") && Math.abs(num(row.netValue)) > 20000000;
      };
      // Selaras dengan tampilan Flow: ambil maksimal lima broker per sisi untuk kategori terkait.
      const topBuy = brokerRows.filter(row => isAllowed(row) && row.buyLots > 0).sort((a,b)=>b.buyLots-a.buyLots).slice(0,5);
      const topSell = brokerRows.filter(row => isAllowed(row) && row.sellLots > 0).sort((a,b)=>b.sellLots-a.sellLots).slice(0,5);
      const groupBuyLots = topBuy.reduce((sum,row)=>sum+row.buyLots,0);
      const groupSellLots = topSell.reduce((sum,row)=>sum+row.sellLots,0);
      const buyPct = totalBuyLots > 0 ? groupBuyLots / totalBuyLots * 100 : 0;
      const sellPct = totalSellLots > 0 ? groupSellLots / totalSellLots * 100 : 0;
      return {pct:buyPct-sellPct,buyPct,sellPct,buyLots:groupBuyLots,sellLots:groupSellLots};
    }
    const retail = categoryPct("retail"), bandar = categoryPct("local"), foreign = categoryPct("foreign");
    dailyRows.push({date,retailPct:retail.pct,bandarPct:bandar.pct,foreignPct:foreign.pct,
      retailBuyPct:retail.buyPct,retailSellPct:-retail.sellPct,
      bandarBuyPct:bandar.buyPct,bandarSellPct:-bandar.sellPct,
      foreignBuyPct:foreign.buyPct,foreignSellPct:-foreign.sellPct,
      retailBuyLots:retail.buyLots,retailSellLots:retail.sellLots,
      bandarBuyLots:bandar.buyLots,bandarSellLots:bandar.sellLots,
      foreignBuyLots:foreign.buyLots,foreignSellLots:foreign.sellLots,
      totalBuyLots,totalSellLots,capped:snapshot.capped===true});
  }
  dailyRows.sort((a,b)=>a.date.localeCompare(b.date));
  // Ringkasan di legend menampilkan persentase untuk tanggal perdagangan terakhir,
  // bukan agregasi lima hari. Ini menghindari nilai 0,00% akibat menjumlahkan field
  // lot yang sebelumnya tidak disimpan pada dailyRows.
  const latestRow = dailyRows[dailyRows.length - 1] || null;
  const latest = latestRow ? {
    retailPct:num(latestRow.retailPct),
    bandarPct:num(latestRow.bandarPct),
    foreignPct:num(latestRow.foreignPct)
  } : {retailPct:0,bandarPct:0,foreignPct:0};
  const sum = key => dailyRows.reduce((total,row)=>total + num(row[key]),0);
  const totalBuyLots = sum("totalBuyLots"), totalSellLots = sum("totalSellLots");
  const payload = {
    schemaVersion:5,
    stock,startDate:dailyRows[0]?.date||"",endDate:dailyRows[dailyRows.length-1]?.date||end,days:dailyRows,
    totals:latest,latest,latestDate:latestRow?.date||"",
    periodTotals:{
      retailPct:(totalBuyLots>0?sum("retailBuyLots")/totalBuyLots*100:0)-(totalSellLots>0?sum("retailSellLots")/totalSellLots*100:0),
      bandarPct:(totalBuyLots>0?sum("bandarBuyLots")/totalBuyLots*100:0)-(totalSellLots>0?sum("bandarSellLots")/totalSellLots*100:0),
      foreignPct:(totalBuyLots>0?sum("foreignBuyLots")/totalBuyLots*100:0)-(totalSellLots>0?sum("foreignSellLots")/totalSellLots*100:0)
    },
    missingDates,fetchErrors,dataWarning:(dailyRows.some(x=>x.capped)?"Sebagian respons harian ZAPI dibatasi Top 10.":"") + (missingDates.length?" Ada tanggal tanpa data harian.":""),
    generatedAt:new Date().toISOString(),source:"NET_DAILY_TOP5"
  };
  if (dailyRows.length) {
    try { await env.DB.prepare("INSERT INTO flow_summary_cache(stock,end_date,days,payload,updated_at) VALUES(?,?,?,?,?) ON CONFLICT(stock,end_date,days) DO UPDATE SET payload=excluded.payload,updated_at=excluded.updated_at").bind(stock,end,requested,JSON.stringify(payload),payload.generatedAt).run(); }
    catch (e) { console.warn("Flow Summary cache save:",String(e.message||e)); }
  }
  return json(payload);
}

async function handleStocks(url,env){
  const q=cleanStock(String(url.searchParams.get("q")||"").trim().toUpperCase().slice(0,5));
  if(!q)return json([]);
  const rows=await env.DB.prepare("SELECT code,name,board,logo FROM stocks WHERE code LIKE ? OR UPPER(name) LIKE ? ORDER BY CASE WHEN code LIKE ? THEN 0 ELSE 1 END,code LIMIT 30")
    .bind(q+"%", "%"+q+"%", q+"%").all();
  let result=rows.results||[];
  // Jika ticker tepat belum ada di database, cari ke endpoint sekali lalu simpan master-nya.
  if(!result.some(x=>String(x.code||"").toUpperCase()===q) && /^[A-Z0-9]{1,4}$/.test(q)){
    const found=await lookupSecurity_(env,q);
    if(found){
      await saveStockMaster_(env,found);
      const row=await env.DB.prepare("SELECT code,name,board,logo FROM stocks WHERE code=?").bind(q).first();
      if(row)result=[row,...result.filter(x=>String(x.code||"").toUpperCase()!==q)];
    }
  }
  return json(result);
}
async function handleDashboard(url,env,ctx){
  const stock=cleanStock(url.searchParams.get("stock"));
  if(!/^[A-Z0-9]{1,4}$/.test(stock))return json({error:"Kode saham tidak valid."},400);
  let start=parseDate(url.searchParams.get("startDate")),end=parseDate(url.searchParams.get("endDate"));
  if(!start)start=latestCompletedDateWIB();
  if(!end)end=start;
  if(start>end)[start,end]=[end,start];

  const useNet=url.searchParams.get("net")==="true";
  const forceRefresh=url.searchParams.get("refresh")==="1";
  let stockRow=null;
  try{stockRow=await env.DB.prepare("SELECT code,name,board,logo FROM stocks WHERE code=?").bind(stock).first();}catch{}
  if(!stockRow){
    try{
      const found=await lookupSecurity_(env,stock);
      if(found){await saveStockMaster_(env,found);stockRow=await env.DB.prepare("SELECT code,name,board,logo FROM stocks WHERE code=?").bind(stock).first();}
    }catch(e){console.error("Stock lookup:",String(e.message||e));}
  }

  // Penting: Broker Summary menggunakan request agregat ZAPI untuk SELURUH periode,
  // bukan penjumlahan hasil harian yang masing-masing dibatasi Top 10.
  let periodData;
  try{
    periodData=await getRangeBrokerSummary_(env,stock,start,end,useNet,forceRefresh);
  }catch(e){
    return json({
      error:"Gagal mengambil Broker Summary agregat dari ZAPI untuk periode "+start+" sampai "+end+". Data harian lama tidak dipakai sebagai pengganti karena hasilnya dapat kehilangan transaksi broker yang tidak masuk Top 10 harian. Detail: "+String(e.message||e).slice(0,350),
      stock,startDate:start,endDate:end,net:useNet,rows:[],dataComplete:false
    },502);
  }

  let netData=null,netWarning="";
  if(!useNet){
    try{netData=await getRangeBrokerSummary_(env,stock,start,end,true,forceRefresh);}
    catch(e){netWarning="Data NET periode tidak tersedia untuk melengkapi SELL yang tidak masuk Top 10 GROSS: "+String(e.message||e).slice(0,250);}
  }

  let masterRows=(await env.DB.prepare("SELECT code,type FROM broker_master").all()).results||[];
  if(!masterRows.length){try{await refreshBrokerMaster(env);}catch(e){console.error("Broker master refresh:",String(e.message||e));}masterRows=(await env.DB.prepare("SELECT code,type FROM broker_master").all()).results||[];}
  const master=Object.fromEntries(masterRows.map(x=>[String(x.code).toUpperCase(),String(x.type).toUpperCase()]));
  function categoryFor_(broker){
    const b=String(broker||"").toUpperCase();
    const t=String(master[b]||"LOCAL").toUpperCase();
    return (["CC","NI","OD","DX"].includes(b)||["INSTITUTION","BUMN","BUMD"].includes(t))?"BUMN":(t==="FOREIGN"||t==="ASING"?"FOREIGN":"LOCAL");
  }

  let rows=[];
  if(useNet){
    const signed=signedNetMap_(periodData);
    rows=Object.values(signed).map(item=>({
      broker:item.broker,
      netLots:item.lots,
      netValue:item.value,
      netAvg:item.lots?Math.abs(item.value/(item.lots*100)):0,
      category:categoryFor_(item.broker)
    }));
    rows.sort((a,b)=>Math.abs(Number(b.netLots||0))-Math.abs(Number(a.netLots||0)));
  }else{
    const rowMap={};
    const ensureRow_=broker=>{
      const b=String(broker||"-").toUpperCase();
      rowMap[b] ||= {broker:b,buyLots:0,buyValue:0,buyAvg:0,sellLots:0,sellValue:0,sellAvg:0,sellDerived:false};
      return rowMap[b];
    };
    for(const item of (periodData.buyers||[])){
      const r=ensureRow_(item.broker);
      r.buyLots+=num(item.lots);r.buyValue+=num(item.value);
    }
    for(const item of (periodData.sellers||[])){
      const r=ensureRow_(item.broker);
      r.sellLots+=num(item.lots);r.sellValue+=num(item.value);
    }

    // Bila broker ada di GROSS BUY tetapi tidak ada di GROSS SELL Top 10,
    // SELL = GROSS BUY - signed NET pada rentang tanggal yang sama.
    const grossSellerCodes=new Set((periodData.sellers||[]).map(x=>String(x.broker||"-").toUpperCase()));
    const signedNet=netData?signedNetMap_(netData):{};
    if(netData){
      for(const item of (periodData.buyers||[])){
        const broker=String(item.broker||"-").toUpperCase();
        if(grossSellerCodes.has(broker)||!Object.prototype.hasOwnProperty.call(signedNet,broker))continue;
        const r=ensureRow_(broker),n=signedNet[broker];
        const inferredLots=Math.max(0,r.buyLots-n.lots);
        const inferredValue=Math.max(0,r.buyValue-n.value);
        if(inferredLots>0){r.sellLots=inferredLots;r.sellValue=inferredValue;r.sellDerived=true;}
      }
    }

    rows=Object.values(rowMap).map(r=>{
      // Rata-rata berbobot dihitung ulang dari total nilai dan total lot;
      // jangan gunakan averagePrice dari hari terakhir sebagai rata-rata periode.
      r.buyAvg=r.buyLots>0?r.buyValue/(r.buyLots*100):0;
      r.sellAvg=r.sellLots>0?r.sellValue/(r.sellLots*100):0;
      r.category=categoryFor_(r.broker);
      return r;
    });
    rows.sort((a,b)=>(b.buyValue+b.sellValue)-(a.buyValue+a.sellValue));
  }

  const staleWarning=periodData.staleCache?"Menggunakan cache periode sebelumnya karena refresh ZAPI gagal.":"";
  const cacheWarning=(!useNet && netData && netData.staleCache)?"Data NET berasal dari cache periode sebelumnya.":"";
  const warnings=[(periodData.capped||(netData&&netData.capped))?"ZAPI menandai hasil periode ini sebagai terbatas (capped).":"",staleWarning,cacheWarning,netWarning].filter(Boolean);
  const logo=stockRow?.logo||"";
  const header={code:stock,name:stockRow?.name||"",sector:"",industry:"",subSector:"",listingBoard:stockRow?.board||"",price:0,change:0,previous:0,changePercent:0,logo,stockbitLogo:logo,updatedAt:periodData.updatedAt||new Date().toISOString()};
  return json({header,startDate:start,endDate:end,latestDate:end,net:useNet,capped:periodData.capped===true,count:rows.length,rows,
    dataComplete:useNet?true:!!netData,missingDates:[],missingNetDates:netWarning?[start+"–"+end]:[],dataWarning:warnings.join(" "),fetchErrors:[],
    cachedDates:[start+"–"+end],flowSource:null,retailFlow:[],bandarFlow:[],foreignFlow:[],ownershipBrokers:[],closePrice:0,
    dataSource:"ZAPI_RANGE_AGGREGATE",rangeCacheStale:!!periodData.staleCache});
}

const BROKER_RANGE_CACHE_TTL_MS = 6 * 60 * 60 * 1000;

async function getRangeBrokerSummary_(env,stock,start,end,useNet,forceRefresh=false){
  await env.DB.prepare("CREATE TABLE IF NOT EXISTS broker_period_cache (stock TEXT NOT NULL,start_date TEXT NOT NULL,end_date TEXT NOT NULL,mode TEXT NOT NULL,payload TEXT NOT NULL,updated_at TEXT NOT NULL,PRIMARY KEY(stock,start_date,end_date,mode))").run();
  const mode=useNet?"net":"gross";
  const cachedRow=await env.DB.prepare("SELECT payload,updated_at FROM broker_period_cache WHERE stock=? AND start_date=? AND end_date=? AND mode=?")
    .bind(stock,start,end,mode).first();
  let cached=null;
  if(cachedRow?.payload){
    try{cached=JSON.parse(cachedRow.payload);cached.updatedAt=cachedRow.updated_at||cached.updatedAt||"";}catch{cached=null;}
  }
  const age=cachedRow?.updated_at?Date.now()-Date.parse(cachedRow.updated_at):Infinity;
  if(cached&&!forceRefresh&&Number.isFinite(age)&&age>=0&&age<BROKER_RANGE_CACHE_TTL_MS)return cached;

  const path=API_BROKER+"?code="+encodeURIComponent(stock)+"&startDate="+encodeURIComponent(start)+"&endDate="+encodeURIComponent(end)+"&net="+(useNet?"true":"false");
  try{
    const raw=await zapi(env,path);
    const result=normalizeBrokerSummary(raw,stock,start,useNet);
    result.code=stock;result.startDate=start;result.endDate=end;result.useNet=useNet;result.updatedAt=new Date().toISOString();
    if(!hasBrokerRows_(result))throw new Error("Respons ZAPI tidak berisi data BUY/SELL untuk rentang ini.");
    await env.DB.prepare("INSERT INTO broker_period_cache(stock,start_date,end_date,mode,payload,updated_at) VALUES(?,?,?,?,?,?) ON CONFLICT(stock,start_date,end_date,mode) DO UPDATE SET payload=excluded.payload,updated_at=excluded.updated_at")
      .bind(stock,start,end,mode,JSON.stringify(result),result.updatedAt).run();
    return result;
  }catch(e){
    if(cached){cached.staleCache=true;cached.refreshError=String(e.message||e).slice(0,250);return cached;}
    throw e;
  }
}

async function handleFlowSource(url,env,ctx){
  const stock=cleanStock(url.searchParams.get("stock"));
  if(!/^[A-Z0-9]{1,4}$/.test(stock))return json({error:"Kode saham tidak valid."},400);
  let start=parseDate(url.searchParams.get("startDate")),end=parseDate(url.searchParams.get("endDate"));
  if(!start)start=latestCompletedDateWIB();
  if(!end)end=start;
  if(start>end)[start,end]=[end,start];
  const forceRefresh=url.searchParams.get("refresh")==="1";

  let gross,net;
  try{
    [gross,net]=await Promise.all([
      getRangeBrokerSummary_(env,stock,start,end,false,forceRefresh),
      getRangeBrokerSummary_(env,stock,start,end,true,forceRefresh)
    ]);
  }catch(e){
    return json({error:"Data agregat GROSS/NET untuk Flow belum tersedia pada periode "+start+" sampai "+end+". "+String(e.message||e).slice(0,350),stock,startDate:start,endDate:end,dataComplete:false},502);
  }

  const signedNet=signedNetMap_(net);
  const grossSellerCodes=new Set((gross.sellers||[]).map(x=>String(x.broker||"-").toUpperCase()));
  const grossBuyByCode=new Map((gross.buyers||[]).map(x=>[String(x.broker||"-").toUpperCase(),x]));
  const inferredSellers=[];
  // Bandar Flow SELL: utamakan GROSS SELL dari ZAPI. Jika broker tidak muncul
  // pada daftar GROSS SELL, turunkan SELL dari GROSS BUY - NET bertanda.
  // Broker NET SELL yang hanya muncul di NET juga tetap ditampilkan sebagai SELL.
  const sellerCandidates=new Set([
    ...(gross.buyers||[]).map(x=>String(x.broker||"-").toUpperCase()),
    ...Object.keys(signedNet)
  ]);
  for(const broker of sellerCandidates){
    if(grossSellerCodes.has(broker)||!Object.prototype.hasOwnProperty.call(signedNet,broker))continue;
    const n=signedNet[broker];
    const buy=grossBuyByCode.get(broker);
    const buyLots=buy?Math.abs(num(buy.lots)):0;
    const buyValue=buy?Math.abs(num(buy.value)):0;
    let lots=Math.max(0,buyLots-n.lots);
    let value=Math.max(0,buyValue-n.value);
    if(!buy && n.lots<0){lots=Math.abs(n.lots);value=Math.abs(n.value);}
    if(lots>0||value>0)inferredSellers.push({broker,lots,value,averagePrice:lots?value/(lots*100):0,derived:true});
  }

  const netRows=Object.values(signedNet).map(item=>({broker:item.broker,lots:item.lots,value:item.value,
    averagePrice:item.lots?Math.abs(item.value/(item.lots*100)):0}));
  let masterRows=(await env.DB.prepare("SELECT code,type FROM broker_master").all()).results||[];
  if(!masterRows.length){try{await refreshBrokerMaster(env);}catch(e){console.error("Broker master refresh:",String(e.message||e));}masterRows=(await env.DB.prepare("SELECT code,type FROM broker_master").all()).results||[];}
  const master=Object.fromEntries(masterRows.map(x=>[String(x.code).toUpperCase(),String(x.type).toUpperCase()]));
  let stockRow=null;
  try{stockRow=await env.DB.prepare("SELECT code,name,board,logo FROM stocks WHERE code=?").bind(stock).first();}catch{}
  const header={code:stock,name:stockRow?.name||"",sector:"",industry:"",subSector:"",listingBoard:stockRow?.board||"",price:0,change:0,previous:0,changePercent:0,logo:stockRow?.logo||"",stockbitLogo:stockRow?.logo||"",updatedAt:gross.updatedAt||new Date().toISOString()};
  const warnings=[];
  if(gross.capped||net.capped)warnings.push("ZAPI menandai salah satu hasil rentang sebagai terbatas (capped); broker di luar hasil API mungkin tidak tersedia.");
  if(gross.staleCache||net.staleCache)warnings.push("Sebagian data Flow memakai cache periode sebelumnya karena refresh ZAPI gagal.");
  return json({stock,startDate:start,endDate:end,buyers:gross.buyers||[],sellers:gross.sellers||[],inferredSellers,netRows,brokerMaster:master,
    retailBrokers:await getRetailBrokers_(env),closePrice:0,header,cachedDates:[start+"–"+end],dataComplete:true,missingDates:[],missingNetDates:[],dataWarning:warnings.join(" "),fetchErrors:[],
    dataSource:"ZAPI_RANGE_AGGREGATE"});
}

function signedNetMap_(data){
  const map={};
  function ensure_(broker){
    const code=String(broker||"-").toUpperCase();
    map[code] ||= {broker:code,lots:0,value:0};
    return map[code];
  }
  for(const item of (data?.buyers||[])){
    const row=ensure_(item.broker);
    row.lots+=Math.abs(num(item.lots));
    row.value+=Math.abs(num(item.value));
  }
  for(const item of (data?.sellers||[])){
    const row=ensure_(item.broker);
    row.lots-=Math.abs(num(item.lots));
    row.value-=Math.abs(num(item.value));
  }
  return map;
}

function aggregateNetSnapshots_(snapshots){
  const map={};
  for(const snap of snapshots){
    const daily=signedNetMap_(snap.data);
    for(const [code,item] of Object.entries(daily)){
      map[code] ||= {broker:code,lots:0,value:0};
      map[code].lots+=item.lots;
      map[code].value+=item.value;
    }
  }
  return Object.values(map).map(item=>({
    broker:item.broker,
    lots:item.lots,
    value:item.value,
    averagePrice:item.lots?Math.abs(item.value/(item.lots*100)):0
  }));
}
function aggregateSide_(items){
  const m={};for(const x of items){const b=String(x.broker||"-").toUpperCase();m[b] ||= {broker:b,lots:0,value:0,averagePrice:0};m[b].lots+=num(x.lots);m[b].value+=num(x.value);}
  for(const x of Object.values(m))x.averagePrice=x.lots?x.value/(x.lots*100):0;
  return Object.values(m);
}
function category(code){
  if(["CC","NI","OD","DX"].includes(code))return "BUMN";
  // Foreign/local classification is hydrated from the broker master table during ingestion when available.
  return ["AK","BK","DP","FS","HD","KZ","LG","MS","PD","SQ","YP","ZP"].includes(code)?"FOREIGN":"LOCAL";
}
async function handleStockbitQuote(url,env){
  const symbol=cleanStock(url.searchParams.get("symbol"));
  if(!symbol || !/^[A-Z0-9]{1,4}$/.test(symbol)) return json({error:"Kode saham tidak valid."},400);
  try{
    let data=await zapi(env,API_STOCKBIT_SYMBOL+"?symbol="+encodeURIComponent(symbol));
    for(let i=0;i<3;i++){const nested=data&&(data.data||data.quote||data.result);if(!nested||nested===data)break;data=nested;}
    return json({symbol, last:data.last ?? data.price ?? null, change:data.change ?? data.changeValue ?? null,
      changePercent:data.changePercent ?? data.change_percent ?? null, iconUrl:data.iconUrl || data.logoUrl || data.logo || "",
      name:data.name || "", provider:data.provider || "stockbit", updatedAt:data.updatedAt || ""},200);
  }catch(e){
    const message=String(e.message||e).slice(0,300);
    return json({error:message,symbol},/429|quota exceeded|monthly quota/i.test(message)?429:502);
  }
}
async function handlePluangQuote(url,env){
  // Endpoint tetap kompatibel dengan frontend lama. Prioritaskan quote Stockbit
  // melalui ZAPI agar harga dan perubahan mengikuti data quote pasar Indonesia.
  // Yahoo hanya menjadi fallback jika Stockbit tidak mengembalikan quote valid.
  const code=cleanStock(url.searchParams.get("code"));
  if(!code || !/^[A-Z0-9]{1,4}$/.test(code)) return json({error:"Kode saham tidak valid."},400);
  try{
    let data=await zapi(env,API_STOCKBIT_SYMBOL+"?symbol="+encodeURIComponent(code));
    for(let i=0;i<3;i++){const nested=data&&(data.data||data.quote||data.result);if(!nested||nested===data)break;data=nested;}
    const price=Number(data.last ?? data.lastPrice ?? data.price ?? data.regularMarketPrice);
    const changeRaw=data.change ?? data.changeValue ?? data.changePoint ?? data.priceChange;
    const pctRaw=data.changePercent ?? data.change_percent ?? data.changePct ?? data.percentChange;
    const change=Number(changeRaw),changePercent=Number(pctRaw);
    if(Number.isFinite(price)&&price>0&&changeRaw!==undefined&&changeRaw!==null&&changeRaw!==""&&Number.isFinite(change)&&pctRaw!==undefined&&pctRaw!==null&&pctRaw!==""&&Number.isFinite(changePercent)){
      return json({code,price,change,changePercent,previousClose:price-change,provider:"stockbit-via-zapi",updatedAt:data.updatedAt||data.timestamp||""});
    }
  }catch(e){
    // Lanjutkan ke fallback Yahoo bila quote Stockbit gagal/tidak lengkap.
  }
  const yahooSymbol=code+".JK";
  try{
    const endpoint="https://query1.finance.yahoo.com/v8/finance/chart/"+encodeURIComponent(yahooSymbol)+"?range=5d&interval=1d&events=history";
    const response=await fetch(endpoint,{
      headers:{"Accept":"application/json","User-Agent":"Mozilla/5.0 stock-dashboard/1.0"},
      signal:AbortSignal.timeout(6000)
    });
    const body=await response.text();
    let payload;
    try{payload=JSON.parse(body)}catch{throw new Error("Yahoo Finance mengembalikan respons yang bukan JSON.")}
    if(!response.ok) throw new Error("Yahoo Finance HTTP "+response.status);
    const result=payload?.chart?.result?.[0];
    const yahooError=payload?.chart?.error;
    if(yahooError) throw new Error(String(yahooError.description||yahooError.code||"Yahoo Finance gagal mengambil harga."));
    if(!result) return json({code,error:"Harga saham "+yahooSymbol+" tidak ditemukan di Yahoo Finance."},404);
    const meta=result.meta||{};
    const closes=(result.indicators?.quote?.[0]?.close||[]).filter(v=>Number.isFinite(Number(v))).map(Number);
    const priceRaw=meta.regularMarketPrice ?? closes[closes.length-1];
    const previousRaw=meta.chartPreviousClose ?? meta.previousClose ?? (closes.length>1?closes[closes.length-2]:null);
    const price=Number(priceRaw), previous=Number(previousRaw);
    if(!Number.isFinite(price)||price<=0) return json({code,error:"Last Price tidak tersedia dari Yahoo Finance."},502);
    const hasPrevious=Number.isFinite(previous)&&previous>0;
    const change=hasPrevious?price-previous:null;
    const changePercent=hasPrevious?change/previous*100:null;
    const updatedAt=Number.isFinite(Number(meta.regularMarketTime))?new Date(Number(meta.regularMarketTime)*1000).toISOString():"";
    return new Response(JSON.stringify({code,price,change,changePercent,previousClose:hasPrevious?previous:null,provider:"yahoo-finance",yahooSymbol,updatedAt}),{
      status:200,headers:{"Content-Type":"application/json; charset=utf-8","Cache-Control":"public, max-age=10","Access-Control-Allow-Origin":"*"}
    });
  }catch(e){
    const message=String(e.message||e).slice(0,300);
    return json({error:message,code,provider:"yahoo-finance"},/timeout|abort/i.test(message)?504:502);
  }
}
async function handleBrokerHistory(url,env){
  const stock=cleanStock(url.searchParams.get("stock")),broker=String(url.searchParams.get("broker")||"").toUpperCase();
  let start=parseDate(url.searchParams.get("startDate")),end=parseDate(url.searchParams.get("endDate"))||start;
  if(!stock||!broker||!start)return json({stock,broker,rows:[],missingDates:[]});
  if(start>end)[start,end]=[end,start];

  const dates=dateRange(start,end).filter(weekdaysOnly);
  const missing=[];
  for(const d of dates){if(!hasBrokerRows_(await getDailyFromDb(env,stock,d,false)))missing.push(d);}
  // History Broker selalu menggunakan payload gross (net=false). Batasi request on-demand
  // agar rentang panjang tidak membuat lonjakan pemanggilan endpoint.
  const queue=missing.slice(0,20),fetchErrors=[];
  let cursor=0;
  async function runner(){while(cursor<queue.length){const date=queue[cursor++];const result=await fetchDailyOnDemand(env,stock,date,false);if(!result.ok)fetchErrors.push({date,error:result.error||"Request gagal",blocked:!!result.blocked});}}
  if(queue.length)await Promise.all(Array.from({length:Math.min(3,queue.length)},runner));

  const rows=[],missingAfter=[];
  for(const d of dates){
    const data=await getDailyFromDb(env,stock,d,false);
    if(!hasBrokerRows_(data)){missingAfter.push(d);continue;}
    const buy=(data.buyers||[]).find(x=>String(x.broker||"").toUpperCase()===broker);
    const sell=(data.sellers||[]).find(x=>String(x.broker||"").toUpperCase()===broker);
    if(buy||sell){const bl=num(buy?.lots),bv=num(buy?.value),sl=num(sell?.lots),sv=num(sell?.value),nl=bl-sl,nv=bv-sv;rows.push({date:d,buyLots:bl,buyValue:bv,buyAvg:bl?bv/(bl*100):0,sellLots:sl,sellValue:sv,sellAvg:sl?sv/(sl*100):0,netLots:nl,netValue:nv,netAvg:nl?Math.abs(nv/nl/100):0,action:nl>0?"BUY":nl<0?"SELL":"HOLD"});}
  }
  return json({stock,broker,startDate:start,endDate:end,rows,missingDates:missingAfter,fetchErrors,source:"CLOUDFLARE_D1_GROSS"});
}
const IDX_OWNERSHIP_PAGE = "https://www.idx.co.id/id/perusahaan-tercatat/data-kepemilikan-saham/";
const IDX_OWNERSHIP_CATEGORY = "lima-persen";
// Cache created by older ownership parsers must be rebuilt once after this release.
const IDX_OWNERSHIP_PARSER_RELEASE = "2026-10-21T14:00:00+07:00";
const IDX_BROKER_SUMMARY_PAGE = "https://www.idx.co.id/id/data-pasar/ringkasan-perdagangan/ringkasan-broker";
const IDX_BROKER_PROFILE_PAGE = "https://www.idx.id/id/anggota-bursa-dan-partisipan/profil-anggota-bursa/";

function decodeIdxEscapes_(value) {
  return String(value || "").replace(/\\u002F/gi, "/").replace(/\\u0026/gi, "&").replace(/\\u003A/gi, ":").replace(/\\\//g, "/").replace(/&amp;/g, "&");
}
function parseIdxDate_(text) {
  const s = String(text || "");
  let m = s.match(/(\d{1,2})\s+(Januari|Februari|Maret|April|Mei|Juni|Juli|Agustus|September|Oktober|November|Desember)\s+(\d{4})/i);
  if (m) {
    const months = {januari:"01",februari:"02",maret:"03",april:"04",mei:"05",juni:"06",juli:"07",agustus:"08",september:"09",oktober:"10",november:"11",desember:"12"};
    return m[3] + "-" + months[m[2].toLowerCase()] + "-" + String(m[1]).padStart(2,"0");
  }
  m = s.match(/(20\d{2})[-_](\d{2})[-_](\d{2})/);
  return m ? `${m[1]}-${m[2]}-${m[3]}` : "";
}
function latestIdxOwnershipFile_(html) {
  const source = decodeIdxEscapes_(html);
  const matches = [];
  // Nuxt page payload contains Description/Prospectus pairs. Restrict to the >5% series.
  const pairRe = /Description\s*:\s*"([^"]+)"[\s\S]{0,1200}?Prospectus\s*:\s*"([^"]+\.xlsx(?:\?[^"]*)?)"/g;
  let m;
  while ((m = pairRe.exec(source))) {
    const description = m[1];
    const url = decodeIdxEscapes_(m[2]);
    if (!/pemegang\s+saham\s+di\s+atas\s+5%/i.test(description) && !/lima-persen\.xlsx/i.test(url)) continue;
    if (!/lima-persen\.xlsx(?:\?|$)/i.test(url)) continue;
    const publishedAt = parseIdxDate_(url) || parseIdxDate_(description);
    if (!publishedAt || !/^https:\/\/www\.idx\.co\.id\//i.test(url)) continue;
    matches.push({description, url, publishedAt, fileName: url.split("/").pop().split("?")[0]});
  }
  matches.sort((a,b)=>b.publishedAt.localeCompare(a.publishedAt));
  return matches[0] || null;
}
function normalizeHeader_(value) {
  return String(value ?? "").normalize("NFKD").replace(/[\u0300-\u036f]/g,"").toLowerCase().replace(/[^a-z0-9%]+/g," ").trim();
}
function parsePercent_(value) {
  if (typeof value === "number") return value > 0 && value <= 1 ? value * 100 : value;
  const s = String(value ?? "").replace(/%/g,"").replace(/\s/g,"").replace(/\.(?=\d{3}(?:\D|$))/g,"").replace(/,/g,".");
  const n = Number(s);
  return Number.isFinite(n) ? (n > 0 && n <= 1 ? n * 100 : n) : 0;
}
function workbookRows_(bytes, requestedStock) {
  const wb = XLSX.read(bytes, {type:"array", cellDates:false, dense:false});
  const wanted = cleanStock(requestedStock);
  for (const sheetName of wb.SheetNames) {
    const matrix = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], {header:1, defval:null, raw:true});
    for (let hi=0; hi<Math.min(matrix.length,40); hi++) {
      const headers = (matrix[hi]||[]).map((_,ci)=>[0,1,2].map(off=>normalizeHeader_((matrix[hi+off]||[])[ci])).filter(Boolean).join(" "));
      const stockIdx = headers.findIndex(h => /\b(kode efek|kode saham|kode emiten|ticker|stock code)\b/.test(h));
      if (stockIdx < 0) continue;
      const issuerIdx = headers.findIndex(h=>/nama emiten|nama perusahaan|issuer/.test(h));
      const accountHolderIdx = headers.findIndex(h=>/nama pemegang rekening efek|nama perusahaan efek|perusahaan efek|nama bank kustodian|bank kustodian|kustodian/.test(h));
      const shareholderIdx = headers.findIndex(h=>/nama pemegang saham/.test(h));
      const statusIdx = headers.findIndex(h=>/status lokal asing|status/.test(h));
      const combinedShareIdx = headers.map((h,i)=>/saham gabungan per investor/.test(h)?i:-1).filter(i=>i>=0);
      const pctIdxs = headers.map((h,i)=>/persentase kepemilikan per investor/.test(h)?i:-1).filter(i=>i>=0);
      const changeIdx = headers.findIndex(h=>/^perubahan$/.test(h.trim()) || /^perubahan jumlah saham$/.test(h.trim()));
      if (shareholderIdx < 0 || combinedShareIdx.length < 1) continue;
      const sharesIdx = combinedShareIdx[combinedShareIdx.length-1];
      const previousSharesIdx = combinedShareIdx.length>1 ? combinedShareIdx[combinedShareIdx.length-2] : -1;
      const pctIdx = pctIdxs.length ? pctIdxs[pctIdxs.length-1] : -1;
      const groups = new Map(); let lastStock="", lastShareholder="", lastIssuer="";
      for(let ri=hi+1;ri<matrix.length;ri++){
        const cells=matrix[ri]||[];
        const cellStock=cleanStock(cells[stockIdx]);
        if(/^[A-Z0-9]{1,4}$/.test(cellStock))lastStock=cellStock;
        const stock=lastStock;
        if(!stock || stock!==wanted) continue;
        const issuer=String(issuerIdx>=0?cells[issuerIdx]??"":"").trim(); if(issuer)lastIssuer=issuer;
        const shareholderCell=String(cells[shareholderIdx]??"").trim(); if(shareholderCell)lastShareholder=shareholderCell;
        const shareholder=shareholderCell||lastShareholder;
        if(!shareholder)continue;
        const accountHolder=String(accountHolderIdx>=0?cells[accountHolderIdx]??"":"").trim();
        const key=stock+"\u0000"+shareholder.toUpperCase();
        if(!groups.has(key))groups.set(key,{stock,issuer:lastIssuer,holder:shareholder,brokerNames:[],status:"L",ownershipPct:0,shares:0,previousShares:null,totalChange:0});
        const g=groups.get(key);
        if(issuer)g.issuer=issuer;
        if(accountHolder && !g.brokerNames.includes(accountHolder))g.brokerNames.push(accountHolder);
        const currentShares=num(cells[sharesIdx]);
        if(currentShares>0)g.shares=currentShares;
        if(previousSharesIdx>=0){const previousShares=num(cells[previousSharesIdx]);if(previousShares>0)g.previousShares=previousShares;}
        if(pctIdx>=0){const pct=parsePercent_(cells[pctIdx]);if(pct>0)g.ownershipPct=pct;}
        if(changeIdx>=0)g.totalChange+=num(cells[changeIdx]);
        const status=String(statusIdx>=0?cells[statusIdx]??"":"").trim().toUpperCase();
        if(status)g.status=status.startsWith("A")?"A":"L";
      }
      const rows=[...groups.values()].filter(g=>g.shares>0 && g.ownershipPct>5).map(g=>({
        stock:g.stock,issuer:g.issuer,holder:g.holder,brokerNames:g.brokerNames,status:g.status,
        ownershipPct:g.ownershipPct,shares:g.shares,previousShares:g.previousShares,
        changePct:g.previousShares>0 && Number.isFinite(g.totalChange)?(g.totalChange/g.previousShares)*100:null,sourceChange:null
      }));
      if(rows.length)return rows;
    }
  }
  throw new Error("Ticker "+wanted+" tidak ditemukan pada kolom saham/pemegang saham di workbook IDX >5%.");
}
async function ensureIdxOwnershipTables_(env) {
  // Hanya simpan ticker yang benar-benar diminta pengguna. Metadata juga per ticker.
  await env.DB.prepare("CREATE TABLE IF NOT EXISTS idx_ownership_stock_meta (stock TEXT PRIMARY KEY, category TEXT NOT NULL, as_of_date TEXT NOT NULL, file_name TEXT NOT NULL, source_url TEXT NOT NULL, row_count INTEGER NOT NULL, updated_at TEXT NOT NULL)").run();
  await env.DB.prepare("CREATE TABLE IF NOT EXISTS idx_ownership_rows (stock TEXT NOT NULL, row_index INTEGER NOT NULL, payload TEXT NOT NULL, updated_at TEXT NOT NULL, PRIMARY KEY(stock,row_index))").run();
  await env.DB.prepare("CREATE TABLE IF NOT EXISTS broker_code_master (broker_name TEXT PRIMARY KEY, broker_code TEXT NOT NULL DEFAULT '', status_validation TEXT NOT NULL DEFAULT 'REVIEW', broker_status TEXT NOT NULL DEFAULT 'UNKNOWN', source_verification TEXT NOT NULL DEFAULT '', aliases TEXT NOT NULL DEFAULT '[]', updated_at TEXT NOT NULL)").run();

}
async function discoverLatestIdxOwnershipFile_(env) {
  let latest=null, discoveryError="";
  try {
    const pageResponse=await fetch(IDX_OWNERSHIP_PAGE,{headers:{"Accept":"text/html,application/xhtml+xml","User-Agent":"Mozilla/5.0 (compatible; BrokerSummary/1.0)"},redirect:"follow"});
    if(!pageResponse.ok)throw new Error("IDX page HTTP "+pageResponse.status);
    latest=latestIdxOwnershipFile_(await pageResponse.text());
    if(!latest)throw new Error("Tautan Excel kategori >5% tidak ditemukan pada halaman IDX.");
  }catch(e){discoveryError=String(e?.message||e);}
  if(!latest){
    try{
      const indexData=await zapi(env,API_BASE+"/finance:idx/ownership-files?category=lima-persen&length=200&start=0");
      let data=indexData;for(let i=0;i<3;i++){const nested=data&&(data.data||data.result);if(!nested||nested===data)break;data=nested;}
      const files=Array.isArray(data)?data:(Array.isArray(data.items)?data.items:(Array.isArray(data.data)?data.data:[]));
      const valid=files.filter(x=>/lima-persen\.xlsx/i.test(String(x.url||""))&&/^https:\/\/www\.idx\.co\.id\//i.test(String(x.url||"")));
      valid.sort((a,b)=>String(b.publishedAt||"").localeCompare(String(a.publishedAt||"")));
      const f=valid[0];if(f)latest={description:String(f.categoryLabel||"Pemegang Saham di Atas 5%"),url:String(f.url),publishedAt:String(parseIdxDate_(f.fileName||f.url)||f.publishedAt||"").slice(0,10),fileName:String(f.fileName||f.url.split("/").pop().split("?")[0])};
    }catch(e){discoveryError+="; fallback ZAPI: "+String(e?.message||e);}
  }
  if(!latest||!latest.publishedAt)throw new Error("Tidak menemukan Excel IDX terbaru >5%. "+discoveryError.slice(0,220));
  return latest;
}
async function refreshLatestIdxOwnership_(env, requestedStock) {
  const stock=cleanStock(requestedStock);
  if(!/^[A-Z0-9]{1,4}$/.test(stock))throw new Error("Ticker saham wajib valid; bulk download tidak diizinkan.");
  await ensureIdxOwnershipTables_(env);
  const latest=await discoverLatestIdxOwnershipFile_(env);
  const cachedMeta=await env.DB.prepare("SELECT stock,category,as_of_date AS asOfDate,file_name AS fileName,source_url AS sourceUrl,row_count AS rowCount,updated_at AS updatedAt FROM idx_ownership_stock_meta WHERE stock=?").bind(stock).first();
  const cachedRows=await env.DB.prepare("SELECT COUNT(*) AS n FROM idx_ownership_rows WHERE stock=?").bind(stock).first();
  if(cachedMeta && cachedMeta.asOfDate===latest.publishedAt && cachedMeta.fileName===latest.fileName && Number(cachedRows?.n||0)>0){
    return {ok:true,skipped:true,reason:"ticker_cache_current",stock,...latest,rowCount:Number(cachedRows.n),updatedAt:cachedMeta.updatedAt};
  }
  const fileResponse=await fetch(latest.url,{headers:{"Accept":"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/octet-stream","User-Agent":"Mozilla/5.0 (compatible; BrokerSummary/1.0)"},redirect:"follow"});
  if(!fileResponse.ok)throw new Error("IDX Excel HTTP "+fileResponse.status);
  const bytes=await fileResponse.arrayBuffer();
  if(bytes.byteLength<100||bytes.byteLength>40*1024*1024)throw new Error("Ukuran Excel IDX tidak valid: "+bytes.byteLength+" bytes");
  // Workbook dibaca di memori, tetapi yang ditulis ke D1 hanya ticker yang sedang dibuka.
  const rows=workbookRows_(bytes,stock).filter(r=>r.stock===stock);
  if(!rows.length)throw new Error("Tidak ada data kepemilikan >5% untuk ticker "+stock+" pada Excel terbaru.");
  const now=new Date().toISOString();
  await env.DB.prepare("DELETE FROM idx_ownership_rows WHERE stock=?").bind(stock).run();
  for(let i=0;i<rows.length;i+=80){
    const batch=rows.slice(i,i+80).map((row,j)=>env.DB.prepare("INSERT INTO idx_ownership_rows(stock,row_index,payload,updated_at) VALUES(?,?,?,?)").bind(stock,i+j,JSON.stringify(row),now));
    await env.DB.batch(batch);
  }
  await env.DB.prepare("INSERT INTO idx_ownership_stock_meta(stock,category,as_of_date,file_name,source_url,row_count,updated_at) VALUES(?,?,?,?,?,?,?) ON CONFLICT(stock) DO UPDATE SET category=excluded.category,as_of_date=excluded.as_of_date,file_name=excluded.file_name,source_url=excluded.source_url,row_count=excluded.row_count,updated_at=excluded.updated_at")
    .bind(stock,IDX_OWNERSHIP_CATEGORY,latest.publishedAt,latest.fileName,latest.url,rows.length,now).run();
  return {ok:true,skipped:false,stock,...latest,rowCount:rows.length,updatedAt:now};
}
async function handleIdxOwnershipRefresh(request, env) {
  if (!env.BACKFILL_TOKEN) return json({error:"Secret BACKFILL_TOKEN belum disetel."},503);
  if (request.headers.get("Authorization") !== "Bearer " + env.BACKFILL_TOKEN) return json({error:"Unauthorized"},401);
  if (request.method !== "POST") return json({error:"Gunakan POST dengan ?stock=SINI."},405);
  const stock=cleanStock(new URL(request.url).searchParams.get("stock"));
  try { return json(await refreshLatestIdxOwnership_(env,stock)); }
  catch(e) { return json({error:String(e?.message||e).slice(0,400),stock},502); }
}
async function handleIdxOwnership(url, env) {
  const stock=cleanStock(url.searchParams.get("stock"));
  if(!/^[A-Z0-9]{1,4}$/.test(stock))return json({error:"Kode saham tidak valid."},400);
  // GitHub Actions + pandas JSON menjadi sumber utama jika OWNERSHIP_JSON_URL disetel.
  if (env.OWNERSHIP_JSON_URL) {
    const response = await fetch(String(env.OWNERSHIP_JSON_URL), { headers: {"Accept":"application/json"}, signal: AbortSignal.timeout(20000) });
    const payload = await response.json().catch(() => ({}));
    if (!response.ok || !payload || !payload.stocks) return json({error:"Gagal membaca JSON hasil GitHub Actions", stock, rows:[]}, 502);
    const rows = Array.isArray(payload.stocks[stock]) ? payload.stocks[stock] : [];
    return json({ok:true, stock, rows:rows.map(r => ({...r, stock, source:"GITHUB_ACTIONS_PANDAS"})), source:"GITHUB_ACTIONS_PANDAS", generatedAt:payload.generatedAt || null}, 200);
  }
  // Jika PANDAS_API_URL dikonfigurasi, jadikan hasil Python/pandas sebagai sumber utama.
  // Jangan fallback diam-diam ke parser JS karena itu dapat mengembalikan CHANGE % berbeda.
  if (env.PANDAS_API_URL) {
    const apiBase = String(env.PANDAS_API_URL).replace(/\/$/, "");
    const apiResponse = await fetch(apiBase + "/api/ownership?stock=" + encodeURIComponent(stock), {
      headers: {"Accept":"application/json"},
      signal: AbortSignal.timeout(25000)
    });
    const apiData = await apiResponse.json().catch(() => ({}));
    if (!apiResponse.ok) return json({error:apiData.error || "Python pandas API gagal", stock, rows:[]}, 502);
    return json({...apiData, stock, source:"PYTHON_PANDAS_API"}, 200);
  }
  await ensureIdxOwnershipTables_(env);
  let meta=await env.DB.prepare("SELECT stock,category,as_of_date AS asOfDate,file_name AS fileName,source_url AS sourceUrl,row_count AS rowCount,updated_at AS updatedAt FROM idx_ownership_stock_meta WHERE stock=?").bind(stock).first();
  const cachedRows=await env.DB.prepare("SELECT COUNT(*) AS n FROM idx_ownership_rows WHERE stock=?").bind(stock).first();
  // Cache per ticker: data yang sama tidak mengunduh Excel berulang kali.
  const cacheFresh=meta && Number(cachedRows?.n||0)>0
    && Date.parse(meta.updatedAt||0)>=Date.parse(IDX_OWNERSHIP_PARSER_RELEASE)
    && (Date.now()-Date.parse(meta.updatedAt||0))<12*60*60*1000;
  if(!cacheFresh){
    try{await refreshLatestIdxOwnership_(env,stock);meta=await env.DB.prepare("SELECT stock,category,as_of_date AS asOfDate,file_name AS fileName,source_url AS sourceUrl,row_count AS rowCount,updated_at AS updatedAt FROM idx_ownership_stock_meta WHERE stock=?").bind(stock).first();}
    catch(e){if(!meta)return json({error:"Data Excel IDX >5% untuk "+stock+" belum tersimpan: "+String(e?.message||e).slice(0,250),stock,rows:[]},502);}
  }
  const result=await env.DB.prepare("SELECT payload FROM idx_ownership_rows WHERE stock=? ORDER BY row_index").bind(stock).all();
  let rows=(result.results||[]).map(r=>{try{return JSON.parse(r.payload)}catch{return null}}).filter(Boolean);
  rows=await resolveOwnershipBrokerCodes_(env,rows);
  return json({stock,meta,rows,source:"IDX_LATEST_XLSX_TICKER_CACHE"});
}
function shortenUnmappedBroker_(value){
  let name=String(value||"-").replace(/^PT\.?\s*/i,"").replace(/^BUT\s*/i,"").replace(/\s*,?\s*(TBK|PT|AG|N\.?A\.?)$/i,"").replace(/\s+/g," ").trim();
  const normalized=name.toUpperCase().replace(/[^A-Z0-9 ]/g,"").replace(/\s+/g," ").trim();
  if(normalized.includes("BANK CENTRAL ASIA"))return "BCA*";
  if(normalized.includes("DEUTSCHE BANK"))return "DEUTSCHE BANK*";
  if(normalized.includes("CITIBANK"))return "CITIBANK*";
  if(name.length>22){
    const words=name.split(" ").filter(Boolean);
    const compact=words.map(w=>w.length<=3?w:w[0].toUpperCase()).join("");
    return compact.length<=14?compact:name.slice(0,19).trim()+"…";
  }
  return name||"-";
}
function decodeHtml_(s){return String(s||"").replace(/&amp;/g,"&").replace(/&quot;/g,'"').replace(/&#39;|&apos;/g,"'").replace(/&lt;/g,"<").replace(/&gt;/g,">").replace(/&#x([0-9a-f]+);/gi,(_,h)=>String.fromCharCode(parseInt(h,16))).replace(/&#(\d+);/g,(_,n)=>String.fromCharCode(Number(n)));}
function normalizeBrokerName_(s){return String(s||"").normalize("NFKD").replace(/[\u0300-\u036f]/g,"").toUpperCase().replace(/[^A-Z0-9]+/g," ").trim().replace(/\s+/g," ");}
const COMMON_BROKER_WORDS_ = new Set(["PT","TBK","SEKURITAS","SECURITIES","SECURITAS","BANK","INDONESIA","INDONESIAN","INDO","THE","DAN","AND","PERSERO","LIMITED","LTD","CO","COMPANY","NATIONAL","ASIA","DIGITAL","CAPITAL","ASSET","INVESTAMA","INVESTMENT","VICKERS","INTERNATIONAL","GLOBAL","MAJU","BUT","BRANCH","CABANG","CUSTODY","REKENING","TAMPUNGAN","KSEI","CLOSED","MEMBER","SAHAM","EFEK","TAMP"]);
// Verified against the IDX Ringkasan Broker workbook supplied for 20-Oct-2026.
// These are exact broker-name aliases, not inferred from shareholder names.
const IDX_BROKER_NAME_ALIASES_ = new Map([
  ["MANDIRI SEKURITAS", "CC"], ["MANDIRI SEKURITAS PT", "CC"],
  ["MNC SEKURITAS", "EP"], ["PT MNC SEKURITAS", "EP"],
  ["BCA SEKURITAS", "SQ"], ["PT BCA SEKURITAS", "SQ"],
  ["KOREA INVESTMENT AND SEKURITAS INDONESIA", "BQ"], ["PT KOREA INVESTMENT AND SEKURITAS INDONESIA", "BQ"],
  ["PLUANG MAJU SEKURITAS", "RO"], ["PT PLUANG MAJU SEKURITAS", "RO"],
  ["OCBC SEKURITAS INDONESIA", "TP"], ["PT OCBC SEKURITAS INDONESIA", "TP"],
  ["ANUGERAH SEKURITAS INDONESIA", "ID"], ["PT ANUGERAH SEKURITAS INDONESIA", "ID"],
  ["BRI DANAREKSA SEKURITAS", "OD"], ["BRI DANAREKSA SEKURITAS PT", "OD"],
  ["SINARMAS SEKURITAS", "DH"], ["SINARMAS SEKURITAS PT", "DH"],
  ["PANIN SEKURITAS", "GR"], ["PANIN SEKURITAS TBK PT", "GR"],
  ["RHB SEKURITAS INDONESIA", "DR"], ["PT RHB SEKURITAS INDONESIA", "DR"],
  ["EKOKAPITAL SEKURITAS", "ES"], ["EKOKAPITAL SEKURITAS PT", "ES"]
]);
function brokerAliasCode_(name){
  const normalized=normalizeBrokerName_(name).replace(/\bPT\b/g," ").replace(/\bTBK\b/g," ").replace(/\s+/g," ").trim();
  const candidates=[normalizeBrokerName_(name),normalized];
  for(const key of candidates){const code=IDX_BROKER_NAME_ALIASES_.get(key);if(code)return code;}
  return "";
}
function distinctiveTokens_(name){return normalizeBrokerName_(name).split(" ").filter(w=>w.length>=3&&!COMMON_BROKER_WORDS_.has(w));}
function parseIdxBrokerDirectory_(html){
  const source=decodeIdxEscapes_(html);
  const found=new Map();
  // The official IDX members profile table is the canonical code/name directory.
  // Match rows with a two-character broker code followed by the brokerage name.
  const trRe=/<tr\b[^>]*>([\s\S]*?)<\/tr>/gi; let tr;
  while((tr=trRe.exec(source))){
    const cells=[...tr[1].matchAll(/<t[dh]\b[^>]*>([\s\S]*?)<\/t[dh]>/gi)].map(m=>decodeHtml_(m[1].replace(/<[^>]*>/g," ").replace(/\s+/g," ").trim()));
    if(cells.length<2)continue;
    let codeIndex=-1;
    for(let i=0;i<Math.min(cells.length,4);i++)if(/^[A-Z0-9]{2}$/.test(cells[i].trim().toUpperCase())){codeIndex=i;break;}
    if(codeIndex<0||!cells[codeIndex+1])continue;
    const code=cells[codeIndex].trim().toUpperCase(), name=cells[codeIndex+1].trim();
    if(!name||/^(KODE|CODE|NAMA|NAME|ANGGOTA|BROKERAGE)/i.test(name))continue;
    if(!found.has(code))found.set(code,{code,name});
  }
  // Some IDX builds serialize table records in Nuxt payload rather than <tr>.
  if(found.size<20){
    const re=/(?:"code"|"Code"|"IDFirm"|"idFirm")\s*:\s*"([A-Z0-9]{2})"[\s\S]{0,500}?(?:"name"|"Name"|"FirmName"|"brokerageFirm"|"Brokerage Firms")\s*:\s*"([^"\\]*(?:\\.[^"\\]*)*)"/g;let m;
    while((m=re.exec(source))){const name=decodeIdxEscapes_(m[2]).replace(/\\"/g,'"').trim();if(name&&!found.has(m[1]))found.set(m[1],{code:m[1],name});}
  }
  return [...found.values()];
}
async function ensureBrokerCodeMasterTable_(env){
  await env.DB.prepare("CREATE TABLE IF NOT EXISTS broker_code_master (broker_name TEXT PRIMARY KEY, broker_code TEXT NOT NULL DEFAULT '', status_validation TEXT NOT NULL DEFAULT 'REVIEW', broker_status TEXT NOT NULL DEFAULT 'UNKNOWN', source_verification TEXT NOT NULL DEFAULT '', aliases TEXT NOT NULL DEFAULT '[]', updated_at TEXT NOT NULL)").run();
}
function latestIdxBrokerWorkbookUrl_(html){
  const source=decodeIdxEscapes_(html);
  const urls=[];
  const re=/(https:\/\/www\.idx\.co\.id\/[^\s"'<>]*?\.xlsx(?:\?[^\s"'<>]*)?|\/Media\/[^\s"'<>]*?\.xlsx(?:\?[^\s"'<>]*)?)/gi;
  let m;
  while((m=re.exec(source))){
    let url=m[0];
    if(url.startsWith("//"))url="https:"+url;
    else if(url.startsWith("/"))url="https://www.idx.co.id"+url;
    url=url.replace(/\\u0026/gi,"&");
    if(/^https:\/\/www\.idx\.co\.id\//i.test(url)&&!urls.includes(url))urls.push(url);
  }
  // Prefer an XLSX explicitly labelled as the broker summary if several files are linked.
  return urls.find(u=>/broker|ringkasan/i.test(decodeURIComponent(u)))||urls[0]||"";
}
function parseIdxBrokerWorkbook_(bytes){
  const wb=XLSX.read(bytes,{type:"array",cellDates:false,dense:false});
  for(const sheetName of wb.SheetNames){
    const matrix=XLSX.utils.sheet_to_json(wb.Sheets[sheetName],{header:1,defval:null,raw:true});
    for(let hi=0;hi<Math.min(matrix.length,20);hi++){
      const headers=(matrix[hi]||[]).map(normalizeHeader_);
      const codeIdx=headers.findIndex(h=>/^(kode perusahaan|kode broker|kode efek|broker code|code)$/.test(h));
      const nameIdx=headers.findIndex(h=>/^(nama perusahaan|nama broker|nama anggota bursa|broker name|company name)$/.test(h));
      if(codeIdx<0||nameIdx<0)continue;
      const rows=[];
      for(let ri=hi+1;ri<matrix.length;ri++){
        const code=String(matrix[ri]?.[codeIdx]??"").trim().toUpperCase();
        const name=String(matrix[ri]?.[nameIdx]??"").trim();
        if(/^[A-Z0-9]{2}$/.test(code)&&name)rows.push({code,name});
      }
      if(rows.length>=20)return [...new Map(rows.map(r=>[r.code,r])).values()];
    }
  }
  return [];
}
async function refreshIdxBrokerMaster_(env){
  await ensureBrokerCodeMasterTable_(env);
  let response=null, sourceUrl="", errors=[];
  // First mirror the manual pandas workflow: discover and download the official IDX broker workbook.
  try{
    const page=await fetch(IDX_BROKER_SUMMARY_PAGE,{headers:{"Accept":"text/html,application/xhtml+xml","User-Agent":"Mozilla/5.0 (compatible; BrokerSummary/1.0)"},redirect:"follow"});
    if(!page.ok)throw new Error("IDX broker page HTTP "+page.status);
    const html=await page.text();
    const workbookUrl=latestIdxBrokerWorkbookUrl_(html);
    if(!workbookUrl)throw new Error("Tautan XLSX broker tidak ditemukan pada HTML halaman IDX.");
    const file=await fetch(workbookUrl,{headers:{"Accept":"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/octet-stream","User-Agent":"Mozilla/5.0 (compatible; BrokerSummary/1.0)"},redirect:"follow"});
    if(!file.ok)throw new Error("IDX broker workbook HTTP "+file.status);
    const bytes=await file.arrayBuffer();
    if(bytes.byteLength<100||bytes.byteLength>20*1024*1024)throw new Error("Ukuran XLSX broker tidak valid: "+bytes.byteLength);
    const parsed=parseIdxBrokerWorkbook_(bytes);
    if(parsed.length<20)throw new Error("Kolom Kode Perusahaan/Nama Perusahaan tidak ditemukan pada XLSX broker.");
    response=parsed;sourceUrl=workbookUrl;
  }catch(e){errors.push("Unduh workbook broker: "+String(e?.message||e));}
  for(const url of response?[]:[IDX_BROKER_SUMMARY_PAGE,IDX_BROKER_PROFILE_PAGE]){
    try{
      const r=await fetch(url,{headers:{"Accept":"text/html,application/xhtml+xml","User-Agent":"Mozilla/5.0 (compatible; BrokerSummary/1.0)"},redirect:"follow"});
      if(!r.ok)throw new Error("HTTP "+r.status);
      const html=await r.text(); const rows=parseIdxBrokerDirectory_(html);
      if(rows.length>=20){response=rows;sourceUrl=url;break;}
      errors.push(url+" tidak mengandung tabel kode/nama yang bisa dibaca ("+rows.length+" baris)");
    }catch(e){errors.push(url+": "+String(e?.message||e));}
  }
  if(!response)throw new Error("Direktori kode broker IDX gagal dibaca. "+errors.join("; ").slice(0,350));
  const now=new Date().toISOString();
  const statements=response.map(r=>env.DB.prepare("INSERT INTO broker_code_master(broker_name,broker_code,status_validation,broker_status,source_verification,aliases,updated_at) VALUES(?,?, 'VALID','UNKNOWN',?,'[]',?) ON CONFLICT(broker_name) DO UPDATE SET broker_code=excluded.broker_code,status_validation='VALID',source_verification=excluded.source_verification,updated_at=excluded.updated_at").bind(r.name,r.code,"IDX official broker directory: "+sourceUrl,now));
  for(let i=0;i<statements.length;i+=80)await env.DB.batch(statements.slice(i,i+80));
  await env.DB.prepare("INSERT INTO app_meta(key,value,updated_at) VALUES('idx_broker_master_sync',?,?) ON CONFLICT(key) DO UPDATE SET value=excluded.value,updated_at=excluded.updated_at").bind(JSON.stringify({count:response.length,sourceUrl,updatedAt:now}),now).run().catch(()=>{});
  return {ok:true,count:response.length,sourceUrl,updatedAt:now,brokers:response};
}
async function handleIdxBrokerMasterRefresh(request,env){
  if(!env.BACKFILL_TOKEN)return json({error:"Secret BACKFILL_TOKEN belum disetel."},503);
  if(request.headers.get("Authorization")!=="Bearer "+env.BACKFILL_TOKEN)return json({error:"Unauthorized"},401);
  if(request.method!=="POST")return json({error:"Gunakan POST."},405);
  try{return json(await refreshIdxBrokerMaster_(env));}catch(e){return json({error:String(e?.message||e).slice(0,500)},502);}
}
async function handleIdxBrokerMaster(url,env){
  await ensureBrokerCodeMasterTable_(env);
  const result=await env.DB.prepare("SELECT broker_code AS code,broker_name AS name,status_validation AS validation,source_verification AS source,updated_at AS updatedAt FROM broker_code_master WHERE broker_code<>'' ORDER BY broker_code").all();
  return json({count:(result.results||[]).length,brokers:result.results||[]});
}
async function resolveOwnershipBrokerCodes_(env, rows){
  await ensureBrokerCodeMasterTable_(env);
  let master=(await env.DB.prepare("SELECT broker_name,broker_code,status_validation,broker_status,aliases FROM broker_code_master").all()).results||[];
  // Bootstrap/update from the official IDX member directory when not yet present.
  if(master.filter(m=>m.status_validation==="VALID"&&m.broker_code).length<20){
    try{await refreshIdxBrokerMaster_(env);master=(await env.DB.prepare("SELECT broker_name,broker_code,status_validation,broker_status,aliases FROM broker_code_master").all()).results||[];}catch(e){console.warn("IDX broker directory refresh:",String(e?.message||e));}
  }
  const normalize=normalizeBrokerName_;
  const entries=master.map(m=>({...m,keys:[m.broker_name,...(()=>{try{return JSON.parse(m.aliases||"[]")}catch{return[]}})()].map(normalize)}));
  const tokenIndex=new Map();
  for(const m of entries){if(m.status_validation!=="VALID"||!m.broker_code)continue;for(const token of new Set(distinctiveTokens_(m.broker_name))){if(!tokenIndex.has(token))tokenIndex.set(token,new Set());tokenIndex.get(token).add(m.broker_code);}}
  const output=[];
  for(const row of rows){
    const sourceNames=Array.isArray(row.brokerNames)&&row.brokerNames.length?row.brokerNames:[row.brokerName].filter(Boolean);
    const codes=[], displayCodes=[];
    for(const brokerName of sourceNames){
      const key=normalize(brokerName);
      const aliasCode=brokerAliasCode_(brokerName);
      let match=aliasCode ? entries.find(m=>m.status_validation==="VALID"&&m.broker_code===aliasCode) : null;
      if(!match) match=entries.find(m=>m.status_validation==="VALID"&&m.broker_code&&m.keys.includes(key));
      if(!match){
        const candidateCodes=new Set();
        for(const token of distinctiveTokens_(key)){const found=tokenIndex.get(token);if(found&&found.size===1)candidateCodes.add([...found][0]);}
        if(candidateCodes.size===1){const code=[...candidateCodes][0];match=entries.find(m=>m.broker_code===code&&m.status_validation==="VALID");}
      }
      const resolvedCode=match?.broker_code||aliasCode||"";
      let label=resolvedCode;
      if(!label){
        label=shortenUnmappedBroker_(brokerName).replace(/\s+BANK\*$/i,"").replace(/\*$/g,"").trim();
        // Match the user's pandas output convention for unresolved institutions such as Deutsche.
        if(/DEUTSCHE/i.test(brokerName))label="DEUTSCHE";
      }
      if(label && !displayCodes.includes(label))displayCodes.push(label);
      if(resolvedCode && /^[A-Z0-9]{2}$/.test(resolvedCode) && !codes.includes(resolvedCode))codes.push(resolvedCode);
      if(!match&&!aliasCode){const now=new Date().toISOString();await env.DB.prepare("INSERT INTO broker_code_master(broker_name,broker_code,status_validation,broker_status,source_verification,aliases,updated_at) VALUES(?, '', 'REVIEW','UNKNOWN','Belum terverifikasi dari direktori IDX','[]',?) ON CONFLICT(broker_name) DO NOTHING").bind(String(brokerName||"Tidak diketahui"),now).run();}
    }
    output.push({...row,brokerNames:sourceNames,brokerCodes:displayCodes.sort(),brokerCode:codes[0]||"",brokerDisplay:displayCodes.join(", ")||shortenUnmappedBroker_(sourceNames[0]||""),brokerIdentity:codes[0]||normalize(sourceNames[0]||""),brokerValidation:codes.length?"VALID":"REVIEW",brokerStatus:row.status||"UNKNOWN"});
  }
  return output;
}

async function handleRetailOwnership(url,env){
  const stock=cleanStock(url.searchParams.get("stock"));
  if(!/^[A-Z0-9]{1,4}$/.test(stock)) return json({error:"Kode saham tidak valid."},400);
  // Pastikan tabel ada agar deployment baru tetap dapat menyimpan hasil on-demand.
  await env.DB.prepare("CREATE TABLE IF NOT EXISTS retail_ownership (stock TEXT NOT NULL, month TEXT NOT NULL, payload TEXT NOT NULL, updated_at TEXT NOT NULL, PRIMARY KEY(stock, month))").run();
  const cached=await env.DB.prepare("SELECT payload FROM retail_ownership WHERE stock=? ORDER BY month DESC LIMIT 1").bind(stock).first();
  if(cached){try{return json(JSON.parse(cached.payload))}catch{}}
  try{
    const raw=await zapi(env,API_KSEI+"?code="+encodeURIComponent(stock)+"&type=EQUITY&length=1000");
    let data=raw;
    for(let i=0;i<3;i++){const nested=data&&(data.data||data.result);if(!nested||nested===data)break;data=nested;}
    let items=Array.isArray(data.items)?data.items:(Array.isArray(data.data)?data.data:[]);
    const item=items.find(x=>cleanStock(x.code)===stock) || items[0];
    if(!item) return json({error:"Data kepemilikan retail belum tersedia dari endpoint KSEI.",code:stock},404);
    const local=item.local||{};
    const individual=num(local.individual);
    const outstanding=num(item.outstanding);
    const recordedTotal=num(item.recordedTotal);
    // Porsi kepemilikan individu lokal dibanding total saham beredar.
    const denominator=outstanding || recordedTotal;
    const percent=denominator>0 ? individual/denominator*100 : 0;
    const date=String(item.date||data.date||"").slice(0,10);
    const month=date.slice(0,7) || isoTodayWIB().slice(0,7);
    const payload={code:stock,percent,month,date,individualShares:individual,outstandingShares:outstanding,source:"KSEI",updatedAt:new Date().toISOString()};
    await env.DB.prepare("INSERT INTO retail_ownership(stock,month,payload,updated_at) VALUES(?,?,?,?) ON CONFLICT(stock,month) DO UPDATE SET payload=excluded.payload,updated_at=excluded.updated_at").bind(stock,month,JSON.stringify(payload),payload.updatedAt).run();
    return json(payload);
  }catch(e){
    const message=String(e.message||e).slice(0,300);
    return json({error:message,code:stock},/429|quota exceeded|monthly quota/i.test(message)?429:502);
  }
}

async function refreshStocks(env){
  // Securities endpoint is paginated; continue until a short/empty page is returned.
  const updatedAt=new Date().toISOString();
  let offset=0, imported=0;
  const pageSize=1000;
  for(let page=0;page<100;page++){
    const sec=await zapi(env,API_SECURITIES+"?length="+pageSize+"&start="+offset);
    const items=Array.isArray(sec)?sec:(Array.isArray(sec.data)?sec.data:(Array.isArray(sec.items)?sec.items:[]));
    if(!items.length)break;
    const batch=[];
    for(const stock of items){
      const code=cleanStock(stock.Code||stock.code||stock.symbol);
      if(!/^[A-Z0-9]{1,4}$/.test(code))continue;
      batch.push(env.DB.prepare("INSERT INTO stocks(code,name,board,updated_at) VALUES(?,?,?,?) ON CONFLICT(code) DO UPDATE SET name=CASE WHEN excluded.name<>'' THEN excluded.name ELSE stocks.name END,board=CASE WHEN excluded.board<>'' THEN excluded.board ELSE stocks.board END,updated_at=excluded.updated_at")
        .bind(code,String(stock.Name||stock.name||stock.companyName||""),String(stock.ListingBoard||stock.listingBoard||stock.board||""),updatedAt));
      const logo=String(stock.Logo||stock.logo||stock.logoUrl||stock.iconUrl||"");
      if(logo){try{await env.DB.prepare("UPDATE stocks SET logo=? WHERE code=?").bind(logo,code).run();}catch{}}
    }
    for(let i=0;i<batch.length;i+=80)await env.DB.batch(batch.slice(i,i+80));
    imported+=batch.length;
    if(items.length<pageSize)break;
    offset+=items.length;
  }
  if(!imported){
    const existing=await env.DB.prepare("SELECT COUNT(*) AS n FROM stocks").first();
    if(!Number(existing?.n||0))throw new Error("Endpoint securities tidak mengembalikan daftar saham; daftar saham belum diisi.");
  }
  return imported;
}
async function saveBrokerDaily(env,code,date,data,useNet=false){
  const record=normalizeBrokerSummary(data,code,date,useNet);
  if(!record.buyers.length && !record.sellers.length) {
    // Jangan simpan respons kosong sebagai data valid agar request berikutnya bisa mencoba ulang.
    throw new Error("Respons ZAPI tidak berisi baris BUY/SELL untuk "+code+" tanggal "+date+
      " (count="+record.count+"). Record kosong tidak disimpan.");
  }
  const serialized=JSON.stringify(record),now=new Date().toISOString();
  // payload = gross/net=false; net_payload = net=true. Upsert hanya memperbarui mode yang diminta.
  const statement="INSERT INTO broker_daily(stock,trade_date,payload,net_payload,updated_at) VALUES(?,?,?,?,?) "+
    "ON CONFLICT(stock,trade_date) DO UPDATE SET "+(useNet?"net_payload=excluded.net_payload":"payload=excluded.payload")+",updated_at=excluded.updated_at";
  await env.DB.prepare(statement).bind(code,date,useNet?"":serialized,useNet?serialized:null,now).run();
}
async function refreshBrokerMaster(env){
  const master={};
  for(const type of ["LOCAL","FOREIGN"]){
    const d=await zapi(env,API_BROKERS+"?type="+type);
    const brokers=Array.isArray(d.items)?d.items:(Array.isArray(d.data)?d.data:[]);
    for(const b of brokers){const code=String(b.code||b.Code||"").toUpperCase();if(code)master[code]=String(b.type||type).toUpperCase();}
  }
  const batch=Object.entries(master).map(([code,type])=>env.DB.prepare("INSERT INTO broker_master(code,type) VALUES(?,?) ON CONFLICT(code) DO UPDATE SET type=excluded.type").bind(code,type));
  for(let i=0;i<batch.length;i+=80)await env.DB.batch(batch.slice(i,i+80));
  return Object.keys(master).length;
}
async function getRetailBrokers_(env){
  const raw=await metaGet(env,"retail_brokers_config");
  if(raw){try{const parsed=JSON.parse(raw);if(Array.isArray(parsed)){const clean=[...new Set(parsed.map(x=>String(x||"").trim().toUpperCase()).filter(x=>/^[A-Z0-9]{2}$/.test(x)))];if(clean.length)return clean;}}catch{}}
  return RETAIL_BROKERS.slice();
}
async function handleRetailBrokers(request,env){
  if(request.method==="POST"){
    let body;try{body=await request.json();}catch{return json({error:"Format JSON tidak valid."},400);}
    if(!Array.isArray(body.brokers))return json({error:"Daftar brokers harus berupa array."},400);
    const brokers=[...new Set(body.brokers.map(x=>String(x||"").trim().toUpperCase()).filter(x=>/^[A-Z0-9]{2}$/.test(x)))];
    if(!brokers.length)return json({error:"Minimal satu broker ritel harus terdaftar."},400);
    if(brokers.length>100)return json({error:"Maksimal 100 broker."},400);
    await metaSet(env,"retail_brokers_config",JSON.stringify(brokers));
    // Cache historis harus dihitung ulang dengan klasifikasi broker terbaru.
    try{await env.DB.prepare("DELETE FROM flow_summary_cache").run();}catch{}
    return json({ok:true,brokers});
  }
  return json({brokers:await getRetailBrokers_(env)});
}

async function metaGet(env,key){const r=await env.DB.prepare("SELECT value FROM app_meta WHERE key=?").bind(key).first();return r?.value??null;}
async function metaSet(env,key,value){await env.DB.prepare("INSERT INTO app_meta(key,value) VALUES(?,?) ON CONFLICT(key) DO UPDATE SET value=excluded.value").bind(key,String(value)).run();}
async function handleBackfill(request, url, env) {
  if (!env.BACKFILL_TOKEN) {
    return json({ error: "Secret BACKFILL_TOKEN belum disetel." }, 503);
  }

  const auth = request.headers.get("Authorization") || "";
  if (auth !== "Bearer " + env.BACKFILL_TOKEN) {
    return json({ error: "Unauthorized" }, 401);
  }

  if (request.method !== "POST") {
    return json({ error: "Gunakan POST." }, 405);
  }

  let body = {};
  try {
    body = await request.json();
  } catch {}

  const action = String(
    body.action || url.searchParams.get("action") || "status"
  ).toLowerCase();

  // STATUS: checkpoint dan ringkasan kegagalan persisten di D1.
  if (action === "status") {
    const raw = await metaGet(env, "backfill_state");
    const state = raw ? JSON.parse(raw) : null;
    const count = await env.DB.prepare(
      "SELECT COUNT(*) AS n FROM backfill_failures"
    ).first();
    const rows = await env.DB.prepare(
      "SELECT stock, trade_date AS date, attempts, last_error AS error " +
      "FROM backfill_failures ORDER BY trade_date, stock LIMIT 20"
    ).all();

    return json({
      ok: true,
      state,
      pendingFailures: Number(count?.n || 0),
      failures: rows.results || []
    });
  }

  // START tidak boleh menghapus checkpoint aktif atau kegagalan yang belum selesai.
  if (action === "start") {
    const oldRaw = await metaGet(env, "backfill_state");
    const pending = Number(
      (await env.DB.prepare(
        "SELECT COUNT(*) AS n FROM backfill_failures"
      ).first())?.n || 0
    );

    if (oldRaw) {
      const old = JSON.parse(oldRaw);
      if (old.status === "running" || old.status === "retry_pending" || pending > 0) {
        return json({
          error: "Checkpoint lama masih aktif atau masih ada kegagalan. Tidak di-reset.",
          state: old,
          pendingFailures: pending
        }, 409);
      }
    }

    const start = parseDate(body.startDate || "2025-01-01");
    const end = parseDate(body.endDate || isoTodayWIB());
    if (!start || !end || start > end) {
      return json({ error: "Rentang tanggal tidak valid." }, 400);
    }

    await refreshStocks(env);
    const total = Number(
      (await env.DB.prepare("SELECT COUNT(*) AS n FROM stocks").first())?.n || 0
    );
    if (!total) {
      return json({ error: "Daftar ticker kosong; backfill tidak dimulai." }, 409);
    }

    const state = {
      status: "running",
      startDate: start,
      endDate: end,
      currentDate: start,
      stockIndex: 0,
      totalStocks: total,
      completedDates: 0,
      success: 0,
      failures: 0,
      startedAt: new Date().toISOString(),
      lastError: ""
    };

    await metaSet(env, "backfill_state", JSON.stringify(state));
    await metaSet(env, "backfill_failures", "[]");
    await metaSet(env, "backfill_date_failures", "{}");

    return json({
      ok: true,
      message: "Checkpoint disiapkan; data broker belum diambil.",
      state
    });
  }

  const raw = await metaGet(env, "backfill_state");
  if (!raw) {
    return json({ error: "Backfill belum dimulai. Jalankan action=start terlebih dahulu." }, 409);
  }
  const state = JSON.parse(raw);

  // RETRY: setiap kegagalan tersimpan per ticker-tanggal di tabel D1.
  if (action === "retry_failed") {
    const batchSize = Math.max(1, Math.min(25, Number(body.batchSize) || 10));
    const result = await env.DB.prepare(
      "SELECT stock, trade_date AS date FROM backfill_failures " +
      "ORDER BY trade_date, stock LIMIT ?"
    ).bind(batchSize).all();
    const todo = result.results || [];
    const errors = [];
    let success = 0;
    let cursor = 0;

    async function retryRunner() {
      while (cursor < todo.length) {
        const item = todo[cursor++];
        try {
          const data = await zapi(
            env,
            API_BROKER + "?code=" + encodeURIComponent(item.stock) +
              "&startDate=" + item.date + "&endDate=" + item.date + "&net=false"
          );
          await saveBrokerDaily(env, item.stock, item.date, data);
          await env.DB.prepare(
            "DELETE FROM backfill_failures WHERE stock=? AND trade_date=?"
          ).bind(item.stock, item.date).run();
          success++;
        } catch (e) {
          const message = String(e.message || e).slice(0, 500);
          await env.DB.prepare(
            "UPDATE backfill_failures SET attempts=attempts+1, last_error=?, updated_at=? " +
            "WHERE stock=? AND trade_date=?"
          ).bind(message, new Date().toISOString(), item.stock, item.date).run();
          errors.push({ stock: item.stock, date: item.date, error: message });
        }
      }
    }

    await Promise.all(
      Array.from({ length: Math.min(5, todo.length) }, retryRunner)
    );

    const pending = Number(
      (await env.DB.prepare("SELECT COUNT(*) AS n FROM backfill_failures").first())?.n || 0
    );
    state.retrySuccess = (state.retrySuccess || 0) + success;
    state.pendingFailures = pending;
    state.lastRetryAt = new Date().toISOString();

    if (state.currentDate > state.endDate) {
      state.status = pending === 0 ? "completed" : "retry_pending";
      if (pending === 0) state.completedAt = new Date().toISOString();
    }

    await metaSet(env, "backfill_state", JSON.stringify(state));
    return json({
      ok: true,
      attempted: todo.length,
      success,
      failedThisBatch: errors.length,
      pendingFailures: pending,
      state,
      errors
    });
  }

  if (action !== "step") {
    return json({ error: "Action harus status, start, step, atau retry_failed." }, 400);
  }

  if (state.status === "completed") {
    return json({ ok: true, state, message: "Backfill sudah selesai." });
  }
  if (state.status === "retry_pending") {
    return json({
      ok: false,
      state,
      message: "Semua tanggal sudah diproses. Jalankan action=retry_failed."
    }, 409);
  }

  const date = state.currentDate;
  if (date > state.endDate) {
    const pending = Number(
      (await env.DB.prepare("SELECT COUNT(*) AS n FROM backfill_failures").first())?.n || 0
    );
    state.pendingFailures = pending;
    state.status = pending === 0 ? "completed" : "retry_pending";
    if (pending === 0) state.completedAt = new Date().toISOString();
    await metaSet(env, "backfill_state", JSON.stringify(state));
    return json({ ok: true, state });
  }

  if (!weekdaysOnly(date)) {
    state.currentDate = addDays(date, 1);
    state.stockIndex = 0;

    if (state.currentDate > state.endDate) {
      const pending = Number(
        (await env.DB.prepare("SELECT COUNT(*) AS n FROM backfill_failures").first())?.n || 0
      );
      state.pendingFailures = pending;
      state.status = pending === 0 ? "completed" : "retry_pending";
      if (pending === 0) state.completedAt = new Date().toISOString();
    }

    await metaSet(env, "backfill_state", JSON.stringify(state));
    return json({ ok: true, skippedWeekend: true, state });
  }

  const stocks = (
    await env.DB.prepare("SELECT code FROM stocks ORDER BY code").all()
  ).results || [];
  if (!stocks.length) return json({ error: "Tabel stocks kosong." }, 409);

  const batchSize = Math.max(1, Math.min(25, Number(body.batchSize) || 10));
  const from = Math.min(state.stockIndex || 0, stocks.length);
  const to = Math.min(stocks.length, from + batchSize);
  const slice = stocks.slice(from, to);
  let success = 0;
  let cursor = 0;
  const errors = [];

  async function runner() {
    while (cursor < slice.length) {
      const item = slice[cursor++];
      try {
        const data = await zapi(
          env,
          API_BROKER + "?code=" + encodeURIComponent(item.code) +
            "&startDate=" + date + "&endDate=" + date + "&net=false"
        );
        await saveBrokerDaily(env, item.code, date, data);
        await env.DB.prepare(
          "DELETE FROM backfill_failures WHERE stock=? AND trade_date=?"
        ).bind(item.code, date).run();
        success++;
      } catch (e) {
        const message = String(e.message || e).slice(0, 500);
        const now = new Date().toISOString();
        await env.DB.prepare(
          "INSERT INTO backfill_failures " +
          "(stock, trade_date, attempts, last_error, updated_at) " +
          "VALUES (?, ?, 1, ?, ?) " +
          "ON CONFLICT(stock, trade_date) DO UPDATE SET " +
          "attempts=backfill_failures.attempts+1, " +
          "last_error=excluded.last_error, updated_at=excluded.updated_at"
        ).bind(item.code, date, message, now).run();
        errors.push({ stock: item.code, date, error: message });
        console.error("Backfill " + date + " " + item.code + ": " + message);
      }
    }
  }

  await Promise.all(
    Array.from({ length: Math.min(5, slice.length) }, runner)
  );

  state.stockIndex = to;
  state.success = (state.success || 0) + success;
  state.failures = (state.failures || 0) + errors.length;
  state.lastBatch = {
    date,
    from,
    to,
    totalStocks: stocks.length,
    success,
    failures: errors.length,
    completedAt: new Date().toISOString()
  };

  if (state.stockIndex >= stocks.length) {
    state.completedDates = (state.completedDates || 0) + 1;
    state.currentDate = addDays(date, 1);
    state.stockIndex = 0;
  }

  if (state.currentDate > state.endDate) {
    const pending = Number(
      (await env.DB.prepare("SELECT COUNT(*) AS n FROM backfill_failures").first())?.n || 0
    );
    state.pendingFailures = pending;
    state.status = pending === 0 ? "completed" : "retry_pending";
    if (pending === 0) state.completedAt = new Date().toISOString();
  }

  await metaSet(env, "backfill_state", JSON.stringify(state));
  return json({
    ok: true,
    state,
    errors,
    remainingStocksForDate: state.currentDate === date ? stocks.length - state.stockIndex : 0
  });
}
function addDays(date,n){const d=new Date(date+"T00:00:00Z");d.setUTCDate(d.getUTCDate()+n);return d.toISOString().slice(0,10);}
async function runDailyIngestion(env){
  const today=isoTodayWIB();
  if(!weekdaysOnly(today))return;
  const exists=await metaGet(env,"ingest_date");
  if(exists===today)return;
  let success=0,failures=0,total=0,error="";
  try{
    await refreshStocks(env);
    const stocks=(await env.DB.prepare("SELECT code FROM stocks ORDER BY code").all()).results||[];
    total=stocks.length;
    let cursor=0;
    async function runner(){while(cursor<stocks.length){const stock=stocks[cursor++];try{
      const data=await zapi(env,API_BROKER+"?code="+encodeURIComponent(stock.code)+"&startDate="+today+"&endDate="+today+"&net=false");
      await saveBrokerDaily(env,stock.code,today,data);success++;
    }catch(e){failures++;console.error("Ingest "+stock.code+": "+e.message);}}}
    await Promise.all(Array.from({length:5},runner));
    try{await refreshBrokerMaster(env);}catch(e){console.error("Broker master: "+e.message);}
    // Never mark the day complete if even one ticker failed.
    if(failures===0&&success===total){await metaSet(env,"ingest_date",today);}else{error="Sebagian ticker gagal; tanggal tidak ditandai selesai dan cron berikutnya dapat mencoba ulang.";}
  }catch(e){failures++;error=String(e.message||e);console.error("Daily ingestion: "+error);}
  await metaSet(env,"ingest_status",JSON.stringify({date:today,success,failures,total,error,completedAt:new Date().toISOString()}));
}
