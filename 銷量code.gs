/**
 * 銷售數據中心 V7.1 - Multi-Select Chips & Smart Alerts (Safety Fix)
 * 核心升級：子通路複選陣列處理、高流速斷貨演算法、嚴格清洗
 */

const CONFIG = {
  // 1. 嚴格資料清洗規則 (新增排除維修收入、運費收入)
  BLACKLIST_KW: ["贈品", "報廢", "轉借出單", "轉廣告費", "維修派工單-消費者", "維修收入", "運費收入"],
  IDX: {
    DATE: 0,
    CHANNEL: 1,
    CAT: 2,
    CODE: 3,
    NAME: 4,
    QTY: 5,
    AMOUNT: 6
  },
  
  // 2. 通路分組定義
  GROUPS: {
    "電商群組": ["P購", "蝦皮", "環球購物", "Friday", "momo", "mo寄倉", "好好運動", "全國電子", "BH官網", "嘖嘖", "BLADEZ官網", "誠品", "citiesocial", "東森", "Y購", "PC寄倉", "全電商", "家樂福線上", "愛料理"],
    "門市群組": ["沙鹿店", "高雄店", "花蓮店", "台南店", "家福經國店", "竹北店", "員林店", "北屯店", "巨城店", "環球青埔店", "頭份店"],
    "經銷群組": ["經銷商", "達康"],
    "其他群組": ["商用", "電話", "外銷"]
  }
};

// 建立全域通路查找表
const CHANNEL_MAP = {};
(function initChannelMap() {
  for (const [group, channels] of Object.entries(CONFIG.GROUPS)) {
    for (const ch of channels) {
      CHANNEL_MAP[ch] = group;
    }
  }
})();

function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('Sales Data Center')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getChannelConfig() {
  return CONFIG.GROUPS;
}

function getChannelGroup(channelName) {
  const ch = String(channelName).trim();
  return CHANNEL_MAP[ch] || "其他群組";
}

function parseDate(val) {
  if (!val) return null;
  if (Object.prototype.toString.call(val) === '[object Date]') return val;
  const d = new Date(val);
  return isNaN(d.getTime()) ? null : d;
}

/**
 * 核心資料查詢
 * @param {string} startDateStr 
 * @param {string} endDateStr 
 * @param {string[]} selectedGroups - 群組複選陣列
 * @param {string[]} selectedSubChannels - 子通路複選陣列 (若包含 'All' 則全選)
 */
function getData(startDateStr, endDateStr, selectedGroups, selectedSubChannels) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const data = sheet.getDataRange().getValues();

    // Safety Checks for Parameters
    if (!selectedGroups) selectedGroups = ['All'];
    if (!selectedSubChannels) selectedSubChannels = ['All'];

    // 1. 定義時間區間
    const start = new Date(startDateStr);
    const end = new Date(endDateStr);
    start.setHours(0,0,0,0);
    end.setHours(23,59,59,999);

    const tStart = start.getTime();
    const tEnd = end.getTime();
    const msDiff = tEnd - tStart;
    const dayDiff = Math.ceil(msDiff / (1000 * 60 * 60 * 24)) || 1;

    const tPrevEnd = tStart - 1;
    const tPrevStart = tPrevEnd - msDiff;
    
    const tRecent7 = tEnd - (7 * 24 * 60 * 60 * 1000);
    const tLast3 = tEnd - (3 * 24 * 60 * 60 * 1000);

    // 2. 資料容器
    let curStats = { rev: 0, qty: 0 };
    let prevStats = { rev: 0, qty: 0 };
    
    // Group Hierarchy: { Group: { rev, qty, subs: { Sub: { rev, qty, prods: {} } } } }
    const groupHierarchy = {
      "電商群組": { rev: 0, qty: 0, subs: {} },
      "門市群組": { rev: 0, qty: 0, subs: {} },
      "經銷群組": { rev: 0, qty: 0, subs: {} },
      "其他群組": { rev: 0, qty: 0, subs: {} }
    };

    const trendAgg = {}; 
    const catRevMap = {}; 
    const prodMap = new Map();

    // 3. 遍歷資料
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      
      // A. 日期篩選
      const dateVal = row[CONFIG.IDX.DATE];
      if (!dateVal) continue;
      
      let tDate;
      if (typeof dateVal.getTime === 'function') {
        tDate = dateVal.getTime();
      } else {
        const pd = parseDate(dateVal);
        if (!pd) continue;
        tDate = pd.getTime();
      }

      if (tDate < tPrevStart) continue;

      const amt = Number(row[CONFIG.IDX.AMOUNT]);
      if (amt <= 0 || isNaN(amt)) continue;

      // B. 排除與篩選
      const chName = String(row[CONFIG.IDX.CHANNEL]).trim();
      const chGroup = getChannelGroup(chName);
      
      // 子通路複選邏輯: 如果 selectedSubChannels 不包含 'All' 且 不包含該通路 -> 跳過
      if (selectedSubChannels.indexOf('All') === -1 && selectedSubChannels.indexOf(chName) === -1) continue;

      // 群組複選邏輯
      if (selectedGroups.indexOf('All') === -1 && selectedGroups.indexOf(chGroup) === -1) continue;

      // 關鍵字排除
      const prodName = String(row[CONFIG.IDX.NAME]);
      let isBlacklisted = false;
      for (let k = 0; k < CONFIG.BLACKLIST_KW.length; k++) {
        if (prodName.includes(CONFIG.BLACKLIST_KW[k])) {
          isBlacklisted = true;
          break;
        }
      }
      if (isBlacklisted) continue;

      // C. 聚合
      const qty = Number(row[CONFIG.IDX.QTY]) || 0;
      const pCode = String(row[CONFIG.IDX.CODE]).trim();
      
      // --- Current Period ---
      if (tDate >= tStart && tDate <= tEnd) {
        curStats.rev += amt;
        curStats.qty += qty;
        
        // Group Hierarchy Aggregation
        if (groupHierarchy[chGroup]) {
          const gNode = groupHierarchy[chGroup];
          gNode.rev += amt;
          gNode.qty += qty;
          
          if (!gNode.subs[chName]) gNode.subs[chName] = { rev: 0, qty: 0, prods: {} };
          const sNode = gNode.subs[chName];
          sNode.rev += amt;
          sNode.qty += qty;
          
          if (!sNode.prods[pCode]) sNode.prods[pCode] = { qty: 0, rev: 0 };
          sNode.prods[pCode].qty += qty;
          sNode.prods[pCode].rev += amt;
        }

        // Trend
        const pCat = String(row[CONFIG.IDX.CAT] || "其他").trim();
        const dObj = new Date(tDate);
        const dateStr = `${dObj.getFullYear()}-${String(dObj.getMonth()+1).padStart(2,'0')}-${String(dObj.getDate()).padStart(2,'0')}`;
        
        if (!trendAgg[dateStr]) trendAgg[dateStr] = {};
        if (!trendAgg[dateStr][pCat]) trendAgg[dateStr][pCat] = 0;
        trendAgg[dateStr][pCat] += amt;

        if (!catRevMap[pCat]) catRevMap[pCat] = 0;
        catRevMap[pCat] += amt;

        // Product Map (Global)
        let p = prodMap.get(pCode);
        if (!p) {
          p = { code: pCode, name: prodName, curRev: 0, curQty: 0, prevRev: 0, prevQty: 0, last3Qty: 0, recent7Rev: 0 };
          prodMap.set(pCode, p);
        }
        p.curRev += amt;
        p.curQty += qty;
        
        if (tDate >= tLast3) p.last3Qty += qty;
        if (tDate >= tRecent7) p.recent7Rev += amt;
      }
      
      // --- Previous Period ---
      else if (tDate >= tPrevStart && tDate <= tPrevEnd) {
        prevStats.rev += amt;
        prevStats.qty += qty;
        
        let p = prodMap.get(pCode);
        if (!p) {
          p = { code: pCode, name: prodName, curRev: 0, curQty: 0, prevRev: 0, prevQty: 0, last3Qty: 0, recent7Rev: 0 };
          prodMap.set(pCode, p);
        }
        p.prevRev += amt;
        p.prevQty += qty;
      }
    }

    // 4. Trend Processing
    const sortedCats = Object.entries(catRevMap).sort((a, b) => b[1] - a[1]);
    const top5Cats = new Set(sortedCats.slice(0, 5).map(e => e[0]));
    
    const trendLabels = Object.keys(trendAgg).sort();
    const displayCats = [...top5Cats, `其他 (含${sortedCats.length > 5 ? sortedCats[5][0] : ''}...)`];
    const datasetMap = {};
    displayCats.forEach(c => datasetMap[c] = new Array(trendLabels.length).fill(0));

    trendLabels.forEach((date, idx) => {
      const dayData = trendAgg[date];
      for (const [cat, val] of Object.entries(dayData)) {
        if (top5Cats.has(cat)) datasetMap[cat][idx] += val;
        else datasetMap[displayCats[displayCats.length - 1]][idx] += val;
      }
    });

    const validDisplayCats = displayCats.filter(cat => datasetMap[cat].some(v => v > 0));
    const trendDatasets = validDisplayCats.map(cat => ({ label: cat, data: datasetMap[cat] }));

    // 5. Post Processing: Alerts (High Velocity Stockout) & Matrix
    const prodList = Array.from(prodMap.values());
    const matrixData = [];
    const paretoDataRaw = [];
    const alerts = { churn: [], stockout: [], profit: [], potential: [] };
    
    const avgRevPerDay = dayDiff > 0 ? (curStats.rev / dayDiff) : 0;
    const potentialThreshold = (avgRevPerDay * 7) * 1.5;

    // A. 找出所有銷量大於 0 的產品，用於排序
    const activeProducts = prodList.filter(p => p.curQty > 0);
    
    // B. 計算「高流速斷貨」: 先取 Top 20 銷量，再檢查 last3Qty == 0
    // Sort by Total Qty Descending
    activeProducts.sort((a, b) => b.curQty - a.curQty);
    const top20Velocity = activeProducts.slice(0, 20);
    
    top20Velocity.forEach(p => {
      // 如果總銷量高，但最近 3 天完全沒賣出，視為異常
      if (p.last3Qty === 0) {
        alerts.stockout.push({ code: p.code, name: p.name, val: p.curQty });
      }
    });

    for (const p of prodList) {
      if (p.curRev > 0) paretoDataRaw.push(p);

      // Matrix
      if (p.curRev > 0) {
        let growth = 0;
        if (p.prevRev > 0) growth = ((p.curRev - p.prevRev) / p.prevRev) * 100;
        else growth = 100;
        
        matrixData.push({
          x: parseFloat(Math.min(Math.max(growth, -100), 300).toFixed(1)),
          y: p.curRev,
          r: p.curQty,
          label: p.code,
          fullName: p.name
        });
      }

      // Other Alerts
      if (p.curRev < p.prevRev && p.prevRev > 10000) {
        alerts.churn.push({ code: p.code, name: p.name, val: p.prevRev - p.curRev });
      }
      if (p.curQty > 0 && p.prevQty > 0 && p.curRev > 5000) {
        const curASP = p.curRev / p.curQty;
        const prevASP = p.prevRev / p.prevQty;
        if (p.curQty > p.prevQty && curASP < prevASP * 0.8) {
           const dropPct = ((prevASP - curASP) / prevASP * 100).toFixed(0);
           alerts.profit.push({ code: p.code, name: p.name, val: dropPct });
        }
      }
      if (p.recent7Rev > potentialThreshold && p.recent7Rev > 10000) {
        alerts.potential.push({ code: p.code, name: p.name, val: p.recent7Rev });
      }
    }

    const sliceTop5 = (list) => list.sort((a,b) => b.val - a.val).slice(0, 5);
    alerts.churn = sliceTop5(alerts.churn);
    alerts.stockout = sliceTop5(alerts.stockout); // 已經是 Top 20 篩選過的，這裡只取前 5 展示
    alerts.profit = sliceTop5(alerts.profit);
    alerts.potential = sliceTop5(alerts.potential);

    // 6. Group Summary Processing
    const groupSummary = [];
    Object.entries(groupHierarchy).forEach(([gName, gData]) => {
      if (gData.rev > 0) {
        const subs = [];
        Object.entries(gData.subs).forEach(([sName, sData]) => {
          if (sData.rev > 0) {
            // Sort Top 5 products by Qty
            const prods = Object.entries(sData.prods)
              .map(([code, pData]) => ({ code, qty: pData.qty, rev: pData.rev }))
              .sort((a, b) => b.qty - a.qty)
              .slice(0, 5);
            
            subs.push({
              name: sName,
              rev: sData.rev,
              qty: sData.qty,
              asp: Math.round(sData.rev / sData.qty),
              prods: prods
            });
          }
        });
        
        subs.sort((a,b) => b.rev - a.rev);

        groupSummary.push({
          name: gName,
          rev: gData.rev,
          qty: gData.qty,
          asp: Math.round(gData.rev / gData.qty),
          subs: subs
        });
      }
    });
    groupSummary.sort((a,b) => b.rev - a.rev);

    // 7. Pareto
    paretoDataRaw.sort((a, b) => b.curRev - a.curRev);
    const paretoData = { labels: [], revs: [], cumulative: [] };
    let runningSum = 0;
    for (const p of paretoDataRaw.slice(0, 20)) {
      runningSum += p.curRev;
      paretoData.labels.push(p.code);
      paretoData.revs.push(p.curRev);
      paretoData.cumulative.push(parseFloat(((runningSum / curStats.rev) * 100).toFixed(1)));
    }

    // 8. KPI
    const calcGrowth = (cur, prev) => (prev === 0 ? (cur > 0 ? 100 : 0) : ((cur - prev) / prev) * 100);
    const kpi = {
      rev: curStats.rev,
      rev_grow: calcGrowth(curStats.rev, prevStats.rev),
      qty: curStats.qty,
      qty_grow: calcGrowth(curStats.qty, prevStats.qty),
      asp: curStats.qty > 0 ? Math.round(curStats.rev / curStats.qty) : 0,
      asp_grow: calcGrowth((curStats.qty > 0 ? curStats.rev / curStats.qty : 0), (prevStats.qty > 0 ? prevStats.rev / prevStats.qty : 0))
    };

    return {
      status: 'success',
      kpi,
      groupSummary,
      alerts,
      trend: { labels: trendLabels, datasets: trendDatasets },
      matrix: matrixData,
      pareto: paretoData,
      meta: { start: startDateStr, end: endDateStr, dayDiff: dayDiff }
    };

  } catch (e) {
    return { status: 'error', message: e.toString() + e.stack };
  }
}
