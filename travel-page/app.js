const data = window.TRIP_DATA;

const state = {
  view: "days",
  city: "全部",
  query: "",
};

const $ = (selector) => document.querySelector(selector);

function text(value) {
  return value || "按当天情况调整";
}

function googleMapsUrl(route) {
  const params = new URLSearchParams({
    api: "1",
    origin: route["地图起点"],
    destination: route["地图终点"],
    travelmode: route["方式"] === "步行" ? "walking" : "transit",
  });
  return `https://www.google.com/maps/dir/?${params.toString()}`;
}

function routeList(routes) {
  return `
    <div class="route-list">
      ${routes.map(route => `
        <a class="route-row" href="${googleMapsUrl(route)}" target="_blank" rel="noopener noreferrer" title="在 Google Maps 打开实时路线">
          <div class="route-points">
            ${route["方案"] ? `<span class="route-option">${route["方案"]}</span>` : ""}
            ${route["起点"]} → ${route["终点"]}
          </div>
          <span class="route-time">${route["方式"]} · ${route["耗时"]}</span>
          ${route["说明"] ? `<span class="route-note">${route["说明"]}</span>` : ""}
        </a>
      `).join("")}
    </div>
    <div class="route-source">Google Maps 于 2026-07-30 日间查询；点击任一路段可按出发时刻查看实时班次。未含逛景点、排队及大型车站内找路时间。</div>
  `;
}

function includesQuery(item, query) {
  if (!query) return true;
  return JSON.stringify(item).toLowerCase().includes(query.toLowerCase());
}

function detectCity(day) {
  const blob = `${day["入住"]} ${day["主线"]} ${day["必去锚点"]}`;
  if (blob.includes("大阪") || blob.includes("USJ") || blob.includes("神户") || blob.includes("姬路")) return "关西";
  if (blob.includes("京都") || blob.includes("宇治") || blob.includes("奈良")) return "京都";
  if (blob.includes("东京") || blob.includes("富士") || blob.includes("镰仓")) return "东京";
  return "全部";
}

function renderHero() {
  $("#tripDates").textContent = data.trip.dates;
  $("#tripTitle").textContent = data.trip.title;
  $("#tripSubtitle").textContent = data.trip.subtitle;
  $("#tripRoute").textContent = data.trip.route;
  $("#updatedAt").textContent = `更新 ${data.updatedAt}`;
}

function renderChips() {
  const chips = ["全部", "关西", "京都", "东京"];
  $("#cityChips").innerHTML = chips
    .map((city) => `<button class="chip ${state.city === city ? "is-active" : ""}" data-city="${city}" type="button">${city}</button>`)
    .join("");
}

function attractionList(attractions) {
  return `
    <div class="attraction-list">
      ${attractions.map(item => `
        <article class="attraction-card">
          <div class="attraction-card-head">
            <h4>${item["名称"]}</h4>
            <span>${item["建议停留"]}</span>
          </div>
          <p>${item["说明"]}</p>
          <p class="attraction-card-tip"><strong>怎么玩：</strong>${item["怎么玩"]}</p>
          <p class="attraction-card-tip"><strong>取舍：</strong>${item["取舍"]}</p>
        </article>
      `).join("")}
    </div>
  `;
}

function passUseList(items) {
  return `
    <div class="pass-use-list">
      ${items.map(item => `
        <article class="pass-use-card">
          <div class="pass-use-kicker">${item["票券"]}</div>
          <h4>${item["项目"]}</h4>
          <p><strong>怎么用：</strong>${item["执行"]}</p>
          <p><strong>临场提醒：</strong>${item["提醒"]}</p>
        </article>
      `).join("")}
    </div>
  `;
}

function dayCard(day, index) {
  const food = day["吃法"] || {};
  const details = day["细节"] || [];
  const attractions = day["景点详解"] || [];
  const passUses = day["票券安排"] || [];
  const routes = day["交通分段"] || [];
  const open = index === 0 ? "open" : "";
  return `
    <details class="day-card" ${open}>
      <summary>
        <div class="day-top">
          <div>
            <div class="day-kicker">${day["天数"]} · ${day["日期"]} · ${day["星期"]}</div>
            <h3 class="day-title">${day["主线"]}</h3>
          </div>
          <span class="stay">${day["入住"]}</span>
        </div>
        <p class="anchors">${day["必去锚点"]}</p>
      </summary>
      <div class="day-body">
        <div class="field"><strong>P人玩法</strong><p>${day["P人友好玩法"]}</p></div>
        ${passUses.length ? `<div class="field"><strong>票券使用 · 当天照此执行</strong>${passUseList(passUses)}</div>` : ""}
        ${attractions.length ? `<div class="field"><strong>景点详解 · 看什么、怎么玩、累了删什么</strong>${attractionList(attractions)}</div>` : ""}
        <div class="field"><strong>交通</strong><p>${day["交通重点"]}</p></div>
        ${routes.length ? `<div class="field"><strong>逐段交通耗时 · ${routes.length}段</strong>${routeList(routes)}</div>` : ""}
        <div class="field"><strong>晚餐/夜间</strong><p>${day["晚餐/夜间建议"]}</p></div>
        <div class="field"><strong>当天吃法</strong><p>${text(food["午餐"])}；${text(food["晚餐"])}。${text(food["P人吃法"])}</p></div>
        ${details.length ? `<div class="field"><strong>分时段</strong><div class="detail-list">${details.map(detail => `
          <div class="detail-row">
            <div class="tag">${detail["时间块"]} · ${detail["城市"]}</div>
            <p>${detail["建议安排"]}</p>
          </div>`).join("")}</div></div>` : ""}
      </div>
    </details>
  `;
}

function renderDays() {
  const filtered = data.days.filter((day) => {
    const cityOk = state.city === "全部" || detectCity(day) === state.city;
    return cityOk && includesQuery(day, state.query);
  });
  $("#dayList").innerHTML = filtered.map(dayCard).join("") || `<div class="info-card"><h3>没有匹配内容</h3><p>换个关键词试试。</p></div>`;
}

function googleMapsSearchUrl(location) {
  const params = new URLSearchParams({
    api: "1",
    query: location["地图查询"],
  });
  return `https://www.google.com/maps/search/?${params.toString()}`;
}

function foodLocationLinks(locations) {
  if (!locations.length) return "";
  return `
    <div class="food-map-links">
      ${locations.map(location => `
        <a class="food-map-link" href="${googleMapsSearchUrl(location)}" target="_blank" rel="noopener noreferrer" title="在 Google Maps 打开 ${location["店名"]}">
          <span aria-hidden="true">📍</span>${location["店名"]}<span aria-hidden="true">↗</span>
        </a>
        ${location["定位说明"] ? `<span class="food-map-note">${location["定位说明"]}</span>` : ""}
      `).join("")}
    </div>
  `;
}

function renderFood() {
  const groups = data.foodMap.reduce((acc, item) => {
    const region = item["区域"] || "其他";
    acc[region] = acc[region] || [];
    acc[region].push(item);
    return acc;
  }, {});
  $("#foodGroups").innerHTML = Object.entries(groups).map(([region, items]) => `
    <div class="food-region">${region}</div>
    ${items.map(item => `
      <article class="food-card">
        <div class="meta"><span>${item["类型"]}</span><span>${item["预算感"]}</span><span>${item["适合日期"]}</span></div>
        <h3>${item["推荐店/吃法"]}</h3>
        ${foodLocationLinks(item["地图店铺"] || [])}
        <div class="field"><strong>预约/排队</strong><p>${item["预约/排队"]}</p></div>
        <div class="field"><strong>点单</strong><p>${item["点单建议"]}</p></div>
        <div class="field"><strong>备选</strong><p>${item["P人备选"]}</p></div>
      </article>
    `).join("")}
  `).join("");
}

function googleMapsFromHereUrl(hotel) {
  const params = new URLSearchParams({
    api: "1",
    destination: hotel["地图查询"],
    travelmode: "transit",
  });
  return `https://www.google.com/maps/dir/?${params.toString()}`;
}

function hotelList(hotels) {
  if (!hotels.length) return "";
  return `
    <div class="hotel-list">
      ${hotels.map(hotel => `
        <div class="hotel-pick">
          <div class="hotel-top">
            <div>
              <h4 class="hotel-name">${hotel["名称"]}</h4>
              <div class="hotel-area">${hotel["区域"]}</div>
            </div>
            <span class="hotel-budget">参考 ${hotel["预算感"]}</span>
          </div>
          <p class="hotel-reason">${hotel["适合理由"]}</p>
          <div class="hotel-actions">
            <a class="hotel-link" href="${googleMapsSearchUrl(hotel)}" target="_blank" rel="noopener noreferrer" title="在 Google Maps 查看 ${hotel["名称"]}">📍 地图定位</a>
            <a class="hotel-link hotel-link--route" href="${googleMapsFromHereUrl(hotel)}" target="_blank" rel="noopener noreferrer" title="从当前位置前往 ${hotel["名称"]}">从我这里出发 ↗</a>
          </div>
        </div>
      `).join("")}
    </div>
    <div class="hotel-location-note">“从我这里出发”不会写死起点；手机需允许 Google Maps 使用当前位置。价格按单人或一间基础房估算，仅作 ¥300–500 筛选参考；请以 2026年9月25日至10月6日 的实时含税总价为准。标有共用卫浴、青旅或胶囊的候选，下单前请再次确认房型。</div>
  `;
}

function lodgingCard(item) {
  const hotels = item["住宿推荐"] || [];
  return `
    <article class="info-card">
      <div class="meta">
        <span>${item["日期"]}</span>
        <span>${item["晚数"]}晚</span>
        <span>备选 ${item["备选区域"]}</span>
      </div>
      <h3>${item["阶段"]} · ${item["首选区域"]}</h3>
      <div class="field"><strong>为什么</strong><p>${item["为什么"]}</p></div>
      <div class="field"><strong>导游建议</strong><p>${item["导游建议"]}</p></div>
      ${hotels.length ? `<div class="field"><strong>具体住宿 · ${hotels.length}家</strong>${hotelList(hotels)}</div>` : ""}
    </article>
  `;
}

function simpleCard(title, meta, fields) {
  return `
    <article class="info-card">
      <div class="meta">${meta.filter(Boolean).map(item => `<span>${item}</span>`).join("")}</div>
      <h3>${title}</h3>
      ${fields.map(([label, value]) => `<div class="field"><strong>${label}</strong><p>${value}</p></div>`).join("")}
    </article>
  `;
}

function usjRating(score) {
  return "★".repeat(score) + "☆".repeat(5 - score);
}

function usjAssetPath(fileName) {
  const inTravelPage = /\/travel-page\/(?:index\.html)?$/.test(window.location.pathname);
  return inTravelPage ? `./assets/usj/${fileName}` : `./travel-page/assets/usj/${fileName}`;
}

function renderUSJRouteMap() {
  return `
    <div class="usj-map-scroll" role="region" aria-label="USJ园区路线示意图，可横向滚动">
      <svg class="usj-route-map" viewBox="0 0 900 660" role="img" aria-labelledby="usjMapTitle usjMapDesc">
        <title id="usjMapTitle">USJ园区逆时针游玩路线示意图</title>
        <desc id="usjMapDesc">从入口出发，依次经过好莱坞、小黄人、超级任天堂、侏罗纪、水世界、亲善村、哈利波特，再回到前区和万圣节夜间项目。</desc>
        <defs>
          <linearGradient id="usjMapBg" x1="0" y1="0" x2="1" y2="1">
            <stop offset="0" stop-color="#eef8f5" />
            <stop offset="1" stop-color="#fff6e6" />
          </linearGradient>
          <marker id="usjArrow" viewBox="0 0 10 10" refX="9" refY="5" markerWidth="7" markerHeight="7" orient="auto-start-reverse">
            <path d="M 0 0 L 10 5 L 0 10 z" fill="#d65f42" />
          </marker>
        </defs>
        <rect x="8" y="8" width="884" height="644" rx="30" fill="url(#usjMapBg)" stroke="#b8d3ca" stroke-width="4" />
        <path d="M30 80 C180 20 340 60 450 28 C600 -5 760 35 870 92 L870 330 C810 300 780 310 735 360 C650 455 555 585 450 630 C330 600 260 535 180 505 C95 475 45 390 30 300 Z" fill="#d8ece7" opacity="0.55" />

        <g class="usj-map-zone usj-map-zone--nintendo"><rect x="34" y="48" width="170" height="105" rx="22" /><text x="119" y="91">超级任天堂世界</text><text x="119" y="119">咚奇刚 · 马里奥 · 耀西</text></g>
        <g class="usj-map-zone"><rect x="68" y="205" width="190" height="118" rx="22" /><text x="163" y="249">侏罗纪公园</text><text x="163" y="277">飞天翼龙 · 乘船游</text></g>
        <g class="usj-map-zone"><rect x="280" y="45" width="165" height="98" rx="22" /><text x="362" y="88">水世界</text><text x="362" y="113">大型特技秀</text></g>
        <g class="usj-map-zone"><rect x="470" y="88" width="165" height="106" rx="22" /><text x="552" y="131">亲善村</text><text x="552" y="158">大白鲨</text></g>
        <g class="usj-map-zone usj-map-zone--harry"><rect x="650" y="148" width="205" height="132" rx="22" /><text x="752" y="197">哈利·波特</text><text x="752" y="225">禁忌之旅 · 鹰马</text><text x="752" y="250">城堡漫步</text></g>
        <g class="usj-map-zone"><rect x="658" y="332" width="180" height="96" rx="22" /><text x="748" y="373">奇境世界</text><text x="748" y="399">亲子项目区</text></g>
        <g class="usj-map-zone usj-map-zone--minion"><rect x="66" y="384" width="190" height="105" rx="22" /><text x="161" y="426">小黄人乐园</text><text x="161" y="454">大恶党任务 · 乘车游</text></g>
        <g class="usj-map-zone"><rect x="278" y="374" width="150" height="96" rx="22" /><text x="353" y="415">旧金山</text><text x="353" y="442">餐饮休整</text></g>
        <g class="usj-map-zone"><rect x="440" y="386" width="175" height="96" rx="22" /><text x="527" y="426">纽约</text><text x="527" y="453">柯南4-D</text></g>
        <g class="usj-map-zone usj-map-zone--hollywood"><rect x="318" y="500" width="260" height="105" rx="22" /><text x="448" y="542">好莱坞</text><text x="448" y="570">美梦乘车游 · SING</text></g>
        <g class="usj-map-entrance"><path d="M398 635 H502 L480 610 H420 Z" /><text x="450" y="649">入口／出口</text></g>

        <polyline class="usj-map-route-line" points="450,620 450,520 160,440 120,105 165,260 185,275 362,95 552,145 752,215 500,435 190,260" marker-end="url(#usjArrow)" />
        ${[
          [1, 450, 620], [2, 450, 520], [3, 160, 440], [4, 120, 105], [5, 165, 242],
          [6, 190, 285], [7, 362, 95], [8, 552, 145], [9, 752, 215], [10, 500, 435], [11, 190, 260]
        ].map(([number, x, y]) => `<g class="usj-map-marker"><circle cx="${x}" cy="${y}" r="17" /><text x="${x}" y="${y + 5}">${number}</text></g>`).join("")}
      </svg>
    </div>
    <div class="usj-map-legend">
      <span><i class="usj-legend-line"></i>推荐步行方向</span>
      <span><i class="usj-legend-dot"></i>数字对应下方时间轴</span>
    </div>
    <p class="usj-inline-note">这是为减少折返绘制的路线示意图，不是精确比例地图。现场定位、洗手间和临时封路请打开USJ官方App地图。</p>
  `;
}

function renderUSJ() {
  const guide = data.usj;
  if (!guide) return;

  const dates = guide.dates.map((item, index) => `
    <article class="usj-date-card ${index === 0 ? "is-primary" : ""}">
      <div class="usj-date-top">
        <h4>${item["日期"]}</h4>
        <span class="usj-status">${item["建议"]}</span>
      </div>
      <div class="usj-date-meta">
        <span>营业：${item["营业时间"]}</span>
        <span>预测：${item["预测平均等待"]} · ${item["年卡情况"]}</span>
        <span>${item["行程影响"]}</span>
      </div>
    </article>
  `).join("");

  const attractions = guide.attractions.map((item, index) => `
    <details class="usj-attraction" ${index < 2 ? "open" : ""}>
      <img class="usj-attraction-image" src="${usjAssetPath(item["图片"])}" alt="${item["名称"]}官方项目图片" loading="lazy" decoding="async" />
      <summary>
        <div class="usj-attraction-top">
          <div class="usj-attraction-title">
            <h4>${item["名称"]}</h4>
            <p>${item["区域"]} · ${item["类型"]}</p>
          </div>
          <span class="usj-wait">${item["预计排队"]}</span>
        </div>
        <div class="usj-rating" aria-label="推荐度${item["推荐度"]}星">${usjRating(item["推荐度"])}</div>
      </summary>
      <div class="usj-attraction-body">
        <div class="field"><strong>项目介绍</strong><p>${item["介绍"]}</p></div>
        <div class="field"><strong>怎么玩</strong><p>${item["玩法"]}</p></div>
        <div class="field"><strong>玩法诀窍</strong><p>${item["操作诀窍"]}</p></div>
        <div class="field"><strong>官网参数</strong><p>${item["官方参数"]}</p></div>
        <div class="field"><strong>预计实际占用</strong><p>${item["实际占用"]}</p></div>
        <div class="field"><strong>乘坐提醒</strong><p>${item["提醒"]}</p></div>
        <div class="field"><strong>速通建议</strong><p>${item["速通建议"]}</p></div>
        <div class="field"><strong>单人通道</strong><p>${item["单人通道"]}</p></div>
        <div class="field"><strong>最终取舍</strong><p>${item["取舍"]}</p></div>
        <a class="usj-official-link" href="${item["官网"]}" target="_blank" rel="noopener noreferrer">打开USJ官方项目页</a>
      </div>
    </details>
  `).join("");

  const halloween = guide.halloween.map(item => `
    <article class="usj-halloween-item">
      <img src="${usjAssetPath(item["图片"])}" alt="${item["名称"]}官方项目图片" loading="lazy" decoding="async" />
      <div class="usj-halloween-copy">
      <h4>${item["名称"]}</h4>
      <p><strong>${item["时间"]}</strong> · ${item["预计等待"]}</p>
      <p>${item["建议"]}</p>
      <a href="${item["官网"]}" target="_blank" rel="noopener noreferrer">查看官网详情 ↗</a>
      </div>
    </article>
  `).join("");

  const routeTimeline = timeline => timeline.map(item => `
    <article class="usj-route-step">
      <div class="usj-route-step-marker">${item["序号"]}</div>
      <div class="usj-route-step-content">
        <div class="usj-route-step-head">
          <div><span>${item["时间"]}</span><h4>${item["安排"]}</h4></div>
          <em>${item["区域"]}</em>
        </div>
        <div class="usj-route-facts">
          <span><strong>排队：</strong>${item["预计等待"]}</span>
          <span><strong>占用：</strong>${item["游玩占用"]}</span>
          <span><strong>通道：</strong>${item["通道"]}</span>
        </div>
        <div class="field"><strong>为什么这样排</strong><p>${item["理由"]}</p></div>
        <div class="field"><strong>现场调整</strong><p>${item["调整"]}</p></div>
      </div>
    </article>
  `).join("");

  const paidPassTimeline = routeTimeline(guide.route_timeline);
  const noPass4Timeline = routeTimeline(guide.route_timeline_no_pass4);

  const strategies = guide.strategies.map(item => `
    <article class="usj-strategy-card ${item["方案"].includes("Pass 4") ? "is-recommended" : ""}">
      <div class="usj-date-top">
        <h4>${item["方案"]}</h4>
        ${item["方案"].includes("Pass 4") ? '<span class="usj-decision">本行程首选</span>' : ""}
      </div>
      <div class="usj-strategy-meta">
        <span><strong>费用：</strong>${item["追加费用"]}</span>
        <span><strong>预计：</strong>${item["预计成果"]}</span>
      </div>
      <div class="field"><strong>优点</strong><p>${item["优点"]}</p></div>
      <div class="field"><strong>缺点</strong><p>${item["缺点"]}</p></div>
      <div class="field"><strong>适合</strong><p>${item["适合"]}</p></div>
    </article>
  `).join("");

  const expressPass4 = guide.express_pass4.map((item, index) => `
    <article class="usj-pass-card ${index === 0 ? "is-recommended" : ""}">
      <div class="usj-date-top">
        <h4>${item["名称"]}</h4>
        <span class="usj-decision">${item["标签"]}</span>
      </div>
      <ol class="usj-pass-projects">
        ${item["项目"].map(project => `<li>${project}</li>`).join("")}
      </ol>
      <div class="field"><strong>为什么选</strong><p>${item["推荐理由"]}</p></div>
      <div class="field"><strong>情侣玩法</strong><p>${item["情侣建议"]}</p></div>
    </article>
  `).join("");

  const singleRiderCouples = guide.single_rider_couples.map(item => `
    <article class="usj-single-card">
      <div class="usj-date-top">
        <h4>${item["项目"]}</h4>
        <span class="usj-single-status">${item["建议"]}</span>
      </div>
      <p>${item["情侣策略"]}</p>
    </article>
  `).join("");

  const storageCards = guide.storage_tips["必须存包"].map(item => `
    <article class="usj-pass-card is-recommended">
      <div class="usj-date-top">
        <h4>${item["项目"]}</h4>
        <span class="usj-decision">必须存包</span>
      </div>
      <div class="field"><strong>开柜方式</strong><p>${item["柜门方式"]}</p></div>
      <div class="field"><strong>操作技巧</strong><p>${item["操作技巧"]}</p></div>
    </article>
  `).join("");

  const sources = guide.sources.map(item => `
    <a class="usj-source-link" href="${item["网址"]}" target="_blank" rel="noopener noreferrer">${item["名称"]}</a>
  `).join("");

  $("#usjGuide").innerHTML = `
    <article class="usj-hero-card">
      <h3>9月28日执行：USJ门票＋禁忌之旅快速通1，不买速通4</h3>
      <p>${guide.summary["结论"]}</p>
      <p class="usj-small-note">${guide.summary["预测说明"]}</p>
      <p class="usj-small-note">当天动作：${guide.summary["当天动作"]}</p>
    </article>

    <section class="usj-section">
      <div class="usj-section-head"><h3>27、28还是29日</h3><span>当前预测，天气变化后需重查</span></div>
      <div class="usj-date-grid">${dates}</div>
    </section>

    <section class="usj-section usj-map-section">
      <div class="usj-section-head"><h3>园区大地图＋路线编号</h3><span>手机可横向滑动查看</span></div>
      ${renderUSJRouteMap()}
      <a class="usj-official-map-link" href="https://www.usj.co.jp/web/ja/jp/service-guide/parkmap" target="_blank" rel="noopener noreferrer">打开USJ官方实时地图</a>
    </section>

    <section class="usj-section">
      <div class="usj-section-head"><h3>存包技巧</h3><span>必须存包3类，其余小包通常可随身</span></div>
      <p class="usj-inline-note">${guide.storage_tips["总原则"]}</p>
      <div class="usj-pass-grid">${storageCards}</div>
      <ul class="usj-closure-list usj-single-rules">${guide.storage_tips["其他规则"].map(item => `<li>${item}</li>`).join("")}</ul>
    </section>

    <section class="usj-section">
      <div class="usj-section-head"><h3>主要项目</h3><span>点击项目展开介绍与取舍</span></div>
      <div class="usj-attraction-list">${attractions}</div>
    </section>

    <section class="usj-section">
      <div class="usj-section-head"><h3>万圣节限定</h3><span>9月27—29日都有普通万圣惊魂夜</span></div>
      <div class="usj-halloween-list">${halloween}</div>
    </section>

    <section class="usj-section">
      <div class="usj-section-head"><h3>Express Pass 4 怎么选</h3><span>两人各买一张，并一次下单</span></div>
      <div class="usj-pass-grid">${expressPass4}</div>
      <p class="usj-inline-note">套餐名称、项目和指定时段可能调整；购买9月27日票券时，以结算页列出的4个项目及区域入场保证为最终依据。</p>
    </section>

    <section class="usj-section">
      <div class="usj-section-head"><h3>若临时加购速通4的时间轴</h3><span>${guide.route_assumption["适用日期"]}</span></div>
      <div class="usj-route-assumption">
        <strong>${guide.route_assumption["目标套餐"]}</strong>
        <p>建议购票时段：${guide.route_assumption["建议时段"]}</p>
        <p>${guide.route_assumption["固定节点"]}</p>
        <p>${guide.route_assumption["说明"]}</p>
      </div>
      <div class="usj-route-timeline">${paidPassTimeline}</div>
    </section>

    <section class="usj-section">
      <div class="usj-section-head"><h3>不加购速通4：自带手环＋禁忌之旅快速通1</h3><span>${guide.no_pass4_assumption["适用日期"]}</span></div>
      <div class="usj-route-assumption">
        <strong>${guide.no_pass4_assumption["目标套餐"]}</strong>
        <p>建议节奏：${guide.no_pass4_assumption["建议时段"]}</p>
        <p>${guide.no_pass4_assumption["固定节点"]}</p>
        <p>${guide.no_pass4_assumption["说明"]}</p>
      </div>
      <div class="usj-route-timeline">${noPass4Timeline}</div>
    </section>

    <section class="usj-section">
      <div class="usj-section-head"><h3>情侣使用单人通道</h3><span>省时间，但必须接受拆开乘坐</span></div>
      <div class="usj-single-grid">${singleRiderCouples}</div>
      <ul class="usj-closure-list usj-single-rules">${guide.single_rider_rules.map(item => `<li>${item}</li>`).join("")}</ul>
    </section>

    <section class="usj-section">
      <div class="usj-section-head"><h3>速通购买方案</h3><span>入园票需要另买</span></div>
      <div class="usj-strategy-grid">${strategies}</div>
    </section>

    <section class="usj-section">
      <div class="usj-section-head"><h3>休止与临场风险</h3><span>出发前一晚再检查</span></div>
      <ul class="usj-closure-list">${guide.closures.map(item => `<li>${item}</li>`).join("")}</ul>
    </section>

    <section class="usj-section">
      <div class="usj-section-head"><h3>官网与数据入口</h3><span>手机可直接打开</span></div>
      <div class="usj-source-list">${sources}</div>
    </section>
  `;
}

function renderSimpleLists() {
  $("#lodgingList").innerHTML = data.lodging.map(lodgingCard).join("");

  $("#planbList").innerHTML = data.alternatives.map(item => simpleCard(
    item["场景"],
    [item["原计划"]],
    [["改法", item["改法"]], ["收益", item["收益"]], ["备注", item["备注"]]]
  )).join("");

  $("#transportList").innerHTML = data.transport.map(item => simpleCard(
    item["事项"],
    [item["建议时间"]],
    [["建议做法", item["建议做法"]], ["风险点", item["风险点"]], ["P人提醒", item["P人提醒"]]]
  )).join("");

  $("#checkList").innerHTML = data.checklist.map(item => simpleCard(
    `${item["类别"]} · ${item["动作"]}`,
    [item["优先级"]],
    [["建议", item["建议"]], ["为什么", item["为什么"]]]
  )).join("");
}

function switchView(view) {
  state.view = view;
  document.querySelectorAll(".tab").forEach(tab => tab.classList.toggle("is-active", tab.dataset.view === view));
  document.querySelectorAll(".content-view").forEach(panel => panel.classList.remove("is-visible"));
  $(`#${view}View`).classList.add("is-visible");
  $("#dayToolbar").style.display = view === "days" ? "block" : "none";
}

function viewFromHash() {
  const requested = window.location.hash.replace("#", "");
  const validViews = ["days", "food", "lodging", "usj", "planb", "prep"];
  return validViews.includes(requested) ? requested : "days";
}

function bindEvents() {
  document.querySelector(".tabs").addEventListener("click", (event) => {
    const button = event.target.closest(".tab");
    if (!button) return;
    switchView(button.dataset.view);
    history.replaceState(null, "", `#${button.dataset.view}`);
  });

  $("#cityChips").addEventListener("click", (event) => {
    const button = event.target.closest(".chip");
    if (!button) return;
    state.city = button.dataset.city;
    renderChips();
    renderDays();
  });

  $("#searchInput").addEventListener("input", (event) => {
    state.query = event.target.value.trim();
    renderDays();
  });

  window.addEventListener("hashchange", () => switchView(viewFromHash()));
}

renderHero();
renderChips();
renderDays();
renderFood();
renderUSJ();
renderSimpleLists();
bindEvents();
const initialView = viewFromHash();
switchView(initialView);
if (initialView !== "days") {
  requestAnimationFrame(() => document.querySelector(".tabs").scrollIntoView());
}
