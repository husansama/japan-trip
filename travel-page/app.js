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

function dayCard(day, index) {
  const food = day["吃法"] || {};
  const details = day["细节"] || [];
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

function bindEvents() {
  document.querySelector(".tabs").addEventListener("click", (event) => {
    const button = event.target.closest(".tab");
    if (!button) return;
    switchView(button.dataset.view);
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
}

renderHero();
renderChips();
renderDays();
renderFood();
renderSimpleLists();
bindEvents();
