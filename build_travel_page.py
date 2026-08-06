from __future__ import annotations

import json
from pathlib import Path

from openpyxl import load_workbook


ROOT = Path("travel-page")
WORKBOOK = Path("日本12天11晚_0925-1006_P人关西加强版_含美食.xlsx")


def segment(
    start: str,
    end: str,
    duration: str,
    mode: str,
    *,
    option: str = "",
    note: str = "",
    map_start: str = "",
    map_end: str = "",
) -> dict[str, str]:
    """统一保存分段交通，并保留用于生成 Google Maps 实时路线的精确地点名。"""
    return {
        "起点": start,
        "终点": end,
        "耗时": duration,
        "方式": mode,
        "方案": option,
        "说明": note,
        "地图起点": map_start or start,
        "地图终点": map_end or end,
    }


# 耗时于 2026-07-30 按 Google Maps 日间推荐路线核对；班次、候车与拥堵会让结果浮动。
SEGMENT_ROUTES: dict[str, list[dict[str, str]]] = {
    "Day1": [
        segment("关西机场 KIX", "大阪站/梅田", "约47分钟", "公共交通", option="住梅田", note="HARUKA 直达", map_start="Kansai International Airport", map_end="Osaka Station"),
        segment("关西机场 KIX", "难波站", "约46分钟", "公共交通", option="住难波", note="南海机场线；特急班次可约42分钟", map_start="Kansai International Airport", map_end="Namba Station Osaka"),
        segment("大阪站/梅田", "道顿堀", "约19分钟", "公共交通", option="住梅田", note="御堂筋线到难波后步行", map_start="Osaka Station", map_end="Dotonbori Osaka"),
        segment("难波站", "道顿堀", "约5分钟", "步行", option="住难波", map_start="Namba Station Osaka", map_end="Dotonbori Osaka"),
    ],
    "Day2": [
        segment("大阪站", "大阪城", "约29分钟", "公共交通", note="含车站至天守阁区域步行", map_start="Osaka Station", map_end="Osaka Castle"),
        segment("大阪城", "梅田", "约28分钟", "公共交通", note="步行接谷町线", map_start="Osaka Castle", map_end="Umeda Station"),
        segment("梅田", "中崎町", "约12分钟", "步行", map_start="Umeda Station", map_end="Nakazakicho Osaka"),
        segment("中崎町", "心斋桥", "约19分钟", "公共交通", note="步行至梅田后乘御堂筋线", map_start="Nakazakicho Station", map_end="Shinsaibashi Station"),
        segment("心斋桥", "道顿堀", "约12分钟", "步行", map_start="Shinsaibashi Station", map_end="Dotonbori Osaka"),
        segment("道顿堀", "大阪站", "约18分钟", "公共交通", note="难波乘御堂筋线", map_start="Dotonbori Osaka", map_end="Osaka Station"),
    ],
    "Day3": [
        segment("大阪站", "姬路站", "约1小时2分钟", "公共交通", note="JR新快速，无需预约", map_start="Osaka Station", map_end="Himeji Station"),
        segment("姬路站", "姬路城", "约20分钟", "步行", map_start="Himeji Station", map_end="Himeji Castle"),
        segment("姬路城", "神户三宫", "约1小时1分钟", "公共交通", note="含步行回姬路站", map_start="Himeji Castle", map_end="Sannomiya Station Kobe"),
        segment("三宫站", "北野异人馆", "约16分钟", "步行", note="上坡路段", map_start="Sannomiya Station Kobe", map_end="Kitano Ijinkan Kobe"),
        segment("北野异人馆", "南京町", "约24分钟", "公共交通", note="公交加步行", map_start="Kitano Ijinkan Kobe", map_end="Nankinmachi Kobe"),
        segment("南京町", "Harborland", "约19分钟", "步行", map_start="Nankinmachi Kobe", map_end="Kobe Harborland"),
        segment("Harborland", "大阪站", "约34分钟", "公共交通", note="从神户站乘JR新快速", map_start="Kobe Harborland", map_end="Osaka Station"),
    ],
    "Day4": [
        segment("大阪站", "环球城站", "约11分钟", "公共交通", note="大阪环状线接JR梦咲线", map_start="Osaka Station", map_end="Universal City Station Osaka"),
        segment("环球城站", "大阪站", "约15分钟", "公共交通", note="散场高峰另预留进站时间", map_start="Universal City Station Osaka", map_end="Osaka Station"),
    ],
    "Day5": [
        segment("大阪站", "四条站", "约38分钟", "公共交通", note="当前最快路线经京都站换乌丸线；拖箱建议多留10分钟", map_start="Osaka Station", map_end="Shijo Station Kyoto"),
        segment("四条站", "清水寺", "约29分钟", "公共交通", note="公交加步行；拥堵时可能更久", map_start="Shijo Station Kyoto", map_end="Kiyomizu-dera Kyoto"),
        segment("清水寺", "三年坂", "约3分钟", "步行", map_start="Kiyomizu-dera Kyoto", map_end="Sannenzaka Kyoto"),
        segment("三年坂", "八坂神社", "约14分钟", "步行", map_start="Sannenzaka Kyoto", map_end="Yasaka Shrine Kyoto"),
        segment("八坂神社", "祇园白川", "约9分钟", "步行", map_start="Yasaka Shrine Kyoto", map_end="Gion Shirakawa Kyoto"),
        segment("祇园白川", "先斗町", "约6分钟", "步行", map_start="Gion Shirakawa Kyoto", map_end="Pontocho Alley Kyoto"),
        segment("先斗町", "四条站", "约15分钟", "步行", map_start="Pontocho Alley Kyoto", map_end="Shijo Station Kyoto"),
    ],
    "Day6": [
        segment("四条站", "阪急岚山站", "约18分钟", "公共交通", note="乌丸站乘阪急，桂站换乘", map_start="Shijo Station Kyoto", map_end="Arashiyama Station Hankyu"),
        segment("阪急岚山站", "天龙寺", "约13分钟", "步行", map_start="Arashiyama Station Hankyu", map_end="Tenryu-ji Kyoto"),
        segment("天龙寺", "竹林小径", "约1–5分钟", "步行", note="入口紧邻天龙寺北门", map_start="Tenryu-ji Kyoto", map_end="Arashiyama Bamboo Forest"),
        segment("竹林小径", "渡月桥", "约13分钟", "步行", map_start="Arashiyama Bamboo Forest", map_end="Togetsukyo Bridge"),
        segment("渡月桥", "金阁寺", "约50分钟", "公共交通", note="公交换乘，候车影响较大", map_start="Togetsukyo Bridge", map_end="Kinkaku-ji Kyoto"),
        segment("金阁寺", "龙安寺", "约11分钟", "公共交通", note="59路公交加步行", map_start="Kinkaku-ji Kyoto", map_end="Ryoan-ji Kyoto"),
        segment("龙安寺", "四条站", "约40分钟", "公共交通", note="公交接乌丸线", map_start="Ryoan-ji Kyoto", map_end="Shijo Station Kyoto"),
    ],
    "Day7": [
        segment("四条站", "JR宇治站", "约34分钟", "公共交通", note="乌丸线至京都站换JR奈良线", map_start="Shijo Station Kyoto", map_end="Uji Station JR"),
        segment("JR宇治站", "平等院", "约9分钟", "步行", map_start="Uji Station JR", map_end="Byodoin Kyoto"),
        segment("平等院", "JR宇治站", "约9分钟", "步行", map_start="Byodoin Kyoto", map_end="Uji Station JR"),
        segment("JR宇治站", "JR奈良站", "约44分钟", "公共交通", note="JR奈良线", map_start="Uji Station JR", map_end="Nara Station JR"),
        segment("JR奈良站", "奈良公园", "约22分钟", "步行", map_start="Nara Station JR", map_end="Nara Park"),
        segment("奈良公园巴士总站", "若草山北口", "约39分钟", "步行", note="公园内部范围很大，此段按明确入口核对", map_start="Nara Park Bus Terminal", map_end="若草山登山口 北ゲート"),
        segment("若草山一带", "四条站", "约1小时21分钟", "公共交通", note="步行至公交站后转近铁/地铁", map_start="Mount Wakakusa", map_end="Shijo Station Kyoto"),
    ],
    "Day8": [
        segment("四条站", "贵船口站", "约58分钟", "公共交通", option="方案A 贵船鞍马", note="公交接叡山电铁", map_start="Shijo Station Kyoto", map_end="Kibuneguchi Station"),
        segment("贵船口站", "贵船神社", "约8分钟", "公共交通", option="方案A 贵船鞍马", note="33路公交；步行约30分钟", map_start="Kibuneguchi Station", map_end="Kifune Shrine Kyoto"),
        segment("贵船神社", "鞍马站", "约21分钟", "公共交通", option="方案A 贵船鞍马", note="公交回贵船口后乘叡山电铁；沿公路步行约43分钟", map_start="Kifune Shrine Kyoto", map_end="Kurama Station Kyoto"),
        segment("鞍马站", "四条站", "约1小时4分钟", "公共交通", option="方案A 贵船鞍马", note="叡山电铁接公交/地铁", map_start="Kurama Station Kyoto", map_end="Shijo Station Kyoto"),
        segment("四条站", "三千院", "约1小时19分钟", "公共交通", option="方案B 大原", note="直达公交加步行", map_start="Shijo Station Kyoto", map_end="Sanzen-in Kyoto"),
        segment("三千院", "宝泉院", "约3分钟", "步行", option="方案B 大原", map_start="Sanzen-in Kyoto", map_end="Hosen-in Kyoto"),
        segment("宝泉院", "四条站", "约1小时5分钟", "公共交通", option="方案B 大原", note="步行至大原站乘17路", map_start="Hosen-in Kyoto", map_end="Shijo Station Kyoto"),
        segment("四条站", "伏见稻荷", "约17分钟", "公共交通", option="方案C 市区", note="乌丸线至京都站换JR奈良线", map_start="Shijo Station Kyoto", map_end="Fushimi Inari Taisha"),
        segment("伏见稻荷", "锦市场", "约24分钟", "公共交通", option="方案C 市区", note="京阪线加步行", map_start="Fushimi Inari Taisha", map_end="Nishiki Market Kyoto"),
        segment("锦市场", "新京极", "约2分钟", "步行", option="方案C 市区", map_start="Nishiki Market Kyoto", map_end="Shinkyogoku Shopping Street"),
        segment("新京极", "四条站", "约14分钟", "步行", option="方案C 市区", map_start="Shinkyogoku Shopping Street", map_end="Shijo Station Kyoto"),
    ],
    "Day9": [
        segment("四条站", "京都站", "约4分钟", "公共交通", note="乌丸线；另留进站与找站台时间", map_start="Shijo Station Kyoto", map_end="Kyoto Station"),
        segment("京都站", "东京站", "约2小时11分钟", "公共交通", note="东海道新干线当前最快班次", map_start="Kyoto Station", map_end="Tokyo Station"),
        segment("东京站", "新宿站", "约14分钟", "公共交通", note="JR中央线", map_start="Tokyo Station", map_end="Shinjuku Station"),
        segment("新宿站", "六本木Hills", "约18分钟", "公共交通", option="夜景A", note="大江户线加步行", map_start="Shinjuku Station", map_end="Roppongi Hills"),
        segment("六本木Hills", "新宿站", "约17分钟", "公共交通", option="夜景A", map_start="Roppongi Hills", map_end="Shinjuku Station"),
        segment("新宿站", "涩谷十字路口", "约8分钟", "公共交通", option="夜景B", note="JR湘南新宿线；山手线也可", map_start="Shinjuku Station", map_end="Shibuya Scramble Crossing"),
    ],
    "Day10": [
        segment("新宿站", "明治神宫", "约7分钟", "公共交通", note="到参宫桥一带；进入主殿另需步行", map_start="Shinjuku Station", map_end="Meiji Jingu"),
        segment("明治神宫主殿", "原宿站", "约10分钟", "步行", map_start="Meiji Jingu Main Hall", map_end="Harajuku Station"),
        segment("原宿站", "表参道站", "约13分钟", "步行", map_start="Harajuku Station", map_end="Omotesando Station"),
        segment("表参道站", "涩谷十字路口", "约17分钟", "步行", map_start="Omotesando Station", map_end="Shibuya Scramble Crossing"),
        segment("涩谷站", "新宿站", "约7分钟", "公共交通", note="JR山手线", map_start="Shibuya Station", map_end="Shinjuku Station"),
    ],
    "Day11": [
        segment("新宿站", "河口湖站", "约1小时46分钟", "公共交通", option="方案A 河口湖", note="Busta新宿高速巴士；需预订", map_start="Shinjuku Station", map_end="Kawaguchiko Station"),
        segment("河口湖站", "新宿站", "约1小时46分钟", "公共交通", option="方案A 河口湖", note="高速巴士；返程也建议预订", map_start="Kawaguchiko Station", map_end="Shinjuku Station"),
        segment("新宿站", "镰仓站", "约1小时", "公共交通", option="方案B 镰仓江之岛", note="JR湘南新宿线", map_start="Shinjuku Station", map_end="Kamakura Station"),
        segment("镰仓站", "江之岛站", "约25分钟", "公共交通", option="方案B 镰仓江之岛", note="江之电", map_start="Kamakura Station", map_end="Enoshima Station"),
        segment("江之岛站", "新宿站", "约1小时19分钟", "公共交通", option="方案B 镰仓江之岛", note="步行至片濑江之岛后乘小田急", map_start="Enoshima Station", map_end="Shinjuku Station"),
        segment("新宿站", "银座站", "约15分钟", "公共交通", option="方案C 东京市区", note="丸之内线", map_start="Shinjuku Station", map_end="Ginza Station"),
        segment("银座站", "浅草寺", "约20分钟", "公共交通", option="方案C 东京市区", note="银座线", map_start="Ginza Station", map_end="Senso-ji"),
        segment("浅草寺", "秋叶原站", "约11分钟", "公共交通", option="方案C 东京市区", note="步行至筑波快线浅草站", map_start="Senso-ji", map_end="Akihabara Station"),
        segment("秋叶原站", "新宿站", "约23分钟", "公共交通", option="方案C 东京市区", note="JR中央·总武线", map_start="Akihabara Station", map_end="Shinjuku Station"),
    ],
    "Day12": [
        segment("新宿站", "浅草寺", "约30分钟", "公共交通", option="上午选浅草", note="中央线接银座线", map_start="Shinjuku Station", map_end="Senso-ji"),
        segment("浅草寺", "成田机场", "约1小时12分钟", "公共交通", option="浅草→成田", note="浅草线/京成线；另留值机时间", map_start="Senso-ji", map_end="Narita International Airport"),
        segment("浅草寺", "羽田T3", "约40分钟", "公共交通", option="浅草→羽田", note="浅草线直通；另留值机时间", map_start="Senso-ji", map_end="Haneda Airport Terminal 3"),
        segment("新宿站", "银座站", "约15分钟", "公共交通", option="上午选银座", note="丸之内线", map_start="Shinjuku Station", map_end="Ginza Station"),
        segment("银座站", "成田机场", "约1小时13分钟", "公共交通", option="银座→成田", note="当前推荐为机场巴士；另留值机时间", map_start="Ginza Station", map_end="Narita International Airport"),
        segment("银座站", "羽田T3", "约25分钟", "公共交通", option="银座→羽田", note="有乐町换山手线及东京单轨；另留值机时间", map_start="Ginza Station", map_end="Haneda Airport Terminal 3"),
    ],
}


def food_pin(name: str, query: str, note: str = "") -> dict[str, str]:
    """保存已核对的店名和地图检索词，前端据此生成可直接点击的 Google Maps 定位。"""
    return {"店名": name, "地图查询": query, "定位说明": note}


# 原攻略中部分条目是区域型推荐；这里为每类补充可落地导航的代表店或明确地点。
FOOD_LOCATIONS: dict[tuple[str, str], list[dict[str, str]]] = {
    ("大阪难波/道顿堀", "大阪烧"): [
        food_pin("美津の", "Mizuno Okonomiyaki Dotonbori Osaka"),
        food_pin("味乃家本店", "Ajinoya Honten Osaka"),
        food_pin("千房道顿堀", "Chibo Dotonbori Building"),
    ],
    ("大阪难波/道顿堀", "章鱼烧"): [
        food_pin("道顿堀くくる", "Dotonbori Kukuru Konamon Museum"),
        food_pin("十八番道顿堀店", "Takoyaki Juhachiban Dotonbori"),
    ],
    ("大阪新世界/难波", "串炸"): [
        food_pin("串かつだるま新世界总本店", "Kushikatsu Daruma Shinsekai Sohonten"),
    ],
    ("大阪梅田", "百货美食"): [
        food_pin("阪神百货梅田本店", "Hanshin Department Store Umeda Main Store"),
        food_pin("KITTE Osaka", "KITTE Osaka"),
    ],
    ("大阪梅田/天满", "寿司/居酒屋"): [
        food_pin("春驹本店", "Harukoma Honten Osaka"),
        food_pin("寿司酒场さしす梅田", "Sushi Sakaba Sashisu Umeda"),
    ],
    ("大阪全城", "芝士蛋糕/甜品"): [
        food_pin("Rikuro老爷爷难波本店", "Rikuro Ojisan no Mise Namba Main Store"),
    ],
    ("姬路", "当地午餐"): [
        food_pin("柊 穴子料理", "Anago Restaurant Hiiragi Himeji"),
        food_pin("姬路おでん能古", "Himeji Oden Nodaya"),
    ],
    ("神户三宫/元町", "神户牛"): [
        food_pin("Steakland Kobe-kan", "Steakland Kobe-kan"),
        food_pin("Mouriya本店", "Mouriya Honten Kobe"),
        food_pin("Wakkoqu新神户", "Wakkoqu Shin-Kobe"),
    ],
    ("神户南京町", "小吃"): [
        food_pin("老祥记", "Roushouki Kobe Nankinmachi"),
        food_pin("YUNYUN", "YUNYUN Kobe Nankinmachi"),
    ],
    ("神户港湾", "甜品/咖啡"): [
        food_pin("Kobe Harborland umie MOSAIC", "Kobe Harborland umie MOSAIC", "商场内可现场选择咖啡和甜品"),
    ],
    ("USJ", "园区餐"): [
        food_pin("Kinopio's Cafe", "Kinopios Cafe Universal Studios Japan"),
        food_pin("Three Broomsticks", "Three Broomsticks Universal Studios Japan"),
    ],
    ("京都祇园/东山", "京料理/天妇罗"): [
        food_pin("祇园天ぷら八坂圓堂", "Gion Tempura Endo Yasaka Kyoto"),
    ],
    ("京都先斗町/木屋町", "居酒屋/京味小菜"): [
        food_pin("Pontocho Robin", "Pontocho Robin Kyoto"),
    ],
    ("京都清水寺周边", "小吃甜品"): [
        food_pin("MACCHA HOUSE清水产宁坂", "MACCHA HOUSE Kyoto Kiyomizu Sannenzaka"),
        food_pin("伊藤久右卫门清水坂店", "Itoh Kyuemon Kiyomizu-saka Store"),
    ],
    ("京都岚山", "汤豆腐/荞麦"): [
        food_pin("松籁庵", "Shoraian Arashiyama Kyoto"),
        food_pin("岚山よしむら", "Arashiyama Yoshimura Kyoto"),
    ],
    ("京都岚山", "咖啡"): [
        food_pin("% ARABICA京都岚山", "% ARABICA Kyoto Arashiyama"),
    ],
    ("京都站", "拉面/车站餐"): [
        food_pin("京都拉面小路", "Kyoto Ramen Koji"),
        food_pin("Kyoto Porta Dining", "Kyoto Porta Dining"),
    ],
    ("宇治", "抹茶正餐/甜品"): [
        food_pin("中村藤吉本店", "Nakamura Tokichi Honten Uji"),
        food_pin("伊藤久右卫门宇治本店", "Itoh Kyuemon Uji Honten"),
        food_pin("通圆茶屋本店", "Tsuen Tea Uji"),
    ],
    ("宇治", "鳗鱼/京料理"): [
        food_pin("京・宇治 抹茶料理辰巳屋", "Kyo-Uji Shunsai Tatsumiya"),
    ],
    ("奈良", "小吃"): [
        food_pin("中谷堂", "Nakatanidou Nara"),
        food_pin("大佛布丁梦风广场店", "Mahoroba Daibutsu Purin Yumekaze Square Nara"),
    ],
    ("奈良", "正餐"): [
        food_pin("志津香公园店", "Shizuka Park Store Nara"),
        food_pin("平宗奈良店", "Hiraso Nara Store"),
    ],
    ("贵船/鞍马", "山里午餐"): [
        food_pin("贵船ひろ文", "Kibune Hirobun Kyoto"),
        food_pin("贵船伝兵衛", "Kibune Denbei Kyoto"),
    ],
    ("大原", "乡土料理"): [
        food_pin("志野 松门", "Shino Shoumon Ohara Kyoto"),
    ],
    ("伏见/京都市区", "清酒/拉面/市场"): [
        food_pin("鸟せい本店", "Torisei Honten Fushimi Kyoto"),
        food_pin("锦市场", "Nishiki Market Kyoto"),
    ],
    ("东京涩谷", "乌冬/夜景餐"): [
        food_pin("TsuruTonTan涩谷", "TsuruTonTan Shibuya Scramble Square"),
    ],
    ("东京原宿/表参道", "逛街午餐"): [
        food_pin("Maisen青山本店", "Tonkatsu Maisen Aoyama Honten"),
        food_pin("bills表参道", "bills Omotesando"),
    ],
    ("东京浅草", "老派早餐/甜品"): [
        food_pin("Ginza Brazil", "Ginza Brazil Asakusa"),
        food_pin("浅草Silk Pudding", "Asakusa Silk Pudding"),
    ],
    ("东京银座/东京站", "购物日餐饮"): [
        food_pin("根室花丸 KITTE丸之内", "Nemuro Hanamaru KITTE Marunouchi"),
        food_pin("六厘舍 东京拉面街", "Rokurinsha Tokyo Station"),
        food_pin("银座 篝本店", "Ginza Kagari Honten"),
    ],
    ("河口湖", "地方菜"): [
        food_pin("ほうとう不动 河口湖站前店", "Houtou Fudou Kawaguchiko Station"),
        food_pin("甲州ほうとう小作 河口湖店", "Kosaku Kawaguchiko"),
    ],
    ("镰仓/江之岛", "海边饭"): [
        food_pin("秋本", "Akimoto Kamakura"),
        food_pin("とびっちょ江之岛本店", "Tobiccho Enoshima"),
        food_pin("bills七里滨", "bills Shichirigahama"),
    ],
    ("机场/伴手礼", "伴手礼"): [
        food_pin("Tokyo Milk Cheese Factory羽田机场", "Tokyo Milk Cheese Factory Haneda Airport"),
    ],
}


# 新增条目优先补足原攻略较少的拉面、寿司、咖啡与东京单店选择。
EXTRA_FOOD_RECOMMENDATIONS: list[dict[str, object]] = [
    {
        "区域": "大阪福岛",
        "类型": "拉面",
        "推荐店/吃法": "燃えよ麺助：鸭肉/贝类系拉面，适合从梅田步行或坐一站去吃",
        "适合日期": "Day2/Day4",
        "预算感": "低-中",
        "预约/排队": "不能预约，常排队；开门前或午后错峰",
        "点单建议": "鸭出汁酱油拉面、贝类盐味拉面",
        "P人备选": "队伍太长就回梅田地下街吃拉面",
        "地图店铺": [food_pin("燃えよ麺助", "Moeyo Mensuke Osaka")],
    },
    {
        "区域": "大阪中央市场",
        "类型": "寿司",
        "推荐店/吃法": "Endo Sushi中央市场店：传统抓握寿司，适合愿意早起的人",
        "适合日期": "Day2",
        "预算感": "中",
        "预约/排队": "早市型店铺，出发前确认营业日与结束时间",
        "点单建议": "先点一盘综合寿司，不够再加",
        "P人备选": "不早起就改梅田春驹或寿司酒场",
        "地图店铺": [food_pin("Endo Sushi中央市场店", "Endo Sushi Central Market Osaka")],
    },
    {
        "区域": "大阪心斋桥",
        "类型": "精品咖啡",
        "推荐店/吃法": "LiLo Coffee Roasters：逛心斋桥时顺路补一杯",
        "适合日期": "Day1/Day2",
        "预算感": "低-中",
        "预约/排队": "无需预约，座位不多",
        "点单建议": "手冲或拿铁，豆子可作轻便伴手礼",
        "P人备选": "满座就外带继续逛",
        "地图店铺": [food_pin("LiLo Coffee Roasters", "LiLo Coffee Roasters Osaka")],
    },
    {
        "区域": "姬路站周边",
        "类型": "乌冬",
        "推荐店/吃法": "Menme：手打乌冬，适合作为姬路城往返途中的午餐",
        "适合日期": "Day3",
        "预算感": "低-中",
        "预约/排队": "店面不大，午餐时段可能排队",
        "点单建议": "冷乌冬更显筋道，天气凉则选热汤乌冬",
        "P人备选": "赶车就改姬路站内定食",
        "地图店铺": [food_pin("Menme", "Menme Himeji Udon")],
    },
    {
        "区域": "神户三宫",
        "类型": "神户洋食",
        "推荐店/吃法": "Grill Ippei：不吃高价神户牛时，可选蛋包饭、炸虾等老派洋食",
        "适合日期": "Day3",
        "预算感": "中",
        "预约/排队": "饭点可能等位，适合提前或延后吃",
        "点单建议": "蛋包饭、炸虾、汉堡排",
        "P人备选": "等位太久就去南京町少量多吃",
        "地图店铺": [food_pin("Grill Ippei三宫", "Grill Ippei Sannomiya Kobe")],
    },
    {
        "区域": "京都银阁寺",
        "类型": "乌冬",
        "推荐店/吃法": "Omen银阁寺本店：蔬菜配料丰富，口味清爽",
        "适合日期": "Day8自由调整",
        "预算感": "中",
        "预约/排队": "热门时段会排队，开门前后更稳",
        "点单建议": "招牌Omen乌冬",
        "P人备选": "不去银阁寺则无需专程跨城",
        "地图店铺": [food_pin("Omen银阁寺本店", "Omen Ginkakuji Kyoto")],
    },
    {
        "区域": "京都冈崎",
        "类型": "乌冬",
        "推荐店/吃法": "山元麺蔵：高人气手打乌冬，只适合愿意为一餐留时间的人",
        "适合日期": "Day5/Day8",
        "预算感": "中",
        "预约/排队": "规则可能调整，出发前看地图页最新公告",
        "点单建议": "牛蒡天妇罗乌冬",
        "P人备选": "不为网红店打乱当天东山动线",
        "地图店铺": [food_pin("山元麺蔵", "Yamamoto Menzou Kyoto")],
    },
    {
        "区域": "京都市中心",
        "类型": "荞麦",
        "推荐店/吃法": "本家尾张屋本店：京都老字号荞麦，适合想吃传统口味",
        "适合日期": "Day5/Day8",
        "预算感": "中",
        "预约/排队": "午餐可能排队，闭店较早",
        "点单建议": "宝来荞麦或天妇罗荞麦",
        "P人备选": "行程不顺路就选京都站或河原町分店型餐饮",
        "地图店铺": [food_pin("本家尾张屋本店", "Honke Owariya Honten Kyoto")],
    },
    {
        "区域": "京都市中心",
        "类型": "咖啡/早餐",
        "推荐店/吃法": "Inoda Coffee本店适合老派早餐；Weekenders适合快速喝精品咖啡",
        "适合日期": "Day5-Day9",
        "预算感": "低-中",
        "预约/排队": "Inoda早餐可能排队；Weekenders座位很少",
        "点单建议": "Inoda早餐套餐；Weekenders手冲/拿铁",
        "P人备选": "按当天动线二选一，不必都打卡",
        "地图店铺": [
            food_pin("Inoda Coffee本店", "Inoda Coffee Honten Kyoto"),
            food_pin("Weekenders Coffee富小路", "Weekenders Coffee Tominokoji Kyoto"),
        ],
    },
    {
        "区域": "京都四条乌丸",
        "类型": "创意寿司",
        "推荐店/吃法": "AWOMB乌丸本店：手织寿司摆盘漂亮，适合安排一顿仪式感午餐",
        "适合日期": "Day5/Day8",
        "预算感": "中-高",
        "预约/排队": "建议提前预约并确认当日套餐",
        "点单建议": "手织寿司套餐",
        "P人备选": "约不到就去锦市场/河原町吃普通寿司",
        "地图店铺": [food_pin("AWOMB乌丸本店", "AWOMB Karasuma Honten Kyoto")],
    },
    {
        "区域": "东京新宿",
        "类型": "乌冬",
        "推荐店/吃法": "Udon Shin：新宿热门手打乌冬，住新宿可较方便地错峰",
        "适合日期": "Day9-Day11",
        "预算感": "低-中",
        "预约/排队": "常排队，先看地图实时繁忙度",
        "点单建议": "培根天妇罗釜玉乌冬或冷乌冬",
        "P人备选": "超过可接受等待时间就回新宿站商场",
        "地图店铺": [food_pin("Udon Shin", "Udon Shin Shinjuku")],
    },
    {
        "区域": "东京新宿",
        "类型": "沾面",
        "推荐店/吃法": "风云儿：浓厚鸡白汤鱼介沾面，适合回酒店前快速吃",
        "适合日期": "Day9-Day11",
        "预算感": "低",
        "预约/排队": "翻台较快但高峰仍会排队",
        "点单建议": "特制沾面，食量小选普通份",
        "P人备选": "不想排就找车站内拉面",
        "地图店铺": [food_pin("风云儿", "Fuunji Shinjuku")],
    },
    {
        "区域": "东京涩谷",
        "类型": "平价寿司",
        "推荐店/吃法": "鱼べい道玄坂店：点单式寿司，购物途中补充体力方便",
        "适合日期": "Day9/Day10",
        "预算感": "低-中",
        "预约/排队": "高峰可能排队，现场取号",
        "点单建议": "少量多轮点，避免一次点太多",
        "P人备选": "队伍长就进涩谷商场餐饮层",
        "地图店铺": [food_pin("鱼べい涩谷道玄坂店", "Uobei Shibuya Dogenzaka")],
    },
    {
        "区域": "东京原宿",
        "类型": "柚子拉面",
        "推荐店/吃法": "AFURI原宿：柚子盐拉面清爽，适合逛街日",
        "适合日期": "Day10",
        "预算感": "低-中",
        "预约/排队": "不能预约，翻台通常较快",
        "点单建议": "柚子盐拉面，想浓一点可选柚子辣露",
        "P人备选": "表参道人多时可留到涩谷再吃",
        "地图店铺": [food_pin("AFURI原宿", "AFURI Harajuku")],
    },
    {
        "区域": "东京浅草",
        "类型": "饭团早餐",
        "推荐店/吃法": "浅草宿六：老牌现做饭团，适合回程日上午吃轻早餐",
        "适合日期": "Day11/Day12",
        "预算感": "低-中",
        "预约/排队": "座位少，营业时段需出发前确认",
        "点单建议": "两枚饭团加味噌汤",
        "P人备选": "回程日遇到排队直接跳过",
        "地图店铺": [food_pin("浅草宿六", "Onigiri Asakusa Yadoroku")],
    },
    {
        "区域": "东京浅草",
        "类型": "铁板烧",
        "推荐店/吃法": "染太郎：日式老屋中的文字烧/大阪烧体验",
        "适合日期": "Day11东京自由",
        "预算感": "中",
        "预约/排队": "用餐时间较长，不建议赶飞机当天去",
        "点单建议": "文字烧与大阪烧各点一种分食",
        "P人备选": "时间紧就吃浅草小吃",
        "地图店铺": [food_pin("染太郎", "Sometaro Asakusa")],
    },
    {
        "区域": "东京日本桥",
        "类型": "海鲜丼",
        "推荐店/吃法": "つじ半日本桥本店：海鲜珠宝丼，最后可加高汤做茶泡饭",
        "适合日期": "Day11东京自由/Day12",
        "预算感": "中",
        "预约/排队": "热门时段排队明显，不适合赶飞机前硬等",
        "点单建议": "基础梅套餐已足够丰富",
        "P人备选": "排队长就改东京站商场寿司",
        "地图店铺": [food_pin("つじ半日本桥本店", "Tsujihan Nihonbashi")],
    },
    {
        "区域": "镰仓由比滨",
        "类型": "荞麦",
        "推荐店/吃法": "松原庵：古民家荞麦与天妇罗，适合镰仓慢游",
        "适合日期": "Day11镰仓",
        "预算感": "中-高",
        "预约/排队": "建议预约；用餐节奏较慢",
        "点单建议": "荞麦加天妇罗套餐",
        "P人备选": "时间紧就小町通或江之岛吃吻仔鱼丼",
        "地图店铺": [food_pin("镰仓松原庵", "Kamakura Matsubaraan")],
    },
]


def hotel_pick(
    name: str,
    area: str,
    budget: str,
    fit: str,
    query: str,
) -> dict[str, str]:
    """保存住宿选择所需的区域、预算定位和地图目标，便于用户从当前位置直接导航。"""
    return {
        "名称": name,
        "区域": area,
        "预算感": budget,
        "适合理由": fit,
        "地图查询": query,
    }


HOTEL_RECOMMENDATIONS: dict[str, list[dict[str, str]]] = {
    "大阪段": [
        hotel_pick(
            "东横INN大阪难波西",
            "难波/元町",
            "约¥350–500/间夜",
            "独立卫浴，通常含简易早餐；步行可到南海难波站，兼顾关西机场与道顿堀。",
            "Toyoko Inn Osaka Namba Nishi",
        ),
        hotel_pick(
            "东横INN大阪难波日本桥",
            "日本桥/难波",
            "约¥350–500/间夜",
            "独立卫浴并含简易早餐，靠近黑门市场；去难波、心斋桥和机场都方便。",
            "Toyoko Inn Osaka Namba Nippombashi",
        ),
        hotel_pick(
            "KOKO HOTEL 大阪难波惠美须町",
            "惠美须町/新世界",
            "约¥300–500/间夜",
            "独立卫浴；地铁堺筋线出行方便，房价通常比难波核心区友好。",
            "KOKO HOTEL Osaka Namba Ebisucho",
        ),
        hotel_pick(
            "KOKO HOTEL 大阪梅田",
            "梅田/大阪站东侧",
            "促销约¥400–500/间夜",
            "独立卫浴；到大阪站约步行10分钟，去姬路、神户和USJ换乘顺，超预算时跳过。",
            "KOKO HOTEL Osaka Umeda",
        ),
        hotel_pick(
            "Hotel Sunplaza 2 Annex",
            "新今宫/动物园前",
            "约¥200–400/间夜",
            "经济型私人客房，部分房型使用共用浴场/卫浴；JR、南海和地铁三线可用。",
            "Hotel Sunplaza 2 Annex Osaka",
        ),
        hotel_pick(
            "Hotel Toyo Osaka",
            "动物园前/新今宫",
            "约¥150–300/间夜",
            "青旅型私人房、共用卫浴，适合单人压低预算；在意隔音和卫浴私密性则不选。",
            "Hotel Toyo Osaka",
        ),
    ],
    "京都段": [
        hotel_pick(
            "THE POCKET HOTEL 京都乌丸五条",
            "五条/乌丸线",
            "约¥250–450/间夜",
            "带锁私人房、共用卫浴；五条站步行约1分钟，适合单人住且交通很省力。",
            "THE POCKET HOTEL Kyoto Karasuma Gojo",
        ),
        hotel_pick(
            "HOTEL TAVINOS Kyoto",
            "河原町五条/清水五条",
            "约¥300–500/间夜",
            "紧凑型独立客房，步行可到清水五条站；去祇园、清水寺与京都站较均衡。",
            "HOTEL TAVINOS Kyoto",
        ),
        hotel_pick(
            "ibis Styles Kyoto Shijo",
            "四条乌丸",
            "约¥350–500/间夜",
            "独立卫浴，四条站和乌丸站步行约4分钟；位置贴合原攻略，促销价合适时优先。",
            "ibis Styles Kyoto Shijo",
        ),
        hotel_pick(
            "HOTEL M's PLUS SHIJO OMIYA",
            "四条大宫",
            "约¥300–500/间夜",
            "独立卫浴，靠近阪急与岚电；去岚山方便，坐一小段车即可到四条河原町。",
            "HOTEL M's PLUS SHIJO OMIYA",
        ),
        hotel_pick(
            "WeBase 京都",
            "四条乌丸",
            "约¥200–450/人夜",
            "青旅/简约酒店，按胶囊床位或私人房选择；四条站步行约5分钟，适合单人。",
            "WeBase Kyoto",
        ),
        hotel_pick(
            "东横INN京都四条大宫",
            "四条大宫",
            "约¥350–500/间夜",
            "独立卫浴并通常含简易早餐，适合看重稳定连锁标准和交通便利的人。",
            "Toyoko Inn Kyoto Shijo Omiya",
        ),
    ],
    "东京段": [
        hotel_pick(
            "HOTEL TAVINOS Asakusa",
            "浅草",
            "约¥350–500/间夜",
            "紧凑型独立客房，适合浅草与上野动线；去新宿约需半小时，优先保预算。",
            "HOTEL TAVINOS Asakusa",
        ),
        hotel_pick(
            "Agora Place Tokyo Asakusa",
            "田原町/浅草",
            "约¥400–500/间夜",
            "独立卫浴，田原町站步行约1分钟；银座线上野、银座和涩谷无需换乘。",
            "Agora Place Tokyo Asakusa",
        ),
        hotel_pick(
            "HOTEL MYSTAYS Asakusa",
            "本所/藏前",
            "约¥350–500/间夜",
            "独立卫浴并有投币洗衣房；位置不在最热闹街区，通常更容易守住预算。",
            "HOTEL MYSTAYS Asakusa",
        ),
        hotel_pick(
            "Ueno New Izu Hotel",
            "上野/稻荷町",
            "约¥350–500/间夜",
            "老牌经济型酒店，步行可到上野站；适合成田机场、浅草和东京站方向。",
            "Ueno New Izu Hotel Tokyo",
        ),
        hotel_pick(
            "Juyoh Hotel",
            "南千住/山谷",
            "约¥200–350/间夜",
            "经济型私人房、共用卫浴；预算优先的单人备选，夜间回程需多留意周边环境。",
            "Juyoh Hotel Tokyo",
        ),
        hotel_pick(
            "Hotel Palace Japan",
            "南千住",
            "约¥200–400/间夜",
            "带锁私人房、共用卫浴，价格通常低于核心区商务酒店；适合单人轻装入住。",
            "Hotel Palace Japan Tokyo",
        ),
        hotel_pick(
            "9h nine hours Shinjuku-North",
            "新大久保/新宿北侧",
            "约¥250–450/人夜",
            "胶囊房，2026年开业；想保留新宿动线且能接受共用卫浴时再选。",
            "9h nine hours Shinjuku-North",
        ),
    ],
    "可选温泉": [
        hotel_pick(
            "东横INN神户三宫1号店",
            "神户三宫",
            "约¥350–500/间夜",
            "独立卫浴并通常含简易早餐；住三宫后白天往返有马温泉，比住温泉旅馆省钱。",
            "Toyoko Inn Kobe Sannomiya No.1",
        ),
        hotel_pick(
            "Kobe Sannomiya Union Hotel",
            "神户三宫东侧",
            "约¥300–500/间夜",
            "独立卫浴的商务酒店，适合把神户作为有马温泉日归基地。",
            "Kobe Sannomiya Union Hotel",
        ),
        hotel_pick(
            "Kobe City Gardens Hotel",
            "JR神户站",
            "约¥250–450/间夜",
            "经济型独立客房，靠近JR神户站；适合先看港区、再搭车去有马。",
            "Kobe City Gardens Hotel",
        ),
        hotel_pick(
            "Hostel Anchorage",
            "神户站/港区",
            "约¥180–350/人夜",
            "青旅床位或简约私人房、共用卫浴；预算最优先时选择。",
            "Hostel Anchorage Kobe",
        ),
    ],
}


def read_sheet(wb, name: str) -> list[dict[str, str]]:
    ws = wb[name]
    rows = list(ws.iter_rows(values_only=True))
    headers = [str(value or "") for value in rows[0]]
    result: list[dict[str, str]] = []
    for row in rows[1:]:
        item = {}
        for header, value in zip(headers, row):
            item[header] = "" if value is None else str(value)
        if any(item.values()):
            result.append(item)
    return result


def build_data() -> dict[str, object]:
    wb = load_workbook(WORKBOOK, read_only=True, data_only=True)
    overview = read_sheet(wb, "12天总览")
    daily_food = read_sheet(wb, "每日吃法")
    food_map = read_sheet(wb, "美食推荐")
    lodging = read_sheet(wb, "住宿区域推荐")
    alternatives = read_sheet(wb, "P人备选方案")
    transport = read_sheet(wb, "交通预约")
    checklist = read_sheet(wb, "行前清单")
    daily_detail = read_sheet(wb, "关西加强每日攻略")

    for item in food_map:
        key = (item.get("区域", ""), item.get("类型", ""))
        item["地图店铺"] = FOOD_LOCATIONS.get(key, [])
    food_map.extend(EXTRA_FOOD_RECOMMENDATIONS)
    for item in lodging:
        item["住宿推荐"] = HOTEL_RECOMMENDATIONS.get(item.get("阶段", ""), [])

    foods_by_day = {item.get("日期", ""): item for item in daily_food}
    details_by_date: dict[str, list[dict[str, str]]] = {}
    for item in daily_detail:
        details_by_date.setdefault(item.get("日期", ""), []).append(item)

    days = []
    for item in overview:
        date = item.get("日期", "")
        short_date = date[5:] if len(date) >= 10 else date
        days.append(
            {
                **item,
                "短日期": short_date,
                "吃法": foods_by_day.get(short_date, {}),
                "细节": details_by_date.get(short_date, []),
                "交通分段": SEGMENT_ROUTES.get(item.get("天数", ""), []),
            }
        )

    return {
        "trip": {
            "title": "日本12天11晚",
            "subtitle": "广州出发 · 关西加强 · 东京收尾",
            "dates": "2026.09.25 - 2026.10.06",
            "route": "大阪4晚 → 京都4晚 → 东京3晚",
        },
        "days": days,
        "foodMap": food_map,
        "lodging": lodging,
        "alternatives": alternatives,
        "transport": transport,
        "checklist": checklist,
        "updatedAt": "2026-07-30",
    }


def write_site(data: dict[str, object]) -> None:
    ROOT.mkdir(exist_ok=True)
    (ROOT / "data.js").write_text(
        "window.TRIP_DATA = " + json.dumps(data, ensure_ascii=False, indent=2) + ";\n",
        encoding="utf-8",
    )
    (ROOT / "index.html").write_text(
        """<!doctype html>
<html lang="zh-CN">
  <head>
    <meta charset="utf-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover" />
    <meta name="theme-color" content="#1e4d4a" />
    <title>日本12天11晚旅行攻略</title>
    <link rel="stylesheet" href="./styles.css" />
  </head>
  <body>
    <header class="hero" id="hero">
      <div class="hero__shade"></div>
      <nav class="topbar">
        <span class="brand">日本旅行</span>
        <a class="ghost-link" href="#today">行程</a>
      </nav>
      <section class="hero__content">
        <p class="eyebrow" id="tripDates"></p>
        <h1 id="tripTitle"></h1>
        <p class="subtitle" id="tripSubtitle"></p>
        <div class="route-pill" id="tripRoute"></div>
      </section>
    </header>

    <main>
      <section class="quick-band" aria-label="旅行摘要">
        <div class="quick-stat">
          <strong>12</strong>
          <span>天</span>
        </div>
        <div class="quick-stat">
          <strong>8</strong>
          <span>晚关西</span>
        </div>
        <div class="quick-stat">
          <strong>3</strong>
          <span>晚东京</span>
        </div>
      </section>

      <div class="tabs" role="tablist" aria-label="攻略分类">
        <button class="tab is-active" data-view="days" type="button">每日</button>
        <button class="tab" data-view="food" type="button">美食</button>
        <button class="tab" data-view="lodging" type="button">住宿</button>
        <button class="tab" data-view="planb" type="button">备选</button>
        <button class="tab" data-view="prep" type="button">准备</button>
      </div>

      <section class="toolbar" id="dayToolbar">
        <label class="search">
          <span>搜索</span>
          <input id="searchInput" type="search" placeholder="京都、USJ、抹茶、富士山" />
        </label>
        <div class="chips" id="cityChips"></div>
      </section>

      <section class="content-view is-visible" id="daysView">
        <div class="section-head" id="today">
          <h2>每日行程</h2>
          <p>按当天体力调整，重点行程都留了删减口。</p>
        </div>
        <div class="day-list" id="dayList"></div>
      </section>

      <section class="content-view" id="foodView">
        <div class="section-head">
          <h2>美食推荐</h2>
          <p>按区域找吃的，点击店名可直接进入 Google Maps；排队太长就用备选。</p>
        </div>
        <div class="food-groups" id="foodGroups"></div>
      </section>

      <section class="content-view" id="lodgingView">
        <div class="section-head">
          <h2>住宿推荐</h2>
          <p>按人民币每晚约300–500元筛选；每家都能查看定位，或从你当前所在位置直接规划路线。</p>
        </div>
        <div class="simple-list" id="lodgingList"></div>
      </section>

      <section class="content-view" id="planbView">
        <div class="section-head">
          <h2>备选方案</h2>
          <p>晚点、下雨、人多、富士山看不见时用。</p>
        </div>
        <div class="simple-list" id="planbList"></div>
      </section>

      <section class="content-view" id="prepView">
        <div class="section-head">
          <h2>准备清单</h2>
          <p>预约、交通、证件和行李放在这里。</p>
        </div>
        <div class="simple-list" id="transportList"></div>
        <div class="simple-list" id="checkList"></div>
      </section>
    </main>

    <footer>
      <span id="updatedAt"></span>
      <span>行程来自本地 Excel · 交通、美食与住宿定位核对 Google Maps</span>
    </footer>

    <script src="./data.js"></script>
    <script src="./app.js"></script>
  </body>
</html>
""",
        encoding="utf-8",
    )
    (ROOT / "styles.css").write_text(
        r""":root {
  color-scheme: light;
  --ink: #172321;
  --muted: #5f6f6b;
  --paper: #fffdf8;
  --band: #f2eee5;
  --line: #ded7c8;
  --green: #1e4d4a;
  --green-2: #2f6f65;
  --coral: #c85e40;
  --gold: #b3832d;
  --shadow: 0 14px 32px rgba(30, 44, 42, 0.12);
  font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", "PingFang SC", "Microsoft YaHei", sans-serif;
}

* {
  box-sizing: border-box;
}

html {
  scroll-behavior: smooth;
}

body {
  margin: 0;
  background: var(--paper);
  color: var(--ink);
}

.hero {
  min-height: 88svh;
  position: relative;
  display: flex;
  flex-direction: column;
  justify-content: space-between;
  background:
    radial-gradient(circle at 18% 16%, rgba(255, 255, 255, 0.22), transparent 25%),
    radial-gradient(circle at 88% 22%, rgba(200, 94, 64, 0.35), transparent 26%),
    linear-gradient(135deg, #173b39 0%, #1e4d4a 42%, #b3832d 100%);
  isolation: isolate;
}

.hero__shade {
  position: absolute;
  inset: 0;
  background:
    linear-gradient(180deg, rgba(0, 0, 0, 0.1), rgba(0, 0, 0, 0.04) 38%, rgba(0, 0, 0, 0.38) 100%),
    repeating-linear-gradient(120deg, rgba(255, 255, 255, 0.08) 0 1px, transparent 1px 18px);
  z-index: -1;
}

.topbar {
  display: flex;
  align-items: center;
  justify-content: space-between;
  padding: max(18px, env(safe-area-inset-top)) 18px 0;
  color: #fff;
}

.brand {
  font-size: 14px;
  font-weight: 700;
}

.ghost-link {
  color: #fff;
  text-decoration: none;
  border: 1px solid rgba(255, 255, 255, 0.62);
  border-radius: 999px;
  padding: 8px 14px;
  font-size: 13px;
  background: rgba(0, 0, 0, 0.18);
  backdrop-filter: blur(10px);
}

.hero__content {
  color: #fff;
  padding: 0 20px 28px;
  max-width: 780px;
}

.eyebrow {
  margin: 0 0 10px;
  font-size: 13px;
  font-weight: 700;
}

h1 {
  margin: 0;
  font-size: clamp(42px, 15vw, 86px);
  line-height: 0.94;
  letter-spacing: 0;
}

.subtitle {
  max-width: 28rem;
  margin: 16px 0 18px;
  font-size: 17px;
  line-height: 1.55;
}

.route-pill {
  display: inline-flex;
  max-width: 100%;
  padding: 10px 14px;
  border-radius: 999px;
  background: rgba(255, 255, 255, 0.88);
  color: var(--green);
  font-size: 14px;
  font-weight: 800;
}

main {
  padding-bottom: 42px;
}

.quick-band {
  display: grid;
  grid-template-columns: repeat(3, 1fr);
  gap: 1px;
  background: var(--line);
  border-bottom: 1px solid var(--line);
}

.quick-stat {
  background: #faf6ed;
  padding: 16px 10px;
  text-align: center;
}

.quick-stat strong {
  display: block;
  font-size: 25px;
  color: var(--coral);
}

.quick-stat span {
  display: block;
  margin-top: 2px;
  font-size: 12px;
  color: var(--muted);
}

.tabs {
  position: sticky;
  top: 0;
  z-index: 8;
  display: flex;
  gap: 8px;
  overflow-x: auto;
  padding: 12px 14px;
  background: rgba(255, 253, 248, 0.94);
  border-bottom: 1px solid var(--line);
  backdrop-filter: blur(18px);
}

.tab {
  flex: 0 0 auto;
  border: 1px solid var(--line);
  background: #fff;
  color: var(--green);
  border-radius: 999px;
  padding: 9px 15px;
  font-weight: 800;
}

.tab.is-active {
  color: #fff;
  background: var(--green);
  border-color: var(--green);
}

.toolbar {
  padding: 14px 16px 4px;
}

.search {
  display: grid;
  gap: 6px;
  color: var(--muted);
  font-size: 12px;
  font-weight: 800;
}

.search input {
  width: 100%;
  min-height: 42px;
  border: 1px solid var(--line);
  border-radius: 8px;
  padding: 10px 12px;
  background: #fff;
  color: var(--ink);
  font: inherit;
}

.chips {
  display: flex;
  gap: 8px;
  overflow-x: auto;
  padding: 12px 0 4px;
}

.chip {
  flex: 0 0 auto;
  border: 1px solid var(--line);
  background: #fff;
  color: var(--muted);
  border-radius: 999px;
  padding: 8px 12px;
  font-size: 13px;
  font-weight: 700;
}

.chip.is-active {
  color: #fff;
  background: var(--coral);
  border-color: var(--coral);
}

.content-view {
  display: none;
}

.content-view.is-visible {
  display: block;
}

.section-head {
  padding: 22px 18px 10px;
}

.section-head h2 {
  margin: 0;
  font-size: 24px;
}

.section-head p {
  margin: 7px 0 0;
  color: var(--muted);
  line-height: 1.45;
}

.day-list,
.simple-list,
.food-groups {
  display: grid;
  gap: 12px;
  padding: 0 14px 18px;
}

.day-card,
.info-card,
.food-card {
  background: #fff;
  border: 1px solid var(--line);
  border-radius: 8px;
  box-shadow: var(--shadow);
  overflow: hidden;
}

.day-card summary {
  list-style: none;
  cursor: pointer;
  padding: 16px;
}

.day-card summary::-webkit-details-marker {
  display: none;
}

.day-top {
  display: flex;
  justify-content: space-between;
  gap: 12px;
  align-items: flex-start;
}

.day-kicker,
.tag {
  color: var(--coral);
  font-size: 12px;
  font-weight: 900;
}

.day-title {
  margin: 6px 0 0;
  font-size: 20px;
  line-height: 1.25;
}

.stay {
  color: var(--green);
  background: #e7f0ed;
  border-radius: 999px;
  padding: 7px 10px;
  font-size: 12px;
  font-weight: 800;
  white-space: nowrap;
}

.anchors {
  margin: 12px 0 0;
  color: var(--muted);
  line-height: 1.5;
}

.day-body {
  border-top: 1px solid var(--line);
  padding: 0 16px 16px;
}

.field {
  margin-top: 14px;
}

.field strong {
  display: block;
  margin-bottom: 4px;
  color: var(--green);
  font-size: 13px;
}

.field p {
  margin: 0;
  color: var(--ink);
  line-height: 1.55;
}

.detail-list {
  display: grid;
  gap: 8px;
  margin-top: 10px;
}

.route-list {
  display: grid;
  gap: 8px;
  margin-top: 8px;
}

.route-row {
  display: grid;
  grid-template-columns: minmax(0, 1fr) auto;
  gap: 4px 10px;
  align-items: center;
  padding: 10px 11px;
  border: 1px solid #dce8e4;
  border-radius: 7px;
  background: #f5faf8;
  color: var(--ink);
  text-decoration: none;
}

.route-row:active {
  background: #e7f0ed;
}

.route-option {
  display: inline-block;
  margin-right: 6px;
  color: var(--coral);
  font-size: 11px;
  font-weight: 900;
}

.route-points {
  min-width: 0;
  font-size: 14px;
  font-weight: 800;
  line-height: 1.4;
}

.route-time {
  color: var(--green);
  font-size: 13px;
  font-weight: 900;
  white-space: nowrap;
}

.route-note {
  grid-column: 1 / -1;
  color: var(--muted);
  font-size: 12px;
  line-height: 1.4;
}

.route-source {
  margin-top: 8px;
  color: var(--muted);
  font-size: 11px;
  line-height: 1.45;
}

.detail-row {
  border-left: 3px solid var(--gold);
  padding: 4px 0 4px 10px;
}

.info-card,
.food-card {
  padding: 15px;
}

.info-card h3,
.food-card h3 {
  margin: 0 0 8px;
  font-size: 17px;
  line-height: 1.35;
}

.meta {
  display: flex;
  flex-wrap: wrap;
  gap: 6px;
  margin-bottom: 10px;
}

.meta span {
  border-radius: 999px;
  padding: 5px 8px;
  background: #f4efe4;
  color: var(--muted);
  font-size: 12px;
  font-weight: 700;
}

.food-region {
  padding: 14px 4px 0;
  color: var(--green);
  font-size: 18px;
}

.food-map-links {
  display: flex;
  flex-wrap: wrap;
  gap: 7px;
  margin: 10px 0 2px;
}

.food-map-link {
  display: inline-flex;
  align-items: center;
  gap: 4px;
  max-width: 100%;
  padding: 7px 9px;
  border: 1px solid #cddfda;
  border-radius: 999px;
  background: #edf6f3;
  color: var(--green);
  font-size: 12px;
  font-weight: 850;
  line-height: 1.3;
  text-decoration: none;
}

.food-map-link:active {
  background: #dcebe7;
}

.food-map-note {
  width: 100%;
  color: var(--muted);
  font-size: 11px;
  line-height: 1.4;
}

.hotel-list {
  display: grid;
  gap: 9px;
  margin-top: 9px;
}

.hotel-pick {
  padding: 11px;
  border: 1px solid #dce8e4;
  border-radius: 7px;
  background: #f7faf9;
}

.hotel-top {
  display: flex;
  justify-content: space-between;
  gap: 8px;
  align-items: flex-start;
}

.hotel-name {
  margin: 0;
  font-size: 14px;
  line-height: 1.4;
}

.hotel-budget {
  flex: 0 0 auto;
  padding: 4px 7px;
  border-radius: 999px;
  background: #f4efe4;
  color: var(--muted);
  font-size: 11px;
  font-weight: 800;
}

.hotel-area {
  margin-top: 3px;
  color: var(--coral);
  font-size: 11px;
  font-weight: 900;
}

.hotel-reason {
  margin: 7px 0 0;
  color: var(--ink);
  font-size: 12px;
  line-height: 1.5;
}

.hotel-actions {
  display: flex;
  flex-wrap: wrap;
  gap: 7px;
  margin-top: 9px;
}

.hotel-link {
  display: inline-flex;
  align-items: center;
  padding: 7px 9px;
  border: 1px solid #cddfda;
  border-radius: 999px;
  background: #fff;
  color: var(--green);
  font-size: 12px;
  font-weight: 850;
  text-decoration: none;
}

.hotel-link--route {
  border-color: var(--green);
  background: var(--green);
  color: #fff;
}

.hotel-location-note {
  margin-top: 9px;
  color: var(--muted);
  font-size: 11px;
  line-height: 1.45;
}

footer {
  display: flex;
  justify-content: space-between;
  gap: 12px;
  padding: 24px 18px calc(24px + env(safe-area-inset-bottom));
  background: var(--green);
  color: rgba(255, 255, 255, 0.82);
  font-size: 12px;
}

@media (min-width: 760px) {
  .hero {
    min-height: 72vh;
  }

  .hero__content {
    padding-left: 7vw;
    padding-bottom: 7vh;
  }

  main {
    max-width: 980px;
    margin: 0 auto;
  }

  .quick-band {
    border-left: 1px solid var(--line);
    border-right: 1px solid var(--line);
  }

  .day-list,
  .simple-list,
  .food-groups {
    grid-template-columns: repeat(2, minmax(0, 1fr));
  }

  .day-list {
    grid-template-columns: 1fr;
  }
}
""",
        encoding="utf-8",
    )
    (ROOT / "app.js").write_text(
        r"""const data = window.TRIP_DATA;

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
""",
        encoding="utf-8",
    )


def write_standalone() -> None:
    html = (ROOT / "index.html").read_text(encoding="utf-8")
    css = (ROOT / "styles.css").read_text(encoding="utf-8")
    data_js = (ROOT / "data.js").read_text(encoding="utf-8")
    app_js = (ROOT / "app.js").read_text(encoding="utf-8")

    html = html.replace('    <link rel="stylesheet" href="./styles.css" />', f"    <style>\n{css}\n    </style>")
    html = html.replace('    <script src="./data.js"></script>\n    <script src="./app.js"></script>', f"    <script>\n{data_js}\n{app_js}\n    </script>")
    html = html.replace(
        '<link rel="preconnect" href="https://images.unsplash.com" />\n    ',
        "",
    )
    Path("日本12天11晚_手机单页攻略.html").write_text(html, encoding="utf-8")


if __name__ == "__main__":
    write_site(build_data())
    write_standalone()
    print((ROOT / "index.html").resolve())
    print(Path("日本12天11晚_手机单页攻略.html").resolve())
