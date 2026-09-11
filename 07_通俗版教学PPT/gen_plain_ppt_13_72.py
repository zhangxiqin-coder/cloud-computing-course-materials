# -*- coding: utf-8 -*-
# 生成《云计算基础》第13-72课时通俗版教学PPT
# 解析 13-72 教案 HTML，复用前12课时的幻灯片构建函数（build_* / build_lesson）。
# 面向职高学生：无手机电脑，教师投影、学生只看屏。大字、少字、生活类比。
import os, re, html, importlib.util

BASE = r"D:\WorkBuddyData\2026-08-11-07-48-22"
PLAN = os.path.join(BASE, "01_教案", "第13-72课时", "云计算基础_第13-72课时教案.html")
OUT = os.path.join(BASE, "07_通俗版教学PPT")
os.makedirs(OUT, exist_ok=True)

# ---- 导入前12课时的构建函数 ----
spec = importlib.util.spec_from_file_location(
    "gen_plain_ppt", os.path.join(OUT, "gen_plain_ppt.py"))
gp = importlib.util.module_from_spec(spec)
spec.loader.exec_module(gp)
build_lesson = gp.build_lesson
CH_COLOR = gp.CH_COLOR
CH_NAME = gp.CH_NAME

# 课时 -> 章节 映射（依据教案目录）
def ch_of(num):
    if 13 <= num <= 20: return 5
    if 21 <= num <= 24: return 6
    if 25 <= num <= 30: return 7
    if 31 <= num <= 36: return 8
    if 37 <= num <= 43: return 9
    if 44 <= num <= 48: return 10
    if 49 <= num <= 54: return 11
    if 55 <= num <= 60: return 12
    if 61 <= num <= 64: return 13
    return 14  # 65-72 综合复习与期末

# 每章一个导入钩子（一句话引出）
CHAPTER_HOOK = {
    5: "上网存东西、传照片，真的安全吗？今天聊云世界里的\"保镖\"。",
    6: "云是谁在卖、怎么卖钱的？带你看看云计算这门生意。",
    7: "你点开一个网页，数据是怎么\"跑\"到眼前的？揭秘网络。",
    8: "聊天记录、订单、成绩……全存在\"数据库\"里，它到底是什么？",
    9: "一台电脑怎么\"变\"成好多台？虚拟化的魔法今天拆开看。",
    10: "云背后的服务器大多跑 Linux，认识一下这个\"黑窗口\"系统。",
    11: "你刷的网页、点的接口，背后是 Web 服务在撑着，今天看门道。",
    12: "阿里云、腾讯云怎么用？带你看公有云平台的\"控制台\"。",
    13: "不想把数据放别人家？自己搭一个私有云，今天试试。",
    14: "全部内容串一遍，查漏补缺，准备期末大考。",
}
# 每章 3 条默认类比（抽不到教案类比时用，保证3张概念页各不相同）
CHAPTER_ANALOGIES = {
    5: ["🛡️ 云安全像小区保安+门锁，厂商管大楼，你管自己房门钥匙",
        "🔐 数据上云像把贵重物品存银行，银行管金库，你管密码",
        "👀 安全三要素像保护一封信：不被偷看(机密)、不被改(完整)、随时能读(可用)"],
    6: ["🛒 云市场像家电卖场，按需租不用买断",
        "💰 云服务像水电表，用多少付多少",
        "🏪 选云厂商像选手机套餐，比价格也比服务"],
    7: ["🚚 网络像快递系统，数据包是包裹，路由是分拣中心",
        "📮 寄信要写地址，上网要写IP，不然送不到",
        "🔌 网线/WiFi像水管，带宽就是管子粗细，越粗水流越大"],
    8: ["📒 数据库像超级大表格，查数据=翻目录找格子",
        "🗂️ 表/行/列像Excel：一张表是一张Sheet，一行是一条记录",
        "🔑 主键像身份证号，唯一定位一条记录"],
    9: ["🏘️ 虚拟化=一幢楼隔成多间，各住各的互不干扰",
        "🧊 一块大冰格分成一格格，各装各的饮料",
        "🖥️ 一台物理机\"分身\"成多台虚拟机，像一人分饰多角"],
    10: ["🖥️ Linux 像出租车的计价器后台，看不见但一直在跑",
         "⌨️ 敲命令像给服务员点单，指令就是菜名",
         "📁 目录结构像图书馆书架，一层层往下找书"],
    11: ["🍔 Web服务像餐厅前台，你点单它出餐",
         "🌐 浏览器像菜单，服务器像厨房，HTTP是点单话术",
         "🔗 一个网址像门牌号，输入就找到那户人家"],
    12: ["🏢 公有云控制台像自助机，点几下就能开一台电脑",
         "☁️ 开云服务器像租共享办公室，随时退租",
         "📦 云产品像积木，按需拼出你要的系统"],
    13: ["🏰 私有云=自家院子，钥匙全在你手里",
         "🔒 自己建机房像自己打井，水干净但要自己维护",
         "🏭 企业私有云像厂内专线，外人进不来"],
    14: ["🧩 复习像拼图，把零散知识点拼成一张大图",
         "🗺️ 总复习像画地图，标出哪座山(重点)哪条河(难点)",
         "✅ 期末像大考前的体检，查漏补缺才放心"],
}

DEMO_KEYWORDS = ["演示", "体验", "初体验", "安装", "操作", "实践", "实验", "搭建", "部署", "配置", "上机"]

def strip_tags(s):
    s = re.sub(r"<[^>]+>", " ", s)
    s = html.unescape(s)
    return re.sub(r"\s+", " ", s).strip()

def shorten(t, n=18):
    t = t.strip()
    for sep in ["（", "(", "：", ":", "，", ",", "。", "；", ";", "——"]:
        i = t.find(sep)
        if 0 < i <= n:
            return t[:i].strip()
    return t[:n].strip()

_NOISE = ["环节", "讲授", "导入", "小结", "视频", "活动", "时间", "类型", "内容",
          "纸面", "目标", "重点", "复习", "布置", "考试", "00-", "10-", "20-", "30-"]

def extract_analogies(text):
    """从去标签纯文本中抽取干净的『类比/比喻』短句，避免带入时间表等噪声。"""
    res = []
    # 1) 类比/比喻 后紧跟冒号的内容
    for m in re.finditer(r"(?:类比|比喻)[：:]\s*([^。；，,\n]{2,45})", text):
        res.append("🎯 " + m.group(1).strip())
    # 2) 「用X类比」的 X（短主语，不含空格/标点）
    for m in re.finditer(r"用([^\s，,。；：:\n]{1,6})类比", text):
        res.append("🎯 用%s类比" % m.group(1).strip())
    clean = []
    for a in res:
        a = re.sub(r"\d{1,2}-\d{1,2}", "", a)   # 去掉 25-40 / 00-10 等时间片段
        a = re.sub(r"\s+", "", a)               # 去空格（时间戳常夹空格）
        for nz in _NOISE:
            i = a.find(nz)
            if i > 0:
                a = a[:i]
        a = a.strip(" 、，,。；;：:").strip()
        if 3 <= len(a) <= 50 and ("类比" in a or "比喻" in a or a.startswith("🎯")):
            clean.append(a)
    seen, out = set(), []
    for a in clean:
        if a not in seen:
            seen.add(a)
            out.append(a)
    return out

def parse_plan():
    src = open(PLAN, encoding="utf-8").read()
    # 按课时分块
    pieces = re.split(r'(?=<div class="lesson">)', src)
    blocks = [b for b in pieces if '<div class="lesson">' in b]
    lessons = []
    for b in blocks:
        m = re.search(r"<span class=\"num\">第(\d+)课时</span>(.*?)</div>", b, re.S)
        if not m:
            continue
        num = int(m.group(1))
        title = strip_tags(m.group(2))
        ch = ch_of(num)

        # 教学目标
        mg = re.search(r"教学目标：</b>(.*?)</div>", b, re.S)
        goal = strip_tags(mg.group(1)) if mg else ""
        goal_list = [g.strip() for g in re.split(r"[；;]", goal) if g.strip()][:3]
        if not goal_list:
            goal_list = [goal]

        # 内容要点
        kp = []
        mk = re.search(r"<ul class=\"content-list\">(.*?)</ul>", b, re.S)
        if mk:
            for li in re.findall(r"<li>(.*?)</li>", mk.group(1), re.S):
                t = strip_tags(li)
                if t:
                    kp.append(t)

        # 视频
        vids = re.findall(
            r"href=\"(https://www\.bilibili\.com/video/[^\"]+)\"[^>]*class=\"video-link\">([^<]*)</a>", b)
        bvs, names = [], []
        for url, name in vids:
            bv = re.search(r"/(BV[0-9A-Za-z]+|av\d+)", url)
            bvs.append(bv.group(1) if bv else url)
            names.append(name.strip())
        video_bv = " / ".join(bvs) if bvs else "（教案未列明）"
        video_name = " / ".join(names) if names else "（见教案视频链接）"

        # 活动：优先 act-box，否则 活动 行
        act = ""
        ma = re.search(r"<div class=\"act-box\">(.*?)</div>", b, re.S)
        if ma:
            act = strip_tags(re.sub(r"<b>.*?</b>", "", ma.group(1), flags=re.S))
        if not act:
            mr = re.search(r'td-type type-act">[^<]*</td><td>(.*?)</td>', b, re.S)
            if mr:
                act = strip_tags(mr.group(1))
        if not act:
            mp = re.search(r"纸面练习：</b>(.*?)</td>", b, re.S)
            if mp:
                act = strip_tags(mp.group(1))
        if not act:
            act = "纸上小结：写出本节课的3个关键词，并各写1句人话解释。"

        # 课后作业
        hw = ""
        mh = re.search(r"<div class=\"hw-box\">(.*?)</div>", b, re.S)
        if mh:
            hw = strip_tags(re.sub(r"<b>.*?</b>", "", mh.group(1), flags=re.S))
        if not hw:
            hw = "回顾本节课3个重点，说给家人或同桌听。"

        # 类比抽取（从去标签纯文本里抓含 类比/比喻 的片段）
        text = strip_tags(b)
        analogies = extract_analogies(text)

        # 演示课判定
        demo = any(k in title for k in DEMO_KEYWORDS) or ('type-demo' in b)

        lessons.append(dict(num=num, ch=ch, title=title, goal=goal_list,
                            kp=kp, video_bv=video_bv, video_name=video_name,
                            activity=act, homework=hw, analogies=analogies, demo=demo))
    return lessons

def build_lesson_data(L):
    ch = L["ch"]
    # 概念要点：去掉纯类比行
    concept_kps = [k for k in L["kp"]
                   if not k.startswith("生活类比") and "类比" not in k[:4] and "比喻" not in k[:4]]
    if not concept_kps:
        concept_kps = L["kp"][:] or ["（本节课以观看与体验为主）"]
    concepts = []
    pool = L["analogies"] or CHAPTER_ANALOGIES[ch][:]
    for i, k in enumerate(concept_kps[:3]):
        h = shorten(k, 18)
        a = pool[i % len(pool)]
        concepts.append({"h": h, "p": k, "a": a})
    while len(concepts) < 3:
        concepts.append({"h": "记住这一条", "p": concept_kps[len(concepts) % len(concept_kps)],
                         "a": pool[len(concepts) % len(pool)]})

    heads = [shorten(k, 12) for k in concept_kps[:3]]
    q1 = "用自己的话，说说什么是「%s」？" % heads[0]
    q2 = ("「%s」和「%s」有什么联系？" % (heads[0], heads[1])) if len(heads) > 1 \
        else "为什么「%s」对学云计算很重要？" % heads[0]
    remember = [shorten(k, 16) for k in concept_kps[:3]]
    while len(remember) < 3:
        remember.append("回顾本节课3个重点")

    intro = CHAPTER_HOOK[ch]
    intro_a = (L["analogies"][0] if L["analogies"] else CHAPTER_ANALOGIES[ch][0])
    vwatch = "看视频时注意：%s 是怎么被老师讲清楚的，记1个关键例子。" % heads[0]

    # 课时13 加密入门：插入"凯撒密码"现场演示页（紧跟"机密性"之后，作为具体例子）
    if L["num"] == 13:
        caesar = {
            "h": "凯撒密码：字母后移3格",
            "p": "加密=把信息藏起来，只有拿钥匙的人能看懂。最古老的玩法叫凯撒密码：把每个字母往后数3个（A→D、B→E…），HELLO 就变成 KHOOR。『后移3格』就是钥匙，往前移回来就能解密。",
            "a": "🔑 类比：全班约定作业本上的字要比真实意思错后3格，只有知道规矩的人能翻译——这就是『密钥=3』。",
        }
        concepts.insert(1, caesar)

    return {
        "num": L["num"], "ch": ch, "title": L["title"], "demo": L["demo"],
        "goal": L["goal"],
        "intro": intro,
        "intro_a": intro_a,
        "concepts": concepts,
        "questions": [q1, q2],
        "video_name": L["video_name"], "video_bv": L["video_bv"],
        "video_watch": vwatch,
        "local": "（在线播放，无本地文件，需联网）",
        "activity": L["activity"],
        "remember": remember,
        "homework": L["homework"],
    }

if __name__ == "__main__":
    raw = parse_plan()
    print("解析到课时数：", len(raw))
    lessons = [build_lesson_data(L) for L in raw]
    # 校验
    assert len(lessons) == 60, "课时数应为60，实为%d" % len(lessons)
    nums = [L["num"] for L in lessons]
    assert nums == list(range(13, 73)), "课时编号不连续"
    for L in lessons:
        assert len(L["concepts"]) in (3, 4), L["title"]
        assert len(L["remember"]) == 3
        assert len(L["questions"]) == 2
    print("数据校验通过。开始生成 60 个通俗版PPT ...")
    for L in lessons:
        path, n = build_lesson(L)
        print("  ✓ 课时%02d %s | %d页 | %s" % (L["num"], L["title"], n, os.path.basename(path)))
    print("全部完成，输出目录：", OUT)
