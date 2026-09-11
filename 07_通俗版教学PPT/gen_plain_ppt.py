# -*- coding: utf-8 -*-
# 生成《云计算基础》前12课时通俗版教学PPT
# 面向职高学生：无手机电脑，教师投影、学生只看屏。大字、少字、生活类比、图解、视频/纸上活动提示。
import os
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn

OUT = r"D:\WorkBuddyData\2026-08-11-07-48-22\07_通俗版教学PPT"
os.makedirs(OUT, exist_ok=True)

FONT = "Microsoft YaHei"
# 章节配色
CH_COLOR = {
    1: RGBColor(0x15, 0x65, 0xC0),  # 简介 蓝
    2: RGBColor(0x2E, 0x7D, 0x32),  # 服务 绿
    3: RGBColor(0xE6, 0x51, 0x00),  # 部署 橙
    4: RGBColor(0x6A, 0x1B, 0x9A),  # 特点 紫
    5: RGBColor(0xC0, 0x39, 0x2B),  # 安全 红
    6: RGBColor(0x16, 0xA0, 0x85),  # 市场 青
    7: RGBColor(0x29, 0x80, 0xB9),  # 网络 蓝
    8: RGBColor(0x8E, 0x44, 0xAD),  # 数据库 紫
    9: RGBColor(0x27, 0xAE, 0x60),  # 虚拟化 绿
    10: RGBColor(0x2C, 0x3E, 0x50),  # Linux 深灰蓝
    11: RGBColor(0xC9, 0x7A, 0x00),  # Web 琥珀
    12: RGBColor(0x0E, 0x7C, 0x7B),  # 公有云 深青
    13: RGBColor(0x6C, 0x34, 0x83),  # 私有云 紫红
    14: RGBColor(0xB0, 0x3A, 0x2E),  # 综合复习 砖红
}
# 标题背景统一蓝色（用户要求：所有课时标题色带保持一致）
TITLE_BLUE = RGBColor(0x15, 0x65, 0xC0)
CH_NAME = {
    1: "云计算的简介", 2: "云计算的服务", 3: "云计算的部署", 4: "云计算的特点",
    5: "云计算安全", 6: "云计算市场", 7: "计算机网络", 8: "数据库基础",
    9: "虚拟化基础", 10: "Linux基础", 11: "Web服务", 12: "公有云平台",
    13: "私有云平台", 14: "综合复习与期末",
}
INK = RGBColor(0x21, 0x21, 0x21)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
CARD_YELLOW = RGBColor(0xFF, 0xF8, 0xE1)
CARD_BLUE = RGBColor(0xE3, 0xF2, 0xFD)
CARD_GREEN = RGBColor(0xE8, 0xF5, 0xE9)
CARD_GRAY = RGBColor(0xF2, 0xF2, 0xF2)
BORDER_ORANGE = RGBColor(0xE6, 0x51, 0x00)


def set_run(run, text, size=24, bold=False, color=INK, font=FONT):
    run.text = text
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.name = font
    rPr = run._r.get_or_add_rPr()
    rPr.set(qn('a:ea'), font)
    rPr.set(qn('a:cs'), font)
    run.font.color.rgb = color


def add_rect(slide, left, top, width, height, fill=None, line=None, line_w=1.0, shape=MSO_SHAPE.RECTANGLE):
    sp = slide.shapes.add_shape(shape, Inches(left), Inches(top), Inches(width), Inches(height))
    if fill is None:
        sp.fill.background()
    else:
        sp.fill.solid()
        sp.fill.fore_color.rgb = fill
    if line is None:
        sp.line.fill.background()
    else:
        sp.line.color.rgb = line
        sp.line.width = Pt(line_w)
    sp.shadow.inherit = False
    return sp


def add_band(slide, ch, num, sub):
    """顶部章节色带（统一蓝色）"""
    color = TITLE_BLUE
    add_rect(slide, 0, 0, 13.333, 1.15, fill=color)
    # 左侧课时号
    tb = slide.shapes.add_textbox(Inches(0.4), Inches(0.12), Inches(3.2), Inches(0.95))
    tf = tb.text_frame; tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]; set_run(p.add_run(), "课时 %02d" % num, size=34, bold=True, color=WHITE)
    # 右侧章节名
    tb2 = slide.shapes.add_textbox(Inches(3.6), Inches(0.12), Inches(9.3), Inches(0.95))
    tf2 = tb2.text_frame; tf2.vertical_anchor = MSO_ANCHOR.MIDDLE; tf2.word_wrap = True
    p2 = tf2.paragraphs[0]; p2.alignment = PP_ALIGN.RIGHT
    set_run(p2.add_run(), "第%d章 · %s" % (ch, CH_NAME[ch]), size=20, bold=True, color=WHITE)


def add_title(slide, text, color=INK, size=30, top=1.35, left=0.5, width=12.3):
    tb = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(0.9))
    tf = tb.text_frame; tf.vertical_anchor = MSO_ANCHOR.MIDDLE; tf.word_wrap = True
    p = tf.paragraphs[0]; set_run(p.add_run(), text, size=size, bold=True, color=color)


def add_lines(slide, left, top, width, height, lines, anchor=MSO_ANCHOR.TOP, gap=8, space_after=None):
    tb = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    tf = tb.text_frame; tf.word_wrap = True; tf.vertical_anchor = anchor
    first = True
    for ln in lines:
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        if ln.get('align'): p.alignment = ln['align']
        sa = ln.get('space_after') if ln.get('space_after') is not None else space_after
        if sa is not None: p.space_after = Pt(sa)
        if ln.get('space_before') is not None: p.space_before = Pt(ln['space_before'])
        set_run(p.add_run(), ln['text'], size=ln.get('size', 24),
                bold=ln.get('bold', False), color=ln.get('color', INK))
    return tb


def add_card(slide, left, top, width, height, fill, border=None):
    return add_rect(slide, left, top, width, height, fill=fill, line=border, line_w=1.25,
                    shape=MSO_SHAPE.ROUNDED_RECTANGLE)


def new_slide(prs):
    return prs.slides.add_slide(prs.slide_layouts[6])


# ============================ 课时数据（通俗易懂版） ============================
lessons = [
{
 'num':1,'ch':1,'title':'什么是云计算','demo':False,
 'goal':['能用自己的话说出"云计算是什么"','能举出2个生活里的云例子'],
 'intro':'你手机里的照片存在哪？百度网盘、抖音的视频是谁在帮你处理？\n这些"看不见的电脑"，就是云计算。',
 'intro_a':'🏠 类比：用电不用自己买发电机，用云上的算力也不用自己买服务器。',
 'concepts':[
   {'h':'云计算 = 把"电脑的力气"放网上','p':'谁要用，谁去租。你出钱，厂商出机器和网络。','a':'💡 像自来水：拧开龙头就有水，不用自家打井。'},
   {'h':'你的手机/电脑是"窗户"','p':'真正干活、存东西的是云。你只负责看和点。','a':'🪟 手机是遥控器，云才是电视台。'},
   {'h':'先混个脸熟：NIST 三五一四','p':'3种服务、5个特征、4种部署。今天先记名字，后面慢慢拆。','a':'🔢 先记住"一二三"，细节以后填。'},
 ],
 'questions':['你今天用过的"云"有哪些？（口头说）','如果断网，你手机里哪些功能会"罢工"？'],
 'video_name':'《什么是云计算?》+《云计算到底是什么?》','video_bv':'BV1RF411C7qu / BV1yh41127T2',
 'video_watch':'注意视频里说的"按需使用""不用自己买服务器"这两句话。',
 'local':'课时01_什么是云计算.mp4',
 'activity':'纸上任务：列出"我身边的云"——至少写5个（微信 / 网盘 / 抖音 / 邮箱 / 导航…）',
 'remember':['云 = 把计算放网上，按需租用','手机是窗户，云是大脑和仓库','NIST 三五一四 先记住名字'],
 'homework':'回家找出家里3个"用云"的东西，明天口头分享。',
},
{
 'num':2,'ch':1,'title':'发展历史与5种计算模式','demo':False,
 'goal':['能说出"从大型机到云"的经历','记住5种计算模式的名字和顺序'],
 'intro':'计算机也在"减肥"：最早占满一整间房，现在装进了你的口袋。\n云计算，是这场减肥的"终点站"。',
 'intro_a':'📉 类比：从"大块头"到"小个子"， computing 越来越轻、越来越省事。',
 'concepts':[
   {'h':'5种计算模式','p':'大型机 → 个人电脑 → 网络计算 → 网格计算 → 云计算。','a':'🚶 像走路变坐车再变坐高铁，一步比一步快。'},
   {'h':'核心理念：像水电一样公用','p':'以前"买软件装自己电脑"，现在"上网直接用"。','a':'⚡ 云是继水、电、气之后的第四种"公用设施"。'},
   {'h':'为什么叫"云"','p':'图里画网络常画成一朵云，代表"看不见、不用管在哪"。','a':'☁️ 你只管用电，不用管发电厂在哪。'},
 ],
 'questions':['你家里还"自己发电"吗？——那为什么计算还要"自己买服务器"？','"公用设施"是什么意思？'],
 'video_name':'《云计算的发展与应用》+《云计算演变之路》','video_bv':'BV1bt411u72Z / BV1Uc411h7oX',
 'video_watch':'看视频里计算机是怎么一步步"变小、变便宜、变公用"的。',
 'local':'课时02_云计算的发展与应用-鲜枣课堂.mp4',
 'activity':'纸上游戏：把5种模式按时间顺序排一排（撕纸条或画线连箭头）。',
 'remember':['云是计算模式演进的终点','核心理念：像水电一样公用','5种模式顺序要记牢'],
 'homework':'回家画一条"计算模式时间轴"，标出5个阶段。',
},
{
 'num':3,'ch':1,'title':'演示课：云端Office协作体验','demo':True,
 'goal':['看懂"多人同时编辑一个文档"就是云协作','知道云协作比"用U盘拷来拷去"好在哪'],
 'intro':'以前传文件靠U盘来回拷，改一版发一版。\n现在一个链接，大家同时改，自动保存。',
 'intro_a':'🔄 类比：以前是"传纸条"，现在是"同桌一起写同一张纸"。',
 'concepts':[
   {'h':'云端Office是什么','p':'腾讯文档 / 金山文档：文档存在云上，打开网页就能写。','a':'📄 像一块"公共黑板"，谁都能上去写。'},
   {'h':'你看到的"别人光标在动"','p':'那就是云在实时把每个人的改动同步给大家。','a':'👀 像几个人同时在一张纸上写字，互相看得见。'},
   {'h':'协作 = 不用拷来拷去','p':'大家看的是同一份，改完自动存，不怕丢。','a':'🚫 跟U盘说再见。'},
 ],
 'questions':['传文件用U盘，最容易出什么麻烦？（丢、乱、版本多）','云协作怎么解决这些麻烦？'],
 'video_name':'《金山文档多人同时编辑》+《腾讯文档保姆级入门》','video_bv':'BV1DD4y1w7PD / BV1X44y1k7R1',
 'video_watch':'看视频里"多人同时编辑""自动保存"的画面，注意观察别人光标。',
 'local':'课时03_金山文档多人同时编辑.mp4',
 'activity':'纸上模拟：4人一组，用箭头画出"传统传文件"和"云协作"的区别。',
 'remember':['云协作 = 同一份、同时改','自动保存，不怕丢','告别U盘拷文件'],
 'homework':'用微信给家人发一条消息，说一个"云"的例子。',
},
{
 'num':4,'ch':2,'title':'SaaS 软件即服务','demo':False,
 'goal':['能解释 SaaS 是什么','随手举出3个 SaaS 例子'],
 'intro':'你不用装软件，打开网页或App就能用——这，就是 SaaS。',
 'intro_a':'🍽️ 类比：去餐厅吃饭(SaaS) vs 买菜回家自己做(自己装软件)。',
 'concepts':[
   {'h':'SaaS = 软件即服务','p':'厂商把软件放云上，你按月或免费直接用。例：微信、钉钉、网易邮箱、百度网盘。','a':'📦 现成的，打开就用。'},
   {'h':'好处：不装、不维护','p':'换个手机登录还能用，更新由厂商管。','a':'🔧 你只管用，修理工是厂商。'},
   {'h':'三兄弟里它最"省心"','p':'SaaS 是 IaaS / PaaS / SaaS 里，用户要操心最少的一个。','a':'🛋️ 类比：SaaS = 拎包入住。'},
 ],
 'questions':['你手机里哪些App是"打开就用、不用装插件"的？','为什么小公司爱用 SaaS？'],
 'video_name':'《用人话说 IaaS-PaaS-SaaS》','video_bv':'BV1hf4y1W7yT',
 'video_watch':'视频用"种水果"比喻三者，重点看 SaaS 那段。',
 'local':'课时04_用人话说IaaS-PaaS-SaaS.mp4',
 'activity':'纸上分类：把"微信 / 美图秀秀网页版 / 网易云音乐"圈出哪些是 SaaS。',
 'remember':['SaaS = 直接用现成软件','不用装、不用维护','微信、钉钉都是 SaaS'],
 'homework':'列出你手机里3个 SaaS 应用。',
},
{
 'num':5,'ch':2,'title':'PaaS 与 IaaS','demo':False,
 'goal':['分清 PaaS 和 IaaS','知道各自"给开发者什么"'],
 'intro':'开网店：你要"租个装修好的店"(PaaS)，还是"租毛坯房自己搞"(IaaS)？',
 'intro_a':'🏠 类比：毛坯房=IaaS，精装房=PaaS，拎包入住=SaaS。',
 'concepts':[
   {'h':'IaaS = 基础设施即服务','p':'给你"毛坯房"（服务器/网络/存储），自己装系统、装软件。例：阿里云ECS、腾讯云CVM。','a':'🧱 地皮和墙给了，装修你自己来。'},
   {'h':'PaaS = 平台即服务','p':'给你"精装房+工具台"（运行环境都备好），你只管写代码放上去。例：云数据库、开放平台。','a':'🛠️ 连工具都备齐，直接上手干。'},
   {'h':'一张图记三兄弟','p':'从上到下：IaaS(毛坯) < PaaS(精装) < SaaS(拎包)。越往下，厂商替你管得越多。','a':'📊 厂商管得越多，你越省心。'},
 ],
 'questions':['想快速上线一个小网站，选 IaaS 还是 PaaS 更省力？','为什么说"越往下厂商管越多"？'],
 'video_name':'《用人话说 IaaS-PaaS-SaaS》+ 天翼云案例','video_bv':'BV1hf4y1W7yT / BV1Mo4y1R7Sb',
 'video_watch':'看天翼云那段，注意它属于哪一类云服务。',
 'local':'课时05_中国电信天翼云.mp4',
 'activity':'排排坐：把 IaaS / PaaS / SaaS 按"厂商管得多 → 少"排队。',
 'remember':['IaaS = 毛坯服务器','PaaS = 带环境的平台','SaaS = 现成软件；越下厂商管越多'],
 'homework':'用"买房比喻"给家人讲清三者区别。',
},
{
 'num':6,'ch':2,'title':'演示课：Gitee 代码托管体验','demo':True,
 'goal':['看懂"代码也存在云上"','明白"版本管理"有什么用'],
 'intro':'写好的作业怕丢？代码也一样——放云上最保险，还能多人一起写。',
 'intro_a':'🗄️ 类比：Gitee 像是代码的"网盘 + 保险箱"。',
 'concepts':[
   {'h':'Gitee / GitHub 是什么','p':'代码的"网盘+保险箱"，还能多人一起改同一份。','a':'📚 像班级共享笔记本，谁都能补两笔。'},
   {'h':'版本管理','p':'每次保存都留记录，改错了能退回上一版。','a':'⏪ 像游戏存档，打输了读档重来。'},
   {'h':'它是云的典型用法','p':'开发者把代码"托管"在云上，不占自己电脑。','a':'☁️ 这就是前面说的 SaaS / PaaS 落地。'},
 ],
 'questions':['如果你的作业只存在桌面，电脑坏了会怎样？','"能退回上一版"为什么重要？'],
 'video_name':'《使用 Gitee 上传代码》','video_bv':'BV1Mh4y1Y73m',
 'video_watch':'看老师怎么把代码"推"到云上，注意版本号的变化。',
 'local':'课时06_使用gitee上传代码.mp4',
 'activity':'纸上画"代码保险箱"：写出 v1 / v2 / v3，画退回箭头。',
 'remember':['代码也存云上','有版本记录能回退','多人协作不冲突'],
 'homework':'注册一个网盘，体验一次"存到云上"。',
},
{
 'num':7,'ch':3,'title':'公有云与社区云','demo':False,
 'goal':['说出公有云的特点','知道谁在用、适合谁'],
 'intro':'自来水厂——谁都能接，按吨付钱。这，就是公有云的思路。',
 'intro_a':'🚰 类比：公有云 = 公共自来水，大家共用、按量付费。',
 'concepts':[
   {'h':'公有云 = 厂商建好，谁都能租','p':'例：阿里云、腾讯云、天翼云。便宜、随开随用，但数据在别人家。','a':'🏊 像公共游泳池，谁买票谁游。'},
   {'h':'优点 vs 缺点','p':'优点：便宜、不用养运维、随开随用。缺点：数据不在自己手里。','a':'⚖️ 省心省钱，但要信任厂商。'},
   {'h':'社区云 = 几家合建共用','p':'几所学校 / 几家单位合建一个云，省钱又合规。','a':'🤝 像小区共享健身房。'}],
 'questions':['"学校官网放阿里云"属于哪种云？','小创业公司为什么喜欢公有云？'],
 'video_name':'《阿里云超级数据中心》+《走进微软数据中心》','video_bv':'BV1kA411N7e5 / BV1cW411t7dk',
 'video_watch':'看数据中心长什么样，理解"成千上万台服务器"就是公有云的家。',
 'local':'课时07_阿里云超级数据中心.mp4',
 'activity':'纸上判断：写出"学校官网 / 小公司官网 / 三校联考系统"各适合哪种云。',
 'remember':['公有云 = 公共租用','便宜、随开随用','社区云 = 几家共用'],
 'homework':'查一下你家宽带的"云"是哪家服务商。',
},
{
 'num':8,'ch':3,'title':'私有云与混合云','demo':False,
 'goal':['分清私有云、混合云','知道各自适合什么场景'],
 'intro':'自家打井 vs 自来水——自己家有井，就是"私有云"。',
 'intro_a':'⛲ 类比：私有云 = 自家水井，混合云 = 井 + 自来水一起用。',
 'concepts':[
   {'h':'私有云 = 一家自己建自己用','p':'银行、政府爱用。数据不出门，安全，但要自己养机器。','a':'🔒 自家保险箱，钥匙自己拿。'},
   {'h':'混合云 = 公有 + 私有拼着用','p':'平时用公有云省钱，敏感数据走私有云保安全。','a':'🏋️ 平时去健身房(公有)，贵重东西锁家里(私有)。'},
   {'h':'四种部署一句话','p':'公有(租用) / 私有(自建) / 混合(拼接) / 社区(几家合建)。','a':'🗺️ 先记四个名字，再看场景。'}],
 'questions':['银行为什么不敢把客户数据放公有云？','"双11"大促时，混合云怎么帮你？'],
 'video_name':'《淄博燃气智慧云数据中心》+《Google Cloud TPU Data Center》','video_bv':'BV1F14y1677t / BV14u4y1C7TG',
 'video_watch':'看"智慧云"怎么把私有云用在企业里。',
 'local':'课时08_淄博燃气智慧云数据中心.mp4',
 'activity':'纸上连线：把"银行 / 小创业公司 / 学校"匹配到 私有 / 公有 / 混合。',
 'remember':['私有云 = 自家建自己用','混合云 = 公私搭配','敏感数据走私有'],
 'homework':'画一张"四种部署"小表格（名字+一句话+例子）。',
},
{
 'num':9,'ch':3,'title':'演示课：云服务器初体验','demo':True,
 'goal':['看懂"在云上开一台电脑"全流程','记住开云服务器的几个步骤'],
 'intro':'想要一台电脑又不想搬主机？云上开一台，关机还不计费。',
 'intro_a':'💻 类比：云服务器 = 云上的"虚拟电脑"，随叫随到。',
 'concepts':[
   {'h':'云服务器 ECS 是什么','p':'云上的一台虚拟电脑：选配置 → 开机 → 远程登录 → 装网站。','a':'🖥️ 像租了台永远不关机的电脑。'},
   {'h':'弹性：人多就加，人少就关','p':'双11人多临时加一台，平时关掉省钱。','a':'🎈 像橡皮筋，需要就拉长。'},
   {'h':'演示：从开机到上线','p':'控制台开ECS → 装环境 → 部署一个简单网站。','a':'🚀 几分钟，网站就上线了。'}],
 'questions':['为什么"关机不计费"对省钱很重要？','开云服务器和买台电脑，哪个更灵活？'],
 'video_name':'《阿里云 ECS 新手搭建网站》+《网站上线部署教程》','video_bv':'av928620141 / BV18a4y1W7e9',
 'video_watch':'跟着视频看"选配置→开机→部署"五步。',
 'local':'课时09_阿里云服务器ECS新手搭建网站.mp4',
 'activity':'纸上填空：开ECS五步走——选___ → 选___ → ___ → ___ → 部署。',
 'remember':['云服务器 = 云上虚拟电脑','按需开关，省钱','五步上线一个网站'],
 'homework':'在纸上写一遍"开云服务器"的步骤。',
},
{
 'num':10,'ch':4,'title':'本质与三大特征','demo':False,
 'goal':['记住云计算3大基础特征','能用大白话解释每个特征'],
 'intro':'云计算不是新魔法，是3个"老想法"凑到了一起。',
 'intro_a':'🧩 类比：三个零件拼起来，就成了"云"。',
 'concepts':[
   {'h':'特征1：按需自助','p':'像自动售货机，自己点、自己拿，不用找管理员。','a':'🥤 扫码付款，饮料自己掉出来。'},
   {'h':'特征2：泛在网络访问','p':'有网就能用，手机、平板、电脑都行。','a':'📶 不像U盘，非得插这台电脑。'},
   {'h':'特征3：资源池化','p':'大家共用一个大资源池，谁用分谁一块，互不看见。','a':'🏊 像公共泳池，每人游各人的。'},
 ],
 'questions':['"自动售货机"对应哪个特征？','为什么"资源池化"对厂商省钱？'],
 'video_name':'《云计算有哪些基本特点》','video_bv':'BV1Zf4y1X7ih',
 'video_watch':'视频直接讲5大特征，重点看前3个。',
 'local':'课时10_云计算有哪些基本特点.mp4',
 'activity':'纸上配对：把"自动售货机 / 共用泳池 / 有网就用"连到3个特征。',
 'remember':['自助 = 像售货机','上网就能用','资源像大池子分着用'],
 'homework':'用3句话给同桌讲清3个特征。',
},
{
 'num':11,'ch':4,'title':'快速弹性与可计量服务','demo':False,
 'goal':['记住剩下2个特征','能说清"弹性"是什么意思'],
 'intro':'双11网站不卡，是因为云能"临时扩容"。这就是"弹性"。',
 'intro_a':'🎈 类比：弹性 = 像橡皮筋，需要就拉长，不需要就收回。',
 'concepts':[
   {'h':'特征4：快速弹性','p':'需要时就多给资源，不需要就收回，伸缩自如。','a':'🌊 像洪水来了加堤坝，退了就撤。'},
   {'h':'特征5：可计量服务','p':'用了多少，按量算钱，像水电表一样清楚。','a':'💡 像电表，用几度交几度钱。'},
   {'h':'5个特征全齐了','p':'按需自助 + 泛在访问 + 资源池化 + 快速弹性 + 可计量。','a':'✅ 这就是 NIST 说的"5个基本特征"。'},
 ],
 'questions':['"双11临时扩容"用的是哪个特征？','"按量付费"对应哪个特征？'],
 'video_name':'《云计算的特点优势及应用领域》+《云计算关键特点》','video_bv':'BV1Su4y1G7sa / BV1Ae4y1Y7Hv',
 'video_watch':'看视频把5个特征串起来记。',
 'local':'课时11_什么是云计算六分钟了解.mp4',
 'activity':'纸上画"橡皮筋"：标出"高峰拉长 / 平时缩短"。',
 'remember':['弹性 = 像橡皮筋伸缩','计量 = 像水电表','5个特征全齐了'],
 'homework':'把这5个特征背给家长听。',
},
{
 'num':12,'ch':4,'title':'演示课：VMware 虚拟机','demo':True,
 'goal':['看懂"一台电脑里跑多台电脑"（虚拟化）'],
 'intro':'一个房间隔成几间出租——这就是"虚拟化"。',
 'intro_a':'🏘️ 类比：虚拟化 = 一幢楼隔成多间，各住各的。',
 'concepts':[
   {'h':'虚拟化 = 一机变多机','p':'用软件把一台物理电脑"切"成多台虚拟电脑，各自独立。','a':'🧊 一块大冰格，分成一格格。'},
   {'h':'它是云计算的"地基"','p':'没有虚拟化，就没有今天又弹性又便宜的云。','a':'🏗️ 楼的地基，看不见但最重要。'},
   {'h':'VMware 是常用工具','p':'装好 VMware，就能在一台电脑里开好几台虚拟机。','a':'🛠️ 一把好用的"隔间"工具。'}],
 'questions':['"一台电脑跑多台电脑"有什么好处？','为什么说虚拟化是云的地基？'],
 'video_name':'《VMware 虚拟机安装教程》+《VMware 虚拟化入门到精通》','video_bv':'BV1nQ4y1c7AW / BV14a411w7D2',
 'video_watch':'看虚拟机怎么被"造"出来、怎么同时跑多个系统。',
 'local':'课时12_VMware虚拟机安装教程.mp4',
 'activity':'纸上画图：一幢楼隔成多间，标出"虚拟化 = 一机变多机"。',
 'remember':['虚拟化 = 一机变多机','是云的地基','VMware 是工具'],
 'homework':'整理前11课笔记，准备期末总复习。',
},
]


def build_cover(prs, L):
    s = new_slide(prs)
    add_band(s, L['ch'], L['num'], L['title'])
    # 标题区
    add_rect(s, 0.5, 2.2, 12.33, 1.5, fill=CARD_GRAY)
    add_lines(s, 0.8, 2.35, 11.7, 1.2,
              [{'text': ('演示课 · ' if L['demo'] else '') + L['title'], 'size': 34, 'bold': True, 'color': TITLE_BLUE}],
              anchor=MSO_ANCHOR.MIDDLE)
    # 一句话目标
    add_lines(s, 0.8, 4.0, 11.7, 1.0,
              [{'text': '🎯 本节课目标', 'size': 22, 'bold': True}])
    add_lines(s, 1.0, 4.45, 11.3, 1.2,
              [{'text': '• ' + g, 'size': 24} for g in L['goal']])
    # 底部提示
    add_rect(s, 0.5, 6.7, 12.33, 0.6, fill=TITLE_BLUE)
    add_lines(s, 0.7, 6.73, 12.0, 0.5,
              [{'text': '教师投影 · 学生只看屏（课堂无手机、无电脑）', 'size': 18, 'bold': True, 'color': WHITE}],
              anchor=MSO_ANCHOR.MIDDLE)


def build_objective(prs, L):
    s = new_slide(prs)
    add_band(s, L['ch'], L['num'], L['title'])
    add_title(s, '🎯 这节课你要会这些')
    y = 2.4
    for g in L['goal']:
        add_card(s, 0.8, y, 11.7, 1.0, fill=CARD_BLUE)
        add_lines(s, 1.1, y + 0.1, 11.2, 0.8, [{'text': '✓  ' + g, 'size': 24, 'bold': True}], anchor=MSO_ANCHOR.MIDDLE)
        y += 1.15
    add_lines(s, 0.8, y + 0.1, 11.7, 0.6, [{'text': '（能用自己的话说出来，就算学会了）', 'size': 18, 'color': RGBColor(0x60,0x60,0x60)}])


def build_intro(prs, L):
    s = new_slide(prs)
    add_band(s, L['ch'], L['num'], L['title'])
    add_title(s, '💡 先聊个生活里的例子')
    add_card(s, 0.8, 2.4, 11.7, 2.4, fill=CARD_GRAY)
    add_lines(s, 1.1, 2.65, 11.1, 2.0,
              [{'text': t, 'size': 24} for t in L['intro'].split('\n')], space_after=10, anchor=MSO_ANCHOR.MIDDLE)
    add_card(s, 0.8, 5.1, 11.7, 1.2, fill=CARD_YELLOW, border=BORDER_ORANGE)
    add_lines(s, 1.1, 5.25, 11.1, 0.95, [{'text': L['intro_a'], 'size': 22, 'bold': True, 'color': BORDER_ORANGE}], anchor=MSO_ANCHOR.MIDDLE)
    add_lines(s, 0.8, 6.5, 11.7, 0.8, [{'text': '👇 带着这个例子，看下面的知识点', 'size': 20, 'color': RGBColor(0x60,0x60,0x60)}])


def build_concept(prs, L, c):
    s = new_slide(prs)
    add_band(s, L['ch'], L['num'], L['title'])
    add_title(s, c['h'], color=TITLE_BLUE, size=28)
    add_card(s, 0.8, 2.45, 11.7, 1.9, fill=CARD_GRAY)
    add_lines(s, 1.1, 2.65, 11.1, 1.5, [{'text': c['p'], 'size': 24}], anchor=MSO_ANCHOR.MIDDLE)
    add_card(s, 0.8, 4.6, 11.7, 1.5, fill=CARD_YELLOW, border=BORDER_ORANGE)
    add_lines(s, 1.1, 4.75, 11.1, 1.2, [{'text': c['a'], 'size': 23, 'bold': True, 'color': BORDER_ORANGE}], anchor=MSO_ANCHOR.MIDDLE)
    add_lines(s, 0.8, 6.3, 11.7, 0.8, [{'text': '🧠 记住这一句就够了', 'size': 20, 'color': RGBColor(0x60,0x60,0x60)}])


def build_question(prs, L):
    s = new_slide(prs)
    add_band(s, L['ch'], L['num'], L['title'])
    add_title(s, '🤔 想一想（举手或同桌讨论）')
    y = 2.45
    for q in L['questions']:
        add_card(s, 0.8, y, 11.7, 1.3, fill=CARD_BLUE)
        add_lines(s, 1.1, y + 0.1, 11.1, 1.1, [{'text': '❓ ' + q, 'size': 22, 'bold': True}], anchor=MSO_ANCHOR.MIDDLE)
        y += 1.45
    add_lines(s, 0.8, y + 0.05, 11.7, 0.6, [{'text': '（不用写，想清楚、能说出来就行）', 'size': 18, 'color': RGBColor(0x60,0x60,0x60)}])


def build_video(prs, L):
    s = new_slide(prs)
    add_band(s, L['ch'], L['num'], L['title'])
    add_title(s, '🎬 看视频：' + L['video_name'], color=RGBColor(0x2E,0x7D,0x32))
    add_card(s, 0.8, 2.45, 11.7, 1.1, fill=CARD_GREEN)
    add_lines(s, 1.1, 2.55, 11.1, 0.9,
              [{'text': '📺 视频号：' + L['video_bv'], 'size': 22, 'bold': True, 'color': RGBColor(0x2E,0x7D,0x32)}], anchor=MSO_ANCHOR.MIDDLE)
    add_lines(s, 0.8, 3.7, 11.7, 0.5, [{'text': '👀 看的时候注意：', 'size': 22, 'bold': True}])
    add_card(s, 0.8, 4.2, 11.7, 1.1, fill=CARD_GRAY)
    add_lines(s, 1.1, 4.3, 11.1, 0.95, [{'text': '• ' + L['video_watch'], 'size': 22}], anchor=MSO_ANCHOR.MIDDLE)
    add_card(s, 0.8, 5.5, 11.7, 1.1, fill=CARD_YELLOW, border=BORDER_ORANGE)
    add_lines(s, 1.1, 5.6, 11.1, 0.95,
              [{'text': '💾 本地文件（U盘直播）：' + L['local'], 'size': 20, 'bold': True, 'color': BORDER_ORANGE}], anchor=MSO_ANCHOR.MIDDLE)
    add_lines(s, 0.8, 6.8, 11.7, 0.6, [{'text': '⚠️ 断网也能播本地文件；在线打不开就跳过，改纸上活动。', 'size': 18, 'color': RGBColor(0x60,0x60,0x60)}])


def build_activity(prs, L):
    s = new_slide(prs)
    add_band(s, L['ch'], L['num'], L['title'])
    add_title(s, '✏️ 课堂小活动（纸上完成）', color=TITLE_BLUE)
    if L['demo']:
        add_card(s, 0.8, 2.4, 11.7, 0.9, fill=CARD_BLUE)
        add_lines(s, 1.1, 2.48, 11.1, 0.75,
                  [{'text': '本节课为教师演示课：学生观看投影 + 纸笔记录，不使用任何设备。', 'size': 20, 'bold': True, 'color': TITLE_BLUE}],
                  anchor=MSO_ANCHOR.MIDDLE)
    add_card(s, 0.8, 3.45, 11.7, 2.6, fill=CARD_GRAY)
    add_lines(s, 1.1, 3.65, 11.1, 2.3, [{'text': L['activity'], 'size': 24}], anchor=MSO_ANCHOR.MIDDLE)
    add_lines(s, 0.8, 6.2, 11.7, 0.8, [{'text': '📝 活动结果同桌互查，老师抽查3组', 'size': 20, 'color': RGBColor(0x60,0x60,0x60)}])


def build_summary(prs, L):
    s = new_slide(prs)
    add_band(s, L['ch'], L['num'], L['title'])
    add_title(s, '✅ 这节课记住这3句话')
    y = 2.45
    for i, r in enumerate(L['remember'], 1):
        add_card(s, 0.8, y, 11.7, 1.15, fill=CARD_GREEN)
        add_lines(s, 1.1, y + 0.1, 11.1, 0.95,
                  [{'text': '%d. %s' % (i, r), 'size': 23, 'bold': True, 'color': RGBColor(0x2E,0x7D,0x32)}], anchor=MSO_ANCHOR.MIDDLE)
        y += 1.3
    add_lines(s, 0.8, y + 0.05, 11.7, 0.6, [{'text': '🗣️ 合上书，能背出这3句 = 达标', 'size': 20, 'color': RGBColor(0x60,0x60,0x60)}])


def build_homework(prs, L):
    s = new_slide(prs)
    add_band(s, L['ch'], L['num'], L['title'])
    add_title(s, '📒 课后小任务')
    add_card(s, 0.8, 2.45, 11.7, 2.2, fill=CARD_YELLOW, border=BORDER_ORANGE)
    add_lines(s, 1.1, 2.65, 11.1, 1.8, [{'text': '✍️ ' + L['homework'], 'size': 24, 'bold': True, 'color': BORDER_ORANGE}], anchor=MSO_ANCHOR.MIDDLE)
    add_lines(s, 0.8, 4.9, 11.7, 1.6,
              [{'text': '📌 本课时对应教材：第%d章 %s' % (L['ch'], CH_NAME[L['ch']]), 'size': 20},
               {'text': '🎬 本课时视频：' + L['video_name'], 'size': 18, 'color': RGBColor(0x60,0x60,0x60)}])


def build_lesson(L):
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    build_cover(prs, L)
    build_objective(prs, L)
    build_intro(prs, L)
    for c in L['concepts']:
        build_concept(prs, L, c)
    build_question(prs, L)
    build_video(prs, L)
    build_activity(prs, L)
    build_summary(prs, L)
    build_homework(prs, L)
    safe_title = L['title'].replace('：', '_').replace(' ', '')
    for _c in '/\\:*?"<>|':
        safe_title = safe_title.replace(_c, '_')
    fname = "课时%02d_%s.pptx" % (L['num'], safe_title)
    path = os.path.join(OUT, fname)
    prs.save(path)
    return path, len(prs.slides._sldIdLst)


if __name__ == "__main__":
    print("开始生成 12 课时通俗版PPT ...")
    for L in lessons:
        path, n = build_lesson(L)
        print("  ✓ 课时%02d %s | %d页 | %s" % (L['num'], L['title'], n, os.path.basename(path)))
    print("全部完成，输出目录：", OUT)
