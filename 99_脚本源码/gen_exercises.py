# -*- coding: utf-8 -*-
"""
生成12课时课堂习题PPT + 答案文档
"""
import sys
sys.stdout.reconfigure(encoding='utf-8')

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

# ========== 颜色定义 ==========
C_BLUE = RGBColor(0x1A, 0x73, 0xE8)
C_DARK_BLUE = RGBColor(0x0D, 0x47, 0xA1)
C_WHITE = RGBColor(0xFF, 0xFF, 0xFF)
C_BLACK = RGBColor(0x33, 0x33, 0x33)
C_GRAY = RGBColor(0x66, 0x66, 0x66)
C_LIGHT_GRAY = RGBColor(0xF0, 0xF0, 0xF0)
C_ORANGE = RGBColor(0xFF, 0x98, 0x00)
C_GREEN = RGBColor(0x4C, 0xAF, 0x50)
C_RED = RGBColor(0xF4, 0x43, 0x36)
C_PURPLE = RGBColor(0x9C, 0x27, 0xB0)
C_TEAL = RGBColor(0x00, 0x97, 0xA7)
C_LIGHT_BLUE = RGBColor(0xE3, 0xF2, 0xFD)
C_LIGHT_GREEN = RGBColor(0xE8, 0xF5, 0xE9)
C_LIGHT_ORANGE = RGBColor(0xFF, 0xF3, 0xE0)
C_LIGHT_RED = RGBColor(0xFF, 0xEB, 0xEE)
C_LIGHT_PURPLE = RGBColor(0xF3, 0xE5, 0xF5)

# ========== 题目数据 ==========
lessons = [
    {
        "num": 1,
        "chapter": "第1章 云计算的简介",
        "title": "什么是云计算",
        "questions": [
            {"type": "选择", "q": "以下哪个不属于NIST定义的云计算基本特征？", "options": ["A. 按需自助服务", "B. 资源池化", "C. 固定硬件配置", "D. 快速弹性"], "answer": "C", "explain": "NIST定义的5个特征中没有\"固定硬件配置\"，云计算的资源是动态分配的。"},
            {"type": "选择", "q": "\"云计算\"概念被广泛提出是在哪一年？", "options": ["A. 1995年", "B. 2006年", "C. 2010年", "D. 2015年"], "answer": "B", "explain": "2006年亚马逊推出EC2服务，谷歌CEO施密特首次公开使用\"Cloud Computing\"一词。"},
            {"type": "判断", "q": "云计算就是网盘，主要用来存文件。", "options": ["A. 正确", "B. 错误"], "answer": "B", "explain": "网盘只是云计算的一种应用（SaaS），云计算涵盖IaaS/PaaS/SaaS三大服务模式，远不止存文件。"},
            {"type": "判断", "q": "NIST是美国国家标准与技术研究院的缩写。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "NIST全称National Institute of Standards and Technology，其SP 800-145文件定义了云计算。"},
            {"type": "填空", "q": "NIST云计算定义的核心框架是\"3-__-4\"，其中\"3\"代表三个____模型。", "answer": "服务", "explain": "3-5-4 = 3个服务模型(SaaS/PaaS/IaaS) + 5个基本特征 + 4个部署模型。"},
        ]
    },
    {
        "num": 2,
        "chapter": "第1章 云计算的简介",
        "title": "发展历史与5种计算模式",
        "questions": [
            {"type": "选择", "q": "以下哪个不是云计算演进的阶段？", "options": ["A. 大型机时代", "B. 个人电脑时代", "C. 互联网时代", "D. 量子计算时代"], "answer": "D", "explain": "云计算的演进路径：大型机→个人电脑→互联网→云计算，量子计算不属于此演进链。"},
            {"type": "选择", "q": "\"多台计算机协同完成一个大型任务\"描述的是哪种计算模式？", "options": ["A. 分布式计算", "B. 并行计算", "C. 网格计算", "D. 以上都是"], "answer": "D", "explain": "分布式、并行、网格计算都强调多机协同，它们都是云计算的技术前身。"},
            {"type": "判断", "q": "云计算是从分布式计算、网格计算等技术发展而来的。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "云计算整合了分布式计算、网格计算、虚拟化等技术，是它们的商业化和产品化。"},
            {"type": "填空", "q": "云计算发展的三个阶段：大型机 → ____ → 云计算。", "answer": "个人电脑", "explain": "大型机（集中）→个人电脑（分散）→云计算（集中+弹性），呈螺旋式上升。"},
            {"type": "判断", "q": "网格计算和云计算是完全相同的东西。", "options": ["A. 正确", "B. 错误"], "answer": "B", "explain": "网格计算偏科研、强调跨组织共享算力；云计算偏商业、强调按需服务和计费。"},
        ]
    },
    {
        "num": 3,
        "chapter": "第1章 云计算的简介",
        "title": "演示课：云端Office协作体验",
        "questions": [
            {"type": "选择", "q": "以下哪个属于SaaS（软件即服务）产品？", "options": ["A. 阿里云ECS", "B. 腾讯文档", "C. VMware Workstation", "D. CentOS Linux"], "answer": "B", "explain": "腾讯文档是直接在浏览器中使用的在线文档工具，免安装、按需使用，属于SaaS。"},
            {"type": "选择", "q": "在线文档协作的核心优势是什么？", "options": ["A. 文件容量更大", "B. 多人同时在线编辑", "C. 只能在手机上用", "D. 必须安装专用软件"], "answer": "B", "explain": "多人实时协作是在线文档区别于传统Office的最大优势，多人在同一文档上同时编辑、实时同步。"},
            {"type": "判断", "q": "金山文档和腾讯文档都需要安装客户端才能使用。", "options": ["A. 正确", "B. 错误"], "answer": "B", "explain": "两者都支持浏览器直接打开使用，无需安装客户端，这正是SaaS的特点。"},
            {"type": "填空", "q": "SaaS的全称是____即服务（填英文）。", "answer": "Software", "explain": "SaaS = Software as a Service（软件即服务），用户直接使用云端软件。"},
            {"type": "判断", "q": "使用在线文档时，文件数据存储在本地电脑硬盘上。", "options": ["A. 正确", "B. 错误"], "answer": "B", "explain": "在线文档的数据存储在云端服务器上，不是本地硬盘，所以才能多端同步、多人协作。"},
        ]
    },
    {
        "num": 4,
        "chapter": "第2章 云计算的服务",
        "title": "SaaS 软件即服务",
        "questions": [
            {"type": "选择", "q": "以下哪个不是SaaS的特点？", "options": ["A. 免安装，浏览器直接用", "B. 用户无需维护服务器", "C. 需要自己购买和管理硬件", "D. 按需付费"], "answer": "C", "explain": "SaaS用户不需要购买和管理硬件，这些都由云服务商负责。用户只需开通账号使用软件。"},
            {"type": "选择", "q": "以下哪个属于SaaS产品？", "options": ["A. Windows 10 操作系统", "B. Office 365 在线办公", "C. VMware Workstation 虚拟机", "D. MySQL 数据库"], "answer": "B", "explain": "Office 365是微软提供的在线办公套件，通过浏览器使用、按月/年订阅付费，是典型的SaaS。"},
            {"type": "判断", "q": "SaaS用户不需要关心软件的维护和升级。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "SaaS模式下，软件的维护、升级、安全补丁都由服务商在云端完成，用户自动获得最新版本。"},
            {"type": "填空", "q": "在SaaS、PaaS、IaaS三种模式中，用户自由度最低的是____。", "answer": "SaaS", "explain": "SaaS用户只能使用服务商提供的软件功能，不能修改底层架构。自由度：IaaS > PaaS > SaaS。"},
            {"type": "判断", "q": "钉钉属于SaaS产品。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "钉钉是阿里提供的在线办公协同平台，用户通过APP/网页使用，无需自己搭建服务器，属于SaaS。"},
        ]
    },
    {
        "num": 5,
        "chapter": "第2章 云计算的服务",
        "title": "PaaS 与 IaaS",
        "questions": [
            {"type": "选择", "q": "\"自己买菜做饭\"（给你厨房和食材，你自己做）类比的是哪种云服务模式？", "options": ["A. SaaS", "B. PaaS", "C. IaaS", "D. 都不是"], "answer": "C", "explain": "IaaS提供基础设施（服务器/存储/网络），相当于给你厨房和食材，你自己做。SaaS=去饭店吃，PaaS=叫外卖半成品。"},
            {"type": "选择", "q": "以下哪个属于IaaS产品？", "options": ["A. 钉钉", "B. Gitee代码托管", "C. 阿里云ECS云服务器", "D. 微信小程序"], "answer": "C", "explain": "阿里云ECS提供虚拟服务器，用户可以在上面安装操作系统和软件，属于IaaS（基础设施即服务）。"},
            {"type": "判断", "q": "PaaS主要面向的是软件开发者。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "PaaS提供开发平台（运行环境/数据库/中间件），开发者只需关注写代码，不用管服务器配置。"},
            {"type": "填空", "q": "IaaS的全称是____即服务（填英文）。", "answer": "Infrastructure", "explain": "IaaS = Infrastructure as a Service（基础设施即服务），提供服务器/存储/网络等底层资源。"},
            {"type": "判断", "q": "使用IaaS时，用户需要自己安装操作系统。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "IaaS只提供虚拟硬件，用户需要自行选择和安装操作系统（如Linux/Windows），以及部署应用。"},
        ]
    },
    {
        "num": 6,
        "chapter": "第2章 云计算的服务",
        "title": "演示课：Gitee代码托管体验",
        "questions": [
            {"type": "选择", "q": "Gitee属于哪种云服务模式？", "options": ["A. SaaS", "B. PaaS", "C. IaaS", "D. 都不是"], "answer": "B", "explain": "Gitee提供代码托管和开发协作平台，开发者在其上管理代码、运行CI/CD，属于PaaS。也有SaaS属性，但更偏PaaS。"},
            {"type": "选择", "q": "代码托管平台的主要功能是什么？", "options": ["A. 存储照片和视频", "B. 管理和分享代码，记录修改历史", "C. 在线看视频", "D. 即时聊天通讯"], "answer": "B", "explain": "代码托管平台核心功能：代码存储、版本管理（记录每次修改）、分支管理、团队协作。"},
            {"type": "判断", "q": "Git和Gitee是同一个东西。", "options": ["A. 正确", "B. 错误"], "answer": "B", "explain": "Git是版本控制工具（软件），Gitee是基于Git的代码托管平台（网站/服务）。Git是工具，Gitee是平台。"},
            {"type": "填空", "q": "Gitee被称为中国版的____（填一个类似的国外代码托管平台名称）。", "answer": "GitHub", "explain": "Gitee和GitHub都是基于Git的代码托管平台，Gitee是国内平台（码云），GitHub是国外平台。"},
            {"type": "判断", "q": "代码托管平台可以记录代码的每一次修改历史。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "版本控制的核心就是记录每次提交（commit），可以随时回退到任意历史版本。"},
        ]
    },
    {
        "num": 7,
        "chapter": "第3章 云计算的部署",
        "title": "公有云与社区云",
        "questions": [
            {"type": "选择", "q": "以下哪个属于公有云？", "options": ["A. 学校内部机房搭建的云", "B. 阿里云", "C. 银行自建的私有数据中心", "D. 军队专用云平台"], "answer": "B", "explain": "阿里云面向公众开放，任何人都可以注册使用，是典型的公有云。"},
            {"type": "选择", "q": "公有云的最大特点是？", "options": ["A. 只对特定组织开放", "B. 对公众开放，按需使用", "C. 完全免费", "D. 不需要网络"], "answer": "B", "explain": "公有云由云服务商建设运营，对公众开放，用户按需注册使用、按量付费。"},
            {"type": "判断", "q": "社区云是多个有共同需求的组织共享使用的云。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "社区云服务于有共同需求的多个组织（如多家医院共享的医疗云、多所高校共享的教育云）。"},
            {"type": "填空", "q": "NIST定义的4种部署模型是：公有云、私有云、____和混合云。", "answer": "社区云", "explain": "4种部署模型：Public Cloud(公有云)、Private Cloud(私有云)、Community Cloud(社区云)、Hybrid Cloud(混合云)。"},
            {"type": "判断", "q": "公有云上所有用户的数据都能互相看到。", "options": ["A. 正确", "B. 错误"], "answer": "B", "explain": "虽然资源是共享的（资源池化），但通过虚拟化和隔离技术，每个用户的数据是相互隔离的，互不可见。"},
        ]
    },
    {
        "num": 8,
        "chapter": "第3章 云计算的部署",
        "title": "私有云与混合云",
        "questions": [
            {"type": "选择", "q": "以下哪个场景最适合使用私有云？", "options": ["A. 个人博客网站", "B. 银行核心交易系统", "C. 免费在线翻译工具", "D. 公开视频分享网站"], "answer": "B", "explain": "银行对数据安全和合规性要求极高，需要完全掌控基础设施，私有云最合适。"},
            {"type": "选择", "q": "混合云是指什么？", "options": ["A. 只使用公有云", "B. 只使用私有云", "C. 公有云和私有云的组合使用", "D. 完全不使用云"], "answer": "C", "explain": "混合云结合公有云的弹性和私有云的安全性，核心数据放私有云，突发流量用公有云。"},
            {"type": "判断", "q": "私有云一定比公有云更安全。", "options": ["A. 正确", "B. 错误"], "answer": "B", "explain": "安全性取决于实施水平。公有云大厂的安全投入远超普通企业自建私有云。私有云只是数据可控性更强。"},
            {"type": "填空", "q": "学校机房搭建的、仅供校内使用的云平台属于____云。", "answer": "私有", "explain": "仅供单一组织内部使用的云是私有云，学校机房云只服务校内师生，属于私有云。"},
            {"type": "判断", "q": "混合云可以兼顾数据安全性和业务灵活性。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "混合云将敏感数据放私有云保安全，将弹性业务放公有云保灵活，兼顾两者优势。"},
        ]
    },
    {
        "num": 9,
        "chapter": "第3章 云计算的部署",
        "title": "演示课：云服务器初体验",
        "questions": [
            {"type": "选择", "q": "阿里云ECS属于哪种云服务模式？", "options": ["A. SaaS", "B. PaaS", "C. IaaS", "D. 都不是"], "answer": "C", "explain": "ECS（Elastic Compute Service）提供虚拟服务器，用户可以自行安装系统和软件，属于IaaS。"},
            {"type": "选择", "q": "购买Linux云服务器后，通常用什么方式远程连接？", "options": ["A. 微信视频通话", "B. SSH（安全Shell）", "C. 蓝牙连接", "D. 用U盘拷贝文件"], "answer": "B", "explain": "SSH是远程管理Linux服务器的标准协议，通过命令行操作。Windows服务器一般用远程桌面(RDP)。"},
            {"type": "判断", "q": "云服务器可以根据需要随时升级CPU和内存配置。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "云服务器支持弹性升降配，业务增长时升级配置，业务低谷时降配省钱，这是云计算的核心优势。"},
            {"type": "填空", "q": "部署网站到云服务器的步骤：买服务器 → 连____ → 装环境 → 传网页 → 外网访问。", "answer": "SSH", "explain": "通过SSH远程连接服务器后，安装Web环境（如Nginx/Apache），上传网页文件，配置域名即可外网访问。"},
            {"type": "判断", "q": "云服务器通常按使用时长（如按小时/按月）计费。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "云服务器支持按量付费（按小时计费）和包年包月两种模式，体现了云计算的可计量服务特征。"},
        ]
    },
    {
        "num": 10,
        "chapter": "第4章 云计算的特点",
        "title": "本质与三大特征",
        "questions": [
            {"type": "选择", "q": "云计算的本质是什么？", "options": ["A. 存储", "B. 弹性", "C. 网络", "D. 虚拟化"], "answer": "B", "explain": "云计算的本质是\"弹性\"——资源能大能小、能伸能缩，按需分配和释放。"},
            {"type": "选择", "q": "\"有网就能用，手机电脑平板都行\"描述的是哪个特征？", "options": ["A. 按需自助服务", "B. 广泛网络访问", "C. 资源池化", "D. 快速弹性"], "answer": "B", "explain": "广泛网络访问(Broad Network Access)指通过各种网络和终端设备随时随地访问云资源。"},
            {"type": "判断", "q": "资源池化意味着多个用户共享同一组物理资源。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "资源池化通过虚拟化技术将物理资源池化，多用户共享，按需动态分配，互不影响。"},
            {"type": "填空", "q": "NIST定义的5个基本特征是：按需自助服务、广泛网络访问、____、快速弹性、可计量服务。", "answer": "资源池化", "explain": "5个特征：On-demand self-service、Broad network access、Resource pooling、Rapid elasticity、Measured service。"},
            {"type": "判断", "q": "按需自助服务需要联系客服人工审批才能获取资源。", "options": ["A. 正确", "B. 错误"], "answer": "B", "explain": "按需自助服务的核心就是用户可以自行获取资源，无需人工干预，像自助餐厅一样想拿就拿。"},
        ]
    },
    {
        "num": 11,
        "chapter": "第4章 云计算的特点",
        "title": "快速弹性与可计量服务",
        "questions": [
            {"type": "选择", "q": "\"双11淘宝自动增加服务器应对流量高峰\"体现了哪个特征？", "options": ["A. 资源池化", "B. 快速弹性", "C. 可计量服务", "D. 广泛网络访问"], "answer": "B", "explain": "快速弹性(Rapid Elasticity)指资源可以快速扩容和缩容，双11增加服务器、双11后释放，就是弹性。"},
            {"type": "选择", "q": "\"用了多少CPU、多少流量就收多少钱\"体现了哪个特征？", "options": ["A. 按需自助服务", "B. 快速弹性", "C. 可计量服务", "D. 资源池化"], "answer": "C", "explain": "可计量服务(Measured Service)指资源使用量可以被监控、统计和计费，像水表电表一样用多少算多少。"},
            {"type": "判断", "q": "快速弹性意味着资源既可以快速扩容，也可以快速缩容。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "弹性是双向的：需求增加时快速扩容，需求减少时快速缩容，这样才不会浪费资源。"},
            {"type": "填空", "q": "弹性的好处是____（不用为峰值买单，按实际使用量付费）。", "answer": "省钱", "explain": "传统自建机房要按峰值采购设备，平时浪费；用云可以平时少租、高峰多租、过完就退，大幅省钱。"},
            {"type": "判断", "q": "可计量服务就是按固定月费收费，不管用不用都收一样的钱。", "options": ["A. 正确", "B. 错误"], "answer": "B", "explain": "可计量服务是按实际使用量计费（用多少算多少），不是固定月费。固定月费是传统IT的模式。"},
        ]
    },
    {
        "num": 12,
        "chapter": "第4章 云计算的特点",
        "title": "演示课：VMware + 阶段测验",
        "questions": [
            {"type": "选择", "q": "VMware Workstation的主要功能是什么？", "options": ["A. 杀毒防毒", "B. 在一台电脑上创建和运行多个虚拟机", "C. 浏览网页", "D. 视频剪辑"], "answer": "B", "explain": "VMware Workstation是桌面级虚拟化软件，可以在一台物理机上虚拟出多台独立的虚拟机。"},
            {"type": "选择", "q": "虚拟化技术与云计算的关系是？", "options": ["A. 毫无关系", "B. 虚拟化是云计算的核心基础技术之一", "C. 完全相同的东西", "D. 互相替代的关系"], "answer": "B", "explain": "虚拟化技术实现了资源池化（一台物理机分多个虚拟机），是云计算IaaS层的核心技术基础。"},
            {"type": "判断", "q": "一台物理计算机可以虚拟出多台独立的虚拟机。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "通过虚拟化软件（如VMware），一台物理机的CPU/内存/硬盘可以被分割成多个虚拟机独立使用。"},
            {"type": "填空", "q": "NIST 3-5-4框架：3个____模型、5个基本特征、4个____模型。", "answer": "服务 / 部署", "explain": "3-5-4 = 3个服务模型(SaaS/PaaS/IaaS) + 5个基本特征 + 4个部署模型(公有/私有/社区/混合)。"},
            {"type": "判断", "q": "虚拟机创建完成后就不能再修改配置了。", "options": ["A. 正确", "B. 错误"], "answer": "B", "explain": "虚拟机的CPU、内存、硬盘等配置都可以随时修改（关机后调整），这正是虚拟化的灵活性所在。"},
        ]
    },
]

# ========== PPT生成 ==========
prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)

blank_layout = prs.slide_layouts[6]  # 空白布局

def add_textbox(slide, left, top, width, height, text, font_size=18, bold=False, color=C_BLACK, align=PP_ALIGN.LEFT, font_name="Microsoft YaHei"):
    txBox = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    tf = txBox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = text
    p.font.size = Pt(font_size)
    p.font.bold = bold
    p.font.color.rgb = color
    p.font.name = font_name
    p.alignment = align
    return txBox

def add_rounded_rect(slide, left, top, width, height, fill_color, line_color=None):
    shape = slide.shapes.add_shape(
        5,  # MSO_SHAPE.ROUNDED_RECTANGLE
        Inches(left), Inches(top), Inches(width), Inches(height)
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_color
    if line_color:
        shape.line.color.rgb = line_color
        shape.line.width = Pt(1.5)
    else:
        shape.line.fill.background()
    shape.shadow.inherit = False
    return shape

def add_rect(slide, left, top, width, height, fill_color, line_color=None):
    shape = slide.shapes.add_shape(
        1,  # MSO_SHAPE.RECTANGLE
        Inches(left), Inches(top), Inches(width), Inches(height)
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_color
    if line_color:
        shape.line.color.rgb = line_color
        shape.line.width = Pt(1)
    else:
        shape.line.fill.background()
    shape.shadow.inherit = False
    return shape

# ---- 封面 ----
slide = prs.slides.add_slide(blank_layout)
add_rect(slide, 0, 0, 13.333, 7.5, C_DARK_BLUE)
add_rect(slide, 0, 0, 13.333, 0.12, C_ORANGE)

add_textbox(slide, 2, 1.8, 9.3, 1.2, "云计算基础", font_size=48, bold=True, color=C_WHITE, align=PP_ALIGN.CENTER)
add_textbox(slide, 2, 3.1, 9.3, 0.8, "课堂习题集（前12课时）", font_size=28, color=C_LIGHT_BLUE, align=PP_ALIGN.CENTER)
add_rect(slide, 4.5, 4.2, 4.3, 0.04, C_ORANGE)
add_textbox(slide, 2, 4.5, 9.3, 0.6, "教材：云计算基础技术与应用(第2版)(微课版)", font_size=16, color=C_LIGHT_GRAY, align=PP_ALIGN.CENTER)
add_textbox(slide, 2, 5.1, 9.3, 0.6, "易海博 池瑞楠 主编 | 人民邮电出版社", font_size=14, color=C_GRAY, align=PP_ALIGN.CENTER)
add_textbox(slide, 2, 6.2, 9.3, 0.5, "适用：职业高中 | 共12课时 | 每课时5题", font_size=14, color=C_GRAY, align=PP_ALIGN.CENTER)

# ---- 目录页 ----
slide = prs.slides.add_slide(blank_layout)
add_rect(slide, 0, 0, 0.15, 7.5, C_BLUE)
add_textbox(slide, 0.8, 0.4, 8, 0.7, "目  录", font_size=32, bold=True, color=C_BLUE)
add_rect(slide, 0.8, 1.1, 4, 0.04, C_ORANGE)

chapters = [
    ("第1章 云计算的简介", "课时 1-3", C_BLUE),
    ("第2章 云计算的服务", "课时 4-6", C_GREEN),
    ("第3章 云计算的部署", "课时 7-9", C_ORANGE),
    ("第4章 云计算的特点", "课时 10-12", C_PURPLE),
]

y = 1.5
for ch_title, ch_range, ch_color in chapters:
    add_rounded_rect(slide, 0.8, y, 11.5, 1.1, C_LIGHT_GRAY)
    add_rect(slide, 0.8, y, 0.12, 1.1, ch_color)
    add_textbox(slide, 1.2, y + 0.15, 7, 0.5, ch_title, font_size=20, bold=True, color=C_BLACK)
    add_textbox(slide, 1.2, y + 0.65, 7, 0.4, ch_range, font_size=14, color=C_GRAY)
    y += 1.35

add_textbox(slide, 0.8, y + 0.3, 11, 0.4, "* 每课时5道题（选择题、判断题、填空题混合），适合课堂纸笔练习", font_size=12, color=C_GRAY)

# ---- 每课时习题页 ----
type_colors = {
    "选择": C_BLUE,
    "判断": C_ORANGE,
    "填空": C_GREEN,
}
type_bg = {
    "选择": C_LIGHT_BLUE,
    "判断": C_LIGHT_ORANGE,
    "填空": C_LIGHT_GREEN,
}

for lesson in lessons:
    # ---- 课时封面 ----
    slide = prs.slides.add_slide(blank_layout)
    chapter_colors = {
        "第1章": C_BLUE,
        "第2章": C_GREEN,
        "第3章": C_ORANGE,
        "第4章": C_PURPLE,
    }
    ch_key = lesson["chapter"][:3]
    ch_color = chapter_colors.get(ch_key, C_BLUE)

    add_rect(slide, 0, 0, 13.333, 7.5, ch_color)
    add_rect(slide, 0, 6.8, 13.333, 0.7, RGBColor(0xFF, 0xFF, 0xFF))

    # 课时编号大圆
    circle = slide.shapes.add_shape(4, Inches(5.17), Inches(1.0), Inches(3), Inches(3))  # OVAL
    circle.fill.solid()
    circle.fill.fore_color.rgb = C_WHITE
    circle.line.fill.background()
    circle.shadow.inherit = False
    add_textbox(slide, 5.17, 1.6, 3, 1.8, str(lesson["num"]), font_size=80, bold=True, color=ch_color, align=PP_ALIGN.CENTER)

    add_textbox(slide, 1, 4.3, 11.3, 0.6, lesson["chapter"], font_size=18, color=C_WHITE, align=PP_ALIGN.CENTER)
    add_textbox(slide, 1, 4.9, 11.3, 0.8, lesson["title"], font_size=28, bold=True, color=C_WHITE, align=PP_ALIGN.CENTER)
    add_textbox(slide, 1, 6.85, 11.3, 0.5, f"本课时共 {len(lesson['questions'])} 道题", font_size=16, color=ch_color, align=PP_ALIGN.CENTER)

    # ---- 题目页 ----
    slide = prs.slides.add_slide(blank_layout)
    # 顶部标题栏
    add_rect(slide, 0, 0, 13.333, 0.9, ch_color)
    add_textbox(slide, 0.5, 0.15, 10, 0.6, f"课时{lesson['num']}：{lesson['title']}", font_size=20, bold=True, color=C_WHITE)
    add_textbox(slide, 10.5, 0.15, 2.5, 0.6, "课堂练习", font_size=14, color=C_WHITE, align=PP_ALIGN.RIGHT)

    y_pos = 1.15
    for qi, q in enumerate(lesson["questions"]):
        tc = type_colors.get(q["type"], C_BLUE)
        tbg = type_bg.get(q["type"], C_LIGHT_BLUE)

        # 题号色块
        add_rect(slide, 0.5, y_pos, 0.5, 0.45, tc)
        add_textbox(slide, 0.5, y_pos + 0.02, 0.5, 0.4, str(qi + 1), font_size=18, bold=True, color=C_WHITE, align=PP_ALIGN.CENTER)

        # 题型标签
        add_rect(slide, 1.1, y_pos, 0.85, 0.45, tbg, line_color=tc)
        add_textbox(slide, 1.1, y_pos + 0.02, 0.85, 0.4, f"[{q['type']}]", font_size=13, bold=True, color=tc, align=PP_ALIGN.CENTER)

        # 题干
        add_textbox(slide, 2.1, y_pos + 0.02, 10.8, 0.45, q["q"], font_size=15, color=C_BLACK)

        y_pos += 0.6

        # 选项
        if "options" in q:
            if q["type"] == "选择":
                for oi, opt in enumerate(q["options"]):
                    col = oi % 2
                    row = oi // 2
                    ox = 2.1 + col * 5.5
                    oy = y_pos + row * 0.4
                    add_textbox(slide, ox, oy, 5, 0.35, opt, font_size=14, color=C_BLACK)
                y_pos += 0.85
            elif q["type"] == "判断":
                add_textbox(slide, 2.1, y_pos, 5, 0.35, "A. 正确        B. 错误", font_size=14, color=C_BLACK)
                y_pos += 0.45
        elif q["type"] == "填空":
            add_textbox(slide, 2.1, y_pos, 10, 0.35, "答：________________", font_size=14, color=C_GRAY)
            y_pos += 0.45

        y_pos += 0.1

    # 底部提示
    add_textbox(slide, 0.5, 7.0, 12, 0.4, "姓名：___________  班级：___________  得分：______", font_size=13, color=C_GRAY)

# 保存PPT
ppt_path = r"D:\WorkBuddyData\2026-08-11-07-48-22\云计算基础_前12课时课堂习题.pptx"
prs.save(ppt_path)
print(f"PPT saved: {ppt_path}")
print(f"Total slides: {len(prs.slides)}")

# ========== 答案文档（HTML）=========
html = """<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>云计算基础 前12课时 课堂习题答案与解析</title>
<style>
  * { margin: 0; padding: 0; box-sizing: border-box; }
  body { font-family: "Microsoft YaHei", "PingFang SC", sans-serif; background: #f5f5f5; color: #333; line-height: 1.8; padding: 20px; }
  .container { max-width: 900px; margin: 0 auto; background: #fff; padding: 40px; border-radius: 8px; box-shadow: 0 2px 12px rgba(0,0,0,0.08); }
  h1 { font-size: 26px; color: #1a73e8; margin-bottom: 8px; border-bottom: 3px solid #1a73e8; padding-bottom: 12px; }
  .subtitle { color: #666; font-size: 14px; margin-bottom: 30px; }
  h2 { font-size: 20px; color: #1a73e8; margin: 30px 0 15px; padding-left: 12px; border-left: 4px solid #1a73e8; }
  h2.ch2 { color: #4caf50; border-left-color: #4caf50; }
  h2.ch3 { color: #ff9800; border-left-color: #ff9800; }
  h2.ch4 { color: #9c27b0; border-left-color: #9c27b0; }
  h3 { font-size: 16px; color: #333; margin: 20px 0 10px; background: #f8f9fa; padding: 8px 12px; border-radius: 4px; }
  .q-item { margin: 12px 0; padding: 12px 15px; border: 1px solid #e0e0e0; border-radius: 6px; }
  .q-item:hover { background: #f8f9fa; }
  .q-header { display: flex; align-items: center; gap: 8px; margin-bottom: 6px; }
  .q-num { background: #1a73e8; color: #fff; width: 26px; height: 26px; border-radius: 50%; display: inline-flex; align-items: center; justify-content: center; font-size: 13px; font-weight: bold; flex-shrink: 0; }
  .q-type { display: inline-block; padding: 2px 8px; border-radius: 10px; font-size: 12px; font-weight: bold; }
  .type-select { background: #e3f2fd; color: #1565c0; }
  .type-true { background: #fff3e0; color: #e65100; }
  .type-fill { background: #e8f5e9; color: #2e7d32; }
  .q-text { font-size: 14px; margin-bottom: 6px; padding-left: 34px; }
  .q-options { font-size: 13px; color: #666; padding-left: 34px; margin-bottom: 6px; }
  .q-answer { font-size: 14px; padding: 6px 12px; background: #fffde7; border: 1px solid #fff59d; border-radius: 4px; margin-left: 34px; margin-bottom: 6px; }
  .q-answer strong { color: #f57f17; }
  .q-explain { font-size: 13px; color: #555; padding: 6px 12px; background: #f1f8e9; border-radius: 4px; margin-left: 34px; }
  .q-explain::before { content: "解析："; font-weight: bold; color: #4caf50; }
  .answer-summary { display: grid; grid-template-columns: repeat(4, 1fr); gap: 10px; margin: 20px 0; }
  .summary-card { background: #f8f9fa; border: 1px solid #e0e0e0; border-radius: 8px; padding: 10px; text-align: center; }
  .summary-card .lesson { font-size: 12px; color: #666; }
  .summary-card .answers { font-size: 14px; font-weight: bold; color: #1a73e8; margin-top: 4px; }
  @media print { .container { box-shadow: none; } .q-item { page-break-inside: avoid; } }
</style>
</head>
<body>
<div class="container">
<h1>云计算基础 前12课时 课堂习题 - 答案与解析</h1>
<p class="subtitle">教材：云计算基础技术与应用(第2版)(微课版) | 易海博 池瑞楠 主编 | 人民邮电出版社<br>
配套文档：《云计算基础_前12课时课堂习题.pptx》</p>

<div class="answer-summary">
"""

# 快速答案总表
for lesson in lessons:
    ch_num = lesson["chapter"][1]
    html += f'<div class="summary-card"><div class="lesson">课时{lesson["num"]}</div><div class="answers">'
    answers = []
    for q in lesson["questions"]:
        if q["type"] == "填空":
            answers.append(q["answer"])
        else:
            answers.append(q["answer"])
    html += " / ".join(answers)
    html += '</div></div>\n'

html += "</div>\n"

# 详细答案
for lesson in lessons:
    ch_num = lesson["chapter"][1]
    ch_class = f"ch{ch_num}"
    html += f'<h2 class="{ch_class}">课时{lesson["num"]}：{lesson["title"]}</h2>\n'
    html += f'<p style="color:#888;font-size:13px;margin-bottom:10px;">{lesson["chapter"]}</p>\n'

    for qi, q in enumerate(lesson["questions"]):
        type_class = {"选择": "type-select", "判断": "type-true", "填空": "type-fill"}.get(q["type"], "type-select")
        html += '<div class="q-item">\n'
        html += f'  <div class="q-header">'
        html += f'<span class="q-num">{qi+1}</span>'
        html += f'<span class="q-type {type_class}">[{q["type"]}]</span>'
        html += '</div>\n'
        html += f'  <div class="q-text">{q["q"]}</div>\n'
        if "options" in q:
            html += '  <div class="q-options">'
            if q["type"] == "选择":
                html += " &nbsp; ".join(q["options"])
            else:
                html += " &nbsp; ".join(q["options"])
            html += '</div>\n'
        html += f'  <div class="q-answer"><strong>答案：{q["answer"]}</strong></div>\n'
        html += f'  <div class="q-explain">{q["explain"]}</div>\n'
        html += '</div>\n'

html += """
</div>
</body>
</html>
"""

ans_path = r"D:\WorkBuddyData\2026-08-11-07-48-22\云计算基础_前12课时课堂习题_答案与解析.html"
with open(ans_path, "w", encoding="utf-8") as f:
    f.write(html)
print(f"Answer doc saved: {ans_path}")
print("Done!")
