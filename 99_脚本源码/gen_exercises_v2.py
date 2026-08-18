# -*- coding: utf-8 -*-
"""
生成12课时课堂习题PPT + 答案文档（扩充版：每课时12题，共144题）
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

# ========== 题目数据（每课时12题） ==========
lessons = [
    {
        "num": 1,
        "chapter": "第1章 云计算的简介",
        "title": "什么是云计算",
        "questions": [
            {"type": "选择", "q": "以下哪个不属于NIST定义的云计算基本特征？", "options": ["A. 按需自助服务", "B. 资源池化", "C. 固定硬件配置", "D. 快速弹性"], "answer": "C", "explain": "NIST定义的5个特征中没有\"固定硬件配置\"，云计算的资源是动态分配的。"},
            {"type": "选择", "q": "\"云计算\"概念被广泛提出是在哪一年？", "options": ["A. 1995年", "B. 2006年", "C. 2010年", "D. 2015年"], "answer": "B", "explain": "2006年亚马逊推出EC2服务，谷歌CEO施密特首次公开使用\"Cloud Computing\"一词。"},
            {"type": "选择", "q": "NIST是哪个机构的缩写？", "options": ["A. 美国国家标准与技术研究院", "B. 国际标准化组织", "C. 欧洲电信标准协会", "D. 中国信息通信研究院"], "answer": "A", "explain": "NIST全称National Institute of Standards and Technology，是美国国家标准与技术研究院。"},
            {"type": "选择", "q": "以下哪个属于云计算的应用？", "options": ["A. 用U盘拷文件", "B. 用百度网盘存照片", "C. 用本地计算器算数", "D. 用打印机打印"], "answer": "B", "explain": "百度网盘是云端存储服务，数据存在云服务器上，属于云计算应用。"},
            {"type": "选择", "q": "NIST云计算定义发布在哪个文件中？", "options": ["A. SP 800-145", "B. SP 800-53", "C. ISO 27001", "D. GB/T 31168"], "answer": "A", "explain": "NIST SP 800-145（2011年9月发布）是云计算定义的标准文献。"},
            {"type": "选择", "q": "关于云计算，以下说法正确的是？", "options": ["A. 云计算就是网盘", "B. 云计算只能用电脑访问", "C. 云计算通过互联网提供资源", "D. 云计算必须自己买服务器"], "answer": "C", "explain": "云计算通过互联网（网络）提供计算资源，网盘只是其中一种应用。"},
            {"type": "判断", "q": "云计算就是网盘，主要用来存文件。", "options": ["A. 正确", "B. 错误"], "answer": "B", "explain": "网盘只是云计算的一种应用（SaaS），云计算涵盖IaaS/PaaS/SaaS三大服务模式。"},
            {"type": "判断", "q": "NIST是美国国家标准与技术研究院的缩写。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "NIST全称National Institute of Standards and Technology，其SP 800-145文件定义了云计算。"},
            {"type": "判断", "q": "云计算的资源是可以按需获取和释放的。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "NIST定义的核心：云计算资源能够被快速分配和释放，只需最少的管理精力。"},
            {"type": "判断", "q": "使用云计算服务必须自己购买物理服务器。", "options": ["A. 正确", "B. 错误"], "answer": "B", "explain": "云计算用户租用云服务商的资源，不需要自己购买物理服务器。"},
            {"type": "填空", "q": "NIST云计算定义的核心框架是\"3-__-4\"，其中\"3\"代表三个____模型。", "answer": "服务", "explain": "3-5-4 = 3个服务模型(SaaS/PaaS/IaaS) + 5个基本特征 + 4个部署模型。"},
            {"type": "填空", "q": "NIST定义中，云计算是一种模型，允许通过____访问共享的计算资源池。", "answer": "网络", "explain": "NIST定义强调\"on-demand network access\"（按需网络访问），网络是云计算的基础。"},
        ]
    },
    {
        "num": 2,
        "chapter": "第1章 云计算的简介",
        "title": "发展历史与5种计算模式",
        "questions": [
            {"type": "选择", "q": "以下哪个不是云计算演进的阶段？", "options": ["A. 大型机时代", "B. 个人电脑时代", "C. 互联网时代", "D. 量子计算时代"], "answer": "D", "explain": "云计算演进路径：大型机->个人电脑->互联网->云计算，量子计算不属于此演进链。"},
            {"type": "选择", "q": "\"多台计算机协同完成一个大型任务\"描述的是哪种计算模式？", "options": ["A. 分布式计算", "B. 并行计算", "C. 网格计算", "D. 以上都是"], "answer": "D", "explain": "分布式、并行、网格计算都强调多机协同，它们都是云计算的技术前身。"},
            {"type": "选择", "q": "大型机时代的计算模式特点是？", "options": ["A. 高度分散", "B. 集中计算、终端只负责输入输出", "C. 每人一台独立电脑", "D. 通过互联网访问"], "answer": "B", "explain": "大型机时代是集中式计算，主机负责所有计算，终端（哑终端）只负责输入和显示。"},
            {"type": "选择", "q": "以下哪个技术是云计算的直接前身？", "options": ["A. 虚拟化技术", "B. 人工智能", "C. 区块链", "D. 量子通信"], "answer": "A", "explain": "虚拟化技术实现了资源池化，是云计算IaaS层的核心技术基础。"},
            {"type": "选择", "q": "效用计算（Utility Computing）的核心理念是？", "options": ["A. 计算资源像水电一样按需使用和计费", "B. 计算机只能用来算效用", "C. 免费提供计算资源", "D. 只用一台超级计算机"], "answer": "A", "explain": "效用计算提出\"计算即公用事业\"，像用水用电一样按量付费，这直接影响了云计算的计费模式。"},
            {"type": "选择", "q": "以下哪种计算模式强调跨组织、跨地域共享算力？", "options": ["A. 分布式计算", "B. 网格计算", "C. 并行计算", "D. 集中式计算"], "answer": "B", "explain": "网格计算强调跨组织、跨地域共享计算资源，常用于科研领域（如SETI@home）。"},
            {"type": "判断", "q": "云计算是从分布式计算、网格计算等技术发展而来的。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "云计算整合了分布式计算、网格计算、虚拟化等技术，是它们的商业化和产品化。"},
            {"type": "判断", "q": "网格计算和云计算是完全相同的东西。", "options": ["A. 正确", "B. 错误"], "answer": "B", "explain": "网格计算偏科研、强调跨组织共享算力；云计算偏商业、强调按需服务和计费。"},
            {"type": "判断", "q": "个人电脑时代的特征是计算能力从集中走向分散。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "个人电脑时代，每台PC都有独立的计算能力，从大型机的集中模式走向分散模式。"},
            {"type": "判断", "q": "虚拟化技术可以在一台物理机上运行多个操作系统。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "虚拟化技术通过Hypervisor在一台物理机上虚拟出多个独立的虚拟机，每个可以运行不同操作系统。"},
            {"type": "填空", "q": "云计算发展的三个阶段：大型机 -> ____ -> 云计算。", "answer": "个人电脑", "explain": "大型机（集中）->个人电脑（分散）->云计算（集中+弹性），呈螺旋式上升。"},
            {"type": "填空", "q": "云计算整合了分布式计算、____计算和虚拟化等技术。", "answer": "网格", "explain": "云计算是分布式计算、网格计算、虚拟化等技术的商业化和产品化。"},
        ]
    },
    {
        "num": 3,
        "chapter": "第1章 云计算的简介",
        "title": "演示课：云端Office协作体验",
        "questions": [
            {"type": "选择", "q": "以下哪个属于SaaS（软件即服务）产品？", "options": ["A. 阿里云ECS", "B. 腾讯文档", "C. VMware Workstation", "D. CentOS Linux"], "answer": "B", "explain": "腾讯文档是直接在浏览器中使用的在线文档工具，免安装、按需使用，属于SaaS。"},
            {"type": "选择", "q": "在线文档协作的核心优势是什么？", "options": ["A. 文件容量更大", "B. 多人同时在线编辑", "C. 只能在手机上用", "D. 必须安装专用软件"], "answer": "B", "explain": "多人实时协作是在线文档区别于传统Office的最大优势。"},
            {"type": "选择", "q": "以下哪个不是在线文档的优势？", "options": ["A. 多人实时协作", "B. 自动保存到云端", "C. 必须安装大型软件", "D. 多设备同步"], "answer": "C", "explain": "在线文档无需安装大型软件，浏览器直接打开即可使用。"},
            {"type": "选择", "q": "金山文档和腾讯文档属于哪种云服务模式？", "options": ["A. IaaS", "B. PaaS", "C. SaaS", "D. 都不是"], "answer": "C", "explain": "用户直接使用云端软件，无需管理底层服务器和平台，属于SaaS。"},
            {"type": "选择", "q": "传统Office软件（如Word 2019安装版）和在线文档的主要区别是？", "options": ["A. 传统软件需要安装，在线文档浏览器直接用", "B. 传统软件功能更强", "C. 在线文档不能编辑", "D. 传统软件免费"], "answer": "A", "explain": "传统Office需本地安装，在线文档通过浏览器使用，这是SaaS与传统软件的核心区别。"},
            {"type": "选择", "q": "在线文档中\"历史版本\"功能的作用是？", "options": ["A. 让文档看起来更老", "B. 查看和恢复之前的编辑记录", "C. 删除所有内容", "D. 加密文档"], "answer": "B", "explain": "在线文档自动保存每次修改的历史版本，可以随时查看和回退到任意版本。"},
            {"type": "判断", "q": "金山文档和腾讯文档都需要安装客户端才能使用。", "options": ["A. 正确", "B. 错误"], "answer": "B", "explain": "两者都支持浏览器直接打开使用，无需安装客户端。"},
            {"type": "判断", "q": "使用在线文档时，文件数据存储在本地电脑硬盘上。", "options": ["A. 正确", "B. 错误"], "answer": "B", "explain": "在线文档的数据存储在云端服务器上，所以才能多端同步、多人协作。"},
            {"type": "判断", "q": "在线文档支持多人同时编辑同一个表格的不同区域。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "在线文档支持多人实时协作，每个人编辑的位置会以不同颜色标注，互不冲突。"},
            {"type": "判断", "q": "断网后在线文档中的所有内容都会丢失。", "options": ["A. 正确", "B. 错误"], "answer": "B", "explain": "数据存在云端，断网只是暂时无法编辑，恢复网络后内容仍然完好。"},
            {"type": "填空", "q": "SaaS的全称是____即服务（填英文）。", "answer": "Software", "explain": "SaaS = Software as a Service（软件即服务），用户直接使用云端软件。"},
            {"type": "填空", "q": "在线文档的多人协作功能依赖____端存储和同步技术。", "answer": "云", "explain": "多人协作依赖云端服务器实时同步各用户的编辑操作。"},
        ]
    },
    {
        "num": 4,
        "chapter": "第2章 云计算的服务",
        "title": "SaaS 软件即服务",
        "questions": [
            {"type": "选择", "q": "以下哪个不是SaaS的特点？", "options": ["A. 免安装，浏览器直接用", "B. 用户无需维护服务器", "C. 需要自己购买和管理硬件", "D. 按需付费"], "answer": "C", "explain": "SaaS用户不需要购买和管理硬件，这些都由云服务商负责。"},
            {"type": "选择", "q": "以下哪个属于SaaS产品？", "options": ["A. Windows 10 操作系统", "B. Office 365 在线办公", "C. VMware Workstation 虚拟机", "D. MySQL 数据库"], "answer": "B", "explain": "Office 365是微软提供的在线办公套件，通过浏览器使用、按月/年订阅付费。"},
            {"type": "选择", "q": "以下哪个不属于SaaS产品？", "options": ["A. 钉钉", "B. 企业微信", "C. 阿里云ECS云服务器", "D. 飞书"], "answer": "C", "explain": "阿里云ECS提供的是虚拟服务器（基础设施），属于IaaS，不是SaaS。"},
            {"type": "选择", "q": "SaaS模式下，软件的升级和维护由谁负责？", "options": ["A. 用户自己", "B. 云服务商", "C. 硬件厂商", "D. 网络运营商"], "answer": "B", "explain": "SaaS模式下，软件的维护、升级、安全补丁都由服务商在云端完成。"},
            {"type": "选择", "q": "SaaS的典型付费方式是？", "options": ["A. 一次性买断", "B. 按月/年订阅", "C. 按CPU核数付费", "D. 完全免费"], "answer": "B", "explain": "SaaS通常采用订阅制（按月或按年付费），用户按需选择套餐。"},
            {"type": "选择", "q": "\"去饭店吃饭，菜做好直接吃\"类比的是哪种云服务？", "options": ["A. SaaS", "B. PaaS", "C. IaaS", "D. 本地部署"], "answer": "A", "explain": "SaaS就像去饭店吃饭，一切准备好，你直接享用成品。"},
            {"type": "判断", "q": "SaaS用户不需要关心软件的维护和升级。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "SaaS模式下，软件的维护、升级、安全补丁都由服务商在云端完成。"},
            {"type": "判断", "q": "钉钉属于SaaS产品。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "钉钉是阿里提供的在线办公协同平台，用户通过APP/网页使用，属于SaaS。"},
            {"type": "判断", "q": "SaaS用户可以自由修改软件的底层代码和架构。", "options": ["A. 正确", "B. 错误"], "answer": "B", "explain": "SaaS用户只能使用服务商提供的功能，不能修改底层代码。自由度最低。"},
            {"type": "判断", "q": "使用SaaS软件需要自己配置数据库和服务器环境。", "options": ["A. 正确", "B. 错误"], "answer": "B", "explain": "SaaS用户无需配置任何底层环境，开通账号即可使用。"},
            {"type": "填空", "q": "在SaaS、PaaS、IaaS三种模式中，用户自由度最低的是____。", "answer": "SaaS", "explain": "自由度：IaaS > PaaS > SaaS。SaaS最省心但自由度最低。"},
            {"type": "填空", "q": "SaaS的英文全称是____ as a Service。", "answer": "Software", "explain": "SaaS = Software as a Service（软件即服务）。"},
        ]
    },
    {
        "num": 5,
        "chapter": "第2章 云计算的服务",
        "title": "PaaS 与 IaaS",
        "questions": [
            {"type": "选择", "q": "\"自己买菜做饭\"（给你厨房和食材）类比的是哪种云服务？", "options": ["A. SaaS", "B. PaaS", "C. IaaS", "D. 都不是"], "answer": "C", "explain": "IaaS提供基础设施（服务器/存储/网络），相当于给你厨房和食材，你自己做。"},
            {"type": "选择", "q": "以下哪个属于IaaS产品？", "options": ["A. 钉钉", "B. Gitee代码托管", "C. 阿里云ECS云服务器", "D. 微信小程序"], "answer": "C", "explain": "阿里云ECS提供虚拟服务器，用户可以在上面安装操作系统和软件，属于IaaS。"},
            {"type": "选择", "q": "PaaS主要面向哪类用户？", "options": ["A. 普通办公人员", "B. 软件开发者", "C. 硬件维修人员", "D. 网络工程师"], "answer": "B", "explain": "PaaS提供开发平台（运行环境/数据库/中间件），开发者只需关注写代码。"},
            {"type": "选择", "q": "以下哪个属于PaaS产品？", "options": ["A. 阿里云ECS", "B. 腾讯文档", "C. 微信小程序云开发", "D. Windows 10"], "answer": "C", "explain": "微信小程序云开发提供开发环境和数据库，开发者只需写代码部署，属于PaaS。"},
            {"type": "选择", "q": "\"叫外卖半成品，自己加热就能吃\"类比的是？", "options": ["A. SaaS", "B. PaaS", "C. IaaS", "D. 本地部署"], "answer": "B", "explain": "PaaS提供了平台和工具（半成品），开发者在此基础上开发应用（加热即食）。"},
            {"type": "选择", "q": "三种服务模式中，用户自由度最高的是？", "options": ["A. SaaS", "B. PaaS", "C. IaaS", "D. 一样高"], "answer": "C", "explain": "自由度：IaaS > PaaS > SaaS。IaaS用户可以自由选择操作系统和软件。"},
            {"type": "判断", "q": "PaaS主要面向的是软件开发者。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "PaaS提供开发平台，开发者只需关注写代码，不用管服务器配置。"},
            {"type": "判断", "q": "使用IaaS时，用户需要自己安装操作系统。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "IaaS只提供虚拟硬件，用户需要自行选择和安装操作系统。"},
            {"type": "判断", "q": "SaaS、PaaS、IaaS三层之间是互相独立、没有关系的。", "options": ["A. 正确", "B. 错误"], "answer": "B", "explain": "三者是层级关系：PaaS建立在IaaS之上，SaaS建立在PaaS之上。"},
            {"type": "判断", "q": "IaaS用户需要自己管理物理服务器的硬件故障。", "options": ["A. 正确", "B. 错误"], "answer": "B", "explain": "物理硬件由云服务商管理，IaaS用户只管虚拟机层面以上的内容。"},
            {"type": "填空", "q": "IaaS的全称是____即服务（填英文）。", "answer": "Infrastructure", "explain": "IaaS = Infrastructure as a Service（基础设施即服务）。"},
            {"type": "填空", "q": "PaaS的全称是____即服务（填英文）。", "answer": "Platform", "explain": "PaaS = Platform as a Service（平台即服务）。"},
        ]
    },
    {
        "num": 6,
        "chapter": "第2章 云计算的服务",
        "title": "演示课：Gitee代码托管体验",
        "questions": [
            {"type": "选择", "q": "Gitee属于哪种云服务模式？", "options": ["A. SaaS", "B. PaaS", "C. IaaS", "D. 都不是"], "answer": "B", "explain": "Gitee提供代码托管和开发协作平台，开发者在其上管理代码，属于PaaS。"},
            {"type": "选择", "q": "代码托管平台的主要功能是什么？", "options": ["A. 存储照片和视频", "B. 管理和分享代码，记录修改历史", "C. 在线看视频", "D. 即时聊天通讯"], "answer": "B", "explain": "代码托管平台核心功能：代码存储、版本管理、分支管理、团队协作。"},
            {"type": "选择", "q": "以下哪个是国外的代码托管平台？", "options": ["A. Gitee（码云）", "B. GitHub", "C. 钉钉", "D. 飞书"], "answer": "B", "explain": "GitHub是全球最大的代码托管平台（美国），Gitee是国内对应的平台。"},
            {"type": "选择", "q": "Git是什么？", "options": ["A. 一个代码托管网站", "B. 一个版本控制工具", "C. 一种编程语言", "D. 一种操作系统"], "answer": "B", "explain": "Git是版本控制软件（工具），Gitee/GitHub是基于Git的代码托管平台（服务）。"},
            {"type": "选择", "q": "代码托管平台上的\"仓库\"（Repository）是指？", "options": ["A. 存放货物的仓库", "B. 存放和管理一个项目代码的空间", "C. 服务器的机柜", "D. 数据库的表"], "answer": "B", "explain": "仓库是代码托管平台的基本单位，每个项目对应一个仓库，包含全部代码和历史记录。"},
            {"type": "选择", "q": "代码托管平台的\"提交\"（Commit）功能是？", "options": ["A. 提交作业", "B. 保存代码的一次修改记录", "C. 删除代码", "D. 编译代码"], "answer": "B", "explain": "每次Commit记录代码的修改内容、时间和修改者，形成版本历史。"},
            {"type": "判断", "q": "Git和Gitee是同一个东西。", "options": ["A. 正确", "B. 错误"], "answer": "B", "explain": "Git是版本控制工具（软件），Gitee是基于Git的代码托管平台（网站）。"},
            {"type": "判断", "q": "代码托管平台可以记录代码的每一次修改历史。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "版本控制的核心就是记录每次提交，可以随时回退到任意历史版本。"},
            {"type": "判断", "q": "Gitee上的代码仓库只能设置为公开，不能设为私有。", "options": ["A. 正确", "B. 错误"], "answer": "B", "explain": "Gitee支持公开和私有仓库，私有仓库只有授权成员才能访问。"},
            {"type": "判断", "q": "多人协作时，代码托管平台可以帮助合并不同人的代码修改。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "通过分支（Branch）和合并请求（Pull/Merge Request）实现多人协作。"},
            {"type": "填空", "q": "Gitee被称为中国版的____（填一个国外代码托管平台名称）。", "answer": "GitHub", "explain": "Gitee（码云）和GitHub都是基于Git的代码托管平台。"},
            {"type": "填空", "q": "Git是一个____控制工具（填三个字）。", "answer": "版本", "explain": "Git是目前最流行的分布式版本控制系统。"},
        ]
    },
    {
        "num": 7,
        "chapter": "第3章 云计算的部署",
        "title": "公有云与社区云",
        "questions": [
            {"type": "选择", "q": "以下哪个属于公有云？", "options": ["A. 学校内部机房搭建的云", "B. 阿里云", "C. 银行自建的私有数据中心", "D. 军队专用云平台"], "answer": "B", "explain": "阿里云面向公众开放，任何人都可以注册使用，是典型的公有云。"},
            {"type": "选择", "q": "公有云的最大特点是？", "options": ["A. 只对特定组织开放", "B. 对公众开放，按需使用", "C. 完全免费", "D. 不需要网络"], "answer": "B", "explain": "公有云由云服务商建设运营，对公众开放，用户按需注册使用。"},
            {"type": "选择", "q": "以下哪个不是国内公有云服务商？", "options": ["A. 阿里云", "B. 腾讯云", "C. 华为云", "D. Oracle Cloud"], "answer": "D", "explain": "Oracle Cloud（甲骨文云）是美国公司，不属于国内公有云服务商。"},
            {"type": "选择", "q": "社区云的典型用户是？", "options": ["A. 个人用户", "B. 有共同需求的多个组织", "C. 单一企业", "D. 所有人"], "answer": "B", "explain": "社区云服务于有共同需求的多个组织（如多家医院共享的医疗云）。"},
            {"type": "选择", "q": "以下哪个场景适合使用社区云？", "options": ["A. 个人存照片", "B. 多家医院共享电子病历系统", "C. 银行核心交易", "D. 个人博客"], "answer": "B", "explain": "多家医院有共同的医疗数据共享需求，社区云正适合这种跨组织协作场景。"},
            {"type": "选择", "q": "公有云用户的数据安全性靠什么保障？", "options": ["A. 物理隔离", "B. 虚拟化和逻辑隔离技术", "C. 用户自己加密", "D. 不需要保障"], "answer": "B", "explain": "公有云通过虚拟化和逻辑隔离技术确保不同用户之间的数据互不可见。"},
            {"type": "判断", "q": "社区云是多个有共同需求的组织共享使用的云。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "社区云服务于有共同需求的多个组织（如医疗云、教育云）。"},
            {"type": "判断", "q": "公有云上所有用户的数据都能互相看到。", "options": ["A. 正确", "B. 错误"], "answer": "B", "explain": "通过虚拟化和隔离技术，每个用户的数据是相互隔离的，互不可见。"},
            {"type": "判断", "q": "公有云的优点包括成本低、无需维护硬件、弹性伸缩。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "公有云由服务商统一维护，用户共享资源，成本最低，弹性最好。"},
            {"type": "判断", "q": "社区云和公有云是一回事。", "options": ["A. 正确", "B. 错误"], "answer": "B", "explain": "社区云只对特定群体开放，公有云对所有人开放。"},
            {"type": "填空", "q": "NIST定义的4种部署模型是：公有云、私有云、____和混合云。", "answer": "社区云", "explain": "4种部署模型：Public/Private/Community/Hybrid Cloud。"},
            {"type": "填空", "q": "阿里云、腾讯云、华为云都属于____云。", "answer": "公有", "explain": "面向公众开放的云平台属于公有云。"},
        ]
    },
    {
        "num": 8,
        "chapter": "第3章 云计算的部署",
        "title": "私有云与混合云",
        "questions": [
            {"type": "选择", "q": "以下哪个场景最适合使用私有云？", "options": ["A. 个人博客网站", "B. 银行核心交易系统", "C. 免费在线翻译工具", "D. 公开视频分享网站"], "answer": "B", "explain": "银行对数据安全和合规性要求极高，需要完全掌控基础设施，私有云最合适。"},
            {"type": "选择", "q": "混合云是指什么？", "options": ["A. 只使用公有云", "B. 只使用私有云", "C. 公有云和私有云的组合使用", "D. 完全不使用云"], "answer": "C", "explain": "混合云结合公有云的弹性和私有云的安全性。"},
            {"type": "选择", "q": "以下哪个是混合云的典型应用场景？", "options": ["A. 个人存照片", "B. 电商平时用私有云，双11用公有云扩容", "C. 只用一台服务器", "D. 完全不用网络"], "answer": "B", "explain": "核心数据放私有云保安全，突发流量用公有云保弹性，这是混合云的经典场景。"},
            {"type": "选择", "q": "私有云的部署方式包括？", "options": ["A. 只能在企业自己的机房", "B. 只能托管在第三方机房", "C. 可自建也可托管", "D. 必须用公有云厂商的私有云方案"], "answer": "C", "explain": "私有云可以在企业自有机房部署，也可以托管在第三方数据中心。"},
            {"type": "选择", "q": "以下哪个不是私有云的优势？", "options": ["A. 数据完全可控", "B. 满足合规要求", "C. 建设成本低", "D. 安全性可定制"], "answer": "C", "explain": "私有云需要自购硬件、自建团队，建设成本通常高于公有云。"},
            {"type": "选择", "q": "混合云解决了什么问题？", "options": ["A. 网速太慢", "B. 公有云不够安全和私有云不够弹性的矛盾", "C. 服务器太贵", "D. 软件不好用"], "answer": "B", "explain": "混合云兼顾安全（私有云）和弹性（公有云），解决两者的矛盾。"},
            {"type": "判断", "q": "私有云一定比公有云更安全。", "options": ["A. 正确", "B. 错误"], "answer": "B", "explain": "安全性取决于实施水平。公有云大厂的安全投入远超普通企业自建私有云。"},
            {"type": "判断", "q": "学校机房搭建的、仅供校内使用的云平台属于私有云。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "仅供单一组织内部使用的云是私有云。"},
            {"type": "判断", "q": "混合云可以兼顾数据安全性和业务灵活性。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "混合云将敏感数据放私有云保安全，弹性业务放公有云保灵活。"},
            {"type": "判断", "q": "私有云的使用成本通常低于公有云。", "options": ["A. 正确", "B. 错误"], "answer": "B", "explain": "私有云需要自购硬件和运维团队，初始成本高；公有云按需付费，门槛低。"},
            {"type": "填空", "q": "学校机房搭建的、仅供校内使用的云平台属于____云。", "answer": "私有", "explain": "仅供单一组织内部使用的云是私有云。"},
            {"type": "填空", "q": "____云结合了公有云的弹性和私有云的安全性。", "answer": "混合", "explain": "混合云 = 公有云 + 私有云的组合。"},
        ]
    },
    {
        "num": 9,
        "chapter": "第3章 云计算的部署",
        "title": "演示课：云服务器初体验",
        "questions": [
            {"type": "选择", "q": "阿里云ECS属于哪种云服务模式？", "options": ["A. SaaS", "B. PaaS", "C. IaaS", "D. 都不是"], "answer": "C", "explain": "ECS（Elastic Compute Service）提供虚拟服务器，属于IaaS。"},
            {"type": "选择", "q": "购买Linux云服务器后，通常用什么方式远程连接？", "options": ["A. 微信视频通话", "B. SSH（安全Shell）", "C. 蓝牙连接", "D. 用U盘拷贝文件"], "answer": "B", "explain": "SSH是远程管理Linux服务器的标准协议。Windows服务器一般用远程桌面(RDP)。"},
            {"type": "选择", "q": "云服务器的\"弹性\"体现在？", "options": ["A. 服务器可以弯曲", "B. 可以随时升级或降低配置", "C. 服务器可以移动", "D. 服务器是软的"], "answer": "B", "explain": "弹性指资源可以随时扩容（升级配置）和缩容（降配省钱）。"},
            {"type": "选择", "q": "以下哪个是云服务器的计费方式？", "options": ["A. 只能按年付费", "B. 按量付费和包年包月", "C. 完全免费", "D. 按点击次数付费"], "answer": "B", "explain": "云服务器支持按量付费（按小时）和包年包月两种模式。"},
            {"type": "选择", "q": "部署网站到云服务器的一般步骤是？", "options": ["A. 买服务器->装环境->传网页->外网访问", "B. 买服务器->传网页->装环境->外网访问", "C. 传网页->买服务器->装环境->外网访问", "D. 装环境->传网页->买服务器->外网访问"], "answer": "A", "explain": "正确顺序：购买服务器->连接SSH->安装Web环境->上传网页文件->配置外网访问。"},
            {"type": "选择", "q": "云服务器选择操作系统时，建站常用的是？", "options": ["A. 只能用Linux", "B. 只能用Windows", "C. Linux或Windows均可", "D. 不需要操作系统"], "answer": "C", "explain": "Linux（如CentOS/Ubuntu）和Windows Server都可以建站，Linux更常用。"},
            {"type": "判断", "q": "云服务器可以根据需要随时升级CPU和内存配置。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "云服务器支持弹性升降配，这是云计算的核心优势。"},
            {"type": "判断", "q": "云服务器通常按使用时长（如按小时/按月）计费。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "体现了云计算的可计量服务特征。"},
            {"type": "判断", "q": "云服务器购买后不能更换操作系统。", "options": ["A. 正确", "B. 错误"], "answer": "B", "explain": "云服务器支持更换操作系统（需重装系统），但原有数据会丢失。"},
            {"type": "判断", "q": "云服务器的数据会自动备份，不需要手动操作。", "options": ["A. 正确", "B. 错误"], "answer": "B", "explain": "云服务器默认不自动备份，需要手动创建快照或配置自动备份策略。"},
            {"type": "填空", "q": "部署网站到云服务器的步骤：买服务器 -> 连____ -> 装环境 -> 传网页 -> 外网访问。", "answer": "SSH", "explain": "通过SSH远程连接Linux服务器进行操作。"},
            {"type": "填空", "q": "阿里云ECS的全称是Elastic ____ Service。", "answer": "Compute", "explain": "ECS = Elastic Compute Service（弹性计算服务）。"},
        ]
    },
    {
        "num": 10,
        "chapter": "第4章 云计算的特点",
        "title": "本质与三大特征",
        "questions": [
            {"type": "选择", "q": "云计算的本质是什么？", "options": ["A. 存储", "B. 弹性", "C. 网络", "D. 虚拟化"], "answer": "B", "explain": "云计算的本质是\"弹性\"--资源能大能小、能伸能缩。"},
            {"type": "选择", "q": "\"有网就能用，手机电脑平板都行\"描述的是哪个特征？", "options": ["A. 按需自助服务", "B. 广泛网络访问", "C. 资源池化", "D. 快速弹性"], "answer": "B", "explain": "广泛网络访问指通过各种网络和终端设备随时随地访问云资源。"},
            {"type": "选择", "q": "\"用户自己点几下就能开通服务器，不用找客服\"描述的是？", "options": ["A. 按需自助服务", "B. 广泛网络访问", "C. 资源池化", "D. 可计量服务"], "answer": "A", "explain": "按需自助服务指用户能自行获取资源，不需人工干预。"},
            {"type": "选择", "q": "\"你不知道数据存在哪台服务器上，但能正常访问\"体现了？", "options": ["A. 按需自助服务", "B. 广泛网络访问", "C. 资源池化", "D. 快速弹性"], "answer": "C", "explain": "资源池化通过虚拟化将物理资源统一管理，用户不关心具体物理位置。"},
            {"type": "选择", "q": "虚拟化技术与资源池化的关系是？", "options": ["A. 虚拟化实现了资源池化", "B. 资源池化实现了虚拟化", "C. 两者无关", "D. 两者完全相同"], "answer": "A", "explain": "虚拟化技术将物理资源抽象为虚拟资源池，实现资源的动态分配。"},
            {"type": "选择", "q": "以下哪个是\"按需自助服务\"的生活类比？", "options": ["A. 自助餐厅，想拿什么自己拿", "B. 去饭店点菜", "C. 叫外卖", "D. 自己做饭"], "answer": "A", "explain": "按需自助服务就像自助餐厅，用户自行获取所需资源。"},
            {"type": "判断", "q": "资源池化意味着多个用户共享同一组物理资源。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "资源池化通过虚拟化技术将物理资源池化，多用户共享，互不影响。"},
            {"type": "判断", "q": "按需自助服务需要联系客服人工审批才能获取资源。", "options": ["A. 正确", "B. 错误"], "answer": "B", "explain": "按需自助服务的核心就是用户可以自行获取资源，无需人工干预。"},
            {"type": "判断", "q": "广泛网络访问意味着只要有网络，任何设备都能访问云服务。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "广泛网络访问支持各种终端（手机、电脑、平板）通过网络访问。"},
            {"type": "判断", "q": "资源池化后，不同用户的数据可能会混淆。", "options": ["A. 正确", "B. 错误"], "answer": "B", "explain": "虽然共享物理资源，但通过隔离技术，每个用户的数据完全独立、互不可见。"},
            {"type": "填空", "q": "NIST定义的5个基本特征是：按需自助服务、广泛网络访问、____、快速弹性、可计量服务。", "answer": "资源池化", "explain": "5个特征：On-demand self-service、Broad network access、Resource pooling、Rapid elasticity、Measured service。"},
            {"type": "填空", "q": "云计算的本质是____（能大能小、能伸能缩）。", "answer": "弹性", "explain": "弹性是云计算最核心的本质特征。"},
        ]
    },
    {
        "num": 11,
        "chapter": "第4章 云计算的特点",
        "title": "快速弹性与可计量服务",
        "questions": [
            {"type": "选择", "q": "\"双11淘宝自动增加服务器应对流量高峰\"体现了哪个特征？", "options": ["A. 资源池化", "B. 快速弹性", "C. 可计量服务", "D. 广泛网络访问"], "answer": "B", "explain": "快速弹性指资源可以快速扩容和缩容。"},
            {"type": "选择", "q": "\"用了多少CPU、多少流量就收多少钱\"体现了哪个特征？", "options": ["A. 按需自助服务", "B. 快速弹性", "C. 可计量服务", "D. 资源池化"], "answer": "C", "explain": "可计量服务指资源使用量可以被监控、统计和计费。"},
            {"type": "选择", "q": "以下哪个是快速弹性的生活类比？", "options": ["A. 水表电表", "B. 弹性腰带", "C. 自助餐厅", "D. 酒店住宿"], "answer": "B", "explain": "弹性腰带能大能小，就像云资源能扩能缩。水表电表类比的是可计量服务。"},
            {"type": "选择", "q": "可计量服务的生活类比是？", "options": ["A. 弹性腰带", "B. 水表电表（用多少算多少）", "C. 自助餐厅", "D. 住酒店"], "answer": "B", "explain": "水表电表按实际用量计费，就像云计算按资源使用量计费。"},
            {"type": "选择", "q": "传统自建机房与云计算在成本上的主要区别是？", "options": ["A. 自建机房按峰值采购，平时浪费；云计算按需付费", "B. 自建机房更便宜", "C. 云计算更贵", "D. 两者成本相同"], "answer": "A", "explain": "自建机房要按最高需求采购设备，平时大量闲置；云计算用多少付多少。"},
            {"type": "选择", "q": "以下哪个场景体现了快速弹性？", "options": ["A. 手机上用钉钉", "B. 直播平台高峰期自动扩容、结束后自动缩容", "C. 云盘存照片", "D. 在线文档协作"], "answer": "B", "explain": "直播高峰期自动增加服务器、结束后自动释放，是快速弹性的典型场景。"},
            {"type": "判断", "q": "快速弹性意味着资源既可以快速扩容，也可以快速缩容。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "弹性是双向的：需求增加时扩容，需求减少时缩容。"},
            {"type": "判断", "q": "可计量服务就是按固定月费收费，不管用不用都收一样的钱。", "options": ["A. 正确", "B. 错误"], "answer": "B", "explain": "可计量服务是按实际使用量计费，不是固定月费。"},
            {"type": "判断", "q": "弹性的好处是可以省钱，不用为峰值买单。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "传统自建机房按峰值采购浪费大；用云可以平时少租、高峰多租、过完就退。"},
            {"type": "判断", "q": "可计量服务意味着用户可以查看自己的资源使用明细。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "可计量服务提供资源使用监控和报告，用户可以在控制台查看用量和费用明细。"},
            {"type": "填空", "q": "弹性的好处是____（不用为峰值买单，按实际使用量付费）。", "answer": "省钱", "explain": "弹性扩缩容避免了为峰值采购设备的浪费。"},
            {"type": "填空", "q": "NIST 5个特征中，\"用多少算多少\"对应的是____服务。", "answer": "可计量", "explain": "可计量服务(Measured Service)按实际使用量监控和计费。"},
        ]
    },
    {
        "num": 12,
        "chapter": "第4章 云计算的特点",
        "title": "演示课：VMware + 阶段测验",
        "questions": [
            {"type": "选择", "q": "VMware Workstation的主要功能是什么？", "options": ["A. 杀毒防毒", "B. 在一台电脑上创建和运行多个虚拟机", "C. 浏览网页", "D. 视频剪辑"], "answer": "B", "explain": "VMware Workstation是桌面级虚拟化软件，可以在一台物理机上虚拟出多台虚拟机。"},
            {"type": "选择", "q": "虚拟化技术与云计算的关系是？", "options": ["A. 毫无关系", "B. 虚拟化是云计算的核心基础技术之一", "C. 完全相同的东西", "D. 互相替代的关系"], "answer": "B", "explain": "虚拟化技术实现了资源池化，是云计算IaaS层的核心技术基础。"},
            {"type": "选择", "q": "以下哪个不是虚拟化的优势？", "options": ["A. 提高资源利用率", "B. 降低硬件成本", "C. 降低软件性能", "D. 方便备份和迁移"], "answer": "C", "explain": "虚拟化会带来轻微性能损耗，但不会\"降低软件性能\"，优势远大于劣势。"},
            {"type": "选择", "q": "虚拟机和物理机的区别是？", "options": ["A. 虚拟机没有CPU", "B. 虚拟机的资源是从物理机虚拟出来的", "C. 虚拟机不能运行操作系统", "D. 虚拟机没有硬盘"], "answer": "B", "explain": "虚拟机的CPU/内存/硬盘等都是从物理机通过虚拟化技术分配出来的。"},
            {"type": "选择", "q": "VMware创建虚拟机时需要指定什么？", "options": ["A. CPU核数、内存大小、硬盘容量", "B. 显示器品牌", "C. 键盘颜色", "D. 鼠标型号"], "answer": "A", "explain": "创建虚拟机需要配置CPU、内存、硬盘等硬件参数，以及选择操作系统。"},
            {"type": "选择", "q": "NIST 3-5-4框架中，\"3\"对应教材哪几章？", "options": ["A. 第1-3章", "B. 第2章", "C. 第4章", "D. 第5章"], "answer": "B", "explain": "\"3\"是3个服务模型(SaaS/PaaS/IaaS)，对应第2章\"云计算的服务\"。"},
            {"type": "判断", "q": "一台物理计算机可以虚拟出多台独立的虚拟机。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "通过虚拟化软件，一台物理机的资源可以被分割成多个虚拟机独立使用。"},
            {"type": "判断", "q": "虚拟机创建完成后就不能再修改配置了。", "options": ["A. 正确", "B. 错误"], "answer": "B", "explain": "虚拟机的CPU、内存、硬盘等配置都可以随时修改（关机后调整）。"},
            {"type": "判断", "q": "虚拟机可以像物理机一样安装操作系统和应用软件。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "虚拟机是完整的计算机系统，可以安装操作系统和各种软件。"},
            {"type": "判断", "q": "NIST 3-5-4框架的\"4\"对应教材第3章。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "\"4\"是4个部署模型，对应第3章\"云计算的部署\"。"},
            {"type": "填空", "q": "NIST 3-5-4框架：3个____模型、5个基本特征、4个____模型。", "answer": "服务 / 部署", "explain": "3个服务模型 + 5个特征 + 4个部署模型。"},
            {"type": "填空", "q": "VMware是一种____化软件（填两个字）。", "answer": "虚拟", "explain": "VMware是虚拟化软件，通过Hypervisor在一台物理机上运行多个虚拟机。"},
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
add_textbox(slide, 2, 6.2, 9.3, 0.5, "适用：职业高中 | 共12课时 | 每课时12题 | 共144题", font_size=14, color=C_GRAY, align=PP_ALIGN.CENTER)

# ---- 目录页 ----
slide = prs.slides.add_slide(blank_layout)
add_rect(slide, 0, 0, 0.15, 7.5, C_BLUE)
add_textbox(slide, 0.8, 0.4, 8, 0.7, "目  录", font_size=32, bold=True, color=C_BLUE)
add_rect(slide, 0.8, 1.1, 4, 0.04, C_ORANGE)

chapters = [
    ("第1章 云计算的简介", "课时 1-3（36题）", C_BLUE),
    ("第2章 云计算的服务", "课时 4-6（36题）", C_GREEN),
    ("第3章 云计算的部署", "课时 7-9（36题）", C_ORANGE),
    ("第4章 云计算的特点", "课时 10-12（36题）", C_PURPLE),
]

y = 1.5
for ch_title, ch_range, ch_color in chapters:
    add_rounded_rect(slide, 0.8, y, 11.5, 1.1, C_LIGHT_GRAY)
    add_rect(slide, 0.8, y, 0.12, 1.1, ch_color)
    add_textbox(slide, 1.2, y + 0.15, 7, 0.5, ch_title, font_size=20, bold=True, color=C_BLACK)
    add_textbox(slide, 1.2, y + 0.65, 7, 0.4, ch_range, font_size=14, color=C_GRAY)
    y += 1.35

add_textbox(slide, 0.8, y + 0.3, 11, 0.4, "* 每课时12道题（选择题6+判断题4+填空题2），适合课堂纸笔练习", font_size=12, color=C_GRAY)

# ---- 题型分布说明 ----
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

# ---- 每课时习题页（拆2页，每页6题） ----
for lesson in lessons:
    chapter_colors = {
        "第1章": C_BLUE,
        "第2章": C_GREEN,
        "第3章": C_ORANGE,
        "第4章": C_PURPLE,
    }
    ch_key = lesson["chapter"][:3]
    ch_color = chapter_colors.get(ch_key, C_BLUE)

    # ---- 课时封面 ----
    slide = prs.slides.add_slide(blank_layout)
    add_rect(slide, 0, 0, 13.333, 7.5, ch_color)
    add_rect(slide, 0, 6.8, 13.333, 0.7, RGBColor(0xFF, 0xFF, 0xFF))

    circle = slide.shapes.add_shape(4, Inches(5.17), Inches(1.0), Inches(3), Inches(3))  # OVAL
    circle.fill.solid()
    circle.fill.fore_color.rgb = C_WHITE
    circle.line.fill.background()
    circle.shadow.inherit = False
    add_textbox(slide, 5.17, 1.6, 3, 1.8, str(lesson["num"]), font_size=80, bold=True, color=ch_color, align=PP_ALIGN.CENTER)

    add_textbox(slide, 1, 4.3, 11.3, 0.6, lesson["chapter"], font_size=18, color=C_WHITE, align=PP_ALIGN.CENTER)
    add_textbox(slide, 1, 4.9, 11.3, 0.8, lesson["title"], font_size=28, bold=True, color=C_WHITE, align=PP_ALIGN.CENTER)
    add_textbox(slide, 1, 6.85, 11.3, 0.5, f"本课时共 {len(lesson['questions'])} 道题（第1页：第1-6题  |  第2页：第7-12题）", font_size=16, color=ch_color, align=PP_ALIGN.CENTER)

    # ---- 题目页1（第1-6题） ----
    slide = prs.slides.add_slide(blank_layout)
    add_rect(slide, 0, 0, 13.333, 0.8, ch_color)
    add_textbox(slide, 0.5, 0.1, 10, 0.55, f"课时{lesson['num']}：{lesson['title']}（1/2）", font_size=18, bold=True, color=C_WHITE)
    add_textbox(slide, 10.5, 0.1, 2.5, 0.55, "第1-6题", font_size=14, color=C_WHITE, align=PP_ALIGN.RIGHT)

    y_pos = 1.0
    for qi in range(6):
        q = lesson["questions"][qi]
        tc = type_colors.get(q["type"], C_BLUE)
        tbg = type_bg.get(q["type"], C_LIGHT_BLUE)

        # 题号色块
        add_rect(slide, 0.5, y_pos, 0.4, 0.35, tc)
        add_textbox(slide, 0.5, y_pos + 0.01, 0.4, 0.33, str(qi + 1), font_size=14, bold=True, color=C_WHITE, align=PP_ALIGN.CENTER)

        # 题型标签
        add_rect(slide, 0.95, y_pos, 0.7, 0.35, tbg, line_color=tc)
        add_textbox(slide, 0.95, y_pos + 0.01, 0.7, 0.33, f"[{q['type']}]", font_size=11, bold=True, color=tc, align=PP_ALIGN.CENTER)

        # 题干
        add_textbox(slide, 1.75, y_pos + 0.01, 11.2, 0.35, q["q"], font_size=13, color=C_BLACK)

        y_pos += 0.42

        # 选项
        if "options" in q:
            if q["type"] == "选择":
                for oi, opt in enumerate(q["options"]):
                    col = oi % 2
                    row = oi // 2
                    ox = 1.75 + col * 5.5
                    oy = y_pos + row * 0.28
                    add_textbox(slide, ox, oy, 5, 0.25, opt, font_size=12, color=C_BLACK)
                y_pos += 0.62
            elif q["type"] == "判断":
                add_textbox(slide, 1.75, y_pos, 5, 0.25, "A. 正确        B. 错误", font_size=12, color=C_BLACK)
                y_pos += 0.3
        elif q["type"] == "填空":
            add_textbox(slide, 1.75, y_pos, 10, 0.25, "答：________________", font_size=12, color=C_GRAY)
            y_pos += 0.3

        y_pos += 0.06

    # 底部
    add_textbox(slide, 0.5, 7.05, 12, 0.35, "姓名：___________  班级：___________  得分：______", font_size=12, color=C_GRAY)

    # ---- 题目页2（第7-12题） ----
    slide = prs.slides.add_slide(blank_layout)
    add_rect(slide, 0, 0, 13.333, 0.8, ch_color)
    add_textbox(slide, 0.5, 0.1, 10, 0.55, f"课时{lesson['num']}：{lesson['title']}（2/2）", font_size=18, bold=True, color=C_WHITE)
    add_textbox(slide, 10.5, 0.1, 2.5, 0.55, "第7-12题", font_size=14, color=C_WHITE, align=PP_ALIGN.RIGHT)

    y_pos = 1.0
    for qi in range(6, 12):
        q = lesson["questions"][qi]
        tc = type_colors.get(q["type"], C_BLUE)
        tbg = type_bg.get(q["type"], C_LIGHT_BLUE)

        add_rect(slide, 0.5, y_pos, 0.4, 0.35, tc)
        add_textbox(slide, 0.5, y_pos + 0.01, 0.4, 0.33, str(qi + 1), font_size=14, bold=True, color=C_WHITE, align=PP_ALIGN.CENTER)

        add_rect(slide, 0.95, y_pos, 0.7, 0.35, tbg, line_color=tc)
        add_textbox(slide, 0.95, y_pos + 0.01, 0.7, 0.33, f"[{q['type']}]", font_size=11, bold=True, color=tc, align=PP_ALIGN.CENTER)

        add_textbox(slide, 1.75, y_pos + 0.01, 11.2, 0.35, q["q"], font_size=13, color=C_BLACK)

        y_pos += 0.42

        if "options" in q:
            if q["type"] == "选择":
                for oi, opt in enumerate(q["options"]):
                    col = oi % 2
                    row = oi // 2
                    ox = 1.75 + col * 5.5
                    oy = y_pos + row * 0.28
                    add_textbox(slide, ox, oy, 5, 0.25, opt, font_size=12, color=C_BLACK)
                y_pos += 0.62
            elif q["type"] == "判断":
                add_textbox(slide, 1.75, y_pos, 5, 0.25, "A. 正确        B. 错误", font_size=12, color=C_BLACK)
                y_pos += 0.3
        elif q["type"] == "填空":
            add_textbox(slide, 1.75, y_pos, 10, 0.25, "答：________________", font_size=12, color=C_GRAY)
            y_pos += 0.3

        y_pos += 0.06

    add_textbox(slide, 0.5, 7.05, 12, 0.35, "姓名：___________  班级：___________  得分：______", font_size=12, color=C_GRAY)

# 保存PPT
ppt_path = r"D:\WorkBuddyData\2026-08-11-07-48-22\云计算基础_前12课时课堂习题_扩充版.pptx"
prs.save(ppt_path)
print(f"PPT saved: {ppt_path}")
print(f"Total slides: {len(prs.slides)}")
total_q = sum(len(l["questions"]) for l in lessons)
print(f"Total questions: {total_q}")

# ========== 答案文档（HTML） ==========
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
  .q-item { margin: 10px 0; padding: 10px 15px; border: 1px solid #e0e0e0; border-radius: 6px; }
  .q-item:hover { background: #f8f9fa; }
  .q-header { display: flex; align-items: center; gap: 8px; margin-bottom: 4px; }
  .q-num { background: #1a73e8; color: #fff; width: 24px; height: 24px; border-radius: 50%; display: inline-flex; align-items: center; justify-content: center; font-size: 12px; font-weight: bold; flex-shrink: 0; }
  .q-type { display: inline-block; padding: 2px 8px; border-radius: 10px; font-size: 11px; font-weight: bold; }
  .type-select { background: #e3f2fd; color: #1565c0; }
  .type-true { background: #fff3e0; color: #e65100; }
  .type-fill { background: #e8f5e9; color: #2e7d32; }
  .q-text { font-size: 13px; margin-bottom: 4px; padding-left: 32px; }
  .q-options { font-size: 12px; color: #666; padding-left: 32px; margin-bottom: 4px; }
  .q-answer { font-size: 13px; padding: 4px 10px; background: #fffde7; border: 1px solid #fff59d; border-radius: 4px; margin-left: 32px; margin-bottom: 4px; display: inline-block; }
  .q-answer strong { color: #f57f17; }
  .q-explain { font-size: 12px; color: #555; padding: 4px 10px; background: #f1f8e9; border-radius: 4px; margin-left: 32px; }
  .q-explain::before { content: "解析："; font-weight: bold; color: #4caf50; }
  .answer-summary { margin: 20px 0; }
  .summary-table { width: 100%; border-collapse: collapse; font-size: 13px; }
  .summary-table th { background: #1a73e8; color: #fff; padding: 8px; text-align: center; }
  .summary-table td { border: 1px solid #e0e0e0; padding: 6px 8px; text-align: center; }
  .summary-table tr:nth-child(even) { background: #f8f9fa; }
  .summary-table .ans { font-weight: bold; color: #1a73e8; }
  .page-break { page-break-before: always; }
  @media print { .container { box-shadow: none; } .q-item { page-break-inside: avoid; } }
</style>
</head>
<body>
<div class="container">
<h1>云计算基础 前12课时 课堂习题 - 答案与解析</h1>
<p class="subtitle">教材：云计算基础技术与应用(第2版)(微课版) | 易海博 池瑞楠 主编 | 人民邮电出版社<br>
配套文档：《云计算基础_前12课时课堂习题.pptx》| 共12课时 x 12题 = 144题</p>

<h3>快速答案总表</h3>
<div class="answer-summary">
<table class="summary-table">
<tr><th>课时</th><th>课题</th><th>1</th><th>2</th><th>3</th><th>4</th><th>5</th><th>6</th><th>7</th><th>8</th><th>9</th><th>10</th><th>11</th><th>12</th></tr>
"""

# 快速答案总表
for lesson in lessons:
    ch_num = lesson["chapter"][1]
    html += f'<tr><td>{lesson["num"]}</td><td style="text-align:left;font-size:12px;">{lesson["title"]}</td>'
    for q in lesson["questions"]:
        ans = q["answer"]
        if len(ans) > 8:
            ans = ans[:7] + ".."
        html += f'<td class="ans">{ans}</td>'
    html += '</tr>\n'

html += "</table>\n</div>\n"

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
