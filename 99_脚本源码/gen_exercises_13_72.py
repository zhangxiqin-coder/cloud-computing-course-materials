# -*- coding: utf-8 -*-
"""
生成第13-72课时课堂习题PPT + 答案文档（每课时12题，共60课时720题）
"""
import sys
sys.stdout.reconfigure(encoding='utf-8')

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

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

# ========== 题目数据 ==========
lessons = [
    # ===== 第二单元 课时13-24 =====
    {
        "num": 13, "chapter": "第5章 云计算安全", "title": "云计算安全概述",
        "questions": [
            {"type": "选择", "q": "信息安全三要素CIA中，C代表什么？", "options": ["A. 完整性", "B. 机密性", "C. 可用性", "D. 可控性"], "answer": "B", "explain": "CIA = Confidentiality(机密性) + Integrity(完整性) + Availability(可用性)。"},
            {"type": "选择", "q": "以下哪个场景属于机密性被破坏？", "options": ["A. 网站无法访问", "B. 文件被别人偷偷看了", "C. 文件被篡改了内容", "D. 服务器断电了"], "answer": "B", "explain": "机密性指信息不被未授权者获取，文件被偷看属于机密性被破坏。"},
            {"type": "选择", "q": "以下哪个场景属于可用性被破坏？", "options": ["A. 网站被DDoS攻击导致无法访问", "B. 密码被泄露", "C. 文件内容被修改", "D. 数据被加密勒索"], "answer": "A", "explain": "可用性指信息系统可被授权者随时访问和使用，DDoS攻击导致服务不可用。"},
            {"type": "选择", "q": "云安全责任共担模型中，云厂商负责什么？", "options": ["A. 用户数据的安全", "B. 用户应用的代码安全", "C. 基础设施的安全", "D. 用户密码的管理"], "answer": "C", "explain": "云厂商负责底层基础设施（物理服务器、网络、虚拟化层）的安全。"},
            {"type": "选择", "q": "云安全责任共担模型可以用什么生活场景类比？", "options": ["A. 买房", "B. 租房", "C. 住酒店", "D. 露营"], "answer": "B", "explain": "租房：房东负责房屋结构安全，租客负责自己财物安全，类似责任共担。"},
            {"type": "选择", "q": "以下哪个不是云计算面临的独特安全挑战？", "options": ["A. 多租户数据隔离", "B. 数据存储在第三方", "C. API接口风险", "D. 键盘进水损坏"], "answer": "D", "explain": "键盘进水是物理硬件损坏，不属于云计算特有的安全挑战。"},
            {"type": "判断", "q": "信息安全三要素中的I代表Integrity，即完整性。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "I = Integrity，指信息在存储、传输过程中不被篡改。"},
            {"type": "判断", "q": "在云安全责任共担模型中，用户不需要负责任何安全工作。", "options": ["A. 正确", "B. 错误"], "answer": "B", "explain": "用户需负责数据安全、应用安全、访问控制等，云厂商只负责基础设施。"},
            {"type": "判断", "q": "多租户是云计算的独特安全挑战之一。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "多租户共享物理资源，需确保不同租户间数据隔离。"},
            {"type": "判断", "q": "云计算安全只需要关注技术手段，不需要关注管理和法律。", "options": ["A. 正确", "B. 错误"], "answer": "B", "explain": "云安全需要技术、管理、法律多维度协同保障。"},
            {"type": "填空", "q": "信息安全三要素缩写为____，分别代表机密性、完整性、可用性。", "answer": "CIA", "explain": "CIA = Confidentiality, Integrity, Availability。"},
            {"type": "填空", "q": "云安全____模型是指云厂商和用户各负责一部分安全工作。", "answer": "责任共担", "explain": "责任共担模型(Shared Responsibility Model)明确划分了双方的安全职责。"},
        ]
    },
    {
        "num": 14, "chapter": "第5章 云计算安全", "title": "密码学基础",
        "questions": [
            {"type": "选择", "q": "未经加密的原始数据称为？", "options": ["A. 密文", "B. 明文", "C. 密钥", "D. 哈希值"], "answer": "B", "explain": "明文(Plaintext)是原始的、未加密的数据。"},
            {"type": "选择", "q": "将明文变为密文的过程称为？", "options": ["A. 解密", "B. 加密", "C. 哈希", "D. 签名"], "answer": "B", "explain": "加密(Encryption)是将明文通过算法和密钥变为密文的过程。"},
            {"type": "选择", "q": "凯撒密码属于哪种密码？", "options": ["A. 现代密码", "B. 对称加密", "C. 古典密码", "D. 非对称加密"], "answer": "C", "explain": "凯撒密码是古老的替换密码，属于古典密码。"},
            {"type": "选择", "q": "密钥的作用是什么？", "options": ["A. 加快加密速度", "B. 控制加密和解密的参数", "C. 存储数据", "D. 传输数据"], "answer": "B", "explain": "密钥是控制加密和解密算法运行的参数，相同算法不同密钥产生不同密文。"},
            {"type": "选择", "q": "以下哪个是古典密码的典型方法？", "options": ["A. AES", "B. RSA", "C. 替换密码", "D. SHA-256"], "answer": "C", "explain": "替换密码（如凯撒密码）是古典密码的代表，AES/RSA/SHA都是现代密码。"},
            {"type": "选择", "q": "现代密码学与古典密码学的主要区别是？", "options": ["A. 现代密码用电脑", "B. 现代密码基于数学难题，密钥而非算法保密", "C. 古典密码更安全", "D. 现代密码不需要密钥"], "answer": "B", "explain": "柯克霍夫原则：现代密码的安全性基于密钥而非算法的保密。"},
            {"type": "判断", "q": "密文是经过加密处理后的数据。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "密文(Ciphertext)是明文经过加密算法和密钥处理后得到的结果。"},
            {"type": "判断", "q": "凯撒密码的安全性很高，现代仍然广泛使用。", "options": ["A. 正确", "B. 错误"], "answer": "B", "explain": "凯撒密码只有25种可能，极易破解，现代不再用于实际安全场景。"},
            {"type": "判断", "q": "加密和解密使用的是相同的算法，但可能使用不同的密钥。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "对称加密加解密用同一密钥，非对称加密用不同密钥（公钥/私钥）。"},
            {"type": "判断", "q": "密码学中的\"密码\"和平时登录账号的\"密码\"是同一个概念。", "options": ["A. 正确", "B. 错误"], "answer": "B", "explain": "密码学中的密码(cipher)指加密算法，登录密码(password)指身份验证口令。"},
            {"type": "填空", "q": "加密过程：____ + 加密算法 + 密钥 = 密文。", "answer": "明文", "explain": "明文经过加密算法和密钥处理后变为密文。"},
            {"type": "填空", "q": "凯撒密码是一种____密码（填两个字，指古代的密码类型）。", "answer": "古典", "explain": "凯撒密码属于古典密码，是简单的字母替换加密。"},
        ]
    },
    {
        "num": 15, "chapter": "第5章 云计算安全", "title": "对称加密算法（DES与AES）",
        "questions": [
            {"type": "选择", "q": "对称加密的特点是？", "options": ["A. 加密和解密使用不同密钥", "B. 加密和解密使用相同密钥", "C. 不需要密钥", "D. 使用公钥加密"], "answer": "B", "explain": "对称加密：发送方和接收方使用同一个密钥进行加解密。"},
            {"type": "选择", "q": "DES算法的密钥长度是多少位？", "options": ["A. 56位", "B. 128位", "C. 256位", "D. 64位"], "answer": "A", "explain": "DES有效密钥长度为56位（总长64位，其中8位为校验位），已不够安全。"},
            {"type": "选择", "q": "以下哪个算法是对称加密算法？", "options": ["A. RSA", "B. AES", "C. SHA-256", "D. MD5"], "answer": "B", "explain": "AES是对称加密算法，RSA是非对称加密，SHA/MD5是哈希算法。"},
            {"type": "选择", "q": "AES相比DES的优势是？", "options": ["A. 速度更慢", "B. 密钥更长更安全", "C. 不需要密钥", "D. 只能加密小文件"], "answer": "B", "explain": "AES支持128/192/256位密钥，安全性远超56位的DES。"},
            {"type": "选择", "q": "对称加密的主要缺点是？", "options": ["A. 速度太慢", "B. 密钥分发困难", "C. 加密结果太长", "D. 不支持大数据"], "answer": "B", "explain": "双方需安全地共享同一密钥，密钥分发是对称加密的核心难题。"},
            {"type": "选择", "q": "以下哪个生活场景可以类比对称加密？", "options": ["A. 一把钥匙锁箱子和开箱子", "B. 信箱投信口和取信口", "C. 公告栏贴通知", "D. 身份证验证"], "answer": "A", "explain": "对称加密就像一把钥匙既锁又开，加解密用同一把\"钥匙\"。"},
            {"type": "判断", "q": "DES算法由于密钥太短，目前已不推荐使用。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "DES 56位密钥可被暴力破解，已被AES取代。"},
            {"type": "判断", "q": "对称加密的加解密速度通常比非对称加密快。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "对称加密算法计算量小，速度快，适合大数据量加密。"},
            {"type": "判断", "q": "AES是目前广泛使用的对称加密标准。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "AES于2001年被NIST采纳为加密标准，广泛应用于各领域。"},
            {"type": "判断", "q": "对称加密中，通信双方各自使用不同的密钥。", "options": ["A. 正确", "B. 错误"], "answer": "B", "explain": "对称加密双方使用相同的密钥，这是\"对称\"的含义。"},
            {"type": "填空", "q": "对称加密中，加密和解密使用____的密钥。", "answer": "相同", "explain": "对称 = 加解密用同一密钥。"},
            {"type": "填空", "q": "目前广泛使用的对称加密标准是____（填三个字母）。", "answer": "AES", "explain": "AES(Advanced Encryption Standard)取代了DES成为新标准。"},
        ]
    },
    {
        "num": 16, "chapter": "第5章 云计算安全", "title": "非对称加密算法（RSA）",
        "questions": [
            {"type": "选择", "q": "非对称加密使用几个密钥？", "options": ["A. 0个", "B. 1个", "C. 2个", "D. 4个"], "answer": "C", "explain": "非对称加密使用一对密钥：公钥和私钥。"},
            {"type": "选择", "q": "在非对称加密中，公钥的作用是？", "options": ["A. 解密", "B. 加密", "C. 签名", "D. 删除"], "answer": "B", "explain": "公钥用于加密，私钥用于解密；公钥可以公开给任何人。"},
            {"type": "选择", "q": "以下哪个是非对称加密算法？", "options": ["A. AES", "B. DES", "C. RSA", "D. MD5"], "answer": "C", "explain": "RSA是最著名的非对称加密算法，AES/DES是对称加密，MD5是哈希。"},
            {"type": "选择", "q": "非对称加密解决了对称加密的什么问题？", "options": ["A. 速度太慢", "B. 密钥分发困难", "C. 加密不够安全", "D. 不支持大文件"], "answer": "B", "explain": "非对称加密无需事先共享密钥，公钥可以公开，解决了密钥分发难题。"},
            {"type": "选择", "q": "RSA算法的安全性基于什么数学难题？", "options": ["A. 大整数分解困难", "B. 线性方程求解", "C. 排序问题", "D. 图论问题"], "answer": "A", "explain": "RSA安全性基于大整数分解难题：将两个大质数的乘积分解回原始质数极其困难。"},
            {"type": "选择", "q": "以下哪个生活场景可以类比非对称加密？", "options": ["A. 一把钥匙锁开箱子", "B. 信箱：投信口公开，取信钥匙私有", "C. 公告栏贴通知", "D. 对讲机通话"], "answer": "B", "explain": "信箱的投信口（公钥）任何人都可以用，但取信需要钥匙（私钥）。"},
            {"type": "判断", "q": "非对称加密中，公钥可以公开给任何人。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "公钥(public key)就是设计为公开的，即使被别人知道也不影响安全。"},
            {"type": "判断", "q": "非对称加密的加解密速度比对称加密快。", "options": ["A. 正确", "B. 错误"], "answer": "B", "explain": "非对称加密计算量大，速度远慢于对称加密，通常用于加密小数据或交换密钥。"},
            {"type": "判断", "q": "RSA算法可以同时用于加密和数字签名。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "RSA私钥签名、公钥验签；公钥加密、私钥解密，两个方向都可以用。"},
            {"type": "判断", "q": "非对称加密中，公钥和私钥是相同的。", "options": ["A. 正确", "B. 错误"], "answer": "B", "explain": "公钥和私钥是数学相关但不同的一对密钥，这正是\"非对称\"的含义。"},
            {"type": "填空", "q": "非对称加密使用一对密钥：____钥和私钥。", "answer": "公", "explain": "公钥(public key)可公开，私钥(private key)需保密。"},
            {"type": "填空", "q": "RSA算法的安全性基于大整数____难题。", "answer": "分解", "explain": "大整数分解：将大合数分解为质因数的乘积在计算上极其困难。"},
        ]
    },
    {
        "num": 17, "chapter": "第5章 云计算安全", "title": "哈希算法与数字签名",
        "questions": [
            {"type": "选择", "q": "哈希算法的特点是？", "options": ["A. 可逆的", "B. 不可逆的（单向）", "C. 需要密钥", "D. 产生不同长度输出"], "answer": "B", "explain": "哈希是单向函数，无法从哈希值反推出原始数据。"},
            {"type": "选择", "q": "以下哪个是哈希算法？", "options": ["A. AES", "B. RSA", "C. SHA-256", "D. DES"], "answer": "C", "explain": "SHA-256是哈希算法，AES/DES是对称加密，RSA是非对称加密。"},
            {"type": "选择", "q": "MD5算法的输出长度是多少位？", "options": ["A. 64位", "B. 128位", "C. 256位", "D. 512位"], "answer": "B", "explain": "MD5产生128位(16字节)的哈希值。"},
            {"type": "选择", "q": "数字签名的作用是？", "options": ["A. 加密数据", "B. 验证身份和防篡改", "C. 压缩文件", "D. 传输数据"], "answer": "B", "explain": "数字签名用于验证发送者身份（不可否认性）和确认数据未被篡改。"},
            {"type": "选择", "q": "数字签名使用什么密钥进行签名？", "options": ["A. 公钥", "B. 私钥", "C. 对称密钥", "D. 不需要密钥"], "answer": "B", "explain": "发送方用自己的私钥签名，接收方用发送方的公钥验签。"},
            {"type": "选择", "q": "哈希算法在密码存储中的应用是？", "options": ["A. 直接存储明文密码", "B. 存储密码的哈希值", "C. 加密密码后存储", "D. 不存储任何信息"], "answer": "B", "explain": "存储密码的哈希值而非明文，即使数据库泄露也无法还原密码。"},
            {"type": "判断", "q": "相同的输入经过同一哈希算法必定得到相同的输出。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "哈希函数的确定性：相同输入永远得到相同输出。"},
            {"type": "判断", "q": "可以从MD5哈希值反推出原始数据。", "options": ["A. 正确", "B. 错误"], "answer": "B", "explain": "哈希是单向不可逆的，无法从哈希值还原原始数据。"},
            {"type": "判断", "q": "数字签名可以防止发送者否认发送过该消息。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "数字签名提供不可否认性，因为只有私钥持有者才能生成有效签名。"},
            {"type": "判断", "q": "MD5目前仍被推荐用于安全要求高的场景。", "options": ["A. 正确", "B. 错误"], "answer": "B", "explain": "MD5已被发现存在碰撞漏洞，安全场景推荐使用SHA-256。"},
            {"type": "填空", "q": "哈希算法是____的，无法从输出反推输入（填两个字）。", "answer": "不可逆", "explain": "哈希是单向函数，不可逆是其核心特性。"},
            {"type": "填空", "q": "数字签名用____钥签名，用公钥验签。", "answer": "私", "explain": "私钥签名保证只有持有者能签，公钥验签让任何人都能验证。"},
        ]
    },
    {
        "num": 18, "chapter": "第5章 云计算安全", "title": "数字证书与PKI",
        "questions": [
            {"type": "选择", "q": "数字证书的作用是？", "options": ["A. 加密数据", "B. 证明公钥属于特定实体", "C. 压缩文件", "D. 传输数据"], "answer": "B", "explain": "数字证书将公钥与持有者身份绑定，由可信第三方(CA)签发。"},
            {"type": "选择", "q": "CA机构的全称是？", "options": ["A. 计算机协会", "B. 证书授权机构", "C. 云计算联盟", "D. 通信管理局"], "answer": "B", "explain": "CA = Certificate Authority，证书授权机构，负责签发和管理数字证书。"},
            {"type": "选择", "q": "PKI的中文含义是？", "options": ["A. 公钥基础设施", "B. 私钥基础设施", "C. 密钥管理接口", "D. 平台密钥接口"], "answer": "A", "explain": "PKI = Public Key Infrastructure，公钥基础设施。"},
            {"type": "选择", "q": "数字证书中包含什么信息？", "options": ["A. 只有公钥", "B. 只有私钥", "C. 持有者信息+公钥+CA签名", "D. 只有CA的地址"], "answer": "C", "explain": "数字证书包含持有者信息、公钥、有效期、CA的数字签名等。"},
            {"type": "选择", "q": "以下哪个是常见的证书格式？", "options": ["A. .docx", "B. .pem / .crt", "C. .mp4", "D. .zip"], "answer": "B", "explain": "PEM和CRT是常见的X.509数字证书文件格式。"},
            {"type": "选择", "q": "浏览器中显示\"安全\"锁标意味着什么？", "options": ["A. 网站没有病毒", "B. 网站使用了有效的HTTPS证书", "C. 网站速度很快", "D. 网站内容真实"], "answer": "B", "explain": "锁标表示该网站使用了有效的SSL/TLS证书，连接是加密的。"},
            {"type": "判断", "q": "数字证书由网站自己签发，不需要第三方机构。", "options": ["A. 正确", "B. 错误"], "answer": "B", "explain": "数字证书需由受信任的CA机构签发，否则浏览器会提示不安全。"},
            {"type": "判断", "q": "PKI是一套管理公钥和数字证书的体系。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "PKI包括CA、证书库、密钥管理等组件，构成完整的公钥管理基础设施。"},
            {"type": "判断", "q": "数字证书有有效期，过期后需要更新。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "证书通常有效期为1-2年，过期后需向CA重新申请。"},
            {"type": "判断", "q": "数字证书可以证明网站的身份真实性。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "CA在签发证书前会验证申请者身份，证书起到了身份证明的作用。"},
            {"type": "填空", "q": "数字证书由____机构签发（填两个字母）。", "answer": "CA", "explain": "CA(Certificate Authority)是受信任的第三方证书授权机构。"},
            {"type": "填空", "q": "PKI的中文全称是公钥____。", "answer": "基础设施", "explain": "PKI = Public Key Infrastructure，公钥基础设施。"},
        ]
    },
    {
        "num": 19, "chapter": "第5章 云计算安全", "title": "SSL/TLS与HTTPS",
        "questions": [
            {"type": "选择", "q": "HTTPS中的S代表什么？", "options": ["A. Safe", "B. Secure", "C. SSL/TLS", "D. Server"], "answer": "C", "explain": "HTTPS = HTTP + SSL/TLS，S指SSL/TLS加密层。"},
            {"type": "选择", "q": "HTTPS使用的默认端口是？", "options": ["A. 80", "B. 443", "C. 8080", "D. 22"], "answer": "B", "explain": "HTTP默认端口80，HTTPS默认端口443。"},
            {"type": "选择", "q": "SSL/TLS协议的主要作用是？", "options": ["A. 加速网页加载", "B. 提供加密通信和身份验证", "C. 压缩网页数据", "D. 过滤广告"], "answer": "B", "explain": "SSL/TLS提供数据加密、身份验证和数据完整性保护。"},
            {"type": "选择", "q": "HTTPS握手过程中，浏览器首先做什么？", "options": ["A. 发送加密数据", "B. 请求服务器的数字证书", "C. 发送私钥", "D. 发送密码"], "answer": "B", "explain": "握手第一步：客户端请求并接收服务器的数字证书，验证证书有效性。"},
            {"type": "选择", "q": "HTTP和HTTPS的主要区别是？", "options": ["A. HTTPS更快", "B. HTTPS加密传输数据", "C. HTTP更安全", "D. 两者完全相同"], "answer": "B", "explain": "HTTPS在HTTP基础上增加了SSL/TLS加密层，数据传输是加密的。"},
            {"type": "选择", "q": "以下哪个不是HTTPS的优势？", "options": ["A. 数据加密传输", "B. 防止中间人窃听", "C. 提高网站加载速度", "D. 验证网站身份"], "answer": "C", "explain": "HTTPS因加密握手会略微增加延迟，不是为了提速。"},
            {"type": "判断", "q": "HTTPS可以对传输的数据进行加密。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "HTTPS通过SSL/TLS加密HTTP报文，防止数据被窃听。"},
            {"type": "判断", "q": "使用HTTPS的网站一定不会被盗号。", "options": ["A. 正确", "B. 错误"], "answer": "B", "explain": "HTTPS只保证传输安全，不能防止用户密码太弱或钓鱼网站等攻击。"},
            {"type": "判断", "q": "SSL和TLS是同一个协议的不同名称。", "options": ["A. 正确", "B. 错误"], "answer": "B", "explain": "TLS是SSL的升级版本，SSL已不再安全，现在使用的是TLS。"},
            {"type": "判断", "q": "网站启用HTTPS需要安装SSL/TLS证书。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "网站需向CA申请并安装数字证书才能启用HTTPS。"},
            {"type": "填空", "q": "HTTPS = HTTP + ____/TLS。", "answer": "SSL", "explain": "HTTPS在HTTP基础上增加SSL/TLS加密层。"},
            {"type": "填空", "q": "HTTPS默认使用端口____。", "answer": "443", "explain": "HTTP默认80，HTTPS默认443。"},
        ]
    },
    {
        "num": 20, "chapter": "第5章 云计算安全", "title": "云安全防护技术",
        "questions": [
            {"type": "选择", "q": "防火墙的主要功能是？", "options": ["A. 查杀病毒", "B. 控制网络流量进出", "C. 加密数据", "D. 备份数据"], "answer": "B", "explain": "防火墙根据安全策略控制进出网络的流量，允许或阻断通信。"},
            {"type": "选择", "q": "IDS的全称是？", "options": ["A. 入侵检测系统", "B. 入侵防御系统", "C. 身份认证系统", "D. 数据加密系统"], "answer": "A", "explain": "IDS = Intrusion Detection System，入侵检测系统，负责检测并报警。"},
            {"type": "选择", "q": "IDS和IPS的区别是？", "options": ["A. IDS只检测报警，IPS可以主动阻断", "B. IDS比IPS更先进", "C. IPS只检测不阻断", "D. 两者完全相同"], "answer": "A", "explain": "IDS(Intrusion Detection)只检测报警，IPS(Intrusion Prevention)可以主动阻断攻击。"},
            {"type": "选择", "q": "云环境中的访问控制主要管理什么？", "options": ["A. 服务器温度", "B. 用户对资源的访问权限", "C. 网络带宽", "D. 存储容量"], "answer": "B", "explain": "访问控制管理用户身份认证和权限分配，确保只有授权用户能访问相应资源。"},
            {"type": "选择", "q": "以下哪个不是云数据安全措施？", "options": ["A. 数据加密存储", "B. 数据备份", "C. 访问控制", "D. 增加CPU频率"], "answer": "D", "explain": "增加CPU频率是性能优化，与数据安全无关。"},
            {"type": "选择", "q": "DDoS攻击的目的是？", "options": ["A. 窃取数据", "B. 篡改网页", "C. 使服务器过载无法服务", "D. 加密文件勒索"], "answer": "C", "explain": "DDoS通过大量请求淹没服务器，使其无法为正常用户提供服务。"},
            {"type": "判断", "q": "防火墙可以完全阻止所有网络攻击。", "options": ["A. 正确", "B. 错误"], "answer": "B", "explain": "防火墙只是安全防护的一环，无法阻止所有攻击（如应用层攻击）。"},
            {"type": "判断", "q": "IPS可以主动阻断检测到的攻击行为。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "IPS(Intrusion Prevention System)不仅能检测还能主动阻断攻击。"},
            {"type": "判断", "q": "云环境中的数据加密包括传输加密和存储加密。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "传输加密(如HTTPS)保护数据在传输中安全，存储加密保护静态数据。"},
            {"type": "判断", "q": "多因素认证比单一密码认证更安全。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "多因素认证(密码+手机验证码+指纹等)大大提高了安全性。"},
            {"type": "填空", "q": "____墙是根据安全策略控制网络流量进出的安全设备。", "answer": "防火", "explain": "防火墙(Firewall)是网络安全的第一道防线。"},
            {"type": "填空", "q": "IPS的全称是入侵____系统。", "answer": "防御", "explain": "IPS = Intrusion Prevention System，入侵防御系统。"},
        ]
    },
    {
        "num": 21, "chapter": "第6章 云计算市场", "title": "云计算市场概述",
        "questions": [
            {"type": "选择", "q": "全球云计算市场份额最大的厂商是？", "options": ["A. 阿里云", "B. AWS", "C. 腾讯云", "D. 华为云"], "answer": "B", "explain": "AWS(亚马逊云)是全球最大的公有云服务商，市场份额长期第一。"},
            {"type": "选择", "q": "中国云计算市场份额最大的厂商是？", "options": ["A. 阿里云", "B. AWS", "C. 腾讯云", "D. 华为云"], "answer": "A", "explain": "阿里云在中国公有云市场份额长期排名第一。"},
            {"type": "选择", "q": "以下哪个不是云计算的发展趋势？", "options": ["A. 混合云成为主流", "B. AI与云融合", "C. 云计算被淘汰", "D. 边缘计算兴起"], "answer": "C", "explain": "云计算仍在快速增长，远未被淘汰。"},
            {"type": "选择", "q": "云计算对传统IT的主要影响是？", "options": ["A. 传统IT完全消失", "B. 企业从买硬件转向租服务", "C. 不再需要IT人员", "D. 网络不再重要"], "answer": "B", "explain": "云计算推动IT从资本支出(CAPEX)转向运营支出(OPEX)，从买硬件转向租服务。"},
            {"type": "选择", "q": "以下哪个是云计算市场的特点？", "options": ["A. 市场规模萎缩", "B. 头部厂商集中度高", "C. 进入门槛极低", "D. 完全垄断"], "answer": "B", "explain": "云计算是重资产行业，头部厂商(AWS/阿里云/微软云)集中度高。"},
            {"type": "选择", "q": "边缘计算的核心思想是？", "options": ["A. 在云端处理所有数据", "B. 在数据源头附近处理数据", "C. 不需要计算", "D. 只在终端处理"], "answer": "B", "explain": "边缘计算将计算能力下沉到数据源附近，减少延迟和带宽消耗。"},
            {"type": "判断", "q": "AWS是全球最大的公有云服务商。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "AWS(亚马逊Web服务)自2006年推出以来一直是全球公有云市场领导者。"},
            {"type": "判断", "q": "云计算市场规模正在持续增长。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "全球云计算市场年均增长率在20%以上，仍在快速增长。"},
            {"type": "判断", "q": "混合云正在成为企业上云的主流选择。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "企业兼顾安全和弹性，混合云是最平衡的选择。"},
            {"type": "判断", "q": "边缘计算会完全取代云计算。", "options": ["A. 正确", "B. 错误"], "answer": "B", "explain": "边缘计算与云计算是互补关系，边缘处理实时数据，云端处理大数据。"},
            {"type": "填空", "q": "全球云计算市场份额最大的厂商是____（填三个字母）。", "answer": "AWS", "explain": "AWS = Amazon Web Services，亚马逊云服务。"},
            {"type": "填空", "q": "中国云计算市场份额最大的厂商是____云。", "answer": "阿里", "explain": "阿里云在中国公有云市场长期排名第一。"},
        ]
    },
    {
        "num": 22, "chapter": "第6章 云计算市场", "title": "主流云服务厂商",
        "questions": [
            {"type": "选择", "q": "AWS是哪个公司的云服务？", "options": ["A. 谷歌", "B. 微软", "C. 亚马逊", "D. IBM"], "answer": "C", "explain": "AWS = Amazon Web Services，是亚马逊公司的云服务平台。"},
            {"type": "选择", "q": "Azure是哪个公司的云服务？", "options": ["A. 谷歌", "B. 微软", "C. 亚马逊", "D. 阿里巴巴"], "answer": "B", "explain": "Microsoft Azure是微软的公有云平台。"},
            {"type": "选择", "q": "以下哪个是腾讯云的特色优势领域？", "options": ["A. 电商", "B. 音视频和游戏", "C. 企业管理", "D. 操作系统"], "answer": "B", "explain": "腾讯云在音视频、游戏、社交等领域有天然优势。"},
            {"type": "选择", "q": "华为云的差异化优势主要在？", "options": ["A. 社交", "B. 硬件能力和政企服务", "C. 电商", "D. 搜索"], "answer": "B", "explain": "华为云在硬件(芯片/服务器)能力和政企客户服务方面有优势。"},
            {"type": "选择", "q": "阿里云的旗舰产品ECS属于什么服务？", "options": ["A. SaaS", "B. PaaS", "C. IaaS", "D. DaaS"], "answer": "C", "explain": "ECS(Elastic Compute Service)是弹性计算服务，属于IaaS。"},
            {"type": "选择", "q": "Google Cloud Platform (GCP)的优势领域是？", "options": ["A. 电商", "B. AI和大数据", "C. 游戏", "D. 政务"], "answer": "B", "explain": "谷歌在AI(TensorFlow)和大数据分析方面有领先优势。"},
            {"type": "判断", "q": "AWS是最早提供大规模公有云服务的厂商。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "AWS于2006年推出EC2，开创了公有云服务模式。"},
            {"type": "判断", "q": "阿里云只在中国市场运营。", "options": ["A. 正确", "B. 错误"], "answer": "B", "explain": "阿里云在东南亚、欧洲、中东等地区都有数据中心和业务。"},
            {"type": "判断", "q": "不同的云厂商提供的IaaS/PaaS/SaaS服务本质上是类似的。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "核心服务类型一致(虚拟机、存储、数据库等)，差异在于价格、性能和特色服务。"},
            {"type": "判断", "q": "华为云在政企领域有较强的市场竞争力。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "华为在政企市场深耕多年，华为云继承了这一优势。"},
            {"type": "填空", "q": "Azure是____公司的云服务平台。", "answer": "微软", "explain": "Microsoft Azure是微软的公有云。"},
            {"type": "填空", "q": "阿里云的弹性计算服务英文缩写是____。", "answer": "ECS", "explain": "ECS = Elastic Compute Service。"},
        ]
    },
    {
        "num": 23, "chapter": "第6章 云计算市场", "title": "企业上云",
        "questions": [
            {"type": "选择", "q": "企业上云的主要动机是？", "options": ["A. 跟风", "B. 降低成本、提高灵活性", "C. 淘汰所有IT人员", "D. 不需要网络"], "answer": "B", "explain": "企业上云可降低IT投入成本、提高业务灵活性和弹性扩展能力。"},
            {"type": "选择", "q": "云迁移中\"重新托管\"（Lift and Shift）是指？", "options": ["A. 重新开发应用", "B. 直接把应用搬到云上不做修改", "C. 修改应用架构", "D. 放弃原有应用"], "answer": "B", "explain": "Lift and Shift（搬迁上云）是最简单的迁移方式，应用不做修改直接部署到云服务器。"},
            {"type": "选择", "q": "企业上云面临的挑战不包括？", "options": ["A. 数据安全顾虑", "B. 应用兼容性", "C. 员工技能差距", "D. 云厂商免费送服务器"], "answer": "D", "explain": "云厂商不会免费送服务器，企业上云需面对安全、兼容性和技能挑战。"},
            {"type": "选择", "q": "以下哪种企业最适合使用混合云？", "options": ["A. 个人博客", "B. 有敏感数据需自控+弹性需求的电商", "C. 只有三台电脑的小店", "D. 不用网络的企业"], "answer": "B", "explain": "混合云兼顾安全(敏感数据放私有云)和弹性(高峰放公有云)，适合大中型企业。"},
            {"type": "选择", "q": "企业上云的一般步骤是？", "options": ["A. 评估->规划->迁移->优化", "B. 直接迁移->评估->规划", "C. 买服务器->装系统->上云", "D. 迁移->规划->评估->丢弃"], "answer": "A", "explain": "标准流程：评估现状->规划架构->执行迁移->持续优化。"},
            {"type": "选择", "q": "云迁移策略中\"重构\"是指？", "options": ["A. 不修改应用", "B. 重新设计应用架构以充分利用云特性", "C. 删除应用", "D. 只改界面"], "answer": "B", "explain": "重构(Refactor)是改变应用架构以利用云原生特性(如微服务、容器)，成本最高但收益最大。"},
            {"type": "判断", "q": "企业上云后就不需要IT运维人员了。", "options": ["A. 正确", "B. 错误"], "answer": "B", "explain": "上云后仍需IT人员，但角色从硬件维护转向云资源管理和应用运维。"},
            {"type": "判断", "q": "所有企业都应该立即全面上云。", "options": ["A. 正确", "B. 错误"], "answer": "B", "explain": "上云需根据业务需求、安全要求、成本等因素综合评估，不能盲目上云。"},
            {"type": "判断", "q": "混合云可以让企业同时享受公有云的弹性和私有云的安全。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "混合云结合两者的优势，是很多企业的选择。"},
            {"type": "判断", "q": "云迁移过程中数据可能面临安全风险。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "迁移过程中数据传输、格式转换等环节都需做好安全防护。"},
            {"type": "填空", "q": "企业上云最简单的迁移方式叫____迁移（把应用直接搬到云上不改）。", "answer": "直接 / 平移", "explain": "Lift and Shift，也叫直接迁移或平移迁移。"},
            {"type": "填空", "q": "____云结合了公有云的弹性和私有云的安全性。", "answer": "混合", "explain": "混合云是很多企业上云的选择。"},
        ]
    },
    {
        "num": 24, "chapter": "第6章 云计算市场", "title": "云计算行业应用与第二单元复习",
        "questions": [
            {"type": "选择", "q": "以下哪个行业大量使用云计算？", "options": ["A. 电商", "B. 医疗", "C. 教育", "D. 以上都是"], "answer": "D", "explain": "云计算已广泛应用于电商、医疗、教育、金融、政务等几乎所有行业。"},
            {"type": "选择", "q": "在线教育平台使用云计算的主要好处是？", "options": ["A. 不需要老师", "B. 弹性应对高峰流量（如开学季）", "C. 不需要网络", "D. 自动批改作业"], "answer": "B", "explain": "教育平台流量波动大，云计算的弹性伸缩可以应对高峰。"},
            {"type": "选择", "q": "信息安全三要素CIA中A代表什么？", "options": ["A. 认证", "B. 授权", "C. 可用性", "D. 审计"], "answer": "C", "explain": "A = Availability，可用性，指系统可被正常访问使用。"},
            {"type": "选择", "q": "以下哪个是复习题：对称加密和非对称加密的主要区别？", "options": ["A. 速度不同", "B. 对称用同一密钥，非对称用公私钥对", "C. 算法不同", "D. 用途不同"], "answer": "B", "explain": "核心区别：对称加密加解密用同一密钥，非对称用公钥/私钥一对密钥。"},
            {"type": "选择", "q": "复习题：以下哪个不属于PKI的组件？", "options": ["A. CA机构", "B. 数字证书", "C. 证书吊销列表(CRL)", "D. 杀毒软件"], "answer": "D", "explain": "杀毒软件不属于PKI体系。PKI包括CA、证书库、CRL、密钥管理等。"},
            {"type": "选择", "q": "复习题：以下哪个不是云安全责任共担的原则？", "options": ["A. 云厂商负责基础设施安全", "B. 用户负责数据安全", "C. 用户负责物理服务器维护", "D. 双方各负责一部分"], "answer": "C", "explain": "物理服务器由云厂商维护，用户不负责物理硬件。"},
            {"type": "判断", "q": "复习：哈希算法是可逆的，可以从哈希值还原原始数据。", "options": ["A. 正确", "B. 错误"], "answer": "B", "explain": "哈希是单向不可逆的，无法从输出还原输入。"},
            {"type": "判断", "q": "复习：HTTPS使用的默认端口是80。", "options": ["A. 正确", "B. 错误"], "answer": "B", "explain": "HTTP默认80，HTTPS默认443。"},
            {"type": "判断", "q": "复习：AWS是全球最大的公有云服务商。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "AWS全球市场份额第一。"},
            {"type": "判断", "q": "复习：数字签名使用公钥签名，私钥验签。", "options": ["A. 正确", "B. 错误"], "answer": "B", "explain": "数字签名使用私钥签名，公钥验签。"},
            {"type": "填空", "q": "复习：信息安全三要素是机密性、完整性、____性。", "answer": "可用", "explain": "CIA = 机密性(Confidentiality)、完整性(Integrity)、可用性(Availability)。"},
            {"type": "填空", "q": "复习：防火墙控制网络流量____（两个字，指出和进）。", "answer": "进出", "explain": "防火墙根据策略控制进出网络的流量。"},
        ]
    },
    # ===== 第三单元 课时25-36 =====
    {
        "num": 25, "chapter": "第7章 计算机网络", "title": "计算机网络概述",
        "questions": [
            {"type": "选择", "q": "计算机网络最基本的功能是？", "options": ["A. 数据通信和资源共享", "B. 打游戏", "C. 看视频", "D. 打印文件"], "answer": "A", "explain": "计算机网络的核心功能是数据通信和资源共享（硬件、软件、数据）。"},
            {"type": "选择", "q": "OSI参考模型共有几层？", "options": ["A. 4层", "B. 5层", "C. 7层", "D. 9层"], "answer": "C", "explain": "OSI七层模型：物理层、数据链路层、网络层、传输层、会话层、表示层、应用层。"},
            {"type": "选择", "q": "OSI模型中最高层是？", "options": ["A. 物理层", "B. 网络层", "C. 传输层", "D. 应用层"], "answer": "D", "explain": "应用层是OSI第七层，是最高层，直接为用户应用程序提供服务。"},
            {"type": "选择", "q": "以下哪个是局域网(LAN)？", "options": ["A. Internet", "B. 学校机房网络", "C. 跨国公司专线", "D. 5G移动网络"], "answer": "B", "explain": "学校机房网络覆盖范围小，属于局域网(LAN)。"},
            {"type": "选择", "q": "Internet属于哪种网络？", "options": ["A. LAN", "B. MAN", "C. WAN", "D. PAN"], "answer": "C", "explain": "Internet是全球范围的广域网(WAN)。"},
            {"type": "选择", "q": "计算机网络在云计算中的作用是？", "options": ["A. 不重要", "B. 云计算的基础，连接用户和云", "C. 只用来上网", "D. 可有可无"], "answer": "B", "explain": "网络是云计算的命脉，所有云服务都通过网络提供。"},
            {"type": "判断", "q": "OSI七层模型中最底层是物理层。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "物理层是第一层，负责传输比特流（0和1的电信号/光信号）。"},
            {"type": "判断", "q": "计算机网络按覆盖范围可分为局域网和广域网。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "按覆盖范围分为LAN(局域网)、MAN(城域网)、WAN(广域网)。"},
            {"type": "判断", "q": "云计算服务完全不需要网络就能使用。", "options": ["A. 正确", "B. 错误"], "answer": "B", "explain": "云计算依赖网络提供服务，\"广泛网络访问\"是NIST定义的5个特征之一。"},
            {"type": "判断", "q": "OSI模型中每层只与相邻层交互。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "OSI模型中每层只使用下层提供的服务并向上层提供服务。"},
            {"type": "填空", "q": "OSI参考模型共有____层。", "answer": "7", "explain": "物理层、数据链路层、网络层、传输层、会话层、表示层、应用层。"},
            {"type": "填空", "q": "学校机房的网络属于____域网（填一个字）。", "answer": "局", "explain": "覆盖范围小（如一栋楼内）的网络是局域网(LAN)。"},
        ]
    },
    {
        "num": 26, "chapter": "第7章 计算机网络", "title": "TCP/IP协议体系",
        "questions": [
            {"type": "选择", "q": "TCP/IP模型共有几层？", "options": ["A. 4层", "B. 5层", "C. 7层", "D. 3层"], "answer": "A", "explain": "TCP/IP四层模型：网络接口层、网络层、传输层、应用层。"},
            {"type": "选择", "q": "TCP/IP模型中，IP协议工作在哪一层？", "options": ["A. 应用层", "B. 传输层", "C. 网络层", "D. 网络接口层"], "answer": "C", "explain": "IP(Internet Protocol)是网络层协议，负责数据包的路由和转发。"},
            {"type": "选择", "q": "以下哪个是应用层协议？", "options": ["A. IP", "B. TCP", "C. HTTP", "D. Ethernet"], "answer": "C", "explain": "HTTP是应用层协议，IP是网络层，TCP是传输层，Ethernet是网络接口层。"},
            {"type": "选择", "q": "TCP/IP模型中传输层的主要协议有？", "options": ["A. TCP和UDP", "B. HTTP和DNS", "C. IP和ARP", "D. FTP和SMTP"], "answer": "A", "explain": "传输层两大协议：TCP(可靠传输)和UDP(快速传输)。"},
            {"type": "选择", "q": "OSI七层模型和TCP/IP四层模型的关系是？", "options": ["A. 完全不同", "B. TCP/IP是OSI的简化版", "C. OSI是TCP/IP的简化版", "D. 两者完全相同"], "answer": "B", "explain": "TCP/IP模型将OSI的七层合并为四层，是实际使用的标准。"},
            {"type": "选择", "q": "以下哪个协议不属于TCP/IP体系？", "options": ["A. TCP", "B. IP", "C. HTTP", "D. USB"], "answer": "D", "explain": "USB是硬件接口标准，不属于TCP/IP网络协议体系。"},
            {"type": "判断", "q": "TCP/IP是互联网事实上的标准协议。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "TCP/IP是Internet使用的标准协议体系，虽然OSI是理论模型但实际使用TCP/IP。"},
            {"type": "判断", "q": "TCP/IP模型中的应用层对应OSI的上三层。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "TCP/IP应用层合并了OSI的应用层、表示层、会话层。"},
            {"type": "判断", "q": "IP协议是面向连接的可靠协议。", "options": ["A. 正确", "B. 错误"], "answer": "B", "explain": "IP是无连接的、不可靠的协议，可靠传输由TCP保证。"},
            {"type": "判断", "q": "TCP/IP模型的网络接口层对应OSI的物理层和数据链路层。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "TCP/IP网络接口层合并了OSI的物理层和数据链路层。"},
            {"type": "填空", "q": "TCP/IP模型共有____层。", "answer": "4", "explain": "网络接口层、网络层、传输层、应用层。"},
            {"type": "填空", "q": "IP协议工作在TCP/IP模型的____层。", "answer": "网络", "explain": "IP是网络层(Internet Layer)的核心协议。"},
        ]
    },
    {
        "num": 27, "chapter": "第7章 计算机网络", "title": "IP地址与子网掩码",
        "questions": [
            {"type": "选择", "q": "IPv4地址由多少位二进制数组成？", "options": ["A. 16位", "B. 32位", "C. 64位", "D. 128位"], "answer": "B", "explain": "IPv4地址是32位，通常表示为4个8位十进制数(如192.168.1.1)。"},
            {"type": "选择", "q": "IP地址192.168.1.100属于哪一类？", "options": ["A. A类", "B. B类", "C. C类", "D. D类"], "answer": "C", "explain": "C类地址范围192.0.0.0-223.255.255.255，192开头属于C类。"},
            {"type": "选择", "q": "以下哪个是私有IP地址？", "options": ["A. 8.8.8.8", "B. 192.168.1.1", "C. 202.96.128.86", "D. 114.114.114.114"], "answer": "B", "explain": "192.168.x.x是私有地址段，仅在局域网内使用，不能在互联网上路由。"},
            {"type": "选择", "q": "IPv6相比IPv4的主要优势是？", "options": ["A. 速度更快", "B. 地址空间更大", "C. 更安全", "D. 不需要网络"], "answer": "B", "explain": "IPv6是128位地址，地址数量远超IPv4的32位，解决了地址枯竭问题。"},
            {"type": "选择", "q": "子网掩码的作用是？", "options": ["A. 加密数据", "B. 区分IP地址中的网络部分和主机部分", "C. 加速网络", "D. 分配IP地址"], "answer": "B", "explain": "子网掩码与IP地址做AND运算，区分网络号和主机号。"},
            {"type": "选择", "q": "默认情况下C类地址的子网掩码是？", "options": ["A. 255.0.0.0", "B. 255.255.0.0", "C. 255.255.255.0", "D. 255.255.255.255"], "answer": "C", "explain": "C类地址默认子网掩码255.255.255.0，前24位为网络号。"},
            {"type": "判断", "q": "IPv4地址总共可以提供约43亿个唯一地址。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "IPv4是32位，2^32约43亿个地址，已基本分配殆尽。"},
            {"type": "判断", "q": "127.0.0.1是回环地址，用于本机测试。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "127.0.0.1(localhost)是回环地址，数据不离开本机。"},
            {"type": "判断", "q": "私有IP地址可以在互联网上直接路由。", "options": ["A. 正确", "B. 错误"], "answer": "B", "explain": "私有IP地址不能在互联网上路由，需通过NAT转换为公网IP。"},
            {"type": "判断", "q": "IPv6地址是128位二进制数。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "IPv6是128位，通常用8组4位十六进制数表示。"},
            {"type": "填空", "q": "IPv4地址由____位二进制数组成。", "answer": "32", "explain": "32位，分4段，每段8位(0-255)。"},
            {"type": "填空", "q": "192.168.x.x是____IP地址（填两个字）。", "answer": "私有", "explain": "私有地址在局域网内使用，不直接暴露在互联网上。"},
        ]
    },
    {
        "num": 28, "chapter": "第7章 计算机网络", "title": "传输层协议：TCP与UDP",
        "questions": [
            {"type": "选择", "q": "TCP协议的特点是？", "options": ["A. 面向连接、可靠传输", "B. 无连接、不可靠", "C. 只能传小数据", "D. 不需要网络"], "answer": "A", "explain": "TCP(传输控制协议)是面向连接的可靠传输协议。"},
            {"type": "选择", "q": "UDP协议的特点是？", "options": ["A. 面向连接、可靠", "B. 无连接、不可靠但快速", "C. 只能传文本", "D. 需要三次握手"], "answer": "B", "explain": "UDP(用户数据报协议)无连接、不保证可靠但速度快，适合实时通信。"},
            {"type": "选择", "q": "TCP三次握手的目的是？", "options": ["A. 加密数据", "B. 建立可靠连接", "C. 传输数据", "D. 关闭连接"], "answer": "B", "explain": "三次握手在数据传输前建立连接，确保双方都准备好通信。"},
            {"type": "选择", "q": "以下哪个应用适合使用UDP？", "options": ["A. 网页浏览", "B. 文件下载", "C. 视频直播", "D. 电子邮件"], "answer": "C", "explain": "视频直播要求实时性，允许少量丢包，UDP比TCP更适合。"},
            {"type": "选择", "q": "以下哪个应用适合使用TCP？", "options": ["A. 视频直播", "B. 网页浏览(HTTP)", "C. DNS查询", "D. 在线游戏位置同步"], "answer": "B", "explain": "网页浏览需要可靠传输，不能丢数据，使用TCP。"},
            {"type": "选择", "q": "TCP三次握手中第一次发送的标志位是？", "options": ["A. SYN", "B. ACK", "C. FIN", "D. RST"], "answer": "A", "explain": "第一次握手客户端发送SYN(同步)标志位，请求建立连接。"},
            {"type": "判断", "q": "TCP在传输数据前需要先建立连接。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "TCP是面向连接的协议，先三次握手建立连接，再传输数据。"},
            {"type": "判断", "q": "UDP比TCP速度快，因为不需要建立连接。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "UDP无需握手建连，直接发送数据，开销小速度快。"},
            {"type": "判断", "q": "TCP能保证数据按顺序到达且不丢失。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "TCP通过序列号、确认重传等机制保证可靠有序传输。"},
            {"type": "判断", "q": "DNS查询既可以使用TCP也可以使用UDP。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "DNS通常用UDP(查询小数据)，区域传送用TCP(大数据传输)。"},
            {"type": "填空", "q": "TCP建立连接需要____次握手。", "answer": "3", "explain": "SYN -> SYN+ACK -> ACK，三次握手建立连接。"},
            {"type": "填空", "q": "视频直播通常使用____协议传输（填三个字母）。", "answer": "UDP", "explain": "UDP速度快，适合实时性要求高、允许少量丢包的场景。"},
        ]
    },
    {
        "num": 29, "chapter": "第7章 计算机网络", "title": "应用层协议",
        "questions": [
            {"type": "选择", "q": "HTTP协议的默认端口是？", "options": ["A. 21", "B. 22", "C. 80", "D. 443"], "answer": "C", "explain": "HTTP默认端口80，HTTPS默认端口443。"},
            {"type": "选择", "q": "DNS协议的主要功能是？", "options": ["A. 传输文件", "B. 域名解析为IP地址", "C. 发送邮件", "D. 分配IP地址"], "answer": "B", "explain": "DNS(域名系统)将域名(如www.baidu.com)解析为IP地址。"},
            {"type": "选择", "q": "FTP协议用于？", "options": ["A. 网页浏览", "B. 文件传输", "C. 邮件发送", "D. 域名解析"], "answer": "B", "explain": "FTP(文件传输协议)用于在网络上传输文件。"},
            {"type": "选择", "q": "DHCP协议的作用是？", "options": ["A. 解析域名", "B. 自动分配IP地址", "C. 传输文件", "D. 发送邮件"], "answer": "B", "explain": "DHCP(动态主机配置协议)自动为网络设备分配IP地址、子网掩码等。"},
            {"type": "选择", "q": "以下哪个协议用于发送电子邮件？", "options": ["A. HTTP", "B. FTP", "C. SMTP", "D. DNS"], "answer": "C", "explain": "SMTP(简单邮件传输协议)用于发送邮件，POP3/IMAP用于接收邮件。"},
            {"type": "选择", "q": "用户在浏览器输入www.baidu.com后，首先发生什么？", "options": ["A. 直接连接服务器", "B. DNS解析域名获取IP", "C. 发送HTTP请求", "D. 建立TCP连接"], "answer": "B", "explain": "浏览器先通过DNS将域名解析为IP地址，然后才能建立TCP连接。"},
            {"type": "判断", "q": "DNS解析过程就是将域名转换为IP地址。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "DNS的核心功能是域名到IP地址的映射。"},
            {"type": "判断", "q": "HTTP是无状态协议，不保存客户端状态。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "HTTP本身是无状态的，用Cookie/Session来维护状态。"},
            {"type": "判断", "q": "SSH默认使用端口22。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "SSH(Secure Shell)默认端口22，用于远程管理服务器。"},
            {"type": "判断", "q": "DHCP可以自动为电脑分配IP地址，无需手动配置。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "DHCP自动分配IP配置，是局域网中最常用的协议之一。"},
            {"type": "填空", "q": "DNS将____解析为IP地址。", "answer": "域名", "explain": "DNS = Domain Name System，域名系统。"},
            {"type": "填空", "q": "HTTP协议默认端口是____。", "answer": "80", "explain": "HTTP默认80端口，HTTPS默认443端口。"},
        ]
    },
    {
        "num": 30, "chapter": "第7章 计算机网络", "title": "网络设备与第七章复习",
        "questions": [
            {"type": "选择", "q": "交换机工作在OSI模型的哪一层？", "options": ["A. 物理层", "B. 数据链路层", "C. 网络层", "D. 应用层"], "answer": "B", "explain": "交换机是数据链路层设备，基于MAC地址转发数据帧。"},
            {"type": "选择", "q": "路由器工作在OSI模型的哪一层？", "options": ["A. 物理层", "B. 数据链路层", "C. 网络层", "D. 应用层"], "answer": "C", "explain": "路由器是网络层设备，基于IP地址转发数据包，连接不同网络。"},
            {"type": "选择", "q": "交换机和路由器的主要区别是？", "options": ["A. 交换机用IP地址，路由器用MAC地址", "B. 交换机用MAC地址在局域网内转发，路由器用IP地址连接不同网络", "C. 没有区别", "D. 路由器更快"], "answer": "B", "explain": "交换机基于MAC在局域网内转发帧，路由器基于IP在不同网络间路由包。"},
            {"type": "选择", "q": "复习：OSI模型共有几层？", "options": ["A. 4层", "B. 5层", "C. 7层", "D. 9层"], "answer": "C", "explain": "OSI七层模型。"},
            {"type": "选择", "q": "复习：TCP三次握手的目的是？", "options": ["A. 加密数据", "B. 建立可靠连接", "C. 传输文件", "D. 关闭连接"], "answer": "B", "explain": "三次握手建立TCP连接。"},
            {"type": "选择", "q": "复习：以下哪个是私有IP地址？", "options": ["A. 8.8.8.8", "B. 10.0.0.1", "C. 202.96.128.86", "D. 114.114.114.114"], "answer": "B", "explain": "10.x.x.x是私有地址段。"},
            {"type": "判断", "q": "复习：TCP是面向连接的可靠传输协议。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "TCP面向连接、可靠传输。"},
            {"type": "判断", "q": "复习：DNS的作用是将域名解析为IP地址。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "DNS做域名到IP的解析。"},
            {"type": "判断", "q": "复习：UDP比TCP速度快但不可靠。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "UDP无连接、不保证可靠但速度快。"},
            {"type": "判断", "q": "交换机可以连接不同的网络（如局域网和互联网）。", "options": ["A. 正确", "B. 错误"], "answer": "B", "explain": "连接不同网络是路由器的功能，交换机用于局域网内部连接。"},
            {"type": "填空", "q": "复习：____机基于MAC地址转发数据帧（填两个字）。", "answer": "交换", "explain": "交换机是数据链路层设备。"},
            {"type": "填空", "q": "复习：____器基于IP地址在不同网络间转发数据包。", "answer": "路由", "explain": "路由器是网络层设备。"},
        ]
    },
    {
        "num": 31, "chapter": "第8章 数据库基础", "title": "数据库基础概念",
        "questions": [
            {"type": "选择", "q": "数据库(DB)的本质是？", "options": ["A. 一个软件", "B. 有组织地存储的数据集合", "C. 一台服务器", "D. 一个文件夹"], "answer": "B", "explain": "数据库是长期存储在计算机中、有组织、可共享的数据集合。"},
            {"type": "选择", "q": "DBMS的中文含义是？", "options": ["A. 数据库管理系统", "B. 数据备份系统", "C. 数据传输服务", "D. 数据安全系统"], "answer": "A", "explain": "DBMS = Database Management System，数据库管理系统。"},
            {"type": "选择", "q": "以下哪个是关系型数据库？", "options": ["A. MongoDB", "B. Redis", "C. MySQL", "D. Cassandra"], "answer": "C", "explain": "MySQL是关系型数据库，MongoDB/Redis/Cassandra是非关系型(NoSQL)数据库。"},
            {"type": "选择", "q": "以下哪个是非关系型数据库(NoSQL)？", "options": ["A. MySQL", "B. Oracle", "C. Redis", "D. SQL Server"], "answer": "C", "explain": "Redis是键值型NoSQL数据库，MySQL/Oracle/SQL Server都是关系型。"},
            {"type": "选择", "q": "关系型数据库使用什么来组织数据？", "options": ["A. 文件夹", "B. 二维表（行和列）", "C. 树形结构", "D. 图结构"], "answer": "B", "explain": "关系型数据库用二维表组织数据，表由行(记录)和列(字段)组成。"},
            {"type": "选择", "q": "以下哪个不是数据库的优点？", "options": ["A. 数据共享", "B. 减少冗余", "C. 数据一致性", "D. 增加数据冗余"], "answer": "D", "explain": "数据库设计目标是减少冗余而非增加。"},
            {"type": "判断", "q": "Excel表格可以算作一种简单的数据库。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "Excel有行列结构，类似二维表，可以理解为简易数据库，但功能远不如专业DBMS。"},
            {"type": "判断", "q": "MySQL是关系型数据库管理系统。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "MySQL是最流行的开源关系型数据库管理系统。"},
            {"type": "判断", "q": "NoSQL数据库完全取代了关系型数据库。", "options": ["A. 正确", "B. 错误"], "answer": "B", "explain": "NoSQL和关系型数据库各有优势，通常配合使用。"},
            {"type": "判断", "q": "数据库管理系统(DBMS)是管理和操作数据库的软件。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "DBMS如MySQL/Oracle是管理数据库的系统软件。"},
            {"type": "填空", "q": "DBMS的中文全称是数据库____系统。", "answer": "管理", "explain": "DBMS = Database Management System。"},
            {"type": "填空", "q": "关系型数据库用____表来组织数据（填两个字）。", "answer": "二维", "explain": "关系型数据库的核心是二维表结构。"},
        ]
    },
    {
        "num": 32, "chapter": "第8章 数据库基础", "title": "关系型数据库与表结构",
        "questions": [
            {"type": "选择", "q": "数据库表中的\"行\"代表什么？", "options": ["A. 字段", "B. 一条记录", "C. 数据类型", "D. 表名"], "answer": "B", "explain": "表中的一行(Row)代表一条记录(Record)，如一个学生的信息。"},
            {"type": "选择", "q": "数据库表中的\"列\"代表什么？", "options": ["A. 一条记录", "B. 一个字段(属性)", "C. 主键", "D. 外键"], "answer": "B", "explain": "表中的一列(Column)代表一个字段(Field)，如姓名、年龄等属性。"},
            {"type": "选择", "q": "主键的作用是？", "options": ["A. 加密数据", "B. 唯一标识表中的每一条记录", "C. 排序数据", "D. 压缩数据"], "answer": "B", "explain": "主键(Primary Key)是表中唯一标识每条记录的字段，不能重复、不能为空。"},
            {"type": "选择", "q": "外键的作用是？", "options": ["A. 加速查询", "B. 建立表与表之间的关联关系", "C. 加密数据", "D. 压缩数据"], "answer": "B", "explain": "外键(Foreign Key)引用另一张表的主键，建立表间关系。"},
            {"type": "选择", "q": "学生表中学号适合作为什么？", "options": ["A. 外键", "B. 主键", "C. 普通字段", "D. 索引"], "answer": "B", "explain": "学号唯一标识每个学生，不重复且不为空，适合做主键。"},
            {"type": "选择", "q": "成绩表中的学号字段（引用学生表）属于什么？", "options": ["A. 主键", "B. 外键", "C. 唯一键", "D. 普通字段"], "answer": "B", "explain": "成绩表中的学号引用学生表的主键，是外键，建立两表关联。"},
            {"type": "判断", "q": "主键值可以重复。", "options": ["A. 正确", "B. 错误"], "answer": "B", "explain": "主键必须唯一，不能有重复值。"},
            {"type": "判断", "q": "一个表只能有一个主键。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "一个表只能定义一个主键（可以是单个字段或多个字段的组合）。"},
            {"type": "判断", "q": "外键用于建立表与表之间的关系。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "外键引用其他表的主键，实现表间关联。"},
            {"type": "判断", "q": "表中的每条记录代表一个实体。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "如学生表中每行记录代表一个学生实体。"},
            {"type": "填空", "q": "表中的____代表一条记录（填一个字）。", "answer": "行", "explain": "行(Row) = 记录(Record)。"},
            {"type": "填空", "q": "____键唯一标识表中的每条记录（填一个字）。", "answer": "主", "explain": "主键(Primary Key)唯一且不为空。"},
        ]
    },
    {
        "num": 33, "chapter": "第8章 数据库基础", "title": "SQL语言入门",
        "questions": [
            {"type": "选择", "q": "SQL的全称是？", "options": ["A. 结构化查询语言", "B. 简单查询语言", "C. 标准查询逻辑", "D. 系统查询库"], "answer": "A", "explain": "SQL = Structured Query Language，结构化查询语言。"},
            {"type": "选择", "q": "以下哪个SQL语句用于查询数据？", "options": ["A. INSERT", "B. SELECT", "C. UPDATE", "D. DELETE"], "answer": "B", "explain": "SELECT用于查询数据，INSERT插入，UPDATE修改，DELETE删除。"},
            {"type": "选择", "q": "以下哪个SQL语句用于插入数据？", "options": ["A. SELECT", "B. INSERT", "C. CREATE", "D. DROP"], "answer": "B", "explain": "INSERT INTO语句用于向表中插入新记录。"},
            {"type": "选择", "q": "CRUD操作中R代表什么？", "options": ["A. Read(查询)", "B. Replace(替换)", "C. Run(运行)", "D. Remove(删除)"], "answer": "A", "explain": "CRUD = Create(增) + Read(查) + Update(改) + Delete(删)。"},
            {"type": "选择", "q": "以下SQL语句：SELECT * FROM students; 的含义是？", "options": ["A. 删除students表", "B. 查询students表的所有记录", "C. 创建students表", "D. 修改students表"], "answer": "B", "explain": "SELECT * 查询所有字段，FROM students指定表名。"},
            {"type": "选择", "q": "DELETE FROM students WHERE id=5; 的含义是？", "options": ["A. 删除整个表", "B. 删除id=5的记录", "C. 查询id=5的记录", "D. 修改id=5的记录"], "answer": "B", "explain": "DELETE删除记录，WHERE id=5限定只删除id为5的那条记录。"},
            {"type": "判断", "q": "SQL语句不区分大小写。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "SQL关键字(SELECT/FROM等)不区分大小写，但习惯上大写。"},
            {"type": "判断", "q": "SELECT语句可以配合WHERE条件筛选数据。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "WHERE子句用于指定查询条件，如WHERE age>18。"},
            {"type": "判断", "q": "UPDATE语句可以修改表中已有的记录。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "UPDATE SET修改已存在记录的字段值。"},
            {"type": "判断", "q": "不带WHERE的DELETE会删除表中所有记录。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "DELETE FROM 表名; 不加WHERE会清空整张表（表结构保留）。"},
            {"type": "填空", "q": "SQL的中文全称是____查询语言。", "answer": "结构化", "explain": "SQL = Structured Query Language。"},
            {"type": "填空", "q": "CRUD中的C代表____（填英文）。", "answer": "Create", "explain": "C=Create(增), R=Read(查), U=Update(改), D=Delete(删)。"},
        ]
    },
    {
        "num": 34, "chapter": "第8章 数据库基础", "title": "SQL查询进阶",
        "questions": [
            {"type": "选择", "q": "WHERE子句的作用是？", "options": ["A. 排序结果", "B. 筛选满足条件的记录", "C. 分组统计", "D. 连接表"], "answer": "B", "explain": "WHERE用于指定查询条件，只返回满足条件的记录。"},
            {"type": "选择", "q": "ORDER BY子句的作用是？", "options": ["A. 筛选记录", "B. 对结果排序", "C. 分组统计", "D. 限制行数"], "answer": "B", "explain": "ORDER BY按指定字段排序，ASC升序(默认)，DESC降序。"},
            {"type": "选择", "q": "GROUP BY子句的作用是？", "options": ["A. 排序", "B. 筛选", "C. 按字段分组统计", "D. 连接表"], "answer": "C", "explain": "GROUP BY按指定字段分组，配合COUNT/SUM/AVG等聚合函数使用。"},
            {"type": "选择", "q": "以下哪个是聚合函数？", "options": ["A. LEFT", "B. COUNT", "C. JOIN", "D. WHERE"], "answer": "B", "explain": "COUNT(计数)、SUM(求和)、AVG(平均)、MAX(最大)、MIN(最小)是聚合函数。"},
            {"type": "选择", "q": "JOIN的作用是？", "options": ["A. 排序数据", "B. 将多个表的数据连接查询", "C. 分组统计", "D. 删除数据"], "answer": "B", "explain": "JOIN通过关联条件将多张表的数据连接在一起查询。"},
            {"type": "选择", "q": "SELECT name, COUNT(*) FROM students GROUP BY class; 的含义是？", "options": ["A. 查询所有学生", "B. 按班级分组统计每班学生数", "C. 排序学生", "D. 删除学生"], "answer": "B", "explain": "GROUP BY class按班级分组，COUNT(*)统计每组人数。"},
            {"type": "判断", "q": "ORDER BY默认是升序(ASC)排列。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "ORDER BY不指定方向时默认升序(从小到大)。"},
            {"type": "判断", "q": "COUNT(*)用于统计记录数量。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "COUNT(*)返回查询结果中的行数。"},
            {"type": "判断", "q": "WHERE子句可以在聚合函数后使用（如WHERE COUNT(*)>5）。", "options": ["A. 正确", "B. 错误"], "answer": "B", "explain": "聚合条件需用HAVING子句，WHERE用于分组前的行级筛选。"},
            {"type": "判断", "q": "JOIN可以将学生表和成绩表通过学号关联查询。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "通过学号(外键)JOIN两表，可以查出学生姓名和对应成绩。"},
            {"type": "填空", "q": "SQL中用____ BY子句对查询结果排序。", "answer": "ORDER", "explain": "ORDER BY用于排序。"},
            {"type": "填空", "q": "统计记录数量用____(*)函数。", "answer": "COUNT", "explain": "COUNT(*)返回行数。"},
        ]
    },
    {
        "num": 35, "chapter": "第8章 数据库基础", "title": "数据库设计与云数据库",
        "questions": [
            {"type": "选择", "q": "E-R图中的E代表什么？", "options": ["A. 扩展", "B. 实体", "C. 关系", "D. 加密"], "answer": "B", "explain": "E-R图 = Entity-Relationship图，E=实体，R=关系。"},
            {"type": "选择", "q": "数据库设计的一般步骤是？", "options": ["A. 需求分析->概念设计->逻辑设计->物理设计", "B. 直接建表", "C. 先写SQL再设计", "D. 先开发再设计"], "answer": "A", "explain": "标准流程：需求分析->概念设计(E-R图)->逻辑设计(表结构)->物理设计。"},
            {"type": "选择", "q": "云数据库的优势是？", "options": ["A. 不需要网络", "B. 自动备份、弹性扩展、免运维", "C. 更不安全", "D. 只能存小数据"], "answer": "B", "explain": "云数据库提供自动备份、弹性扩容、高可用、免运维等优势。"},
            {"type": "选择", "q": "以下哪个是云数据库产品？", "options": ["A. 本地MySQL", "B. 阿里云RDS", "C. Excel文件", "D. txt文件"], "answer": "B", "explain": "阿里云RDS(Relational Database Service)是云关系型数据库服务。"},
            {"type": "选择", "q": "数据库设计中\"规范化\"的主要目的是？", "options": ["A. 增加数据冗余", "B. 减少数据冗余和依赖", "C. 加速查询", "D. 加密数据"], "answer": "B", "explain": "规范化通过分解表来消除冗余和异常，提高数据一致性。"},
            {"type": "选择", "q": "云数据库与传统自建数据库的主要区别是？", "options": ["A. 云数据库不需要SQL", "B. 云数据库由云服务商托管运维", "C. 云数据库更慢", "D. 没有区别"], "answer": "B", "explain": "云数据库由云服务商负责部署、备份、扩容、高可用等运维工作。"},
            {"type": "判断", "q": "E-R图是数据库概念设计的工具。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "E-R图用于概念设计阶段，描述实体、属性和实体间关系。"},
            {"type": "判断", "q": "云数据库可以自动备份和恢复数据。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "自动备份是云数据库的基本功能之一。"},
            {"type": "判断", "q": "数据库设计不需要考虑未来扩展需求。", "options": ["A. 正确", "B. 错误"], "answer": "B", "explain": "好的数据库设计需要考虑未来数据增长和功能扩展。"},
            {"type": "判断", "q": "云数据库支持弹性扩展存储容量和计算资源。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "云数据库可以在线升级CPU/内存/存储，不影响业务。"},
            {"type": "填空", "q": "E-R图中E代表____（填两个字）。", "answer": "实体", "explain": "Entity = 实体，如学生、课程等。"},
            {"type": "填空", "q": "阿里云的云数据库服务叫____（填三个字母）。", "answer": "RDS", "explain": "RDS = Relational Database Service。"},
        ]
    },
    {
        "num": 36, "chapter": "第8章 数据库基础", "title": "数据库安全与第三单元复习",
        "questions": [
            {"type": "选择", "q": "SQL注入攻击是指？", "options": ["A. 物理破坏服务器", "B. 通过构造恶意SQL语句欺骗服务器执行", "C. 窃取数据库密码", "D. 删除数据库文件"], "answer": "B", "explain": "SQL注入是通过在输入中嵌入恶意SQL片段，使服务器执行非预期操作。"},
            {"type": "选择", "q": "防止SQL注入的有效方法是？", "options": ["A. 使用更复杂的密码", "B. 使用参数化查询(预编译语句)", "C. 关闭数据库", "D. 不使用SQL"], "answer": "B", "explain": "参数化查询将SQL语句和数据分离，有效防止注入。"},
            {"type": "选择", "q": "复习：TCP/IP模型有几层？", "options": ["A. 4层", "B. 5层", "C. 7层", "D. 3层"], "answer": "A", "explain": "TCP/IP四层模型。"},
            {"type": "选择", "q": "复习：以下哪个是关系型数据库？", "options": ["A. MongoDB", "B. Redis", "C. MySQL", "D. Cassandra"], "answer": "C", "explain": "MySQL是关系型数据库。"},
            {"type": "选择", "q": "复习：SELECT语句中用于条件筛选的是？", "options": ["A. ORDER BY", "B. GROUP BY", "C. WHERE", "D. JOIN"], "answer": "C", "explain": "WHERE用于筛选满足条件的记录。"},
            {"type": "选择", "q": "复习：HTTP默认端口是？", "options": ["A. 21", "B. 22", "C. 80", "D. 443"], "answer": "C", "explain": "HTTP默认端口80。"},
            {"type": "判断", "q": "复习：交换机工作在网络层。", "options": ["A. 正确", "B. 错误"], "answer": "B", "explain": "交换机工作在数据链路层，路由器工作在网络层。"},
            {"type": "判断", "q": "复习：IPv4地址是32位二进制数。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "IPv4是32位。"},
            {"type": "判断", "q": "SQL注入是常见的数据库安全威胁。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "SQL注入是OWASP Top 10安全威胁之一。"},
            {"type": "判断", "q": "复习：主键可以重复。", "options": ["A. 正确", "B. 错误"], "answer": "B", "explain": "主键必须唯一，不能重复。"},
            {"type": "填空", "q": "复习：SQL注入可以通过____查询来防范（填三个字）。", "answer": "参数化", "explain": "参数化查询(预编译语句)将代码和数据分离。"},
            {"type": "填空", "q": "复习：DNS将域名解析为____地址。", "answer": "IP", "explain": "DNS做域名到IP地址的解析。"},
        ]
    },
    # ===== 第四单元 课时37-48 =====
    {
        "num": 37, "chapter": "第9章 虚拟化基础", "title": "虚拟化技术概述",
        "questions": [
            {"type": "选择", "q": "虚拟化技术的核心思想是？", "options": ["A. 增加硬件数量", "B. 将物理资源抽象为可分配的虚拟资源", "C. 加速网络", "D. 加密数据"], "answer": "B", "explain": "虚拟化通过软件将物理资源(CPU/内存/存储/网络)抽象为可灵活分配的虚拟资源。"},
            {"type": "选择", "q": "虚拟化与云计算的关系是？", "options": ["A. 完全无关", "B. 虚拟化是云计算的核心基础技术", "C. 云计算是虚拟化的基础", "D. 两者完全相同"], "answer": "B", "explain": "虚拟化实现了资源池化，是云计算IaaS层的核心技术基础。"},
            {"type": "选择", "q": "一台物理机虚拟出多台虚拟机后，每台虚拟机可以？", "options": ["A. 共享同一IP", "B. 独立运行不同的操作系统", "C. 只能运行Linux", "D. 不能联网"], "answer": "B", "explain": "每台虚拟机都是独立的计算机系统，可以安装不同的操作系统。"},
            {"type": "选择", "q": "虚拟化的主要优势不包括？", "options": ["A. 提高资源利用率", "B. 降低硬件成本", "C. 提高散热效果", "D. 方便备份迁移"], "answer": "C", "explain": "提高散热效果不是虚拟化的优势。"},
            {"type": "选择", "q": "服务器虚拟化后，一台物理机上运行多个虚拟机，这体现了？", "options": ["A. 资源池化", "B. 数据加密", "C. 网络加速", "D. 存储压缩"], "answer": "A", "explain": "物理资源被池化为虚拟资源池，按需分配给各虚拟机。"},
            {"type": "选择", "q": "以下哪个不是虚拟化的类型？", "options": ["A. 服务器虚拟化", "B. 桌面虚拟化", "C. 网络虚拟化", "D. 空调虚拟化"], "answer": "D", "explain": "空调虚拟化不存在，虚拟化类型包括服务器/桌面/网络/存储虚拟化。"},
            {"type": "判断", "q": "虚拟化技术可以在一台物理机上运行多个独立的虚拟机。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "这是虚拟化最基本的功能。"},
            {"type": "判断", "q": "虚拟机和物理机一样可以安装操作系统和应用软件。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "虚拟机是完整的计算机系统模拟。"},
            {"type": "判断", "q": "虚拟化技术可以减少物理服务器的数量，降低能耗。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "多台虚拟机整合到一台物理机上，减少硬件和能耗。"},
            {"type": "判断", "q": "虚拟化就是云计算，两者没有区别。", "options": ["A. 正确", "B. 错误"], "answer": "B", "explain": "虚拟化是技术手段，云计算是服务模式，虚拟化是云计算的基础之一。"},
            {"type": "填空", "q": "虚拟化技术的核心思想是将物理资源____为虚拟资源（填两个字）。", "answer": "抽象", "explain": "抽象(abstraction)是虚拟化的核心。"},
            {"type": "填空", "q": "虚拟化是____计算的核心基础技术（填两个字）。", "answer": "云", "explain": "没有虚拟化就没有现代云计算。"},
        ]
    },
    {
        "num": 38, "chapter": "第9章 虚拟化基础", "title": "虚拟化原理与Hypervisor",
        "questions": [
            {"type": "选择", "q": "Hypervisor也称为？", "options": ["A. 虚拟机监视器(VMM)", "B. 操作系统", "C. 数据库", "D. 网络协议"], "answer": "A", "explain": "Hypervisor = Virtual Machine Monitor(VMM)，虚拟机监视器。"},
            {"type": "选择", "q": "Type 1 Hypervisor（裸金属型）直接运行在什么上？", "options": ["A. 操作系统上", "B. 物理硬件上", "C. 虚拟机内", "D. 浏览器中"], "answer": "B", "explain": "Type 1直接运行在物理硬件上，性能最好，如VMware ESXi。"},
            {"type": "选择", "q": "Type 2 Hypervisor（托管型）运行在什么上？", "options": ["A. 物理硬件上", "B. 操作系统上", "C. 虚拟机内", "D. 路由器上"], "answer": "B", "explain": "Type 2运行在宿主操作系统上，如VMware Workstation。"},
            {"type": "选择", "q": "以下哪个是Type 1 Hypervisor？", "options": ["A. VMware Workstation", "B. VirtualBox", "C. VMware ESXi", "D. QEMU(非KVM)"], "answer": "C", "explain": "ESXi是裸金属型(Type 1)，直接运行在硬件上。Workstation/VirtualBox是Type 2。"},
            {"type": "选择", "q": "虚拟机的组成部分包括？", "options": ["A. 只有CPU", "B. 虚拟CPU、虚拟内存、虚拟磁盘、虚拟网卡", "C. 只有硬盘", "D. 只有网络"], "answer": "B", "explain": "虚拟机包含完整的虚拟硬件：vCPU、vMem、vDisk、vNIC等。"},
            {"type": "选择", "q": "以下哪个是Type 2 Hypervisor？", "options": ["A. VMware ESXi", "B. Hyper-V(裸金属模式)", "C. VMware Workstation", "D. Xen"], "answer": "C", "explain": "VMware Workstation需要先装操作系统再安装，是Type 2。"},
            {"type": "判断", "q": "Hypervisor是创建和运行虚拟机的软件层。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "Hypervisor管理和调度虚拟机对物理资源的访问。"},
            {"type": "判断", "q": "Type 1 Hypervisor性能优于Type 2。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "Type 1直接运行在硬件上，没有宿主OS的开销，性能更好。"},
            {"type": "判断", "q": "VMware Workstation属于裸金属型虚拟化。", "options": ["A. 正确", "B. 错误"], "answer": "B", "explain": "VMware Workstation是托管型(Type 2)，需要宿主操作系统。"},
            {"type": "判断", "q": "一台物理机上的多个虚拟机共享物理CPU和内存。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "Hypervisor负责将物理资源分配给各虚拟机共享使用。"},
            {"type": "填空", "q": "Hypervisor也叫虚拟机____器（填两个字）。", "answer": "监视", "explain": "VMM = Virtual Machine Monitor。"},
            {"type": "填空", "q": "Type ____型Hypervisor直接运行在物理硬件上（填数字）。", "answer": "1", "explain": "Type 1(裸金属型)直接运行在硬件上。"},
        ]
    },
    {
        "num": 39, "chapter": "第9章 虚拟化基础", "title": "VMware虚拟化",
        "questions": [
            {"type": "选择", "q": "VMware Workstation属于哪种类型的虚拟化软件？", "options": ["A. Type 1(裸金属)", "B. Type 2(托管型)", "C. 容器", "D. 云平台"], "answer": "B", "explain": "VMware Workstation运行在操作系统上，是Type 2。"},
            {"type": "选择", "q": "VMware ESXi主要用于？", "options": ["A. 个人电脑", "B. 企业级服务器虚拟化", "C. 手机", "D. 路由器"], "answer": "B", "explain": "ESXi是企业级Type 1 Hypervisor，用于数据中心服务器虚拟化。"},
            {"type": "选择", "q": "创建虚拟机时需要指定什么？", "options": ["A. CPU核数、内存大小、硬盘容量", "B. 显示器品牌", "C. 键盘颜色", "D. 鼠标型号"], "answer": "A", "explain": "创建虚拟机需配置CPU、内存、硬盘等虚拟硬件参数。"},
            {"type": "选择", "q": "VMware虚拟机的快照功能的作用是？", "options": ["A. 加速网络", "B. 保存虚拟机某一时刻的状态，便于回退", "C. 加密数据", "D. 压缩文件"], "answer": "B", "explain": "快照保存虚拟机当前状态，出问题时可快速恢复到快照点。"},
            {"type": "选择", "q": "VMware Workstation和ESXi的主要区别是？", "options": ["A. 功能相同", "B. Workstation是桌面级(Type 2)，ESXi是企业级(Type 1)", "C. ESXi更便宜", "D. Workstation更快"], "answer": "B", "explain": "Workstation是桌面虚拟化软件，ESXi是服务器级裸金属虚拟化平台。"},
            {"type": "选择", "q": "虚拟机的克隆功能可以？", "options": ["A. 删除虚拟机", "B. 创建一个与原虚拟机相同的副本", "C. 加速虚拟机", "D. 压缩虚拟机"], "answer": "B", "explain": "克隆创建虚拟机的完整副本，分为链接克隆和完整克隆。"},
            {"type": "判断", "q": "VMware Workstation可以在Windows上运行Linux虚拟机。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "虚拟机可以安装不同的操作系统。"},
            {"type": "判断", "q": "VMware ESXi需要先安装Windows才能使用。", "options": ["A. 正确", "B. 错误"], "answer": "B", "explain": "ESXi是Type 1 Hypervisor，直接安装在物理硬件上，不需要宿主OS。"},
            {"type": "判断", "q": "虚拟机快照可以在系统崩溃后恢复到之前的状态。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "快照就是保存状态以便恢复。"},
            {"type": "判断", "q": "虚拟机的配置创建后就不能修改了。", "options": ["A. 正确", "B. 错误"], "answer": "B", "explain": "虚拟机关机后可以修改CPU、内存等配置。"},
            {"type": "填空", "q": "VMware ____是桌面级虚拟化软件（填英文）。", "answer": "Workstation", "explain": "VMware Workstation是最流行的桌面虚拟化软件。"},
            {"type": "填空", "q": "虚拟机____功能可以保存当前状态以便恢复（填两个字）。", "answer": "快照", "explain": "Snapshot快照是虚拟化的重要功能。"},
        ]
    },
    {
        "num": 40, "chapter": "第9章 虚拟化基础", "title": "KVM虚拟化",
        "questions": [
            {"type": "选择", "q": "KVM是哪个操作系统内核中的虚拟化模块？", "options": ["A. Windows", "B. Linux", "C. macOS", "D. Android"], "answer": "B", "explain": "KVM(Kernel-based Virtual Machine)是Linux内核的虚拟化模块。"},
            {"type": "选择", "q": "KVM属于哪种类型的虚拟化？", "options": ["A. Type 2", "B. Type 1(内核级)", "C. 容器", "D. 模拟器"], "answer": "B", "explain": "KVM将Linux内核变为Hypervisor，属于Type 1（内核级）虚拟化。"},
            {"type": "选择", "q": "KVM通常与什么配合使用来管理虚拟机？", "options": ["A. VMware", "B. QEMU", "C. Docker", "D. Hyper-V"], "answer": "B", "explain": "KVM负责CPU/内存虚拟化，QEMU负责设备模拟，两者配合使用。"},
            {"type": "选择", "q": "KVM在云计算中的典型应用是？", "options": ["A. 个人桌面", "B. OpenStack等云平台的底层虚拟化", "C. 手机应用", "D. 路由器固件"], "answer": "B", "explain": "KVM是OpenStack等开源云平台最常用的虚拟化后端。"},
            {"type": "选择", "q": "以下哪个不是KVM的特点？", "options": ["A. 开源免费", "B. 性能好(内核级)", "C. 需要Windows系统", "D. 与Linux内核集成"], "answer": "C", "explain": "KVM是Linux内核模块，不需要Windows。"},
            {"type": "选择", "q": "KVM虚拟化需要CPU支持什么技术？", "options": ["A. 超线程", "B. 硬件虚拟化(Intel VT-x / AMD-V)", "C. 多核", "D. 超频"], "answer": "B", "explain": "KVM需要CPU支持硬件虚拟化扩展(Intel VT-x或AMD-V)。"},
            {"type": "判断", "q": "KVM是开源的虚拟化技术。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "KVM是Linux内核的一部分，完全开源。"},
            {"type": "判断", "q": "KVM和QEMU是同一个软件。", "options": ["A. 正确", "B. 错误"], "answer": "B", "explain": "KVM是内核模块(CPU/内存虚拟化)，QEMU是用户态工具(设备模拟/管理)。"},
            {"type": "判断", "q": "KVM需要硬件虚拟化支持才能运行。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "KVM依赖CPU的硬件虚拟化扩展。"},
            {"type": "判断", "q": "KVM常用于开源云平台如OpenStack。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "KVM是OpenStack默认的虚拟化驱动。"},
            {"type": "填空", "q": "KVM是____内核中的虚拟化模块（填英文操作系统名）。", "answer": "Linux", "explain": "KVM = Kernel-based Virtual Machine，Linux内核虚拟化。"},
            {"type": "填空", "q": "KVM通常与____配合管理虚拟机（填四个字母）。", "answer": "QEMU", "explain": "QEMU提供设备模拟和管理接口。"},
        ]
    },
    {
        "num": 41, "chapter": "第9章 虚拟化基础", "title": "容器技术Docker",
        "questions": [
            {"type": "选择", "q": "Docker容器与虚拟机的主要区别是？", "options": ["A. 容器更重", "B. 容器共享宿主OS内核，无需独立操作系统", "C. 容器不能运行应用", "D. 容器需要Hypervisor"], "answer": "B", "explain": "容器共享宿主OS内核，不需要每个容器都装操作系统，更轻量。"},
            {"type": "选择", "q": "Docker的主要优势是？", "options": ["A. 启动慢", "B. 轻量、快速、一致性好", "C. 资源占用大", "D. 只能在Linux运行"], "answer": "B", "explain": "容器秒级启动、资源占用少、开发测试生产环境一致。"},
            {"type": "选择", "q": "Docker镜像(Image)是什么？", "options": ["A. 运行中的程序", "B. 创建容器的只读模板", "C. 虚拟机", "D. 操作系统"], "answer": "B", "explain": "镜像是静态的只读模板，包含应用和依赖，用于创建容器实例。"},
            {"type": "选择", "q": "Docker容器(Container)是什么？", "options": ["A. 静态模板", "B. 镜像的运行实例", "C. 操作系统", "D. 物理机"], "answer": "B", "explain": "容器是镜像启动后的运行实例，类似面向对象中对象与类的关系。"},
            {"type": "选择", "q": "以下哪个不是容器的优势？", "options": ["A. 启动速度快", "B. 资源占用少", "C. 完全隔离(独立内核)", "D. 环境一致性"], "answer": "C", "explain": "容器共享内核，隔离性不如虚拟机(每个VM有独立内核)。"},
            {"type": "选择", "q": "Docker在云计算中的典型应用场景是？", "options": ["A. 替代操作系统", "B. 微服务部署和CI/CD", "C. 硬件维修", "D. 网络布线"], "answer": "B", "explain": "Docker广泛用于微服务架构和持续集成/持续部署(CI/CD)。"},
            {"type": "判断", "q": "Docker容器比虚拟机启动更快。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "容器不需要启动操作系统，秒级启动；虚拟机需几分钟。"},
            {"type": "判断", "q": "Docker容器每个都有独立的操作系统内核。", "options": ["A. 正确", "B. 错误"], "answer": "B", "explain": "容器共享宿主机内核，这是它比虚拟机轻量的原因。"},
            {"type": "判断", "q": "Docker镜像可以看作容器的模板。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "镜像是静态模板，容器是运行实例。"},
            {"type": "判断", "q": "容器的隔离性比虚拟机强。", "options": ["A. 正确", "B. 错误"], "answer": "B", "explain": "虚拟机有独立内核，隔离更强；容器共享内核，隔离性较弱。"},
            {"type": "填空", "q": "Docker____是创建容器的只读模板（填两个字）。", "answer": "镜像", "explain": "Image = 镜像。"},
            {"type": "填空", "q": "Docker____是镜像的运行实例（填两个字）。", "answer": "容器", "explain": "Container = 容器。"},
        ]
    },
    {
        "num": 42, "chapter": "第9章 虚拟化基础", "title": "虚拟化网络与存储",
        "questions": [
            {"type": "选择", "q": "虚拟交换机的作用是？", "options": ["A. 加速物理网络", "B. 连接虚拟机之间的网络通信", "C. 加密数据", "D. 分配IP地址"], "answer": "B", "explain": "虚拟交换机(vSwitch)是软件实现的交换机，连接同一物理机上虚拟机的网络。"},
            {"type": "选择", "q": "虚拟网卡的类型不包括？", "options": ["A. 桥接模式(Bridged)", "B. NAT模式", "C. 仅主机模式(Host-only)", "D. 蓝牙模式"], "answer": "D", "explain": "虚拟网卡常见模式：桥接/NAT/仅主机，没有蓝牙模式。"},
            {"type": "选择", "q": "NAT模式下虚拟机如何访问外网？", "options": ["A. 直接用物理网卡", "B. 通过宿主机的NAT转发访问外网", "C. 不能访问外网", "D. 通过蓝牙"], "answer": "B", "explain": "NAT模式：虚拟机通过宿主机的网络地址转换访问外部网络。"},
            {"type": "选择", "q": "虚拟化存储技术VMDK是什么？", "options": ["A. 网络协议", "B. 虚拟机磁盘文件格式", "C. 数据库", "D. 操作系统"], "answer": "B", "explain": "VMDK是VMware的虚拟磁盘文件格式，存储虚拟机的硬盘数据。"},
            {"type": "选择", "q": "桥接模式下虚拟机的网络特点是？", "options": ["A. 与宿主机共享IP", "B. 虚拟机像物理机一样直接接入物理网络", "C. 不能上网", "D. 只能内网通信"], "answer": "B", "explain": "桥接模式：虚拟机直接接入物理网络，有独立IP，相当于网络中的一台物理机。"},
            {"type": "选择", "q": "存储虚拟化的好处是？", "options": ["A. 减少存储容量", "B. 统一管理多个存储设备，提高利用率", "C. 加速CPU", "D. 降低网络安全"], "answer": "B", "explain": "存储虚拟化将多个物理存储设备抽象为统一的存储池。"},
            {"type": "判断", "q": "虚拟交换机是软件实现的交换机。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "vSwitch是纯软件实现，运行在Hypervisor中。"},
            {"type": "判断", "q": "桥接模式下虚拟机可以获得与物理网络同网段的IP地址。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "桥接模式虚拟机直接在物理网络中，获得同网段IP。"},
            {"type": "判断", "q": "虚拟机的虚拟磁盘文件可以像普通文件一样复制和迁移。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "虚拟磁盘(VMDK等)就是文件，便于备份和迁移。"},
            {"type": "判断", "q": "NAT模式下虚拟机可以直接被外部网络访问。", "options": ["A. 正确", "B. 错误"], "answer": "B", "explain": "NAT模式虚拟机在宿主机后面，默认不能被外部直接访问(需端口转发)。"},
            {"type": "填空", "q": "虚拟____是连接虚拟机网络通信的软件设备（填两个字）。", "answer": "交换", "explain": "vSwitch = 虚拟交换机。"},
            {"type": "填空", "q": "____模式下虚拟机像物理机一样直接接入物理网络。", "answer": "桥接", "explain": "Bridged模式，虚拟机获得物理网络IP。"},
        ]
    },
    {
        "num": 43, "chapter": "第9章 虚拟化基础", "title": "虚拟化与云计算关系 + 第九章复习",
        "questions": [
            {"type": "选择", "q": "虚拟化在云计算中的核心地位体现在？", "options": ["A. 虚拟化就是云计算", "B. 虚拟化实现资源池化，是IaaS的基础", "C. 虚拟化只用于桌面", "D. 云计算不需要虚拟化"], "answer": "B", "explain": "虚拟化将物理资源池化为可按需分配的虚拟资源，是IaaS的核心。"},
            {"type": "选择", "q": "复习：Type 1 Hypervisor直接运行在什么上？", "options": ["A. 操作系统", "B. 物理硬件", "C. 虚拟机", "D. 浏览器"], "answer": "B", "explain": "Type 1(裸金属型)直接运行在硬件上。"},
            {"type": "选择", "q": "复习：Docker容器与虚拟机的主要区别？", "options": ["A. 容器更重", "B. 容器共享宿主OS内核", "C. 容器需要独立OS", "D. 没有区别"], "answer": "B", "explain": "容器共享内核，轻量快速。"},
            {"type": "选择", "q": "复习：VMware ESXi属于什么类型？", "options": ["A. Type 2", "B. Type 1", "C. 容器", "D. 操作系统"], "answer": "B", "explain": "ESXi是Type 1裸金属Hypervisor。"},
            {"type": "选择", "q": "复习：KVM是哪个操作系统的内核模块？", "options": ["A. Windows", "B. macOS", "C. Linux", "D. Android"], "answer": "C", "explain": "KVM是Linux内核虚拟化模块。"},
            {"type": "选择", "q": "云计算IaaS层为用户提供的是？", "options": ["A. 完整的应用", "B. 虚拟化的计算、存储、网络资源", "C. 开发平台", "D. 数据库服务"], "answer": "B", "explain": "IaaS提供虚拟化的基础设施资源(虚拟机/虚拟网络/虚拟存储)。"},
            {"type": "判断", "q": "复习：虚拟化技术可以在一台物理机上运行多个虚拟机。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "这是虚拟化的基本功能。"},
            {"type": "判断", "q": "复习：Docker容器比虚拟机启动更快。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "容器秒级启动，虚拟机需几分钟。"},
            {"type": "判断", "q": "复习：虚拟机快照可以保存状态便于恢复。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "快照是虚拟化的重要功能。"},
            {"type": "判断", "q": "虚拟化技术只用于云计算，没有其他用途。", "options": ["A. 正确", "B. 错误"], "answer": "B", "explain": "虚拟化也用于桌面虚拟化、测试开发、安全沙箱等场景。"},
            {"type": "填空", "q": "复习：虚拟化实现了资源____化，是IaaS的基础。", "answer": "池", "explain": "Resource Pooling 资源池化。"},
            {"type": "填空", "q": "复习：Docker____是容器的运行实例模板（填两个字）。", "answer": "镜像", "explain": "镜像(Image) -> 容器(Container)。"},
        ]
    },
    {
        "num": 44, "chapter": "第10章 Linux基础", "title": "Linux操作系统概述",
        "questions": [
            {"type": "选择", "q": "Linux操作系统的创始人是谁？", "options": ["A. 比尔·盖茨", "B. 林纳斯·托瓦兹(Linus Torvalds)", "C. 史蒂夫·乔布斯", "D. 拉里·埃里森"], "answer": "B", "explain": "Linus Torvalds于1991年创建了Linux内核。"},
            {"type": "选择", "q": "Linux属于什么类型的软件？", "options": ["A. 商业闭源", "B. 开源自由软件", "C. 只能付费使用", "D. 病毒软件"], "answer": "B", "explain": "Linux是开源自由软件，遵循GPL协议。"},
            {"type": "选择", "q": "以下哪个是Linux的发行版？", "options": ["A. Windows 10", "B. macOS", "C. Ubuntu", "D. iOS"], "answer": "C", "explain": "Ubuntu是流行的Linux发行版，Windows/macOS/iOS都不是Linux。"},
            {"type": "选择", "q": "Linux在云计算中的重要地位体现在？", "options": ["A. 云服务器大多运行Linux", "B. Linux是最快的操作系统", "C. Linux不需要硬件", "D. Linux只能用于手机"], "answer": "A", "explain": "绝大多数云服务器(如阿里云ECS)使用Linux操作系统。"},
            {"type": "选择", "q": "以下哪个不是Linux的特点？", "options": ["A. 开源免费", "B. 稳定性好", "C. 安全性高", "D. 必须付费购买"], "answer": "D", "explain": "Linux是开源免费的，不需要付费购买。"},
            {"type": "选择", "q": "Android(安卓)系统的底层基于什么？", "options": ["A. Windows", "B. Linux内核", "C. macOS", "D. iOS"], "answer": "B", "explain": "Android基于Linux内核开发。"},
            {"type": "判断", "q": "Linux操作系统是开源的。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "Linux内核遵循GPL开源协议。"},
            {"type": "判断", "q": "CentOS和Ubuntu都是Linux的发行版。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "CentOS和Ubuntu都是基于Linux内核的发行版。"},
            {"type": "判断", "q": "Linux在服务器领域比Windows Server更普及。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "Linux在服务器/云计算领域占绝对优势。"},
            {"type": "判断", "q": "Linux只能通过图形界面操作。", "options": ["A. 正确", "B. 错误"], "answer": "B", "explain": "Linux支持图形界面和命令行界面，服务器通常用命令行。"},
            {"type": "填空", "q": "Linux的创始人是Linus____（填姓氏）。", "answer": "Torvalds", "explain": "Linus Torvalds(林纳斯·托瓦兹)。"},
            {"type": "填空", "q": "云服务器大多运行____操作系统。", "answer": "Linux", "explain": "Linux是云计算的主流操作系统。"},
        ]
    },
    {
        "num": 45, "chapter": "第10章 Linux基础", "title": "Linux文件系统与目录结构",
        "questions": [
            {"type": "选择", "q": "Linux文件系统的根目录用什么表示？", "options": ["A. C:\\", "B. /", "C. \\", "D. root"], "answer": "B", "explain": "Linux根目录是/，与Windows的C:\\不同。"},
            {"type": "选择", "q": "Linux中用户的主目录(home)是？", "options": ["A. /root", "B. /home/用户名", "C. /usr", "D. /tmp"], "answer": "B", "explain": "普通用户主目录在/home/用户名下，root用户主目录是/root。"},
            {"type": "选择", "q": "Linux中存放系统配置文件的目录是？", "options": ["A. /etc", "B. /home", "C. /tmp", "D. /var"], "answer": "A", "explain": "/etc存放系统配置文件(如网络配置、服务配置等)。"},
            {"type": "选择", "q": "Linux中\"一切皆____\"是其核心设计理念。", "options": ["A. 文件", "B. 进程", "C. 网络", "D. 命令"], "answer": "A", "explain": "Linux的核心理念：一切皆文件(everything is a file)，设备、进程等都以文件形式表示。"},
            {"type": "选择", "q": "/tmp目录的特点是？", "options": ["A. 永久存储", "B. 临时文件目录，可能被自动清理", "C. 系统配置", "D. 用户数据"], "answer": "B", "explain": "/tmp用于存放临时文件，系统重启后可能被清空。"},
            {"type": "选择", "q": "Linux中存放可执行程序的目录是？", "options": ["A. /home", "B. /etc", "C. /usr/bin", "D. /tmp"], "answer": "C", "explain": "/usr/bin存放用户可执行程序，/bin存放系统基本命令。"},
            {"type": "判断", "q": "Linux的根目录是/而不是C:\\。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "Linux没有盘符概念，所有文件从根目录/开始。"},
            {"type": "判断", "q": "Linux中\"一切皆文件\"意味着设备也可以用文件表示。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "如/dev/sda表示硬盘设备文件。"},
            {"type": "判断", "q": "/etc目录存放用户个人文件。", "options": ["A. 正确", "B. 错误"], "answer": "B", "explain": "/etc存放系统配置文件，用户个人文件在/home下。"},
            {"type": "判断", "q": "Linux路径用正斜杠/分隔，Windows用反斜杠\\分隔。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "Linux: /home/user，Windows: C:\\Users\\user。"},
            {"type": "填空", "q": "Linux根目录用____表示（填一个符号）。", "answer": "/", "explain": "根目录是/。"},
            {"type": "填空", "q": "Linux的核心设计理念是\"一切皆____\"。", "answer": "文件", "explain": "Everything is a file。"},
        ]
    },
    {
        "num": 46, "chapter": "第10章 Linux基础", "title": "Linux常用命令（一）",
        "questions": [
            {"type": "选择", "q": "ls命令的作用是？", "options": ["A. 列出目录内容", "B. 创建目录", "C. 删除文件", "D. 复制文件"], "answer": "A", "explain": "ls(list)列出当前目录下的文件和子目录。"},
            {"type": "选择", "q": "cd命令的作用是？", "options": ["A. 复制文件", "B. 切换当前工作目录", "C. 创建文件", "D. 查看文件内容"], "answer": "B", "explain": "cd(change directory)切换工作目录，如cd /home。"},
            {"type": "选择", "q": "pwd命令的作用是？", "options": ["A. 修改密码", "B. 显示当前工作目录的路径", "C. 创建目录", "D. 删除目录"], "answer": "B", "explain": "pwd(print working directory)显示当前所在目录的绝对路径。"},
            {"type": "选择", "q": "mkdir命令的作用是？", "options": ["A. 创建文件", "B. 创建目录", "C. 删除目录", "D. 移动文件"], "answer": "B", "explain": "mkdir(make directory)创建新目录。"},
            {"type": "选择", "q": "rm命令的作用是？", "options": ["A. 重命名文件", "B. 删除文件或目录", "C. 复制文件", "D. 查看文件"], "answer": "B", "explain": "rm(remove)删除文件，rm -r删除目录及其内容。"},
            {"type": "选择", "q": "cp命令的作用是？", "options": ["A. 移动文件", "B. 复制文件或目录", "C. 删除文件", "D. 创建文件"], "answer": "B", "explain": "cp(copy)复制文件，cp -r复制目录。"},
            {"type": "判断", "q": "ls -l可以以长格式显示文件详细信息。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "-l选项显示权限、所有者、大小、修改时间等详细信息。"},
            {"type": "判断", "q": "rm -rf可以强制递归删除目录及其内容。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "-r递归删除，-f强制不提示，非常危险需谨慎使用。"},
            {"type": "判断", "q": "cd ..可以返回上一级目录。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "..表示父目录，cd ..返回上一级。"},
            {"type": "判断", "q": "mv命令既可以移动文件也可以重命名文件。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "mv(move)既可移动文件到其他目录，也可在同一目录下重命名。"},
            {"type": "填空", "q": "____命令列出目录内容（填两个字母）。", "answer": "ls", "explain": "ls = list。"},
            {"type": "填空", "q": "____命令切换当前工作目录（填两个字母）。", "answer": "cd", "explain": "cd = change directory。"},
        ]
    },
    {
        "num": 47, "chapter": "第10章 Linux基础", "title": "Linux常用命令（二）与用户管理",
        "questions": [
            {"type": "选择", "q": "chmod命令的作用是？", "options": ["A. 修改文件所有者", "B. 修改文件权限", "C. 创建用户", "D. 查看文件内容"], "answer": "B", "explain": "chmod(change mode)修改文件/目录的读/写/执行权限。"},
            {"type": "选择", "q": "chown命令的作用是？", "options": ["A. 修改文件权限", "B. 修改文件所有者", "C. 创建目录", "D. 删除用户"], "answer": "B", "explain": "chown(change owner)修改文件/目录的所有者和所属组。"},
            {"type": "选择", "q": "Linux文件权限rwx中r代表什么？", "options": ["A. 运行(run)", "B. 读取(read)", "C. 重命名(rename)", "D. 删除(remove)"], "answer": "B", "explain": "r=read(读)，w=write(写)，x=execute(执行)。"},
            {"type": "选择", "q": "Linux中超级用户(root)的UID是？", "options": ["A. 0", "B. 1", "C. 100", "D. 1000"], "answer": "A", "explain": "root用户的UID是0，拥有最高权限。"},
            {"type": "选择", "q": "useradd命令的作用是？", "options": ["A. 删除用户", "B. 添加新用户", "C. 修改密码", "D. 查看用户"], "answer": "B", "explain": "useradd创建新用户账户。"},
            {"type": "选择", "q": "cat命令的作用是？", "options": ["A. 复制文件", "B. 查看文件内容", "C. 创建目录", "D. 修改权限"], "answer": "B", "explain": "cat(concatenate)显示文件内容，可同时查看多个文件。"},
            {"type": "判断", "q": "Linux中每个文件都有读(r)、写(w)、执行(x)三种权限。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "rwx权限分别针对所有者、所属组、其他用户。"},
            {"type": "判断", "q": "root用户可以执行任何操作，不受权限限制。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "root是超级用户，拥有系统最高权限。"},
            {"type": "判断", "q": "chmod 755 file 将文件权限设为rwxr-xr-x。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "7=rwx(所有者), 5=r-x(组), 5=r-x(其他)。"},
            {"type": "判断", "q": "passwd命令可以修改用户密码。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "passwd 用户名可修改指定用户的密码。"},
            {"type": "填空", "q": "____命令修改文件权限（填五个字母）。", "answer": "chmod", "explain": "chmod = change mode。"},
            {"type": "填空", "q": "Linux超级用户是____（填四个字母）。", "answer": "root", "explain": "root用户UID为0，拥有最高权限。"},
        ]
    },
    {
        "num": 48, "chapter": "第10章 Linux基础", "title": "Linux在云计算中的应用 + 第四单元复习",
        "questions": [
            {"type": "选择", "q": "Linux在云计算中的主要应用是？", "options": ["A. 只用于桌面", "B. 作为云服务器的操作系统", "C. 替代Windows", "D. 只用于手机"], "answer": "B", "explain": "Linux是云服务器的主流操作系统。"},
            {"type": "选择", "q": "复习：Hypervisor也称为？", "options": ["A. 虚拟机监视器(VMM)", "B. 操作系统", "C. 数据库", "D. 网络"], "answer": "A", "explain": "Hypervisor = VMM。"},
            {"type": "选择", "q": "复习：Linux根目录是？", "options": ["A. C:\\", "B. /", "C. \\", "D. root"], "answer": "B", "explain": "Linux根目录是/。"},
            {"type": "选择", "q": "复习：ls命令的作用是？", "options": ["A. 创建目录", "B. 列出目录内容", "C. 删除文件", "D. 切换目录"], "answer": "B", "explain": "ls = list，列出目录内容。"},
            {"type": "选择", "q": "复习：Docker容器与虚拟机的主要区别？", "options": ["A. 容器共享宿主OS内核", "B. 容器更重", "C. 没有区别", "D. 容器需要独立OS"], "answer": "A", "explain": "容器共享内核，轻量快速。"},
            {"type": "选择", "q": "以下哪个Linux命令可以查看文件内容？", "options": ["A. ls", "B. cd", "C. cat", "D. mkdir"], "answer": "C", "explain": "cat命令可以显示文件内容。"},
            {"type": "判断", "q": "复习：Linux是开源操作系统。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "Linux是开源的。"},
            {"type": "判断", "q": "复习：虚拟化是云计算的核心基础技术。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "虚拟化实现资源池化。"},
            {"type": "判断", "q": "复习：chmod命令修改文件所有者。", "options": ["A. 正确", "B. 错误"], "answer": "B", "explain": "chmod修改权限，chown修改所有者。"},
            {"type": "判断", "q": "SSH是远程管理Linux服务器的常用协议。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "SSH(Secure Shell)默认端口22，是远程管理Linux的标准方式。"},
            {"type": "填空", "q": "复习：Linux根目录是____。", "answer": "/", "explain": "根目录/。"},
            {"type": "填空", "q": "复习：虚拟化实现了资源____化。", "answer": "池", "explain": "Resource Pooling。"},
        ]
    },
    # ===== 第五单元 课时49-60 =====
    {
        "num": 49, "chapter": "第11章 Web服务", "title": "Web服务概述",
        "questions": [
            {"type": "选择", "q": "Web服务基于什么模型？", "options": ["A. P2P对等模型", "B. 客户端-服务器(C/S)模型", "C. 主从模型", "D. 分布式模型"], "answer": "B", "explain": "Web服务基于客户端(浏览器)-服务器(Web服务器)模型。"},
            {"type": "选择", "q": "静态网页和动态网页的主要区别是？", "options": ["A. 静态网页更漂亮", "B. 静态网页内容固定，动态网页内容根据请求生成", "C. 动态网页有动画", "D. 没有区别"], "answer": "B", "explain": "静态网页是预先写好的HTML文件，动态网页由服务器端程序根据请求实时生成。"},
            {"type": "选择", "q": "以下哪个是Web服务器软件？", "options": ["A. Chrome", "B. Apache", "C. MySQL", "D. Windows"], "answer": "B", "explain": "Apache是Web服务器软件，Chrome是浏览器(客户端)，MySQL是数据库。"},
            {"type": "选择", "q": "URL中https://www.example.com/page.html的域名部分是？", "options": ["A. https", "B. www.example.com", "C. /page.html", "D. com"], "answer": "B", "explain": "URL结构：协议://域名/路径，www.example.com是域名。"},
            {"type": "选择", "q": "Web服务在云计算中的作用是？", "options": ["A. 不重要", "B. 云计算通过Web服务向用户提供界面和API", "C. 只用于浏览网页", "D. 可有可无"], "answer": "B", "explain": "云计算的管理控制台、API接口等都基于Web服务。"},
            {"type": "选择", "q": "以下哪个是动态网站的技术？", "options": ["A. HTML", "B. CSS", "C. PHP", "D. JPEG"], "answer": "C", "explain": "PHP是服务器端脚本语言，用于生成动态网页。HTML/CSS/JPEG都是静态内容。"},
            {"type": "判断", "q": "浏览器属于Web服务中的客户端。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "浏览器发送HTTP请求，是客户端；Web服务器响应请求，是服务端。"},
            {"type": "判断", "q": "静态网页的内容不会随用户或时间变化。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "静态网页是固定的HTML文件，内容对所有用户都一样。"},
            {"type": "判断", "q": "Apache和Nginx都是Web服务器软件。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "Apache和Nginx是最流行的两款Web服务器软件。"},
            {"type": "判断", "q": "动态网页一定比静态网页加载慢。", "options": ["A. 正确", "B. 错误"], "answer": "B", "explain": "现代动态网页有缓存机制，不一定慢；且动态网页可以个性化内容。"},
            {"type": "填空", "q": "Web服务基于____-服务器模型（填两个字）。", "answer": "客户端", "explain": "Client-Server模型。"},
            {"type": "填空", "q": "网页内容固定不变的是____态网页（填一个字）。", "answer": "静", "explain": "静态网页内容固定。"},
        ]
    },
    {
        "num": 50, "chapter": "第11章 Web服务", "title": "HTTP协议详解",
        "questions": [
            {"type": "选择", "q": "HTTP协议是无状态协议，\"无状态\"的含义是？", "options": ["A. 服务器没有状态", "B. 服务器不保留客户端的请求状态", "C. 不能传输数据", "D. 没有连接"], "answer": "B", "explain": "无状态指服务器处理完请求后不保留客户端信息，每次请求都是独立的。"},
            {"type": "选择", "q": "HTTP请求方法GET的作用是？", "options": ["A. 删除资源", "B. 获取资源", "C. 创建资源", "D. 修改资源"], "answer": "B", "explain": "GET方法用于从服务器获取/查询资源，是最常用的HTTP方法。"},
            {"type": "选择", "q": "HTTP请求方法POST的作用是？", "options": ["A. 获取数据", "B. 向服务器提交数据", "C. 删除数据", "D. 重定向"], "answer": "B", "explain": "POST方法用于向服务器提交数据(如提交表单、上传文件)。"},
            {"type": "选择", "q": "HTTP状态码200表示什么？", "options": ["A. 服务器错误", "B. 请求成功", "C. 资源未找到", "D. 重定向"], "answer": "B", "explain": "200 OK表示请求成功处理。"},
            {"type": "选择", "q": "HTTP状态码404表示什么？", "options": ["A. 请求成功", "B. 资源未找到", "C. 服务器内部错误", "D. 未授权"], "answer": "B", "explain": "404 Not Found表示请求的资源不存在。"},
            {"type": "选择", "q": "HTTP状态码500表示什么？", "options": ["A. 请求成功", "B. 资源未找到", "C. 服务器内部错误", "D. 重定向"], "answer": "C", "explain": "500 Internal Server Error表示服务器端发生错误。"},
            {"type": "判断", "q": "HTTP是无状态协议。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "HTTP本身不保存客户端状态，用Cookie/Session来维护。"},
            {"type": "判断", "q": "GET请求可以通过URL传递参数。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "GET参数附在URL后面(如?name=value)，POST参数在请求体中。"},
            {"type": "判断", "q": "HTTP状态码以5开头表示客户端错误。", "options": ["A. 正确", "B. 错误"], "answer": "B", "explain": "5xx表示服务器端错误，4xx才是客户端错误(如404)。"},
            {"type": "判断", "q": "HTTP请求由请求行、请求头、请求体组成。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "HTTP请求结构：请求行(方法+URL+版本)、请求头、空行、请求体。"},
            {"type": "填空", "q": "HTTP请求方法____用于获取资源（填三个字母）。", "answer": "GET", "explain": "GET = 获取资源。"},
            {"type": "填空", "q": "HTTP状态码____表示资源未找到（填三个数字）。", "answer": "404", "explain": "404 Not Found。"},
        ]
    },
    {
        "num": 51, "chapter": "第11章 Web服务", "title": "Web服务器Apache",
        "questions": [
            {"type": "选择", "q": "Apache HTTP Server是哪种软件？", "options": ["A. 浏览器", "B. Web服务器", "C. 数据库", "D. 操作系统"], "answer": "B", "explain": "Apache是最老牌的开源Web服务器软件。"},
            {"type": "选择", "q": "Apache的配置文件通常是？", "options": ["A. httpd.conf", "B. config.txt", "C. settings.json", "D. web.xml"], "answer": "A", "explain": "Apache主配置文件是httpd.conf(或apache2.conf)。"},
            {"type": "选择", "q": "Apache的虚拟主机功能允许？", "options": ["A. 虚拟机管理", "B. 一台服务器上托管多个网站", "C. 网络虚拟化", "D. 存储虚拟化"], "answer": "B", "explain": "虚拟主机(Virtual Host)让一台Apache服务器根据域名区分并服务多个网站。"},
            {"type": "选择", "q": "Apache的默认监听端口是？", "options": ["A. 22", "B. 80", "C. 443", "D. 8080"], "answer": "B", "explain": "HTTP默认端口80，Apache监听80端口。"},
            {"type": "选择", "q": "Apache的主要特点不包括？", "options": ["A. 开源免费", "B. 模块化设计", "C. 跨平台", "D. 只能运行在Windows上"], "answer": "D", "explain": "Apache可以运行在Linux/Windows/macOS等多种平台上。"},
            {"type": "选择", "q": "LAMP架构中的A代表什么？", "options": ["A. Android", "B. Apache", "C. AWS", "D. Azure"], "answer": "B", "explain": "LAMP = Linux + Apache + MySQL + PHP。"},
            {"type": "判断", "q": "Apache是开源的Web服务器软件。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "Apache HTTP Server是Apache软件基金会的开源项目。"},
            {"type": "判断", "q": "Apache可以通过虚拟主机功能在一台服务器上运行多个网站。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "虚拟主机是Apache的重要功能。"},
            {"type": "判断", "q": "Apache只能运行在Linux上。", "options": ["A. 正确", "B. 错误"], "answer": "B", "explain": "Apache是跨平台的，也支持Windows。"},
            {"type": "判断", "q": "LAMP架构中的MySQL可以用其他数据库替换。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "可以用PostgreSQL/MariaDB等替换MySQL，形成LAPP等变体。"},
            {"type": "填空", "q": "LAMP中的A代表____。", "answer": "Apache", "explain": "Linux+Apache+MySQL+PHP。"},
            {"type": "填空", "q": "Apache的默认端口是____。", "answer": "80", "explain": "HTTP默认端口80。"},
        ]
    },
    {
        "num": 52, "chapter": "第11章 Web服务", "title": "Web服务器Nginx",
        "questions": [
            {"type": "选择", "q": "Nginx相比Apache的主要优势是？", "options": ["A. 功能更多", "B. 高并发处理能力更强，资源占用少", "C. 更安全", "D. 更容易配置"], "answer": "B", "explain": "Nginx采用事件驱动架构，高并发性能优异，内存占用少。"},
            {"type": "选择", "q": "反向代理的作用是？", "options": ["A. 加速客户端", "B. 代理服务器接收请求后转发给后端服务器", "C. 加密数据", "D. 压缩网页"], "answer": "B", "explain": "反向代理：客户端请求到代理服务器，代理转发给后端真实服务器，隐藏后端。"},
            {"type": "选择", "q": "负载均衡的作用是？", "options": ["A. 加密数据", "B. 将请求分发到多台服务器，避免单台过载", "C. 压缩数据", "D. 存储数据"], "answer": "B", "explain": "负载均衡将流量分配到多台服务器，提高整体处理能力和可用性。"},
            {"type": "选择", "q": "Nginx常用于什么场景？", "options": ["A. 个人办公", "B. 反向代理和负载均衡", "C. 数据库管理", "D. 图像处理"], "answer": "B", "explain": "Nginx最常用于反向代理、负载均衡和静态内容服务。"},
            {"type": "选择", "q": "以下哪个是Nginx的特点？", "options": ["A. 线程驱动", "B. 事件驱动、高并发", "C. 只支持静态文件", "D. 闭源付费"], "answer": "B", "explain": "Nginx采用事件驱动异步架构，高并发性能突出。"},
            {"type": "选择", "q": "LNMP架构中的N代表什么？", "options": ["A. Network", "B. Nginx", "C. Node.js", "D. .NET"], "answer": "B", "explain": "LNMP = Linux + Nginx + MySQL + PHP。"},
            {"type": "判断", "q": "Nginx在高并发场景下性能优于Apache。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "Nginx事件驱动架构在高并发时资源占用远低于Apache。"},
            {"type": "判断", "q": "反向代理可以隐藏后端真实服务器的IP地址。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "客户端只能看到代理服务器的IP，后端服务器IP被隐藏。"},
            {"type": "判断", "q": "负载均衡可以提高网站的可用性。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "一台服务器故障时，流量自动转到其他服务器。"},
            {"type": "判断", "q": "Nginx是闭源商业软件。", "options": ["A. 正确", "B. 错误"], "answer": "B", "explain": "Nginx有开源版本(免费)和商业版本(NGINX Plus)。"},
            {"type": "填空", "q": "____代理是代理服务器接收请求后转发给后端服务器。", "answer": "反向", "explain": "Reverse Proxy反向代理。"},
            {"type": "填空", "q": "____均衡将请求分发到多台服务器避免过载。", "answer": "负载", "explain": "Load Balancing负载均衡。"},
        ]
    },
    {
        "num": 53, "chapter": "第11章 Web服务", "title": "Web应用与动态网站",
        "questions": [
            {"type": "选择", "q": "LAMP架构包括哪些组件？", "options": ["A. Linux+Apache+MySQL+PHP", "B. Linux+AWS+MongoDB+Python", "C. Windows+Apache+MySQL+PHP", "D. Linux+Android+MySQL+PHP"], "answer": "A", "explain": "LAMP = Linux + Apache + MySQL + PHP。"},
            {"type": "选择", "q": "以下哪个是服务器端编程语言？", "options": ["A. HTML", "B. CSS", "C. PHP", "D. JavaScript(前端)"], "answer": "C", "explain": "PHP是服务器端语言，HTML/CSS/前端JavaScript运行在浏览器中。"},
            {"type": "选择", "q": "动态网站的\"动态\"主要指什么？", "options": ["A. 页面有动画效果", "B. 内容根据用户请求实时生成", "C. 页面会自动刷新", "D. 页面有视频"], "answer": "B", "explain": "动态网页内容由服务器端程序根据请求参数、数据库等实时生成。"},
            {"type": "选择", "q": "以下哪个不是Web开发语言？", "options": ["A. PHP", "B. Python", "C. Java", "D. Photoshop"], "answer": "D", "explain": "Photoshop是图像处理软件，不是编程语言。"},
            {"type": "选择", "q": "Cookie的作用是？", "options": ["A. 加密数据", "B. 在客户端保存用户状态信息", "C. 加速网页", "D. 压缩文件"], "answer": "B", "explain": "Cookie在浏览器端保存少量数据，用于维持HTTP会话状态。"},
            {"type": "选择", "q": "Session和Cookie的主要区别是？", "options": ["A. 没有区别", "B. Cookie存在客户端，Session存在服务器端", "C. Cookie更安全", "D. Session更慢"], "answer": "B", "explain": "Cookie存储在浏览器端，Session存储在服务器端，Session更安全。"},
            {"type": "判断", "q": "LAMP是常见的动态网站技术架构。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "LAMP是经典的Web开发架构。"},
            {"type": "判断", "q": "PHP代码在服务器端执行。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "PHP是服务器端脚本语言，在服务器上执行后生成HTML发给浏览器。"},
            {"type": "判断", "q": "JavaScript只能在浏览器端运行。", "options": ["A. 正确", "B. 错误"], "answer": "B", "explain": "JavaScript也可在服务器端运行(Node.js)。"},
            {"type": "判断", "q": "Session比Cookie更安全，因为数据存在服务器端。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "Session数据在服务器端，客户端无法直接查看或修改。"},
            {"type": "填空", "q": "LAMP中M代表____（填数据库名）。", "answer": "MySQL", "explain": "Linux+Apache+MySQL+PHP。"},
            {"type": "填空", "q": "____保存在客户端浏览器中，用于维持会话状态。", "answer": "Cookie", "explain": "Cookie在客户端，Session在服务器端。"},
        ]
    },
    {
        "num": 54, "chapter": "第11章 Web服务", "title": "Web安全与CDN + 第十一章复习",
        "questions": [
            {"type": "选择", "q": "CDN的全称是？", "options": ["A. 内容分发网络", "B. 云数据网络", "C. 内容压缩工具", "D. 网络加速协议"], "answer": "A", "explain": "CDN = Content Delivery Network，内容分发网络。"},
            {"type": "选择", "q": "CDN的主要作用是？", "options": ["A. 加密数据", "B. 将内容缓存到全球节点，加速用户访问", "C. 压缩文件", "D. 防黑客"], "answer": "B", "explain": "CDN将静态内容缓存在离用户最近的节点，减少延迟。"},
            {"type": "选择", "q": "XSS攻击是指？", "options": ["A. 破坏服务器硬件", "B. 在网页中注入恶意脚本代码", "C. 窃取数据库", "D. DDoS攻击"], "answer": "B", "explain": "XSS(跨站脚本攻击)在网页中注入恶意JavaScript，窃取用户信息。"},
            {"type": "选择", "q": "复习：HTTP默认端口是？", "options": ["A. 22", "B. 80", "C. 443", "D. 8080"], "answer": "B", "explain": "HTTP默认端口80。"},
            {"type": "选择", "q": "复习：Nginx的主要优势是？", "options": ["A. 功能最多", "B. 高并发性能强", "C. 更安全", "D. 更简单"], "answer": "B", "explain": "Nginx事件驱动，高并发性能突出。"},
            {"type": "选择", "q": "复习：HTTP状态码404表示？", "options": ["A. 请求成功", "B. 资源未找到", "C. 服务器错误", "D. 重定向"], "answer": "B", "explain": "404 Not Found。"},
            {"type": "判断", "q": "CDN可以加速用户访问网站的速度。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "CDN将内容缓存到离用户更近的节点。"},
            {"type": "判断", "q": "复习：Apache和Nginx都是Web服务器软件。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "两者都是流行的Web服务器。"},
            {"type": "判断", "q": "复习：GET请求用于向服务器提交数据。", "options": ["A. 正确", "B. 错误"], "answer": "B", "explain": "GET用于获取资源，POST用于提交数据。"},
            {"type": "判断", "q": "复习：LAMP架构中的P代表PHP。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "LAMP = Linux+Apache+MySQL+PHP。"},
            {"type": "填空", "q": "CDN的中文全称是内容____网络。", "answer": "分发", "explain": "Content Delivery Network。"},
            {"type": "填空", "q": "复习：HTTP状态码____表示请求成功。", "answer": "200", "explain": "200 OK。"},
        ]
    },
    {
        "num": 55, "chapter": "第12章 公有云平台", "title": "公有云平台概述",
        "questions": [
            {"type": "选择", "q": "公有云平台的核心特点是？", "options": ["A. 只对特定组织开放", "B. 对公众开放，按需使用", "C. 完全免费", "D. 不需要网络"], "answer": "B", "explain": "公有云对公众开放，用户注册即可使用，按需付费。"},
            {"type": "选择", "q": "以下哪个不是公有云的服务类型？", "options": ["A. IaaS", "B. PaaS", "C. SaaS", "D. DaaS(桌面即服务常见但非NIST标准)"], "answer": "D", "explain": "NIST定义的三种服务模型是IaaS/PaaS/SaaS。"},
            {"type": "选择", "q": "公有云相比私有云的主要优势是？", "options": ["A. 更安全", "B. 成本低、弹性好、无需运维硬件", "C. 更可控", "D. 更合规"], "answer": "B", "explain": "公有云无需自购硬件、按需付费、弹性伸缩。"},
            {"type": "选择", "q": "公有云的计费模式通常包括？", "options": ["A. 只能按年付费", "B. 按量付费和包年包月", "C. 完全免费", "D. 按人收费"], "answer": "B", "explain": "公有云支持按量(按小时/秒)和包年包月两种计费模式。"},
            {"type": "选择", "q": "以下哪个是全球最大的公有云平台？", "options": ["A. 阿里云", "B. AWS", "C. 腾讯云", "D. 华为云"], "answer": "B", "explain": "AWS是全球市场份额最大的公有云。"},
            {"type": "选择", "q": "公有云用户的数据安全性主要依靠？", "options": ["A. 用户自己加密", "B. 云厂商的隔离和安全措施", "C. 物理隔离", "D. 不需要保护"], "answer": "B", "explain": "公有云通过虚拟化隔离、加密、访问控制等技术保障数据安全。"},
            {"type": "判断", "q": "公有云对公众开放，任何人都可以注册使用。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "公有云面向公众开放。"},
            {"type": "判断", "q": "公有云用户需要自己购买和维护物理服务器。", "options": ["A. 正确", "B. 错误"], "answer": "B", "explain": "公有云由厂商负责硬件，用户只需使用。"},
            {"type": "判断", "q": "公有云支持弹性伸缩，可以根据需求增减资源。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "弹性是公有云的核心优势。"},
            {"type": "判断", "q": "公有云和私有云完全没有交集。", "options": ["A. 正确", "B. 错误"], "answer": "B", "explain": "混合云结合了公有云和私有云。"},
            {"type": "填空", "q": "____云对公众开放，按需使用（填两个字）。", "answer": "公有", "explain": "Public Cloud。"},
            {"type": "填空", "q": "全球最大的公有云平台是____。", "answer": "AWS", "explain": "Amazon Web Services。"},
        ]
    },
    {
        "num": 56, "chapter": "第12章 公有云平台", "title": "阿里云实践",
        "questions": [
            {"type": "选择", "q": "阿里云ECS属于哪种云服务？", "options": ["A. SaaS", "B. PaaS", "C. IaaS", "D. DaaS"], "answer": "C", "explain": "ECS(弹性计算服务)提供虚拟服务器，属于IaaS。"},
            {"type": "选择", "q": "阿里云OSS是什么服务？", "options": ["A. 虚拟服务器", "B. 对象存储", "C. 数据库", "D. CDN"], "answer": "B", "explain": "OSS(Object Storage Service)是阿里云的对象存储服务，用于存储文件/图片/视频等。"},
            {"type": "选择", "q": "阿里云RDS是什么服务？", "options": ["A. 虚拟服务器", "B. 云数据库", "C. 对象存储", "D. CDN"], "answer": "B", "explain": "RDS(Relational Database Service)是阿里云的关系型数据库服务。"},
            {"type": "选择", "q": "阿里云OSS中存储的文件通过什么访问？", "options": ["A. 只能通过FTP", "B. 通过URL/API访问", "C. 只能通过U盘", "D. 不能访问"], "answer": "B", "explain": "OSS对象存储通过HTTP URL或API访问，支持公开或私有读取。"},
            {"type": "选择", "q": "阿里云在中国云计算市场的地位是？", "options": ["A. 第二", "B. 第一", "C. 第三", "D. 不在前五"], "answer": "B", "explain": "阿里云是中国公有云市场份额第一。"},
            {"type": "选择", "q": "阿里云ECS创建后可以做什么？", "options": ["A. 只能浏览网页", "B. 安装操作系统和软件、部署网站", "C. 只能存储文件", "D. 不能联网"], "answer": "B", "explain": "ECS是虚拟服务器，可以像物理机一样安装软件和部署应用。"},
            {"type": "判断", "q": "阿里云ECS支持按量付费和包年包月两种计费方式。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "这是公有云的标准计费模式。"},
            {"type": "判断", "q": "阿里云OSS适合存储大量图片和视频文件。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "OSS是对象存储，适合存储非结构化数据(图片/视频/文档)。"},
            {"type": "判断", "q": "阿里云RDS需要用户自己安装数据库软件。", "options": ["A. 正确", "B. 错误"], "answer": "B", "explain": "RDS是托管数据库服务，厂商负责安装、备份、升级等运维。"},
            {"type": "判断", "q": "阿里云ECS购买后不能更换操作系统。", "options": ["A. 正确", "B. 错误"], "answer": "B", "explain": "ECS支持更换操作系统(需重装系统)。"},
            {"type": "填空", "q": "阿里云的弹性计算服务缩写是____。", "answer": "ECS", "explain": "ECS = Elastic Compute Service。"},
            {"type": "填空", "q": "阿里云的对象存储服务缩写是____。", "answer": "OSS", "explain": "OSS = Object Storage Service。"},
        ]
    },
    {
        "num": 57, "chapter": "第12章 公有云平台", "title": "腾讯云实践",
        "questions": [
            {"type": "选择", "q": "腾讯云CVM对应阿里云的哪个产品？", "options": ["A. OSS", "B. ECS", "C. RDS", "D. CDN"], "answer": "B", "explain": "腾讯云CVM(云服务器)和阿里云ECS都是IaaS虚拟服务器。"},
            {"type": "选择", "q": "腾讯云的特色优势领域是？", "options": ["A. 电商", "B. 音视频和游戏", "C. 政务", "D. 硬件"], "answer": "B", "explain": "腾讯在社交、游戏、音视频领域有天然优势，腾讯云继承了这些优势。"},
            {"type": "选择", "q": "腾讯云COS是什么服务？", "options": ["A. 虚拟服务器", "B. 对象存储", "C. 数据库", "D. CDN"], "answer": "B", "explain": "COS(Cloud Object Storage)是腾讯云的对象存储服务。"},
            {"type": "选择", "q": "以下哪个是腾讯云的产品？", "options": ["A. ECS", "B. CVM", "C. EC2", "D. S3"], "answer": "B", "explain": "CVM是腾讯云的云服务器产品，ECS是阿里云，EC2是AWS，S3也是AWS。"},
            {"type": "选择", "q": "阿里云ECS和腾讯云CVM的本质区别是？", "options": ["A. 功能完全不同", "B. 本质相同，都是IaaS虚拟服务器", "C. CVM更好", "D. ECS更好"], "answer": "B", "explain": "都是IaaS层的虚拟服务器产品，功能本质相同。"},
            {"type": "选择", "q": "腾讯云在哪个领域有独特优势？", "options": ["A. 操作系统", "B. 直播和即时通讯", "C. 硬件芯片", "D. 数据库"], "answer": "B", "explain": "腾讯的社交和游戏基因使腾讯云在直播、IM(即时通讯)领域有优势。"},
            {"type": "判断", "q": "腾讯云CVM和阿里云ECS都是IaaS服务。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "两者都是虚拟服务器IaaS产品。"},
            {"type": "判断", "q": "腾讯云在游戏行业有较强的竞争力。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "腾讯是游戏巨头，腾讯云在游戏云服务领域有优势。"},
            {"type": "判断", "q": "不同云厂商的同类产品(IaaS)功能本质类似。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "核心功能一致，差异在于价格、性能、特色服务和生态。"},
            {"type": "判断", "q": "腾讯云COS和阿里云OSS都是对象存储服务。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "两者都是对象存储服务，功能类似。"},
            {"type": "填空", "q": "腾讯云的云服务器产品叫____。", "answer": "CVM", "explain": "CVM = Cloud Virtual Machine。"},
            {"type": "填空", "q": "腾讯云的特色优势领域是音视频和____。", "answer": "游戏", "explain": "腾讯在游戏和社交领域的优势延伸到云服务。"},
        ]
    },
    {
        "num": 58, "chapter": "第12章 公有云平台", "title": "AWS云服务",
        "questions": [
            {"type": "选择", "q": "AWS是哪个公司的云服务？", "options": ["A. 谷歌", "B. 微软", "C. 亚马逊", "D. IBM"], "answer": "C", "explain": "AWS = Amazon Web Services，亚马逊公司的云服务。"},
            {"type": "选择", "q": "AWS EC2对应阿里云的哪个产品？", "options": ["A. OSS", "B. ECS", "C. RDS", "D. CDN"], "answer": "B", "explain": "EC2(Elastic Compute Cloud)和ECS都是IaaS虚拟服务器。"},
            {"type": "选择", "q": "AWS S3是什么服务？", "options": ["A. 虚拟服务器", "B. 对象存储", "C. 数据库", "D. CDN"], "answer": "B", "explain": "S3(Simple Storage Service)是AWS的对象存储服务。"},
            {"type": "选择", "q": "AWS RDS是什么服务？", "options": ["A. 虚拟服务器", "B. 云数据库", "C. 对象存储", "D. CDN"], "answer": "B", "explain": "RDS(Relational Database Service)是AWS的关系型数据库服务。"},
            {"type": "选择", "q": "AWS与国内云厂商的主要区别是？", "options": ["A. AWS功能更多", "B. AWS全球覆盖广，国内云在中国有本地化优势", "C. AWS更便宜", "D. 没有区别"], "answer": "B", "explain": "AWS全球区域最多，国内云厂商在中国有合规和本地化服务优势。"},
            {"type": "选择", "q": "AWS是全球第几个推出大规模公有云服务的厂商？", "options": ["A. 第一个", "B. 第二个", "C. 第三个", "D. 不确定"], "answer": "A", "explain": "AWS于2006年推出EC2，开创了公有云服务模式。"},
            {"type": "判断", "q": "AWS是全球最大的公有云服务商。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "AWS全球市场份额第一。"},
            {"type": "判断", "q": "AWS EC2和阿里云ECS都是IaaS服务。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "两者都是弹性计算服务(IaaS)。"},
            {"type": "判断", "q": "AWS在中国有本地区域(由光环新网和西云数据运营)。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "AWS中国区由本地合作伙伴运营，符合中国法规。"},
            {"type": "判断", "q": "AWS S3和阿里云OSS都是对象存储服务。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "两者功能类似，都是对象存储。"},
            {"type": "填空", "q": "AWS是____公司的云服务（填两个字）。", "answer": "亚马逊", "explain": "Amazon Web Services。"},
            {"type": "填空", "q": "AWS的虚拟服务器产品叫____（填三个字符）。", "answer": "EC2", "explain": "EC2 = Elastic Compute Cloud。"},
        ]
    },
    {
        "num": 59, "chapter": "第12章 公有云平台", "title": "云服务选型与对比",
        "questions": [
            {"type": "选择", "q": "云服务选型时最重要的考量因素是？", "options": ["A. 价格最低", "B. 满足业务需求并平衡成本、性能、安全", "C. 品牌最大", "D. 功能最多"], "answer": "B", "explain": "选型需综合考虑业务需求、成本、性能、安全性、合规性等因素。"},
            {"type": "选择", "q": "个人小型网站适合选择什么云服务？", "options": ["A. 自建私有云", "B. 公有云低配虚拟服务器", "C. 混合云", "D. 社区云"], "answer": "B", "explain": "小型网站流量小，公有云低配ECS/CVM即可满足，成本低。"},
            {"type": "选择", "q": "银行核心系统适合选择什么云？", "options": ["A. 公有云", "B. 私有云", "C. 社区云", "D. 不用云"], "answer": "B", "explain": "银行对安全合规要求极高，私有云最合适。"},
            {"type": "选择", "q": "需要弹性应对流量高峰的电商适合？", "options": ["A. 纯私有云", "B. 纯公有云", "C. 混合云", "D. 不用云"], "answer": "C", "explain": "核心数据放私有云保安全，高峰用公有云保弹性。"},
            {"type": "选择", "q": "云迁移的一般流程是？", "options": ["A. 评估->规划->迁移->优化", "B. 直接迁移", "C. 先买硬件再迁移", "D. 迁移->评估->规划"], "answer": "A", "explain": "标准流程：评估现状->规划架构->执行迁移->持续优化。"},
            {"type": "选择", "q": "以下哪个不是云服务选型的考量因素？", "options": ["A. 成本", "B. 性能", "C. 安全合规", "D. 服务器颜色"], "answer": "D", "explain": "服务器颜色与选型无关。"},
            {"type": "判断", "q": "不同业务场景应选择不同的云服务方案。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "没有万能方案，需根据业务特点选择。"},
            {"type": "判断", "q": "价格是云服务选型的唯一标准。", "options": ["A. 正确", "B. 错误"], "answer": "B", "explain": "需综合考虑成本、性能、安全、合规等多个因素。"},
            {"type": "判断", "q": "混合云适合需要兼顾安全性和弹性的企业。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "混合云兼顾私有云安全和公有云弹性。"},
            {"type": "判断", "q": "企业上云后不需要持续优化。", "options": ["A. 正确", "B. 错误"], "answer": "B", "explain": "上云后需持续监控和优化资源使用，控制成本。"},
            {"type": "填空", "q": "银行核心系统适合使用____云。", "answer": "私有", "explain": "私有云安全可控。"},
            {"type": "填空", "q": "云选型需平衡成本、性能和____。", "answer": "安全", "explain": "安全是云选型的重要考量。"},
        ]
    },
    {
        "num": 60, "chapter": "第12章 公有云平台", "title": "公有云安全与合规 + 第五单元复习",
        "questions": [
            {"type": "选择", "q": "云安全合规主要关注什么？", "options": ["A. 服务器速度", "B. 数据保护法规遵从(如数据不出境)", "C. 服务器颜色", "D. 网页美观"], "answer": "B", "explain": "云安全合规关注数据保护、隐私法规、数据主权(如数据不出境)等法律合规要求。"},
            {"type": "选择", "q": "复习：HTTP协议默认端口是？", "options": ["A. 22", "B. 80", "C. 443", "D. 8080"], "answer": "B", "explain": "HTTP默认端口80。"},
            {"type": "选择", "q": "复习：阿里云ECS属于什么服务？", "options": ["A. SaaS", "B. PaaS", "C. IaaS", "D. DaaS"], "answer": "C", "explain": "ECS是IaaS。"},
            {"type": "选择", "q": "复习：Nginx的主要用途是？", "options": ["A. 数据库", "B. 反向代理和负载均衡", "C. 操作系统", "D. 杀毒"], "answer": "B", "explain": "Nginx常用于反向代理和负载均衡。"},
            {"type": "选择", "q": "复习：CDN的作用是？", "options": ["A. 加密数据", "B. 内容分发加速", "C. 存储数据", "D. 删除数据"], "answer": "B", "explain": "CDN将内容缓存到全球节点加速访问。"},
            {"type": "选择", "q": "中国的数据安全法规要求？", "options": ["A. 数据可以随意出境", "B. 重要数据需在境内存储", "C. 不需要遵守法规", "D. 所有数据必须公开"], "answer": "B", "explain": "中国《数据安全法》《个人信息保护法》等要求数据合规处理。"},
            {"type": "判断", "q": "复习：LAMP架构包括Linux、Apache、MySQL、PHP。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "LAMP是经典Web架构。"},
            {"type": "判断", "q": "复习：AWS是全球最大的公有云服务商。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "AWS全球市场份额第一。"},
            {"type": "判断", "q": "公有云用户数据需遵守所在国家的数据保护法规。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "数据合规是云服务的重要要求。"},
            {"type": "判断", "q": "复习：反向代理可以隐藏后端服务器IP。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "反向代理隐藏后端真实服务器。"},
            {"type": "填空", "q": "复习：HTTP状态码404表示资源____。", "answer": "未找到", "explain": "404 Not Found。"},
            {"type": "填空", "q": "复习：AWS的虚拟服务器叫____。", "answer": "EC2", "explain": "EC2 = Elastic Compute Cloud。"},
        ]
    },
    # ===== 第六单元 课时61-72 =====
    {
        "num": 61, "chapter": "第13章 私有云平台", "title": "私有云概述",
        "questions": [
            {"type": "选择", "q": "私有云的主要特点是？", "options": ["A. 对公众开放", "B. 仅供单一组织内部使用", "C. 完全免费", "D. 不需要网络"], "answer": "B", "explain": "私有云由单一组织独占使用，不对外开放。"},
            {"type": "选择", "q": "私有云相比公有云的主要优势是？", "options": ["A. 成本更低", "B. 数据完全可控、安全性更高", "C. 弹性更好", "D. 更方便"], "answer": "B", "explain": "私有云数据在组织内部，完全可控，安全性更高。"},
            {"type": "选择", "q": "以下哪个场景最适合使用私有云？", "options": ["A. 个人博客", "B. 银行核心交易系统", "C. 免费在线翻译", "D. 公开视频网站"], "answer": "B", "explain": "银行对数据安全要求极高，需要完全掌控，私有云最合适。"},
            {"type": "选择", "q": "私有云的部署方式包括？", "options": ["A. 只能自建", "B. 只能托管", "C. 可自建也可托管", "D. 必须用公有云"], "answer": "C", "explain": "私有云可在企业自有机房部署，也可托管在第三方数据中心。"},
            {"type": "选择", "q": "私有云的主要缺点是？", "options": ["A. 不安全", "B. 建设成本高、弹性不如公有云", "C. 速度慢", "D. 不支持网络"], "answer": "B", "explain": "私有云需自购硬件、自建运维团队，初始投入大，弹性受硬件限制。"},
            {"type": "选择", "q": "学校机房搭建的仅供校内使用的云属于？", "options": ["A. 公有云", "B. 私有云", "C. 社区云", "D. 混合云"], "answer": "B", "explain": "仅供单一组织(学校)内部使用的云是私有云。"},
            {"type": "判断", "q": "私有云仅供单一组织内部使用。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "私有云的核心特征是单一组织独占。"},
            {"type": "判断", "q": "私有云的建设成本通常低于公有云。", "options": ["A. 正确", "B. 错误"], "answer": "B", "explain": "私有云需自购硬件和运维团队，初始成本高于公有云。"},
            {"type": "判断", "q": "私有云可以部署在企业自有机房或第三方数据中心。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "私有云部署方式灵活，可自建也可托管。"},
            {"type": "判断", "q": "私有云的弹性扩展能力通常优于公有云。", "options": ["A. 正确", "B. 错误"], "answer": "B", "explain": "私有云受物理硬件限制，弹性不如公有云(可即时扩容)。"},
            {"type": "填空", "q": "____云仅供单一组织内部使用（填两个字）。", "answer": "私有", "explain": "Private Cloud。"},
            {"type": "填空", "q": "私有云的主要优势是数据完全____。", "answer": "可控", "explain": "数据在组织内部，完全可控。"},
        ]
    },
    {
        "num": 62, "chapter": "第13章 私有云平台", "title": "OpenStack架构",
        "questions": [
            {"type": "选择", "q": "OpenStack是什么？", "options": ["A. 操作系统", "B. 开源私有云管理平台", "C. 数据库", "D. 浏览器"], "answer": "B", "explain": "OpenStack是开源的云计算管理平台，用于构建和管理私有云。"},
            {"type": "选择", "q": "OpenStack的授权方式是？", "options": ["A. 商业闭源", "B. 开源(Apache 2.0许可)", "C. 免费但不开源", "D. 试用版"], "answer": "B", "explain": "OpenStack是开源项目，遵循Apache 2.0许可证。"},
            {"type": "选择", "q": "OpenStack在私有云中的地位是？", "options": ["A. 不重要", "B. 最流行的开源私有云平台", "C. 只用于公有云", "D. 已被淘汰"], "answer": "B", "explain": "OpenStack是全球最流行的开源私有云构建方案。"},
            {"type": "选择", "q": "OpenStack整体架构的特点是？", "options": ["A. 单一组件", "B. 模块化、分布式架构", "C. 不能扩展", "D. 只能运行在Windows"], "answer": "B", "explain": "OpenStack由多个独立组件组成，分布式部署，高度可扩展。"},
            {"type": "选择", "q": "OpenStack的底层虚拟化技术通常使用？", "options": ["A. VMware Workstation", "B. KVM", "C. VirtualBox", "D. Docker"], "answer": "B", "explain": "KVM是OpenStack最常用的底层虚拟化驱动。"},
            {"type": "选择", "q": "以下哪个不是OpenStack的特点？", "options": ["A. 开源免费", "B. 模块化设计", "C. 只支持Windows", "D. 可扩展性强"], "answer": "C", "explain": "OpenStack主要运行在Linux上，不支持只运行在Windows。"},
            {"type": "判断", "q": "OpenStack是开源的私有云管理平台。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "OpenStack是最流行的开源私有云平台。"},
            {"type": "判断", "q": "OpenStack采用模块化架构，由多个组件协作。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "OpenStack由Nova、Neutron、Cinder等多个组件组成。"},
            {"type": "判断", "q": "OpenStack只能用于私有云，不能用于公有云。", "options": ["A. 正确", "B. 错误"], "answer": "B", "explain": "OpenStack也可用于构建公有云，但主要用于私有云。"},
            {"type": "判断", "q": "OpenStack通常运行在Linux操作系统上。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "OpenStack主要部署在Linux(如Ubuntu/CentOS)上。"},
            {"type": "填空", "q": "____是最流行的开源私有云平台（填一个英文单词）。", "answer": "OpenStack", "explain": "OpenStack。"},
            {"type": "填空", "q": "OpenStack采用____化架构（填两个字）。", "answer": "模块", "explain": "模块化设计，各组件独立部署。"},
        ]
    },
    {
        "num": 63, "chapter": "第13章 私有云平台", "title": "OpenStack核心组件",
        "questions": [
            {"type": "选择", "q": "OpenStack中Nova组件负责什么？", "options": ["A. 网络", "B. 计算(虚拟机管理)", "C. 存储", "D. 身份认证"], "answer": "B", "explain": "Nova是OpenStack的计算组件，负责虚拟机的创建、调度和管理。"},
            {"type": "选择", "q": "OpenStack中Neutron组件负责什么？", "options": ["A. 计算", "B. 网络(虚拟网络管理)", "C. 存储", "D. 镜像"], "answer": "B", "explain": "Neutron是网络组件，提供虚拟网络、路由、防火墙等网络服务。"},
            {"type": "选择", "q": "OpenStack中Cinder组件负责什么？", "options": ["A. 计算", "B. 网络", "C. 块存储", "D. 身份认证"], "answer": "C", "explain": "Cinder是块存储组件，为虚拟机提供持久化存储卷。"},
            {"type": "选择", "q": "OpenStack中Glance组件负责什么？", "options": ["A. 计算", "B. 镜像(虚拟机镜像管理)", "C. 网络", "D. 存储"], "answer": "B", "explain": "Glance是镜像服务组件，存储和管理虚拟机镜像。"},
            {"type": "选择", "q": "OpenStack中Keystone组件负责什么？", "options": ["A. 计算", "B. 网络", "C. 身份认证和授权", "D. 存储"], "answer": "C", "explain": "Keystone是身份认证组件，管理用户、租户和权限。"},
            {"type": "选择", "q": "OpenStack中Swift组件负责什么？", "options": ["A. 计算", "B. 对象存储", "C. 网络", "D. 身份认证"], "answer": "B", "explain": "Swift是对象存储组件，类似AWS S3，存储非结构化数据。"},
            {"type": "判断", "q": "Nova是OpenStack的计算组件。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "Nova负责虚拟机生命周期管理。"},
            {"type": "判断", "q": "Keystone负责OpenStack的身份认证。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "Keystone是认证授权中心。"},
            {"type": "判断", "q": "OpenStack各组件之间通过API相互协作。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "各组件通过RESTful API通信。"},
            {"type": "判断", "q": "Cinder提供的是对象存储服务。", "options": ["A. 正确", "B. 错误"], "answer": "B", "explain": "Cinder是块存储，Swift才是对象存储。"},
            {"type": "填空", "q": "OpenStack中____组件负责计算（填四个字母）。", "answer": "Nova", "explain": "Nova = 计算服务。"},
            {"type": "填空", "q": "OpenStack中____组件负责身份认证（填八个字母）。", "answer": "Keystone", "explain": "Keystone = 认证服务。"},
        ]
    },
    {
        "num": 64, "chapter": "第13章 私有云平台", "title": "私有云部署与应用 + 第十三章复习",
        "questions": [
            {"type": "选择", "q": "私有云部署的一般步骤是？", "options": ["A. 规划->准备环境->安装部署->测试验证", "B. 直接安装", "C. 先买硬件再说", "D. 不需要规划"], "answer": "A", "explain": "标准流程：需求规划->环境准备->安装部署->测试验证->运维。"},
            {"type": "选择", "q": "复习：OpenStack中Neutron负责什么？", "options": ["A. 计算", "B. 网络", "C. 存储", "D. 认证"], "answer": "B", "explain": "Neutron是网络组件。"},
            {"type": "选择", "q": "复习：私有云的主要优势是？", "options": ["A. 成本低", "B. 数据完全可控", "C. 弹性好", "D. 免运维"], "answer": "B", "explain": "数据可控是私有云的核心优势。"},
            {"type": "选择", "q": "私有云适合部署在什么环境？", "options": ["A. 只能云端", "B. 企业自有机房或第三方数据中心", "C. 只能在家", "D. 不需要场地"], "answer": "B", "explain": "私有云可自建也可托管。"},
            {"type": "选择", "q": "复习：OpenStack通常使用什么虚拟化技术？", "options": ["A. VirtualBox", "B. KVM", "C. Docker", "D. Hyper-V"], "answer": "B", "explain": "KVM是OpenStack最常用的虚拟化驱动。"},
            {"type": "选择", "q": "私有云在各行业的典型应用不包括？", "options": ["A. 银行核心系统", "B. 政务云", "C. 个人博客", "D. 军队信息系统"], "answer": "C", "explain": "个人博客适合公有云，不需要私有云。"},
            {"type": "判断", "q": "复习：OpenStack是开源私有云平台。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "OpenStack是开源的。"},
            {"type": "判断", "q": "复习：私有云的建设成本通常高于公有云。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "需自购硬件和运维团队。"},
            {"type": "判断", "q": "复习：Nova是OpenStack的计算组件。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "Nova负责虚拟机管理。"},
            {"type": "判断", "q": "私有云部署前需要进行需求规划和架构设计。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "好的规划是私有云成功部署的基础。"},
            {"type": "填空", "q": "复习：OpenStack中____负责计算（填四个字母）。", "answer": "Nova", "explain": "Nova = 计算。"},
            {"type": "填空", "q": "复习：____云数据完全可控（填两个字）。", "answer": "私有", "explain": "私有云数据在组织内部。"},
        ]
    },
    {
        "num": 65, "chapter": "综合复习", "title": "综合复习一：云计算基础与安全（第1-6章）",
        "questions": [
            {"type": "选择", "q": "NIST定义的云计算3-5-4框架中3代表什么？", "options": ["A. 3个部署模型", "B. 3个服务模型", "C. 3个特征", "D. 3个厂商"], "answer": "B", "explain": "3=3个服务模型(SaaS/PaaS/IaaS)，5=5个特征，4=4个部署模型。"},
            {"type": "选择", "q": "以下哪个属于SaaS？", "options": ["A. 阿里云ECS", "B. 腾讯文档", "C. VMware", "D. CentOS"], "answer": "B", "explain": "腾讯文档是SaaS(软件即服务)。"},
            {"type": "选择", "q": "信息安全三要素CIA中I代表什么？", "options": ["A. 认证", "B. 完整性", "C. 可用性", "D. 授权"], "answer": "B", "explain": "I = Integrity(完整性)。"},
            {"type": "选择", "q": "对称加密和非对称加密的主要区别是？", "options": ["A. 速度不同", "B. 对称用同一密钥，非对称用公私钥对", "C. 算法不同", "D. 用途不同"], "answer": "B", "explain": "对称=同一密钥，非对称=公钥/私钥一对。"},
            {"type": "选择", "q": "HTTPS使用的默认端口是？", "options": ["A. 80", "B. 443", "C. 22", "D. 8080"], "answer": "B", "explain": "HTTPS默认端口443。"},
            {"type": "选择", "q": "全球云计算市场份额最大的厂商是？", "options": ["A. 阿里云", "B. AWS", "C. 腾讯云", "D. 华为云"], "answer": "B", "explain": "AWS全球市场份额第一。"},
            {"type": "判断", "q": "NIST定义的5个特征包括资源池化和快速弹性。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "5个特征：按需自助、广泛网络访问、资源池化、快速弹性、可计量服务。"},
            {"type": "判断", "q": "哈希算法是可逆的。", "options": ["A. 正确", "B. 错误"], "answer": "B", "explain": "哈希是单向不可逆的。"},
            {"type": "判断", "q": "混合云结合了公有云的弹性和私有云的安全性。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "混合云 = 公有云 + 私有云。"},
            {"type": "判断", "q": "数字签名使用私钥签名、公钥验签。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "私钥签名保证不可否认性。"},
            {"type": "填空", "q": "NIST 3-5-4中3代表____个服务模型。", "answer": "3", "explain": "SaaS/PaaS/IaaS。"},
            {"type": "填空", "q": "CIA中C代表____性（填两个字）。", "answer": "机密", "explain": "Confidentiality = 机密性。"},
        ]
    },
    {
        "num": 66, "chapter": "综合复习", "title": "综合复习二：网络与数据库（第7-8章）",
        "questions": [
            {"type": "选择", "q": "OSI模型共有几层？", "options": ["A. 4层", "B. 5层", "C. 7层", "D. 9层"], "answer": "C", "explain": "OSI七层模型。"},
            {"type": "选择", "q": "TCP三次握手的目的是？", "options": ["A. 加密数据", "B. 建立可靠连接", "C. 传输数据", "D. 关闭连接"], "answer": "B", "explain": "三次握手建立TCP连接。"},
            {"type": "选择", "q": "以下哪个是私有IP地址？", "options": ["A. 8.8.8.8", "B. 192.168.1.1", "C. 202.96.128.86", "D. 114.114.114.114"], "answer": "B", "explain": "192.168.x.x是私有地址。"},
            {"type": "选择", "q": "SQL中用于查询数据的关键字是？", "options": ["A. INSERT", "B. SELECT", "C. UPDATE", "D. DELETE"], "answer": "B", "explain": "SELECT用于查询。"},
            {"type": "选择", "q": "以下哪个是关系型数据库？", "options": ["A. MongoDB", "B. Redis", "C. MySQL", "D. Cassandra"], "answer": "C", "explain": "MySQL是关系型数据库。"},
            {"type": "选择", "q": "DNS的作用是？", "options": ["A. 传输文件", "B. 域名解析为IP地址", "C. 发送邮件", "D. 分配IP"], "answer": "B", "explain": "DNS做域名到IP的解析。"},
            {"type": "判断", "q": "TCP是面向连接的可靠传输协议。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "TCP面向连接、可靠。"},
            {"type": "判断", "q": "主键可以重复。", "options": ["A. 正确", "B. 错误"], "answer": "B", "explain": "主键必须唯一。"},
            {"type": "判断", "q": "交换机工作在数据链路层。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "交换机基于MAC地址转发。"},
            {"type": "判断", "q": "SQL注入可以通过参数化查询防范。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "参数化查询有效防止SQL注入。"},
            {"type": "填空", "q": "TCP建立连接需要____次握手。", "answer": "3", "explain": "三次握手。"},
            {"type": "填空", "q": "SQL中____用于查询数据（填六个字母）。", "answer": "SELECT", "explain": "SELECT查询语句。"},
        ]
    },
    {
        "num": 67, "chapter": "综合复习", "title": "综合复习三：虚拟化与Linux（第9-10章）",
        "questions": [
            {"type": "选择", "q": "Type 1 Hypervisor直接运行在什么上？", "options": ["A. 操作系统", "B. 物理硬件", "C. 虚拟机", "D. 浏览器"], "answer": "B", "explain": "Type 1(裸金属型)直接运行在硬件上。"},
            {"type": "选择", "q": "Docker容器与虚拟机的主要区别？", "options": ["A. 容器更重", "B. 容器共享宿主OS内核", "C. 没有区别", "D. 容器需要独立OS"], "answer": "B", "explain": "容器共享内核，轻量快速。"},
            {"type": "选择", "q": "Linux根目录用什么表示？", "options": ["A. C:\\", "B. /", "C. \\", "D. root"], "answer": "B", "explain": "Linux根目录是/。"},
            {"type": "选择", "q": "ls命令的作用是？", "options": ["A. 创建目录", "B. 列出目录内容", "C. 删除文件", "D. 切换目录"], "answer": "B", "explain": "ls = list。"},
            {"type": "选择", "q": "KVM是哪个操作系统内核的虚拟化模块？", "options": ["A. Windows", "B. macOS", "C. Linux", "D. Android"], "answer": "C", "explain": "KVM是Linux内核模块。"},
            {"type": "选择", "q": "chmod命令的作用是？", "options": ["A. 修改文件所有者", "B. 修改文件权限", "C. 创建用户", "D. 查看文件"], "answer": "B", "explain": "chmod = change mode。"},
            {"type": "判断", "q": "虚拟化是云计算的核心基础技术。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "虚拟化实现资源池化。"},
            {"type": "判断", "q": "Linux是开源操作系统。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "Linux遵循GPL开源协议。"},
            {"type": "判断", "q": "Docker容器比虚拟机启动更快。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "容器秒级启动。"},
            {"type": "判断", "q": "VMware ESXi是Type 1 Hypervisor。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "ESXi是裸金属型。"},
            {"type": "填空", "q": "Linux根目录是____。", "answer": "/", "explain": "根目录/。"},
            {"type": "填空", "q": "虚拟化实现了资源____化。", "answer": "池", "explain": "Resource Pooling。"},
        ]
    },
    {
        "num": 68, "chapter": "综合复习", "title": "综合复习四：Web服务与云平台（第11-13章）",
        "questions": [
            {"type": "选择", "q": "HTTP默认端口是？", "options": ["A. 22", "B. 80", "C. 443", "D. 8080"], "answer": "B", "explain": "HTTP默认端口80。"},
            {"type": "选择", "q": "Nginx的主要用途是？", "options": ["A. 数据库", "B. 反向代理和负载均衡", "C. 操作系统", "D. 杀毒"], "answer": "B", "explain": "Nginx常用于反向代理和负载均衡。"},
            {"type": "选择", "q": "阿里云ECS属于什么服务？", "options": ["A. SaaS", "B. PaaS", "C. IaaS", "D. DaaS"], "answer": "C", "explain": "ECS是IaaS。"},
            {"type": "选择", "q": "OpenStack中Nova负责什么？", "options": ["A. 网络", "B. 计算", "C. 存储", "D. 认证"], "answer": "B", "explain": "Nova是计算组件。"},
            {"type": "选择", "q": "CDN的作用是？", "options": ["A. 加密数据", "B. 内容分发加速", "C. 存储数据", "D. 删除数据"], "answer": "B", "explain": "CDN加速内容访问。"},
            {"type": "选择", "q": "私有云的主要优势是？", "options": ["A. 成本低", "B. 数据完全可控", "C. 弹性好", "D. 免运维"], "answer": "B", "explain": "私有云数据在组织内部完全可控。"},
            {"type": "判断", "q": "LAMP架构包括Linux、Apache、MySQL、PHP。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "LAMP是经典Web架构。"},
            {"type": "判断", "q": "AWS是全球最大的公有云服务商。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "AWS全球第一。"},
            {"type": "判断", "q": "OpenStack是开源私有云平台。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "OpenStack开源。"},
            {"type": "判断", "q": "HTTP状态码404表示资源未找到。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "404 Not Found。"},
            {"type": "填空", "q": "HTTP状态码____表示资源未找到。", "answer": "404", "explain": "404 Not Found。"},
            {"type": "填空", "q": "____是全球最大的公有云厂商（填三个字母）。", "answer": "AWS", "explain": "Amazon Web Services。"},
        ]
    },
    {
        "num": 69, "chapter": "综合复习", "title": "模拟考试（上）",
        "questions": [
            {"type": "选择", "q": "云计算的本质是什么？", "options": ["A. 存储", "B. 弹性", "C. 网络", "D. 虚拟化"], "answer": "B", "explain": "云计算的本质是弹性——资源能大能小。"},
            {"type": "选择", "q": "以下哪个不属于NIST定义的5个特征？", "options": ["A. 按需自助服务", "B. 资源池化", "C. 固定硬件配置", "D. 快速弹性"], "answer": "C", "explain": "\"固定硬件配置\"不属于5个特征。"},
            {"type": "选择", "q": "银行核心系统适合使用什么云？", "options": ["A. 公有云", "B. 私有云", "C. 社区云", "D. 不用云"], "answer": "B", "explain": "银行安全要求高，适合私有云。"},
            {"type": "选择", "q": "AES属于哪种加密算法？", "options": ["A. 对称加密", "B. 非对称加密", "C. 哈希算法", "D. 签名算法"], "answer": "A", "explain": "AES是对称加密算法。"},
            {"type": "选择", "q": "TCP/IP模型有几层？", "options": ["A. 4层", "B. 5层", "C. 7层", "D. 3层"], "answer": "A", "explain": "TCP/IP四层模型。"},
            {"type": "选择", "q": "Docker容器的特点是？", "options": ["A. 每个容器有独立OS", "B. 共享宿主OS内核，轻量快速", "C. 需要Hypervisor", "D. 启动慢"], "answer": "B", "explain": "容器共享内核，轻量快速。"},
            {"type": "判断", "q": "SaaS用户不需要关心软件的维护和升级。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "SaaS由服务商负责维护升级。"},
            {"type": "判断", "q": "非对称加密使用公钥加密、私钥解密。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "公钥加密，私钥解密。"},
            {"type": "判断", "q": "Linux中一切皆文件。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "Linux核心设计理念。"},
            {"type": "判断", "q": "CDN可以加速网站访问。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "CDN将内容缓存到全球节点。"},
            {"type": "填空", "q": "云计算的本质是____。", "answer": "弹性", "explain": "弹性是核心本质。"},
            {"type": "填空", "q": "NIST 3-5-4中5代表____个基本特征。", "answer": "5", "explain": "5个特征。"},
        ]
    },
    {
        "num": 70, "chapter": "综合复习", "title": "模拟考试（下）与讲评",
        "questions": [
            {"type": "选择", "q": "云安全责任共担模型中用户负责什么？", "options": ["A. 物理服务器维护", "B. 数据安全和应用安全", "C. 网络线路", "D. 机房供电"], "answer": "B", "explain": "用户负责数据和应用层面的安全。"},
            {"type": "选择", "q": "以下哪个是Type 1 Hypervisor？", "options": ["A. VMware Workstation", "B. VirtualBox", "C. VMware ESXi", "D. Docker"], "answer": "C", "explain": "ESXi是Type 1裸金属型。"},
            {"type": "选择", "q": "SQL中删除数据的语句是？", "options": ["A. SELECT", "B. INSERT", "C. DELETE", "D. CREATE"], "answer": "C", "explain": "DELETE删除记录。"},
            {"type": "选择", "q": "HTTP状态码500表示什么？", "options": ["A. 请求成功", "B. 资源未找到", "C. 服务器内部错误", "D. 重定向"], "answer": "C", "explain": "500 = Internal Server Error。"},
            {"type": "选择", "q": "OpenStack中Keystone负责什么？", "options": ["A. 计算", "B. 网络", "C. 身份认证", "D. 存储"], "answer": "C", "explain": "Keystone是认证组件。"},
            {"type": "选择", "q": "以下哪个不是公有云的优势？", "options": ["A. 成本低", "B. 弹性好", "C. 数据完全可控", "D. 无需运维硬件"], "answer": "C", "explain": "数据完全可控是私有云的优势。"},
            {"type": "判断", "q": "对称加密的加解密速度比非对称加密快。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "对称加密计算量小速度快。"},
            {"type": "判断", "q": "反向代理可以隐藏后端服务器IP。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "反向代理隐藏后端真实服务器。"},
            {"type": "判断", "q": "数字证书由CA机构签发。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "CA(Certificate Authority)签发证书。"},
            {"type": "判断", "q": "UDP比TCP速度快但不可靠。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "UDP无连接不保证可靠但速度快。"},
            {"type": "填空", "q": "OpenStack中____负责身份认证。", "answer": "Keystone", "explain": "Keystone = 认证服务。"},
            {"type": "填空", "q": "HTTP状态码____表示服务器内部错误。", "answer": "500", "explain": "500 Internal Server Error。"},
        ]
    },
    {
        "num": 71, "chapter": "综合复习", "title": "期末考试",
        "questions": [
            {"type": "选择", "q": "NIST定义的3个服务模型是？", "options": ["A. SaaS/PaaS/IaaS", "B. LAN/MAN/WAN", "C. TCP/UDP/IP", "D. LAMP/LNMP/WAMP"], "answer": "A", "explain": "SaaS、PaaS、IaaS是三种服务模型。"},
            {"type": "选择", "q": "以下哪个不属于4种部署模型？", "options": ["A. 公有云", "B. 私有云", "C. 社区云", "D. 社交云"], "answer": "D", "explain": "4种部署模型：公有/私有/社区/混合云，没有社交云。"},
            {"type": "选择", "q": "防火墙的主要功能是？", "options": ["A. 查杀病毒", "B. 控制网络流量进出", "C. 加密数据", "D. 备份数据"], "answer": "B", "explain": "防火墙控制进出网络的流量。"},
            {"type": "选择", "q": "IP地址192.168.1.1属于哪类地址？", "options": ["A. A类", "B. B类", "C. C类", "D. D类"], "answer": "C", "explain": "192开头属于C类地址。"},
            {"type": "选择", "q": "以下哪个是Linux命令？", "options": ["A. dir", "B. ls", "C. copy", "D. del"], "answer": "B", "explain": "ls是Linux命令，dir/copy/del是Windows命令。"},
            {"type": "选择", "q": "负载均衡的作用是？", "options": ["A. 加密数据", "B. 将请求分发到多台服务器", "C. 压缩数据", "D. 存储数据"], "answer": "B", "explain": "负载均衡将流量分配到多台服务器。"},
            {"type": "判断", "q": "虚拟化技术是云计算的核心基础技术之一。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "虚拟化实现资源池化。"},
            {"type": "判断", "q": "HTTPS = HTTP + SSL/TLS。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "HTTPS在HTTP上增加SSL/TLS加密。"},
            {"type": "判断", "q": "OpenStack是开源的公有云平台。", "options": ["A. 正确", "B. 错误"], "answer": "B", "explain": "OpenStack主要用于私有云，也可用于公有云，但不是\"公有云平台\"。"},
            {"type": "判断", "q": "SSH默认端口是22。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "SSH默认端口22。"},
            {"type": "填空", "q": "NIST定义的3个服务模型是SaaS、PaaS和____。", "answer": "IaaS", "explain": "IaaS = 基础设施即服务。"},
            {"type": "填空", "q": "4种部署模型：公有云、私有云、社区云和____云。", "answer": "混合", "explain": "混合云 = 公有云 + 私有云。"},
        ]
    },
    {
        "num": 72, "chapter": "综合复习", "title": "试卷讲评与课程总结",
        "questions": [
            {"type": "选择", "q": "云计算课程总共覆盖了几章内容？", "options": ["A. 10章", "B. 13章", "C. 15章", "D. 8章"], "answer": "B", "explain": "教材共13章，从云计算简介到私有云平台。"},
            {"type": "选择", "q": "以下哪个是本课程最重要的核心概念？", "options": ["A. NIST 3-5-4框架", "B. 具体命令记忆", "C. 厂商产品价格", "D. 历史年份"], "answer": "A", "explain": "NIST 3-5-4框架(3服务模型+5特征+4部署模型)是云计算的核心知识。"},
            {"type": "选择", "q": "以下哪个技术贯穿了整个云计算课程？", "options": ["A. 虚拟化", "B. 杀毒", "C. 打印", "D. 蓝牙"], "answer": "A", "explain": "虚拟化是云计算的基础，在多个章节中反复出现。"},
            {"type": "选择", "q": "以下哪个不是学习云计算需要掌握的基础知识？", "options": ["A. 网络", "B. 操作系统(Linux)", "C. 数据库", "D. 汽车驾驶"], "answer": "D", "explain": "汽车驾驶与云计算无关。"},
            {"type": "选择", "q": "以下哪个是云计算的发展方向？", "options": ["A. 淘汰", "B. AI+云融合、边缘计算", "C. 回归本地部署", "D. 不再发展"], "answer": "B", "explain": "AI+云融合、边缘计算、混合云是云计算的发展趋势。"},
            {"type": "选择", "q": "学习云计算后可以从事什么工作？", "options": ["A. 云运维工程师", "B. 云架构师", "C. 云安全工程师", "D. 以上都可以"], "answer": "D", "explain": "云计算领域有运维、架构、安全等多种岗位。"},
            {"type": "判断", "q": "云计算是一个快速发展的技术领域。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "云计算技术持续演进，需要持续学习。"},
            {"type": "判断", "q": "虚拟化是云计算的基础技术之一。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "虚拟化实现资源池化，是IaaS的核心。"},
            {"type": "判断", "q": "Linux是云服务器的主流操作系统。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "绝大多数云服务器运行Linux。"},
            {"type": "判断", "q": "掌握云计算需要持续学习和实践。", "options": ["A. 正确", "B. 错误"], "answer": "A", "explain": "技术更新快，需持续学习。"},
            {"type": "填空", "q": "本课程共覆盖____章内容（填数字）。", "answer": "13", "explain": "第1-13章。"},
            {"type": "填空", "q": "____化是云计算的核心基础技术。", "answer": "虚拟", "explain": "虚拟化技术。"},
        ]
    },
]

# ========== 单元划分 ==========
UNITS = [
    ("第二单元 云安全与云市场（课时13-24）", 13, 24),
    ("第三单元 云网络与云数据库（课时25-36）", 25, 36),
    ("第四单元 虚拟化与Linux（课时37-48）", 37, 48),
    ("第五单元 Web服务与公有云（课时49-60）", 49, 60),
    ("第六单元 私有云与综合复习（课时61-72）", 61, 72),
]

FONT = "微软雅黑"

def add_tb(slide, l, t, w, h):
    tb = slide.shapes.add_textbox(Inches(l), Inches(t), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = True
    return tf

def set_para(p, text, size=14, bold=False, color=C_BLACK, align=PP_ALIGN.LEFT, space_after=4):
    p.alignment = align
    p.space_after = Pt(space_after)
    r = p.add_run()
    r.text = text
    r.font.size = Pt(size)
    r.font.bold = bold
    r.font.color.rgb = color
    r.font.name = FONT

def add_rect(slide, l, t, w, h, fill):
    from pptx.enum.shapes import MSO_SHAPE
    sh = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(l), Inches(t), Inches(w), Inches(h))
    sh.fill.solid()
    sh.fill.fore_color.rgb = fill
    sh.line.fill.background()
    return sh

def lesson_by_num(n):
    for les in lessons:
        if les["num"] == n:
            return les
    return None

# ========== 生成PPT ==========
prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
blank = prs.slide_layouts[6]

# --- 封面 ---
s = prs.slides.add_slide(blank)
add_rect(s, 0, 0, 13.333, 7.5, C_DARK_BLUE)
tf = add_tb(s, 1.5, 2.2, 10.3, 1.2)
set_para(tf.paragraphs[0], "云计算基础 第13-72课时 课堂习题", 40, True, C_WHITE, PP_ALIGN.CENTER)
tf = add_tb(s, 1.5, 3.8, 10.3, 0.8)
set_para(tf.paragraphs[0], "共60课时 · 每课时12题（6选择+4判断+2填空）· 合计720题", 20, False, C_WHITE, PP_ALIGN.CENTER)
tf = add_tb(s, 1.5, 4.7, 10.3, 0.6)
set_para(tf.paragraphs[0], "教材：《云计算基础技术与应用（第2版）（微课版）》", 14, False, RGBColor(0xBB, 0xDE, 0xFB), PP_ALIGN.CENTER)

# --- 目录（每单元1页）---
for unit_name, lo, hi in UNITS:
    s = prs.slides.add_slide(blank)
    add_rect(s, 0, 0, 13.333, 1.0, C_BLUE)
    tf = add_tb(s, 0.6, 0.18, 12.1, 0.7)
    set_para(tf.paragraphs[0], "目录 · " + unit_name, 24, True, C_WHITE)
    tf = add_tb(s, 0.6, 1.35, 6.0, 5.8)
    tf2 = add_tb(s, 6.9, 1.35, 6.0, 5.8)
    for i, n in enumerate(range(lo, hi + 1)):
        les = lesson_by_num(n)
        target = tf if i < 6 else tf2
        p = target.add_paragraph() if i >= 6 or i > 0 else target.paragraphs[0]
        set_para(p, "第%d课时  %s" % (n, les["title"]), 16, False, C_BLACK, space_after=14)

# --- 每课时：封面1页 + 习题2页 ---
for les in lessons:
    n = les["num"]
    # 课时封面
    s = prs.slides.add_slide(blank)
    add_rect(s, 0, 0, 13.333, 7.5, C_BLUE)
    tf = add_tb(s, 1.0, 2.0, 11.3, 1.0)
    set_para(tf.paragraphs[0], "第%d课时" % n, 44, True, C_WHITE, PP_ALIGN.CENTER)
    tf = add_tb(s, 1.0, 3.2, 11.3, 0.9)
    set_para(tf.paragraphs[0], les["title"], 30, True, C_WHITE, PP_ALIGN.CENTER)
    tf = add_tb(s, 1.0, 4.3, 11.3, 0.7)
    set_para(tf.paragraphs[0], les["chapter"], 18, False, RGBColor(0xBB, 0xDE, 0xFB), PP_ALIGN.CENTER)
    tf = add_tb(s, 1.0, 5.3, 11.3, 0.6)
    set_para(tf.paragraphs[0], "课堂习题：12题（6选择 + 4判断 + 2填空）", 16, False, C_WHITE, PP_ALIGN.CENTER)

    # 习题页 ×2
    for page_i, qset in enumerate((les["questions"][:6], les["questions"][6:])):
        s = prs.slides.add_slide(blank)
        add_rect(s, 0, 0, 13.333, 0.85, C_BLUE)
        tf = add_tb(s, 0.5, 0.12, 12.3, 0.6)
        set_para(tf.paragraphs[0], "第%d课时 %s · 课堂习题（%d-%d题）" % (n, les["title"], page_i * 6 + 1, page_i * 6 + 6), 20, True, C_WHITE)
        add_rect(s, 0, 7.28, 13.333, 0.22, C_ORANGE)
        tf = add_tb(s, 0.6, 1.05, 12.1, 6.1)
        first = True
        for qi, q in enumerate(qset):
            gno = page_i * 6 + qi + 1
            p = tf.paragraphs[0] if first else tf.add_paragraph()
            first = False
            set_para(p, "%d. [%s] %s" % (gno, q["type"], q["q"]), 15, True, C_DARK_BLUE, space_after=2)
            if "options" in q:
                p2 = tf.add_paragraph()
                set_para(p2, "        ".join(q["options"]), 13, False, C_GRAY, space_after=10)
            else:
                p2 = tf.add_paragraph()
                set_para(p2, "        答案：____________", 13, False, C_GRAY, space_after=10)

out_ppt = r"D:\WorkBuddyData\2026-08-11-07-48-22\云计算基础_第13-72课时课堂习题.pptx"
prs.save(out_ppt)
print("PPT saved:", out_ppt, "| slides:", len(prs.slides.__iter__.__self__._sldIdLst))

# ========== 生成HTML答案与解析 ==========
html = []
html.append("""<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<title>云计算基础 第13-72课时 课堂习题 答案与解析</title>
<style>
body { font-family: "Microsoft YaHei", sans-serif; background:#f5f6fa; margin:0; padding:20px; color:#333; }
.container { max-width: 1100px; margin: 0 auto; }
h1 { text-align:center; color:#0d47a1; }
.sub { text-align:center; color:#666; margin-bottom:30px; }
h2 { color:#0d47a1; border-left:6px solid #1a73e8; padding-left:12px; background:#e3f2fd; padding:10px 14px; border-radius:4px; }
h3 { color:#1a73e8; margin-bottom:6px; }
table { border-collapse: collapse; width:100%; background:#fff; margin:10px 0 24px 0; }
th, td { border:1px solid #d0d0d0; padding:6px 8px; text-align:center; font-size:14px; }
th { background:#1a73e8; color:#fff; }
td.qno { background:#f0f0f0; font-weight:bold; width:70px; }
.lesson { background:#fff; border-radius:8px; padding:18px 22px; margin:16px 0; box-shadow:0 1px 4px rgba(0,0,0,0.08); }
.q { margin:10px 0; padding:10px 12px; background:#fafafa; border-radius:6px; border-left:4px solid #1a73e8; }
.q .stem { font-weight:bold; }
.q .ans { color:#c62828; font-weight:bold; margin-top:4px; }
.q .exp { color:#555; margin-top:4px; font-size:14px; }
.tag { display:inline-block; background:#e3f2fd; color:#0d47a1; border-radius:4px; padding:1px 8px; font-size:12px; margin-right:6px; }
@media print { .lesson { box-shadow:none; border:1px solid #ccc; page-break-inside:avoid; } }
</style>
</head>
<body>
<div class="container">
<h1>云计算基础 第13-72课时 课堂习题 答案与解析</h1>
<div class="sub">共60课时 · 720题 · 教材：《云计算基础技术与应用（第2版）（微课版）》</div>
<h2>答案速查表</h2>
""")

for unit_name, lo, hi in UNITS:
    html.append("<h3>%s</h3>" % unit_name)
    html.append("<table><tr><th>课时</th>" + "".join("<th>%d</th>" % i for i in range(1, 13)) + "</tr>")
    for n in range(lo, hi + 1):
        les = lesson_by_num(n)
        html.append("<tr><td class='qno'>第%d课时</td>" % n)
        for q in les["questions"]:
            html.append("<td>%s</td>" % q["answer"])
        html.append("</tr>")
    html.append("</table>")

html.append("<h2>逐题答案与解析</h2>")
for unit_name, lo, hi in UNITS:
    html.append("<h3>%s</h3>" % unit_name)
    for n in range(lo, hi + 1):
        les = lesson_by_num(n)
        html.append("<div class='lesson'>")
        html.append("<h3>第%d课时 · %s <small style='color:#888;font-weight:normal'>(%s)</small></h3>" % (n, les["title"], les["chapter"]))
        for qi, q in enumerate(les["questions"]):
            html.append("<div class='q'>")
            html.append("<div class='stem'><span class='tag'>%d. %s</span>%s</div>" % (qi + 1, q["type"], q["q"]))
            if "options" in q:
                html.append("<div style='color:#666;margin-top:2px;'>%s</div>" % "　".join(q["options"]))
            html.append("<div class='ans'>答案：%s</div>" % q["answer"])
            html.append("<div class='exp'>解析：%s</div>" % q["explain"])
            html.append("</div>")
        html.append("</div>")

html.append("</div></body></html>")

out_html = r"D:\WorkBuddyData\2026-08-11-07-48-22\云计算基础_第13-72课时课堂习题_答案与解析.html"
with open(out_html, "w", encoding="utf-8") as f:
    f.write("\n".join(html))
print("HTML saved:", out_html)
print("Done. Lessons:", len(lessons))
