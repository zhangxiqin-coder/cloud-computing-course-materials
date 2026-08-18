#!/bin/bash
# B站视频批量下载脚本
# 使用 yt-dlp + ffmpeg 合并音视频

FFMPEG_DIR="C:\\Users\\Administrator\\.workbuddy\\binaries\\python\\envs\\default\\Lib\\site-packages\\imageio_ffmpeg\\binaries"
YT_DLP="C:\\Users\\Administrator\\.workbuddy\\binaries\\python\\envs\\default\\Scripts\\yt-dlp.exe"
OUT_DIR="D:/WorkBuddyData/2026-08-11-07-48-22/bilibili_videos"

# 视频列表：课时|BV号|备注
VIDEOS=(
  "01|BV1RF411C7qu|什么是云计算"
  "01|BV1yh41127T2|云计算到底是什么"
  "02|BV1bt411u72Z|云计算的发展与应用-鲜枣课堂"
  "02|BV1Uc411h7oX|云计算演变之路"
  "02|BV1cE411W7Ls|杨哥带你走进云计算的世界"
  "04|BV1hf4y1W7yT|用人话说IaaS-PaaS-SaaS"
  "06|BV1Mh4y1Y73m|使用gitee上传代码"
  "07|BV1kA411N7e5|阿里云超级数据中心"
  "07|BV1cW411t7dk|走进微软数据中心"
  "08|BV1F14y1677t|淄博燃气智慧云数据中心"
  "08|BV14u4y1C7TG|Google-Cloud-TPU-Data-Center"
  "09|av928620141|阿里云服务器ECS新手搭建网站"
  "09|BV18a4y1W7e9|通俗易懂的网站上线部署发布教程"
  "12|BV1N7wveKEmx|VMware虚拟机安装教程"
  "12|BV14a411w7D2|VMware虚拟化入门到精通"
)

SUCCESS=0
FAIL=0
FAILED_LIST=""

for entry in "${VIDEOS[@]}"; do
  IFS='|' read -r LESSON BVID DESC <<< "$entry"
  URL="https://www.bilibili.com/video/${BVID}"
  echo ""
  echo "============================================"
  echo "下载: 课时${LESSON} | ${BVID} | ${DESC}"
  echo "============================================"

  "$YT_DLP" \
    --ffmpeg-location "$FFMPEG_DIR" \
    -f "bestvideo[height<=720]+bestaudio/best[height<=720]/best" \
    --merge-output-format mp4 \
    --no-playlist \
    --no-warnings \
    -o "${OUT_DIR}/课时${LESSON}_${DESC}.%(ext)s" \
    "$URL" 2>&1

  if [ $? -eq 0 ]; then
    echo "[OK] 课时${LESSON} ${DESC} 下载完成"
    ((SUCCESS++))
  else
    echo "[FAIL] 课时${LESSON} ${DESC} 下载失败"
    ((FAIL++))
    FAILED_LIST="${FAILED_LIST}课时${LESSON} ${BVID} ${DESC}\n"
  fi
done

echo ""
echo "============================================"
echo "下载完成！成功: ${SUCCESS} 个, 失败: ${FAIL} 个"
if [ -n "$FAILED_LIST" ]; then
  echo -e "失败列表:\n${FAILED_LIST}"
fi
echo "视频保存在: ${OUT_DIR}"
echo "============================================"
