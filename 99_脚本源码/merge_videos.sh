#!/bin/bash
# 合并已下载的视频(.mp4)和音频(.m4a)文件

FFMPEG="C:\\Users\\Administrator\\.workbuddy\\binaries\\python\\envs\\default\\Lib\\site-packages\\imageio_ffmpeg\\binaries\\ffmpeg-win-x86_64-v7.1.exe"
DIR="D:/WorkBuddyData/2026-08-11-07-48-22/bilibili_videos"

cd "$DIR" || exit 1

# 找出所有 .mp4 视频文件（排除已合并的）
for vfile in *.mp4; do
  # 跳过已合并的文件（不含 .f 后缀的）
  if [[ "$vfile" != *.f*.mp4 ]]; then
    continue
  fi

  # 提取基础文件名（去掉 .fXXXXX.mp4）
  base=$(echo "$vfile" | sed 's/\.f[0-9]*\.mp4$//')
  afile="${base}.f30280.m4a"

  # 如果对应的音频文件不存在，尝试其他格式
  if [ ! -f "$afile" ]; then
    afile="${base}.f30251.m4a"
  fi
  if [ ! -f "$afile" ]; then
    # 查找任何 .m4a 文件
    afile=$(ls "${base}".*.m4a 2>/dev/null | head -1)
  fi

  outfile="${base}.mp4"

  if [ -f "$afile" ]; then
    echo "合并: ${vfile} + ${afile} -> ${outfile}"
    "$FFMPEG" -i "$vfile" -i "$afile" -c:v copy -c:a aac -strict experimental -y "$outfile" 2>&1 | tail -3

    if [ $? -eq 0 ] && [ -f "$outfile" ]; then
      echo "[OK] $outfile"
      rm -f "$vfile" "$afile"
    else
      echo "[FAIL] $outfile"
    fi
  else
    echo "[SKIP] $vfile - 未找到对应音频文件"
  fi
  echo ""
done

echo "============================================"
echo "合并完成！最终文件列表："
echo "============================================"
ls -lh "$DIR"/*.mp4 2>/dev/null
