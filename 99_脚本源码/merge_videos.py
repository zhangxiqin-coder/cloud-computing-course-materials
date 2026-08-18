# -*- coding: utf-8 -*-
"""合并B站视频分片 - 不删除源文件"""
import os
import re
import sys
import subprocess

sys.stdout.reconfigure(encoding='utf-8')

FFMPEG = r"C:\Users\Administrator\.workbuddy\binaries\python\envs\default\Lib\site-packages\imageio_ffmpeg\binaries\ffmpeg-win-x86_64-v7.1.exe"
DIR = r"D:\WorkBuddyData\2026-08-11-07-48-22\bilibili_videos"

files = os.listdir(DIR)
video_parts = sorted([f for f in files if re.search(r'\.f\d+\.mp4$', f)])

success = 0
skip = 0
fail = 0

for vname in video_parts:
    base = re.sub(r'\.f\d+\.mp4$', '', vname)
    vpath = os.path.join(DIR, vname)
    outpath = os.path.join(DIR, base + ".mp4")
    
    # 找音频
    audio_parts = [f for f in files if f.startswith(base + '.') and f.endswith('.m4a') and '.f' in f]
    if not audio_parts:
        print(f"[SKIP] {vname} - 无音频")
        skip += 1
        continue
    apath = os.path.join(DIR, audio_parts[0])
    
    # 已合并则跳过
    if os.path.exists(outpath) and os.path.getsize(outpath) > 1024*1024:
        print(f"[EXIST] {base}.mp4 ({os.path.getsize(outpath)//1024//1024}MB)")
        skip += 1
        continue
    
    print(f"合并: {base}.mp4 ...", end=" ", flush=True)
    cmd = [FFMPEG, "-i", vpath, "-i", apath, "-c:v", "copy", "-c:a", "aac", "-strict", "experimental", "-y", outpath]
    result = subprocess.run(cmd, capture_output=True)
    
    if result.returncode == 0 and os.path.exists(outpath) and os.path.getsize(outpath) > 1024*1024:
        print(f"OK ({os.path.getsize(outpath)//1024//1024}MB)")
        success += 1
    else:
        print("FAIL")
        fail += 1

print(f"\n{'='*50}")
print(f"成功: {success}, 跳过: {skip}, 失败: {fail}")
print(f"{'='*50}")

# 最终列表
print("\n最终视频文件：")
final = sorted([f for f in os.listdir(DIR) if f.endswith('.mp4') and '.f' not in f])
total = 0
for f in final:
    s = os.path.getsize(os.path.join(DIR, f))
    total += s
    print(f"  {f}  ({s//1024//1024}MB)")
print(f"\n共 {len(final)} 个视频，总大小 {total//1024//1024//1024}GB{total//1024//1024%1024}MB")
