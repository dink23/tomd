"""PyInstaller용 버전 정보 파일을 만든다. exe 속성(자세히)에 제품 이름·설명·버전이 보이게 한다.

사용: python installer/make_version.py 1.0.15  →  installer/version_info.txt
"""
import sys
from pathlib import Path

version = sys.argv[1] if len(sys.argv) > 1 else "1.0.0"
parts = [int(p) for p in (version.split(".") + ["0", "0", "0", "0"])[:4] if p.isdigit()]
parts = (parts + [0, 0, 0, 0])[:4]
numbers = ", ".join(str(p) for p in parts)

text = f"""# -*- coding: utf-8 -*-
VSVersionInfo(
  ffi=FixedFileInfo(filevers=({numbers}), prodvers=({numbers}), mask=0x3f, flags=0x0, OS=0x40004, fileType=0x1, subtype=0x0, date=(0, 0)),
  kids=[
    StringFileInfo([StringTable('041204b0', [
      StringStruct('CompanyName', 'ToMD'),
      StringStruct('FileDescription', 'ToMD - 문서를 마크다운으로 변환'),
      StringStruct('FileVersion', '{version}'),
      StringStruct('InternalName', 'tomd'),
      StringStruct('OriginalFilename', 'tomd.exe'),
      StringStruct('ProductName', 'ToMD'),
      StringStruct('ProductVersion', '{version}')])]),
    VarFileInfo([VarStruct('Translation', [0x0412, 1200])])
  ]
)
"""
out = Path(__file__).with_name("version_info.txt")
out.write_text(text, encoding="utf-8")
print(f"wrote {out} ({version})")
