"""ToMD: 문서 파일을 마크다운으로 변환한다."""
import configparser
import os
import sys

__version__ = "14"


def version_label():
    """설치본이면 설치 파일의 버전(GitHub Releases 번호와 같음)을, 아니면 소스 버전을 돌려준다."""
    try:
        ini = os.path.join(os.path.dirname(os.path.abspath(sys.executable)), "version.ini")
        if os.path.isfile(ini):
            parser = configparser.ConfigParser()
            parser.read(ini, encoding="utf-8")
            value = parser.get("tomd", "version", fallback="").strip()
            if value:
                return value
    except Exception:
        pass
    return f"v{__version__}"
