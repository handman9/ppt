# 예배 자막 PPTX 생성기

가사를 한 줄씩 입력하면 업로드한 샘플 PPTX의 자막 디자인을 복제해
한 줄당 한 장의 슬라이드를 생성하는 로컬 웹 도구입니다.

## 준비

- Windows
- Python 3.10 이상

Python이 없다면 먼저 설치합니다.

https://www.python.org/downloads/

설치할 때 `Add python.exe to PATH`를 체크하면 `run_server.bat`로 바로 실행할 수 있습니다.
첫 실행 때 필요한 Python 패키지가 자동 설치됩니다.

## 실행

`run_server.bat`를 더블클릭합니다.

또는 PowerShell에서 실행합니다.

```powershell
cd D:\util\lyric_caption_web
python app.py
```

브라우저에서 `http://127.0.0.1:8765`로 접속합니다.

## Render 배포

Render에서 Web Service를 만들고 GitHub 저장소를 연결합니다.

```text
Build Command: pip install -r requirements.txt
Start Command: python app.py
```

저장소의 `.python-version` 파일로 Python 3.12를 사용하게 지정되어 있습니다.

## 사용

1. 자막 스타일로 쓸 PPTX 파일을 업로드합니다.
2. 가사를 한 줄씩 붙여 넣습니다.
3. PPTX를 생성합니다.

템플릿 PPTX의 첫 번째 텍스트 포함 슬라이드를 기준으로 삼고,
그 슬라이드 안에서 가장 큰 텍스트 박스를 자막 박스로 사용합니다.
배경, 이미지, 텍스트 위치, 폰트 스타일은 템플릿을 그대로 유지합니다.

## 결과 파일

생성된 PPTX는 브라우저로 다운로드되고, 서버 폴더의 `outputs` 폴더에도 저장됩니다.

## 주의

- 샘플 PPTX는 `.pptx` 형식이어야 합니다.
- 템플릿 안에 자막으로 쓸 텍스트 박스가 하나 이상 있어야 합니다.
- PowerPoint에서 복구 메시지가 뜨지 않도록 슬라이드 관계 ID 충돌을 피해서 생성합니다.
