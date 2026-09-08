# CMC_scheduler
[Python/PyQt5] 엑셀 VBA 매크로 및 스크립트 멀티 예약 실행 프로그램 (백그라운드/트레이/자동시작 기능 포함)

Windows에서 Python 스크립트(`.py`), 실행 파일(`.exe`), Excel VBA 매크로(`.xlsm`, `.xlsb`)를 지정한 시간에 자동 실행합니다.

## 실행 방법

```powershell
python main.py
```

가상환경을 쓰는 경우:

```powershell
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe main.py
```

실행 파일은 `dist\CMC스케줄러\CMC스케줄러.exe`입니다. 폴더 전체를 복사해서 사용하세요.

## 주요 기능

- 한 번만 / 매일 / 매주 / 매월 반복 예약
- 창을 닫아도 트레이에서 백그라운드 동작 (옵션)
- PC 부팅 시 자동 시작
- Excel 파일 선택 시 VBA 매크로 목록 자동 표시
- 작업 즉시 실행, 일시중지/재개, 실행 로그 저장

## 문서

- [사용설명서.txt](사용설명서.txt)
- [CMC_Scheduler_User_Manual.txt](CMC_Scheduler_User_Manual.txt)

## 다시 빌드

```powershell
.\build_exe.ps1
```

## 라이선스

MIT License
