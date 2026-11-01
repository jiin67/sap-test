# SAP ADT 설치 (Windows 재설치 후)

## 실행
PowerShell 을 **관리자 권한**으로 열고:

```powershell
Set-ExecutionPolicy -Scope Process Bypass -Force
cd C:\sap-test\setup
.\install_adt.ps1
```

로그: `C:\SAP_ADT\install_adt_yyyyMMdd_HHmm.log`

## 설치되는 것
| 순서 | 프로그램 | 방법 | 용도 |
|---|---|---|---|
| 1 | Eclipse Temurin JDK 21 | winget | Eclipse/ADT 실행용 Java |
| 2 | VC++ 2013 x64 런타임 | winget | ADT 의 SAP JCo 네이티브 라이브러리 필수 (없으면 시스템 연결 실패) |
| 3 | VC++ 2015-2022 x64/x86 | winget | SAP GUI 필수 |
| 4 | Git, Python 3.12 | winget (`-SkipExtras` 로 생략) | 이 폴더 복구·결과서 스크립트 |
| 5 | Eclipse IDE for Java Developers | zip → `C:\SAP_ADT\eclipse` | ADT 호스트 |
| 6 | ABAP Development Tools | p2 director (`https://tools.hana.ondemand.com/latest`) | ADT 플러그인 |
| 7 | SAP GUI for Windows 7.70/8.00 | **수동** (사내 배포 패키지) | 스크립트는 설치 여부만 확인 |

## 자동 설치가 안 될 때
- ADT 기능 ID 를 자동으로 못 찾으면: Eclipse 실행 → Help → Install New Software →
  `https://tools.hana.ondemand.com/latest` → **ABAP Development Tools** 체크 → 설치 후 재시작.
- 회사 프록시로 다운로드가 막히면 Eclipse 에서 Window → Preferences → Network Connections 설정.

## 설치 후
1. SAP GUI 설치 → SAP Logon 에 KDD/KDQ 등록 (KDP 는 ADT 에서도 조회 외 작업 금지).
2. Eclipse → Window → Perspective → ABAP → File → New → ABAP Project → SAP Logon 항목 선택 → 클라이언트·사용자로 로그온.
3. 이 저장소 복구: `git clone https://github.com/jiin67/sap-test C:\sap-test`
