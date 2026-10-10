# add_process_sheet_v1.ps1 : "프로세스흐름" 시트 (SR03000035400 결과서 v1.00, 2026-10-02 현재 설정 S01~S10 기준)
#   fill_template_subs.ps1 다음에 실행. -After 로 어느 시트 뒤에 둘지 지정(기본: 6번째 FAIL내역 뒤).
param([Parameter(Mandatory = $true)][string]$Path, [int]$After = 6)
$ErrorActionPreference = "Stop"
$HDR = 2050429; $IN = 14804223; $OK = 14348258; $NG = 13551615

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false; $xl.ScreenUpdating = $false
$script:r = 1
try {
    $wb = $xl.Workbooks.Open($Path)
    foreach ($s in @($wb.Worksheets)) { if ($s.Name -eq "프로세스흐름") { $s.Delete() } }
    $ws = $wb.Worksheets.Add([Type]::Missing, $wb.Worksheets.Item($After))
    $ws.Name = "프로세스흐름"

    function Title($t) { $c = $ws.Cells.Item($script:r, 1); $c.Value2 = $t; $c.Font.Size = 14; $c.Font.Bold = $true; $script:r += 2 }
    function Section($t) { $c = $ws.Cells.Item($script:r, 1); $c.Value2 = $t; $c.Font.Size = 11; $c.Font.Bold = $true; $c.Font.Color = $HDR; $script:r++ }
    function Head($cells) { for ($i = 0; $i -lt $cells.Count; $i++) { $c = $ws.Cells.Item($script:r, $i + 1); $c.Value2 = $cells[$i]; $c.Interior.Color = $HDR; $c.Font.Color = 16777215; $c.Font.Bold = $true; $c.HorizontalAlignment = -4108; $c.Borders.LineStyle = 1 }; $script:r++ }
    function Row($cells, $fill = $null) { for ($i = 0; $i -lt $cells.Count; $i++) { $c = $ws.Cells.Item($script:r, $i + 1); $c.NumberFormat = "@"; $c.Value2 = $cells[$i]; $c.Borders.LineStyle = 1; $c.VerticalAlignment = -4160; $c.WrapText = $true; if ($fill) { $c.Interior.Color = $fill } }; $script:r++ }
    function Mono($t) { $c = $ws.Cells.Item($script:r, 1); $c.Value2 = $t; $c.Font.Name = "Consolas"; $c.Font.Size = 10; $script:r++ }
    function Note($t) { $c = $ws.Cells.Item($script:r, 1); $c.Value2 = $t; $c.Font.Italic = $true; $script:r += 2 }
    function Gap { $script:r++ }

    Title "SR03000035400 보일러 서비스 구독 대리점 비용 정산 - 전체 프로세스"
    $ws.Cells.Item(2, 1).Value2 = "보일러를 일시불로 산 고객이 서비스 구독으로 전환하면 AS 는 무상이 된다. 그 비용(센터 무상 부품, 대리점 회수품 지원금, 대리점 AS 무상지원금)을 일반 고객 건과 분리해 집계·정산·전기하는 것이 이번 개선이다. 아래는 2026-10-02 QA(KDQ 210) 설정 기준으로 실제 확인한 흐름이다."
    $ws.Cells.Item(2, 1).Font.Italic = $true
    $script:r = 4

    Section "1. 한 장 요약 - 상담 1건이 거치는 길"
    Mono "  [CIC 상담시스템]  고객 상담 → 조치 완료(기사 방문, 부품 교체)"
    Mono "        │  ① RFC ZCRMF02_RFC_ZCRM_CICSVC_ALL 가 SAP 으로 적재 (구독여부 SUBS_TYPE, 보증만료일 WRTDATE, 보증잔여개월 SUBSWRTMONTH 포함)"
    Mono "        ▼"
    Mono "  ZCRM_CICSVC (센터·회수품 계열)                  ZCRM_CICSVC_NEW (지원금 계열)"
    Mono "        │                                                │"
    Mono "   ┌────┴────┐                                           │"
    Mono "   ▼         ▼                                           ▼"
    Mono "  ② 센터매출  ③ 회수품접수                           ⑤ 무상지원금 산정"
    Mono "  ZCRM177A/AS  ZCRM052/052S                         ZCRM058A/058S"
    Mono "  오더→납품→출고  적합판정·접수(ZCRM_RECPT)           확정 → 대리점확정 → FI 전표"
    Mono "   │              │                                      │"
    Mono "   ▼              ▼                                      ▼"
    Mono "  출고 FI 전표   ④ 회수품정산 ZCRM171/171S          ZCRM_ASSPTFEEH1 (키 + SUBS_TYPE)"
    Mono "  (5328130)      본사확정 → FI 전표                        │"
    Mono "                 ZCRM_RECPTF (키 + SUBS_TYPE)              ▼"
    Mono "                                                     ⑥ 지원금 상세 ZCRM060A/060S"
    Gap
    Note "※ 대리점으로 가는 돈은 두 갈래다. ④ 회수품 지원금(부품 단가 기준, 구독 N년이후는 소비자가 71,000 / 그 외 대리점가 39,500 + VAT)과 ⑤ AS 무상지원금(건당 5,000원 등 지급기준별). 센터가 쓴 부품(②)은 돈이 아니라 재고 출고 원가로 잡힌다."

    Section "2. 단계별 프로그램과 구독 분기"
    Head @("단계", "프로그램 / T-code", "하는 일", "구독 분기 방식", "구독 전용 결과 (설정·테이블)")
    Row @("① CIC 적재", "기능그룹 ZCRMG02`nZCRMF02_RFC_ZCRM_CICSVC_ALL", "CIC 상담·조치 데이터를 ZCRM_CICSVC / ZCRM_CICSVC_NEW 에 넣는다", "Native SQL 에 SUBS_TYPE·WRTDATE·SUBSWRTMONTH 3개 필드 추가 (부호 보존)", "SUBS_TYPE='S' 가 구독의 유일한 기준. SUBSWRTMONTH>0 이면 'N년이후', 0 이하면 'N년이내'")
    Row @("② 센터 매출", "ZCRMR177_1`nZCRM177A(일반) / ZCRM177AS(구독)", "센터 기사가 쓴 무상 부품을 판매오더 → 납품 → 출고전기로 처리", "sy-tcode 로 gv_subs_mode 결정(70행) → 혼재 판정(ZF_CHK_DUP_CSLTNO) 뒤 상담번호 단위로 구독/일반 DELETE(152~156행)", "구독 화면 제목 '[구독] AS 비용처리', 구독여부 문구 '구독무상(n년이후/이내)', 구독 오더는 ZCRM_ORDER_AUTO 프로그램명 ZCRM177AS. 출고 원가 계정·코스트센터는 일반과 동일(5328130 / 출고 센터)")
    Row @("③ 회수품 접수", "ZCRMR052`nZCRM052(일반) / ZCRM052S(구독)`nZCRM053/053S(조회전용)", "대리점이 반납한 회수품의 적합/부적합을 판정하고 접수(ZCRM_RECPT)", "sy-tcode → SUBS_TYPE 레인지로 조회 분리. 구독 모드에서만 소비자가(ZPA2)·Z16 자동판정 적용", "구독 N년이후 단가 = 소비자가 ZPA2(71,000) / N년이내 = 대리점가 PR00(39,500). 회수가능제조일 > 부품제조일이면 자동 부적합 Z16(구독 기간 만료). 센터조치·서비스점조치 라디오 모두 표시")
    Row @("④ 회수품 정산", "ZCRMR171`nZCRM171(일반) / ZCRM171S(구독)", "기간월(적합판정일의 월)별 대리점 지원금을 집계 → 본사확정 → FI 전표", "sy-tcode → 정산유형 S08(구독) / Z08(일반). ZCRM_RECPTF 키에 SUBS_TYPE", "FM ZCRMF03_AS_SETTLE_FIPOSTING 이 ZCRM_FIPOST_T(S08: 계정 5328210 A/S부품비) + ZCRM_KOSTL_T(S08 × 센터: 코스트센터 = 대리점 센터, 예산센터 = 240074 구독사업팀)를 읽는다. 일반 Z08 은 예산센터 260002 CS기획팀. 요청일자 기본값 = 이번 달")
    Row @("⑤ 무상지원금", "ZCRMR058_1`nZCRM058A/059A(일반) / ZCRM058S/059S(구독)", "기간월(조치일의 월)별 대리점 AS 무상지원금 산정 → 확정 → 대리점확정 → FI 전표", "sy-tcode → 정산유형 S01(구독) / Z01(일반). ZCRM_ASSPTFEEH1 키에 SUBS_TYPE, 조회·DELETE·UPDATE 모두 subs_type 조건", "S01: 계정 5328170 A/S지원금, 코스트센터 = 대리점 센터, 예산센터 = 240074. 구독은 CIC 확정일(ZCRM_CHRG_CNFRM)이 있어야 하고 화면에 대리점확정 버튼이 없다(확정 → 전표). 예산계정 5328150 에 240074 예산 필요")
    Row @("⑥ 지원금 상세", "ZCRMR060_1`nZCRM060A(일반) / ZCRM060S(구독)", "무상지원금 상세 리스트 조회·전송", "sy-tcode → SUBS_TYPE 레인지(ZCRM_CICSVC_NEW)", "구독 모드에서만 기술료·출장비 컬럼 표시")
    Row @("공용", "기능그룹 ZCRMG03`nZCRMF03_GET_SUBS_TEXT", "구독 문구 생성(설치일·보증만료일 → 구독기간 n년 / n.n년)", "052·171·177 이 같은 FM 호출", "12의 배수면 'n년', 아니면 소수점 1자리(예 1.3년), 0 이하 'N년'. 171 정산 줄은 '구독무상' 고정")
    Note "※ 058/060 은 대리점의 KNVV 소속 센터 기준으로, 171 은 상담건의 센터(ZCRM_CICSVC-VKBUR) 기준으로 센터 조건을 거른다. ZCRMR194(무상청구)는 이번 개선 대상이 아니다."

    Section "3. 정산유형·계정·코스트센터 설정 (2026-10-02 KDQ 기준 - '현재 설정대로 진행' 가정)"
    Head @("정산유형", "쓰는 프로그램", "ZCRM_FIPOST_T 계정 (1100)", "ZCRM_KOSTL_T 코스트센터 / 예산센터", "예산 통제")
    Row @("S08 [구독]AS부품비", "ZCRM171S", "2103100(31) / 5328210 A/S부품비(40), 세금 A1", "0490·0520·CT01~CT24 25행: 코스트센터 = 센터별(예 CT18 260014), 예산센터 = 240074", "5328210 은 통제 계정 아님") $OK
    Row @("Z08 AS부품비", "ZCRM171", "2103100 / 5328210, A1", "센터별 코스트센터, 예산센터 = 260002 CS기획팀", "-") $OK
    Row @("S01 [구독]AS무상지원금", "ZCRM058S", "2103100 / 5328170 A/S지원금, A1", "27행: 코스트센터 = 센터별, 예산센터 = 240074", "예산계정 5328150 - 240074 에 예산 필요 (QA 는 ZCO012 40051 → ZCO013 으로 편성)") $IN
    Row @("Z01 AS무상지원금", "ZCRM058A", "2103100 / 5328170, A1", "센터별 코스트센터, 예산센터 = 260002", "260002 예산") $OK
    Row @("S09 [구독]AS정산", "(예비)", "-", "24행, 예산센터 240074", "-") $IN
    Row @("S02~S07·S10", "(예비 - Z02~Z07·Z10 의 구독 복사본)", "Z 와 동일", "행 없음", "-") $IN
    Note "※ 정산유형명은 ZCRM_STRART_T(SM30), 계정은 ZCRM_FIPOST_V, 코스트센터는 AS정산 코스트센터 유지보수뷰에서 관리. 운영 반영 전 체크리스트는 C:\sap-test\OPEN_ITEMS.md (O-3 조기확정 통제, O-4 예산센터·운영 예산, O-5 설정 이관 경로, O-6 무상출고 원가)."

    Section "4. 정산 시점·운영 규칙"
    Head @("항목", "내용")
    Row @("접수 마감", "월 마감 이후에는 그 기간월로 추가 접수를 받지 않는다. 추가 접수는 새 상담번호로 들어와 다음 기간월에 정산된다 (2026-09-21 업무 확인)")
    Row @("확정 시점", "④ 회수품 정산의 본사확정·전표생성, ⑤ 무상지원금 확정·전표는 기간월 마감 뒤에 수행한다. 프로그램에는 월 중 확정을 막는 장치가 없다(O-3)")
    Row @("되돌리기", "171: 노랑(확정)은 본사확정취소, 초록(전표)은 전표취소 → 확정취소. 058: 전표삭제 → 대리점확정취소 → 무상지원금취소 순. 구독(058S)은 CIC 확정일이 있는 건은 취소가 막힌다")
    Row @("구독·일반 분리 보장", "같은 대리점·같은 기간월이라도 ZCRM_RECPTF·ZCRM_ASSPTFEEH1 에 구독(S)/일반(공백) 행이 따로 남고, 한쪽 화면의 처리가 다른 쪽 행을 건드리지 않는다 (v0.07 D-08 재테스트로 확인)")
    Gap

    $ws.Columns.Item(1).ColumnWidth = 24
    $ws.Columns.Item(2).ColumnWidth = 36
    $ws.Columns.Item(3).ColumnWidth = 44
    $ws.Columns.Item(4).ColumnWidth = 52
    $ws.Columns.Item(5).ColumnWidth = 60
    $ws.Rows("1:$script:r").EntireRow.AutoFit() | Out-Null
    $ws.Activate(); $xl.ActiveWindow.DisplayGridlines = $false
    $wb.Worksheets.Item(1).Activate()
    $wb.Save(); $wb.Close($false)
    "process sheet added ($script:r rows)"
}
finally { $xl.ScreenUpdating = $true; $xl.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl); [GC]::Collect() }
