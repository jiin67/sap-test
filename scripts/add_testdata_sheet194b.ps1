# add_testdata_sheet194.ps1 : "테스트데이터" sheet for the 2026-09-14 ZCRMR194 → ZCRM052 test (v0.01).
param(
    [Parameter(Mandatory = $true)][string]$Path,
    [string]$SheetName = "테스트데이터"
)
$ErrorActionPreference = "Stop"
$HDR = 2050429; $IN = 14804223; $OK = 14348258; $NG = 13551615

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false; $xl.ScreenUpdating = $false
$script:r = 1
try {
    $wb = $xl.Workbooks.Open($Path)
    foreach ($s in @($wb.Worksheets)) { if ($s.Name -eq $SheetName) { $s.Delete() } }
    $ws = $wb.Worksheets.Add([Type]::Missing, $wb.Worksheets.Item(6))
    $ws.Name = $SheetName

    function Title($t) { $c = $ws.Cells.Item($script:r, 1); $c.Value2 = $t; $c.Font.Size = 14; $c.Font.Bold = $true; $script:r += 2 }
    function Section($t) { $c = $ws.Cells.Item($script:r, 1); $c.Value2 = $t; $c.Font.Size = 11; $c.Font.Bold = $true; $c.Font.Color = $HDR; $script:r++ }
    function Head($cells) { for ($i = 0; $i -lt $cells.Count; $i++) { $c = $ws.Cells.Item($script:r, $i + 1); $c.Value2 = $cells[$i]; $c.Interior.Color = $HDR; $c.Font.Color = 16777215; $c.Font.Bold = $true; $c.HorizontalAlignment = -4108; $c.Borders.LineStyle = 1 }; $script:r++ }
    function Row($cells, $fill = $null) { for ($i = 0; $i -lt $cells.Count; $i++) { $c = $ws.Cells.Item($script:r, $i + 1); $c.NumberFormat = "@"; $c.Value2 = $cells[$i]; $c.Borders.LineStyle = 1; $c.VerticalAlignment = -4160; $c.WrapText = $true; if ($fill) { $c.Interior.Color = $fill } }; $script:r++ }
    function Gap { $script:r++ }
    function Note($t) { $c = $ws.Cells.Item($script:r, 1); $c.Value2 = $t; $c.Font.Italic = $true; $script:r += 2 }

    Title "테스트에 사용한 데이터 (2026-09-18 ZCRMR194 → ZCRM052 재테스트)"
    $ws.Cells.Item(2, 1).Value2 = "KDQ 210 / 사용자 KKR003-1 / 1차(09-14) FAIL 3건 수정 후 새 데이터로 전체 재수행. 데이터는 삭제하지 않고 남겨 둠"
    $ws.Cells.Item(2, 1).Font.Italic = $true

    Section "0. 1차 FAIL 항목 재테스트 결과"
    Head @("ID", "1차(09-14) 현상", "재테스트(09-18) 결과", "판정")
    Row @("D-02", "ZCRM053S 실행 시 덤프 LOAD_PROGRAM_NOT_FOUND (TSTC 가 ZCRMR053 을 가리킴)", "TSTC ZCRM053S → ZCRMR052. 정상 실행, 구독 2건 조회, 제목 '[구독전용] 회수품 적합성 관리 조회 (2)', 저장 비활성", "PASS") $OK
    Row @("C-05", "ZCRM052S 결과에 구독구분 컬럼이 숨겨져 있음", "ZCRM052S·053S 첫 컬럼에 '구독' = S 표시", "PASS") $OK
    Row @("C-07", "ZCRM052S 결과 제목에 조회건수 없음", "'[구독전용] 회수품 적합성 관리 (2)' 로 건수 표시", "PASS") $OK
    Row @("B-03 (신규)", "1차에는 ZCRM052(비구독)에 구독구분 컬럼 없음", "ZCRM052·053 에도 '구독' 컬럼이 표시됨(값 공백) - 공통 레이아웃 영향으로 추정", "FAIL") $NG
    Gap

    Section "1. 생성한 테스트 데이터 (SE16N 삽입, 생성자 ZTEST0918B)"
    Head @("케이스", "상담번호", "ZCRM_CICSVC", "ZCRM_CICCOST_M", "ZCRM_CICCOST_SUB", "ZCRM_CICCOST_ADD", "기대 표기")
    Row @("TC1 일시불", "2026091891001", "CLMART A / PRICETYPE N / RESULTORG A / 센터 CT05 / TAGNO 9281 / SUBS_TYPE 공백 / SUBSWRTMONTH 0", "ASKR / CT05 / INDATE 2026.09.18 / ISFREE Y / PAYMETHOD C / TAX_YN N", "ITEM 10 / 30037654A / QTY 1 / COST 0 KRW", "NAME1 ZTEST_20260918B", "무상") $IN
    Row @("TC2 구독이내", "2026091891002", "CLMART A / PRICETYPE N / RESULTORG B / 서비스점 100090 / 센터 CT04 / TAGNO 9282 / SUBS_TYPE S / SUBSWRTMONTH 0", "ASKR / CT04 / INDATE 2026.09.18 / ISFREE Y", "ITEM 10 / 30031356A / QTY 1 / COST 0 KRW", "NAME1 ZTEST_20260918B", "구독무상(N년이내)") $IN
    Row @("TC3 구독이후", "2026091891003", "CLMART A / PRICETYPE N / RESULTORG B / 서비스점 100090 / 센터 CT04 / TAGNO 9283 / SUBS_TYPE S / SUBSWRTMONTH 5", "ASKR / CT04 / INDATE 2026.09.18 / ISFREE Y", "ITEM 10 / 30031356A / QTY 1 / COST 0 KRW", "NAME1 ZTEST_20260918B", "구독무상(N년이후)") $IN
    Note "※ 복사 원본: TC1 = 2026062600687, TC2·TC3 = 2026063003075 (1차와 동일). TC2 의 SUBSWRTMONTH 는 SE16N 이 음수를 저장하지 못해 0 사용."

    Section "2. ZCRMR194 실행 (SE38, 2026.09.18 06:01)"
    Head @("항목", "입력 / 결과")
    Row @("선택조건", "S_VKORG ASKR / S_CSLTNO 2026091891001 ~ 2026091891003 / S_INDAT 2026.09.18 / S_VKBUR 공백")
    Row @("리스트", "청구 저장/확정완료 / 전체 : 3 / 에러 혹은 유상 : 3")
    Gap
    Head @("상담번호", "ZCRM_REQ", "무상매출 (ZCRM_CICCOST_M)", "ZCRM_AUTO_ERRLOG", "판정")
    Row @("2026091891001", "2026091800001-10 / PRDREQDATE 2026.09.18 / REQ_CONF X / TAGNO1 9281 / CHECK X", "상태 공백", "MAKE_ORDER '자재 30037654A 가용수량이 부족합니다'", "PASS (청구 생성, 재고 부족으로 매출 제외)") $OK
    Row @("2026091891002", "2026091800001-20 / 2026.09.18 / X / 9282 / X", "상태 공백", "MAKE_ORDER '자재 30031356A 가용수량이 부족합니다'", "PASS") $OK
    Row @("2026091891003", "2026091800001-30 / 2026.09.18 / X / 9283 / X", "상태 공백", "MAKE_ORDER '자재 30031356A 가용수량이 부족합니다'", "PASS") $OK
    Note "※ 이번에는 세 건 모두 가용재고가 없어 무상오더가 만들어지지 않았다(1차 때 CT04 재고 1개 출고). 오더→납품→출고 정상 경로는 1차 결과서(v0.01) 참조. ZCRM_REQNO 2026.09.18 순번 1."

    Section "3. 회수품 적합성 관리 조회 결과"
    Head @("T-code", "조건", "결과 건수 / 제목", "상담번호", "서비스 처리결과", "판정")
    Row @("ZCRM052", "센터조치 / 미접수 / 상담번호 TC1~TC3", "1 / 회수품적합성관리 (1)", "2026091891001", "무상", "PASS (단, 구독 컬럼 표시 - B-03)") $OK
    Row @("ZCRM052", "서비스점조치 / 미접수 / 상담번호 TC1~TC3", "0", "-", "- (구독건 제외)", "PASS") $OK
    Row @("ZCRM052S", "서비스점조치 / 미접수 / 상담번호 TC1~TC3", "2 / [구독전용] 회수품 적합성 관리 (2)", "2026091891002 / 2026091891003", "구독무상(N년이내) / 구독무상(N년이후), 구독 S", "PASS") $OK
    Row @("ZCRM053", "센터조치 / 미접수 / 상담번호 TC1~TC3", "1 / 회수품적합성관리 조회 (1)", "2026091891001", "무상, 저장 비활성", "PASS") $OK
    Row @("ZCRM053S", "서비스점조치 / 미접수 / 상담번호 TC1~TC3", "2 / [구독전용] 회수품 적합성 관리 조회 (2)", "2026091891002 / 2026091891003", "구독무상(N년이내) / (N년이후), 저장 비활성", "PASS") $OK
    Gap

    Section "4. STEP8 회수품 접수 저장 (ZCRM052S)"
    Head @("항목", "값")
    Row @("대상", "2026091891002 (TC2) - 미접수 조회")
    Row @("입력", "적합수량 1 / 책임사업장 1100 / 회수품입고일자 2026.09.18 / 접수부서명·접수자 ZTEST_20260918B")
    Row @("결과", "확인 팝업 '예' → '정상적으로 저장되었습니다.'") $OK
    Row @("ZCRM_RECPT", "접수번호 2026091800014 / 품목 10 / 요청번호 2026091800001-20 / 접수일자 2026.09.18 / NCF Y / RSPC 1100 / STATUS X / NETWR 39,500 KRW") $OK
    Row @("접수 조회", "ZCRM052S 회수품 접수 + 접수일자 2026.09.18 → TC2 1건 표시") $OK
    Gap

    Section "5. KDQ 에 남아 있는 데이터 (정리 시 참고)"
    Head @("테이블", "식별 조건", "건수")
    Row @("ZCRM_CICSVC / ZCRM_CICCOST_M / ZCRM_CICCOST_SUB / ZCRM_CICCOST_ADD", "CSLTNO 2026091891001 ~ 2026091891003 (생성자 ZTEST0918B)", "각 3")
    Row @("ZCRM_REQ", "REQNO 2026091800001 (품목 10/20/30)", "3")
    Row @("ZCRM_REQNO", "CURRDT 2026.09.18 / REQNO 1", "1")
    Row @("ZCRM_AUTO_ERRLOG", "CSLTNO 2026091891001 ~ 003", "3")
    Row @("ZCRM_RECPT (+ 입고일자 ZCRM_RECPT_WEAR)", "RECPTNO 2026091800014 (REMARK1·REMARK2 ZTEST_20260918B)", "1")
    Note "※ 2026091891002 의 접수(적합판정일 2026.09.18)는 ZCRM171S 2026.09 정산의 대리점 100090 구독 수량에 1건 더해진다(09-18 ZCRMR171 테스트 데이터와 같은 달·대리점)."

    $w = @(22, 30, 36, 34, 30, 26, 22)
    for ($i = 0; $i -lt $w.Count; $i++) { $ws.Columns.Item($i + 1).ColumnWidth = $w[$i] }
    $ws.Activate(); $xl.ActiveWindow.DisplayGridlines = $false
    $wb.Worksheets.Item(1).Activate()
    $wb.Save()
    "sheet added: $SheetName rows=$($script:r)"
    $wb.Close($true)
}
finally {
    $xl.ScreenUpdating = $true; $xl.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl); [GC]::Collect()
}
