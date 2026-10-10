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

    Title "테스트에 사용한 데이터 (2026-09-14 ZCRMR194 → ZCRM052 테스트)"
    $ws.Cells.Item(2, 1).Value2 = "KDQ 210 / 사용자 KKR003-1 / 테스트 데이터와 처리 결과는 사용자 결정으로 삭제하지 않고 남겨 둠"
    $ws.Cells.Item(2, 1).Font.Italic = $true

    Section "0. 데이터 준비 방식"
    Head @("항목", "내용")
    Row @("생성 방법", "SE16N 편집으로 새 행 삽입(사용자 승인). 원본 행은 편집모드로 열지 않고, 조회 화면에서 읽은 값을 새 행에 입력")
    Row @("복사 원본", "TC1: 2026062600687 (대전센터 CT05, 전기오븐 GEONB412M00BL, 부품 30037654A - ZCRMR194 로 매출 완료된 정상건) / TC2·TC3: 2026063003075 (대리점 100090 (주)평택경동, 보일러 PNCB3540026K001, 부품 30031356A)")
    Row @("식별자", "상담번호 2026091490001~003 (오늘 날짜 + 9000번대, 기존 데이터 없음 확인), 생성자 ZTEST0914, 고객명·주소·전화 ZTEST_20260914 / 01000000000 으로 대체")
    Row @("구독 개월수", "TC2 는 시나리오 -3 대신 0 사용. SE16N 에 -3 또는 3- 입력 시 오류 없이 저장되지 않음(2회 확인, DB 미반영)")
    Gap

    Section "1. 생성한 테스트 데이터"
    Head @("케이스", "상담번호", "ZCRM_CICSVC", "ZCRM_CICCOST_M", "ZCRM_CICCOST_SUB", "ZCRM_CICCOST_ADD", "기대 표기")
    Row @("TC1 일시불", "2026091490001", "CLMART A / PRICETYPE N / RESULTORG A / 센터 CT05 / TAGNO 9141 / SUBS_TYPE 공백 / SUBSWRTMONTH 0", "ASKR / CT05 / INDATE 2026.09.14 / ISFREE Y / PAYMETHOD C / TAX_YN N", "ITEM 10 / 30037654A / QTY 1 / COST 0 KRW", "NAME1 ZTEST_20260914", "무상") $IN
    Row @("TC2 구독이내", "2026091490002", "CLMART A / PRICETYPE N / RESULTORG B / 서비스점 100090 / 센터 CT04 / TAGNO 9142 / SUBS_TYPE S / SUBSWRTMONTH 0", "ASKR / CT04 / INDATE 2026.09.14 / ISFREE Y", "ITEM 10 / 30031356A / QTY 1 / COST 0 KRW", "NAME1 ZTEST_20260914", "구독무상(N년이내)") $IN
    Row @("TC3 구독이후", "2026091490003", "CLMART A / PRICETYPE N / RESULTORG B / 서비스점 100090 / 센터 CT04 / TAGNO 9143 / SUBS_TYPE S / SUBSWRTMONTH 5", "ASKR / CT04 / INDATE 2026.09.14 / ISFREE Y", "ITEM 10 / 30031356A / QTY 1 / COST 0 KRW", "NAME1 ZTEST_20260914", "구독무상(N년이후)") $IN
    Gap

    Section "2. ZCRMR194 실행 (SE38, 2026.09.14 21:32)"
    Head @("항목", "입력 / 결과")
    Row @("선택조건", "S_VKORG ASKR / S_CSLTNO 2026091490001 ~ 2026091490003 / S_INDAT 2026.09.14 / S_VKBUR 공백")
    Row @("리스트", "청구 저장/확정완료 / 전체 : 3 / 에러 혹은 유상 : 2")
    Gap
    Head @("상담번호", "ZCRM_REQ", "무상매출 (ZCRM_CICCOST_M)", "ZCRM_ORDER_AUTO", "ZCRM_AUTO_ERRLOG", "판정")
    Row @("2026091490001", "2026091400001-10 / PRDREQDATE 2026.09.14 / REQ_CONF X / TAGNO1 9141 / TAGNO1_CHECK X", "상태 공백", "-", "MAKE_ORDER '자재 30037654A 가용수량이 부족합니다'", "PASS (청구 생성, 재고 부족으로 매출 제외)") $OK
    Row @("2026091490002", "2026091400001-20 / 2026.09.14 / X / 9142 / X", "STATUS1~3_FD 녹색 / VBELN_VA_FD 4293318 / VBELN_VL_FD 85857197", "4293318 / 85857197 / 4921917570", "-", "PASS") $OK
    Row @("2026091490003", "2026091400001-30 / 2026.09.14 / X / 9143 / X", "상태 공백", "-", "MAKE_ORDER '자재 30031356A 가용수량이 부족합니다'", "PASS (TC2 가 CT04 재고 1개 먼저 출고)") $OK
    Note "※ ZCRM_REQNO 2026.09.14 순번 1. 오더 4293318: 유형 ZFD1, 판매조직 ASKR/10, 영업소 CT04, PO번호 2026091490002. 납품 85857197: 유형 ZLFD, 실제출고일 2026.09.14."

    Section "3. 회수품 적합성 관리 조회 결과"
    Head @("T-code", "조건", "결과 건수", "상담번호", "서비스 처리결과", "판정")
    Row @("ZCRM052", "센터조치 / 미접수 / 상담번호 TC1", "1", "2026091490001", "무상", "PASS") $OK
    Row @("ZCRM052", "센터조치 / 미접수 / 상담번호 TC1~TC3", "1", "2026091490001", "무상", "PASS") $OK
    Row @("ZCRM052", "서비스점조치 / 미접수 / 상담번호 TC1~TC3", "0", "-", "- (구독건 제외)", "PASS") $OK
    Row @("ZCRM052S", "서비스점조치 / 미접수 / 상담번호 TC1~TC3 (STEP8 전)", "2", "2026091490002 / 2026091490003", "구독무상(N년이내) / 구독무상(N년이후)", "PASS (구독구분 컬럼 미표시 - C-05 FAIL)") $OK
    Row @("ZCRM053", "센터조치 / 미접수 / 상담번호 TC1~TC3", "1", "2026091490001", "무상", "PASS") $OK
    Row @("ZCRM053S", "실행", "-", "-", "덤프 LOAD_PROGRAM_NOT_FOUND", "FAIL (D-02)") $NG
    Gap

    Section "4. STEP8 회수품 접수 저장 (ZCRM052S)"
    Head @("항목", "값")
    Row @("대상", "2026091490002 (TC2) - 미접수 조회")
    Row @("입력", "적합수량 1 / 부적합수량 0 / 책임사업장 1100 / 회수품입고일자 2026.09.14 / 접수부서명·접수자 ZTEST_20260914")
    Row @("결과", "확인 팝업 '예' → '정상적으로 저장되었습니다.'") $OK
    Row @("ZCRM_RECPT", "접수번호 2026091400001 / 품목 10 / 요청번호 2026091400001-20 / 접수일자 2026.09.14 / NCF Y / RSPC 1100 / STATUS X / NETWR 39,500 KRW") $OK
    Row @("접수 조회", "ZCRM052S 회수품 접수 + 접수일자 2026.09.14 → TC2 1건 표시, 미접수 목록에서는 제외(TC3 만 남음)") $OK
    Gap

    Section "5. KDQ 에 남아 있는 데이터 (정리 시 참고)"
    Head @("테이블 / 문서", "식별 조건", "건수")
    Row @("ZCRM_CICSVC / ZCRM_CICCOST_M / ZCRM_CICCOST_SUB / ZCRM_CICCOST_ADD", "CSLTNO 2026091490001 ~ 2026091490003 (생성자 ZTEST0914)", "각 3")
    Row @("ZCRM_REQ", "REQNO 2026091400001 (품목 10/20/30)", "3")
    Row @("ZCRM_REQNO", "CURRDT 2026.09.14 / REQNO 1", "1")
    Row @("ZCRM_AUTO_ERRLOG", "CSLTNO 2026091490001, 2026091490003", "2")
    Row @("ZCRM_ORDER_AUTO", "CSLTNO 2026091490002 (4293318 / 85857197 / 4921917570)", "3")
    Row @("SD / MM", "판매오더 4293318 (ZFD1), 납품 85857197 (출고완료), 자재전표 4921917570 - 부품 30031356A 1EA 출고", "-")
    Row @("ZCRM_RECPT (+ 입고일자 ZCRM_RECPT_WEAR)", "RECPTNO 2026091400001 (REMARK1·REMARK2 ZTEST_20260914)", "1")

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
