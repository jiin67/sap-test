# add_process_sheets.ps1 : "프로세스흐름" + "테스트데이터" 시트 (SR03000035400 전체 프로세스 결과서용)
param([Parameter(Mandatory = $true)][string]$Path)
$ErrorActionPreference = "Stop"
$HDR = 2050429; $IN = 14804223; $OK = 14348258; $NG = 13551615; $BOX = 15790320

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false; $xl.ScreenUpdating = $false
$script:r = 1
try {
    $wb = $xl.Workbooks.Open($Path)

    function NewSheet($name, $after) {
        foreach ($s in @($wb.Worksheets)) { if ($s.Name -eq $name) { $s.Delete() } }
        $ws = $wb.Worksheets.Add([Type]::Missing, $wb.Worksheets.Item($after))
        $ws.Name = $name
        $script:r = 1
        return $ws
    }
    function Title($ws, $t) { $c = $ws.Cells.Item($script:r, 1); $c.Value2 = $t; $c.Font.Size = 14; $c.Font.Bold = $true; $script:r += 2 }
    function Section($ws, $t) { $c = $ws.Cells.Item($script:r, 1); $c.Value2 = $t; $c.Font.Size = 11; $c.Font.Bold = $true; $c.Font.Color = $HDR; $script:r++ }
    function Head($ws, $cells) { for ($i = 0; $i -lt $cells.Count; $i++) { $c = $ws.Cells.Item($script:r, $i + 1); $c.Value2 = $cells[$i]; $c.Interior.Color = $HDR; $c.Font.Color = 16777215; $c.Font.Bold = $true; $c.HorizontalAlignment = -4108; $c.Borders.LineStyle = 1 }; $script:r++ }
    function Row($ws, $cells, $fill = $null) { for ($i = 0; $i -lt $cells.Count; $i++) { $c = $ws.Cells.Item($script:r, $i + 1); $c.NumberFormat = "@"; $c.Value2 = $cells[$i]; $c.Borders.LineStyle = 1; $c.VerticalAlignment = -4160; $c.WrapText = $true; if ($fill) { $c.Interior.Color = $fill } }; $script:r++ }
    function Mono($ws, $t) { $c = $ws.Cells.Item($script:r, 1); $c.Value2 = $t; $c.Font.Name = "Consolas"; $c.Font.Size = 10; $script:r++ }
    function Note($ws, $t) { $c = $ws.Cells.Item($script:r, 1); $c.Value2 = $t; $c.Font.Italic = $true; $script:r += 2 }
    function Gap { $script:r++ }

    # ===================================================== 프로세스흐름
    $ws = NewSheet "프로세스흐름" 6
    Title $ws "SR03000035400 보일러 서비스 구독 대리점 비용 정산 - 업무 프로세스"
    $ws.Cells.Item(2, 1).Value2 = "보일러를 일시불로 산 고객이 3년 뒤 서비스 구독으로 전환한다. 구독 고객의 AS 는 무상으로 처리되는데, 그 비용을 대리점에 정산·지급하는 흐름을 기존 일반 건과 분리하는 것이 이번 개선의 목적이다."
    $ws.Cells.Item(2, 1).Font.Italic = $true
    $script:r = 4

    Section $ws "1. 개선 대상 오브젝트 (QA 소스 전수 검색 결과)"
    Head $ws @("단계", "오브젝트", "T-code", "역할", "구독 분기 방식")
    Row $ws @("① CIC 연동", "기능그룹 ZCRMG02`nFM ZCRMF02_RFC_ZCRM_CICSVC_ALL", "(화면의 '상담번호 동기화' 버튼)", "외부 CIC 상담시스템에서 상담·조치 데이터를 SAP 으로 가져온다", "Native SQL SELECT/INTO 에 SUBS_TYPE·WRTDATE·SUBSWRTMONTH 3필드 추가")
    Row $ws @("② 센터 매출", "ZCRMR177_1", "ZCRM177A(일반)`nZCRM177AS(구독)", "센터 기사가 쓴 무상 부품을 판매오더→납품→출고전기로 처리한다", "sy-tcode 로 모드 판별 → 상담번호 단위로 구독/일반 분리, 구독여부 TEXT 표시")
    Row $ws @("③ 회수품 접수", "ZCRMR052", "ZCRM052(일반)`nZCRM052S(구독)`nZCRM053/053S(조회)", "대리점이 반납한 회수품의 적합성을 판정하고 접수한다", "sy-tcode 로 모드 판별 → SUBS_TYPE 레인지로 조회 분리")
    Row $ws @("④ 회수품 정산", "ZCRMR171", "ZCRM171(일반)`nZCRM171S(구독)", "월 단위로 대리점별 회수품 지원금을 집계해 본사확정하고 FI 전표를 만든다", "ZCRM_RECPTF 키에 SUBS_TYPE 추가 → 같은 대리점도 구독/일반 행이 분리")
    Row $ws @("⑤ 무상지원금", "ZCRMR058_1", "ZCRM058A/059A(일반)`nZCRM058S/059S(구독)", "월 단위로 대리점 AS 무상지원금을 산정하고 확정·전표 처리한다", "ZCRM_ASSPTFEEH1 키에 SUBS_TYPE 추가, 조회·저장 모두 구독구분으로 분리")
    Row $ws @("⑥ 지원금 상세", "ZCRMR060_1", "ZCRM060A(일반)`nZCRM060S(구독)", "무상지원금 상세 리스트 조회 및 전송", "구독 모드에서만 기술료·출장비 컬럼 표시, SUBS_TYPE 레인지로 조회 분리")
    Note $ws "※ ZCRMR194(무상청구)는 이번 개선 대상이 아니다. 소스에 SR03000035400 표식이 없고, 테스트 데이터를 만들 때만 사용했다. 기능그룹 ZCRMG03 은 RFC 테스트용 INCLUDE(fbgenmac) 한 줄만 추가됐다."

    Section $ws "2. 데이터 흐름"
    Mono $ws "  [CIC 상담시스템 - 외부 MSSQL]"
    Mono $ws "        │  ① ZCRMF02_RFC_ZCRM_CICSVC_ALL (Native SQL)"
    Mono $ws "        ▼"
    Mono $ws "  ZCRM_CICSVC (회수품·센터 계열)        ZCRM_CICSVC_NEW (지원금 계열)"
    Mono $ws "        │                                        │"
    Mono $ws "        ├────────────┐                           │"
    Mono $ws "        │            │                           │"
    Mono $ws "   [센터 처리]   [대리점 처리]                     │"
    Mono $ws "        │            │                           │"
    Mono $ws "        ▼            ▼                           ▼"
    Mono $ws "  ② ZCRM177A/AS  ③ ZCRM052/052S           ⑤ ZCRM058A/S · 059A/S"
    Mono $ws "   오더·납품·출고    적합판정·접수                  지원금 산정 → 대리점확정"
    Mono $ws "        │            │                           │      → 본사확정 → FI 전표"
    Mono $ws "        ▼            ▼                           ▼"
    Mono $ws "  ZCRM_ORDER_AUTO  ZCRM_RECPT                ZCRM_ASSPTFEEH1 (키+SUBS_TYPE)"
    Mono $ws "                     │                           │"
    Mono $ws "                     ▼                           ▼"
    Mono $ws "              ④ ZCRM171/171S              ⑥ ZCRM060A/S 상세·전송"
    Mono $ws "               본사확정 → FI 전표"
    Mono $ws "                     ▼"
    Mono $ws "              ZCRM_RECPTF (키+SUBS_TYPE)"
    Gap
    Note $ws "※ 대리점에게 가는 돈은 두 갈래다. ④ 회수품 지원금(부품 1건당 43,450원, ZCRM_RECPTF)과 ⑤ AS 무상지원금(건당 5,000원 등 지급기준별, ZCRM_ASSPTFEEH1). 원천·집계단위·전표가 모두 달라 서로 간섭하지 않는다."

    Section $ws "3. 공통 설계 패턴"
    Head $ws @("항목", "내용")
    Row $ws @("모드 판별", "모든 프로그램이 sy-tcode 를 상수와 비교해 구독 모드를 정한다(예: ZCRMR177_1 은 GC_TCODE_SUBS='ZCRM177AS'). 구독/일반 T-code 가 같은 프로그램을 공유한다")
    Row $ws @("조회 분리", "RANGES 를 만들어 구독이면 I/EQ 'S', 일반이면 E/EQ 'S' 를 넣고 WHERE ... subs_type IN gr_subs 로 거른다")
    Row $ws @("저장 분리", "결과 테이블 키에 SUBS_TYPE 을 추가한다(ZCRM_RECPTF, ZCRM_ASSPTFEEH1). 같은 대리점·같은 월이라도 구독분과 일반분이 별도 행으로 남는다")
    Row $ws @("화면 표기", "내부코드 'S' 를 그대로 보여주지 않고 업무 문구로 바꾼다. ZCRMR177_1 은 유상/무상/구독무상(N년이후)/구독무상(N년이내), ZCRMR052 는 서비스 처리결과 컬럼")
    Row $ws @("구독 판정값", "SUBS_TYPE='S' 면 구독(도메인상 공백/S 만 허용). SUBSWRTMONTH(INT2, 부호 있음) 가 양수면 N년이후, 0 이하면 N년이내")
    Gap

    Section $ws "4. 정산 시점 규칙 (2026-09-21 업무 확인)"
    Head $ws @("항목", "내용")
    Row $ws @("접수 마감", "월 마감 이후에는 그 기간월로 추가 접수를 받지 않는다. SAP 으로 넘어오는 상담건은 결제까지 끝난 건이고, 이후 접수가 더 생기면 새 상담번호로 들어와 다음 기간월에 정산된다")
    Row $ws @("확정 시점", "④ 회수품 정산(ZCRM171/171S)의 본사확정·전표생성은 해당 기간월이 마감된 뒤에 수행한다. 따라서 확정된 뒤 같은 기간월로 접수가 추가되는 상황은 생기지 않는다")
    Row $ws @("유의 (프로그램 통제 없음)", "프로그램에는 기간월이 끝나기 전에 확정하지 못하게 막는 장치가 없다. 월 중에 잘못 확정하면 정산 화면 수량(ZCRM_RECPT 실시간 집계)과 ZCRM_RECPTF(확정 시점 스냅샷)가 어긋난다. 되돌리려면 노랑(확정, 전표 전)은 확정취소 후 재확정, 초록(전표 완료)은 전표취소까지 거쳐야 한다") $NG
    Note $ws "※ v0.01 에서 FAIL(P-01) 로 올렸던 건이다. 테스트가 기간월 2026.09 를 월 중(09-18)에 확정한 탓이며 운영 절차와 달라, 업무 확인 후 v0.02 에서 철회했다(E-22 PASS)."

    $ws.Columns.Item(1).ColumnWidth = 26
    $ws.Columns.Item(2).ColumnWidth = 34
    $ws.Columns.Item(3).ColumnWidth = 20
    $ws.Columns.Item(4).ColumnWidth = 48
    $ws.Columns.Item(5).ColumnWidth = 52
    $ws.Rows("1:$script:r").EntireRow.AutoFit() | Out-Null
    $ws.Activate(); $xl.ActiveWindow.DisplayGridlines = $false

    # ===================================================== 테스트데이터
    $ws = NewSheet "테스트데이터" 7
    Title $ws "E2E 테스트에 사용한 데이터 (2026-09-19, KDQ 210)"
    $ws.Cells.Item(2, 1).Value2 = "사용자 KKR003-1 / 기간월 2026.09 / 데이터와 생성 문서는 사용자 방침에 따라 삭제하지 않고 남겨 둠"
    $ws.Cells.Item(2, 1).Font.Italic = $true
    $script:r = 4

    Section $ws "1. 따라하기 순서"
    Head $ws @("단계", "T-code / 작업", "입력값")
    Row $ws @("1", "SE16N - ZCRM_CICSVC 삽입", "복사 원본 2026091890001(일반)/2026091890002(구독). CSLTNO·TAGNO·SUBS_TYPE·SUBSWRTMONTH·WRTDATE·KUNNR_SVC·VKBUR·RSLTDATE(2026.09.19)만 변경. CLMART='A', PRICETYPE='N', RESULTORG='B' 필수")
    Row $ws @("2", "SE16N - ZCRM_CICSVC_NEW 삽입", "복사 원본 2026063003075. 같은 값 + NIGHTWEEK='D'. 지원금(058/060)은 이 테이블을 읽는다")
    Row $ws @("3", "SE16N - ZCRM_REQ 삽입", "복사 원본 2026091890001. REQNO=상담번호, REQNO_ITEM=10, PRDREQDATE=2026.09.19, TAGNO1=꼬리표")
    Row $ws @("4", "ZCRM052S / ZCRM052", "서비스점조치 + 회수품 미접수 + 상담번호로 조회 → 적합수량 1, 책임사업장 1100, 회수품입고일자 2026.09.19, 접수부서·접수자 입력 → 저장(Ctrl+S) → 확인 팝업 '예'")
    Row $ws @("5", "ZCRM171S / ZCRM171", "정산 라디오 + 기간월 2026.09 → 실행 → 대리점 행 선택 → 본사확정 → 확인 → 다시 선택 → 전표생성 → 확인 → 전기일 달력에서 '오늘'")
    Row $ws @("6", "ZCRM058S / ZCRM058A", "전체 라디오 + 기간월 2026.09 + 센터 CT17 + 대리점 범위 → 실행")
    Row $ws @("7", "ZCRM060S / ZCRM060A", "기간월 2026.09 + 센터 CT17 + 대리점 범위 → 실행")
    Note $ws "※ 4단계에서 회수품입고일자(WEARDATS)를 빼면 'E 회수품 입고일자를 입력해주세요' 가 나온다. 5단계 전표생성 시 선택이 풀리므로 행을 다시 선택해야 한다."

    Section $ws "2. 생성한 상담건 (생성자 ZTEST0919E, 조치일자 2026.09.19)"
    Head $ws @("상담번호", "대리점 / 센터", "구독구분", "꼬리표", "접수번호", "정산 결과", "지원금(058)")
    Row $ws @("2026091993001", "100032 수도상사 / CT17", "구독 S", "9501", "2026091900001", "기존 확정분에 합산(월 중 확정 상태라 미반영 - 프로세스흐름 4 참조)", "구독 10,000 에 포함") $IN
    Row $ws @("2026091993002", "100032 수도상사 / CT17", "구독 S", "9502", "2026091900002", "동일", "구독 10,000 에 포함") $IN
    Row $ws @("2026091993003", "100032 수도상사 / CT17", "일반(공백)", "9503", "2026091900004", "본사확정 → 전표 3100069938", "일반 5,000") $OK
    Row $ws @("2026091993004", "100090 (주)평택경동 / CT04", "구독 S", "9504", "2026091900003", "기존 노랑에 합산(미확정)", "센터 CT04 - 이번 조회 범위 밖") $IN
    Row $ws @("2026091993005", "100047 (주)백일 / CT17", "구독 S", "9505", "(2026091900005)", "본사확정 → 전표 3100069937", "구독 5,000") $OK
    Note $ws "※ 공통: 자재 PNCB3540026K001, 부품 30031356A(룸콘 NR-63S), 부품가 39,500 → 회수품 정산 43,450(VAT 3,950). 무상지원금은 평일 기준 건당 5,000원."

    Section $ws "3. 생성된 문서·전표"
    Head $ws @("구분", "번호", "내용", "상태")
    Row $ws @("회수품 접수", "2026091900001~004", "ZCRM_RECPT, 적합판정일 2026.09.19", "남겨 둠") $IN
    Row $ws @("FI 전표(구독)", "3100069937", "100047 (주)백일 / 2건 86,900 / 전기일 2026.09.19 / 영업:A/S부품비", "정상, 역분개 없음") $OK
    Row $ws @("FI 전표(일반)", "3100069938", "100032 수도상사 / 1건 43,450 / 전기일 2026.09.19", "정상, 역분개 없음") $OK
    Row $ws @("정산 테이블", "ZCRM_RECPTF 2026.09", "100032 공백(435/3100069938), 100032 S(869/3100069925), 100047 S(869/3100069937)", "구독·일반 분리 저장 확인") $OK
    Gap

    Section $ws "4. 단계별 조회 결과 요약"
    Head $ws @("단계", "구독 화면", "일반 화면", "확인 내용")
    Row $ws @("② 센터매출 (177)", "7건", "3건", "상담번호 2026091990001~013 기준. 혼재 건은 구독 화면에만") $OK
    Row $ws @("③ 회수품접수 (052)", "3건", "1건", "E2E 상담번호 기준, 서로 섞이지 않음") $OK
    Row $ws @("④ 회수품정산 (171)", "12개 대리점", "10개 대리점", "대리점+구독구분 단위 집계") $OK
    Row $ws @("⑤ 무상지원금 (058)", "100032 10,000 / 100047 5,000", "100032 5,000", "같은 대리점이 구독·일반으로 나뉘어 산정") $OK
    Row $ws @("⑥ 지원금상세 (060)", "100032 2건 / 100047 1건", "100032 1건", "구독 화면에만 기술료·출장비 컬럼") $OK

    $ws.Columns.Item(1).ColumnWidth = 20
    $ws.Columns.Item(2).ColumnWidth = 30
    $ws.Columns.Item(3).ColumnWidth = 24
    $ws.Columns.Item(4).ColumnWidth = 14
    $ws.Columns.Item(5).ColumnWidth = 22
    $ws.Columns.Item(6).ColumnWidth = 40
    $ws.Columns.Item(7).ColumnWidth = 26
    $ws.Rows("1:$script:r").EntireRow.AutoFit() | Out-Null
    $ws.Activate(); $xl.ActiveWindow.DisplayGridlines = $false

    $wb.Worksheets.Item(1).Activate()
    $wb.Save()
    $wb.Close($false)
    "process + data sheets added to $Path"
}
finally {
    $xl.ScreenUpdating = $true
    $xl.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
}
