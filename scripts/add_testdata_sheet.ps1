# add_testdata_sheet.ps1 : add a "테스트데이터" sheet (오늘 수행한 두 회차의 실제 데이터) to the result workbook.
# The workbook comes from the DRM-wrapped company template, so it must be handled through Excel COM.
param(
    [Parameter(Mandatory = $true)][string]$Path,
    [string]$SheetName = "테스트데이터"
)
$ErrorActionPreference = "Stop"

$HDR = 2050429      # dark blue
$SEC = 15849925     # light blue
$IN  = 14804223     # light yellow (입력값)
$OK  = 14348258     # light green

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$xl.ScreenUpdating = $false
$script:r = 1

try {
    $wb = $xl.Workbooks.Open($Path)
    # 기존 동일 이름 시트가 있으면 삭제
    foreach ($s in @($wb.Worksheets)) { if ($s.Name -eq $SheetName) { $s.Delete() } }
    $after = $wb.Worksheets.Item(6)          # 6 = FAIL내역
    $ws = $wb.Worksheets.Add([Type]::Missing, $after)
    $ws.Name = $SheetName

    function Title($t) {
        $c = $ws.Cells.Item($script:r, 1); $c.Value2 = $t
        $c.Font.Size = 14; $c.Font.Bold = $true
        $script:r += 2
    }
    function Section($t) {
        $c = $ws.Cells.Item($script:r, 1); $c.Value2 = $t
        $c.Font.Size = 11; $c.Font.Bold = $true; $c.Font.Color = $HDR
        $script:r++
    }
    function Head($cells) {
        for ($i = 0; $i -lt $cells.Count; $i++) {
            $c = $ws.Cells.Item($script:r, $i + 1); $c.Value2 = $cells[$i]
            $c.Interior.Color = $HDR; $c.Font.Color = 16777215; $c.Font.Bold = $true
            $c.HorizontalAlignment = -4108; $c.Borders.LineStyle = 1
        }
        $script:r++
    }
    function Row($cells, $fill = $null) {
        for ($i = 0; $i -lt $cells.Count; $i++) {
            $c = $ws.Cells.Item($script:r, $i + 1)
            $c.NumberFormat = "@"                 # 상담번호가 숫자로 바뀌지 않도록
            $c.Value2 = $cells[$i]
            $c.Borders.LineStyle = 1
            $c.VerticalAlignment = -4160
            $c.WrapText = $true
            if ($fill) { $c.Interior.Color = $fill }
        }
        $script:r++
    }
    function Gap { $script:r++ }

    Title "테스트에 사용한 데이터 (2026-09-13 수행 2회차)"
    $ws.Cells.Item(2, 1).Value2 = "AS 무상지원금 일시불구독 개선 / KDQ 210 / 기간월 2026.07 / 모든 데이터는 테스트 후 원복 완료"
    $ws.Cells.Item(2, 1).Font.Italic = $true

    Section "0. 데이터 개요"
    Head @("항목", "내용")
    Row @("구독 구분 데이터", "ZCRM_CICSVC_NEW-SUBS_TYPE 을 'S' 로 변경 (회차별 4건). 원래값은 공백이며 테스트 후 공백으로 복원")
    Row @("확정 데이터", "ZCRM_CHRG_CNFRM. RFC ZCRMF03_RFC_ZCRM_CHRG_CNFRM 으로 등록(S)·취소(C). 테스트 후 0건")
    Row @("선정 조건", "무상(PRICETYPE=N) / CSLTNO_ITEM=1 / 제외사유(DRRSN) 공백 / 기사(TECHNAME) 있음 / SUBS_TYPE 공백 / 서비스점이 센터코드(CT..) 아님")
    Row @("대리점 확인", "KNVV 에 VKORG=ASKR 행 존재 (058 실행 시 '센터'에 넣을 VKBUR 확인), ZCRM_ASSPTFEEH1 2026.07 에 해당 대리점 행 없음")
    Row @("재현용 조회", "SE16N ZCRM_CICSVC_NEW / 조치일자 2026.07.01~2026.07.03 / 유무상구분자 N / 상담번호의 seq(행)번호 1")
    Gap

    # ================= 회차 1 =================
    Section "1. 회차 1 — RFC 수정 후 재테스트 (2026-09-12 23:25 ~ 2026-09-13 05:10)"
    Head @("역할", "상담번호(CSLTNO)", "대리점(KUNNR_SVC)", "센터", "조치일자", "클레임유형", "기사(TECHNAME)", "SUBS_TYPE 전→후")
    Row @("X-1", "2026070100278", "117786 제주진영서비스점", "CT11", "2026.07.01", "D1", "0000130784", "공백 → S → 공백") $IN
    Row @("X-2", "2026070100450", "117786", "CT11", "2026.07.02", "D1", "0000132094", "공백 → S → 공백") $IN
    Row @("X-3", "2026070100862", "117786", "CT11", "2026.07.01", "D1", "0000134384", "공백 → S → 공백") $IN
    Row @("Y-1", "2026070102721", "100011 (주)승보상사", "CT02", "2026.07.02", "A", "0000130144", "공백 → S → 공백") $IN
    Gap

    Section "1-1. 확정 데이터 등록·취소 순서 (RFC)"
    Head @("순서", "케이스", "처리구분", "상담번호", "확정일자(CFDATE1)", "RFC 결과", "그 시점 058S CFDATE1")
    Row @("1", "3-1", "-", "(확정 0건)", "-", "-", "X 공백 / Y 공백")
    Row @("2", "2-5 부분성공", "S", "2026070100278 + 9999999999999", "2026.08.05", "EV_RESULT=E (유효건만 저장)", "X 공백 (1/3)")
    Row @("3", "2-9 중복", "S", "2026070100450 (같은 줄 2회)", "2026.08.06", "EV_RESULT=S, 1건만 저장", "X 공백 (2/3) = 3-2")
    Row @("4", "2-1 정상등록", "S", "2026070100862 / 2026070102721 / 2026070100278(재전송)", "2026.08.07 / 2026.08.04 / 2026.08.05", "EV_RESULT=S, 3건", "X 2026.08.07 (3/3) = 3-3, Y 2026.08.04 = 3-5")
    Row @("5", "2-2 취소", "C", "2026070100862", "-", "EV_RESULT=S, 1건 삭제", "X 공백 (2/3) = 3-4")
    Row @("6", "취소 대상 없음", "C", "2026070100862 (재취소)", "-", "EV_RESULT=E '취소할 확정 데이터가 없습니다'", "변화 없음")
    Row @("7", "5-3 원복", "C", "2026070100278 / 0450 / 2721", "-", "EV_RESULT=S, 3건 삭제 → 테이블 0건", "X·Y 0행") $OK
    Gap

    Section "1-2. 화면 금액 (기간월 2026.07)"
    Head @("T-code", "대리점", "센터", "세팅 전", "세팅 후", "원복 후", "비고")
    Row @("ZCRM058A", "117786", "CT11", "190,000", "170,000", "190,000", "구독 3건(20,000) 만큼 감소")
    Row @("ZCRM058S", "117786", "CT11", "0행", "20,000", "0행", "심야 10,000 + 주간 10,000")
    Row @("ZCRM058A", "100011", "CT02", "65,000", "60,000", "65,000", "구독 1건(5,000) 만큼 감소")
    Row @("ZCRM058S", "100011", "CT02", "0행", "5,000", "0행", "")
    Row @("ZCRM060S", "117786", "CT11", "0행", "기사 3명 합계 20,000", "0행", "0000134384 고대의 5,000 / 0000132094 김태현 10,000 / 0000130784 김덕경 5,000")
    Row @("ZCRM060A", "117786", "CT11", "-", "기사 8명 합계 170,000", "-", "일반건")
    Gap

    Section "1-3. 오류 케이스용 부가 데이터"
    Head @("용도", "값", "설명", "DB 반영")
    Row @("존재하지 않는 상담번호", "9999999999999", "2-4 / 2-5 에서 사용", "없음")
    Row @("확정일 누락 (2-6)", "2026070100450 + 확정일 공백", "메시지 '확정일자 누락' 확인", "없음")
    Row @("미래일자 (2-7)", "2026070100450 + 2026.12.31", "메시지 '미래일자 입력 불가' 확인", "없음")
    Row @("SPMON 파생 검증 (5-2)", "2026062403120 (대리점 119410, 조치일자 2026.07.01)", "상담번호는 202606 이지만 SPMON 이 2026.07 로 저장되는지 확인", "등록 후 즉시 C 로 삭제")
    Gap

    # ================= 회차 2 =================
    Section "2. 회차 2 — 다른 테스트 데이터로 전체 재수행 (2026-09-13 06:45 ~ 07:00)"
    Head @("역할", "상담번호(CSLTNO)", "대리점(KUNNR_SVC)", "센터", "조치일자", "클레임유형", "기사(TECHNAME)", "SUBS_TYPE 전→후")
    Row @("X2-1", "2026063002988", "119410 대성상사(인천중구대리점)", "CT03", "2026.07.01", "A", "0000131701", "공백 → S → 공백") $IN
    Row @("X2-2", "2026063002994", "119410", "CT03", "2026.07.02", "A", "0000131701", "공백 → S → 공백") $IN
    Row @("X2-3", "2026063003684", "119410", "CT03", "2026.07.01", "D2", "0000131701", "공백 → S → 공백") $IN
    Row @("Y2-1", "2026070103092", "100013 노원경동나비엔(노원대리점)", "CT01", "2026.07.02", "D2", "0000131742", "공백 → S → 공백") $IN
    Gap
    $ws.Cells.Item($script:r, 1).Value2 = "※ X2 3건은 상담번호가 202606 으로 시작하지만 조치일자는 7월이다. SPMON 이 조치일자 기준으로 파생되는지(5-2) 본 데이터로 함께 검증했다."
    $ws.Cells.Item($script:r, 1).Font.Italic = $true
    $script:r += 2

    Section "2-1. 확정 데이터 등록·취소 순서 (RFC)"
    Head @("순서", "케이스", "처리구분", "상담번호", "확정일자(CFDATE1)", "RFC 결과", "그 시점 058S CFDATE1")
    Row @("1", "3-1", "-", "(확정 0건)", "-", "-", "X2 공백 / Y2 공백")
    Row @("2", "2-5 부분성공", "S", "2026063002988 + 9999999999999", "2026.08.10", "EV_RESULT=E (유효건만 저장)", "X2 공백 (1/3)")
    Row @("3", "2-9 중복", "S", "2026063002994 (같은 줄 2회)", "2026.08.11", "EV_RESULT=S, 1건만 저장", "X2 공백 (2/3) = 3-2")
    Row @("4", "2-1 정상등록", "S", "2026063003684 / 2026070103092 / 2026063002988(재전송)", "2026.08.12 / 2026.08.09 / 2026.08.10", "EV_RESULT=S, 3건", "X2 2026.08.12 (3/3) = 3-3, Y2 2026.08.09 = 3-5")
    Row @("5", "2-2 취소", "C", "2026063003684", "-", "EV_RESULT=S, 1건 삭제", "X2 공백 (2/3) = 3-4")
    Row @("6", "취소 대상 없음", "C", "2026063003684 (재취소)", "-", "EV_RESULT=E '취소할 확정 데이터가 없습니다'", "변화 없음")
    Row @("7", "5-3 원복", "C", "2026063002988 / 2026063002994 / 2026070103092", "-", "EV_RESULT=S, 3건 삭제 → 테이블 0건", "X2·Y2 0행") $OK
    Gap

    Section "2-2. 화면 금액 (기간월 2026.07)"
    Head @("T-code", "대리점", "센터", "세팅 전", "세팅 후", "원복 후", "비고")
    Row @("ZCRM058A", "119410", "CT03", "190,000", "175,000", "190,000", "구독 3건(15,000) 만큼 감소")
    Row @("ZCRM058S", "119410", "CT03", "0행", "15,000", "0행", "")
    Row @("ZCRM058A", "100013", "CT01", "65,000", "60,000", "65,000", "구독 1건(5,000) 만큼 감소")
    Row @("ZCRM058S", "100013", "CT01", "0행", "5,000", "0행", "")
    Row @("ZCRM060S", "119410", "CT03", "0행", "기사 1명 3건 15,000", "0행", "0000131701 (3건이 같은 기사)")
    Row @("ZCRM060A", "119410", "CT03", "-", "일반건 포함 9행", "-", "")
    Gap

    Section "2-3. 오류 케이스용 부가 데이터"
    Head @("용도", "값", "설명", "DB 반영")
    Row @("존재하지 않는 상담번호", "9999999999999", "2-4 / 2-5 에서 사용", "없음")
    Row @("확정일 누락 (2-6)", "2026063002994 + 확정일 공백", "메시지 '확정일자 누락' 확인", "없음")
    Row @("미래일자 (2-7)", "2026063002994 + 2026.12.31", "메시지 '미래일자 입력 불가' 확인", "없음")
    Gap

    # ================= 원복 =================
    Section "3. 원복 확인 (두 회차 공통)"
    Head @("대상", "확인 방법", "결과")
    Row @("ZCRM_CHRG_CNFRM", "SE16N 전체 조회", "0건 ('값이 없습니다')") $OK
    Row @("ZCRM_CICSVC_NEW-SUBS_TYPE", "지정 상담번호 재조회", "8건 모두 공백으로 복원") $OK
    Row @("ZCRM058A", "기간월 2026.07 / 대리점별 조회", "117786 190,000 · 100011 65,000 · 119410 190,000 · 100013 65,000 (세팅 전과 동일)") $OK
    Row @("ZCRM058S / ZCRM060S", "동일 조건 조회", "모두 0행") $OK
    Gap
    $ws.Cells.Item($script:r, 1).Value2 = "※ 이 시트의 상담번호·금액은 실행 로그(C:\sap-test\logs)와 화면 캡처(테스트증적 시트)로 확인한 실측값이다."
    $ws.Cells.Item($script:r, 1).Font.Italic = $true

    # 서식
    $w = @(14, 34, 30, 8, 12, 12, 16, 34)
    for ($i = 0; $i -lt $w.Count; $i++) { $ws.Columns.Item($i + 1).ColumnWidth = $w[$i] }
    $ws.Rows("1:$($script:r)").VerticalAlignment = -4160
    $ws.Range("A1").Select() | Out-Null
    $xl.ActiveWindow.DisplayGridlines = $false

    $wb.Worksheets.Item(1).Activate()
    $wb.Save()
    "sheet added: $SheetName (rows=$($script:r))"
    "sheets: " + (($wb.Worksheets | ForEach-Object { $_.Name }) -join " / ")
    $wb.Close($true)
}
finally {
    $xl.ScreenUpdating = $true
    $xl.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
    [GC]::Collect()
}
