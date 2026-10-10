# mk0930_e2e.ps1 : SR03000035400 2026-09-30 변경분(구독 정산유형 S01~S10) 통합 테스트 데이터 10건
#   같은 상담건을 ZCRM_CICSVC·ZCRM_CICSVC_NEW·ZCRM_REQ·ZCRM_CICCOST_M/SUB/ADD 에 넣고, 구독 건은 058S 용 CIC 대리점확정일(ZCRM_CHRG_CNFRM)도 넣는다.
#   조치일자·요청일자 2026.09.30, 생성자 ZTEST0930E, 부품 30031356A (센터 창고 재고: CT18 4 / CT11 3 / CT08 2 / CT19 2)
#   -Set T (오늘 통합 테스트, 2026093090011~020) | -Set U (내일 사용자 테스트용, 2026093090021~030)
#   -Part svc | new | req | m | sub | cnfrm
param([Parameter(Mandatory = $true)][ValidateSet('T', 'U')][string]$Set,
      [Parameter(Mandatory = $true)][ValidateSet('svc', 'new', 'req', 'm', 'sub', 'cnfrm')][string]$Part)
Set-Location C:\sap-test
[Console]::OutputEncoding = [Text.Encoding]::GetEncoding(949)

$SVCG = 'results\raw\S12_ZCRM_CICSVC_tpl_svc_gen.tsv'
$SVCS = 'results\raw\S12_ZCRM_CICSVC_tpl_svc_sub.tsv'
$NEW  = 'results\raw\S12_ZCRM_CICSVC_NEW_tpl_svcnew.tsv'
$REQ  = 'results\raw\S12_ZCRM_REQ_tpl_req.tsv'
$TM   = 'results\raw\S12_ZCRM_CICCOST_M_tpl_m.tsv'
$TSUB = 'results\raw\S12_ZCRM_CICCOST_SUB_tpl_sub.tsv'
$TADD = 'results\raw\S12_ZCRM_CICCOST_ADD_tpl_add.tsv'
$TCF  = 'results\raw\S12_ZCRM_CHRG_CNFRM_e6.tsv'

# 대리점 : 번호, 이름(ZTEST 표시용), 소속 센터(KNVV)
$D = @{
    '100184' = @('ZTEST_포항남부대리점', 'CT18')
    '100250' = @('ZTEST_서귀포대리점',   'CT11')
    '100215' = @('ZTEST_울산중앙대리점', 'CT08')
    '100085' = @('ZTEST_기흥수지대리점', 'CT19')
    # -Set U (내일 사용자 테스트용)
    '100234' = @('ZTEST_경동ENG진주',     'CT14')
    '100279' = @('ZTEST_강진대리점',     'CT23')
    '100256' = @('ZTEST_남원대리점',     'CT06')
    '100095' = @('ZTEST_대전동구중앙',   'CT05')
    '100005' = @('ZTEST_강서서부대리점', 'CT02')
    '100157' = @('ZTEST_대구달서대리점', 'CT09')
    '100200' = @('ZTEST_부산중구대리점', 'CT07')
    '100133' = @('ZTEST_원주대리점',     'CT10')
}
# 케이스 : 상담번호 꼬리, 조건, 대리점, 구독여부, SUBSWRTMONTH, 설치일, 보증만료일, 부품제조일
#   A 구독 N년이후 정상 / B 구독 N년이내 정상 / C 구독 기간 만료(부품제조일 과거 → Z16) / D 일반
$CASES_T = @(
    @('11', 'A', '100184', 'S', '12', '2023.09.01', '2026.09.01', '2025.12.13'),   # 36개월 3년이후
    @('12', 'B', '100184', 'S', '0',  '2024.09.01', '2026.09.01', '2025.12.13'),   # 24개월 2년이내
    @('13', 'C', '100184', 'S', '12', '2023.09.01', '2026.09.01', '2020.01.01'),   # 3년이후 + Z16
    @('14', 'D', '100184', '',  '0',  '',           '',           '2025.12.13'),
    @('15', 'A', '100250', 'S', '24', '2021.09.01', '2026.09.01', '2025.12.13'),   # 60개월 5년이후
    @('16', 'B', '100250', 'S', '0',  '2025.09.01', '2026.09.01', '2025.12.13'),   # 12개월 1년이내
    @('17', 'A', '100215', 'S', '12', '2022.09.01', '2026.09.01', '2025.12.13'),   # 48개월 4년이후
    @('18', 'C', '100215', 'S', '0',  '2024.09.01', '2026.09.01', '2020.01.01'),   # 2년이내 + Z16
    @('19', 'B', '100085', 'S', '0',  '2024.09.01', '2026.09.01', '2025.12.13'),   # 24개월 2년이내
    @('20', 'D', '100085', '',  '0',  '',           '',           '2025.12.13')
)
$CASES_U = @(
    @('21', 'A', '100234', 'S', '12', '2023.09.01', '2026.09.01', '2025.12.13'),   # 36개월 3년이후
    @('22', 'B', '100234', 'S', '0',  '2023.09.01', '2026.09.01', '2025.12.13'),   # 36개월 3년이내
    @('23', 'A', '100279', 'S', '12', '2024.09.01', '2026.09.01', '2025.12.13'),   # 24개월 2년이후
    @('24', 'C', '100279', 'S', '24', '2022.09.01', '2026.09.01', '2020.01.01'),   # 4년이후 + Z16
    @('25', 'A', '100256', 'S', '36', '2020.09.01', '2026.09.01', '2025.12.13'),   # 72개월 6년이후
    @('26', 'B', '100095', 'S', '0',  '2025.09.01', '2026.09.01', '2025.12.13'),   # 12개월 1년이내
    @('27', 'B', '100005', 'S', '0',  '2024.09.01', '2026.09.01', '2025.12.13'),   # 24개월 2년이내
    @('28', 'C', '100157', 'S', '0',  '2025.09.01', '2026.09.01', '2020.01.01'),   # 1년이내 + Z16
    @('29', 'D', '100200', '',  '0',  '',           '',           '2025.12.13'),
    @('30', 'D', '100133', '',  '0',  '',           '',           '2025.12.13')
)
$CASES = if ($Set -eq 'T') { $CASES_T } else { $CASES_U }
if (-not $CASES) { "no case set $Set"; exit 1 }

$COMMON = 'RSLTDATE=2026.09.30;REGDATE=2026.09.30;PRDDATE_NEW=2026.09.30;NAME1=ZTEST_20260930E;CITY1=ZTEST_20260930E;TELF1=01000000000;TELF2=01000000000;ERNAM=ZTEST0930E;AENAM=ZTEST0930E;ERDAT=2026.09.30;AEDAT=2026.09.30'
$ops = @()
foreach ($cs_row in $CASES) {  # 주의: PowerShell 변수는 대소문자 구분이 없어 $c 를 쓰면 공통값 변수와 충돌한다(2026-09-30 실제 발생)
    $c = $cs_row
    $cs = "20260930900" + $c[0]
    $dl = $D[$c[2]]
    $subs = "SUBS_TYPE=$($c[3]);SUBSWRTMONTH=$($c[4]);ISTDATE=$($c[5]);WRTDATE=$($c[6]);PRDDATE_PRT=$($c[7]);TAGNO=99$($c[0])"
    if ($c[3] -eq '') { $subs = "SUBS_TYPE=;SUBSWRTMONTH=0;WRTDATE=;PRDDATE_PRT=$($c[7]);TAGNO=99$($c[0])" }
    $dealer = "KUNNR_SVC=$($c[2]);SVCNAME=$($dl[0]);VKBUR=$($dl[1])"
    switch ($Part) {
        'svc' { $tpl = if ($c[3] -eq 'S') { $SVCS } else { $SVCG }; $ops += "ins|ZCRM_CICSVC|$cs|$tpl|$COMMON;$dealer;$subs" }
        'new' { $ops += "ins|ZCRM_CICSVC_NEW|$cs|$NEW|$COMMON;$dealer;$subs;NIGHTWEEK=D;TELF1=01000000000;TELF2=01000000000;ENGTEL=01000000000" }
        'req' { $ops += "ins|ZCRM_REQ|$cs|$REQ|REQNO=$cs;REQNO_ITEM=10;PRDREQDATE=2026.09.30;REQ_PERSON=ZTEST_20260930E;TAGNO1=99$($c[0]);ERNAM=ZTEST0930E;ERDAT=2026.09.30" }
        'm'   { $ops += "ins|ZCRM_CICCOST_M|$cs|$TM|CENTER_CD=$($dl[1]);CENTER_NM=ZTEST;TECHNAME=ZTEST;ERNAM=ZTEST0930E;ERDAT=2026.09.30;INDATE=2026.09.30;ISFREE=Y;TEXTINFO=ZTEST_20260930E $($c[1])" }
        'sub' { $ops += "ins|ZCRM_CICCOST_SUB|$cs|$TSUB|ERNAM=ZTEST0930E;ERDAT=2026.09.30"
                $ops += "ins|ZCRM_CICCOST_ADD|$cs|$TADD|NAME1=ZTEST_20260930E;CITY1=ZTEST;ERNAM=ZTEST0930E;ERDAT=2026.09.30" }
        'cnfrm' { if ($c[3] -eq 'S') { $ops += "ins|ZCRM_CHRG_CNFRM|$cs|$TCF|SPMON=2026.09;CFDATE1=2026.09.30" } }
    }
}
cscript //nologo scripts\S19_se16n_insert.vbs "e30${Set}_$Part" @ops
"exit=$LASTEXITCODE"
