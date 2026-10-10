# mk1002_e2e.ps1 : SR03000035400 v1.0 전체 프로세스 테스트 데이터 10건 (2026-10-02, 현재 설정 S01~S10 기준)
#   같은 상담건을 ZCRM_CICSVC·ZCRM_CICSVC_NEW·ZCRM_REQ·ZCRM_CICCOST_M/SUB/ADD 에 넣고, 구독 건은 058S 용 CIC 대리점확정일(ZCRM_CHRG_CNFRM)도 넣는다.
#   조치일자·요청일자 2026.10.02, 생성자 ZTEST1002E, 부품 30031356A (센터 창고 재고: CT18 4 / CT08 2 / CT14 2 / CT23 2)
#   -Set V (2026100290001~010, 대리점 100187·100216·100241·100282 - 이전 테스트 미사용 정상점)
#   -Part svc | new | req | m | sub | cnfrm
param([Parameter(Mandatory = $true)][ValidateSet('V','W','W2','W3','W4','W5','W6','W7')][string]$Set,
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
    '100187' = @('ZTEST_포항북부대리점', 'CT18')
    '100216' = @('ZTEST_울산남구대리점', 'CT08')
    '100241' = @('ZTEST_통영대리점',     'CT14')
    '100282' = @('ZTEST_보성대리점',     'CT23')
    '100085' = @('ZTEST_기흥수지대리점', 'CT19')   # -Set W (D-09 재확인용, 재고 CT19 2)
    '100005' = @('ZTEST_강서서부대리점', 'CT02')   # -Set W2 (CT19 ATP 선점으로 CT02 로 재시도, 재고 1)
    '100590' = @('ZTEST_CT02대리점2', 'CT02')   # -Set W5 (0원 혼합 선택용 두 번째 CT02 정상점, KNVV AUFSD 공백)
    '100879' = @('ZTEST_CT02대리점3', 'CT02')   # -Set W5 대체 후보
}
# 케이스 : 상담번호 꼬리, 조건, 대리점, 구독여부, SUBSWRTMONTH, 설치일, 보증만료일, 부품제조일
#   A 구독 N년이후 정상 / B 구독 N년이내 정상 / C 구독 기간 만료(부품제조일 과거 → Z16) / D 일반
$CASES_V = @(
    @('01', 'A', '100187', 'S', '12', '2023.10.01', '2026.10.01', '2025.12.13'),   # 36개월 3년이후
    @('02', 'B', '100187', 'S', '0',  '2024.10.01', '2026.10.01', '2025.12.13'),   # 24개월 2년이내
    @('03', 'C', '100187', 'S', '12', '2023.10.01', '2026.10.01', '2020.01.01'),   # 3년이후 + Z16
    @('04', 'D', '100187', '',  '0',  '',           '',           '2025.12.13'),
    @('05', 'A', '100216', 'S', '24', '2021.10.01', '2026.10.01', '2025.12.13'),   # 60개월 5년이후
    @('06', 'D', '100216', '',  '0',  '',           '',           '2025.12.13'),
    @('07', 'A', '100241', 'S', '12', '2022.10.01', '2026.10.01', '2025.12.13'),   # 48개월 4년이후
    @('08', 'B', '100241', 'S', '0',  '2025.10.01', '2026.10.01', '2025.12.13'),   # 12개월 1년이내
    @('09', 'C', '100282', 'S', '0',  '2024.10.01', '2026.10.01', '2020.01.01'),   # 2년이내 + Z16
    @('10', 'D', '100282', '',  '0',  '',           '',           '2025.12.13')
)
$CASES_W = @(
    ,@('13', 'A', '100085', 'S', '12', '2023.10.01', '2026.10.01', '2025.12.13')    # 구독 3년이후, D-09 재확인  (케이스 1건이면 앞에 , 를 붙여야 배열이 풀리지 않는다 - 2026-10-02 사고)
)
$CASES_W2 = @(
    ,@('14', 'A', '100005', 'S', '12', '2023.10.01', '2026.10.01', '2025.12.13')    # 구독 3년이후, D-09 납품·출고 단계 확인
)
$CASES_W3 = @(
    ,@('15', 'A', '100005', 'S', '12', '2023.10.01', '2026.10.01', '2025.12.13')    # 구독 3년이후, D-10 다건 합산 확인(014 와 같은 대리점, 기술료 7,000/출장비 20,000 은 S26 로)
)
$CASES_W4 = @(
    ,@('16', 'A', '100005', 'S', '12', '2023.10.01', '2026.10.01', '2025.12.13')    # 구독 3년이후, D-10 다건 합산 증명용 3번째 상담(기술료 3,000/출장비 5,000 은 S26 로)
)
$CASES_W5 = @(
    ,@('17', 'A', '100590', 'S', '12', '2023.10.01', '2026.10.01', '2025.12.13')    # 구독 3년이후, 0원 대리점(100005)과 혼합 선택 확정용(기술료 4,000/출장비 6,000 은 S26). cnfrm 은 넣지 않음(취소 가능하게)
)
$CASES_W6 = @(
    ,@('18', 'A', '100241', 'S', '12', '2023.10.01', '2026.10.01', '2025.12.13')    # 구독, 2026.09 CT14 혼합 선택용(100234 0원 + 100241 금액). 삽입 후 S26 으로 INDATE 2026.09.30 / 기술료 4,000 / 출장비 6,000. cnfrm 없음(취소 가능)
)
$CASES_W7 = @(
    ,@('19', 'B', '100005', 'S', '0', '2024.10.01', '2026.10.01', '2025.12.13')    # 구독 2년이내(SUBSWRTMONTH 0), D-16 확인용 - 기술료 9,000/출장비 9,000 은 S26 로
)
$CASES = $null
if ($Set -eq 'W') { $CASES = $CASES_W } elseif ($Set -eq 'W2') { $CASES = $CASES_W2 } elseif ($Set -eq 'W3') { $CASES = $CASES_W3 } elseif ($Set -eq 'W4') { $CASES = $CASES_W4 } elseif ($Set -eq 'W5') { $CASES = $CASES_W5 } elseif ($Set -eq 'W6') { $CASES = $CASES_W6 } elseif ($Set -eq 'W7') { $CASES = $CASES_W7 } else { $CASES = $CASES_V }   # if 식 대입은 1건 배열을 풀어 버린다
$COMMON = 'RSLTDATE=2026.10.02;REGDATE=2026.10.02;PRDDATE_NEW=2026.10.02;NAME1=ZTEST_20261002E;CITY1=ZTEST_20261002E;TELF1=01000000000;TELF2=01000000000;ERNAM=ZTEST1002E;AENAM=ZTEST1002E;ERDAT=2026.10.02;AEDAT=2026.10.02'
$ops = @()
foreach ($cs_row in $CASES) { if ($cs_row.Count -ne 8 -or $cs_row[0].Length -ne 2) { "bad case row: $cs_row"; exit 1 }  # 주의: PowerShell 변수는 대소문자 구분이 없어 $c 를 쓰면 공통값 변수와 충돌한다(2026-09-30 실제 발생)
    $c = $cs_row
    $cs = "20261002900" + $c[0]
    $dl = $D[$c[2]]
    $subs = "SUBS_TYPE=$($c[3]);SUBSWRTMONTH=$($c[4]);ISTDATE=$($c[5]);WRTDATE=$($c[6]);PRDDATE_PRT=$($c[7]);TAGNO=99$($c[0])"
    if ($c[3] -eq '') { $subs = "SUBS_TYPE=;SUBSWRTMONTH=0;WRTDATE=;PRDDATE_PRT=$($c[7]);TAGNO=99$($c[0])" }
    $dealer = "KUNNR_SVC=$($c[2]);SVCNAME=$($dl[0]);VKBUR=$($dl[1])"
    switch ($Part) {
        'svc' { $tpl = if ($c[3] -eq 'S') { $SVCS } else { $SVCG }; $ops += "ins|ZCRM_CICSVC|$cs|$tpl|$COMMON;$dealer;$subs" }
        'new' { $ops += "ins|ZCRM_CICSVC_NEW|$cs|$NEW|$COMMON;$dealer;$subs;NIGHTWEEK=D;TELF1=01000000000;TELF2=01000000000;ENGTEL=01000000000" }
        'req' { $ops += "ins|ZCRM_REQ|$cs|$REQ|REQNO=$cs;REQNO_ITEM=10;PRDREQDATE=2026.10.02;REQ_PERSON=ZTEST_20261002E;TAGNO1=99$($c[0]);ERNAM=ZTEST1002E;ERDAT=2026.10.02" }
        'm'   { $ops += "ins|ZCRM_CICCOST_M|$cs|$TM|CENTER_CD=$($dl[1]);CENTER_NM=ZTEST;TECHNAME=ZTEST;ERNAM=ZTEST1002E;ERDAT=2026.10.02;INDATE=2026.10.02;ISFREE=Y;TEXTINFO=ZTEST_20261002E $($c[1])" }
        'sub' { $ops += "ins|ZCRM_CICCOST_SUB|$cs|$TSUB|ERNAM=ZTEST1002E;ERDAT=2026.10.02"
                $ops += "ins|ZCRM_CICCOST_ADD|$cs|$TADD|NAME1=ZTEST_20261002E;CITY1=ZTEST;ERNAM=ZTEST1002E;ERDAT=2026.10.02" }
        'cnfrm' { if ($c[3] -eq 'S') { $ops += "ins|ZCRM_CHRG_CNFRM|$cs|$TCF|SPMON=2026.10;CFDATE1=2026.10.02" } }
    }
}
cscript //nologo scripts\S19_se16n_insert.vbs "e02${Set}_$Part" @ops
"exit=$LASTEXITCODE"
