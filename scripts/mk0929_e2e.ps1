# mk0929_e2e.ps1 : SR03000035400 통합 진행 확인 데이터 (2026-09-29)
#   같은 상담건을 ZCRM_CICSVC 와 ZCRM_CICSVC_NEW 에 모두 넣어 ZCRM177A/AS, ZCRM052/052S, ZCRM171/171S, ZCRM058A/S, ZCRM060A/S 가 함께 진행되는지 본다.
#   대리점 100033 코지하우징(정상점, KNVV CT01, LFA1 123033, 이전 테스트 미사용) / 센터 CT01 (1120/CT01 재고 3EA) / 생성자 ZTEST0929E
#   상담번호 2026092990011 구독 3년이후 / 012 구독 2년이내 / 013 일반   (구독기간 12개월 단위)
#   -Part svc | new | req | m | sub
param([Parameter(Mandatory = $true)][ValidateSet('svc', 'new', 'req', 'm', 'sub', 'cnfrm')][string]$Part)
Set-Location C:\sap-test
[Console]::OutputEncoding = [Text.Encoding]::GetEncoding(949)

$SVCG = 'results\raw\S12_ZCRM_CICSVC_tpl_svc_gen.tsv'
$SVCS = 'results\raw\S12_ZCRM_CICSVC_tpl_svc_sub.tsv'
$NEW = 'results\raw\S12_ZCRM_CICSVC_NEW_tpl_svcnew.tsv'
$REQ = 'results\raw\S12_ZCRM_REQ_tpl_req.tsv'
$TM = 'results\raw\S12_ZCRM_CICCOST_M_tpl_m.tsv'
$TSUB = 'results\raw\S12_ZCRM_CICCOST_SUB_tpl_sub.tsv'
$TADD = 'results\raw\S12_ZCRM_CICCOST_ADD_tpl_add.tsv'

$C = 'RSLTDATE=2026.09.29;REGDATE=2026.09.29;PRDDATE_NEW=2026.09.29;NAME1=ZTEST_20260929E;CITY1=ZTEST_20260929E;ERNAM=ZTEST0929E;AENAM=ZTEST0929E;ERDAT=2026.09.29;AEDAT=2026.09.29'
$D33 = 'KUNNR_SVC=100033;SVCNAME=ZTEST_코지하우징;VKBUR=CT01'
# 상담건별 구독 속성 (두 테이블에 같은 값)
$S1 = 'TAGNO=9911;SUBS_TYPE=S;SUBSWRTMONTH=12;ISTDATE=2023.09.01;WRTDATE=2026.09.01;PRDDATE_PRT=2025.12.13'   # 3년이후
$S2 = 'TAGNO=9912;SUBS_TYPE=S;SUBSWRTMONTH=0;ISTDATE=2024.09.01;WRTDATE=2026.09.01;PRDDATE_PRT=2025.12.13'    # 2년이내
$S3 = 'TAGNO=9913;SUBS_TYPE=;SUBSWRTMONTH=0;WRTDATE=;PRDDATE_PRT=2025.12.13'                                    # 일반
$CM = 'CENTER_CD=CT01;CENTER_NM=ZTEST;TECHNAME=ZTEST;ERNAM=ZTEST0929E;ERDAT=2026.09.29;INDATE=2026.09.29;ISFREE=Y'

switch ($Part) {
    'svc' {
        $ops = @(
            "ins|ZCRM_CICSVC|2026092990011|$SVCS|$C;$D33;$S1",
            "ins|ZCRM_CICSVC|2026092990012|$SVCS|$C;$D33;$S2",
            "ins|ZCRM_CICSVC|2026092990013|$SVCG|$C;$D33;$S3"
        )
        cscript //nologo scripts\S19_se16n_insert.vbs e3_svc @ops
    }
    'new' {
        # ZCRM058_1 / ZCRM060_1 이 읽는 테이블. PRICETYPE='N', CLMART='A', CSLTNO_ITEM=1, NIGHTWEEK=D(주간)
        $ops = @(
            "ins|ZCRM_CICSVC_NEW|2026092990011|$NEW|$C;$D33;$S1;NIGHTWEEK=D",
            "ins|ZCRM_CICSVC_NEW|2026092990012|$NEW|$C;$D33;$S2;NIGHTWEEK=D",
            "ins|ZCRM_CICSVC_NEW|2026092990013|$NEW|$C;$D33;$S3;NIGHTWEEK=D"
        )
        cscript //nologo scripts\S19_se16n_insert.vbs e3_new @ops
    }
    'req' {
        $ops = 1..3 | ForEach-Object {
            $cs = "20260929900$(10 + $_)"
            "ins|ZCRM_REQ|$cs|$REQ|REQNO=$cs;REQNO_ITEM=10;PRDREQDATE=2026.09.29;REQ_PERSON=ZTEST_20260929E;TAGNO1=99$(10 + $_);ERNAM=ZTEST0929E;ERDAT=2026.09.29"
        }
        cscript //nologo scripts\S19_se16n_insert.vbs e3_req @ops
    }
    'm' {
        $ops = @(
            "ins|ZCRM_CICCOST_M|2026092990011|$TM|$CM;TEXTINFO=ZTEST_20260929E e2e-subs-3y",
            "ins|ZCRM_CICCOST_M|2026092990012|$TM|$CM;TEXTINFO=ZTEST_20260929E e2e-subs-2y",
            "ins|ZCRM_CICCOST_M|2026092990013|$TM|$CM;TEXTINFO=ZTEST_20260929E e2e-general"
        )
        cscript //nologo scripts\S19_se16n_insert.vbs e3_m @ops
    }
    'sub' {
        $ops = @('2026092990011', '2026092990012', '2026092990013') | ForEach-Object {
            "ins|ZCRM_CICCOST_SUB|$_|$TSUB|ERNAM=ZTEST0929E;ERDAT=2026.09.29"
            "ins|ZCRM_CICCOST_ADD|$_|$TADD|NAME1=ZTEST_20260929E;CITY1=ZTEST;ERNAM=ZTEST0929E;ERDAT=2026.09.29"
        }
        cscript //nologo scripts\S19_se16n_insert.vbs e3_sub @ops
    }
}
"exit=$LASTEXITCODE"

# 'cnfrm' : 구독(058S) 전표에 필요한 CIC 대리점확정일. 운영에서는 CIC 가 RFC(ZCRMF03_RFC_ZCRM_CHRG_CNFRM)로 넣는 값.
#            2026-09-29 11:06 사용자가 9/19 E2E 상담건에 SE16N 으로 넣은 것과 같은 방식으로 구독 2건만 넣는다.
if ($Part -eq 'cnfrm') {
    $TCF = 'results\raw\S12_ZCRM_CHRG_CNFRM_e6.tsv'
    $ops = @('2026092990011', '2026092990012') | ForEach-Object { "ins|ZCRM_CHRG_CNFRM|$_|$TCF|SPMON=2026.09;CFDATE1=2026.09.29" }
    cscript //nologo scripts\S19_se16n_insert.vbs e3_cnfrm @ops
    "exit=$LASTEXITCODE"
}
