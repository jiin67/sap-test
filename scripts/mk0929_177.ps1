# mk0929_177.ps1 : ZCRM177A / ZCRM177AS 문서처리(주문→납품→출고) 확인용 테스트 상담건 (2026-09-29)
#   2026092990001 구독 무상 / 2026092990002 일반 무상 - 센터 CT08 (1120/CT08 에 30031356A 재고 4EA)
#   2026091990008(CT04) 은 CT04 재고 0 이라 무상오더 4293368 에서 납품이 막혀 새로 만든다.
#   -Part svc | m | sub      (svc = ZCRM_CICSVC, m = ZCRM_CICCOST_M, sub = ZCRM_CICCOST_SUB + ZCRM_CICCOST_ADD)
param([Parameter(Mandatory = $true)][ValidateSet('svc', 'm', 'sub')][string]$Part)
Set-Location C:\sap-test
[Console]::OutputEncoding = [Text.Encoding]::GetEncoding(949)

$SVCG = 'results\raw\S12_ZCRM_CICSVC_tpl_svc_gen.tsv'
$SVCS = 'results\raw\S12_ZCRM_CICSVC_tpl_svc_sub.tsv'
$TM = 'results\raw\S12_ZCRM_CICCOST_M_tpl_m.tsv'
$TSUB = 'results\raw\S12_ZCRM_CICCOST_SUB_tpl_sub.tsv'
$TADD = 'results\raw\S12_ZCRM_CICCOST_ADD_tpl_add.tsv'

$C = 'NAME1=ZTEST_20260929;CITY1=ZTEST_20260929;RSLTDATE=2026.09.29;REGDATE=2026.09.29;PRDDATE_NEW=2026.09.29;ERNAM=ZTEST0929;AENAM=ZTEST0929;ERDAT=2026.09.29;AEDAT=2026.09.29;VKBUR=CT08'
$CM = 'CENTER_CD=CT08;CENTER_NM=ZTEST;TECHNAME=ZTEST;ERNAM=ZTEST0929;ERDAT=2026.09.29;INDATE=2026.09.29;ISFREE=Y'

switch ($Part) {
    'svc' {
        $ops = @(
            # 구독 N년이후, 설치 2023.09.01 ~ 보증 2026.09.01 = 36개월 (3년)
            "ins|ZCRM_CICSVC|2026092990001|$SVCS|$C;TAGNO=9901;SUBS_TYPE=S;SUBSWRTMONTH=12;ISTDATE=2023.09.01;WRTDATE=2026.09.01",
            # 일반
            "ins|ZCRM_CICSVC|2026092990002|$SVCG|$C;TAGNO=9902;SUBS_TYPE=;SUBSWRTMONTH=0;WRTDATE="
        )
        cscript //nologo scripts\S19_se16n_insert.vbs d0929_177svc @ops
    }
    'm' {
        $ops = @(
            "ins|ZCRM_CICCOST_M|2026092990001|$TM|$CM;TEXTINFO=ZTEST_20260929 177-doc-subs-CT08",
            "ins|ZCRM_CICCOST_M|2026092990002|$TM|$CM;TEXTINFO=ZTEST_20260929 177-doc-general-CT08"
        )
        cscript //nologo scripts\S19_se16n_insert.vbs d0929_177m @ops
    }
    'sub' {
        $ops = @('2026092990001', '2026092990002') | ForEach-Object {
            "ins|ZCRM_CICCOST_SUB|$_|$TSUB|ERNAM=ZTEST0929;ERDAT=2026.09.29"
            "ins|ZCRM_CICCOST_ADD|$_|$TADD|NAME1=ZTEST_20260929;CITY1=ZTEST;ERNAM=ZTEST0929;ERDAT=2026.09.29"
        }
        cscript //nologo scripts\S19_se16n_insert.vbs d0929_177sub @ops
    }
}
"exit=$LASTEXITCODE"
