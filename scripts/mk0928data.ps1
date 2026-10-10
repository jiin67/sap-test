# mk0928data.ps1 : 2026-09-28 추가 개선건(Z16 / 실제 년수 / 구독 코스트센터 240074) 테스트 데이터
#   상담번호 2026092890001~003 / 대리점 100050(정상점, 2026.09 일반 정산은 이미 전표 3100069940) / 센터 CT03 (0520 아님)
#   -Part svc | req
param([Parameter(Mandatory = $true)][ValidateSet('svc', 'req')][string]$Part)
Set-Location C:\sap-test
[Console]::OutputEncoding = [Text.Encoding]::GetEncoding(949)

$SVCG = 'results\raw\S12_ZCRM_CICSVC_tpl_svc_gen.tsv'
$SVCS = 'results\raw\S12_ZCRM_CICSVC_tpl_svc_sub.tsv'
$REQ = 'results\raw\S12_ZCRM_REQ_tpl_req.tsv'

$C = 'RSLTDATE=2026.09.28;REGDATE=2026.09.28;PRDDATE_NEW=2026.09.28;NAME1=ZTEST_20260928;CITY1=ZTEST_20260928;ERNAM=ZTEST0928;AENAM=ZTEST0928;ERDAT=2026.09.28;AEDAT=2026.09.28'
$D50 = 'KUNNR_SVC=100050;SVCNAME=ZTEST_계양대리점;VKBUR=CT03'

switch ($Part) {
    'svc' {
        $ops = @(
            # 1) 구독 N년이후, 설치 2023.09.01 ~ 보증 2026.09.01 = 36개월 → '구독무상(3년이후)', 제조일 정상 → 저장 → 171S(CT03) 정산·전표
            "ins|ZCRM_CICSVC|2026092890001|$SVCS|$C;$D50;TAGNO=9801;SUBS_TYPE=S;SUBSWRTMONTH=12;ISTDATE=2023.09.01;WRTDATE=2026.09.01;PRDDATE_PRT=2025.12.13",
            # 2) 구독 N년이내, 설치 2024.03.01 ~ 보증 2026.09.01 = 30개월 → '구독무상(2.5년이내)', 부품제조일 과거 → Z16 (저장 안 함)
            "ins|ZCRM_CICSVC|2026092890002|$SVCS|$C;$D50;TAGNO=9802;SUBS_TYPE=S;SUBSWRTMONTH=0;ISTDATE=2024.03.01;WRTDATE=2026.09.01;PRDDATE_PRT=2020.01.01",
            # 3) 비구독, 부품제조일 과거 → ZCRM052 에서 Z16 이 붙지 않아야 함 (저장 안 함)
            "ins|ZCRM_CICSVC|2026092890003|$SVCG|$C;$D50;TAGNO=9803;SUBS_TYPE=;SUBSWRTMONTH=0;WRTDATE=;PRDDATE_PRT=2020.01.01"
        )
        cscript //nologo scripts\S19_se16n_insert.vbs d0928_svc @ops
    }
    'req' {
        $ops = 1..3 | ForEach-Object {
            $cs = "202609289000$_"
            "ins|ZCRM_REQ|$cs|$REQ|REQNO=$cs;REQNO_ITEM=10;PRDREQDATE=2026.09.28;REQ_PERSON=ZTEST_20260928;TAGNO1=980$_;ERNAM=ZTEST0928;ERDAT=2026.09.28"
        }
        cscript //nologo scripts\S19_se16n_insert.vbs d0928_req @ops
    }
}
"exit=$LASTEXITCODE"
