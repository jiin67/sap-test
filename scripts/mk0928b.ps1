# mk0928b.ps1 : ZCRMR171 구독 문구(조회=실제 년수 / 정산=구독무상 고정) 재테스트용 새 대리점 데이터
#   상담번호 2026092891001~005 / 대리점 100068 경동성남동부대리점(정상점, 센터 CT13, 공급업체 123068, 이전 테스트 미사용)
#   구독기간(보증만료일 - 설치일)은 12개월 단위
#   -Part svc | req
param([Parameter(Mandatory = $true)][ValidateSet('svc', 'req')][string]$Part)
Set-Location C:\sap-test
[Console]::OutputEncoding = [Text.Encoding]::GetEncoding(949)

$SVCG = 'results\raw\S12_ZCRM_CICSVC_tpl_svc_gen.tsv'
$SVCS = 'results\raw\S12_ZCRM_CICSVC_tpl_svc_sub.tsv'
$REQ = 'results\raw\S12_ZCRM_REQ_tpl_req.tsv'

$C = 'RSLTDATE=2026.09.28;REGDATE=2026.09.28;PRDDATE_NEW=2026.09.28;NAME1=ZTEST_20260928;CITY1=ZTEST_20260928;ERNAM=ZTEST0928;AENAM=ZTEST0928;ERDAT=2026.09.28;AEDAT=2026.09.28'
$D68 = 'KUNNR_SVC=100068;SVCNAME=ZTEST_성남동부대리점;VKBUR=CT13'

switch ($Part) {
    'svc' {
        $ops = @(
            # 1) 구독 N년이후, 2023.09.01 ~ 2026.09.01 = 36개월 → '구독무상(3년이후)', 정상 부품 → 저장 (ZPA2 71,000)
            "ins|ZCRM_CICSVC|2026092891001|$SVCS|$C;$D68;TAGNO=9811;SUBS_TYPE=S;SUBSWRTMONTH=12;ISTDATE=2023.09.01;WRTDATE=2026.09.01;PRDDATE_PRT=2025.12.13",
            # 2) 구독 N년이내, 2024.09.01 ~ 2026.09.01 = 24개월 → '구독무상(2년이내)', 정상 부품 → 저장 (PR00 39,500)
            "ins|ZCRM_CICSVC|2026092891002|$SVCS|$C;$D68;TAGNO=9812;SUBS_TYPE=S;SUBSWRTMONTH=0;ISTDATE=2024.09.01;WRTDATE=2026.09.01;PRDDATE_PRT=2025.12.13",
            # 3) 구독 N년이후, 2021.09.01 ~ 2026.09.01 = 60개월 → '구독무상(5년이후)', 정상 부품 → 저장 (ZPA2 71,000)
            "ins|ZCRM_CICSVC|2026092891003|$SVCS|$C;$D68;TAGNO=9813;SUBS_TYPE=S;SUBSWRTMONTH=24;ISTDATE=2021.09.01;WRTDATE=2026.09.01;PRDDATE_PRT=2025.12.13",
            # 4) 구독 N년이내, 2025.09.01 ~ 2026.09.01 = 12개월 → '구독무상(1년이내)', 부품제조일 과거 → Z16 (저장 안 함)
            "ins|ZCRM_CICSVC|2026092891004|$SVCS|$C;$D68;TAGNO=9814;SUBS_TYPE=S;SUBSWRTMONTH=0;ISTDATE=2025.09.01;WRTDATE=2026.09.01;PRDDATE_PRT=2020.01.01",
            # 5) 비구독 → ZCRM052 저장 (PR00 39,500), 정산 '무상'
            "ins|ZCRM_CICSVC|2026092891005|$SVCG|$C;$D68;TAGNO=9815;SUBS_TYPE=;SUBSWRTMONTH=0;WRTDATE=;PRDDATE_PRT=2025.12.13"
        )
        cscript //nologo scripts\S19_se16n_insert.vbs d0928b_svc @ops
    }
    'req' {
        $ops = 1..5 | ForEach-Object {
            $cs = "202609289100$_"
            "ins|ZCRM_REQ|$cs|$REQ|REQNO=$cs;REQNO_ITEM=10;PRDREQDATE=2026.09.28;REQ_PERSON=ZTEST_20260928;TAGNO1=981$_;ERNAM=ZTEST0928;ERDAT=2026.09.28"
        }
        cscript //nologo scripts\S19_se16n_insert.vbs d0928b_req @ops
    }
}
"exit=$LASTEXITCODE"
