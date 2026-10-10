# mk0921data.ps1 : 2026-09-21 추가 개선건 테스트 데이터
#   상담번호 2026092190001~004 / 센터 0520(구독사업팀) / 생성자 ZTEST0921
#   -Part svc | new | req
param([Parameter(Mandatory = $true)][ValidateSet('svc', 'new', 'req')][string]$Part)
Set-Location C:\sap-test
[Console]::OutputEncoding = [Text.Encoding]::GetEncoding(949)

$SVCG = 'results\raw\S12_ZCRM_CICSVC_tpl_svc_gen.tsv'
$SVCS = 'results\raw\S12_ZCRM_CICSVC_tpl_svc_sub.tsv'
$NEW = 'results\raw\S12_ZCRM_CICSVC_NEW_tpl_svcnew.tsv'
$REQ = 'results\raw\S12_ZCRM_REQ_tpl_req.tsv'

$C = 'RSLTDATE=2026.09.21;REGDATE=2026.09.21;PRDDATE_NEW=2026.09.21;NAME1=ZTEST_20260921;CITY1=ZTEST_20260921;ERNAM=ZTEST0921;AENAM=ZTEST0921;ERDAT=2026.09.21;AEDAT=2026.09.21'
$D34 = 'KUNNR_SVC=100034;SVCNAME=ZTEST_종로대리점'      # 구독 테스트 (2026.09 RECPTF 행 없음)
$D50 = 'KUNNR_SVC=100050;SVCNAME=ZTEST_계양대리점'      # 비구독 회귀

switch ($Part) {
    'svc' {
        $ops = @(
            # 1) 구독 N년이후 + 센터 0520 + 제조일 정상 → NETWR 이 소비자가(ZPA2 71,000)인지
            "ins|ZCRM_CICSVC|2026092190001|$SVCS|$C;$D34;VKBUR=0520;TAGNO=9601;SUBS_TYPE=S;SUBSWRTMONTH=12;WRTDATE=2027.09.21;PRDDATE_PRT=2025.12.13",
            # 2) 구독 N년이후 + 센터 0520 + 부품제조일이 회수가능제조일보다 과거 → Z02 자동 부적합
            "ins|ZCRM_CICSVC|2026092190002|$SVCS|$C;$D34;VKBUR=0520;TAGNO=9602;SUBS_TYPE=S;SUBSWRTMONTH=12;WRTDATE=2027.09.21;PRDDATE_PRT=2020.01.01",
            # 3) 비구독 회귀 (기존 센터)
            "ins|ZCRM_CICSVC|2026092190003|$SVCG|$C;$D50;VKBUR=CT03;TAGNO=9603;SUBS_TYPE=;SUBSWRTMONTH=0;WRTDATE=;PRDDATE_PRT=2025.12.13",
            # 4) 구독 N년이내(0) + 센터 0520 → N년이후와 금액 비교용
            "ins|ZCRM_CICSVC|2026092190004|$SVCS|$C;$D34;VKBUR=0520;TAGNO=9604;SUBS_TYPE=S;SUBSWRTMONTH=0;WRTDATE=2026.09.21;PRDDATE_PRT=2025.12.13"
        )
        cscript //nologo scripts\S19_se16n_insert.vbs d0921_svc @ops
    }
    'new' {
        $ops = @(
            "ins|ZCRM_CICSVC_NEW|2026092190001|$NEW|$C;$D34;VKBUR=0520;TAGNO=9601;SUBS_TYPE=S;SUBSWRTMONTH=12;WRTDATE=2027.09.21;NIGHTWEEK=D",
            "ins|ZCRM_CICSVC_NEW|2026092190002|$NEW|$C;$D34;VKBUR=0520;TAGNO=9602;SUBS_TYPE=S;SUBSWRTMONTH=12;WRTDATE=2027.09.21;NIGHTWEEK=D",
            "ins|ZCRM_CICSVC_NEW|2026092190003|$NEW|$C;$D50;VKBUR=CT03;TAGNO=9603;SUBS_TYPE=;SUBSWRTMONTH=0;WRTDATE=;NIGHTWEEK=D",
            "ins|ZCRM_CICSVC_NEW|2026092190004|$NEW|$C;$D34;VKBUR=0520;TAGNO=9604;SUBS_TYPE=S;SUBSWRTMONTH=0;WRTDATE=2026.09.21;NIGHTWEEK=D"
        )
        cscript //nologo scripts\S19_se16n_insert.vbs d0921_new @ops
    }
    'req' {
        $ops = 1..4 | ForEach-Object {
            $cs = "202609219000$_"
            "ins|ZCRM_REQ|$cs|$REQ|REQNO=$cs;REQNO_ITEM=10;PRDREQDATE=2026.09.21;REQ_PERSON=ZTEST_20260921;TAGNO1=960$_;ERNAM=ZTEST0921;ERDAT=2026.09.21"
        }
        cscript //nologo scripts\S19_se16n_insert.vbs d0921_req @ops
    }
}
"exit=$LASTEXITCODE"
