import json
from peer import load,facts,CIK
END='2026-06-30'
def ttm(F,tags):
    for tag in tags:
        a=facts(F,tag)
        if not a: continue
        d={(x['start'],x['end']):x['val'] for x in a if 'start' in x}
        fy=('2025-01-01','2025-12-31'); y=('2026-01-01',END); py=('2025-01-01','2025-06-30')
        if fy in d and y in d and py in d: return d[fy]+d[y]-d[py]
    return None
def inst(F,tags,end=END):
    for tag in tags:
        a=facts(F,tag)
        if not a: continue
        v=[x['val'] for x in a if 'start' not in x and x['end']==end]
        if v: return v[-1]
    return None
out={}
for t in CIK:
    F=load(t)
    r=dict(
    rev=ttm(F,['Revenues','RevenueFromContractWithCustomerExcludingAssessedTax']),
    opinc=ttm(F,['OperatingIncomeLoss']), ni=ttm(F,['NetIncomeLoss']),
    da=ttm(F,['DepreciationDepletionAndAmortization','DepreciationAndAmortization']),
    rent=ttm(F,['AircraftRental']), cfo=ttm(F,['NetCashProvidedByUsedInOperatingActivities']),
    capex=ttm(F,['PaymentsToAcquireProductiveAssets']) if t!='DAL' else None,
    debt_c=inst(F,['LongTermDebtAndCapitalLeaseObligationsCurrent']), debt_lt=inst(F,['LongTermDebtAndCapitalLeaseObligations']),
    ol_c=inst(F,['OperatingLeaseLiabilityCurrent']), ol_nc=inst(F,['OperatingLeaseLiabilityNoncurrent']),
    pension=inst(F,['PensionAndOtherPostretirementDefinedBenefitPlansLiabilitiesNoncurrent','DefinedBenefitPensionPlanLiabilitiesNoncurrent']),
    cash=inst(F,['CashAndCashEquivalentsAtCarryingValue','Cash']),
    sti=inst(F,['ShortTermInvestments','AvailableForSaleSecuritiesDebtSecuritiesCurrent','MarketableSecuritiesCurrent']),
    equity=inst(F,['StockholdersEquity']))
    if t=='DAL':
        # capex = flight equipment + other productive assets (tags changed in 2026)
        fl=facts(F,'PaymentsForFlightEquipment'); ot=facts(F,'PaymentsToAcquireOtherProductiveAssets')
        g=lambda a,s,e:[x['val'] for x in a if x.get('start')==s and x['end']==e][0]
        r['capex']=g(facts(F,'PaymentsToAcquireProductiveAssets'),'2025-01-01','2025-12-31')+g(fl,'2026-01-01',END)+g(ot,'2026-01-01',END)-g(facts(F,'PaymentsToAcquireProductiveAssets'),'2025-01-01','2025-06-30')
    out[t]={k:(v/1e6 if v is not None else None) for k,v in r.items()}
    print(t,{k:(round(v) if v is not None else None) for k,v in out[t].items()})
json.dump(out,open('data/peers_ttm_q2_2026.json','w'),indent=1)
