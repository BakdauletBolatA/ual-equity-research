import json,sys
from datetime import date
CIK={'UAL':'0000100517','DAL':'0000027904','AAL':'0000006201','LUV':'0000092380','ALK':'0000766421','JBLU':'0001158463'}
def load(t): return json.load(open(f'data/facts_{CIK[t]}.json'))['facts']
def facts(F,tag):
    for ns in ('us-gaap','dei'):
        if tag in F.get(ns,{}):
            u=F[ns][tag]['units']; return u.get('USD') or u.get('shares') or list(u.values())[0]
    return None
def ttm(F,tags):
    for tag in tags:
        a=facts(F,tag)
        if not a: continue
        d={}
        for x in a:
            if 'start' in x: d[(x['start'],x['end'])]=x['val']
        fy=[k for k in d if k[0]=='2025-01-01' and k[1]=='2025-12-31']
        ytd=sorted([k for k in d if k[0]=='2026-01-01'],key=lambda k:k[1])
        if not fy or not ytd: continue
        e=ytd[-1]; pe=('2025-01-01',str(int(e[1][:4])-1)+e[1][4:])
        if pe not in d: continue
        return d[fy[0]]+d[e]-d[pe], tag, e[1]
    return None,None,None
def inst(F,tags):
    for tag in tags:
        a=facts(F,tag)
        if not a: continue
        b=sorted([x for x in a if 'start' not in x],key=lambda x:x['end'])
        if b and b[-1]['end']>='2026-06-30': return b[-1]['val'],tag,b[-1]['end']
    return None,None,None
if __name__=='__main__':
    t=sys.argv[1]; F=load(t)
    R={
    'rev':ttm(F,['Revenues','RevenueFromContractWithCustomerExcludingAssessedTax','OperatingRevenue']),
    'opinc':ttm(F,['OperatingIncomeLoss']),
    'ni':ttm(F,['NetIncomeLoss']),
    'da':ttm(F,['DepreciationDepletionAndAmortization','DepreciationAndAmortization','DepreciationAmortizationAndAccretionNet','Depreciation']),
    'rent':ttm(F,['AircraftRental','OperatingLeaseExpense','AircraftRent']),
    'olcost':ttm(F,['OperatingLeaseCost','OperatingLeasePayments']),
    'cfo':ttm(F,['NetCashProvidedByUsedInOperatingActivities']),
    'capex':ttm(F,['PaymentsToAcquireProductiveAssets','PaymentsToAcquirePropertyPlantAndEquipment','PaymentsForFlightEquipment']),
    'fuel':ttm(F,['FuelCosts','AircraftFuelAndRelatedTaxes','AircraftFuel']),
    'interest':ttm(F,['InterestExpenseNonoperating','InterestExpense','InterestExpenseDebt']),
    'debt':inst(F,['LongTermDebtAndCapitalLeaseObligationsIncludingCurrentMaturities','LongTermDebt','DebtAndCapitalLeaseObligations','LongTermDebtAndFinanceLeaseObligations']),
    'debt_cur':inst(F,['LongTermDebtAndCapitalLeaseObligationsCurrent','LongTermDebtCurrent','DebtCurrent']),
    'debt_lt':inst(F,['LongTermDebtAndCapitalLeaseObligations','LongTermDebtNoncurrent']),
    'finlease':inst(F,['FinanceLeaseLiability']),
    'oplease':inst(F,['OperatingLeaseLiability']),
    'oplease_c':inst(F,['OperatingLeaseLiabilityCurrent']),
    'oplease_nc':inst(F,['OperatingLeaseLiabilityNoncurrent']),
    'cash':inst(F,['CashAndCashEquivalentsAtCarryingValue']),
    'sti':inst(F,['ShortTermInvestments','AvailableForSaleSecuritiesDebtSecuritiesCurrent','MarketableSecuritiesCurrent']),
    'equity':inst(F,['StockholdersEquity']),
    'shares':inst(F,['EntityCommonStockSharesOutstanding']),
    'pension':inst(F,['DefinedBenefitPensionPlanLiabilitiesNoncurrent','PensionAndOtherPostretirementDefinedBenefitPlansLiabilitiesNoncurrent']),
    }
    print(t, json.dumps({k:v for k,v in R.items()}))
