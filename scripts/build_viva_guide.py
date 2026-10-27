"""Build a compact, source-traced ten-page viva guide without editing Excel."""
import argparse, hashlib, html, json, pathlib, re
import openpyxl
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, PageBreak, Table, TableStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib import colors
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.pagesizes import A4
from pypdf import PdfReader

ap=argparse.ArgumentParser();ap.add_argument('--workbook',default='outputs/overview-short/KOHC_IBF_FINAL_DECLUTTERED.xlsx');ap.add_argument('--output',default='output/pdf/KOHC_Viva_Guide_Updated_10_Pages.pdf');ap.add_argument('--markdown',default='work/github_delivery/Viva_Guide.md');args=ap.parse_args()
src=pathlib.Path(args.workbook);before=hashlib.sha256(src.read_bytes()).hexdigest();w=openpyxl.load_workbook(src,data_only=True)
cos=['KOHC','LUCK','DGKC','MLCF','FCCL'];ratio_rows=[111,112,113,114,116,117,118,119,120,121,123,124,125,126,128,129,130,131,132,133,134,136,137,138,139,140,141]
names=['Current Ratio','Quick Ratio','Cash Ratio','Net Working Capital','Inventory Turnover','Receivables Turnover','Days Inventory Outstanding','Days Sales Outstanding','Fixed Asset Turnover','Total Asset Turnover','Debt Ratio','Debt-to-Equity','Debt-to-Capital','Times Interest Earned','Gross Profit Margin','Operating / EBIT Margin','Net Profit Margin','Return on Assets','Return on Equity','Return on Invested Capital','Basic Earning Power','Earnings per Share','Book Value per Share','Price-to-Earnings','Market-to-Book','Dividend Yield','EV/EBITDA']
forms=['CA / CL','(CA - inventory) / CL','Available cash equivalents / CL','CA - CL','COGS / average inventory','Net sales / average net receivables','365 / inventory turnover','365 x average net receivables / sales','Sales / closing net PPE including CWIP','Sales / closing total assets','Total liabilities / total assets','IBD / closing common equity','IBD / (IBD + common equity)','EBIT / expensed interest','Gross profit / sales','EBIT / sales','PAT / sales','PAT / closing total assets','PAT / closing common equity','EBIT x (1 - actual ETR) / (IBD + equity)','EBIT / closing total assets','PAT / compatible weighted shares','Common equity / closing ordinary shares','Official PSX close / compatible EPS','Official PSX close / compatible BVPS','FY cash DPS / official PSX close','(Market equity + IBD + other claims - cash) / EBITDA']
pct_ids={11,13,15,16,17,18,19,20,21,26};day_ids={7,8};amount_ids={4,22,23}
def val(co,y,row):return w[co].cell(row,y-2017).value
def fmt(i,x):
 if not isinstance(x,(int,float)):return str(x)
 return f'{x*100:.2f}%' if i in pct_ids else f'{x:,.2f}'+(' days' if i in day_ids else '' if i in amount_ids else 'x')
fontdir=pathlib.Path('C:/Windows/Fonts')
if fontdir.exists():font_files=[('Body',fontdir/'calibri.ttf'),('Bold',fontdir/'calibrib.ttf')]
else:
 import reportlab
 fontdir=pathlib.Path(reportlab.__file__).parent/'fonts';font_files=[('Body',fontdir/'Vera.ttf'),('Bold',fontdir/'VeraBd.ttf')]
for n,f in font_files:pdfmetrics.registerFont(TTFont(n,str(f)))
pdfmetrics.registerFontFamily('Body',normal='Body',bold='Bold',italic='Body',boldItalic='Bold')
green=colors.HexColor('#204F3D');ink=colors.HexColor('#1C3028');gray=colors.HexColor('#526359');pale=colors.HexColor('#EDF3EC')
styles={
 'body':ParagraphStyle('body',fontName='Body',fontSize=10.1,leading=13.1,textColor=ink,spaceAfter=6),
 'small':ParagraphStyle('small',fontName='Body',fontSize=8.4,leading=10.5,textColor=gray,spaceAfter=5),
 'title':ParagraphStyle('title',fontName='Bold',fontSize=19,leading=22,textColor=green,spaceAfter=12),
 'h':ParagraphStyle('h',fontName='Bold',fontSize=11.5,leading=14,textColor=green,spaceBefore=8,spaceAfter=5,keepWithNext=True),
 'cell':ParagraphStyle('cell',fontName='Body',fontSize=8.8,leading=10.8,textColor=ink),
 'tiny':ParagraphStyle('tiny',fontName='Body',fontSize=8.0,leading=9.5,textColor=ink),
 'head':ParagraphStyle('head',fontName='Bold',fontSize=8.5,leading=10,textColor=colors.white)}
story=[];md=[];titles=[]
def clean(t):return re.sub(r'(?<=[a-z])(?=\d)',' ',str(t)).replace('1July','1 July').replace('28April','28 April')
def esc(t):return html.escape(clean(t)).replace('\n','<br/>')
def para(t,kind='body'):story.append(Paragraph(esc(t),styles[kind]));md.append(clean(t)+'\n')
def head(t):story.append(Paragraph(esc(t),styles['h']));md.append('### '+clean(t)+'\n')
def page(t):
 if story:story.append(PageBreak())
 titles.append(t);story.append(Paragraph(esc(t),styles['title']));md.append('## '+t+'\n')
def tab(rows,widths,tiny=False,pad=4):
 cells=[[Paragraph(esc(c),styles['head'] if i==0 else styles['tiny' if tiny else 'cell'])for c in row]for i,row in enumerate(rows)]
 t=Table(cells,colWidths=widths,repeatRows=1,hAlign='LEFT');cmd=[('BACKGROUND',(0,0),(-1,0),green),('VALIGN',(0,0),(-1,-1),'TOP'),('LEFTPADDING',(0,0),(-1,-1),5),('RIGHTPADDING',(0,0),(-1,-1),5),('TOPPADDING',(0,0),(-1,-1),pad),('BOTTOMPADDING',(0,0),(-1,-1),pad)]
 for i in range(1,len(rows)):
  if i%2:cmd.append(('BACKGROUND',(0,i),(-1,i),pale))
  cmd.append(('LINEBELOW',(0,i),(-1,i),.25,colors.HexColor('#D4DFD7')))
 t.setStyle(TableStyle(cmd));story.extend([t,Spacer(1,7)]);md.append('| '+' | '.join(str(x).replace('\n',' ')for x in rows[0])+' |');md.append('| '+' | '.join('---'for x in rows[0])+' |');md.extend('| '+' | '.join(str(x).replace('\n',' ')for x in r)+' |'for r in rows[1:]);md.append('')
def qa(q,a,trace=None):
 head(q);para(a)
 if trace:para('Trace: '+trace,'small')
def source(t):para('Evidence: '+t,'small')

page('1. Companies, competition and your opening answer')
para('KOHC Financial Analysis | FY2021-FY2025 | Updated viva guide','small')
para('We compare Kohat Cement with Lucky Cement, D.G. Khan Cement, Maple Leaf Cement and Fauji Cement using audited unconsolidated accounts, official PSX year-end closing prices and one standardized definition per measure. The workbook retains 27 measures, including all 19 mandatory ratios, 675 company-year results, 135 sample means and 54 charts.')
tab([['Company','Business and the point to remember'],['KOHC - Kohat Cement','Main company; grey and white cement. Strong latest margins and recorded liquidity; slower inventory movement and no applicable FY dividends need discussion.'],['LUCK - Lucky Cement','Cement and clinker producer/exporter with Karachi and Pezu operations. Investments influence standalone earnings. FY2025 has a five-for-one share split.'],['DGKC - D.G. Khan Cement','Ordinary Portland and sulphate-resistant cement. A FY2023 tax-related loss is followed by a sharp earnings recovery.'],['MLCF - Maple Leaf Cement','Grey and white cement; Kohinoor Maple Leaf Group. Large stores/spares balances and FY2025 subsidiary dividend/investment income affect interpretation.'],['FCCL - Fauji Cement','Fauji/Askari cement brands and specialist products. The Askari merger changes the business perimeter between FY2021 and FY2022.']],[110,409])
head('The sector context is mixed')
para('KOHC\'s FY2025 industry review reports domestic cement dispatches down 2.7%, exports up 29.5% and total dispatches up 2.3%. A company can improve profit through cost control or financing relief while local demand remains weak. Do not describe the whole period as a simple demand boom.')
head('How to explain a number in the viva')
para('State the ratio and period; identify which numerator and denominator changed; trace the statement line and note; cross-check a related ratio; then explain the implication and a limitation. A supported numerical driver is stronger evidence than an unsupported general story.')
para('Amounts in the case studies are PKR billion unless stated; formula-sheet NWC is PKR million and per-share amounts are rupees. D-H are FY2021-FY2025. Hidden column I holds FY2020 openings: inventory I54 and receivables I20.','small')
para('Group: '+str(w['Overview']['B7'].value).replace('\n','; '),'small')
source('Latest submission Overview; official company profiles/products; KOHC FY2025 directors\' report, printed pp24-25. Prepared 27 October 2026. No workbook calculation was changed.')

page('2. One consolidated formula sheet')
para('CA = current assets; CL = current liabilities; IBD = interest-bearing debt principal plus leases; PAT = profit after levy/tax. Averages are (opening + closing)/2. Ratios retain full precision in Excel.','small')
tab([['Measure','Standardized formula / basis','KOHC FY2025']]+[[f'{i}. {names[i-1]}',forms[i-1],fmt(i,val('KOHC',2025,ratio_rows[i-1]))]for i in range(1,28)],[142,292,85],tiny=True,pad=3)
para('EBIT = pre-levy/pre-tax profit + expensed interest, including other income. ETR = (levy + income-tax expense) / pre-levy/pre-tax profit, uncapped. EBITDA = EBIT + audited depreciation/amortization. Available audited cash equivalents are used; book financing is the EV debt-value proxy.','small')
para('Qualification: the existing KOHC debt classification retains PKR3.704884m of FY2025 short-term accrued mark-up despite nil principal utilization; D/E, debt-to-capital and ROIC remain qualified. These are preserved workbook figures, not an assertion that this known issue has been corrected.','small')

page('3. All 27 measures: FY2025 comparison')
rows=[['Measure','KOHC','LUCK','DGKC','MLCF','FCCL','Mean']]
for i,row in enumerate(ratio_rows,1):
 values=[val(co,2025,row)for co in cos];numeric=[v for v in values if isinstance(v,(int,float))];mean=sum(numeric)/len(numeric)
 rows.append([names[i-1]]+[fmt(i,v)for v in values]+[fmt(i,mean)])
tab(rows,[149]+[61.6667]*6,tiny=True,pad=3)
head('What this comparison establishes')
para('KOHC leads liquidity cover, latest margins and interest coverage. LUCK leads cash cover and inventory turnover; DGKC has the lowest DSO; MLCF leads total-asset turnover; FCCL has the highest FY2025 Dividend Yield. There is no single winner for every objective.')
para('Mean = equal-weight five-company benchmark including KOHC, not the whole industry. Raw EPS/BVPS levels are not size-neutral quality rankings and LUCK\'s split limits comparisons of denominations. NWC is an absolute amount. Genuine zeros count; n/m does not.','small')

page('4. Methodology and accounting notes to defend')
tab([['Choice','Reason and implication'],['Unconsolidated accounts','Keep the same reporting boundary. Parent accounts include investment assets and subsidiary dividends; consolidated accounts combine subsidiary activity and eliminate intra-group dividends.'],['Balance timing','Inventory and receivables use opening/closing averages. PPE, assets, equity and invested capital use closing balances. Do not silently swap conventions to reproduce an issuer ratio.'],['Inventory / PPE','Inventory includes stores, spares and stock-in-trade. Quick ratio removes that whole inventory amount. PPE includes CWIP, which may not yet generate sales.'],['Debt / interest / cash','Debt Ratio uses all liabilities. D/E uses principal and leases, excluding interest-free financing and accrued interest by policy, subject to the disclosed KOHC implementation issue. TIE uses expensed interest, not total finance cost, capitalized interest or cash paid. Cash uses audited qualifying available components.'],['Actual taxation / ROIC','Use the actual uncapped ETR, including levy. A tax rate over 100% can produce negative NOPAT; do not cap it to create a positive ROIC.'],['PSX / share bases','Last official session on/before June 30. LUCK split trading: 28-Apr-2025; KOHC: 25-Aug-2025, after FY2025. Prices, EPS, BVPS and DPS must use compatible units.'],['n/m / averages / charts','The project display rule marks negative/non-meaningful observations n/m; signed accounting inputs remain. Means exclude n/m but include numeric zeros. Chart gaps for n/m are intentional.']],[111,408],pad=4)
head('Company-specific notes')
para('KOHC: FY2020 has an immaterial one-rupee footing difference. The separate accrued mark-up issue remains disclosed. LUCK: retain the note-supported PKR0.400m deferred-tax correction and PKR10.653m amortization. Its published ROE may use comprehensive income, while ours uses PAT.')
para('DGKC: FY2021-FY2023 NWC deficits are real; FY2023 loss/tax effects are retained. MLCF: freight reclassification can move cost from COGS to distribution, changing gross margin without the same change in operating profit; related operating supplier payables are not automatically debt. FCCL: merger comparatives, interest-free funding and capitalized interest require care.')
para('Audit: 437 of 455 published comparisons are numerically reconstructed; 18 retain qualified differences. Eight concern FCCL asset-turnover inconsistencies; the remaining exact published bases are unconfirmed where stated. Earlier 63-case forensic records remain in the Audit. Do not force agreement with hardcoded figures.','small')

page('5. Liquidity: investigate the inventory gap')
qa('1. Why was MLCF current cover 1.25x but quick cover only 0.29x in FY2022?', 'CA was 20.322, CL 16.193 and inventory 15.549. Current cover = 20.322/16.193 = 1.25x. Removing inventory leaves 4.773 of quick assets: 4.773/16.193 = 0.29x. Inventory represented 76.5% of CA. This is a high reliance on inventory within recorded short-term resources.', 'MLCF E111:E112 -> E24, E48 and E54; E54 = E18 + E19.')
para('The inventory note is essential: 12.854 was stores/spares/tools, 2.696 was stock-in-trade, and only 0.514 was finished goods. Inventory turnover was 2.64x, or 138 days. The gap supports an inventory-dependence explanation, but it does not prove future sales or that all stock was unsold cement.')
qa('2. Does a widening current-quick gap always mean stockpiling?', 'No. DGKC\'s gap widened from 0.63x to 0.79x in FY2024-FY2025 while inventory rose only 1.2% (20.253 -> 20.495). CL fell 19.6% (32.226 -> 25.911). Current cover improved 1.30x -> 1.85x and quick cover 0.67x -> 1.06x. The smaller denominator mainly explains the wider gap.', 'DGKC G:H rows48,54,111,112. CR - QR = inventory / CL.')
qa('3. How can KOHC quick cover be 2.67x but cash cover only 0.12x?', 'KOHC has 26.942 of short-term investments versus 1.451 cash/bank balances. Quick assets are broader than audited qualifying cash equivalents. Do not call all quick assets instantly spendable cash. Conversely, LUCK\'s audited cash equivalents include 61.298 of qualifying investments plus 2.790 bank cash, giving 64.088 and cash cover of 1.09x.', 'Rows21-22 and113; cash workings below the ratio table. LUCK FY2025 note36.2.')
head('A strong oral answer')
para('"I would first reconcile the mathematical gap, then inspect the inventory note to separate fuel/stores, raw materials, work-in-process and finished goods. I would compare turnover and actual dispatches before deciding whether the balance reflects production needs, slow movement or anticipated demand."')
source('MLCF FY2022 notes25-26, printed p231; LUCK FY2025 note36.2, printed p273; latest company sheets. Potential causes without official confirmation remain hypotheses.')

page('6. KOHC: inventory and receivable warning signs')
qa('4. Why did inventory turnover worsen when total inventory fell?', 'FY2024-FY2025 turnover fell 3.60x -> 3.02x and DIO rose 102 -> 121 days. Average inventory fell only 1.0% (7.618 -> 7.542), while cost of sales fell 16.7% (27.391 -> 22.814). The numerator contracted much faster than the denominator.', 'KOHC G:H rows54,65,116,118. The turnover denominator uses the preceding closing balance too.')
tab([['Inventory component','FY2024','FY2025','What it tells us'],['Stores and spares','5.786','4.991','Declined'],['Stock-in-trade','1.794','2.514','Increased about40%'],['Work-in-process within stock','0.723','1.313','Increased about82%'],['Finished goods within stock','0.671','0.776','Increased about16%'],['Total inventory','7.579','7.505','Small fall masks a mix change']],[192,69,69,189])
para('Cement dispatches fell about10%. Higher WIP/finished goods alongside weaker dispatches deserves investigation. However, a lower cost of sales also mechanically lowers turnover; this is not a pure physical-volume measure and does not prove obsolete inventory.')
qa('5. Why did DSO rise although closing receivables declined?', 'Closing net receivables fell 2.110 -> 1.692, but average receivables rose 1.658 -> 1.901. Sales fell 38.648 -> 37.536. DSO therefore rose 15.66 -> 18.49 days. The high FY2024 close becomes the FY2025 opening, so looking only at the latest balance misses the actual denominator.', 'KOHC F20:H20; sales row64; receivables turnover117 and DSO119.')
para('The loss allowance also increased from0.0416 to0.0857. A fall in net receivables cannot automatically be described entirely as cash collection. Read gross receivables, allowance movement and ageing. Net sales is the consistent credit-sales proxy, so DSO is not a verified invoice-by-invoice collection period.','small')
source('KOHC FY2025 notes18-20.1 and operating-performance discussion, printed pp24-25,99-100. Inventory and receivable calculations use unrounded workbook inputs.')

page('7. KOHC: stronger profits despite weaker sales')
qa('6. Why did profit rise when revenue fell?', 'Sales declined 2.9%, but cost of sales declined16.7%. Gross profit rose30.8%; gross margin increased29.13% ->39.22%. PAT rose8.893 ->11.575 and net margin23.01% ->30.84%. The financial explanation is that costs fell faster than revenue.', 'KOHC G:H rows64-79; gross margin128 and net margin130.')
tab([['PKR billion','FY2024','FY2025'],['Sales','38.648','37.536'],['Cost of sales, positive magnitude','27.391','22.814'],['Gross profit','11.256','14.722'],['Other income','4.454','5.281'],['Profit after tax','8.893','11.575']],[321,99,99])
para('Management links the improvement to coal/fuel management and renewable-energy efficiency. Dispatches nevertheless fell10%, so "stronger demand" would be the wrong explanation for KOHC\'s latest profit improvement.')
qa('7. Why is EBIT margin48.10%, above gross margin39.22%?', 'The selected EBIT measure includes recognized other income. Core operating profit before other income is12.816, or34.14% of sales; other income is5.281. The full EBIT margin should not be called a pure cement-manufacturing margin.', 'Rows71-73,82,128-129. EBIT is reconstructed consistently from pre-levy/pre-tax profit plus expensed interest.')
qa('8. Why did interest coverage jump22.32x ->58.96x?', 'EBIT rose25.2% from14.422 to18.053, while expensed interest fell52.6% from0.646 to0.306. Both changes improve TIE. This is strong accounting earnings cover, but it is not coverage of principal repayments or proof of equivalent cash generation.', 'Rows82,85,126. Full finance cost, cash interest paid and expensed interest are different numbers.')
head('Management implication')
para('Preserve the cost gains, investigate slower stock conversion and monitor collection quality. Strong margins do not remove working-capital or investment funding needs. Read the cash-flow statement before recommending distributions.')
source('KOHC FY2025 directors\' report, printed pp24-25; income statement and finance-cost notes. Business explanations are management-supported; no quantified attribution to individual initiatives is invented.')

page('8. MLCF and DGKC: earnings quality and recovery')
qa('9. Why did MLCF profit more than triple on only3.7% sales growth?', 'FY2024-FY2025 PAT rose5.273 ->17.036 (about223%); net margin7.93% ->24.71%. Other income rose0.238 ->8.463. Note39 includes6.660 dividend income, of which6.401 came from Maple Leaf Power, plus1.694 unrealised investment gains. Core operating performance also improved, but the whole profit jump is not cement-sales growth.', 'MLCF G:H rows64,71-73,79,130,132; FY2025 note39.')
para('A subsidiary dividend received in parent accounts differs from a dividend paid to MLCF shareholders. Thus MLCF can record dividend income while its shareholder Dividend Yield remains zero. In consolidated accounts, intra-group dividends are eliminated.')
qa('10. How did DGKC make a FY2023 loss after positive pre-tax profit?', 'Pre-levy/pre-tax profit was3.163, combined levy/tax6.799 and the net loss3.636. The actual ETR was214.97%. The report explains increased deferred-tax liability linked to super tax and expired tax credits; high interest had already reduced pre-tax profit. Do not replace the loss with zero or cap the tax rate.', 'DGKC F76:F79 and F86. Earnings/return n/m displays retain signed underlying figures.')
para('DGKC also had CA39.027 below CL48.676: a real9.649 working-capital deficit. Its liquidity deficit and its tax-driven net loss are distinct issues; do not use one as an automatic explanation of the other.')
qa('11. Why did DGKC P/E collapse while its price rose?', 'FY2024-FY2025 price roseRs90.27 ->Rs165.56, but EPS grewRs1.24 ->Rs19.80. P/E fell72.92x ->8.36x because earnings recovered much faster than price. Gross margin improved15.94% ->25.74%; interest fell7.925 ->3.817 and TIE improved1.36x ->4.41x. The old earnings base was unusually small.', 'DGKC G:H rows85,105,126,128,136,138.')
para('A lower P/E here is not evidence that the price declined, nor proof of undervaluation. Check whether the recovered earnings are sustainable.','small')
source('MLCF FY2025 note39, printed p264; DGKC FY2023 financial-performance review, printed p104; latest standardized statement inputs and ratio tables.')

page('9. FCCL and LUCK: denominator and share effects')
qa('12. Why did FCCL ratios weaken although sales and profit grew?', 'FY2021-FY2022 assets rose34.052 ->113.698 and sales24.271 ->54.243, while current cover fell2.02x ->1.06x, ROA10.19% ->6.26% and TIE56.45x ->12.58x. The Askari amalgamation was effective1July2021. Note1.3 states FY2021 comparatives were not revised and represent the pre-acquisition business.', 'FCCL D:E rows25,55,64,79,85,111,126,131; FY2022 note1.3, printed pp190-191.')
para('This is a change in business perimeter. Growth is not wholly organic, and ratio deterioration does not automatically demonstrate an equivalent collapse in management quality. Assets and financing obligations expanded faster than the earnings needed to maintain the earlier ratios.')
qa('13. Can FCCL D/E fall even when debt rises?', 'Yes. FY2024-FY2025 IBD increased36.590 ->40.036 (9.4%), while equity increased73.399 ->84.272 (14.8%). D/E therefore fell0.50x ->0.48x. Lower leverage as a proportion does not necessarily mean absolute debt repayment.', 'FCCL G:H rows55,29 and124.')
qa('14. Did LUCK EPS collapse fromRs94.54 toRs22.59?', 'No. The five-for-one split changes the share unit. Put FY2024 on the same post-split unit:94.54/5 =Rs18.91. Comparable EPS rose18.91 ->22.59, while PAT rose28.107 ->33.092. BVPS and raw prices also require denomination care. LUCK split trading began28April2025.', 'LUCK G:H rows79,103-105,136-138; share-basis disclosures and Overview.')
qa('15. Why did LUCK ROE slip despite higher profit?', 'PAT rose17.7%, but closing equity grew19.1% from147.761 to175.910. ROE therefore eased19.02% ->18.81%. Returns depend on the earnings and the capital base together.', 'LUCK G:H rows79,29,132. Closing equity is the project convention.')
head('The broader lesson')
para('Always ask whether the business perimeter, asset base, funding mix or number of shares changed before attributing a ratio movement to sales performance. The same arithmetic direction can have very different economic explanations.')
source('FCCL FY2022 note1.3; audited share and equity records; official PSX corporate-action evidence retained in the Audit. No retrospective KOHC FY2025 split adjustment is applied.')

page('10. Dividends, decisions and final viva checks')
tab([['FY2025','DPS (Rs)','PSX close (Rs)','Dividend Yield']]+[[co,f'{val(co,2025,106):.2f}',f'{val(co,2025,105):.2f}',fmt(26,val(co,2025,140))]for co in cos],[126,100,136,157])
qa('16. Can yield fall while the dividend rises?', 'FCCL DPS roseRs1.00 ->Rs1.25, but price roseRs22.91 ->Rs44.67. Yield fell4.36% ->2.80% because price grew faster. Dividend Yield is DPS divided by the matching year-end share price, not the growth rate in the dividend.')
qa('17. Why is KOHC yield zero when the cash flow deducts dividends?', 'The project attributes dividends to financial years, including applicable final proposals. Payments may settle old liabilities. FY2025 cash dividends ofPKR3.158m match the fall in dividend payable from36.835m to33.677m. This does not establish a new FY2025 dividend. Management cites investments and buyback activity in its no-dividend decision.', 'KOHC rows45,98,106,140; FY2025 dividend-payable note and directors\' report.')
para('Verified zero years: KOHC and MLCF FY2021-FY2025; LUCK FY2021-FY2022; DGKC FY2023-FY2024; FCCL FY2021-FY2023. All17 observations remain numeric zero and count in the means. FY2025 mean yield is1.03%, including the two zeros.','small')
head('Decisions you can defend')
para('KOHC: preserve margins and investigate stock/receivables before recommending payout. LUCK: distinguish cement operations, investment income and split effects. DGKC: examine financing relief, tax and the sustainability of recovery. MLCF: separate recurring operations from dividend/fair-value income. FCCL: assess merger comparability, expansion funding and debt service alongside its dividend.')
head('Short answers to common challenges')
para('A current ratio of2 is a classroom reference, not a verified cement-sector target. Quick cover below1 identifies reliance on other funding/assets, not automatic insolvency. High ROE may reflect leverage or a small equity base. ROIC alone does not prove value creation without a compatible required return. Low multiples do not prove a bargain. n/m is not zero. The mean of company DIO is not365 divided by mean turnover.')
para('The known KOHC mark-up classification qualification and18 published differences remain disclosed; this guide does not certify them away. The Excel was not edited to prepare this PDF. The full41-page report and unchanged Audit retain detailed methodology and reconciliation evidence.','small')
source('Official audited reports and PSX evidence retained in the Audit. Report portals: kohatcement.com/financial_reports.aspx; lucky-cement.com; dgcement.com; kmlg.com/mlcfl/financials; fccl.com.pk/eng/financial-reports. Detailed case references appear on pp5-9.')

class Doc(SimpleDocTemplate):
 def afterFlowable(self,f):
  if isinstance(f,Paragraph) and f.style.name=='title':
   key='p'+str(self.page);self.canv.bookmarkPage(key);self.canv.addOutlineEntry(f.getPlainText(),key,0,False)
def footer(c,d):
 c.saveState();c.setStrokeColor(colors.HexColor('#D4DFD7'));c.line(38,40,A4[0]-38,40);c.setFont('Body',8);c.setFillColor(gray);c.drawString(38,27,'KOHC | Viva preparation | FY2021-FY2025');c.drawRightString(A4[0]-38,27,f'{d.page} / 10');c.restoreState()
out=pathlib.Path(args.output);out.parent.mkdir(parents=True,exist_ok=True)
Doc(str(out),pagesize=A4,rightMargin=38,leftMargin=38,topMargin=35,bottomMargin=50,title='KOHC Financial Analysis - Updated Viva Guide',author='IBF Project Group',pageCompression=1).build(story,onFirstPage=footer,onLaterPages=footer)
reader=PdfReader(out);assert len(reader.pages)==10,len(reader.pages)
text='\n'.join(p.extract_text() or '' for p in reader.pages)
assert all(n in text for n in names)
assert all(f'{i}.' in text for i in range(1,18))
assert hashlib.sha256(src.read_bytes()).hexdigest()==before
pathlib.Path(args.markdown).parent.mkdir(parents=True,exist_ok=True);pathlib.Path(args.markdown).write_text('# KOHC Financial Analysis - Updated Viva Guide\n\n'+'\n'.join(md),encoding='utf8')
result={'pages':len(reader.pages),'ratio_formulas':27,'company_comparisons':27,'numbered_cases':17,'source_workbook_sha256':before,'source_workbook_unchanged':True,'output_sha256':hashlib.sha256(out.read_bytes()).hexdigest(),'page_word_counts':[len((p.extract_text()or'').split())for p in reader.pages]}
out.with_suffix('.validation.json').write_text(json.dumps(result,indent=2));print(json.dumps(result,indent=2))
