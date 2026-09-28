#JCars Logistics Analysis
##1. Project Background & Objective
###JCars Logistics imports, sells and delivers vehicles to customers across Kenya. Management supplied a single raw, uncleaned transactional export (Jcars_data.csv) covering sales, customers, vehicles, branches, sales representatives, payments, deliveries, logistics costs, returns, cancellations and customer experience and asked for a reliable, interactive Power BI solution that turns that data into management decision support: understanding sales and revenue performance, cost and profitability, vehicle and branch performance, sales channels, payments, logistics, returns and unusual transactions that deserve further investigation.
###This repository documents that end-to-end journey:
### 1.  Data quality audit.
### 2.  Currency standardization.
### 3.  Power Query cleaning.
### 4.  Data modelling.
### 5.  DAX.
### 6.  Dashboard and report design.
### 7.  Investigation and recommendations.
## 2.   Dataset & Grain
●	Source file: Jcars_data.csv is a single flat table with 32 columns.

●	Grain: one row contains one vehicle sales order line, a single order for one or more units of one vehicle 

●	configuration, sold by one sales rep, to one customer, through one branch.

●	Row count: 276 order lines, representing 466 total units sold.

●	Date coverage: Order Date spans 1 Jan 2025 – 1 Dec 2026.

●	Business entities represented in the columns: customer (name, type, age), location (region, county, city), branch, sales rep, lead source/channel, vehicle (make, model, type, year, fuel, transmission, color), transaction economics (units, price, cost, discount, delivery fee, revenue recorded), payment (method, status), logistics (delivery status, delivery date, logistics cost), and customer experience (rating, review count, returned flag).

## 3. Data Quality Audit
### Below are the significant issues identified, why each mattered, and how each was resolved in Power Query.


```

	Issue Found	Why It's a Problem	How It Was Handled
1	Inconsistent/misspelled categorical text across 15 fields (Region, County, City, Branch, Car Make, Car Model, Fuel Type, Transmission, Color, Payment Method, Payment Status, Delivery Status, Customer Type, Lead Source, Returned e.g. "totoya", "toyta", "Mercedes Benz", "cental", "nrb", "mtkenya", "harier"	Splits one real-world category into many, silently understating totals in every chart grouped by that field	Built explicit lookup/mapping tables per field so every known variant is normalized to one canonical label; anything unrecognized falls back to a cleaned proper-case value instead of being dropped
2	Mixed currencies in monetary fields (unmarked values, KES/KSh, $/USD, EUR, ZAR/R)	Summing mixed currencies as if they were all KES massively distorts revenue, cost and profit	Currency detected from the raw text and converted to KES using one fixed rate per currency, applied consistently.
3	Shorthand monetary values ending in "M" (e.g. values meant to read as millions)	Read literally, these numbers are ~1,000,000× too small	Detected the trailing "M" and multiplied by 1,000,000 before currency conversion
4	Placeholder / error tokens standing in for missing values ("", "N/A", "NULL", "TBD", "-", "#VALUE!", "#ERROR", "#DATE!") scattered across text, numeric and date fields	If left as text, these values break numeric aggregation or get counted.	Explicitly recognized and converted to true nulls before typing.
5	Inconsistent date formats and Excel serial date numbers in Order Date / Delivery Date	Silent misanalyses’ shift orders into the wrong period, corrupting every trend visual	Custom analysis tries serial-number conversion first (values in the 30000–60000 range), then en-GB, then en-US text formats, before falling back to null
6	Inconsistent Order ID formats (ORD, LCL, CAR, LC prefixes mixed with stray characters)	Breaks any attempt to use Order ID as a clean identifier or to de-duplicate	Prefix detected and preserved, non-alphanumeric noise stripped, standardised to PREFIX + digits
7	Units Sold recorded as words ("one", "two", "three") instead of numbers in some rows	Text values are excluded from SUM() and silently understate volume	Word-to-number mapping applied before numeric typing
8	Discount recorded in mixed formats (words like "fifteen", "ten percent", decimals, percentages) and some implausible values above 50%	Mixed formats break numeric math; a discount above 50% on a vehicle sale is not credible and likely a data-entry error	Analysed to a decimal fraction; values above 50% treated as invalid and excluded rather than trusted
9	Customer Rating recorded in mixed formats ("excellent", "4/5", "3 out of 5") and out-of-range values	Inconsistent scales cannot be averaged meaningfully	Normalized to a 0–5 numeric scale; anything outside 0–5 treated as invalid
10	Review Count partly recorded as words (e.g. "ten")	Same issue as Units Sold  breaks SUM()/AVERAGE()	Word-to-number conversion applied
11	Customer Age outside a plausible adult range (below 18 or above 100)	Data-entry errors	Treated as invalid (nulled), then given a documented default
12	Sales Rep names containing a systematic typo pattern (a digit "1" in place of the letter "i", e.g. names rendering incorrectly)	Splits one sales rep into two identities, corrupting rep-level performance analysis	Character substitution applied before matching against the known rep list
13	Zero-value Logistics Cost on orders marked "Delivered"	A completed delivery with zero recorded logistics cost is not credible and would understate true delivery cost	Flagged as suspicious and nulled rather than treated as a true zero, so it doesn't silently deflate logistics-cost analysis
14	Zero Revenue Recorded on orders that were neither Cancelled nor Refunded	A live, non-cancelled order showing zero revenue is inconsistent with the business logic and would distort revenue reporting if trusted	Flagged as suspicious and nulled; Revenue is independently recalculated rather than relying on this field
15	Duplicate category spellings for the same lead source (e.g. "Face Book" vs "Facebook")	Splits one channel into two in every lead-source chart	Explicit text replacement to merge the variant into the canonical label
16	Recorded "Revenue" not independently verifiable	The raw file's revenue-style field could not be trusted at face value given Issue 14	A transaction-level Revenue measure was rebuilt independently as Units Sold × Unit Selling Price × (1 − Discount) + Delivery Fee, and validated against the (cleaned) recorded figure rather than assumed correct — see Data Validation

```
###Every row that could not be corrected with confidence was retained, not deleted, and stamped with a `Data Quality Flag` column recording exactly what was uncertain about it (missing Order Date, missing Delivery Date, estimated Price, estimated Cost). 76 of 276 rows (27.5%) carry at least one flag.
## 4. Currency Standardization
###All monetary values are reported in Kenya Shillings (KES). Where a value did not explicitly state a currency, it was assumed to be KES. Where another currency was explicitly indicated, it was converted using one fixed rate applied consistently throughout the project:
```
Currency detected	Marker(s) matched	Rate to KES applied
KES / KSh	"KES", "KSH", or no marker (default)	No conversion
US Dollar	$, "USD"	129.54
Euro	"EUR"	147.84
South African Rand	"ZAR", or values prefixed "R "	7.93

```
###Currency conversion is applied before the "M" (millions) shorthand adjustment and before any downstream calculation, so Revenue, Unit Selling Price, Unit Cost and Logistics Cost are all expressed on a common KES basis before they are combined.## 5. Data Cleaning & Preparation (Power Query)
###All cleaning lives in the Jcars_data-CLEANED query, built from reusable custom functions rather than one-off manual steps, so the logic is transparent and repeatable:
●	fnCat — normalizes free text against a mapping table (case/space/punctuation-insensitive key match); unmapped-but-recognizable text is title-cased rather than discarded and recognized "empty" tokens become "Unknown".
●	fnDate — parses Order Date / Delivery Date, handling Excel serial numbers and both en-GB/en-US text formats.
●	fnMoneyKES — detects currency markers, strips formatting, applies the "M" multiplier and the currency conversion table above, and treats error/placeholder tokens as null.
●	fnOrder — standardises Order ID prefixes.
●	fnAge,  fnYear, fnUnits, fnDisc, fnRating, fnCount — field-specific parsing and validity-range checks for Customer Age, Vehicle Year, Units Sold, Discount, Customer Rating and Review Count respectively.
## 6. Data Validation
### Checks performed on the analysis-ready data before modelling:
●	Type check: dates typed as date, monetary fields typed as Currency, so no numeric field is silently treated as text.
●	Row-count check: row count in Jcars_data-CLEANED matches Jcars_data-ORIGINAL no records were dropped during cleaning.
●	Crossfield business-rule check: Logistics Cost = 0 on "Delivered" orders, and Revenue Recorded = 0 on non-Cancelled/non-Refunded orders, were both checked against the business logic and treated as suspicious.
●	Recalculation vs recorded value: the raw "Revenue Recorded" field was compared against an independently rebuilt Revenue measure (Units Sold × Price × (1 − Discount) + Delivery Fee); the rebuilt figure is the one used throughout the model, with the discrepancy documented rather than silently overwritten.
●	Category coverage check: every categorical mapping table was built directly from the distinct raw values present in the file, so no known variant is left unmapped.
## 7. Data Model
###The raw flat file was restructured into a star schema:
```
 
                DIM-CUSTOMER  (Customer Type, CUST KEY)
                      │
DIM-LEAD SOURCE ──┐   │
 (Lead Source,    │   │
  Index)          │   │
                   ▼   ▼
              FACT TABLE 2  ◄── DIM-VEHICLE (Car Make, Car Model, Vehicle Type,
        (order-line grain:                    Vehicle Year, Fuel Type,
     dates, region/county/city/branch,         Transmission, Colour, Index)
   units, price, cost, discount, fees,
  payment, delivery, rating, returned,   ◄── DIM-SALES REP (Sales Rep, Index)
   Revenue, Cost of Goods Sold, etc.)
```
●	Fact table: FACT TABLE 2  one row per order line, built from a staging query (REFERENCE, sourced from Jcars_data-CLEANED) merged against each dimension to attach surrogate keys, with the descriptive dimension attributes removed from the fact table afterwards to avoid duplication.
●	Dimensions: DIM-CUSTOMER (by Customer Type), DIM-VEHICLE (by the full vehicle attribute combination  Make/Model/Type/Year/Fuel/Transmission/Colour), DIM-SALES REP, DIM-LEAD SOURCE. Each dimension is built with Table.Distinct + a surrogate index key.
●	Relationships: all four are many-to-one from FACT TABLE 2 to the dimension, with bidirectional cross-filtering enabled so slicers on either side filter the other.
●	Date intelligence: Power BI's auto date/time tables support the Order Date and Delivery Date hierarchies (year/quarter/month/day) used in the trend visuals.
●	
##  8. DAX Measures & Calculated Columns
```
Type	Name	DAX	Purpose
Measure	Total revenue	SUM('FACT TABLE 2'[Revenue])	Core revenue KPI
Measure	Total Units Sold	SUM('Jcars_data-CLEANED'[Units Sold])	Core volume KPI
Measure	Total Gross Profit	[Total revenue] - SUM('FACT TABLE 2'[Cost of goods Sold])	Core profitability KPI
Measure	Total Gross Profit Margin	([Total Gross Profit]/[Total revenue])*100	Profitability ratio (%)
Calculated column	Cost of goods Sold (on FACT TABLE 2)	'FACT TABLE 2'[Unit Cost] * 'FACT TABLE 2'[Units Sold]	Line-level cost basis feeding Gross Profit
Calculated column	revenue per unit (on FACT TABLE 2)	'FACT TABLE 2'[Revenue] - 'FACT TABLE 2'[Unit Cost]	Line-level margin contribution
```
### Business definitions used: Revenue = independently rebuilt transaction value (units × price, net of discount, plus delivery fee) rather than the raw "Revenue Recorded" field; Gross Profit = Revenue − Cost of Goods Sold (unit cost × units); Gross Profit Margin = Gross Profit ÷ Revenue.
## 9. Report Pages
## 1. Executive Dashboard (JCARS DASHBOARD)
A single-page overview built around four headline KPI cards (Total Units Sold, Total Revenue, Total Gross Profit, Total Gross Profit Margin), supported by:
●	A map of Gross Profit by Branch
●	Gross Profit and Units Sold trend by Customer Type
●	Gross Profit and Margin by Car Make (combo chart)
●	Gross Profit share by Lead Source (pie)
●	Gross Profit by Region (funnel)
●	Units Sold by Sales Rep
●	Units Sold by Delivery Status
●	Gross Profit, Revenue and order count by quarter (Delivery Date)
●	A Vehicle Type slicer that cross-filters the whole page
## 2.  Detailed Analysis (MESSY BLUEPRINT)
This is a deeper investigation page, carrying the KPI cards forward plus: a treemap of Units Sold by Car Model, pivot tables (Gross Profit/Revenue/Units/COGS by Vehicle Type; Returned status by Vehicle Type), a map of Gross Profit by Branch, Gross Profit by Region (funnel), Units Sold by Sales Rep, Gross Profit share by Lead Source, Gross Profit/Revenue/orders by quarter, Payment Method share (donut), Payment Status by Units and Revenue (combo), Delivery Status by Units, Customer Type trend, a Logistics Cost / Gross Profit table by Vehicle Type and Year, a KPI visual, a navigation button and a Vehicle Type slicer.
### Interactivity implemented
●	Cross-page slicer (Vehicle Type) and cross-filtering between visuals on both pages
●	An in-report navigation button on the detail page
●	 There was integration of drill throughs for city, county, fuel type, and returned status.
##10.  Assumptions & Business Rules
●	Monetary values with no explicit currency marker are assumed to be KES.
●	Fixed exchange rates (USD 129.54, EUR 147.84, ZAR 7.93) are applied consistently project-wide rather than varying by transaction date cite your source/date in the final write-up.
●	Revenue is defined and calculated as Units Sold × Unit Selling Price × (1 − Discount) + Delivery Fee, not taken from the raw "Revenue Recorded" field.
●	Gross Profit = Revenue − (Unit Cost × Units Sold); Gross Profit Margin = Gross Profit ÷ Revenue.
●	Discount values above 50% are treated as invalid rather than trusted.
●	Customer Rating is normalized to a 0–5 scale; values outside that range are treated as invalid.
●	Customer Age outside 18–100 is treated as invalid.
●	Logistics Cost of 0 on a "Delivered" order, and Revenue Recorded of 0 on a non-cancelled/non-refunded order, are both treated as suspicious rather than genuine zeros.
●	Where a genuinely missing value could not be corrected, it was retained and flagged (via Data Quality Flag) rather than deleted; for measures that must still aggregate, the following documented defaults are used: Customer Age → 35, Vehicle Year → 2022, Units Sold → 1, Discount → 0, Customer Rating → 3, Review Count → 0, Delivery Fee → 0, Logistics Cost → 0, Order ID → "UNKNOWN", missing Unit Selling Price/Unit Cost → 0 (and separately flagged as estimated).
## 11.  Key Business Metrics (from the analysis-ready data)
Metric	Value
Total Units Sold	466
Total Revenue	KES 1,897,538,016
Total Cost of Goods Sold	KES 1,482,051,475
Total Gross Profit	KES 415,486,541
Gross Profit Margin	21.9%
Orders analysed	276
Rows carrying a data-quality flag	76 (27.5%)
Returned = "Yes"	88 orders (31.9%)
Delivery Status = Cancelled	37 orders (13.4%)
Payment Status = Cancelled	27 orders (9.8%)

Top contributors: Toyota and Volkswagen together generate 69% of total Gross Profit; Thika and Nairobi HQ branches together generate 65% of total Gross Profit; SUVs are the leading vehicle type by units sold (178 of 466 units).
12. Analyst-Defined Business Questions
Five additional questions developed after exploring the data, three of which are answered directly below using the analysis-ready dataset:
1.	Is any lead source actually losing the business money once cost is accounted for? The answer was yes. Orders sourced via WhatsApp show a negative aggregate Gross Profit of roughly −KES 3.49 million, the only channel in negative territory, despite Walk-in and Website each contributing over KES 140 million in Gross Profit. This deserves investigation before further budget is put behind that channel.
2.	How concentrated is profitability in a small number of makes and branches, and what does that mean for risk? Toyota and Volkswagen alone account for 69% of total Gross Profit, and Thika and Nairobi HQ alone account for 65%. That level of concentration means a supply disruption for either make, or an operational issue at either branch, would have an outsized impact on company profitability.
3.	Does the order pipeline show a genuine slowdown in late 2026, or is this a data-completeness artefact? (Answered.) Order volume drops sharply after Q2 2026 (only 1 order in Q3 2026 and 3 in Q4 2026, versus 34–84 orders per quarter through mid-2026). This pattern is far more consistent with the export having been taken partway through the quarter than with a real, sudden collapse in demand  worth confirming with the data owner rather than reporting as a genuine downturn.
## 13. Investigation & Exceptions
### Patterns worth flagging to management rather than treating as routine:
●	WhatsApp lead source: negative Gross Profit (KES 3.49M) , the only channel losing money in aggregate.
●	31.9% of orders flagged as Returned = "Yes” a high return rate for vehicle sales; worth segmenting by make/model/branch/rep to see if it concentrates anywhere before assuming it's evenly spread.
●	2025 Q2 carries 62% of the entire period's Gross Profit (KES 259.4M of KES 415.5M) despite being one of eight quarters in the dataset a single unusually strong quarter (or a small number of very large orders within it) rather than a steady run-rate; should be decomposed by order before being used to set forward targets.
●	27.5% of rows carry at least one data-quality flag, largely missing Order/Delivery dates or estimated Price/Cost. This is high enough that any KPI filtered to a narrow date range should be cross-checked against the flag rate for that slice.
●	Sparse data in Q3–Q4 2026 - treat as likely export-cutoff rather than a genuine sales collapse 14. Management Insights

4.	Profitability is heavily concentrated, not evenly spread. Two car makes (Toyota, Volkswagen) generate roughly seven of every ten shillings of gross profit, and two branches (Thika, Nairobi HQ) generate roughly two of every three. This is a business built on a narrow profitable core rather than broad, even performance  useful to know when deciding where to protect supply and staffing.
5.	Not every lead source is worth the same investment. Walk-in and Website traffic are the two strongest channels by Gross Profit (over KES 140M each), while WhatsApp-sourced orders are net loss-making in aggregate. Marketing/lead-generation spend allocated evenly across channels would be misallocated relative to what the data shows.
6.	A meaningful share of "Delivered" and "Paid" business still carries data-quality risk. 76 of 276 order lines (27.5%) needed a missing date or an estimated price/cost filled in to be usable at all — management dashboards built on this data should be read with that caveat in mind, especially at granular (branch/rep/month) cuts where a handful of flagged rows can swing a percentage meaningfully.
7.	Returns run higher than a "background noise" rate would suggest. Nearly a third of orders (31.9%) are flagged Returned = "Yes." Whether this reflects a genuine quality or fulfilment issue, or an artefact of how "Returned" was recorded, cannot be determined from this field alone and needs to be cross-checked against Delivery Status and specific vehicle records before management treats it as a real operational problem.
8.	SUVs lead by volume (178 of 466 units, 38%), but volume leadership and profit leadership are not the same thing Car Make, not Vehicle Type, is where Gross Profit concentrates most sharply. Volume-based decisions (e.g. inventory mix) and profit-based decisions (e.g. where to negotiate better supplier terms) should be made from different cuts to this data, not the same one.
15. Management Recommendations
9.	Investigate the WhatsApp lead-generation channel before continuing to invest in it. It is the only lead source with negative aggregate Gross Profit in the dataset. Confirm whether this reflects genuinely lower-margin deals coming through that channel, a data-recording issue specific to WhatsApp-sourced orders, or a small number of outlier transactions then decide whether to adjust discounting rules for that channel, reduce its budget, or fix the recording issue.
10.	Review supplier and inventory concentration risk in Toyota and Volkswagen, and operational concentration in the Thika and Nairobi HQ branches. With 69% of Gross Profit riding on two makes and 65% on two branches, a disruption to (supply delay, branch-level operational issue) would have a disproportionate effect on company profitability. This does not mean diversifying away from what's working it means confirming contingency plans exist for the makes and branches the business is most dependent on.
11.	Audit the "Returned = Yes" population (88 orders) by make model, branch and sales rep before treating the 31.9% rate as normal. If returns cluster around specific vehicles or locations, that points to a fixable quality or fulfilment issue; if they're evenly spread, it may simply reflect how the field was recorded, and the definition of "Returned" itself should be reviewed.
12.	Confirm whether the sharp order drop in Q3–Q4 2026 reflects a real business event or an incomplete data export before it is used in any performance narrative or forecast. Using it as-is risks either wrongly alarming management about a "slowdown" or, if corrected later, requiring a retraction.
## 16. Repository Structure
JCARS PROJECT FOLDER
●	README.md File
DATA FOLDER
●	EXCEL ORIGINAL DATA SOURCE
DATA ANALYSIS FOLDER
●	POWER BI FILE
IMAGES FOLDER
●	SCREENSHOTS

 
## 17. Tools Used
●	Power BI Desktop - data modelling, Power Query (M), DAX, report design
●	Power Query (M) -data quality remediation, currency standardization, star-schema construction
## 18. Challenges & Learnings
The single biggest challenge in this dataset was that almost every column had its own distinct failure mode categorical typos, mixed currencies, mixed units of measure ("M" for millions), numbers written as words, and business-rule violations (zero logistics cost on delivered orders) so a single generic cleaning pass would not have caught most of it. Building small, reusable, purpose-built functions per problem type (fnCat, fnDate, fnMoneyKES, etc.) rather than one-off manual fixes made the cleaning transparent, auditable, and easy to re-run if the raw export is refreshed. The clearest lesson: a "Revenue" or "Cost" field recorded directly in a raw export should never be trusted at face value without being cross-checked against an independently rebuilt calculation in this dataset, that check is what surfaced the zero-revenue and zero-logistics-cost anomalies used above.



