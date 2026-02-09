# Paper-Products-Pricing-Model
This pricing model enables **tissue paper converters** (manufacturers who purchase parent reels and convert them into finished products) to accurately calculate product costs and generate export quotations for international markets.
https://docs.google.com/spreadsheets/d/1HmTRq3jr9xv4m3NLMx6gq8msM0yMplKO/edit?usp=sharing&ouid=117695434224299199084&rtpof=true&sd=true

**Why it exists:** Export pricing for paper products involves complex calculations across material costs, converting operations, packaging, logistics, and market-specific factors. This tool automates these calculations, reducing quote preparation time from hours to minutes while improving accuracy and consistency.

---

## 🛠️ Tech Stack

| Component | Technology | Purpose |
|-----------|------------|---------|
| **Platform** | Microsoft Excel 2016+ | Core application |
| **Formulas** | VLOOKUP, INDEX-MATCH, IFERROR | Data lookups and calculations |
| **Data Validation** | Excel Lists | Dropdown selections |
| **Structure** | Excel Tables | Organized data management |
| **Compatibility** | Google Sheets (limited) | Alternative platform |

### Why Excel?

- ✅ No programming knowledge required
- ✅ Universally available in business environments
- ✅ Easy to customize and maintain
- ✅ Works offline
- ✅ Familiar interface for sales and operations teams

---

## 📁 Data Source

### Data Structure Overview

![Data Structure Overview](https://github.com/medhatripathi/Paper-Products-Pricing-Model/blob/main/Data%20Structure%20Overview.png)

The model uses **7 interconnected sheets** organized into three categories:

### Primary Data Tables

| Table | Records | Description |
|-------|---------|-------------|
| **Parent Reels** | 367 | Reel specifications with pricing (1-ply, 2-ply, 3-ply) |
| **Finished Products** | 151 | Product configurations (toilet tissue, kitchen towel, napkins) |
| **Destinations** | 9 | Export markets with freight rates |
| **Cost Parameters** | 50+ | Converting, packaging, overhead, logistics costs |

### Data Fields: Parent Reels (INPUT_Reels)

| Field | Type | Example | Description |
|-------|------|---------|-------------|
| Reel_Code | Text | R-2P-17-V | Unique identifier |
| Supplier | Text | Supplier A | Source mill |
| Ply | Number | 2 | Number of plies |
| GSM_Per_Ply | Number | 17 | Grams per square meter |
| Virgin_Pct | Percent | 100% | Virgin fiber content |
| Wet_Strength | Text | Yes/No | Wet strength treatment |
| Color | Text | White | Paper color |
| Softness | Text | Premium | Softness grade |
| Price_INR_MT | Currency | 108,000 | Price per metric ton |

### Data Fields: Finished Products (INPUT_Products)

| Field | Type | Example | Description |
|-------|------|---------|-------------|
| Product_Code | Text | TT-2P-200-S | Unique identifier |
| Category | Text | TT | Product category |
| Ply | Number | 2 | Number of plies |
| GSM_Per_Ply | Number | 17 | Weight per ply |
| Sheet_Length_mm | Number | 100 | Sheet length |
| Sheet_Width_mm | Number | 114 | Sheet width |
| Sheets_Per_Unit | Number | 200 | Sheets per roll |
| Units_Per_Carton | Number | 48 | Packing configuration |
| Matched_Reel_Code | Text | R-1P-17-V | Links to parent reel |

### Data Refresh Requirements

| Data Type | Frequency | Source |
|-----------|-----------|--------|
| Exchange Rates | Weekly | Bank/Financial service |
| Reel Prices | Monthly | Supplier price lists |
| Freight Rates | Monthly | Freight forwarder |
| Converting Costs | Quarterly | Internal accounts |
| Overhead Costs | Annually | Internal accounts |

---

## ✨ Features & Highlights

### Business Problem

**Tissue paper converters face several pricing challenges:**

1. **Complex Cost Structure**
   - Material costs vary by reel specification (GSM, ply, virgin %, softness)
   - Converting costs differ by product complexity (embossing, printing)
   - Packaging costs depend on product size and configuration
   - Export costs vary significantly by destination

2. **Time-Consuming Quote Process**
   - Manual calculations take 2-4 hours per quote
   - High risk of calculation errors
   - Inconsistent pricing across sales team
   - Difficulty comparing product profitability

3. **Market Responsiveness**
   - Slow response to customer inquiries
   - Inability to quickly adjust for raw material price changes
   - Limited visibility into cost drivers

4. **Profitability Visibility**
   - Unknown margin by product/destination
   - Difficulty identifying cost reduction opportunities
   - No standardized pricing methodology

### Goal of the Dashboard

| Objective | Target | Metric |
|-----------|--------|--------|
| **Reduce quote preparation time** | 90% reduction | From 2-4 hours to 10-15 minutes |
| **Improve pricing accuracy** | <2% variance | Compared to actual costs |
| **Standardize pricing methodology** | 100% adoption | Across all sales team |
| **Enable quick scenario analysis** | Real-time | Margin/destination/volume changes |
| **Professional customer communication** | Consistent | Branded quote documents |

### Core Features

![Core feature](https://github.com/medhatripathi/Paper-Products-Pricing-Model/blob/main/Core%20Features.png)

### Business Impact & Insights

#### Quantifiable Benefits

| Impact Area | Before | After | Improvement |
|-------------|--------|-------|-------------|
| Quote preparation time | 2-4 hours | 10-15 minutes | **90% faster** |
| Calculation errors | 5-10% quotes | <1% quotes | **90% reduction** |
| Pricing consistency | Variable | Standardized | **100% consistent** |
| Cost visibility | Limited | Full breakdown | **Complete transparency** |
| Scenario analysis | Manual | Instant | **Real-time decisions** |

#### Strategic Insights Enabled

1. **Product Profitability Analysis**
   - Identify highest margin products by destination
   - Compare virgin vs recycled fiber economics
   - Evaluate premium vs economy positioning

2. **Market Optimization**
   - Determine most profitable export destinations
   - Analyze freight impact on pricing competitiveness
   - Identify markets for focused sales effort

3. **Cost Driver Visibility**
   - Understand material cost as % of total
   - Quantify certification cost impact
   - Measure converting complexity effects

4. **Pricing Strategy Support**
   - Set data-driven margin targets by product
   - Model volume discount scenarios
   - Evaluate private label vs branded pricing

---

## 🖥️ Dashboard Walkthrough

### Sheet 1: Input Costs

**Purpose:** Central repository for all cost parameters

![image](https://github.com/medhatripathi/Paper-Products-Pricing-Model/blob/main/Pg1%20Input%20Costs.png)


**Key Elements:**
- Yellow cells indicate user inputs
- All other sheets reference these values
- Update frequency: Weekly for FX, Monthly for costs

---

### Sheet 2: Input Reels

**Purpose:** Database of parent reel specifications and pricing

![image](https://github.com/medhatripathi/Paper-Products-Pricing-Model/blob/main/Pg2%20Input%20Reels.png)


**Key Elements:**
- Unique reel code naming convention
- Multiple suppliers for comparison
- Automatic USD conversion

---

### Sheet 3: Input Products

**Purpose:** Database of finished product specifications

![image](https://github.com/medhatripathi/Paper-Products-Pricing-Model/blob/main/Pg3%20Input%20Products.png)


**Key Elements:**
- Complete product specifications
- Links to parent reel database
- Packaging configuration for logistics

---

### Sheet 4: LookUp Value

**Purpose:** Pull values for selected product

![image](https://github.com/medhatripathi/Paper-Products-Pricing-Model/blob/main/Pg4%20LookUp%20Value.png)


**Key Insights from this sheet:**
- Material (reel) cost is largest component (~65%)
- Identifies cost reduction opportunities
- Instant margin calculations

---

### Sheet 5: Calc Product

**Purpose:** Calculate ex-works cost for selected product

![image](https://github.com/medhatripathi/Paper-Products-Pricing-Model/blob/main/Pg5%20Calc%20Product.png)


**Key Insights from this sheet:**
- Material (reel) cost is largest component (~65%)
- Identifies cost reduction opportunities
- Instant margin calculations

---

### Sheet 6: Calc Export

**Purpose:** Calculate landed cost for export destinations

![image](https://github.com/medhatripathi/Paper-Products-Pricing-Model/blob/main/Pg6%20Calc%20Export.png)


**Key Insights from this sheet:**
- Complete landed cost visibility
- Container optimization
- Easy comparison across destinations

---

### Sheet 7: Quote Summary

**Purpose:** Generate professional customer quotation

![image](https://github.com/medhatripathi/Paper-Products-Pricing-Model/blob/main/Pg7%20Quote%20Summary.png)


**Key Features:**
- Professional, print-ready format
- All data auto-populated
- Customizable terms and branding

---

## 🗺️ Roadmap

**PHASE 1: FOUNDATION ✅ (Current Release)**
- Basic cost calculation
- Export pricing
- Quote generation
- Product/Reel databases

**PHASE 2: REFINEMENT 🔲 (Planned)**
- Multi-product quoting
- RFQ input template
- Sensitivity analysis
- Dashboard visuals
- Competitor comparison

**PHASE 3: AUTOMATION 🔲 (Future)**
- Customer database
- Price history tracking
- Order management
- Profitability analysis
- External data integration
