# Grand Millennium Dubai - Competitive Set Analysis Project

## 📊 Project Overview

This repository contains a comprehensive competitive set analysis for **Grand Millennium Dubai**, focusing on identifying optimal competitor hotels within a 5 km radius for revenue management and strategic positioning.

## 🎯 Objectives

- Identify non-luxury hotels within 5 km of Grand Millennium Dubai
- Analyze hotels based on multiple criteria including facilities, proximity, room mix, business mix, star rating, and reviews
- Score and rank hotels to determine primary and secondary competitive sets
- Provide strategic recommendations for revenue management

## 📁 Project Files

### Data Files
- **`Compset by distance.json`** - Complete list of 106 hotels within 5-7 km radius with distance data
- **`Compset Analysis Tool.json`** - Scoring criteria and metrics configuration
- **`filtered_hotels.json`** - 56 non-luxury hotels within 5 km (filtered results)

### Analysis Files
- **`Grand_Millennium_Dubai_CompSet_Analysis.xlsx`** - Main Excel workbook with complete analysis
  - Executive Summary Report
  - Charts & Diagrams (6 interactive Excel charts)
  - Compset Scoring Analysis
  - Hotels Database (56 hotels with 35 data points each)

### Python Scripts
- **`filter_hotels.py`** - Filters hotels by distance and category
- **`create_compset_analysis.py`** - Creates initial Excel structure
- **`populate_research_data.py`** - Populates hotel data and calculates scores
- **`create_excel_charts.py`** - Generates interactive Excel charts
- **`create_summary_report.py`** - Creates executive summary report

### Supporting Files
- **`NEW - STR Competitor Set Analysis.xlsx`** - STR analysis template
- **`Copy of STR Competitor Set Analysis.xlsx`** - Backup STR analysis

## 🏆 Key Findings

### Primary Compset (Score ≥ 90)
1. **Millennium Place Barsha Heights Hotel** - Score: 91.0
   - Distance: 1.0 km
   - Rooms: 468 + 447 apartments
   - Strong MICE capability and similar brand positioning

### Secondary Compset (Score 75-89)
1. Radisson Blu Hotel Dubai Media City - Score: 88.0
2. Golden Tulip Media Hotel - Score: 83.0
3. Mercure Dubai Barsha Heights - Score: 81.0
4. Staybridge Suites Dubai Internet City - Score: 79.0
5. Atana Hotel - Score: 77.0
6. Millennium Place Dubai Marina - Score: 76.0

## 📊 Analysis Statistics

- **Total Hotels Analyzed**: 56 non-luxury hotels within 5 km
- **Hotels with Complete Research Data**: 8 hotels
- **Average Compset Score**: 80.4/100
- **Hotels within 2 km**: 22 (immediate competition)
- **4-Star Upscale Hotels**: 40
- **3-Star Midscale Hotels**: 16

## 🔍 Scoring Methodology

Hotels are scored on a 100-point scale based on weighted criteria:

### Major Criteria
- **Location & Proximity** (7%) - Distance from Grand Millennium
- **Property & Brand Positioning** (15%) - Star rating, brand affiliation, loyalty programs
- **Room Mix & Inventory** (18%) - Total rooms, apartments, executive lounge, room types
- **Meetings & Events** (15%) - Meeting space, ballroom capacity, MICE capability
- **Facilities & Amenities** (12%) - Pool, spa, gym, restaurants, unique features
- **Online Reputation** (10%) - TripAdvisor and Booking.com ratings
- **Business Mix Alignment** (10%) - Corporate, leisure, and MICE segments
- **Technology & Innovation** (3%) - Modern amenities and tech infrastructure
- **Additional Factors** (10%) - Various other competitive considerations

### Score Interpretation
- **90-100**: Primary Compset - Direct competitors for benchmarking
- **75-89**: Secondary Compset - Important market references
- **60-74**: Extended Reference - Broader market context
- **Below 60**: Not Recommended for regular compset tracking

## 📈 Excel Charts Included

The workbook includes 6 interactive Excel charts:
1. Top 10 Hotels by Overall Compset Score (Bar Chart)
2. Guest Review Ratings Comparison (Horizontal Bar Chart)
3. Distance vs. Compset Score Analysis (Scatter Chart)
4. Compset Recommendation Distribution (Pie Chart)
5. Meeting Space Comparison (Bar Chart)
6. Total Room Inventory Comparison (Bar Chart)

## 💡 Strategic Recommendations

1. **Primary Benchmarking**: Focus on Millennium Place Barsha Heights as primary competitor
2. **Rate Positioning**: Maintain ADR within 10-15% of primary compset average
3. **Market Segments**: Target corporate accounts in Media City, Internet City, Knowledge Village
4. **Differentiators**: Leverage apartment offerings and executive lounge
5. **Review Scores**: Target 8.5+ ratings to compete with top-rated properties
6. **MICE Strategy**: Emphasize meeting space capabilities and technology

## 🔄 Next Steps

- [ ] Complete research on remaining 48 hotels
- [ ] Gather April 2025 rate data for ADR comparisons
- [ ] Conduct detailed MICE capability assessments
- [ ] Validate corporate account overlap
- [ ] Set up monthly STR compset monitoring
- [ ] Quarterly compset review and adjustments

## 🛠️ Requirements

### Python Libraries
```bash
pip install pandas openpyxl
```

### Software
- Microsoft Excel 2016 or later (for viewing charts)
- Python 3.8+ (for running scripts)

## 📝 Usage

1. Clone the repository
2. Open `Grand_Millennium_Dubai_CompSet_Analysis.xlsx` to view the complete analysis
3. Run Python scripts to regenerate analysis:
   ```bash
   python filter_hotels.py
   python create_compset_analysis.py
   python populate_research_data.py
   python create_excel_charts.py
   python create_summary_report.py
   ```

## 👥 Author

**Grand Millennium Dubai Revenue Management Team**

## 📅 Last Updated

November 2025

## 📄 License

This project is proprietary and confidential. For internal use only.

---

**Note**: This analysis is based on publicly available data from hotel websites, booking platforms, and review sites as of November 2025. For the most accurate competitive intelligence, supplement with STR reports and direct market intelligence.
