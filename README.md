Call Centre Performance Analysis: AstroSage vs AstroGuru

An end-to-end Excel analysis of call centre data (calls and chats) from two astrology platforms, AstroSage (App) and AstroGuru (Gurucool). The project cleans the raw data, answers 14 objective and 10 subjective business questions with Pivot Tables, charts and formulas, and ends with a recommendation for allocating a ₹1 crore investment.

Table of Contents
Project Overview
Dataset
Tools and Techniques
Data Cleaning
Key Findings
Strategic Recommendations
Investment Allocation (₹1 Crore)
Dashboard
Workbook Structure
How to Use
Project Overview

The goal is to understand call centre performance and decide where to invest: hiring more agents, training, or technology. The analysis covers:

Call and chat volume trends (daily and monthly)
Operational cost and sales by consultation type
Agent (guru) workload and performance
Customer retention (repeat callers)
Customer satisfaction and what drives it
Platform comparison: AstroSage (App) vs AstroGuru (Gurucool)
Dataset
Property	Value
Source table	data sheet (1 primary table)
Records	28,027
Attributes	36
Period	1 Dec 2023 to 3 Jan 2024 (34 operational days)
Call sessions	8,365
Agents (gurus)	131

Attribute groups

Customer and guru: user, uid, guru, guruName, gid
Consultation: consultationType, Updated ConsultationType, Websites, Updated Websites
Call metrics: chatSeconds, callStatus, astrologerCallStatus, userCallStatus, astrologerOnCallDuration, timeDuration
Revenue: amount, Updated amount, astrologersEarnings, Updated AstroEarnings, netAmount
Satisfaction: rating
Derived time fields: Date, Month, Year, Hour
Other: refundStatus, region, Updated region, Repeat Caller, Trimmed_gur_name

Of the 36 attributes, 24 are categorical, 8 are continuous, and 4 are date/time related.

Tools and Techniques
Microsoft Excel: Pivot Tables, charts, slicers, KPI cards
Functions: IF, IFERROR, VLOOKUP, TRIM, PROPER, UPPER, LOWER, TEXT, SUM, COUNTA, COUNTIF, AVERAGE, MAX, MIN, CORREL
Charts: line, column, bar, pie, doughnut, scatter plot
Data Cleaning

A Cleaned Data sheet was built from the raw data to handle missing values and inconsistent text:

excel
=PROPER(TRIM([@guruName]))                                  // standardise guru names
=TEXT([@createdAT],"yyyy")                                  // extract year
=IF(chatStatus="","Unknown",chatStatus)                     // fill blank chat status
=IF(consultationType="","UNKNOWN",UPPER(TRIM(consultationType)))  // standardise consultation types
=TRIM(LOWER(Websites))                                      // standardise website names
=IF(amount="","0",amount)                                   // replace blank amounts
Key Findings
Volume and operations
Metric	Result
Total call sessions	8,365
Average daily call volume	246 (peak 430, minimum 107)
Average daily chat volume	573.94
Highest-volume month	December (8,090 calls)
Lowest-volume month	January (418 calls)
December vs January	~19.35x more calls
Total operational cost	₹99,146.57 (Dec ₹93,786.16, Jan ₹5,360.41)
Avg calls per agent per day	1.9 (8,365 / (131 x 34))
Customers and revenue
Metric	Result
Repeat callers	1,275 (56.6% of calls)
Total sales	₹213,987.32
CALL sales	₹168,442.04 (79%)
CHAT sales	₹45,494.68 (21%)
PUBLIC_LIVE_CALL sales	₹50.60
Most active guru	Astro Krishaa (1,580 calls)
Most active user	UID 33017 (158 calls)
Customer satisfaction
Metric	Result
Average rating	2.93 (out of 8)
Correlation: call duration vs rating	-0.0007 (no meaningful relationship)
Highest-rated guru	Tarot Mystical (7.50)
Lowest-rated guru	Tarot Rittika (0.00)
Best month	December (2.95)
Weakest month	January (2.68)
What drives satisfaction
Call duration has virtually no effect on ratings, so shorter calls will not improve satisfaction.
Net amount has the strongest relationship with rating (-0.191): users who pay more tend to rate lower.
Chat makes up about 69.6% of interactions but has the lowest average rating (2.7).
App users rate slightly higher (2.98) than Gurucool users (2.92).
Late-night hours (9 to 11 PM) show noticeably lower ratings.
Platform comparison
Platform	Calls	Avg rating
Gurucool (AstroGuru)	20,225	2.92
App (AstroSage)	7,800	2.98

Gurucool handles about 2.59x the volume with similar satisfaction.

Workload balance
The busiest agent handles 1,580 calls versus 743 for the lowest of the high-volume agents: a gap of 837 calls (ratio 2.13x).
Two agents (GID 287 and 239) handle about 10.8% of all interactions.
Some agents combine high volume with very low ratings (for example GID 201: 735 calls, rating 1).
Strategic Recommendations
Prioritise agent training. Focus on GIDs rated below 3 and on chat communication skills.
Postpone large-scale hiring. Workload per agent is manageable; revisit only if volume grows significantly.
Plan peak-hour staffing. Add capacity and route experienced agents to 9 to 11 PM slots.
Balance workloads. Cap calls per agent, redistribute from overloaded agents, and give fewer calls to agents with very long average durations.
Review pricing and value perception. Higher net amounts correlate with lower ratings.
Improve the Gurucool web experience to close the satisfaction gap with the App.
Run a mentor programme pairing top-rated agents (GID 97, 305) with low-rated ones (GID 201, 219).
Adopt supporting technology: AI-based call routing, CRM, workforce management tools, real-time dashboards, and automated feedback surveys.
Investment Allocation (₹1 Crore)
Share	Amount	Area	Focus
50%	₹50 Lakhs	Agent training and skill development	Communication, problem-solving, first-call resolution
30%	₹30 Lakhs	Technology and process optimisation	Intelligent routing, queue management, CRM, performance monitoring
20%	₹20 Lakhs	Workforce planning and peak-hour management	Peak staffing, workload balancing
Dashboard

The final dashboard combines:

KPI cards: Total Calls, Total Revenue, Average Satisfaction, Repeat Caller %, Average Calls per Agent, Peak Daily Calls
Charts: Daily Call Volume Trend, Customer Satisfaction Trend, Call Status Distribution, Rating-wise User and Guru Distribution, Platform Performance, Revenue by Consultation Type, Top Rated Astrologers
Slicers: Year, Month, Website, Consultation Type
Workbook Structure
Sheet	Purpose
data	Raw source table
Cleaned Data	Standardised, analysis-ready data
Objective Analysis	Pivot Tables and charts for Q1 to Q14
Subjective Analysis	Supporting analysis for Q1 to Q10
Dashboard	KPI cards, charts and slicers
How to Use
Clone or download this repository.
Open the Excel workbook in Microsoft Excel (2016 or later recommended).
Review the Cleaned Data sheet, then the analysis sheets and the dashboard.
Use the slicers on the dashboard to filter by year, month, website and consultation type.
License

Add your preferred license here (for example MIT).
