## Health Insurance Member Onboarding Automation System
Excel + VBA · End-to-End Member Registration & Reporting Automation

Overview
A production-grade, VBA-powered Excel system that automates the full member onboarding lifecycle for a health insurance operation — from initial registration through to PDF report generation.
Built from a real operational problem: a fragmented, manual registration process that was consuming significant team hours each month. After deployment, manual processing time dropped by 60%.
The system replaces ad-hoc spreadsheets and manual data entry with a structured, validated, auditable workflow — all inside Excel, with no external dependencies.

The Problem
BeforeAfterManual data entry across disconnected sheetsSingle structured registration form with dropdownsNo duplicate checking — re-registrations slipped throughAuto duplicate detection on ID number at point of entryMember IDs assigned manuallySequential IDs auto-generated on registrationNo audit trailEvery action logged with timestamp and Windows userMonthly report built manuallyOne-click PDF export to desktopStatus updates done ad hocValidated status pipeline with colour-coded tracking

Workbook Architecture
ADH_Member_Onboarding_System.xlsx
│
├── DASHBOARD              ← Live KPI cards, pipeline status, recent registrations
├── NEW MEMBER REGISTRATION ← Structured intake form with data validation & macros
├── MEMBER_DB              ← Master member database with auto-filter & status colours
├── EMPLOYER_GROUPS        ← Corporate client registry with contract tracking
├── REPORTS                ← Auto-calculated monthly summary with PDF export button
└── VBA MODULE             ← Full macro code — copy into VBA editor to activate

Features
Registration Form

Structured 4-section intake form (Personal Info, Address, Plan & Cover, Emergency Contact)
Dropdown validation for Gender, Marital Status, Plan Type, Benefit Option, Premium Category, Nationality
Required field validation — registration blocked if mandatory fields are empty
Auto-generated Member ID displayed live on the form

Automation (VBA Macros)

RegisterMember — validates inputs, checks for duplicate ID numbers, writes to MEMBER_DB, applies formatting, logs the action, offers to clear form
ClearRegistrationForm — resets all input cells, returns cursor to first field
UpdateMemberStatus — finds member by ID, validates new status, updates DB and colour-coding, logs change
SearchMember — searches by Member ID or full name, displays summary popup
GeneratePDFReport — exports the Reports sheet as a timestamped PDF to desktop
FormatDBRow / FormatStatusCell — reusable helpers for consistent styling
LogActivity — writes every action to an ACTIVITY_LOG sheet (auto-created on first use)

Dashboard

Live KPI cards: Total Members, Active, Pending, New This Month, Employer Groups, Family Plans
Pipeline tracker showing member count at each onboarding stage with RAG indicators
Recent registrations panel pulling the latest entries from MEMBER_DB

Member Database

16-column structured database with auto-filter and frozen header row
Colour-coded status column: Active (green), Pending/Application Received (amber), Suspended (red), In-progress stages (blue)
Alternating row shading for readability

Reporting

Auto-calculated executive summary pulling live counts from MEMBER_DB
Efficiency metrics including processing time saved and pending SLA breach flags
One-click PDF export via GeneratePDFReport macro
