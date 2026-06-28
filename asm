Bank Dashboard — Smart Upgrade Plan 🚀
সব features একসাথে implement করব। বর্তমান bank_dashboard_new_search_ready.py ফাইলটাই upgrade করব — কোনো নতুন ফাইল structure ভাঙব না।

কী কী করব
1. ⚠️ Smart Alerts & Anomaly Detection
Dashboard-এর শুরুতে একটা Alerts Panel থাকবে (collapsible)
Overdue LC alert: Maturity Date পেরিয়ে গেছে কিন্তু payment আসেনি → 🔴
Due Soon alert: আগামী ৭ দিনে maturity → 🟡
Sudden drop alert: কোনো firm/sales person এর submission হঠাৎ ৫০%+ কমে গেলে → 🟠
Sidebar-এ alert count badge দেখাবে
2. 📈 Forecasting (Linear Trend)
Overview tab-এ "Next Month Forecast" section যোগ হবে
Monthly data দিয়ে linear regression → আগামী মাসের predicted submission count ও invoice value
Confidence interval সহ chart
3. 🔍 Smart Search (Natural Language Filter)
Sidebar-এ একটা search box যোগ হবে
Pattern match: "unpaid DBBL June", "overdue ABC firm", "paid last week"
LC Number দিয়ে direct search → single record popup
4. 🗓️ Due Date Tracker Tab (নতুন tab)
নতুন "📅 Due Dates" tab
আগামী ৭ / ১৫ / ৩০ / ৬০ দিনের maturity LC গুলো grouped table
Color-coded: 🔴 Overdue | 🟡 Due in 7 days | 🟢 Safe
Countdown days column
Export to Excel/PDF
5. 🏆 Sales Person Leaderboard (নতুন section)
Firm & Sales Person tab-এ gamified leaderboard
Rank medals: 🥇🥈🥉 + star ratings
Month-over-month change (↑↓ arrows)
Target vs Actual progress bar (target = avg of last 3 months × 1.1)
6. 📧 Auto PDF Report (upgrade existing)
বর্তমান export button আছে → এটাকে smarter করব
Daily Summary PDF: KPIs + charts screenshot + top parties + overdue list
Weekly Report PDF: Leaderboard + trend chart + alerts summary
Email (optional): smtp config দিলে send করবে
7. 🔐 Login System
Session state দিয়ে simple login
Users: admin (সব দেখবে), manager (firm level), sales (নিজের data)
Credentials: .streamlit/users.toml ফাইলে রাখব
Logout button
Proposed Changes
[MODIFY] bank_dashboard_new_search_ready.py
Login system (শুরুতে)
Smart search box (sidebar)
Alerts panel (header area)
Forecasting section (Overview tab)
Due Date Tracker (নতুন tab)
Sales Leaderboard upgrade (Firm & Sales Person tab)
PDF report upgrade
[NEW] .streamlit/users.toml
User credentials config
Verification Plan
App run করব → login screen দেখাবে কিনা
admin দিয়ে login → সব tab দেখাবে কিনা
Smart search test → "DBBL" লিখলে filter হবে কিনা
Alerts panel → overdue records দেখাবে কিনা
Due Date tab → upcoming maturities দেখাবে কিনা
NOTE

Email feature optional থাকবে — SMTP config না দিলে শুধু PDF generate হবে।

IMPORTANT

Login system খুব simple হবে (session_state based) — production-grade নয়। Real security চাইলে পরে Firebase auth integrate করা যাবে।