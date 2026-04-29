# ElderCare Architecture

ElderCare is a mobile-first PWA and Telegram bot system for tracking an elderly parent's daily care tasks, medication schedules, supplement reordering, bill payments, food expiry, and health metrics — with Google Sheets as the data layer and Google Apps Script as the backend and host.

```mermaid
flowchart TD
    User["Caregiver / Elder\n(Mobile Browser)"]
    PWA["PWA Frontend\n(Index.html + App.html + Styles.html)\nHosted on Apps Script Web App"]
    Backend["Apps Script Backend\n(Code.gs)\ndoGet · getTodayTasks · markTaskDone\ngetSupplies · markOrdered\ngetMedications · confirmMedication · snoozeMedication\ngetBills · markBillPaid\ngetFoodItems · addFoodItem\nlogHealth · getHealthTrends"]
    Sheet["Google Sheets\n(Data Layer)"]
    TASKS["TASKS\n(daily/weekly/monthly/once)"]
    SUPPLIES["SUPPLIES\n(qty · daily usage · reorder date)"]
    MEDS["MEDS + MED_LOG\n(schedules · confirmations)"]
    BILLS["BILLS\n(due dates · recurrence)"]
    FOOD["FOOD_LOG\n(cooked · expiry · consumed)"]
    HEALTH["HEALTH_LOG\n(energy · sleep · mood · focus)"]
    SETTINGS["SETTINGS\n(elder_name · app_title)"]
    LOG["LOG\n(audit trail)"]
    Telegram["Telegram.gs\n(Scheduled Triggers)"]
    TelegramAPI["Telegram Bot API"]
    ElderPhone["Elder's Phone\n(Telegram)"]
    CarerPhone["Carer's Phone\n(Telegram)"]

    User -->|"HTTPS"| PWA
    PWA -->|"google.script.run"| Backend
    Backend -->|"SpreadsheetApp"| Sheet
    Sheet --- TASKS
    Sheet --- SUPPLIES
    Sheet --- MEDS
    Sheet --- BILLS
    Sheet --- FOOD
    Sheet --- HEALTH
    Sheet --- SETTINGS
    Sheet --- LOG

    Telegram -->|"reads same Sheet"| Sheet
    Telegram -->|"UrlFetchApp POST"| TelegramAPI
    TelegramAPI --> ElderPhone
    TelegramAPI --> CarerPhone

    Backend -->|"_decrementMedSupply\nlinks MEDS → SUPPLIES"| Sheet

    subgraph "Scheduled Triggers (Apps Script)"
        T1["dailyMorningAlert\n(tasks + supply alerts → carer)"]
        T2["checkDueMedicationReminders\n(med due → elder)"]
        T3["checkEscalations\n(no confirm >30 min → carer)"]
        T4["checkDueBills\n(bill due within remind window → group)"]
        T5["weeklySummaryAlert\n(health trend → carer)"]
    end

    Telegram --- T1
    Telegram --- T2
    Telegram --- T3
    Telegram --- T4
    Telegram --- T5
```

## Component Summary

| Component | Role |
|---|---|
| `phase1/setup.gs` | One-time setup: creates all Sheet tabs with headers and seed data |
| `phase2/Code.gs` | Apps Script backend — serves PWA via `doGet`, exposes all data API functions |
| `phase2/Index.html` | PWA shell with tab navigation (Tasks, Supplies, Meds, Bills, Food, Health) |
| `phase2/App.html` | Frontend JS — calls `google.script.run`, renders UI, handles swipe-to-complete |
| `phase2/Styles.html` | Mobile-first CSS, swipe animations, status colour coding |
| `phase3/Telegram.gs` | Scheduled trigger handlers that read the Sheet and push alerts via Telegram Bot API |
| Google Sheets | Persistent data store — TASKS, SUPPLIES, MEDS, MED_LOG, BILLS, FOOD_LOG, HEALTH_LOG, SETTINGS, LOG |
| Telegram Bot API | Delivers push notifications to elder's phone and carer's phone |

## Data Flow — Medication Reminder

```mermaid
sequenceDiagram
    participant Trigger as Apps Script Trigger
    participant Telegram as Telegram.gs
    participant Sheet as Google Sheets
    participant Bot as Telegram Bot API
    participant Elder as Elder's Phone
    participant PWA as PWA (Carer)

    Trigger->>Telegram: checkDueMedicationReminders() every 10 min
    Telegram->>Sheet: Read MEDS + MED_LOG
    Telegram->>Bot: sendMessage → elder_chat_id
    Bot->>Elder: "Time to take Metformin"
    Elder-->>PWA: Opens app
    PWA->>Sheet: confirmMedication(eventId) via google.script.run
    Sheet-->>PWA: { success: true }
    Note over Telegram,Sheet: If no confirm in 30 min → checkEscalations() → carer alert
```
