# Swiggy Rider Recruitment CRM

A comprehensive CRM system for recruiting Swiggy delivery riders and tracking their 31-delivery completion within 7 days.

## Overview

This CRM helps recruitment teams:
- **Call & recruit** potential delivery riders
- **Track delivery progress** (31 deliveries in 7 days target)
- **Auto-remove riders** who fail to complete 31 deliveries within the 7-day window
- **Manage agents** with break/washroom/meeting timers
- **Generate daily reports** automatically at 6:30 PM IST

## Features

### Agent Panel (`/agent`)
- Auto-dialer with disposition system
- Manual lead addition (no file upload required)
- Recruit riders with: Name, Phone, Vehicle Type, Area, City
- Track delivery count for recruited riders
- Break, Washroom, Meeting timers
- Leaderboard & Rankings
- Daily PDF report generation

### Admin Panel (`/admin`)
- Upload number sheets (.xlsx/.csv) for bulk dialing
- Real-time agent monitoring (active, break, meeting, washroom)
- Recruited riders panel with delivery tracking (X/31)
- 7-day countdown timer per rider
- Auto-failure marking after 7 days if < 31 deliveries
- Delivery Completed panel (31/31 riders)
- Delivery Failed panel (expired riders)
- EID management with roles (Agent, TL, Client, Admin)
- DND number management
- Call script upload
- Generated Reports Archive (auto-generated daily)

### Client Panel (`/client`)
- View disposition stats (Daily/Weekly/Monthly/Yearly)
- View recruited riders, followups, converted customers
- Download generated PDF reports
- DND numbers list

## Disposition System
- **Interested** → Rider Recruited (enters 7-day delivery tracking)
- **Followup** → Call back later (max 2 followups)
- **Not Interested** → Permanent removal
- **CNC (Dead)** → Call Not Connected (retry next day, max 2x)
- **CNR** → Call Not Received (retry next day, max 2x)
- **Switch Off** → Phone switched off (retry next day, max 2x)
- **Not-Eligible (Discard)** → Permanent removal
- **DND** → Do Not Disturb (permanent, never dial again)

## Delivery Tracking Flow
1. Agent marks lead as "Interested" → Rider is recruited
2. 7-day countdown starts from recruitment date
3. Agent calls rider daily and updates delivery count
4. When count reaches 31 → Auto-marked as "Delivery Completed"
5. If 7 days pass and count < 31 → Auto-marked as "Failed"

## Setup

```bash
npm install
npm start
```

Server runs on port 3000 by default.

- Admin Panel: http://localhost:3000/admin
- Agent Panel: http://localhost:3000/agent
- Client Panel: http://localhost:3000/client

## Default Login
- Client Panel: EID `9000` (username: "swiggy")

## Data Persistence
Data is stored in `~/.autolead-crm/` by default (outside the project folder).
Override with `AUTOLEAD_DATA_DIR` environment variable.

## Tech Stack
- Node.js + Express
- Socket.IO (real-time updates)
- PDFKit (report generation)
- XLSX (spreadsheet parsing)
- No database required (JSON file storage)
