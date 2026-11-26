# Allo Innoware recurring task automation

A Google Sheets and Apps Script prototype for tracking recurring tasks, recording completion, and escalating overdue work.

| File | Contents |
| --- | --- |
| [Assignment and event evidence](Allo%20Innoware%20Assignment%20and%20Event%20Evidence.pdf) | Original assignment brief, results for all eight tests, the simulated event timeline, and the live email verification. |
| [Recurring operations workbook](Allo%20Innoware%20-%20Recurring%20Operations.xlsx) | Task master, four employee queues, manager queue, alert log, settings, and demo walkthrough. |
| [Apps Script](Allo%20Ops.gs) | Cycle generation, proof capture, manager acknowledgement, and email escalation. |
| [Local test harness](test.mjs) | Dependency-free Node.js checks using mock Sheets, trigger, and email services. |

## Run the demo

1. Upload the workbook to Drive and save it as a native Google Sheet.
2. Open **Extensions → Apps Script**, replace `Code.gs` with the supplied source, save, and reload the sheet.
3. Choose **Allo Ops → Initialize / verify workbook**. Follow the **Demo & Tests** tab from the starting clock of **26 Nov 2026, 16:45 IST**. Edit `Settings!B9` to advance time.
4. Enter a portal Ticket ID or `POSTED` for the WhatsApp task in an employee row, then choose **Submit proof for selected row**. The manager acknowledges separately in **Manager Queue**.

The demo uses simulated time and records email payloads in **Alert Log**. The configured deadline is 17:00 IST. Leadership is notified once if an open task misses the 60-minute acknowledgement window or stays open after 180 minutes. The **Start Here** tab covers live setup and the five-minute polling trigger.

## Verification

All eight assignment scenarios passed in native Google Sheets. A separate live run delivered exactly one manager alert and one leadership alert to a controlled inbox. The test trigger was removed after delivery. The PDF appendix records the demo and live test events.

Run the mock-based checks locally with `node test.mjs`.
