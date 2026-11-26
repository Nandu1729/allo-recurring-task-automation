# Allo Innoware recurring task automation

AI Operations / Founder’s Office assignment · Nanda Kishore

The prototype gives employees one place to see and complete their recurring tasks. The manager uses a single queue to acknowledge overdue work and track resolution. Leadership receives an escalation when review or resolution misses its deadline.

| File | Contents |
| --- | --- |
| [Assignment and event evidence](Allo%20Innoware%20Assignment%20and%20Event%20Evidence.pdf) | Original assignment brief, results for all eight tests, the simulated event timeline, and the live email verification. |
| [Recurring operations workbook](Allo%20Innoware%20-%20Recurring%20Operations.xlsx) | Task master, four employee queues, manager queue, alert log, settings, and demo walkthrough. |
| [Apps Script](Allo%20Ops.gs) | Cycle generation, proof capture, manager acknowledgement, and email escalation. |
| [Local test harness](test.mjs) | Dependency-free Node.js checks using mock Sheets, trigger, and email services. |

## Run the demo

1. Upload the workbook to Drive and save it as a native Google Sheet.
2. Open **Extensions → Apps Script**, replace `Code.gs` with the supplied source, save, and reload the sheet.
3. Choose **Allo Ops → Initialize / verify workbook**. The demo starts at **26 Nov 2026, 16:45 IST**. Follow **Demo & Tests**. When instructed to advance time, edit `Settings!B9` and run **Allo Ops → Refresh queues and check alerts**.
4. Enter a portal Ticket ID or `POSTED` for the WhatsApp task in an employee row, then choose **Submit proof for selected row**. The manager acknowledges separately in **Manager Queue**.

The demo uses simulated time and records email payloads in **Alert Log**. Employee A’s WhatsApp task recurs on weekdays. The portal tasks recur 7, 15, or 30 calendar days after the last completion date.

The first check at or after **17:00 IST** alerts the manager. Leadership receives one escalation per open cycle if the manager misses the **60-minute** acknowledgement deadline or if the task remains open after **180 minutes**. Both windows start from the manager alert attempt. **Start Here** covers live setup and the five-minute polling trigger.

## Verification

All eight assignment scenarios passed in native Google Sheets. A separate live run delivered exactly one manager alert and one leadership alert to a controlled inbox. The test trigger was removed after delivery. The PDF appendix records the demo and live test events.

Run the mock-based checks locally with `node test.mjs`.
