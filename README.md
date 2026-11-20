# Daily Sales Loader — case study

A Windows desktop tool that loads an online store's daily sales into the company's accounting spreadsheet and reconciles the whole month against the store's export.

**Result:** daily sales entry went from 30–40 minutes to 1–2 minutes. In its first monthly close, the reconciliation found 58 sales that had never been recorded (about 6% of that month's orders) and around 15 unrecorded refunds. The owners validated the reconciliation.

> This repository is a case study. The tool was built for a real business and is used in production, so the source code and data are private. Everything shown here uses made-up sample data and a generic store name. I'm happy to walk through the code in a call.

## The problem

Every day, someone copied the day's orders from the store admin into a shared Excel workbook by hand: order number, amount, and the amount again under the right payment-method column. It took 30–40 minutes and had two failure modes nobody could see:

1. **Late payments disappeared.** An order that was still unpaid on the day it was placed was skipped. If the customer paid three days later, nothing brought that order back, so it never reached the books.
2. **Typing errors passed the spreadsheet's own check.** The sheet had a control column, but a slip in a formula range could leave it reporting "ok" on every row.

## What it does

**Daily mode.** Copy the order list from the store admin, open the tool, and it builds the block of rows to paste into the spreadsheet:

- Reads the open workbook to find the next free row, the next sequence number, and the last order already loaded. The only thing the user confirms is the date.
- Keeps only paid orders and maps each payment method to its column.
- Marks in red anything it does not recognize. The "Copy block" button stays locked until the user has reviewed every red row.

**Month-end mode.** Load the store's monthly CSV export and the tool compares it with what is in the spreadsheet for that period:

- Lists every paid order that is missing and offers them as a block to paste.
- Handles refunds, including partial ones, as negative rows that point to the original order.
- Refuses to load a refund whose original sale is not in the spreadsheet, and says why.

### Example (sample data)

Input copied from the store admin:

| Order | Total | Payment status | Payment method |
|---|---|---|---|
| #1041 | 18,500 | Paid | Bank transfer |
| #1042 | 9,200 | Paid | Cash |
| #1043 | 31,000 | Pending | Bank transfer |
| #1044 | 12,750 | Paid | Gift voucher |

What the tool shows:

| Order | Result |
|---|---|
| #1041 | Ready: 18,500 under "Transfer" |
| #1042 | Ready: 9,200 under "Cash" |
| #1043 | Skipped: not paid yet. It will show up in the month-end check once paid |
| #1044 | **Red: unknown payment method.** Not loaded until a person decides |

## Design decisions

| Decision | Why |
|---|---|
| The tool never writes to the spreadsheet. It produces a block and a person pastes it | The workbook belongs to the company. A program that cannot write to it cannot corrupt it |
| It never guesses. An unknown payment method or status is flagged, not mapped to the closest match | A wrong guess looks exactly like a correct row. A red row gets looked at |
| A single `.exe`, compiled with the C# compiler that ships with Windows | The office machines had no Node, Python, or .NET SDK, and the tool had to be handed to coworkers as one file with nothing to install |
| Formulas are copied from an existing row and re-pointed, not hard-coded | The tool keeps working if the company changes a formula |
| Two modes with different inputs: clipboard for the day, CSV export for the month | The daily flow stays fast; the monthly check catches what the daily flow cannot see by design |
| The month-end check only looks at rows dated inside the period | Older rows in the same column used a different numbering and could collide with order numbers |
| Rollout by parallel run: old script and new tool side by side, comparing output, before switching | The only way to show the new tool does the same thing, not just that it does not crash |

## Architecture

Five parts, with the side effects kept at the edges:

| Part | Responsibility | Touches Excel | Touches the UI |
|---|---|---|---|
| Store reader | Pasted text or CSV → list of sales | No | No |
| Spreadsheet reader | Sheet → loaded sales, last row, last sequence number, formulas | Yes | No |
| Reconciler | The two lists → what is missing | No | No |
| Block builder | Sales to load → tab-separated text ready to paste | No | No |
| Window | Wires the other four together | No | Yes |

Three of the five parts touch neither Excel nor the screen, so they are covered by automated tests that run without the workbook.

Lessons from working with Excel through COM:

- Open workbooks with `IgnoreReadOnlyRecommended`, or `Workbooks.Open` can return null without raising an error.
- Read ranges in one call into an array. The sheet has about 16,000 rows, and cell-by-cell access is too slow.
- Never kill Excel processes by name: the others belong to the user and may hold unsaved work.

## Stack

C# · Windows Forms · .NET Framework 4.8 · Excel COM automation · CSV parsing · automated tests for the pure components

## About

Built by Facundo Chireno while running the online operation of an ecommerce store. It replaced an earlier PowerShell script I had written for the same job.
