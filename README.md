# Datum Fieldnote – Excel Add-in Metadata Inspector

**Datum Fieldnote** is a tool in progress Excel task pane add-in that provides a clean, intuitive interface for inspecting metadata about selected cells in your workbook. The add-in connects with a designated log sheet and surfaces change history, values, and notes—all directly within a sidebar panel.

Built using the [Excel JavaScript API](https://learn.microsoft.com/en-us/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview), this tool streamlines version tracking and supports better spreadsheet documentation.

---

## Features

- **Metadata Inspector** – Instantly displays the selected cell's address, value, and any change logs.
- **Log Sheet Integration** – Pulls metadata from a `Log` sheet to show previous value, new value, user, timestamp, and notes.
- **Go to Log Entry** – Jumps to the matching row in the log sheet and highlights it.
- **Table & Chart Creation** – Bonus tools for quickly creating a formatted expense table and a chart.

---

## How to Run This Project

### Prerequisites

- **Node.js** (latest LTS version) – [Install Node.js](https://nodejs.org/)
- **Microsoft 365 with Excel Desktop** – A subscription that supports Office Add-ins. Try [Microsoft 365 E5 Developer](https://developer.microsoft.com/microsoft-365/dev-program) or [start a free trial](https://www.microsoft.com/microsoft-365/try?rtc=1).

---

### Steps to Launch

1. **Open the Office Add-ins Development Kit** in VS Code.
2. Click **Preview Your Office Add-in (F5)** → select **Excel Desktop (Edge Chromium)**.
3. Excel will launch and sideload the add-in.
4. Select any cell in your workbook (on any sheet except `Log`) to view change logs in the sidebar.

> Make sure you have a sheet named **Log** with columns like `Timestamp`, `Cell`, `Sheet`, `Previous Value`, `New Value`, `User`, and `Notes`.

---

## Project Structure

| File | Purpose |
|------|---------|
| `manifest.xml` | Describes the add-in’s metadata and permissions |
| `src/taskpane/taskpane.html` | The layout of the Metadata Inspector |
| `src/taskpane/taskpane.css` | Styling for the task pane |
| `src/taskpane/taskpane.js` | Core logic: Excel integration, event handling, UI updates |

---

## Development Notes

- The metadata inspector updates every time a new cell is selected.
- If the selected cell matches an entry in the `Log` sheet, its change history is displayed.
- You can customize or expand the metadata columns used in the `Log` sheet as needed.

---

## Troubleshooting

- Confirm your workbook includes a sheet named `Log`.
- Close all instances of Excel before restarting.
- Use the **Stop Previewing Add-in** command in VS Code to reset sideloading.
- See the [Office Dev Troubleshooting Guide](https://learn.microsoft.com/office/dev/add-ins/testing/troubleshoot-development-errors) for more help.

---

## Resources

- [Office Add-ins Documentation](https://learn.microsoft.com/office/dev/add-ins/)
- [Validate Manifest File](https://learn.microsoft.com/office/dev/add-ins/testing/validate-your-office-add-in-manifest)
- [Office Add-ins Community Call](https://learn.microsoft.com/office/dev/add-ins/overview/office-add-ins-community-call)

---

## License

© 2024 Microsoft Corporation. All rights reserved.

**Disclaimer**: *This code is provided “as-is” without any warranties. Use at your own risk.*