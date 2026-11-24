# QC Pipeline Ops console

The console runs in the QC server's Windows logon session. Start it with the existing logon shortcut or launch it explicitly:

~~~powershell
powershell.exe -STA -NoProfile -File .\scripts\service\Start-QCOpsConsole.ps1
~~~

Use `-AppSettingsPath` to select a different profile. `-NoGui` opens the existing text console. Closing the window leaves the Session 0 scheduled pipeline running.

- **Overview** shows watcher and worker health, audit watermark age, oldest pending file, lane holds, and reconciliation progress and schedule.
- **Queue** opens on failures. Filter by state, type, or identifying text. Select a row for the JSON and any requeue restriction. Held lane records include their durable recovery data. See the [prepend recovery procedure](../engineering/remote-worker-host-guardrails.md) before retrying an uncertain upload. The grid loads up to 250 recent files per state plus deferred Status Set work.
- **Logs** opens newest-first. Reverse the order, pause live updates, or filter by component, severity, time, and text. Selecting a row shows the original multiline message and event fields. Refresh reads up to 250 lines from each of the 12 latest files per component.
- **History** shows audit polls, reconciliations, audit events, job outcomes, delivery failures, and automation errors. Use **Explore in SQL** for more rows and filters.
- **Sheets** shows package search results, a lane PDF matrix link, and state comparisons between ProjectWise and SQL. Prefer a selected sheet package when the same stem exists in several folders.
- **Actions** shows configured project and folder presets, supports custom folder paths, and presents script-specific controls. Use **Preview** before **Apply** where offered. Verify the displayed profile and arguments before applying. Output and elapsed time appear in the window. A timeout ends the child process. Check live state before retrying a timed-out write.
- **SQL Explorer** discovers queryable application tables and views and highlights curated datasets. Choose filters, columns, sort direction and page size, then select **Load results**. Paging and export reuse those filters. CSV export includes the selected columns and up to 5,000 matching rows, and prefixes spreadsheet formula-like text with an apostrophe. A selected row shows all returned fields, and identifiers can open related datasets.

Grid times display in Arizona wall time. CSV date and time fields retain ISO 8601 representation. Window dimensions, selected filters, and SQL column order are stored locally in `%LOCALAPPDATA%\ProjectWiseQC\ops-console.json`.

## Manual watcher and audit cursor

**Request full scan** sends an operations request to the active watcher. **Run one watcher pass** is available only on the QC server when its scheduled task is disabled and stopped. It invokes the normal watcher with `-Once`, so continuous polling in the selected profile cannot turn it into a loop. The pass can run due reconciliation and enqueue jobs. All watcher invocations use an exclusive queue handle at `_watcher/active-watcher.lock`; the file can remain after the process exits. Deploy and restart the scheduled watcher so its running process uses the same lease as manual invocations.

The **Rewind audit watermark** dialog in Actions requires a stopped pipeline and a typed confirmation. It holds the watcher lease while rewinding the SQL cursor to an earlier instant.

## Custom SELECT

Set `QC_OPS_SQL_READONLY_CONNECTION_STRING` in the environment inherited by the interactive operator console to enable custom queries. Use a dedicated SQL login granted only SELECT on the intended QC application objects. The custom editor uses that connection exclusively. It accepts one SELECT that can be used as a derived table, with named, unique result columns. SQL takes care of enforcing login permissions. The console caps output at 250 rows and 64 KiB per text or binary field where SQL Server applies TEXTSIZE. Connection timeout is five seconds and command timeout is ten seconds.

SQL supplies telemetry and the package index. Execution state remains in the JSON queue.
