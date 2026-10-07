"use strict";

(() => {
  const $ = id => document.getElementById(id);
  const ACCOUNT = "Cloud Account Number";
  const FRIENDLY = "Cloud Account Friendly Name";
  const PAGE_SIZE = 10;

  const state = {
    source: null,
    destination: null,
    mapping: null,
    output: null,
    page: 0,
    busy: false
  };

  function status(message, error = false) {
    $("status").textContent = message;
    $("status").className = error ? "status error" : "status";
  }

  function normalizeHeader(value) {
    return String(value ?? "").trim().toLowerCase().replace(/\s+/g, " ");
  }

  function accountKey(value) {
    const text = String(value ?? "").trim();
    if (!text) return "";
    if (!/^\d{1,12}$/.test(text)) {
      throw new Error(
        `Invalid account number "${text}". Use up to 12 digits, without punctuation.`
      );
    }
    return text.padStart(12, "0");
  }

  function isBlank(row) {
    return row.every(value => value === "" || value == null);
  }

  function invalidate() {
    state.output = null;
    state.page = 0;
    $("merge-results").hidden = true;
    $("download-output").disabled = true;
  }

  function setStep(id, complete, active) {
    $(id).className =
      "step" + (complete ? " complete" : active ? " active" : "");
  }

  function updateControls() {
    const ready = Boolean(window.XLSX);
    $("load-demo").disabled = !ready || state.busy;
    $("source-file").disabled = !ready || state.busy;
    $("destination-file").disabled = !ready || state.busy;
    $("source-sheet").disabled = state.busy;
    $("destination-sheet").disabled = state.busy;
    $("run-merge").disabled =
      state.busy || !state.mapping || !state.destination;

    setStep("step-source", Boolean(state.mapping), !state.mapping);
    setStep(
      "step-destination",
      Boolean(state.destination),
      Boolean(state.mapping) && !state.destination
    );
    setStep(
      "step-merge",
      Boolean(state.output),
      Boolean(state.mapping && state.destination) && !state.output
    );
  }

  function metrics(id, items) {
    const container = $(id);
    container.replaceChildren();

    for (const [label, value] of items) {
      const box = document.createElement("div");
      box.className = "metric";
      const small = document.createElement("small");
      small.textContent = label;
      const strong = document.createElement("strong");
      strong.textContent = value;
      box.append(small, strong);
      container.append(box);
    }
  }

  function table(id, headers, rows) {
    const element = document.createElement("table");
    const head = document.createElement("thead");
    const headerRow = document.createElement("tr");

    for (const value of headers) {
      const cell = document.createElement("th");
      cell.textContent = String(value ?? "");
      headerRow.append(cell);
    }

    head.append(headerRow);
    const body = document.createElement("tbody");

    for (const row of rows) {
      const tr = document.createElement("tr");
      for (let column = 0; column < headers.length; column++) {
        const cell = document.createElement("td");
        cell.textContent = String(row[column] ?? "");
        tr.append(cell);
      }
      body.append(tr);
    }

    element.append(head, body);
    $(id).replaceChildren(element);
  }

  function worksheetRows(sheet) {
    if (!sheet["!ref"]) throw new Error("This worksheet is empty.");

    const range = XLSX.utils.decode_range(sheet["!ref"]);
    if (range.e.r > 100000 || range.e.c > 499) {
      throw new Error("Please use a worksheet under 100,000 rows and 500 columns.");
    }

    const rows = XLSX.utils.sheet_to_json(sheet, {
      header: 1,
      range: {
        s: { r: 0, c: 0 },
        e: range.e
      },
      raw: true,
      defval: "",
      blankrows: true
    });

    return rows.map(row =>
      Array.from({ length: range.e.c + 1 }, (_, index) => row[index] ?? "")
    );
  }

  function selectedData(kind) {
    const file = state[kind];
    const name = $(`${kind}-sheet`).value;
    const sheet = file.book.Sheets[name];
    return { name, sheet, rows: worksheetRows(sheet) };
  }

  function inspectSource() {
    state.mapping = null;
    $("source-metrics").replaceChildren();
    $("source-preview").replaceChildren();

    const { rows } = selectedData("source");
    const headers = rows[0].map(normalizeHeader);
    const accountColumns = headers
      .map((value, index) => value === normalizeHeader(ACCOUNT) ? index : -1)
      .filter(index => index >= 0);
    const friendlyColumns = headers
      .map((value, index) => value === normalizeHeader(FRIENDLY) ? index : -1)
      .filter(index => index >= 0);

    if (accountColumns.length !== 1 || friendlyColumns.length !== 1) {
      throw new Error(
        `Source row 1 must contain exactly one "${ACCOUNT}" and one "${FRIENDLY}" header.`
      );
    }

    const map = new Map();
    let duplicates = 0;
    let total = 0;

    for (let index = 1; index < rows.length; index++) {
      const row = rows[index];
      if (isBlank(row)) continue;
      total++;

      const key = accountKey(row[accountColumns[0]]);
      const name = String(row[friendlyColumns[0]] ?? "").trim();

      if (!key || !name) {
        throw new Error(`Source row ${index + 1} needs an account number and friendly name.`);
      }

      if (map.has(key)) {
        duplicates++;
        if (map.get(key) !== name) {
          throw new Error(
            `Account ${key} has conflicting friendly names. Correct the source file first.`
          );
        }
      } else {
        map.set(key, name);
      }
    }

    if (!map.size) throw new Error("No source mappings were found.");
    state.mapping = map;

    metrics("source-metrics", [
      ["Total Source Rows", total],
      ["Unique Mappings Loaded", map.size],
      ["Duplicate Source Rows", duplicates]
    ]);

    table(
      "source-preview",
      ["Normalized Account Number", FRIENDLY],
      Array.from(map.entries()).slice(0, 5)
    );
  }

  function inspectDestination() {
    $("destination-metrics").replaceChildren();
    const { rows } = selectedData("destination");

    if (normalizeHeader(rows[0][0]) !== normalizeHeader(ACCOUNT)) {
      throw new Error(`Destination cell A1 must be "${ACCOUNT}".`);
    }

    if (rows[0].some(value => normalizeHeader(value) === normalizeHeader(FRIENDLY))) {
      throw new Error("This destination already has a friendly-name column. Use the original file.");
    }

    const accounts = rows.slice(1)
      .filter(row => !isBlank(row) && String(row[0]).trim() !== "")
      .map(row => accountKey(row[0]));
    const unique = new Set(accounts);

    metrics("destination-metrics", [
      ["Original Columns", rows[0].length],
      ["Destination Rows (including blank rows)", rows.length - 1],
      ["Unique Accounts", unique.size],
      ["Repeated Account Rows", accounts.length - unique.size]
    ]);
  }

  function installFile(kind, book, filename) {
    state[kind] = { book, filename };
    $(`${kind}-filename`).textContent = filename;
    $(`${kind}-details`).hidden = false;

    const select = $(`${kind}-sheet`);
    select.replaceChildren();

    for (const name of book.SheetNames) {
      const option = document.createElement("option");
      option.value = name;
      option.textContent = name;
      select.append(option);
    }

    if (kind === "source") inspectSource();
    else inspectDestination();
  }

  async function upload(kind, file) {
    if (!file || state.busy || !window.XLSX) return;

    invalidate();
    state[kind] = null;
    if (kind === "source") state.mapping = null;
    $(`${kind}-details`).hidden = true;
    state.busy = true;
    updateControls();
    status(`Reading ${file.name}…`);

    try {
      if (!/\.(xlsx|xls|csv)$/i.test(file.name)) {
        throw new Error("Choose an .xlsx, .xls, or .csv file.");
      }
      if (file.size > 25 * 1024 * 1024) {
        throw new Error("Please use a file smaller than 25 MB.");
      }

      const book = XLSX.read(await file.arrayBuffer(), {
        type: "array",
        raw: true,
        cellFormula: true
      });

      if (!book.SheetNames.length) throw new Error("No worksheets were found.");
      installFile(kind, book, file.name);
      status("File loaded. Check the selected worksheet, then run the merge.");
    } catch (error) {
      status(error.message, true);
    } finally {
      state.busy = false;
      updateControls();
    }
  }

  function renderOutput() {
    if (!state.output) return;
    const query = $("filter-rows").value.trim().toLowerCase();
    const rows = state.output.rows.slice(1).filter(row =>
      !query || row.some(value => String(value).toLowerCase().includes(query))
    );

    const pages = Math.max(1, Math.ceil(rows.length / PAGE_SIZE));
    state.page = Math.min(state.page, pages - 1);
    const start = state.page * PAGE_SIZE;

    table(
      "output-preview",
      state.output.rows[0],
      rows.slice(start, start + PAGE_SIZE)
    );

    $("preview-count").textContent = rows.length
      ? `Showing ${start + 1}–${Math.min(start + PAGE_SIZE, rows.length)} of ${rows.length} rows`
      : "No matching rows";
    $("page-number").textContent = `Page ${state.page + 1} of ${pages}`;
    $("previous-page").disabled = state.page === 0;
    $("next-page").disabled = state.page >= pages - 1;
  }

  function runMerge() {
    invalidate();

    try {
      if (!state.mapping || !state.destination) {
        throw new Error("Load both files first.");
      }

      inspectSource();
      inspectDestination();
      const { name, sheet, rows } = selectedData("destination");

      if (sheet["!merges"]?.length) {
        throw new Error("Merged cells are not supported. Use a plain data worksheet.");
      }

      for (const [address, cell] of Object.entries(sheet)) {
        if (address.startsWith("!")) continue;
        if (cell.f || cell.F) {
          throw new Error("Destination formulas are not supported. Use a values-only copy.");
        }
        if (cell.t === "e") {
          throw new Error(`Destination cell ${address} contains an Excel error.`);
        }
      }

      let matched = 0;
      let unmatched = 0;
      let blankNames = 0;

      const output = rows.map((row, index) => {
        let nameToInsert = "";
        if (index === 0) {
          nameToInsert = FRIENDLY;
        } else if (!isBlank(row)) {
          const key = accountKey(row[0]);
          if (key && state.mapping.has(key)) {
            nameToInsert = state.mapping.get(key);
            matched++;
          } else {
            unmatched++;
          }
          if (!nameToInsert) blankNames++;
        }
        return [row[0], nameToInsert, ...row.slice(1)];
      });

      const workbook = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(
        workbook,
        XLSX.utils.aoa_to_sheet(output),
        name
      );

      const bytes = XLSX.write(workbook, {
        bookType: "xlsx",
        type: "array"
      });

      const reread = XLSX.read(bytes, { type: "array" });
      const exported = worksheetRows(reread.Sheets[name]);

      const rowCheck = exported.length === rows.length;
      const columnCheck = exported.every(row => row.length === rows[0].length + 1);
      const valuesCheck = rowCheck && rows.every((row, r) =>
        row.every((value, c) => exported[r][c === 0 ? 0 : c + 1] === value)
      );
      const namesCheck = rowCheck && output.every((row, r) =>
        exported[r][1] === row[1]
      );

      const checks = [
        ["Original row count and order", rowCheck && valuesCheck],
        ["Original cell values preserved", valuesCheck],
        ["Exactly one new column inserted", columnCheck],
        ["Column B friendly names verified", namesCheck],
        ["Repeated account rows retained", rowCheck && valuesCheck]
      ];

      $("audit-checks").replaceChildren();
      for (const [label, passed] of checks) {
        const item = document.createElement("div");
        item.className = "audit-check" + (passed ? "" : " fail");
        item.textContent = `${passed ? "✓" : "✗"} ${label}`;
        $("audit-checks").append(item);
      }

      if (checks.some(([, passed]) => !passed)) {
        throw new Error("Export validation failed. Download has been blocked.");
      }

      state.output = { rows: output, bytes };
      $("filter-rows").value = "";

      metrics("audit-metrics", [
        ["Source Mappings", state.mapping.size],
        ["Destination Rows", rows.length - 1],
        ["Matched Destination Rows", matched],
        ["Unmatched Nonblank Rows", unmatched],
        ["Blank Friendly Names (nonblank rows)", blankNames],
        ["Columns (Original → Final)", `${rows[0].length} → ${rows[0].length + 1}`]
      ]);

      $("audit-message").className =
        "audit-message" + (unmatched ? " warning" : "");
      $("audit-message").textContent = unmatched
        ? `Value checks passed. ${unmatched} nonblank rows have no mapping; their friendly names are blank. Review before downloading.`
        : "All value and row checks passed. The updated worksheet is ready to download.";

      $("merge-results").hidden = false;
      $("download-output").disabled = false;
      renderOutput();
      status("Merge finished. Only the selected destination worksheet is exported.");
    } catch (error) {
      state.output = null;
      $("download-output").disabled = true;
      status(error.message, true);
    }
    updateControls();
  }

  function demoBook(rows, name) {
    const book = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(book, XLSX.utils.aoa_to_sheet(rows), name);
    return book;
  }

  function loadDemo() {
    invalidate();

    try {
      const mappings = [
        ["270161484344", "Account 5"],
        ["344888049787", "Account 7"],
        ["007179095593", "Core Billing Prod"],
        ["882910482019", "Dev Sandbox East"],
        ["192837465012", "Analytics Data Warehouse"],
        ["551928374019", "Security GuardDuty Ops"]
      ];

      const records = [
        ["270161484344", "Product A", "P001", 1, "us-east-1", 120.5],
        ["270161484344", "Product B", "P002", 2, "us-west-2", 340],
        ["344888049787", "Product C", "P003", 3, "eu-central-1", 85.2],
        ["270161484344", "Product D", "P004", 4, "us-east-1", 510],
        ["007179095593", "EC2 Reserved Instances", "P005", 12, "us-east-1", 1420],
        ["007179095593", "S3 Standard Storage", "P006", 50, "us-east-1", 230.75],
        ["882910482019", "Lambda Executions", "P007", 1000, "us-west-2", 45.1],
        ["192837465012", "Redshift Cluster", "P008", 2, "us-east-1", 3100],
        ["551928374019", "GuardDuty Threat Detection", "P009", 1, "us-east-1", 95],
        ["270161484344", "CloudWatch Logs", "P010", 25, "us-east-1", 110.25]
      ];

      $("source-file").value = "";
      $("destination-file").value = "";

      installFile(
        "source",
        demoBook([[ACCOUNT, FRIENDLY], ...mappings], "Unique Accounts"),
        "Demo_Source_Mapping.xlsx"
      );

      installFile(
        "destination",
        demoBook([
          [ACCOUNT, "Product Name", "Product ID", "Quantity", "Region", "Monthly Cost ($)"],
          ...records
        ], "Account Records"),
        "Demo_Destination_Records.xlsx"
      );

      runMerge();
    } catch (error) {
      status(error.message, true);
      updateControls();
    }
  }

  for (const kind of ["source", "destination"]) {
    $(`${kind}-file`).addEventListener("change", event => {
      upload(kind, event.target.files[0]);
    });

    $(`${kind}-sheet`).addEventListener("change", () => {
      invalidate();
      try {
        if (kind === "source") inspectSource();
        else inspectDestination();
        status("Worksheet selected. Run the merge to update the output.");
      } catch (error) {
        status(error.message, true);
      }
      updateControls();
    });

    const zone = $(`${kind}-drop`);
    zone.addEventListener("dragover", event => {
      event.preventDefault();
      zone.classList.add("dragover");
    });
    zone.addEventListener("dragleave", () => zone.classList.remove("dragover"));
    zone.addEventListener("drop", event => {
      event.preventDefault();
      zone.classList.remove("dragover");
      upload(kind, event.dataTransfer.files[0]);
    });
  }

  $("load-demo").addEventListener("click", loadDemo);
  $("run-merge").addEventListener("click", runMerge);

  $("filter-rows").addEventListener("input", () => {
    state.page = 0;
    renderOutput();
  });

  $("previous-page").addEventListener("click", () => {
    state.page = Math.max(0, state.page - 1);
    renderOutput();
  });

  $("next-page").addEventListener("click", () => {
    state.page++;
    renderOutput();
  });

  $("download-output").addEventListener("click", () => {
    if (!state.output) return;

    const blob = new Blob([state.output.bytes], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = "Updated_Cloud_Account_Friendly_Names.xlsx";
    document.body.append(link);
    link.click();
    link.remove();
    setTimeout(() => URL.revokeObjectURL(url), 60000);
  });

  document.querySelector(".format-note").textContent =
    "Use row 1 for headers. Account keys are normalized to 12 digits for matching; " +
    "original destination values remain unchanged. Export contains only the selected " +
    "worksheet's values. Formatting and other workbook features are not copied. " +
    "Destination formulas and merged cells are blocked.";

  updateControls();
  status(window.XLSX
    ? "Ready. Upload your files or click Load Demo Files."
    : "Spreadsheet library is missing. Add vendor/xlsx.full.min.js to finish setup.",
    !window.XLSX
  );
})();
