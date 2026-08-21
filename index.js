// ---------- Element references ----------
const fileInput = document.getElementById("excelFile");
const dropzone = document.getElementById("dropzone");
const generateBtn = document.getElementById("generateBtn");
const genHint = document.getElementById("genHint");

const ledgerWrap = document.getElementById("ledgerWrap");
const ledgerBody = document.getElementById("ledgerBody");
const ledgerCount = document.getElementById("ledgerCount");
const ledgerMore = document.getElementById("ledgerMore");

const progressWrap = document.getElementById("progressWrap");
const progressFill = document.getElementById("progressFill");
const progressText = document.getElementById("progressText");
const progressPct = document.getElementById("progressPct");

const doneWrap = document.getElementById("doneWrap");
const doneTitle = document.getElementById("doneTitle");
const doneSub = document.getElementById("doneSub");

const certificate = document.getElementById("certificate");
const certName = document.getElementById("certName");
const certAge = document.getElementById("certAge");
const certAmount = document.getElementById("certAmount");
const certDate = document.getElementById("certDate");

const today = new Date().toLocaleDateString("en-IN", {
  day: "2-digit",
  month: "short",
  year: "numeric"
});

let parsedRecords = [];
let validCount = 0;

// ---------- Helpers ----------
function isValid(record) {
  const name = String(record.Name || "").trim();
  const age = String(record.Age || "").trim();
  const amount = Number(record.Amount);
  return !!name && !!age && Number.isFinite(amount) && amount > 0;
}

function renderLedger(records) {
  ledgerBody.innerHTML = "";
  validCount = 0;
  const shown = records.slice(0, 8);

  shown.forEach((record, i) => {
    const ok = isValid(record);
    if (ok) validCount++;

    const tr = document.createElement("tr");
    if (!ok) tr.classList.add("row-invalid");

    const amt = Number(record.Amount);
    tr.innerHTML = `
      <td class="num">${String(i + 1).padStart(2, "0")}</td>
      <td>${String(record.Name || "—")}</td>
      <td class="num">${String(record.Age || "—")}</td>
      <td class="amt">${Number.isFinite(amt) ? "₹" + amt.toLocaleString("en-IN") : "—"}</td>
      <td>${ok ? '<span class="stamp ok">approved</span>' : '<span class="stamp skip">skipped</span>'}</td>
    `;
    ledgerBody.appendChild(tr);
  });

  // count full valid total across all records, not just shown
  validCount = records.filter(isValid).length;

  ledgerCount.textContent = `${records.length} row${records.length === 1 ? "" : "s"} · ${validCount} valid`;
  ledgerMore.textContent = records.length > shown.length
    ? `+ ${records.length - shown.length} more row${records.length - shown.length === 1 ? "" : "s"} not shown`
    : "";
  ledgerWrap.style.display = "block";
}

async function handleFile(file) {
  if (!file) return;

  dropzone.classList.add("has-file");
  dropzone.querySelector(".drop-title").textContent = file.name;
  dropzone.querySelector(".drop-sub").innerHTML =
    `<span class="filechip"><span class="dot"></span>Ready to parse</span>`;

  try {
    const data = await file.arrayBuffer();
    const workbook = XLSX.read(data, { type: "array" });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    parsedRecords = XLSX.utils.sheet_to_json(sheet, { defval: "" });

    if (!parsedRecords.length) {
      genHint.textContent = "That file has no rows — try another.";
      generateBtn.disabled = true;
      ledgerWrap.style.display = "none";
      return;
    }

    renderLedger(parsedRecords);
    generateBtn.disabled = validCount === 0;
    genHint.textContent = validCount
      ? `Ready — ${validCount} certificate${validCount === 1 ? "" : "s"} will be generated`
      : "No valid rows found — check Name, Age, and Amount columns";
  } catch (err) {
    console.error(err);
    genHint.textContent = "Couldn't read that file — is it a valid .xlsx?";
    generateBtn.disabled = true;
  }
}

// ---------- Upload interactions ----------
dropzone.addEventListener("click", () => fileInput.click());

dropzone.addEventListener("keydown", e => {
  if (e.key === "Enter" || e.key === " ") fileInput.click();
});

fileInput.addEventListener("change", () => handleFile(fileInput.files[0]));

["dragenter", "dragover"].forEach(evt =>
  dropzone.addEventListener(evt, e => {
    e.preventDefault();
    dropzone.classList.add("drag");
  })
);

["dragleave", "drop"].forEach(evt =>
  dropzone.addEventListener(evt, e => {
    e.preventDefault();
    dropzone.classList.remove("drag");
  })
);

dropzone.addEventListener("drop", e => {
  const file = e.dataTransfer.files[0];
  if (file) {
    fileInput.files = e.dataTransfer.files;
    handleFile(file);
  }
});

// ---------- Generate certificates ----------
generateBtn.addEventListener("click", async () => {
  if (!parsedRecords.length) return;

  generateBtn.disabled = true;
  doneWrap.classList.remove("active");
  progressWrap.classList.add("active");
  progressFill.style.width = "0%";

  const zip = new JSZip();
  const total = parsedRecords.length;
  let generated = 0, skipped = 0;

  for (let i = 0; i < total; i++) {
    const record = parsedRecords[i];
    const name = String(record.Name || "").trim();
    const age = String(record.Age || "").trim();
    const amount = Number(record.Amount);

    if (!name || !age || !Number.isFinite(amount) || amount <= 0) {
      skipped++;
    } else {
      certName.textContent = name;
      certAge.textContent = age;
      certAmount.textContent = "₹" + amount.toLocaleString("en-IN");
      certDate.textContent = today;

      const copy = certificate.cloneNode(true);
      copy.removeAttribute("id");
      copy.classList.add("pdf-copy");
      copy.style.display = "block";
      document.body.appendChild(copy);

      // let the browser paint before capturing
      await new Promise(resolve =>
        requestAnimationFrame(() => requestAnimationFrame(resolve))
      );

      const canvas = await html2canvas(copy, {
        scale: 2,
        backgroundColor: "#fffdf6",
        useCORS: true,
        width: copy.scrollWidth,
        height: copy.scrollHeight
      });

      document.body.removeChild(copy);

      const { jsPDF } = window.jspdf;
      const pdf = new jsPDF({ orientation: "landscape", unit: "mm", format: "a4" });

      const pageWidth = 297, pageHeight = 210, margin = 10;
      const maxWidth = pageWidth - margin * 2;
      const maxHeight = pageHeight - margin * 2;
      const aspect = canvas.height / canvas.width;

      let imgWidth = maxWidth;
      let imgHeight = imgWidth * aspect;
      if (imgHeight > maxHeight) {
        imgHeight = maxHeight;
        imgWidth = imgHeight / aspect;
      }

      const x = (pageWidth - imgWidth) / 2;
      const y = (pageHeight - imgHeight) / 2;

      pdf.addImage(canvas.toDataURL("image/jpeg", 0.98), "JPEG", x, y, imgWidth, imgHeight);

      const safeName = name.replace(/[^a-z0-9]/gi, "-").replace(/-+/g, "-");
      const pdfBlob = pdf.output("blob");
      zip.file(`Donation-Certificate-${safeName}.pdf`, pdfBlob);

      generated++;
    }

    const pct = Math.round(((i + 1) / total) * 100);
    progressFill.style.width = pct + "%";
    progressPct.textContent = pct + "%";
    progressText.textContent = `Generated ${generated} of ${total} · ${skipped} skipped`;
  }

  progressText.textContent = `Bundling ${generated} certificate${generated === 1 ? "" : "s"} into a ZIP…`;
  const zipBlob = await zip.generateAsync({ type: "blob" });

  const link = document.createElement("a");
  link.href = URL.createObjectURL(zipBlob);
  link.download = "Donation-Certificates.zip";
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);

  progressWrap.classList.remove("active");
  doneWrap.classList.add("active");
  doneTitle.textContent = `${generated} certificate${generated === 1 ? "" : "s"} generated`;
  doneSub.textContent = skipped
    ? `Downloaded as one ZIP · ${skipped} row${skipped === 1 ? "" : "s"} skipped for missing data.`
    : "Downloaded as one ZIP file.";

  generateBtn.disabled = false;
});