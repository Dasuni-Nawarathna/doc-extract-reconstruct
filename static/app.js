/**
 * app.js — Frontend logic for DocRecon web app.
 * Handles drag-and-drop file upload, conversion API calls,
 * progress states, and download.
 */

document.addEventListener("DOMContentLoaded", () => {
    // ── DOM Elements ──────────────────────────────────────────────
    const dropZone = document.getElementById("dropZone");
    const fileInput = document.getElementById("fileInput");
    const filePreview = document.getElementById("filePreview");
    const fileName = document.getElementById("fileName");
    const fileSize = document.getElementById("fileSize");
    const removeFile = document.getElementById("removeFile");
    const convertBtn = document.getElementById("convertBtn");
    const uploadCard = document.getElementById("uploadCard");
    const processingCard = document.getElementById("processingCard");
    const processingStatus = document.getElementById("processingStatus");
    const resultCard = document.getElementById("resultCard");
    const resultMeta = document.getElementById("resultMeta");
    const downloadBtn = document.getElementById("downloadBtn");
    const reportSection = document.getElementById("reportSection");
    const reportContent = document.getElementById("reportContent");
    const anotherBtn = document.getElementById("anotherBtn");
    const errorCard = document.getElementById("errorCard");
    const errorMessage = document.getElementById("errorMessage");
    const retryBtn = document.getElementById("retryBtn");

    let selectedFile = null;
    let downloadUrl = null;

    // ── File Type Icons ───────────────────────────────────────────
    const FILE_COLORS = {
        ".docx": "#3b82f6",
        ".pdf":  "#ef4444",
        ".png":  "#10b981",
        ".jpg":  "#f59e0b",
        ".jpeg": "#f59e0b",
        ".tiff": "#8b5cf6",
        ".bmp":  "#6366f1",
    };

    // ── Drag & Drop ───────────────────────────────────────────────
    dropZone.addEventListener("click", () => fileInput.click());

    dropZone.addEventListener("dragover", (e) => {
        e.preventDefault();
        dropZone.classList.add("drag-over");
    });

    dropZone.addEventListener("dragleave", () => {
        dropZone.classList.remove("drag-over");
    });

    dropZone.addEventListener("drop", (e) => {
        e.preventDefault();
        dropZone.classList.remove("drag-over");
        if (e.dataTransfer.files.length > 0) {
            handleFile(e.dataTransfer.files[0]);
        }
    });

    fileInput.addEventListener("change", () => {
        if (fileInput.files.length > 0) {
            handleFile(fileInput.files[0]);
        }
    });

    // ── File Selection ────────────────────────────────────────────
    function handleFile(file) {
        const ext = "." + file.name.split(".").pop().toLowerCase();
        const allowed = [".docx", ".pdf", ".png", ".jpg", ".jpeg", ".tiff", ".bmp"];

        if (!allowed.includes(ext)) {
            showError(`Unsupported file type "${ext}". Accepted: ${allowed.join(", ")}`);
            return;
        }

        selectedFile = file;

        // Update preview
        fileName.textContent = file.name;
        fileSize.textContent = formatSize(file.size);

        // Set icon color by file type
        const iconColor = FILE_COLORS[ext] || "#3b82f6";
        document.getElementById("fileTypeIcon").style.background =
            `linear-gradient(135deg, ${iconColor}, ${iconColor}dd)`;

        // Show preview, hide drop zone
        dropZone.style.display = "none";
        filePreview.style.display = "block";
    }

    // ── Remove File ───────────────────────────────────────────────
    removeFile.addEventListener("click", () => {
        resetToUpload();
    });

    // ── Convert ───────────────────────────────────────────────────
    convertBtn.addEventListener("click", async () => {
        if (!selectedFile) return;

        // Show processing state
        showCard(processingCard);
        updateStatus("Uploading file...");

        const formData = new FormData();
        formData.append("file", selectedFile);

        try {
            updateStatus("Converting document...");

            const response = await fetch("/api/convert", {
                method: "POST",
                body: formData,
            });

            const data = await response.json();

            if (!response.ok || data.error) {
                showError(data.error || "Unknown error occurred");
                return;
            }

            // Success — show result
            downloadUrl = `/api/download/${data.output_file}`;

            // Build metadata chips
            resultMeta.innerHTML = `
                <span class="meta-chip">
                    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/></svg>
                    ${data.input_type}
                </span>
                <span class="meta-chip">
                    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="10"/><polyline points="12 6 12 12 16 14"/></svg>
                    ${data.elapsed}s
                </span>
            `;

            // Show confidence report if available
            if (data.report) {
                reportContent.textContent = data.report;
                reportSection.style.display = "block";
            } else {
                reportSection.style.display = "none";
            }

            showCard(resultCard);

        } catch (err) {
            showError(`Connection error: ${err.message}`);
        }
    });

    // ── Download ──────────────────────────────────────────────────
    downloadBtn.addEventListener("click", () => {
        if (downloadUrl) {
            window.location.href = downloadUrl;
        }
    });

    // ── Another File ──────────────────────────────────────────────
    anotherBtn.addEventListener("click", resetToUpload);
    retryBtn.addEventListener("click", resetToUpload);

    // ── Helpers ───────────────────────────────────────────────────
    function showCard(card) {
        [uploadCard, processingCard, resultCard, errorCard].forEach(c => {
            c.style.display = "none";
        });
        card.style.display = "block";
        // Re-trigger animation
        card.style.animation = "none";
        card.offsetHeight; // force reflow
        card.style.animation = "";
    }

    function showError(message) {
        errorMessage.textContent = message;
        showCard(errorCard);
    }

    function updateStatus(text) {
        processingStatus.textContent = text;
    }

    function resetToUpload() {
        selectedFile = null;
        downloadUrl = null;
        fileInput.value = "";
        dropZone.style.display = "block";
        filePreview.style.display = "none";
        showCard(uploadCard);
    }

    function formatSize(bytes) {
        if (bytes < 1024) return bytes + " B";
        if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + " KB";
        return (bytes / (1024 * 1024)).toFixed(1) + " MB";
    }
});
