document.addEventListener("DOMContentLoaded", () => {
  const SESSION_ID =
    sessionStorage.getItem("bos_session_id") || "sess-" + crypto.randomUUID();
  sessionStorage.setItem("bos_session_id", SESSION_ID);

  const contentInput = document.getElementById("userInput");
  const countDisplay = document.getElementById("charCount");
  const fileInput = document.getElementById("fileInput");
  const fileList = document.getElementById("fileList");
  const generateBtn = document.getElementById("generateBtn");

  // Sections
  const inputSection = document.getElementById("inputSection");
  const resultSection = document.getElementById("resultSection");
  const errorSection = document.getElementById("errorSection");
  const actionPanel = document.getElementById("actionPanel");

  // Camera Elements
  const cameraBtn = document.getElementById("cameraBtn");
  const cameraModal = document.getElementById("cameraModal");
  const cameraVideo = document.getElementById("cameraVideo");
  const captureBtn = document.getElementById("captureBtn");
  const closeCameraBtn = document.getElementById("closeCameraBtn");
  let stream = null;

  let files = [];
  let fileIdCounter = 0;

  // Character counter
  if (contentInput && countDisplay) {
    contentInput.addEventListener("input", () => {
      countDisplay.textContent = contentInput.value.length.toLocaleString();
    });
  }

  // File Upload Handling
  if (fileInput) {
    fileInput.addEventListener("change", async () => {
      if (!fileInput.files.length) return;
      for (const file of fileInput.files) {
        await uploadFile(file);
      }
      fileInput.value = "";
    });
  }

  // Camera Handling
  if (cameraBtn) {
    cameraBtn.addEventListener("click", async () => {
      try {
        // Request rear camera (environment) with optimized constraints for document capture
        const constraints = {
          video: {
            facingMode: { ideal: "environment" }, // Prefer rear camera
            width: { ideal: 1920 },
            height: { ideal: 1440 },
            aspectRatio: { ideal: 4 / 3 },
          },
          audio: false,
        };

        stream = await navigator.mediaDevices.getUserMedia(constraints);
        cameraVideo.srcObject = stream;
        cameraModal.classList.remove("hidden");

        // Show visual feedback
        console.log("✅ Rear camera initialized for document capture");
      } catch (err) {
        // Fallback: try any available camera
        try {
          console.warn(
            "⚠️ Rear camera not available, trying any camera:",
            err.message,
          );
          const fallbackConstraints = {
            video: { width: { ideal: 1920 }, height: { ideal: 1440 } },
            audio: false,
          };
          stream =
            await navigator.mediaDevices.getUserMedia(fallbackConstraints);
          cameraVideo.srcObject = stream;
          cameraModal.classList.remove("hidden");
          console.log("✅ Fallback camera initialized");
        } catch (fallbackErr) {
          alert("Camera access denied or unavailable: " + fallbackErr.message);
          console.error("❌ Camera initialization failed:", fallbackErr);
        }
      }
    });
  }

  if (closeCameraBtn) {
    closeCameraBtn.addEventListener("click", closeCamera);
  }

  if (captureBtn) {
    captureBtn.addEventListener("click", () => {
      if (!stream) return;
      captureBtn.disabled = true;
      captureBtn.textContent = "📸 Processing...";

      // Step 1: Capture at full resolution
      const fullCanvas = document.createElement("canvas");
      fullCanvas.width = cameraVideo.videoWidth;
      fullCanvas.height = cameraVideo.videoHeight;
      const ctx = fullCanvas.getContext("2d");
      ctx.drawImage(cameraVideo, 0, 0, fullCanvas.width, fullCanvas.height);

      // Step 2: Resize to optimized dimensions (900x1200) for AI text processing
      const optimizedCanvas = document.createElement("canvas");
      const maxWidth = 900;
      const maxHeight = 1200;
      let targetWidth = fullCanvas.width;
      let targetHeight = fullCanvas.height;

      if (targetWidth > maxWidth || targetHeight > maxHeight) {
        const aspectRatio = fullCanvas.width / fullCanvas.height;
        if (targetWidth > targetHeight) {
          targetWidth = Math.min(maxWidth, targetWidth);
          targetHeight = Math.round(targetWidth / aspectRatio);
        } else {
          targetHeight = Math.min(maxHeight, targetHeight);
          targetWidth = Math.round(targetHeight * aspectRatio);
        }
      }

      optimizedCanvas.width = targetWidth;
      optimizedCanvas.height = targetHeight;
      const optimizedCtx = optimizedCanvas.getContext("2d");
      optimizedCtx.drawImage(fullCanvas, 0, 0, targetWidth, targetHeight);

      // Step 3: Convert to optimized JPEG (quality: 70%, file size ~80-150KB)
      optimizedCanvas.toBlob(
        async (blob) => {
          const file = new File([blob], `camera_capture_${Date.now()}.jpg`, {
            type: "image/jpeg",
          });
          closeCamera();
          captureBtn.disabled = false;
          captureBtn.textContent = "Capture";
          await uploadFile(file);
        },
        "image/jpeg",
        0.7,
      );
    });
  }

  function closeCamera() {
    if (stream) {
      stream.getTracks().forEach((track) => track.stop());
      stream = null;
    }
    cameraModal.classList.add("hidden");
  }

  async function uploadFile(file) {
    const id = ++fileIdCounter;

    const row = document.createElement("div");
    row.className = "file-item";
    row.dataset.id = id;

    const isImage = file.type.startsWith("image/");
    const thumbSrc = isImage ? URL.createObjectURL(file) : "";

    const rawSvg = `<svg xmlns="http://www.w3.org/2000/svg" width="36" height="36" viewBox="0 0 24 24" fill="none" stroke="#9ca3af" stroke-width="2"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path><polyline points="14 2 14 8 20 8"></polyline><line x1="16" y1="13" x2="8" y2="13"></line><line x1="16" y1="17" x2="8" y2="17"></line><polyline points="10 9 9 9 8 9"></polyline></svg>`;
    const fallbackSvg = `data:image/svg+xml;utf8,${encodeURIComponent(rawSvg)}`;

    row.innerHTML = `
      <img src="${isImage ? thumbSrc : fallbackSvg}" class="file-thumb" alt="thumbnail" onerror="this.onerror=null; this.src='${fallbackSvg}';">
      <span class="file-name">Uploading: ${file.name}...</span>
      <button class="btn-remove-file" title="Remove">✕</button>
    `;

    const removeBtn = row.querySelector(".btn-remove-file");
    removeBtn.disabled = true;
    fileList.appendChild(row);

    const form = new FormData();
    form.append("file", file);

    try {
      const r = await fetch("/api/upload-file", {
        method: "POST",
        headers: { "X-Session-ID": SESSION_ID },
        body: form,
      });
      const data = await r.json();
      if (!data.success) throw new Error(data.error);

      files.push({
        id,
        serverPath: data.filePath,
        mimeType: data.mimeType,
        url: data.staticUrl,
      });
      row.querySelector(".file-name").textContent = file.name;

      removeBtn.disabled = false;
      removeBtn.addEventListener("click", () => {
        files = files.filter((f) => f.id !== id);
        row.remove();
      });
    } catch (err) {
      alert("Upload failed: " + err.message);
      row.remove();
    }
  }

  // ─── View State Management ───
  function showView(view) {
    // view = "input" | "result" | "error"
    inputSection.style.display = view === "input" ? "block" : "none";
    resultSection.style.display = view === "result" ? "block" : "none";
    errorSection.style.display = view === "error" ? "block" : "none";
  }

  window.resetUI = function () {
    showView("input");
  };

  // ─── Generate / Extract ───
  if (generateBtn) {
    generateBtn.addEventListener("click", async () => {
      const prompt = contentInput.value.trim();
      if (!prompt) {
        alert("Please enter instructions for the AI.");
        return;
      }
      if (files.length === 0) {
        alert("Please upload at least one image or PDF.");
        return;
      }

      generateBtn.disabled = true;
      document.getElementById("btnContent").style.display = "none";
      document.getElementById("btnLoading").style.display = "flex";

      try {
        const filePaths = files.map((f) => f.serverPath);

        // Use hybrid progressive batch endpoint
        const response = await fetch("/api/extract-batch", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ prompt, filePaths }),
        });

        const data = await response.json();

        if (!response.ok || !data.success) {
          throw new Error(data.error || "Data extraction failed.");
        }

        const batchId = data.batchId;
        const totalBatches = data.totalBatches;
        let allResults = [{ batch: 1, result: data.result }];

        // Display first result immediately
        showBatchResults(allResults, totalBatches, 1);
        showView("result");

        // Poll for remaining batches if more than one batch
        if (totalBatches > 1) {
          await pollBatchResults(batchId, allResults, totalBatches);
        }
      } catch (e) {
        console.error("Extraction error:", e);
        document.getElementById("errorMessage").textContent = e.message;
        showView("error");
      } finally {
        generateBtn.disabled = false;
        document.getElementById("btnContent").style.display = "flex";
        document.getElementById("btnLoading").style.display = "none";
      }
    });
  }

  // Poll for batch results
  async function pollBatchResults(batchId, allResults, totalBatches) {
    let currentBatch = 1;
    let maxAttempts = 120; // 2 minutes max
    let attempts = 0;

    while (currentBatch < totalBatches && attempts < maxAttempts) {
      attempts++;
      await new Promise((resolve) => setTimeout(resolve, 1000)); // Wait 1 second

      try {
        const response = await fetch(`/api/get-batch-result/${batchId}`);
        const data = await response.json();

        if (data.success && data.results.length > allResults.length) {
          // New results available
          allResults = data.results;
          showBatchResults(allResults, totalBatches, data.currentBatch);
          currentBatch = data.currentBatch;
        }

        // All batches complete
        if (!data.pending) {
          break;
        }
      } catch (err) {
        console.warn("Polling error:", err);
      }
    }
  }

  // Display batch results progressively
  function showBatchResults(results, totalBatches, currentBatch) {
    const converter = new showdown.Converter({
      tables: true,
      strikethrough: true,
      tasklists: true,
      openLinksInNewWindow: true,
    });

    const resultContent = document.getElementById("resultContent");

    // Build batched result display
    let html = `<div class="batch-progress">📊 Processing: Batch ${currentBatch} of ${totalBatches}</div>`;

    results.forEach((batchItem, index) => {
      const isComplete = batchItem.status === "complete";
      html += `
        <div class="batch-result-card ${isComplete ? "complete" : "loading"}">
          <div class="batch-header">
            <span class="batch-number">Batch ${batchItem.batch}</span>
            ${isComplete ? '<span class="batch-status">✓ Complete</span>' : '<span class="batch-status loading">⏳ Processing...</span>'}
          </div>
          <div class="batch-content">
            ${converter.makeHtml(batchItem.result)}
          </div>
        </div>
      `;
    });

    resultContent.innerHTML = html;
  }

  // ─── Copy Button ───
  const copyBtn = document.getElementById("copyBtn");
  if (copyBtn) {
    copyBtn.addEventListener("click", () => {
      const resultContent = document.getElementById("resultContent");
      const copyBtnText = document.getElementById("copyBtnText");

      // Copy both rich text (HTML) and plain text
      const htmlContent = resultContent.innerHTML;
      const textContent = resultContent.innerText;

      // Try to copy rich text first for pasting into Word/Docs
      if (navigator.clipboard && navigator.clipboard.write) {
        const htmlBlob = new Blob([htmlContent], { type: "text/html" });
        const textBlob = new Blob([textContent], { type: "text/plain" });
        const item = new ClipboardItem({
          "text/html": htmlBlob,
          "text/plain": textBlob,
        });
        navigator.clipboard
          .write([item])
          .then(() => {
            copyBtn.classList.add("copied");
            copyBtnText.textContent = "Copied!";
            setTimeout(() => {
              copyBtn.classList.remove("copied");
              copyBtnText.textContent = "Copy All";
            }, 2500);
          })
          .catch(() => {
            // Fallback to plain text
            navigator.clipboard.writeText(textContent).then(() => {
              copyBtn.classList.add("copied");
              copyBtnText.textContent = "Copied!";
              setTimeout(() => {
                copyBtn.classList.remove("copied");
                copyBtnText.textContent = "Copy All";
              }, 2500);
            });
          });
      } else {
        navigator.clipboard.writeText(textContent).then(() => {
          copyBtn.classList.add("copied");
          copyBtnText.textContent = "Copied!";
          setTimeout(() => {
            copyBtn.classList.remove("copied");
            copyBtnText.textContent = "Copy All";
          }, 2500);
        });
      }
    });
  }

  // ─── Extract Another / Retry ───
  const extractAnotherBtn = document.getElementById("extractAnotherBtn");
  if (extractAnotherBtn) {
    extractAnotherBtn.addEventListener("click", () => {
      showView("input");
    });
  }

  const errorRetryBtn = document.getElementById("errorRetryBtn");
  if (errorRetryBtn) {
    errorRetryBtn.addEventListener("click", () => {
      showView("input");
    });
  }
});
