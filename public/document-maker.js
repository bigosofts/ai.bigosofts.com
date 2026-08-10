document.addEventListener("DOMContentLoaded", () => {
  console.log("[document-maker] init");
  try {
    const content = document.getElementById("content");
    const count = document.getElementById("count");
    const imagesEl = document.getElementById("images");
    const button = document.getElementById("generateBtn");
    const result = document.getElementById("result");
    const error = document.getElementById("error");

    if (!content || !count || !imagesEl || !button) {
      console.error("[document-maker] Missing required DOM nodes", {
        content: !!content,
        count: !!count,
        images: !!imagesEl,
        button: !!button,
      });
      return;
    }

    const SESSION_ID =
      sessionStorage.getItem("bos_session_id") || "sess-" + crypto.randomUUID();
    sessionStorage.setItem("bos_session_id", SESSION_ID);
    let imageId = 0;
    let imageRecords = [];

    content.addEventListener(
      "input",
      () => (count.textContent = content.value.length.toLocaleString()),
    );

    const addImageBtn = document.getElementById("addImage");
    if (addImageBtn) {
      addImageBtn.addEventListener("click", () => {
        const id = ++imageId;
        const row = document.createElement("div");
        row.className = "image-row";
        row.dataset.id = id;
        row.innerHTML = `<input type="file" accept="image/*" aria-label="Image file"><input type="text" placeholder="Describe placement, e.g. team photo after intro" aria-label="Image description"><button type="button" aria-label="Remove image">×</button>`;
        const [file, description, remove] =
          row.querySelectorAll("input,button");
        file.addEventListener("change", async () => {
          if (!file.files[0]) return;
          file.disabled = true;
          const form = new FormData();
          form.append("image", file.files[0]);
          try {
            const r = await fetch("/api/upload-image", {
              method: "POST",
              headers: { "X-Session-ID": SESSION_ID },
              body: form,
            });
            const data = await r.json();
            if (!data.success) throw new Error(data.error);
            imageRecords = imageRecords.filter((x) => x.id !== id);
            imageRecords.push({
              id,
              type: "upload",
              serverPath: data.filePath,
              url: data.staticUrl,
            });
            file.title = "Uploaded: " + file.files[0].name;
          } catch (e) {
            alert("Image upload failed: " + e.message);
            file.value = "";
          } finally {
            file.disabled = false;
          }
        });
        remove.addEventListener("click", () => {
          imageRecords = imageRecords.filter((x) => x.id !== id);
          row.remove();
        });
        imagesEl.appendChild(row);
      });
    }

    // Helper for reset UI
    window.resetUI = function () {
      if (typeof setViewState === "function") {
        setViewState("state-idle");
      }
      if (result) result.hidden = true;
      if (error) error.hidden = true;
      const progressBar = document.getElementById("progressBar");
      if (progressBar) progressBar.style.width = "0%";
    };

    button.addEventListener("click", async () => {
      const value = content.value.trim();
      if (!value) {
        alert("Please enter brief instructions or content for your document.");
        return;
      }

      button.disabled = true;
      const btnText = document.getElementById("btnContent");
      const loading = document.getElementById("btnLoading");
      if (btnText) btnText.hidden = true;
      if (loading) loading.style.display = "flex";

      if (error) error.hidden = true;
      if (result) result.hidden = true;

      const images = [...imagesEl.children].flatMap((row) => {
        const record = imageRecords.find(
          (x) => x.id === Number(row.dataset.id),
        );
        return record
          ? [
              {
                ...record,
                description:
                  row.querySelectorAll("input")[1].value.trim() ||
                  "Supporting image",
              },
            ]
          : [];
      });

      try {
        if (typeof setViewState === "function") {
          setViewState("state-loading");
        }
        const activityMessage = document.getElementById("activityMessage");
        if (activityMessage)
          activityMessage.textContent =
            "Reading your topic and identifying structure";
        const progressBar = document.getElementById("progressBar");
        if (progressBar) progressBar.style.width = "20%";

        const response = await fetch("/api/generate-document", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ content: value, images }),
        });
        const data = await response.json();
        if (!response.ok || !data.success)
          throw new Error(data.error || "Document generation failed.");

        const downloadsList = document.getElementById("downloadsList");
        if (downloadsList) {
          downloadsList.innerHTML = "";
          const item = document.createElement("div");
          item.className = "download-item";
          item.innerHTML = `<a href="/api/download-document/${data.fileId}" download="${data.fileName || "AI_Document.docx"}">${data.fileName || "AI_Document.docx"}</a>`;
          downloadsList.appendChild(item);
        }
        const downloadBtn = document.getElementById("downloadBtn");
        if (downloadBtn) {
          downloadBtn.href = "/api/download-document/" + data.fileId;
          downloadBtn.download = data.fileName || "AI_Document.docx";
          downloadBtn.hidden = false;
        }

        const progressBar2 = document.getElementById("progressBar");
        if (progressBar2) progressBar2.style.width = "100%";
        const progressTitle = document.getElementById("progressTitle");
        if (progressTitle)
          progressTitle.textContent = "Finalizing your document";
        if (activityMessage)
          activityMessage.textContent = "Your document is ready to download";
        if (typeof setViewState === "function") {
          setViewState("state-success");
        }

        const resultName = document.getElementById("resultName");
        if (resultName)
          resultName.textContent = data.fileName || "Editable .docx file";
        const link = document.getElementById("download");
        if (link) {
          link.href = "/api/download-document/" + data.fileId;
          link.download = data.fileName || "AI_Document.docx";
        }
        if (result) result.hidden = false;
        if (result) result.scrollIntoView({ behavior: "smooth", block: "nearest" });
      } catch (e) {
        console.error("[document-maker] generation error", e);
        if (typeof stopActivityFeed === "function") stopActivityFeed();
        if (typeof setViewState === "function") {
          setViewState("state-error");
        }
        const errMsg = document.getElementById("errorMessage");
        if (errMsg) errMsg.textContent = e.message || "Generation failed.";

        if (error) {
          error.textContent = e.message;
          error.hidden = false;
        }
      } finally {
        button.disabled = false;
        if (btnText) btnText.hidden = false;
        if (loading) loading.style.display = "none";
      }
    });
  } catch (e) {
    console.error("[document-maker] init error", e);
  }
});