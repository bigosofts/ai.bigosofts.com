// Auth UI and flow for gated pages
(function () {
  const pageMap = {
    "slide-maker.html": "Slide Maker",
    "document-maker.html": "Document Maker",
    "data-extractor.html": "Data Extractor",
  };
  const path = window.location.pathname.split("/").pop();
  const pageKey = path || "slide-maker.html";
  const pageName = pageMap[pageKey] || pageKey;
  const localKey = `authorized_${pageKey}`;
  if (localStorage.getItem(localKey)) return; // already authorized

  // Create modal container
  const tpl = document.createElement("div");
  tpl.innerHTML = `
  <div id="authOverlay" class="auth-overlay">
    <div class="auth-modal" role="dialog" aria-modal="true" aria-labelledby="authTitle">
      <div class="auth-header">
        <div class="auth-brand">B</div>
        <div>
          <h2 id="authTitle">Request access to ${pageName}</h2>
          <div class="auth-sub">Secure access — one-time code will be emailed to you</div>
        </div>
      </div>
      <div class="auth-body" id="authBody">
        <div class="auth-error" id="authError"></div>
        <form id="authForm">
          <div class="auth-field"><label>Full name<input name="name" required autocomplete="name" /></label></div>
          <div class="auth-field"><label>Age (optional)<input name="age" type="number" min="1" inputmode="numeric" /></label></div>
          <div class="auth-field"><label>Email address<input name="email" type="email" required autocomplete="email" /></label></div>
          <div class="auth-note">We will send a one-time code to your email to verify access.</div>
          <div class="auth-actions">
            <button type="submit" class="auth-btn" id="requestBtn">Request access</button>
            <a class="auth-cancel" href="https://bigosofts.com/ai-creation">Cancel</a>
          </div>
        </form>
      </div>
      <div class="auth-footer"><div class="auth-small">Privacy: we only store name & email when access is granted.</div></div>
    </div>
  </div>
  `;
  document.body.appendChild(tpl);

  const overlay = document.getElementById("authOverlay");
  overlay.style.display = "flex";

  // focus management
  const focusables =
    'button, [href], input, select, textarea, [tabindex]:not([tabindex="-1"])';
  function trapFocus(container) {
    const els = container.querySelectorAll(focusables);
    const first = els[0];
    const last = els[els.length - 1];
    first && first.focus();
    container.addEventListener("keydown", function (e) {
      if (e.key !== "Tab") return;
      if (e.shiftKey && document.activeElement === first) {
        e.preventDefault();
        last.focus();
      } else if (!e.shiftKey && document.activeElement === last) {
        e.preventDefault();
        first.focus();
      }
    });
  }

  async function postJson(url, data) {
    const res = await fetch(url, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(data),
    });
    return res.json();
  }
  const authForm = document.getElementById("authForm");
  const authError = document.getElementById("authError");
  trapFocus(document.querySelector(".auth-modal"));

  function showError(msg) {
    authError.textContent = msg;
    authError.style.display = msg ? "block" : "none";
  }

  authForm.addEventListener("submit", async function (e) {
    e.preventDefault();
    showError("");
    const f = new FormData(e.target);
    const name = f.get("name");
    const age = f.get("age");
    const email = f.get("email");
    const page = pageKey;
    try {
      const btn = document.getElementById("requestBtn");
      btn.disabled = true;
      btn.textContent = "Sending...";
      const r = await postJson("/api/request-access", {
        name,
        age,
        email,
        page,
      });
      if (!r || !r.success)
        throw new Error(r && r.error ? r.error : "Request failed");
      showOtpForm(name, email);
    } catch (err) {
      showError(err.message || String(err));
      document.getElementById("requestBtn").disabled = false;
      document.getElementById("requestBtn").textContent = "Request access";
    }
  });

  function showOtpForm(name, email) {
    const body = document.getElementById("authBody");
    body.innerHTML = `
      <div class="auth-error" id="authErrorOtp"></div>
      <p>Thanks ${name}. We sent a one-time code to <strong>${email}</strong>. Enter it below to continue. The code expires in 10 minutes.</p>
      <form id="otpForm">
        <div class="auth-field"><label>Enter code<input name="otp" inputmode="numeric" pattern="\\d{6}" placeholder="123456" class="otp-input" required /></label></div>
        <div class="auth-actions">
          <button type="submit" class="auth-btn" id="verifyBtn">Verify code</button>
          <a class="auth-cancel" href="https://bigosofts.com/ai-creation">Cancel</a>
        </div>
      </form>
    `;
    const otpError = document.getElementById("authErrorOtp");
    const otpForm = document.getElementById("otpForm");
    otpForm.addEventListener("submit", async function (e) {
      e.preventDefault();
      otpError.style.display = "none";
      const otp = new FormData(e.target).get("otp");
      try {
        const btn = document.getElementById("verifyBtn");
        btn.disabled = true;
        btn.textContent = "Verifying...";
        const r = await postJson("/api/verify-otp", {
          email,
          otp,
          page: pageKey,
        });
        if (!r || !r.success)
          throw new Error(r && r.error ? r.error : "Verification failed");
        localStorage.setItem(localKey, "1");
        overlay.remove();
      } catch (err) {
        otpError.textContent = err.message || String(err);
        otpError.style.display = "block";
        document.getElementById("verifyBtn").disabled = false;
        document.getElementById("verifyBtn").textContent = "Verify code";
      }
    });
    trapFocus(document.querySelector(".auth-modal"));
  }
})();
