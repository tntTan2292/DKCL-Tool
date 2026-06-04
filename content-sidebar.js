const DKCL_SIDEBAR_ID = "dkcl-report-extension-sidebar";
const DKCL_TOGGLE_ID = "dkcl-report-extension-toggle";
const DKCL_PIN_STORAGE_KEY = "dkclSidebarPinned";

function mountDkclSidebar() {
  if (document.getElementById(DKCL_SIDEBAR_ID)) return;

  const sidebar = document.createElement("aside");
  sidebar.id = DKCL_SIDEBAR_ID;
  sidebar.style.cssText = [
    "position: fixed",
    "top: 72px",
    "right: 16px",
    "width: 380px",
    "max-width: min(380px, calc(100vw - 32px))",
    "height: calc(100vh - 104px)",
    "z-index: 2147483000",
    "background: #07111f",
    "box-shadow: -16px 18px 48px rgba(2, 6, 23, 0.28)",
    "border: 1px solid rgba(255,255,255,0.18)",
    "border-radius: 24px",
    "transition: transform 220ms ease, opacity 220ms ease",
    "overflow: hidden",
    `transform: ${localStorage.getItem(DKCL_PIN_STORAGE_KEY) === "true" ? "translateX(0)" : "translateX(calc(100% + 28px))"}`,
    "opacity: 0.96"
  ].join(";");

  const iframe = document.createElement("iframe");
  iframe.title = "DKCL Báo cáo BĐHN";
  iframe.src = chrome.runtime.getURL("popup.html?embedded=1");
  iframe.style.cssText = [
    "width: 100%",
    "height: 100%",
    "border: 0",
    "display: block",
    "background: transparent"
  ].join(";");

  const toggle = document.createElement("button");
  toggle.id = DKCL_TOGGLE_ID;
  toggle.type = "button";
  toggle.textContent = "DKCL";
  toggle.title = "Ẩn/hiện DKCL Sidebar";
  toggle.style.cssText = [
    "position: fixed",
    "top: 50%",
    "right: 12px",
    "transform: translateY(-50%)",
    "z-index: 2147483001",
    "height: 42px",
    "padding: 0 12px",
    "border: 0",
    "border-radius: 999px",
    "background: linear-gradient(135deg, #fbbf24, #f97316)",
    "color: #111827",
    "font: 700 12px Arial, sans-serif",
    "letter-spacing: .04em",
    "box-shadow: 0 10px 24px rgba(0,0,0,.22)",
    "cursor: pointer",
    "transition: right 220ms ease, transform 220ms ease"
  ].join(";");

  let collapsed = localStorage.getItem(DKCL_PIN_STORAGE_KEY) !== "true";
  let pinned = localStorage.getItem(DKCL_PIN_STORAGE_KEY) === "true";
  toggle.style.right = collapsed ? "12px" : "408px";
  toggle.textContent = pinned ? "Đã ghim" : collapsed ? "DKCL" : "Ẩn";

  toggle.addEventListener("click", () => {
    if (pinned) return;
    collapsed = !collapsed;
    sidebar.style.transform = collapsed ? "translateX(calc(100% + 28px))" : "translateX(0)";
    toggle.style.right = collapsed ? "12px" : "408px";
    toggle.textContent = collapsed ? "DKCL" : "Ẩn";
  });

  window.addEventListener("message", (event) => {
    if (event.source !== iframe.contentWindow || event.data?.source !== "dkcl-report-popup" || event.data?.type !== "pin-state") return;
    pinned = Boolean(event.data.pinned);
    localStorage.setItem(DKCL_PIN_STORAGE_KEY, String(pinned));
    collapsed = false;
    sidebar.style.transform = "translateX(0)";
    toggle.style.right = "408px";
    toggle.textContent = pinned ? "Đã ghim" : "Ẩn";
    toggle.title = pinned ? "Giao diện DKCL đang được ghim" : "Ẩn/hiện DKCL Sidebar";
  });

  sidebar.appendChild(iframe);
  document.documentElement.appendChild(sidebar);
  document.documentElement.appendChild(toggle);
}

if (document.readyState === "loading") {
  document.addEventListener("DOMContentLoaded", mountDkclSidebar, { once: true });
} else {
  mountDkclSidebar();
}

// DKCL có thể đổi route sau đăng nhập không reload full page, nên kiểm tra lại vài lần.
let retryCount = 0;
const retryTimer = setInterval(() => {
  mountDkclSidebar();
  retryCount += 1;
  if (retryCount >= 10) clearInterval(retryTimer);
}, 1000);

