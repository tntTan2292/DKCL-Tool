const DKCL_ORIGIN = "https://dkcl.vnpost.vn/";
const SIDEBAR_PATH = "popup.html";

chrome.runtime.onInstalled.addListener(async () => {
  await enableActionClickSidebar();
  await openSidebarForExistingDkclTabs();
});

chrome.runtime.onStartup.addListener(async () => {
  await enableActionClickSidebar();
  await openSidebarForExistingDkclTabs();
});

chrome.tabs.onUpdated.addListener((tabId, changeInfo, tab) => {
  const nextUrl = changeInfo.url || tab.url || "";
  if (isDkclPage(nextUrl) && (changeInfo.status === "complete" || changeInfo.url)) {
    openSidebar(tabId, tab.windowId);
  }
});

chrome.tabs.onActivated.addListener(async ({ tabId }) => {
  try {
    const tab = await chrome.tabs.get(tabId);
    if (isDkclPage(tab.url)) {
      openSidebar(tab.id, tab.windowId);
    }
  } catch (error) {
    console.warn("Không đọc được tab hiện tại:", error);
  }
});

async function enableActionClickSidebar() {
  await chrome.sidePanel.setPanelBehavior({ openPanelOnActionClick: true });
}

async function openSidebarForExistingDkclTabs() {
  const tabs = await chrome.tabs.query({ url: `${DKCL_ORIGIN}*` });
  tabs.forEach((tab) => openSidebar(tab.id, tab.windowId));
}

function isDkclPage(url = "") {
  return url.startsWith(DKCL_ORIGIN);
}

async function openSidebar(tabId, windowId) {
  if (!tabId || !windowId) return;

  try {
    await chrome.sidePanel.setOptions({ tabId, path: SIDEBAR_PATH, enabled: true });
    await chrome.sidePanel.open({ windowId });
  } catch (error) {
    console.warn("Không thể tự mở DKCL sidebar:", error);
  }
}
