/* Shared UI primitives: DOM factory, menus, dialogs, palettes, toasts.
 *
 * el() is the single place where the application builds DOM.  It understands
 * a few conveniences beyond setAttribute: class, text, html, dataset, on*
 * listeners and nested children.  Text and title/placeholder/aria-label
 * attributes are localized on creation, so modules only ever write English.
 */

import { icon } from "./icons.js";
import { localizeNode, translateAttribute } from "./i18n.js";

export const $ = (sel, root = document) => root.querySelector(sel);
export const $$ = (sel, root = document) => Array.from(root.querySelectorAll(sel));

const BOOLEAN_ATTRS = new Set([
  "disabled", "checked", "readonly", "required", "selected", "multiple",
  "hidden", "autofocus", "inert",
]);

export function el(tag, attrs = {}, children = []) {
  const node = document.createElement(tag);
  const untranslated = attrs.i18n === false;
  if (untranslated) node.setAttribute("data-no-i18n", "");
  for (const [key, value] of Object.entries(attrs)) {
    if (value === undefined || value === null) continue;
    if (key === "i18n") continue;
    if (key === "class") node.className = value;
    else if (key === "html") node.innerHTML = value;
    else if (key === "text") {
      const text = document.createTextNode(String(value));
      node.append(text);
      if (!untranslated) localizeNode(text);
    }
    else if (key === "dataset") Object.assign(node.dataset, value);
    else if (key === "style" && typeof value === "object") Object.assign(node.style, value);
    else if (key.startsWith("on") && typeof value === "function") node.addEventListener(key.slice(2), value);
    else if (BOOLEAN_ATTRS.has(key.toLowerCase())) {
      // setAttribute("disabled", "false") still disables the control, so
      // boolean attributes are always toggled by presence.
      node.toggleAttribute(key, !!value);
    }
    else {
      node.setAttribute(key, String(value));
      if (!untranslated) translateAttribute(node, key);
    }
  }
  for (const child of [].concat(children)) {
    if (child === null || child === undefined || child === false) continue;
    if (child.nodeType) {
      node.append(child);
      continue;
    }
    const text = document.createTextNode(String(child));
    node.append(text);
    if (!untranslated) localizeNode(text);
  }
  return node;
}

export function escapeHtml(text) {
  return String(text ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

export function clamp(value, min, max) {
  return Math.min(max, Math.max(min, value));
}

export function setIcon(node, name, cls = "") {
  node.innerHTML = icon(name, cls);
  return node;
}

/* ---------------------------------------------------------------- controls */
export function field(label, control, hint = "") {
  const wrap = el("label", { class: "field" });
  if (label) wrap.append(el("span", { text: label }));
  wrap.append(control);
  if (hint) wrap.append(el("small", { class: "dialog-note", text: hint }));
  return wrap;
}

export function textInput(value = "", attrs = {}) {
  return el("input", { type: "text", value, ...attrs });
}

export function numberInput(value = 0, attrs = {}) {
  return el("input", { type: "number", value: String(value), ...attrs });
}

export function selectControl(options, value, attrs = {}) {
  const select = el("select", attrs);
  for (const option of options) {
    const item = el("option", { value: String(option.value) }, option.label);
    if (String(option.value) === String(value)) item.selected = true;
    select.append(item);
  }
  return select;
}

export function checkRow(label, checked = false, onChange = null) {
  const input = el("input", { type: "checkbox" });
  input.checked = !!checked;
  if (onChange) input.addEventListener("change", () => onChange(input.checked));
  return el("label", { class: "check-row" }, [input, el("span", { text: label })]);
}

export function segmented(options, value, onPick) {
  const wrap = el("div", { class: "seg" });
  const buttons = [];
  for (const option of options) {
    const button = el("button", { type: "button", text: option.label });
    if (String(option.value) === String(value)) button.classList.add("active");
    button.addEventListener("click", () => {
      for (const other of buttons) other.classList.toggle("active", other === button);
      onPick?.(option.value);
    });
    buttons.push(button);
    wrap.append(button);
  }
  return wrap;
}

export function iconButton(name, title, onClick, cls = "") {
  const button = el("button", { class: "tb-icon " + cls, title, type: "button", html: icon(name) });
  if (onClick) button.addEventListener("click", onClick);
  return button;
}

/* ------------------------------------------------------------------ toasts */
export function toast(message, kind = "", timeout = 2600) {
  const layer = $("#toast-layer");
  if (!layer) return null;
  const node = el("div", { class: "toast " + kind, text: message });
  layer.append(node);
  setTimeout(() => {
    node.style.transition = "opacity .2s";
    node.style.opacity = "0";
    setTimeout(() => node.remove(), 220);
  }, timeout);
  return node;
}

/* ------------------------------------------------------------------- menus */
const menuLayer = () => $("#menu-layer");
let openMenus = [];

export function closeMenus() {
  for (const menu of openMenus) menu.remove();
  openMenus = [];
  menuLayer()?.classList.remove("open");
}

function buildMenu(items, options = {}) {
  const menu = el("div", { class: "menu" + (options.big ? " menu-big" : ""), role: "menu" });
  if (options.minWidth) menu.style.minWidth = options.minWidth + "px";
  if (options.maxWidth) menu.style.maxWidth = options.maxWidth + "px";

  const makeItems = (list, container) => {
    for (const item of list) {
      if (!item) continue;
      if (item.type === "sep") {
        container.append(el("div", { class: "menu-sep" }));
        continue;
      }
      if (item.type === "head") {
        container.append(el("div", { class: "menu-head", text: item.label }));
        continue;
      }
      if (item.type === "custom") {
        container.append(item.node);
        continue;
      }
      const button = el("button", { class: "menu-item", type: "button", role: "menuitem" });
      if (item.disabled) button.disabled = true;
      if (item.checked) {
        button.append(el("span", { class: "m-check", html: icon("check") }));
      } else if (options.checkable) {
        button.append(el("span", { class: "m-check" }));
      }
      if (item.icon) button.append(el("span", { class: "m-icon", html: icon(item.icon) }));
      button.append(el("span", { class: "m-label", text: item.label }));
      if (item.hint) button.append(el("span", { class: "m-sub", text: item.hint }));
      if (item.shortcut) button.append(el("span", { class: "m-key", text: item.shortcut }));
      if (item.submenu) button.append(el("span", { class: "m-sub", html: icon("chevronRight") }));
      if (item.color) {
        const dot = el("span", { class: "m-check" });
        dot.style.cssText = `width:12px;height:12px;border-radius:2px;background:${item.color};border:1px solid rgba(0,0,0,.15)`;
        button.prepend(dot);
      }
      if (item.submenu) {
        let sub = null;
        let hoverTimer = null;
        const openSub = () => {
          if (sub) return;
          sub = buildMenu(item.submenu, { ...options, submenu: true });
          menuLayer().append(sub);
          const rect = button.getBoundingClientRect();
          const subRect = sub.getBoundingClientRect();
          let left = rect.right - 2;
          let top = rect.top - 4;
          if (left + subRect.width > window.innerWidth - 8) left = rect.left - subRect.width + 2;
          if (top + subRect.height > window.innerHeight - 8) top = Math.max(8, window.innerHeight - subRect.height - 8);
          sub.style.left = left + "px";
          sub.style.top = top + "px";
          openMenus.push(sub);
          button.classList.add("focused");
        };
        const closeSub = () => {
          if (sub) {
            openMenus = openMenus.filter((m) => m !== sub);
            sub.remove();
            sub = null;
          }
          button.classList.remove("focused");
        };
        button.addEventListener("mouseenter", () => {
          clearTimeout(hoverTimer);
          openSub();
        });
        button.addEventListener("mouseleave", () => {
          hoverTimer = setTimeout(closeSub, 220);
        });
        sub?.addEventListener("mouseenter", () => clearTimeout(hoverTimer));
      } else {
        button.addEventListener("click", () => {
          if (item.disabled) return;
          closeMenus();
          item.onClick?.();
        });
      }
      container.append(button);
    }
  };

  makeItems(items, menu);
  return menu;
}

function placeMenu(menu, anchor, options) {
  const rect = anchor.getBoundingClientRect();
  const menuRect = menu.getBoundingClientRect();
  let left = options.align === "right" ? rect.right - menuRect.width : rect.left;
  let top = options.fromTop ? rect.top : rect.bottom + 3;
  if (left + menuRect.width > window.innerWidth - 8) left = window.innerWidth - menuRect.width - 8;
  if (left < 8) left = 8;
  if (top + menuRect.height > window.innerHeight - 8) top = Math.max(8, rect.top - menuRect.height - 3);
  menu.style.left = Math.round(left) + "px";
  menu.style.top = Math.round(top) + "px";
}

export function openMenu(anchor, items, options = {}) {
  closeMenus();
  const menu = buildMenu(items, options);
  menuLayer().append(menu);
  openMenus.push(menu);
  menuLayer().classList.add("open");
  placeMenu(menu, anchor, options);
  const first = menu.querySelector(".menu-item:not(:disabled)");
  menu.addEventListener("keydown", (event) => {
    const focusable = Array.from(menu.querySelectorAll(".menu-item:not(:disabled)"));
    const index = focusable.indexOf(document.activeElement);
    if (event.key === "ArrowDown") {
      event.preventDefault();
      focusable[(index + 1) % focusable.length]?.focus();
    } else if (event.key === "ArrowUp") {
      event.preventDefault();
      focusable[(index - 1 + focusable.length) % focusable.length]?.focus();
    }
  });
  menu.tabIndex = -1;
  setTimeout(() => (first || menu).focus({ preventScroll: true }), 0);
  return menu;
}

export function openContextMenu(x, y, items, options = {}) {
  closeMenus();
  const menu = buildMenu(items, options);
  menuLayer().append(menu);
  openMenus.push(menu);
  menuLayer().classList.add("open");
  const rect = menu.getBoundingClientRect();
  const left = x + rect.width > window.innerWidth - 8 ? window.innerWidth - rect.width - 8 : x;
  const top = y + rect.height > window.innerHeight - 8 ? window.innerHeight - rect.height - 8 : y;
  menu.style.left = Math.max(6, left) + "px";
  menu.style.top = Math.max(6, top) + "px";
  return menu;
}

document.addEventListener("mousedown", (event) => {
  if (openMenus.length && !event.target.closest(".menu")) closeMenus();
});
document.addEventListener("keydown", (event) => {
  if (event.key === "Escape" && openMenus.length) {
    closeMenus();
    event.stopPropagation();
  }
}, true);
window.addEventListener("blur", closeMenus);
window.addEventListener("resize", closeMenus);

/* ----------------------------------------------------------------- dialogs */
const dialogStack = [];

export function closeAllDialogs() {
  for (const entry of dialogStack.slice().reverse()) entry.close(null);
}

function topDialog() {
  return dialogStack[dialogStack.length - 1] || null;
}

document.addEventListener("keydown", (event) => {
  const entry = topDialog();
  if (!entry) return;
  const { dialog, options } = entry;
  if (event.key === "Escape" && options.escapeClose !== false) {
    event.preventDefault();
    event.stopPropagation();
    entry.close(null);
    return;
  }
  if (event.key === "Enter" && options.defaultButton !== false) {
    const active = document.activeElement;
    const tag = active?.tagName;
    if (tag === "TEXTAREA" || tag === "BUTTON" || tag === "SELECT" || active?.isContentEditable) return;
    if (active && !dialog.contains(active)) return;
    const primary = dialog.querySelector(".dialog-foot .btn.primary");
    if (primary) {
      event.preventDefault();
      primary.click();
    }
    return;
  }
  if (event.key === "Tab") {
    const focusable = Array.from(dialog.querySelectorAll(
      'button:not(:disabled), input:not([type="hidden"]):not(:disabled), select:not(:disabled), textarea:not(:disabled), [tabindex="0"]'
    )).filter((node) => node.offsetParent !== null);
    if (!focusable.length) return;
    const first = focusable[0];
    const last = focusable[focusable.length - 1];
    if (event.shiftKey && (document.activeElement === first || !dialog.contains(document.activeElement))) {
      event.preventDefault();
      last.focus();
    } else if (!event.shiftKey && document.activeElement === last) {
      event.preventDefault();
      first.focus();
    }
  }
}, true);

export function showDialog(options) {
  const layer = $("#dialog-layer");
  if (!layer) return null;
  const backdrop = el("div", { class: "dialog-backdrop" });
  const dialog = el("div", {
    class: "dialog " + (options.size || ""),
    role: "dialog",
    "aria-modal": "true",
  });
  const head = el("div", { class: "dialog-head" }, [
    el("h2", { text: options.title || "" }),
    el("button", { class: "dialog-close", title: "Close", type: "button", html: icon("close") }),
  ]);
  const body = el("div", { class: "dialog-body" });
  if (typeof options.body === "string") body.innerHTML = options.body;
  else if (options.body) body.append(options.body);

  const foot = el("div", { class: "dialog-foot" });
  dialog.append(head, body);
  if (options.buttons !== false) dialog.append(foot);
  backdrop.append(dialog);
  layer.append(backdrop);
  layer.classList.add("open");

  const entry = { backdrop, dialog, options, close: null };
  entry.close = (value) => {
    if (!backdrop.isConnected) return;
    const index = dialogStack.indexOf(entry);
    if (index >= 0) dialogStack.splice(index, 1);
    backdrop.remove();
    if (!dialogStack.length) layer.classList.remove("open");
    options.onClose?.(value);
  };
  dialogStack.push(entry);

  head.querySelector(".dialog-close").addEventListener("click", () => entry.close(null));
  if (options.backdropClose !== false) backdrop.addEventListener("mousedown", (event) => {
    if (event.target === backdrop) entry.close(null);
  });

  for (const button of options.buttons || []) {
    const node = el("button", { class: "btn " + (button.kind || ""), text: button.label, type: "button" });
    if (button.left) node.classList.add("left");
    node.addEventListener("click", async () => {
      if (button.onClick) {
        const result = await button.onClick({ close: entry.close, dialog, body });
        if (result === false) return;
      }
      if (button.close !== false) entry.close(button.value ?? button.label);
    });
    foot.append(node);
  }

  options.onOpen?.({ dialog, body, close: entry.close, foot });
  setTimeout(() => {
    const focusable = body.querySelector("input:not([type=hidden]):not([type=file]), select, textarea")
      || foot.querySelector(".btn.primary");
    focusable?.focus();
    if (focusable?.select && focusable.type !== "number" && focusable.type !== "color") focusable.select();
  }, 30);

  return { dialog, body, close: entry.close };
}

export function confirmDialog(title, message, confirmLabel = "OK", kind = "primary") {
  return new Promise((resolve) => {
    let settled = false;
    const finish = (value) => {
      if (settled) return;
      settled = true;
      resolve(value);
    };
    showDialog({
      title,
      size: "narrow",
      body: el("p", { text: message, style: "margin:0;font-size:12.5px;line-height:1.6;color:var(--text-soft)" }),
      buttons: [
        { label: "Cancel", kind: "ghost" },
        { label: confirmLabel, kind, onClick: () => finish(true) },
      ],
      onClose: () => finish(false),
    });
  });
}

export function promptDialog({ title, label, value = "", placeholder = "", type = "text", note = "" }) {
  return new Promise((resolve) => {
    let settled = false;
    const finish = (result) => {
      if (settled) return;
      settled = true;
      resolve(result);
    };
    const input = el("input", { type, value, placeholder });
    input.style.cssText = "width:100%;height:32px;border:1px solid var(--border-strong);border-radius:4px;padding:0 9px;outline:none";
    const body = el("div", {}, [
      label ? field(label, input) : input,
      note ? el("p", { class: "dialog-note", text: note }) : null,
    ]);
    showDialog({
      title,
      size: "narrow",
      body,
      buttons: [
        { label: "Cancel", kind: "ghost" },
        { label: "OK", kind: "primary", onClick: () => finish(input.value) },
      ],
      onClose: () => finish(null),
      onOpen: () => setTimeout(() => { input.focus(); input.select(); }, 40),
    });
  });
}

/* ---------------------------------------------------------------- palettes */
export const THEME_COLORS = [
  "#000000", "#FFFFFF", "#E7E6E6", "#44546A", "#4472C4", "#ED7D31",
  "#A5A5A5", "#FFC000", "#5B9BD5", "#70AD47", "#264478", "#9E480E",
  "#636363", "#997300", "#255E91", "#43682B", "#7F7F7F", "#F2F2F2",
  "#D0CECE", "#D6DCE4", "#D9E2F3", "#FBE4D5", "#EDEDED", "#FFF2CC",
  "#DDEBF7", "#E2EFDA", "#1F4E79", "#843C0C", "#375623", "#2F5597",
];

export const STANDARD_COLORS = [
  "#C00000", "#FF0000", "#FFC000", "#FFFF00", "#92D050", "#00B050",
  "#00B0F0", "#0070C0", "#002060", "#7030A0", "#FF00FF", "#A6A6A6",
  "#7F7F7F", "#404040", "#000000", "#FFFFFF", "#E7E6E6", "#F2F2F2",
];

const AUTO_COLOR = `<svg viewBox="0 0 24 24" width="16" height="16"><path d="M4 20h16" stroke="#c00000" stroke-width="3"/><path d="M8 16L11.6 5h1.2L16.4 16" fill="none" stroke="currentColor" stroke-width="1.6"/></svg>`;

export function colorMenu(anchor, { onPick, allowNone = false, title = "Theme Colors" } = {}) {
  const mk = (color) => {
    const swatch = el("button", { class: "swatch", type: "button", title: color });
    swatch.style.background = color;
    swatch.addEventListener("click", () => {
      closeMenus();
      onPick(color);
    });
    return swatch;
  };
  const grid = el("div", { class: "palette-grid" }, [
    el("div", { class: "palette-label", text: title }),
    ...THEME_COLORS.map(mk),
    el("div", { class: "palette-label", text: "Standard Colors" }),
    ...STANDARD_COLORS.map(mk),
  ]);

  const extras = [];
  if (allowNone) {
    const auto = el("button", { class: "menu-item", type: "button" }, [
      el("span", { class: "m-icon", html: AUTO_COLOR }),
      el("span", { class: "m-label", text: "Automatic" }),
    ]);
    auto.addEventListener("click", () => { closeMenus(); onPick(""); });
    extras.push(auto);
  }
  const more = el("button", { class: "menu-item", type: "button" }, [
    el("span", { class: "m-icon", html: icon("palette") }),
    el("span", { class: "m-label", text: "More Colors…" }),
  ]);
  more.addEventListener("click", () => {
    closeMenus();
    const input = el("input", { type: "color", value: "#C00000" });
    input.style.cssText = "width:100%;height:48px;border:0;padding:0;background:transparent";
    showDialog({
      title: "Colors",
      size: "narrow",
      body: input,
      buttons: [
        { label: "Cancel", kind: "ghost" },
        { label: "OK", kind: "primary", onClick: () => onPick(input.value) },
      ],
    });
  });
  extras.push(more);

  const menu = el("div", { class: "menu", style: "padding:4px" });
  menu.append(...extras, el("div", { class: "menu-sep" }), grid);
  closeMenus();
  menuLayer().append(menu);
  openMenus.push(menu);
  menuLayer().classList.add("open");
  placeMenu(menu, anchor, {});
  return menu;
}

/* ------------------------------------------------------------- small helpers */
export function formatDate(ts) {
  if (ts === null || ts === undefined || ts === "") return "";
  try {
    return new Date(ts).toLocaleString(undefined, {
      year: "numeric", month: "short", day: "numeric", hour: "2-digit", minute: "2-digit",
    });
  } catch {
    return "";
  }
}

export function debounce(fn, wait = 200) {
  let timer = null;
  return (...args) => {
    clearTimeout(timer);
    timer = setTimeout(() => fn(...args), wait);
  };
}

export function later(fn, wait = 0) {
  return setTimeout(fn, wait);
}
