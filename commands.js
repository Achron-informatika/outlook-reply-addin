/*
 * Reply with Attachments - Outlook add-in (function commands)
 *
 * What it does when a user clicks a button:
 *   1. Reads the attachments of the selected email (skips inline images such as signature logos).
 *   2. Downloads each attachment's content from Exchange via Office.js (nothing leaves Microsoft 365).
 *   3. Opens a normal Reply / Reply All form with those attachments added (Mailbox requirement set 1.15).
 *
 * Written in plain ES5 (no async/await, no arrow functions) so it parses in every Outlook client.
 */
(function () {
  "use strict";

  var NOTIFY_ID = "rwaStatus";

  // ---------- localisation ----------
  var STRINGS = {
    en: {
      preparing: "Preparing attachments…",
      unsupported: "This Outlook version can't reply with attachments. Please update Outlook or use new Outlook / Outlook on the web.",
      skipped: "Some attachments could not be copied (e.g. cloud links or blocked files): ",
      error: "Couldn't open the reply with attachments: ",
      noItem: "Select a single email first."
    },
    hu: {
      preparing: "Mellékletek előkészítése…",
      unsupported: "Ez az Outlook-verzió nem támogatja a mellékletekkel való választ. Frissítse az Outlookot, vagy használja az új Outlookot / Outlook webes verzióját.",
      skipped: "Néhány melléklet nem másolható (pl. felhőhivatkozás vagy tiltott fájl): ",
      error: "Nem sikerült megnyitni a választ a mellékletekkel: ",
      noItem: "Előbb jelöljön ki egyetlen e-mailt."
    }
  };

  function t(key) {
    var lang = "en";
    try {
      var dl = (Office.context.displayLanguage || "").toLowerCase();
      if (dl.indexOf("hu") === 0) { lang = "hu"; }
    } catch (e) { /* ignore */ }
    return STRINGS[lang][key] || STRINGS.en[key];
  }

  // ---------- notifications (info bar above the email) ----------
  function notify(item, type, text) {
    if (!item || !item.notificationMessages) { return; }
    if (text.length > 150) { text = text.substring(0, 147) + "..."; }
    var msg;
    if (type === "error") {
      msg = { type: Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage, message: text };
    } else {
      msg = {
        type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
        message: text,
        icon: "Reply.16",
        persistent: false
      };
    }
    try { item.notificationMessages.replaceAsync(NOTIFY_ID, msg); } catch (e) { /* ignore */ }
  }

  function clearNotify(item) {
    try { item.notificationMessages.removeAsync(NOTIFY_ID); } catch (e) { /* ignore */ }
  }

  // ---------- helpers ----------
  function utf8ToBase64(str) {
    // works in all browsers, including legacy ones
    return window.btoa(unescape(encodeURIComponent(str)));
  }

  function looksLikeBase64(str) {
    if (!str || str.length < 8) { return false; }
    var sample = str.substring(0, 2000).replace(/[\r\n]/g, "");
    return /^[A-Za-z0-9+\/]+={0,2}$/.test(sample) && sample.indexOf(" ") === -1;
  }

  function ensureExtension(name, ext) {
    name = (name || "attachment").replace(/[\\\/:*?"<>|]/g, "_");
    if (name.toLowerCase().slice(-ext.length) !== ext) { name += ext; }
    return name;
  }

  function isCloud(att) {
    var cloud = Office.MailboxEnums.AttachmentType && Office.MailboxEnums.AttachmentType.Cloud;
    return att.attachmentType === cloud || att.attachmentType === "cloud";
  }

  // Convert one AttachmentContent result into a ReplyFormAttachment (or null if not possible)
  function toReplyAttachment(att, content) {
    var F = Office.MailboxEnums.AttachmentContentFormat;
    var fmt = content.format;
    if (fmt === F.Base64 || fmt === "base64") {
      return { type: "base64", name: att.name, base64file: content.content };
    }
    if (fmt === F.Eml || fmt === "eml") {
      // attached email (.msg / .eml) -> re-attach as .eml
      var eml = looksLikeBase64(content.content) ? content.content.replace(/[\r\n]/g, "") : utf8ToBase64(content.content);
      return { type: "base64", name: ensureExtension(att.name, ".eml"), base64file: eml };
    }
    if (fmt === F.ICalendar || fmt === "iCalendar") {
      var ics = looksLikeBase64(content.content) ? content.content.replace(/[\r\n]/g, "") : utf8ToBase64(content.content);
      return { type: "base64", name: ensureExtension(att.name, ".ics"), base64file: ics };
    }
    return null; // "url" = cloud attachment, can't be copied as a file
  }

  // Sequentially fetch attachment contents (keeps memory use and request load low)
  function collect(item, list, index, out, skipped, done) {
    if (index >= list.length) { done(out, skipped); return; }
    var att = list[index];
    if (isCloud(att)) {
      skipped.push(att.name);
      collect(item, list, index + 1, out, skipped, done);
      return;
    }
    item.getAttachmentContentAsync(att.id, function (res) {
      if (res.status === Office.AsyncResultStatus.Succeeded) {
        var r = toReplyAttachment(att, res.value);
        if (r) { out.push(r); } else { skipped.push(att.name); }
      } else {
        skipped.push(att.name);
      }
      collect(item, list, index + 1, out, skipped, done);
    });
  }

  function openReply(item, replyAll, attachments, cb) {
    var formData = { attachments: attachments };
    var fnAsync = replyAll ? item.displayReplyAllFormAsync : item.displayReplyFormAsync;
    try {
      if (typeof fnAsync === "function") {
        fnAsync.call(item, formData, function (res) {
          if (res.status === Office.AsyncResultStatus.Failed) {
            cb(res.error ? res.error.message : "unknown error");
          } else { cb(null); }
        });
      } else {
        if (replyAll) { item.displayReplyAllForm(formData); } else { item.displayReplyForm(formData); }
        cb(null);
      }
    } catch (e) {
      cb(e && e.message ? e.message : String(e));
    }
  }

  // ---------- main ----------
  function run(event, replyAll) {
    var finished = false;
    function complete() {
      if (finished) { return; }
      finished = true;
      try { event.completed(); } catch (e) { /* ignore */ }
    }

    var item;
    try { item = Office.context.mailbox.item; } catch (e) { item = null; }
    if (!item) { complete(); return; }

    try {
      if (!Office.context.requirements.isSetSupported("Mailbox", "1.15")) {
        notify(item, "error", t("unsupported"));
        complete();
        return;
      }

      var all = item.attachments || [];
      var real = [];
      for (var i = 0; i < all.length; i++) {
        if (!all[i].isInline) { real.push(all[i]); }
      }

      // No attachments -> behave like a normal reply
      if (real.length === 0) {
        openReply(item, replyAll, [], function (err) {
          if (err) { notify(item, "error", t("error") + err); }
          complete();
        });
        return;
      }

      notify(item, "info", t("preparing"));

      collect(item, real, 0, [], [], function (attachments, skipped) {
        openReply(item, replyAll, attachments, function (err) {
          if (err) {
            notify(item, "error", t("error") + err);
          } else if (skipped.length) {
            notify(item, "info", t("skipped") + skipped.join(", "));
          } else {
            clearNotify(item);
          }
          complete();
        });
      });
    } catch (e) {
      notify(item, "error", t("error") + (e && e.message ? e.message : String(e)));
      complete();
    }
  }

  function replyWithAttachments(event) { run(event, false); }
  function replyAllWithAttachments(event) { run(event, true); }

  // Expose for XML-manifest function commands (global lookup) ...
  window.replyWithAttachments = replyWithAttachments;
  window.replyAllWithAttachments = replyAllWithAttachments;

  // ... and register via the modern API where available
  Office.onReady(function () {
    if (Office.actions && typeof Office.actions.associate === "function") {
      Office.actions.associate("replyWithAttachments", replyWithAttachments);
      Office.actions.associate("replyAllWithAttachments", replyAllWithAttachments);
    }
  });
})();
