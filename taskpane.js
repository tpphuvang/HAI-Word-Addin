/* global Word, Office */

var sessions = [];
var currentSessionId = null;
var docContext = "";
var attachedFiles = [];
var memApiKey = "";
var modalCallback = null;

// ==============================================================================
// NÃO SỐ 1: ĐẺ VĂN BẢN MỚI (CHỐT CHUẨN THỂ THỨC UBND XÃ PHÚ VINH)
// ==============================================================================
var SYSTEM_PROMPT_FULL = `Bạn là chuyên gia soạn thảo văn bản hành chính xuất sắc. BẠN BẮT BUỘC TRẢ VỀ TOÀN BỘ NỘI DUNG DƯỚI DẠNG MÃ HTML (Chỉ trả về HTML, không giải thích).
Bắt buộc áp dụng Thể thức NĐ30/2020. Dấu gạch đầu dòng bắt buộc dùng "-". 

LƯU Ý SỐNG CÒN VỀ TÊN CƠ QUAN VÀ ĐỊA DANH:
1. Ủy ban nhân dân xã Phú Vinh KHÔNG có cơ quan chủ quản ghi phía trên. Tên cơ quan ban hành đứng độc lập, in đậm là: ỦY BAN NHÂN DÂN XÃ PHÚ VINH.
2. TRỪ KHI người dùng yêu cầu soạn văn bản do một Bộ phận/Ban chuyên môn ký ban hành, thì ghi cơ quan chủ quản là "ỦY BAN NHÂN DÂN XÃ PHÚ VINH" (chữ thường), và cơ quan ban hành là tên Bộ phận/Ban đó (in hoa, đậm).
3. Địa danh ghi ở dòng ngày tháng năm bắt buộc là "Phú Vinh".

Dưới đây là khung HTML chuẩn (Hãy tự động điều chỉnh linh hoạt nếu là văn bản của phòng chuyên môn, mặc định là của UBND xã):

<table width="100%" style="border-collapse: collapse; border: none; font-family: 'Times New Roman', serif; font-size: 13pt;">
  <tr>
    <td width="40%" align="center" valign="top" style="border: none; padding: 0;">
      <b>ỦY BAN NHÂN DÂN XÃ PHÚ VINH</b><br>
      <hr style="width: 40%; border: 0.5px solid black; margin-top: 2px;">
    </td>
    <td width="60%" align="center" valign="top" style="border: none; padding: 0;">
      <b>CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM</b><br>
      <b>Độc lập - Tự do - Hạnh phúc</b><br>
      <hr style="width: 50%; border: 0.5px solid black; margin-top: 2px;">
    </td>
  </tr>
  <tr>
    <td align="center" valign="top" style="border: none; padding: 5px 0 0 0;">
      Số: .../UBND-VP
    </td>
    <td align="center" valign="top" style="border: none; padding: 5px 0 0 0;">
      <i>Phú Vinh, ngày ... tháng ... năm 202...</i>
    </td>
  </tr>
</table>
<br>
<div style="text-align: center; font-family: 'Times New Roman', serif; font-size: 14pt;">
  <b>TÊN LOẠI VĂN BẢN</b><br>
  <b>V/v: Trích yếu nội dung</b>
</div>
<br>
<div style="font-family: 'Times New Roman', serif; font-size: 14pt; text-align: justify; line-height: 1.5;">
  <p style="margin: 0 0 6pt 0;">Kính gửi: [Người/Cơ quan nhận]</p>
  <p style="text-indent: 1cm; margin: 0 0 6pt 0;">[Nội dung chi tiết của văn bản ở đây. Lùi đầu dòng 1cm. Dùng dấu "-" để liệt kê.]</p>
</div>
<br>
<table width="100%" style="border-collapse: collapse; border: none; font-family: 'Times New Roman', serif; font-size: 12pt;">
  <tr>
    <td width="50%" align="left" valign="top" style="border: none; padding: 0;">
      <b><i>Nơi nhận:</i></b><br>
      - Như trên;<br>
      - Lưu: VT.
    </td>
    <td width="50%" align="center" valign="top" style="border: none; padding: 0;">
      <b>TM. ỦY BAN NHÂN DÂN</b><br>
      <b>CHỦ TỊCH</b><br><br><br><br>
      <b>[Tên Chủ tịch]</b>
    </td>
  </tr>
</table>`;

// ==============================================================================
// NÃO SỐ 2: CHUYÊN ĐI SỬA CHỮ (Chế độ Copilot Khắc nghiệt)
// ==============================================================================
var SYSTEM_PROMPT_EDIT = `Bạn là biên tập viên hành chính nhà nước kỷ luật và khắt khe nhất. 
Nhiệm vụ của bạn: 
1. Sửa TOÀN BỘ lỗi chính tả, dấu câu, lỗi diễn đạt lủng củng.
2. Ép câu chữ về đúng văn phong hành chính: súc tích, khách quan, trang trọng, không dùng từ ngữ cảm xúc hay đa nghĩa.
3. CHỈ TRẢ VỀ mã HTML của đoạn văn bản đã sửa.
CẤM TUYỆT ĐỐI KHÔNG ĐƯỢC CHÈN QUỐC HIỆU, TIÊU NGỮ, TÊN CƠ QUAN, HAY CHỮ KÝ.
Định dạng đầu ra phải là HTML thô (dùng các thẻ <p style="text-indent: 1cm; text-align: justify; margin-bottom: 6pt;">). Chỉ trả mã HTML, không giải thích gì thêm.`;

// --- TIỆN ÍCH UI ---
function lsGet(key) { try { return localStorage.getItem(key); } catch(e) { return null; } }
function lsSet(key, val) { try { localStorage.setItem(key, val); } catch(e) {} }

function showConfirm(msg, onOk) {
  var msgEl = document.getElementById("modalMsg");
  var overlay = document.getElementById("modalOverlay");
  if(!msgEl || !overlay) return;
  msgEl.textContent = msg;
  overlay.classList.add("show");
  modalCallback = onOk;
}
function hideModal() {
  var overlay = document.getElementById("modalOverlay");
  if(overlay) overlay.classList.remove("show");
  modalCallback = null;
}
function setStatus(msg, type) {
  var el = document.getElementById("status");
  if(!el) return;
  el.textContent = msg;
  el.className = "status " + (type || "");
}

// --- KHỞI TẠO APP ---
Office.onReady(function(info) {
  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', initApp);
  } else {
    initApp();
  }
});

function initApp() {
  var savedKey = lsGet("geminiApiKey") || memApiKey;
  if (savedKey && document.getElementById("apiKey")) document.getElementById("apiKey").value = savedKey;
  
  loadSessions();
  if (sessions.length === 0) createNewSession();
  else switchSession(sessions[sessions.length - 1].id);

  attachEvent("btnSave", "click", saveKey);
  attachEvent("btnNew", "click", createNewSession);
  attachEvent("btnSend", "click", function() { sendMessage("chat"); });
  attachEvent("btnInsertLast", "click", function() { sendMessage("edit_inline"); });
  
  // NÚT XÓA CHAT ĐÃ ĐƯỢC LÀM LẠI HOÀN TOÀN
  attachEvent("btnClear", "click", function() { 
    showConfirm("Bạn muốn xóa trắng Hội thoại này?", doClearCurrentSession); 
  });
  
  attachEvent("btnReadSel", "click", readSelection);
  attachEvent("btnUpload", "click", function() {
    var fi = document.getElementById("fileInput");
    if(fi) { fi.value = ""; fi.click(); }
  });
  
  attachEvent("fileInput", "change", handleFileSelect);
  attachEvent("modalOk", "click", function() { if (modalCallback) modalCallback(); hideModal(); });
  attachEvent("modalCancel", "click", hideModal);
  attachEvent("focusHint", "click", function() {
    var prompt = document.getElementById("prompt");
    if(prompt) prompt.blur();
  });
}

function attachEvent(id, eventType, callback) {
  var el = document.getElementById(id);
  if(el) el.addEventListener(eventType, callback);
}

// --- QUẢN LÝ SESSION (FIX LỖI BÓNG MA) ---
function loadSessions() {
  try {
    var raw = lsGet("haiSessions") || "[]";
    sessions = Array.isArray(JSON.parse(raw)) ? JSON.parse(raw) : [];
  } catch(e) { sessions = []; }
}
function saveSessions() { lsSet("haiSessions", JSON.stringify(sessions)); }

function getCurrentSession() {
  for (var i = 0; i < sessions.length; i++) {
    if (sessions[i].id === currentSessionId) return sessions[i];
  }
  return null;
}

function createNewSession() {
  var id = Date.now().toString();
  sessions.push({ id: id, name: "Chat " + (sessions.length + 1), history: [], lastResponse: "" });
  saveSessions();
  switchSession(id);
}

// Hàm Xóa 1 thẻ Session khỏi danh sách
function deleteSessionUI(id, event) {
  event.stopPropagation(); // Chặn sự kiện click nhầm vào việc đổi thẻ
  showConfirm("Xóa vĩnh viễn thẻ này?", function() {
    if (sessions.length <= 1) { 
        // Nếu chỉ còn 1 thẻ, thì xóa trắng nội dung của nó thay vì xóa thẻ
        doClearCurrentSession();
        return; 
    }
    sessions = sessions.filter(function(s) { return s.id !== id; });
    saveSessions();
    if (currentSessionId === id) switchSession(sessions[sessions.length - 1].id);
    else renderSessionChips();
  });
}

function switchSession(id) {
  currentSessionId = id;
  docContext = "";
  attachedFiles = [];
  renderFilesArea();
  var btn = document.getElementById("btnReadSel");
  if(btn) btn.classList.remove("active");
  renderSessionChips();
  renderChatHistory();
}

function renderSessionChips() {
  var list = document.getElementById("sessionList");
  if(!list) return;
  list.innerHTML = "";
  sessions.forEach(function(s) {
    var wrap = document.createElement("div");
    wrap.className = "chip-wrap";
    
    var chip = document.createElement("div");
    chip.className = "session-chip" + (s.id === currentSessionId ? " active" : "");
    // Hiển thị tên rõ ràng, giới hạn độ dài
    chip.textContent = s.name.substring(0, 15) + (s.name.length > 15 ? "..." : "");
    chip.addEventListener("click", function() { switchSession(s.id); });
    wrap.appendChild(chip);
    
    var del = document.createElement("button");
    del.className = "chip-del";
    del.textContent = "x";
    del.title = "Xóa thẻ này";
    del.addEventListener("click", function(e) { deleteSessionUI(s.id, e); });
    wrap.appendChild(del);
    
    list.appendChild(wrap);
  });
}

function renderChatHistory() {
  var box = document.getElementById("chatBox");
  if(!box) return;
  box.innerHTML = "";
  var session = getCurrentSession();
  if (!session || !session.history || session.history.length === 0) return;
  
  session.history.forEach(function(msg) {
    var text = msg.parts[0].text || "";
    if (msg.role === "user") addBubble("Bạn", text, "bubble-user", null, false);
    else addBubble("HAI", text, "bubble-ai", text, false);
  });
  box.scrollTop = box.scrollHeight;
}

// Hàm XÓA SẠCH NỘI DUNG của thẻ hiện tại (Giết chết Quota Exceeded)
function doClearCurrentSession() {
  var session = getCurrentSession();
  if (!session) return;
  
  // Ép trắng mảng ngầm
  session.history = []; 
  session.lastResponse = "";
  session.name = "Chat Mới"; // Đặt lại tên mặc định
  
  docContext = ""; 
  attachedFiles = [];
  renderFilesArea(); 
  
  var btn = document.getElementById("btnReadSel");
  if(btn) btn.classList.remove("active");
  
  saveSessions(); 
  renderChatHistory(); // Render lại Box chat (sẽ trống không)
  renderSessionChips(); // Render lại danh sách thẻ (đổi tên)
  
  setStatus("Đã dọn dẹp sạch sẽ bộ nhớ!", "ok");
}

function saveKey() {
  var keyEl = document.getElementById("apiKey");
  if(!keyEl) return;
  memApiKey = keyEl.value.trim();
  lsSet("geminiApiKey", memApiKey);
  setStatus("Đã lưu khóa API!", "ok");
}

// --- THAO TÁC FILE / WORD ---
function readSelection() {
  setStatus("Đang đọc vùng bôi đen...", "");
  Word.run(function(context) {
    var sel = context.document.getSelection();
    sel.load("text");
    return context.sync().then(function() {
      docContext = sel.text.trim();
      if(docContext) {
        var btn = document.getElementById("btnReadSel");
        if(btn) btn.classList.add("active");
        setStatus("Đã nạp đoạn văn bản bạn đang bôi đen.", "ok");
      } else {
        setStatus("Bạn chưa bôi đen gì cả!", "err");
      }
    });
  }).catch(function(e) { setStatus("Lỗi: " + e.message, "err"); });
}

function insertHtmlToWord(htmlText) {
  if (!htmlText) { setStatus("Không có nội dung!", "err"); return; }
  var cleanHtml = htmlText.replace(/```html/gi, "").replace(/```/g, "").trim();

  Word.run(function(context) {
    var sel = context.document.getSelection();
    sel.insertHtml(cleanHtml, Word.InsertLocation.replace);
    return context.sync();
  })
  .then(function() { setStatus("Đã chèn thành công vào Word!", "ok"); })
  .catch(function(e) { setStatus("Lỗi chèn HTML: " + e.message, "err"); });
}

function handleFileSelect(e) {
  var files = e.target.files;
  if (!files || files.length === 0) return;
  Array.from(files).forEach(function(file) {
    var reader = new FileReader();
    reader.onload = function(ev) {
      attachedFiles.push({ base64: ev.target.result.split(",")[1], mimeType: file.type, name: file.name });
      renderFilesArea();
    };
    reader.readAsDataURL(file);
  });
}

function removeFile(idx) {
  attachedFiles.splice(idx, 1);
  renderFilesArea();
}

function renderFilesArea() {
  var area = document.getElementById("filesArea");
  if(!area) return;
  area.innerHTML = "";
  attachedFiles.forEach(function(f, idx) {
    var chip = document.createElement("div");
    chip.className = "file-chip";
    chip.innerHTML = `<span>File đính kèm</span> <button class="chip-rm" onclick="removeFile(${idx})">x</button>`;
    area.appendChild(chip);
  });
}

// --- GỬI API (TỐI ƯU QUOTA) ---
function sendMessage(mode) {
  var apiKey = document.getElementById("apiKey").value.trim();
  var prompt = document.getElementById("prompt").value.trim();
  
  if (!apiKey) { setStatus("Vui lòng nhập API Key!", "err"); return; }
  
  if (mode === "edit_inline") {
    // CHẾ ĐỘ COPILOT
    Word.run(function(context) {
      var sel = context.document.getSelection();
      sel.load("text");
      return context.sync().then(function() {
        if (!sel.text.trim()) {
          setStatus("Vui lòng bôi đen đoạn văn bản cần sửa trước!", "err");
          return;
        }
        var copilotPrompt = "Chỉ sửa đoạn văn sau theo yêu cầu: [" + (prompt || "Sửa lỗi chính tả, câu cú lủng củng") + "]. Văn bản gốc:\n\n" + sel.text;
        callGeminiAPI(apiKey, copilotPrompt, true);
      });
    });
  } else {
    // CHẾ ĐỘ CHAT
    if (!prompt && attachedFiles.length === 0) return;
    addBubble("Bạn", prompt, "bubble-user", null, true);
    document.getElementById("prompt").value = "";
    
    var fullText = prompt;
    if (docContext) fullText = "Ngữ cảnh:\n" + docContext + "\n\nYêu cầu: " + prompt;
    callGeminiAPI(apiKey, fullText, false);
  }
}

function callGeminiAPI(apiKey, promptText, isCopilotMode) {
  showThinking();
  var session = getCurrentSession();
  
  // TỐI ƯU QUOTA: Xây dựng mảng nội dung gửi đi (Chỉ gửi 1 câu hiện tại)
  var parts = [{ text: promptText }];
  attachedFiles.forEach(function(file) {
    parts.push({ inline_data: { mime_type: file.mimeType, data: file.base64 } });
  });
  
  var activeSystemPrompt = isCopilotMode ? SYSTEM_PROMPT_EDIT : SYSTEM_PROMPT_FULL;
  var url = "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=" + apiKey;
  
  // TỐI ƯU QUOTA: Bỏ 'session.history' ra khỏi lệnh gửi. Chỉ gửi Zero-shot.
  var bodyObj = {
    system_instruction: { parts: [{ text: activeSystemPrompt }] },
    contents: [{ role: "user", parts: parts }] 
  };
  
  fetch(url, { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(bodyObj) })
  .then(res => res.json())
  .then(data => {
    removeThinking();
    if (data.error) {
      setStatus("Lỗi API: " + data.error.message.substring(0, 50), "err");
      return;
    }
    
    var reply = data.candidates[0].content.parts.map(p => p.text).join("");
    
    if (isCopilotMode) {
      insertHtmlToWord(reply);
      setStatus("Đã biến hóa xong vùng bôi đen!", "ok");
    } else {
      session.lastResponse = reply;
      
      // Đặt tên Session tự động nếu đây là câu đầu tiên
      if (session.history.length === 0 && promptText.length > 5) {
        var newName = promptText.replace("Ngữ cảnh:\n" + docContext + "\n\nYêu cầu: ", "");
        session.name = newName.substring(0, 20) + "...";
      }
      
      session.history.push({ role: "user", parts: [{ text: promptText }] });
      session.history.push({ role: "model", parts: [{ text: reply }] });
      saveSessions();
      
      renderSessionChips();
      addBubble("HAI", "Đã tạo xong khung văn bản.", "bubble-ai", reply, true);
      setStatus("Hoàn thành! Bấm chèn để xuất ra Word.", "ok");
    }
    
    attachedFiles = [];
    renderFilesArea();
    docContext = ""; 
    var btn = document.getElementById("btnReadSel");
    if(btn) btn.classList.remove("active");
  })
  .catch(e => {
    removeThinking();
    setStatus("Lỗi mạng: " + e.message, "err");
  });
}

function addBubble(sender, text, cssClass, insertText, scroll) {
  var box = document.getElementById("chatBox");
  if(!box) return;
  var div = document.createElement("div");
  div.className = "bubble " + cssClass;
  
  var labelDiv = document.createElement("div");
  labelDiv.className = "bubble-label";
  labelDiv.innerHTML = `<span>${sender}</span>`;
  
  if (cssClass === "bubble-ai" && insertText) {
    var actions = document.createElement("div");
    actions.className = "bubble-actions";
    actions.innerHTML = `<button class="btn-bubble green" onclick='insertHtmlToWord(getCurrentSession().lastResponse)'>Chèn Toàn Bộ Form</button>`;
    labelDiv.appendChild(actions);
  }
  
  div.appendChild(labelDiv);
  var textDiv = document.createElement("div");
  textDiv.className = "bubble-text";
  textDiv.textContent = text.substring(0, 150) + (text.length > 150 ? "... (Đã ẩn mã HTML)" : "");
  div.appendChild(textDiv);
  box.appendChild(div);
  if (scroll) box.scrollTop = box.scrollHeight;
}

function showThinking() {
  var box = document.getElementById("chatBox");
  if(!box) return;
  var div = document.createElement("div");
  div.className = "bubble bubble-ai thinking-bubble";
  div.innerHTML = `<div class="bubble-label"><span>Đang xử lý</span></div><div class="thinking"><div class="dot"></div><div class="dot"></div><div class="dot"></div></div>`;
  box.appendChild(div);
  box.scrollTop = box.scrollHeight;
}
function removeThinking() {
  var t = document.querySelector(".thinking-bubble");
  if (t) t.remove();
}