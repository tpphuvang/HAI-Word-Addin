// Thiết lập mặc định khi khởi động
Office.onReady((info) => {
    if (info.host === Office.HostType.Word) {
        document.getElementById("btnGenerate").onclick = mainProcess;
    }
});

// --- HÀM ĐIỀU PHỐI CHÍNH ---
async function mainProcess() {
    const apiKey = document.getElementById("apiKey").value;
    const docType = document.getElementById("docType").value;
    const promptInput = document.getElementById("promptInput").value;
    const mode = document.querySelector('input[name="mode"]:checked').value;
    
    const btn = document.getElementById("btnGenerate");
    const spinner = document.getElementById("loadingSpinner");

    if (!apiKey || !promptInput.trim()) {
        updateStatus("Thiếu API Key hoặc Nội dung!", "red");
        return;
    }

    btn.disabled = true;
    spinner.style.display = "block";

    try {
        // 1. Xử lý Template (nếu có)
        if (docType !== "tu_do") {
            updateStatus("Đang tải mẫu chuẩn...", "#2b579a");
            await loadTemplate(docType);
        }

        // 2. Gọi AI lấy nội dung
        updateStatus("AI đang tham mưu nội dung...", "#2b579a");
        const generatedText = await callAI(apiKey, docType, promptInput);

        // 3. Phân luồng Editor hay Chat
        if (mode === "chat") {
            showChatBox(generatedText);
            updateStatus("Đã soạn xong! Xem ở khung Chat.", "#2b579a");
        } else {
            await insertToWord(generatedText);
            updateStatus("Đã chèn trực tiếp vào văn bản!", "green");
        }

    } catch (err) {
        updateStatus("Lỗi: " + err.message, "red");
    } finally {
        btn.disabled = false;
        spinner.style.display = "none";
    }
}

// --- CÁC HÀM BỔ TRỢ ---

async function loadTemplate(type) {
    const url = `https://tpphuvang.github.io/HAI-Word-Addin/${type}.docx`;
    const res = await fetch(url);
    if (!res.ok) throw new Error("Không thấy file mẫu trên GitHub. Check lại tên hoa/thường!");
    
    const buffer = await res.arrayBuffer();
    const base64 = arrayBufferToBase64(buffer);

    await Word.run(async (context) => {
        context.document.body.clear();
        context.document.body.insertFileFromBase64(base64, "Start");
        await context.sync();
    });
}

async function callAI(key, type, prompt) {
    // --- KHU VỰC CẤU HÌNH (Sếp chỉ cần đổi ở đây) ---
    const provider = "gemini"; // Có thể đổi thành: "claude", "perplexity", hoặc "local"
    
    let url = "";
    let headers = { "Content-Type": "application/json" };
    let body = {};

    const systemMsg = `Bạn là trợ lý hành chính xã Phú Vinh. Soạn nội dung cho ${type}. KHÔNG viết Header/Footer.`;

    // --- LOGIC TỰ ĐỘNG XOAY TRỤC ---
    if (provider === "gemini") {
        url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=${key}`;
        body = {
            contents: [{ parts: [{ text: prompt }] }],
            systemInstruction: { parts: [{ text: systemMsg }] }
        };
    } 
    else if (provider === "claude") {
        url = "https://api.anthropic.com/v1/messages";
        headers["x-api-key"] = key;
        headers["anthropic-version"] = "2023-06-01";
        body = {
            model: "claude-3-sonnet-20240229",
            max_tokens: 2000,
            system: systemMsg,
            messages: [{ role: "user", content: prompt }]
        };
    }
    else if (provider === "perplexity" || provider === "local") {
        // Chuẩn OpenAI (Perplexity và LM Studio dùng chung chuẩn này)
        url = (provider === "local") ? "http://IP_MAY_KIA:1234/v1/chat/completions" : "https://api.perplexity.ai/chat/completions";
        headers["Authorization"] = `Bearer ${key}`;
        body = {
            model: (provider === "local") ? "local-model" : "pplx-7b-online",
            messages: [
                { role: "system", content: systemMsg },
                { role: "user", content: prompt }
            ]
        };
    }

    // --- THỰC THI GỬI TIN ---
    const res = await fetch(url, { method: "POST", headers: headers, body: JSON.stringify(body) });
    if (!res.ok) throw new Error("Lỗi kết nối API. Check lại Key hoặc Hạn mức.");
    
    const data = await res.json();
    
    // Trích xuất kết quả tùy theo hãng (vì mỗi ông nhả JSON một kiểu)
    if (provider === "gemini") return data.candidates[0].content.parts[0].text;
    if (provider === "claude") return data.content[0].text;
    return data.choices[0].message.content; // Cho Perplexity và LM Studio
}

async function insertToWord(text) {
    await Word.run(async (context) => {
        const body = context.document.body;
        // Tìm và thay thế tất cả placeholder
        const searchResults = body.search("{{NOI_DUNG_AI}}", { matchCase: true });
        context.load(searchResults, 'font');
        await context.sync();

        if (searchResults.items.length > 0) {
            for (let item of searchResults.items) {
                item.insertText(text, "Replace");
                formatFont(item);
            }
        } else {
            // Nếu không có placeholder, chèn tại con trỏ
            const sel = context.document.getSelection();
            sel.insertText(text, "Replace");
            formatFont(sel);
        }
        await context.sync();
    });
}

function formatFont(range) {
    range.font.name = "Times New Roman";
    range.font.size = 14;
    range.font.color = "black";
}

function showChatBox(text) {
    const box = document.getElementById("chatResult");
    const btn = document.getElementById("btnInsertManual");
    box.innerText = text;
    box.style.display = "block";
    btn.style.display = "block";
    btn.onclick = () => insertToWord(text);
}

function updateStatus(msg, color) {
    const s = document.getElementById("status");
    s.innerText = msg;
    s.style.color = color;
}

function arrayBufferToBase64(buffer) {
    let binary = '';
    let bytes = new Uint8Array(buffer);
    for (let i = 0; i < bytes.byteLength; i++) binary += String.fromCharCode(bytes[i]);
    return window.btoa(binary);
}
