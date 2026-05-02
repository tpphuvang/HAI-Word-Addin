Office.onReady((info) => {
    if (info.host === Office.HostType.Word) {
        // Gắn sự kiện khi bấm nút
        document.getElementById("btnGenerate").onclick = processDocument;
    }
});

async function processDocument() {
    const docType = document.getElementById("docType").value;
    const promptInput = document.getElementById("promptInput").value;
    const apiKey = document.getElementById("apiKey").value;
    const statusDiv = document.getElementById("status");
    const spinner = document.getElementById("loadingSpinner");
    const btn = document.getElementById("btnGenerate");

    // Rào lỗi cơ bản
    if (!apiKey) {
        showStatus("Lỗi: Vui lòng nhập Gemini API Key phía dưới!", "red");
        return;
    }
    if (!promptInput.trim()) {
        showStatus("Lỗi: Cần nhập nội dung tham mưu!", "red");
        return;
    }

    // Khóa giao diện
    btn.disabled = true;
    spinner.style.display = "block";

    try {
        // BƯỚC 1: XỬ LÝ TEMPLATE (Nếu không phải văn bản tự do)
        if (docType !== "tu_do") {
            showStatus("Đang tải biểu mẫu hành chính...", "#2b579a");
            
            // Link gọi file từ GitHub của sếp
            const templateName = docType + ".docx"; 
            const templateUrl = "https://tpphuvang.github.io/HAI-Word-Addin/" + templateName;
            
            const response = await fetch(templateUrl);
            if (!response.ok) throw new Error("Chưa tìm thấy file mẫu " + templateName + " trên GitHub. Sếp đã up chưa?");
            
            const arrayBuffer = await response.arrayBuffer();
            const base64File = arrayBufferToBase64(arrayBuffer);

            await Word.run(async (context) => {
                const body = context.document.body;
                body.clear(); // Xóa sạch tài liệu cũ
                body.insertFileFromBase64(base64File, Word.InsertLocation.start); // Chèn file mẫu vào
                await context.sync();
            });
        }

        // BƯỚC 2: GỌI GEMINI API
        showStatus("AI đang soạn nội dung lõi...", "#2b579a");
        const generatedText = await callGemini(apiKey, docType, promptInput);

        // BƯỚC 3: ĐIỀN VÀO CHỖ TRỐNG TRÊN WORD
        showStatus("Đang hoàn thiện văn bản...", "#2b579a");
        await Word.run(async (context) => {
            if (docType === "tu_do") {
                // Nếu tự do: Đổ thẳng văn bản vào trang
                const body = context.document.body;
                body.insertText(generatedText, Word.InsertLocation.end);
            } else {
                // Nếu dùng mẫu: Tìm biến {{NOI_DUNG_AI}} và thay thế
                const searchResults = context.document.body.search("{{NOI_DUNG_AI}}", { matchCase: true, matchWholeWord: true });
                context.load(searchResults, 'font');
                await context.sync();

                if (searchResults.items.length > 0) {
                    // Nếu tìm thấy chữ đó, ghi đè AI vào
                    searchResults.items[0].insertText(generatedText, Word.InsertLocation.replace);
                } else {
                    // Nếu sếp quên đặt chữ đó trong mẫu, nó sẽ tự chèn xuống cuối
                    context.document.body.insertText("\n\n" + generatedText, Word.InsertLocation.end);
                }
            }
            await context.sync();
        });

        showStatus("Thành công! Văn bản đã sẵn sàng.", "green");

    } catch (error) {
        console.error(error);
        showStatus("Lỗi hệ thống: " + error.message, "red");
    } finally {
        btn.disabled = false;
        spinner.style.display = "none";
    }
}

// Hàm chuyển đổi File sang dạng Base64
function arrayBufferToBase64(buffer) {
    let binary = '';
    const bytes = new Uint8Array(buffer);
    const len = bytes.byteLength;
    for (let i = 0; i < len; i++) {
        binary += String.fromCharCode(bytes[i]);
    }
    return window.btoa(binary);
}

// Hàm phụ trợ hiển thị trạng thái
function showStatus(message, color) {
    const statusDiv = document.getElementById("status");
    statusDiv.innerText = message;
    statusDiv.style.color = color;
}

// Hàm gọi API Gemini
async function callGemini(apiKey, type, prompt) {
    const url = "https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-pro:generateContent?key=" + apiKey;
    
    // System Prompt kết hợp (Căn dặn AI)
    const systemInstruction = `Bạn là chuyên viên hành chính tại UBND xã Phú Vinh, thành phố Huế. 
Nhiệm vụ: Soạn phần thân của văn bản (loại: ${type}) theo đúng văn phong Nghị định 30/2020/NĐ-CP.
KHÔNG viết Quốc hiệu, Tiêu ngữ, Tên cơ quan vì đã có sẵn trong file mẫu. Chỉ viết phần nội dung cốt lõi và kết luận.`;

    const requestBody = {
        contents: [{
            role: "user",
            parts: [{ text: prompt }]
        }],
        systemInstruction: {
            role: "system",
            parts: [{ text: systemInstruction }]
        },
        generationConfig: {
            temperature: 0.3 // Giữ độ sáng tạo thấp để văn phong nghiêm túc
        }
    };

    const response = await fetch(url, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(requestBody)
    });

    if (!response.ok) {
        throw new Error("Lỗi kết nối Gemini API. Có thể do Key sai hoặc hết hạn mức.");
    }

    const data = await response.json();
    return data.candidates[0].content.parts[0].text;
}
