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

        // GỌI HÀM XỬ LÝ KẾT QUẢ MỚI
        await handleAIOutput(generatedText);

    } catch (error) {
        console.error(error);
        showStatus("Lỗi hệ thống: " + error.message, "red");
    } finally {
        btn.disabled = false;
        spinner.style.display = "none";
    }
};
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
// --- HÀM XỬ LÝ KẾT QUẢ: EDITOR HAY CHAT ---
async function handleAIOutput(generatedText) {
    const mode = document.querySelector('input[name="mode"]:checked').value;
    const chatResult = document.getElementById("chatResult");
    const btnInsertManual = document.getElementById("btnInsertManual");

    if (mode === "chat") {
        chatResult.innerText = generatedText;
        chatResult.style.display = "block";
        btnInsertManual.style.display = "block";
        btnInsertManual.onclick = () => insertToWord(generatedText); 
        showStatus("Đã soạn xong! Xem ở khung dưới.", "#2b579a");
    } else {
        chatResult.style.display = "none";
        btnInsertManual.style.display = "none";
        await insertToWord(generatedText);
        showStatus("Đã chèn văn bản thành công!", "green");
    }
}

// --- HÀM CHÈN THÔNG MINH ---
async function insertToWord(text) {
    await Word.run(async (context) => {
        const docType = document.getElementById("docType").value;
        
        if (docType !== "tu_do") {
            const searchResults = context.document.body.search("{{NOI_DUNG_AI}}", { matchCase: true });
            context.load(searchResults);
            await context.sync();
            
            if (searchResults.items.length > 0) {
                searchResults.items[0].insertText(text, Word.InsertLocation.replace);
                await context.sync();
                return;
            }
        }
        
        const selection = context.document.getSelection();
        selection.insertText(text, Word.InsertLocation.replace); 
        await context.sync();
    });
}
