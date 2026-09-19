(function () {
    "use strict";

    // ===== PDF.js worker config =====
    if (window.pdfjsLib) {
        pdfjsLib.GlobalWorkerOptions.workerSrc = "js/lib/pdf.worker.min.js";
    }

    // ===== DOM elements =====
    var uploadScreen = document.getElementById("uploadScreen");
    var editorScreen = document.getElementById("editorScreen");
    var dropZone = document.getElementById("dropZone");
    var fileInput = document.getElementById("fileInput");
    var fileInfo = document.getElementById("fileInfo");
    var fileName = document.getElementById("fileName");
    var progressBar = document.getElementById("progressBar");
    var confirmBtn = document.getElementById("confirmBtn");
    var errorMsg = document.getElementById("errorMsg");
    var backBtn = document.getElementById("backBtn");
    var docPaper = document.getElementById("docPaper");
    var pagesWrapper = document.getElementById("pagesWrapper");
    var selectionOverlay = document.getElementById("selectionOverlay");
    var signaturesLayer = document.getElementById("signaturesLayer");
    var toolbarHint = document.getElementById("toolbarHint");
    var loadingHint = document.getElementById("loadingHint");
    var modal = document.getElementById("signatureModal");
    var closePadBtn = document.getElementById("closePadBtn");
    var canvas = document.getElementById("signatureCanvas");
    var canvasHint = document.getElementById("canvasHint");
    var clearBtn = document.getElementById("clearBtn");
    var undoBtn = document.getElementById("undoBtn");
    var confirmSignBtn = document.getElementById("confirmSignBtn");
    var downloadBar = document.getElementById("downloadBar");
    var downloadBtn = document.getElementById("downloadBtn");
    var docStatus = document.getElementById("docStatus");
    var ctx = canvas.getContext("2d");

    // ===== State =====
    var selectedFile = null;
    var fileType = null;
    var originalFileBytes = null;
    var pageElements = [];
    var signatures = [];
    var docReady = false;

    // Signature canvas state
    var drawing = false;
    var lastX = 0, lastY = 0;
    var strokes = [];
    var currentStroke = [];
    var hasDrawn = false;

    // Selection state
    var isSelecting = false;
    var selStartX = 0, selStartY = 0;
    var currentSelBox = null;
    var activeSelection = null;

    // ===== Upload Screen =====

    dropZone.addEventListener("click", function () { fileInput.click(); });

    dropZone.addEventListener("dragover", function (e) {
        e.preventDefault();
        dropZone.classList.add("drag-over");
    });
    dropZone.addEventListener("dragleave", function () { dropZone.classList.remove("drag-over"); });
    dropZone.addEventListener("drop", function (e) {
        e.preventDefault();
        dropZone.classList.remove("drag-over");
        if (e.dataTransfer.files.length > 0) handleFile(e.dataTransfer.files[0]);
    });
    fileInput.addEventListener("change", function () {
        if (fileInput.files.length > 0) handleFile(fileInput.files[0]);
    });

    function handleFile(file) {
        var ext = file.name.split(".").pop().toLowerCase();
        if (ext !== "docx" && ext !== "pdf") {
            showError("仅支持 .docx 和 .pdf 格式");
            return;
        }
        selectedFile = file;
        fileName.textContent = file.name;
        fileInfo.style.display = "block";
        confirmBtn.style.display = "block";
        errorMsg.style.display = "none";
    }

    function showError(msg) {
        errorMsg.textContent = msg;
        errorMsg.style.display = "block";
    }

    confirmBtn.addEventListener("click", function () {
        if (!selectedFile) return;
        showEditor();
    });

    // ===== Editor Screen =====

    function showEditor() {
        uploadScreen.style.display = "none";
        editorScreen.style.display = "block";
        loadingHint.style.display = "block";
        toolbarHint.style.display = "none";

        var ext = selectedFile.name.split(".").pop().toLowerCase();
        fileType = ext;

        if (ext === "pdf") {
            renderPdf(selectedFile);
        } else {
            renderDocx(selectedFile);
        }
    }

    backBtn.addEventListener("click", function () {
        // Reset everything
        editorScreen.style.display = "none";
        uploadScreen.style.display = "block";
        pagesWrapper.innerHTML = "";
        selectionOverlay.innerHTML = "";
        signaturesLayer.innerHTML = "";
        signatures = [];
        pageElements = [];
        docReady = false;
        downloadBar.style.display = "none";
        docStatus.textContent = "待签字";
        docStatus.classList.remove("signed");
        fileInfo.style.display = "none";
        confirmBtn.style.display = "none";
        fileInput.value = "";
        selectedFile = null;
    });

    // ===== PDF Rendering (PDF.js) =====

    async function renderPdf(file) {
        try {
            originalFileBytes = await file.arrayBuffer();
            var pdf = await pdfjsLib.getDocument({ data: originalFileBytes.slice(0) }).promise;

            for (var i = 1; i <= pdf.numPages; i++) {
                var page = await pdf.getPage(i);
                var viewport = page.getViewport({ scale: 1.5 });

                var pageCanvas = document.createElement("canvas");
                pageCanvas.width = viewport.width;
                pageCanvas.height = viewport.height;
                pageCanvas.className = "page-canvas";

                var pageCtx = pageCanvas.getContext("2d");
                await page.render({ canvasContext: pageCtx, viewport: viewport }).promise;

                var pageDiv = document.createElement("div");
                pageDiv.className = "pdf-page";
                pageDiv.setAttribute("data-page", i - 1);
                pageDiv.appendChild(pageCanvas);
                pagesWrapper.appendChild(pageDiv);

                pageElements.push({
                    el: pageDiv,
                    canvas: pageCanvas,
                    pageNum: i - 1,
                    naturalW: viewport.width,
                    naturalH: viewport.height,
                    isDocx: false,
                });
            }

            loadingHint.style.display = "none";
            toolbarHint.style.display = "block";
            toolbarHint.textContent = "在文档上拖拽框选签字区域，然后点击框选区签名";
            docReady = true;
        } catch (err) {
            loadingHint.textContent = "PDF 渲染失败: " + err.message;
            loadingHint.style.color = "#ef4444";
        }
    }

    // ===== DOCX Rendering (docx-preview) =====

    async function renderDocx(file) {
        try {
            var buffer = await file.arrayBuffer();
            originalFileBytes = buffer;

            var container = document.createElement("div");
            container.id = "docxContainer";
            container.className = "docx-container";
            pagesWrapper.appendChild(container);

            await window.docx.renderAsync(buffer, container, null, {
                className: "docx-page",
                inWrapper: true,
                ignoreWidth: false,
                ignoreHeight: false,
                breakPages: true,
            });

            setTimeout(function () {
                loadingHint.style.display = "none";
                toolbarHint.style.display = "block";
                toolbarHint.textContent = "在文档上拖拽框选签字区域，然后点击框选区签名";
                collectDocxPages();
                docReady = true;
            }, 300);
        } catch (err) {
            loadingHint.textContent = "文档渲染失败: " + err.message;
            loadingHint.style.color = "#ef4444";
        }
    }

    function collectDocxPages() {
        var pages = docPaper.querySelectorAll(".docx-wrapper .docx-page, .docx-page");
        pageElements = [];
        if (pages.length > 0) {
            pages.forEach(function (page, idx) {
                pageElements.push({
                    el: page,
                    pageNum: idx,
                    naturalW: page.offsetWidth,
                    naturalH: page.offsetHeight,
                    isDocx: true,
                });
            });
        } else {
            pageElements.push({
                el: docPaper,
                pageNum: 0,
                naturalW: docPaper.offsetWidth,
                naturalH: docPaper.offsetHeight,
                isDocx: true,
            });
        }
    }

    // ===== Selection Logic =====

    function getPointInPaper(e) {
        var rect = docPaper.getBoundingClientRect();
        var clientX, clientY;
        if (e.touches && e.touches.length > 0) {
            clientX = e.touches[0].clientX;
            clientY = e.touches[0].clientY;
        } else {
            clientX = e.clientX;
            clientY = e.clientY;
        }
        return {
            x: clientX - rect.left,
            y: clientY - rect.top,
        };
    }

    function findPageAt(y) {
        for (var i = 0; i < pageElements.length; i++) {
            var pe = pageElements[i];
            var top = pe.el.offsetTop;
            var height = pe.el.offsetHeight;
            if (y >= top && y < top + height) return pe;
        }
        if (pageElements.length > 0) {
            var best = pageElements[0];
            var bestDist = Math.abs(pageElements[0].el.offsetTop - y);
            for (var j = 1; j < pageElements.length; j++) {
                var d = Math.abs(pageElements[j].el.offsetTop - y);
                if (d < bestDist) { best = pageElements[j]; bestDist = d; }
            }
            return best;
        }
        return null;
    }

    function startSelection(e) {
        if (e.target.closest && (e.target.closest(".sel-box") || e.target.closest(".sig-placed"))) return;
        if (!docReady) return;

        var pt = getPointInPaper(e);
        isSelecting = true;
        selStartX = pt.x;
        selStartY = pt.y;

        currentSelBox = document.createElement("div");
        currentSelBox.className = "sel-box drawing";
        currentSelBox.style.left = pt.x + "px";
        currentSelBox.style.top = pt.y + "px";
        currentSelBox.style.width = "0px";
        currentSelBox.style.height = "0px";
        selectionOverlay.appendChild(currentSelBox);

        e.preventDefault();
    }

    function updateSelection(e) {
        if (!isSelecting || !currentSelBox) return;
        e.preventDefault();
        var pt = getPointInPaper(e);
        var x = Math.min(selStartX, pt.x);
        var y = Math.min(selStartY, pt.y);
        var w = Math.abs(pt.x - selStartX);
        var h = Math.abs(pt.y - selStartY);
        currentSelBox.style.left = x + "px";
        currentSelBox.style.top = y + "px";
        currentSelBox.style.width = w + "px";
        currentSelBox.style.height = h + "px";
    }

    function endSelection(e) {
        if (!isSelecting || !currentSelBox) return;
        isSelecting = false;

        var w = parseFloat(currentSelBox.style.width);
        var h = parseFloat(currentSelBox.style.height);
        var x = parseFloat(currentSelBox.style.left);
        var y = parseFloat(currentSelBox.style.top);

        if (w < 40 || h < 30) {
            currentSelBox.remove();
            currentSelBox = null;
            return;
        }

        var pageEl = findPageAt(y + h / 2);
        if (!pageEl) {
            currentSelBox.remove();
            currentSelBox = null;
            return;
        }

        var pageTop = pageEl.el.offsetTop;
        var pageLeft = pageEl.el.offsetLeft;
        var relX = x - pageLeft;
        var relY = y - pageTop;
        var displayW = pageEl.el.offsetWidth;
        var displayH = pageEl.el.offsetHeight;

        var selId = "sel-" + Date.now();
        currentSelBox.id = selId;
        currentSelBox.classList.remove("drawing");
        currentSelBox.classList.add("active");

        var signBtn = document.createElement("div");
        signBtn.className = "sel-sign-btn";
        signBtn.innerHTML = "<span>在此签字</span>";
        currentSelBox.appendChild(signBtn);

        var delBtn = document.createElement("div");
        delBtn.className = "sel-del-btn";
        delBtn.innerHTML = "&times;";
        currentSelBox.appendChild(delBtn);

        // Direct click on sign button - most reliable
        signBtn.addEventListener("click", function (ev) {
            ev.stopPropagation();
            ev.preventDefault();
            if (!selData.signed) {
                activeSelection = selData;
                openModal();
            }
        });

        signBtn.addEventListener("touchstart", function (ev) {
            ev.stopPropagation();
            ev.preventDefault();
            if (!selData.signed) {
                activeSelection = selData;
                openModal();
            }
        }, { passive: false });

        var selData = {
            id: selId,
            x: x, y: y, w: w, h: h,
            page: pageEl.pageNum,
            relX: relX, relY: relY,
            displayW: displayW, displayH: displayH,
            isDocx: pageEl.isDocx,
            el: currentSelBox,
            signed: false,
            dataUrl: null,
        };
        selections.push(selData);

        currentSelBox.addEventListener("click", function (ev) {
            if (ev.target.closest && ev.target.closest(".sel-del-btn")) return;
            if (!selData.signed) {
                ev.stopPropagation();
                ev.preventDefault();
                activeSelection = selData;
                openModal();
            }
        });

        currentSelBox.addEventListener("mousedown", function (ev) {
            if (ev.target.closest && ev.target.closest(".sel-del-btn")) return;
            if (!selData.signed) {
                ev.stopPropagation();
                ev.preventDefault();
            }
        });

        currentSelBox.addEventListener("touchstart", function (ev) {
            if (ev.target.closest && ev.target.closest(".sel-del-btn")) return;
            if (!selData.signed) {
                ev.stopPropagation();
                ev.preventDefault();
                activeSelection = selData;
                openModal();
            }
        }, { passive: false });

        delBtn.addEventListener("click", function (ev) {
            ev.stopPropagation();
            ev.preventDefault();
            selData.el.remove();
            var placed = document.querySelector('[data-sel-id="' + selId + '"]');
            if (placed) placed.remove();
            selections = selections.filter(function (s) { return s.id !== selId; });
            updateStatus();
        });

        delBtn.addEventListener("touchstart", function (ev) {
            ev.stopPropagation();
            ev.preventDefault();
            selData.el.remove();
            var placed = document.querySelector('[data-sel-id="' + selId + '"]');
            if (placed) placed.remove();
            selections = selections.filter(function (s) { return s.id !== selId; });
            updateStatus();
        }, { passive: false });

        currentSelBox = null;
        toolbarHint.textContent = "已创建框选区域，点击「在此签字」开始签名";
    }

    // ===== Signature Canvas =====

    function resizeCanvas() {
        var rect = canvas.getBoundingClientRect();
        var ratio = window.devicePixelRatio || 1;
        canvas.width = rect.width * ratio;
        canvas.height = rect.height * ratio;
        ctx.scale(ratio, ratio);
        ctx.lineCap = "round";
        ctx.lineJoin = "round";
        ctx.strokeStyle = "#1e293b";
        ctx.lineWidth = 2.5;
        ctx.fillStyle = "#fff";
        ctx.fillRect(0, 0, canvas.width, canvas.height);
        redrawAll();
    }

    function getCanvasPoint(e) {
        var rect = canvas.getBoundingClientRect();
        if (e.touches && e.touches.length > 0) {
            return { x: e.touches[0].clientX - rect.left, y: e.touches[0].clientY - rect.top };
        }
        return { x: e.clientX - rect.left, y: e.clientY - rect.top };
    }

    function startDraw(e) {
        e.preventDefault();
        drawing = true;
        var p = getCanvasPoint(e);
        lastX = p.x; lastY = p.y;
        currentStroke = [{ x: p.x, y: p.y }];
        canvasHint.classList.add("hidden");
    }

    function draw(e) {
        if (!drawing) return;
        e.preventDefault();
        var p = getCanvasPoint(e);
        ctx.beginPath();
        ctx.moveTo(lastX, lastY);
        ctx.lineTo(p.x, p.y);
        ctx.stroke();
        lastX = p.x; lastY = p.y;
        currentStroke.push({ x: p.x, y: p.y });
        hasDrawn = true;
    }

    function endDraw(e) {
        if (!drawing) return;
        e.preventDefault();
        drawing = false;
        if (currentStroke.length > 1) strokes.push(currentStroke);
        currentStroke = [];
    }

    function redrawAll() {
        var rect = canvas.getBoundingClientRect();
        ctx.fillStyle = "#fff";
        ctx.fillRect(0, 0, rect.width, rect.height);
        strokes.forEach(function (stroke) {
            if (stroke.length < 2) return;
            ctx.beginPath();
            ctx.moveTo(stroke[0].x, stroke[0].y);
            for (var i = 1; i < stroke.length; i++) ctx.lineTo(stroke[i].x, stroke[i].y);
            ctx.stroke();
        });
    }

    function clearCanvas() {
        strokes = []; currentStroke = []; hasDrawn = false;
        var rect = canvas.getBoundingClientRect();
        ctx.fillStyle = "#fff";
        ctx.fillRect(0, 0, rect.width, rect.height);
        canvasHint.classList.remove("hidden");
    }

    function undoStroke() {
        if (strokes.length === 0) return;
        strokes.pop();
        redrawAll();
        if (strokes.length === 0) { hasDrawn = false; canvasHint.classList.remove("hidden"); }
    }

    function hasCanvasContent() {
        if (strokes.length > 0) return true;
        try {
            var w = canvas.width, h = canvas.height;
            var data = ctx.getImageData(0, 0, w, h).data;
            for (var i = 0; i < data.length; i += 4) {
                if (data[i] < 250 || data[i+1] < 250 || data[i+2] < 250) return true;
            }
        } catch (e) {}
        return false;
    }

    function getCanvasDataURL() {
        var w = canvas.width, h = canvas.height;
        var minX = w, minY = h, maxX = 0, maxY = 0;
        var data = ctx.getImageData(0, 0, w, h).data;
        var found = false;
        for (var y = 0; y < h; y++) {
            for (var x = 0; x < w; x++) {
                var idx = (y * w + x) * 4;
                if (data[idx] < 250 || data[idx+1] < 250 || data[idx+2] < 250) {
                    found = true;
                    if (x < minX) minX = x;
                    if (x > maxX) maxX = x;
                    if (y < minY) minY = y;
                    if (y > maxY) maxY = y;
                }
            }
        }
        if (!found) return canvas.toDataURL("image/png");
        var pad = 10;
        minX = Math.max(0, minX - pad);
        minY = Math.max(0, minY - pad);
        maxX = Math.min(w, maxX + pad);
        maxY = Math.min(h, maxY + pad);
        var cw = maxX - minX, ch = maxY - minY;
        var tmp = document.createElement("canvas");
        tmp.width = cw; tmp.height = ch;
        var tmpCtx = tmp.getContext("2d");
        tmpCtx.fillStyle = "#fff";
        tmpCtx.fillRect(0, 0, cw, ch);
        tmpCtx.drawImage(canvas, minX, minY, cw, ch, 0, 0, cw, ch);
        return tmp.toDataURL("image/png");
    }

    // ===== Modal =====

    function openModal() {
        modal.style.display = "flex";
        strokes = []; hasDrawn = false;
        setTimeout(function () { resizeCanvas(); }, 50);
    }

    function closeModal() {
        modal.style.display = "none";
        activeSelection = null;
    }

    function placeSignature(dataUrl) {
        if (!activeSelection) return;
        var sel = activeSelection;
        var wrapper = document.createElement("div");
        wrapper.className = "sig-placed";
        wrapper.setAttribute("data-sel-id", sel.id);
        wrapper.style.left = sel.x + "px";
        wrapper.style.top = sel.y + "px";
        wrapper.style.width = sel.w + "px";
        wrapper.style.height = sel.h + "px";

        var img = document.createElement("img");
        img.src = dataUrl;
        img.style.width = "100%";
        img.style.height = "100%";
        img.style.objectFit = "contain";
        wrapper.appendChild(img);

        var badge = document.createElement("div");
        badge.className = "sig-badge";
        badge.textContent = "已签";
        wrapper.appendChild(badge);

        signaturesLayer.appendChild(wrapper);
        sel.el.style.display = "none";
        sel.signed = true;
        sel.dataUrl = dataUrl;
        updateStatus();
    }

    function updateStatus() {
        var count = selections.filter(function (s) { return s.signed; }).length;
        if (count > 0) {
            docStatus.textContent = "已签 " + count + " 处";
            docStatus.classList.add("signed");
            downloadBar.style.display = "block";
            toolbarHint.textContent = "已签 " + count + " 处，可继续框选其他位置，或下载文档";
        } else {
            docStatus.textContent = "待签字";
            docStatus.classList.remove("signed");
            downloadBar.style.display = "none";
        }
    }

    function confirmSign() {
        if (!hasCanvasContent()) { alert("请先手写签名"); return; }
        var dataUrl = getCanvasDataURL();
        placeSignature(dataUrl);
        closeModal();
        strokes = []; hasDrawn = false;
    }

    // ===== Download: Create signed PDF =====

    async function downloadSigned() {
        var signedCount = selections.filter(function (s) { return s.signed; }).length;
        if (signedCount === 0) { alert("请先签名"); return; }

        downloadBtn.disabled = true;
        downloadBtn.textContent = "正在生成 PDF...";

        try {
            if (fileType === "pdf") {
                await downloadSignedPdf();
            } else {
                await downloadSignedDocx();
            }
        } catch (err) {
            alert("生成 PDF 失败: " + err.message);
        }

        downloadBtn.disabled = false;
        downloadBtn.textContent = "下载已签字文档 (PDF)";
    }

    async function downloadSignedPdf() {
        var pdfDoc = await PDFLib.PDFDocument.load(originalFileBytes);
        var signedSigs = selections.filter(function (s) { return s.signed; });

        for (var i = 0; i < signedSigs.length; i++) {
            var sig = signedSigs[i];
            var page = pdfDoc.getPage(sig.page);
            var pageSize = page.getSize();
            var pdfW = pageSize.width;
            var pdfH = pageSize.height;

            // Scale: display pixels -> PDF points
            var scaleX = pdfW / sig.displayW;
            var scaleY = pdfH / sig.displayH;

            // Position in PDF points (pdf-lib Y is from bottom)
            var x = sig.relX * scaleX;
            var y = pdfH - (sig.relY + sig.h) * scaleY;
            var w = sig.w * scaleX;
            var h = sig.h * scaleY;

            // Embed signature image
            var pngBytes = dataUrlToUint8Array(sig.dataUrl);
            var pngImage = await pdfDoc.embedPng(pngBytes);
            page.drawImage(pngImage, { x: x, y: y, width: w, height: h });
        }

        var pdfBytes = await pdfDoc.save();
        triggerDownload(new Blob([pdfBytes], { type: "application/pdf" }), "signed_document.pdf");
    }

    async function downloadSignedDocx() {
        var pdfDoc = await PDFLib.PDFDocument.create();
        var captureScale = 2;

        for (var p = 0; p < pageElements.length; p++) {
            var pe = pageElements[p];

            // Capture page as canvas with html2canvas
            var pageCanvas = await html2canvas(pe.el, {
                scale: captureScale,
                useCORS: true,
                backgroundColor: "#ffffff",
            });

            // Add page image to PDF
            var pngBytes = dataUrlToUint8Array(pageCanvas.toDataURL("image/png"));
            var pngImage = await pdfDoc.embedPng(pngBytes);
            var pdfPage = pdfDoc.addPage([pageCanvas.width, pageCanvas.height]);
            pdfPage.drawImage(pngImage, { x: 0, y: 0, width: pageCanvas.width, height: pageCanvas.height });

            // Draw signatures on this page
            var pageSigs = selections.filter(function (s) { return s.signed && s.page === pe.pageNum; });
            for (var s = 0; s < pageSigs.length; s++) {
                var sig = pageSigs[s];
                var sigPngBytes = dataUrlToUint8Array(sig.dataUrl);
                var sigImage = await pdfDoc.embedPng(sigPngBytes);

                var sx = (pageCanvas.width / pe.el.offsetWidth);
                var sy = (pageCanvas.height / pe.el.offsetHeight);
                var sigX = sig.relX * sx;
                var sigY = pageCanvas.height - (sig.relY + sig.h) * sy;
                var sigW = sig.w * sx;
                var sigH = sig.h * sy;
                pdfPage.drawImage(sigImage, { x: sigX, y: sigY, width: sigW, height: sigH });
            }
        }

        var pdfBytes = await pdfDoc.save();
        triggerDownload(new Blob([pdfBytes], { type: "application/pdf" }), "signed_document.pdf");
    }

    function dataUrlToUint8Array(dataUrl) {
        var base64 = dataUrl.split(",")[1];
        var binary = atob(base64);
        var bytes = new Uint8Array(binary.length);
        for (var i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
        return bytes;
    }

    function triggerDownload(blob, filename) {
        var url = URL.createObjectURL(blob);
        var a = document.createElement("a");
        a.href = url;
        a.download = filename;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        setTimeout(function () { URL.revokeObjectURL(url); }, 1000);
    }

    // ===== Event Bindings =====

    // Selection events
    docPaper.addEventListener("mousedown", startSelection);
    document.addEventListener("mousemove", updateSelection);
    document.addEventListener("mouseup", endSelection);
    docPaper.addEventListener("touchstart", startSelection, { passive: false });
    document.addEventListener("touchmove", updateSelection, { passive: false });
    document.addEventListener("touchend", endSelection, { passive: false });

    // Canvas events
    canvas.addEventListener("mousedown", function (e) { e.stopPropagation(); startDraw(e); });
    canvas.addEventListener("mousemove", draw);
    canvas.addEventListener("mouseup", endDraw);
    canvas.addEventListener("mouseleave", endDraw);
    canvas.addEventListener("touchstart", function (e) { e.stopPropagation(); startDraw(e); }, { passive: false });
    canvas.addEventListener("touchmove", draw, { passive: false });
    canvas.addEventListener("touchend", endDraw, { passive: false });
    canvas.addEventListener("touchcancel", endDraw, { passive: false });

    // Modal controls
    closePadBtn.addEventListener("click", closeModal);
    clearBtn.addEventListener("click", clearCanvas);
    undoBtn.addEventListener("click", undoStroke);
    confirmSignBtn.addEventListener("click", confirmSign);
    modal.addEventListener("click", function (e) { if (e.target === modal) closeModal(); });

    // Download
    downloadBtn.addEventListener("click", downloadSigned);

    // Resize handler
    var resizeTimer;
    window.addEventListener("resize", function () {
        clearTimeout(resizeTimer);
        resizeTimer = setTimeout(function () {
            if (modal.style.display !== "none") resizeCanvas();
            if (fileType === "docx" && docReady) collectDocxPages();
        }, 200);
    });
})();
