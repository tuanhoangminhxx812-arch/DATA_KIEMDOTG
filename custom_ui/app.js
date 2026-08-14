/* 
========================================================================
   HỆ THỐNG KIỂM DÒ TÀI KHOẢN TRUNG GIAN - STREAMLIT COMPONENT LOGIC
   Drilldown 3 cấp tương tác giống hệt Excel
========================================================================
*/

// Trạng thái ứng dụng ở phía Client
const state = {
    fileUploaded: false,
    hasAnalyzed: false,
    data: null, // Dữ liệu phân tích nhận được từ Python
    selectedAccount: '', // Mã tài khoản đang chọn
    selectedCreator: '', // Người hạch toán đang chọn (rỗng = tất cả)
    currentPage: 1,
    rowsPerPage: 15,
    filteredChitietRows: []
};

// Khởi động ứng dụng sau khi DOM sẵn sàng
document.addEventListener('DOMContentLoaded', () => {
    initApp();
});

// Hàm gửi tin nhắn về Python Streamlit
function sendToStreamlit(action, payload = {}) {
    window.parent.postMessage({
        isStreamlitMessage: true,
        type: "streamlit:setComponentValue",
        value: { action, ...payload }
    }, "*");
}

function initApp() {
    // Đăng ký DOM Elements
    const elements = {
        connectionStatus: document.getElementById('connection-status'),
        statusText: document.querySelector('.status-text'),
        currentFilename: document.getElementById('current-filename'),
        
        btnSelectExcel: document.getElementById('btn-add-file'),
        excelFileInput: document.getElementById('excel-file-input'),
        btnAnalyze: document.getElementById('btn-analyze'),
        
        emptyState: document.getElementById('empty-state'),
        loadingState: document.getElementById('loading-state'),
        dashboardPanel: document.getElementById('analysis-dashboard'),
        
        kpiAccUnbalanced: document.getElementById('kpi-acc-unbalanced'),
        kpiTotalDebit: document.getElementById('kpi-total-debit'),
        kpiTotalCredit: document.getElementById('kpi-total-credit'),
        kpiNetDiff: document.getElementById('kpi-net-diff'),
        kpiDiffCard: document.getElementById('kpi-diff-card'),
        
        filterAccountSelect: document.getElementById('filter-account-select'),
        
        tbodyTableA: document.getElementById('tbody-table-a'),
        totalADebit: document.getElementById('total-a-debit'),
        totalACredit: document.getElementById('total-a-credit'),
        totalADiff: document.getElementById('total-a-diff'),
        totalAStatus: document.getElementById('total-a-status'),
        
        tableBSection: document.getElementById('table-b-section'),
        selectedAccountBTitle: document.getElementById('selected-account-b-title'),
        tbodyTableB: document.getElementById('tbody-table-b'),
        totalBDebit: document.getElementById('total-b-debit'),
        totalBCredit: document.getElementById('total-b-credit'),
        totalBDiff: document.getElementById('total-b-diff'),
        totalBStatus: document.getElementById('total-b-status'),
        
        tableCSection: document.getElementById('table-c-section'),
        selectedCreatorCTitle: document.getElementById('selected-creator-c-title'),
        btnExportFiltered: document.getElementById('btn-export-filtered'),
        tbodyChitiet: document.getElementById('tbody-chitiet'),
        searchInput: document.getElementById('search-input'),
        pageStartRow: document.getElementById('page-start-row'),
        pageEndRow: document.getElementById('page-end-row'),
        totalRowsCount: document.getElementById('total-rows-count'),
        currentPageNum: document.getElementById('current-page-num'),
        totalPagesNum: document.getElementById('total-pages-num'),
        btnPrevPage: document.getElementById('btn-prev-page'),
        btnNextPage: document.getElementById('btn-next-page')
    };

    // Lưu elements vào đối tượng toàn cục để dễ truy cập
    window.appElements = elements;

    // Đăng ký sự kiện
    registerEvents(elements);

    // Đăng ký lắng nghe dữ liệu từ Python gửi xuống
    window.addEventListener("message", (event) => {
        if (event.data.type === "streamlit:render") {
            const args = event.data.args;
            if (args && args.state) {
                handleStreamlitState(args.state);
            }
        }
    });

    // Thông báo cho Streamlit parent là component đã sẵn sàng hoạt động (để ẩn thanh cảnh báo màu vàng)
    window.parent.postMessage({
        isStreamlitMessage: true,
        type: "streamlit:componentReady",
        apiVersion: 1
    }, "*");

    // Thiết lập độ cao khung hình ban đầu cho iframe
    window.parent.postMessage({
        isStreamlitMessage: true,
        type: "streamlit:setFrameHeight",
        height: 1000
    }, "*");

    // Gửi tín hiệu init để Python biết component đã tải xong
    sendToStreamlit("init");
}

// Xử lý luồng dữ liệu và trạng thái từ Streamlit Cloud gửi xuống
function handleStreamlitState(st) {
    const el = window.appElements;
    
    // 1. Quản lý trạng thái File Đã Tải Lên
    if (st.uploaded_filename) {
        state.fileUploaded = true;
        el.currentFilename.innerText = `File đang tải: ${st.uploaded_filename}`;
        el.btnAnalyze.disabled = false;
        el.connectionStatus.className = "status-indicator online";
        el.statusText.innerText = "Đã nhận file Excel";
    } else {
        state.fileUploaded = false;
        state.hasAnalyzed = false;
        el.currentFilename.innerText = "Vui lòng tải lên file Excel phát sinh để bắt đầu đối chiếu";
        el.btnAnalyze.disabled = true;
        el.connectionStatus.className = "status-indicator offline";
        el.statusText.innerText = "Đang chờ tải file...";
        
        el.emptyState.style.display = 'flex';
        el.dashboardPanel.style.display = 'none';
        resetKPIs();
        resetFilters();
    }
    
    // 2. Quản lý kết quả phân tích đối chiếu chéo
    if (st.has_analyzed && st.analysis_data) {
        state.hasAnalyzed = true;
        state.data = st.analysis_data;
        
        // Tắt loading spinner
        el.loadingState.style.display = 'none';
        el.btnAnalyze.disabled = false;
        el.btnSelectExcel.disabled = false;
        const spinnerIcon = el.btnAnalyze.querySelector('.icon-spin-target');
        if (spinnerIcon) spinnerIcon.classList.remove('spin');
        
        // Vẽ lại giao diện
        renderAll();
    }
    
    // 3. Kích hoạt tải xuống Excel in-memory serverless khi Python gửi Base64 về
    if (st.download_trigger) {
        triggerDownload(st.download_trigger.filename, st.download_trigger.data);
        // Báo cho Python là tải xong để reset trigger trong session_state
        sendToStreamlit("download_done");
    }
}

// Tải file xuống máy khách hàng trực tiếp từ byte array Base64
function triggerDownload(filename, base64Data) {
    try {
        const byteCharacters = atob(base64Data);
        const byteNumbers = new Array(byteCharacters.length);
        for (let i = 0; i < byteCharacters.length; i++) {
            byteNumbers[i] = byteCharacters.charCodeAt(i);
        }
        const byteArray = new Uint8Array(byteNumbers);
        const blob = new Blob([byteArray], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
        
        const link = document.createElement('a');
        link.href = window.URL.createObjectURL(blob);
        link.download = filename;
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        
        showToast("Tải file lệch thành công!", "success");
    } catch (err) {
        console.error("Lỗi khi tải file xuống:", err);
        showToast("Lỗi tải xuống báo cáo!", "error");
    }
}

// Đăng ký các sự kiện tương tác
function registerEvents(el) {
    // A. Chọn file Excel qua Drag & Drop ẩn
    el.btnSelectExcel.addEventListener('click', () => {
        el.excelFileInput.click();
    });

    el.excelFileInput.addEventListener('change', (e) => {
        if (e.target.files.length === 0) return;
        
        const file = e.target.files[0];
        const reader = new FileReader();
        
        el.statusText.innerText = "Đang đọc file...";
        el.btnSelectExcel.disabled = true;
        
        reader.onload = function(evt) {
            const base64Data = evt.target.result.split(',')[1];
            
            // Gửi base64 dữ liệu file lên Python
            sendToStreamlit("upload", {
                filename: file.name,
                data: base64Data
            });
            
            el.btnSelectExcel.disabled = false;
            el.excelFileInput.value = '';
            showToast(`Đã nhận file: ${file.name}`, "success");
        };
        
        reader.onerror = function() {
            showToast("Lỗi đọc file Excel!", "error");
            el.btnSelectExcel.disabled = false;
            el.excelFileInput.value = '';
        };
        
        reader.readAsDataURL(file);
    });

    // B. Chạy Phân Tích Đối Chiếu
    el.btnAnalyze.addEventListener('click', () => {
        if (!state.fileUploaded) return;
        
        el.emptyState.style.display = 'none';
        el.dashboardPanel.style.display = 'none';
        el.loadingState.style.display = 'flex';
        el.btnAnalyze.disabled = true;
        el.btnSelectExcel.disabled = true;
        
        const spinnerIcon = el.btnAnalyze.querySelector('.icon-spin-target');
        if (spinnerIcon) spinnerIcon.classList.add('spin');
        
        // Gửi yêu cầu phân tích về Python
        sendToStreamlit("analyze");
    });

    // C. Thay đổi Dropdown chọn Tài khoản (Ô C5)
    el.filterAccountSelect.addEventListener('change', function() {
        const accCode = this.value;
        if (!accCode) return;
        
        selectAccount(accCode);
    });

    // D. Ô tìm kiếm bảng Chi Tiết
    el.searchInput.addEventListener('input', () => {
        state.currentPage = 1;
        filterAndPaginateChitiet();
    });

    // E. Phân trang
    el.btnPrevPage.addEventListener('click', () => {
        if (state.currentPage > 1) {
            state.currentPage--;
            renderChitietTable();
        }
    });

    el.btnNextPage.addEventListener('click', () => {
        const totalPages = Math.ceil(state.filteredChitietRows.length / state.rowsPerPage);
        if (state.currentPage < totalPages) {
            state.currentPage++;
            renderChitietTable();
        }
    });

    // F. Xuất File Lệch Cấp độ 3 (Bảng C)
    el.btnExportFiltered.addEventListener('click', () => {
        if (!state.selectedAccount) {
            showToast("Vui lòng chọn tài khoản để xuất file lệch!", "error");
            return;
        }
        
        showToast("Đang chuẩn bị xuất file Excel lệch...", "success");
        // Gọi Python tạo excel in-memory và gửi Base64 về
        sendToStreamlit("export", {
            account: state.selectedAccount,
            creator: state.selectedCreator
        });
    });

    // G. Click trên ô Tổng cộng chênh lệch Bảng B
    el.totalBDiff.addEventListener('click', () => {
        if (!state.selectedAccount) return;
        selectCreator(''); // Click vào tổng cộng thì xem tất cả creator
    });
}

// Reset các KPI cards
function resetKPIs() {
    const el = window.appElements;
    el.kpiAccUnbalanced.innerText = "-";
    el.kpiTotalDebit.innerText = "-";
    el.kpiTotalCredit.innerText = "-";
    el.kpiNetDiff.innerText = "-";
    el.kpiDiffCard.className = "kpi-card bg-glass";
    el.kpiDiffCard.style.borderLeft = "";
}

// Reset các bộ lọc và các phần bảng con
function resetFilters() {
    const el = window.appElements;
    el.filterAccountSelect.innerHTML = '<option value="" disabled selected>-- Chọn tài khoản trung gian --</option>';
    el.tableBSection.style.display = 'none';
    el.tableCSection.style.display = 'none';
}

// Định dạng tiền tệ VND
function formatMoney(val) {
    if (val === undefined || val === null) return "0";
    const rounded = Math.round(val);
    const absVal = Math.abs(rounded);
    const formatted = absVal.toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",");
    return rounded < 0 ? `-${formatted}` : formatted;
}

// Làm tròn số thập phân
function round(value, decimals) {
    return Number(Math.round(value + 'e' + decimals) + 'e-' + decimals);
}

// Tạo thông báo Toast
function showToast(msg, type = 'success') {
    const container = document.getElementById('toast-container');
    const toast = document.createElement('div');
    toast.className = `toast toast-${type}`;
    
    let svgIcon = '';
    if (type === 'success') {
        svgIcon = `<svg class="toast-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><polyline points="20 6 9 17 4 12"/></svg>`;
    } else {
        svgIcon = `<svg class="toast-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="12"/><line x1="12" y1="16" x2="12.01" y2="16"/></svg>`;
    }
    
    toast.innerHTML = `
        ${svgIcon}
        <span class="toast-message">${msg}</span>
    `;
    
    container.appendChild(toast);
    
    setTimeout(() => {
        toast.style.animation = 'fadeOut 0.3s forwards';
        setTimeout(() => {
            toast.remove();
        }, 300);
    }, 3500);
}

// Render giao diện sau khi phân tích
function renderAll() {
    const el = window.appElements;
    el.emptyState.style.display = 'none';
    el.loadingState.style.display = 'none';
    el.dashboardPanel.style.display = 'block';
    
    renderKPIs();
    populateAccountDropdown();
    renderTableA();
    
    // Nếu có tài khoản đang chọn sẵn, khôi phục lại
    if (state.selectedAccount) {
        selectAccount(state.selectedAccount);
    } else {
        el.tableBSection.style.display = 'none';
        el.tableCSection.style.display = 'none';
    }
}

// Render các khối KPI
function renderKPIs() {
    if (!state.data) return;
    const kpis = state.data.kpi_cards;
    const el = window.appElements;
    
    el.kpiAccUnbalanced.innerText = `${kpis.total_unbalanced_accs} TK lệch`;
    el.kpiTotalDebit.innerText = `${formatMoney(kpis.total_unpaired_debit)} đ`;
    el.kpiTotalCredit.innerText = `${formatMoney(kpis.total_unpaired_credit)} đ`;
    el.kpiNetDiff.innerText = `${formatMoney(kpis.total_discrepancy)} đ`;
    
    if (Math.abs(Math.round(kpis.total_discrepancy)) === 0) {
        el.kpiDiffCard.className = "kpi-card bg-teal";
    } else {
        el.kpiDiffCard.className = "kpi-card bg-glass";
        el.kpiDiffCard.style.borderLeft = "5px solid var(--color-error-fill)";
    }
}

// Điền tài khoản vào Dropdown chọn nhanh
function populateAccountDropdown() {
    const select = window.appElements.filterAccountSelect;
    select.innerHTML = '<option value="" disabled selected>-- Chọn tài khoản trung gian --</option>';
    
    if (!state.data) return;
    
    const accounts = state.data.doi_chieu_bang_a.map(item => item.account);
    accounts.sort().forEach(acc => {
        const opt = document.createElement('option');
        opt.value = acc;
        opt.innerText = acc;
        select.appendChild(opt);
    });
}

// --- CẤP ĐỘ 1: RENDER BẢNG A ---
function renderTableA() {
    const el = window.appElements;
    const tbody = el.tbodyTableA;
    tbody.innerHTML = '';
    
    if (!state.data) return;
    
    const items = state.data.doi_chieu_bang_a;
    
    // Sort theo mã tài khoản tăng dần
    const sortedItems = [...items].sort((a, b) => a.account.localeCompare(b.account));
    
    let sumDebit = 0;
    let sumCredit = 0;
    let sumDiff = 0;
    
    sortedItems.forEach(item => {
        sumDebit += item.debit;
        sumCredit += item.credit;
        sumDiff += item.diff;
        
        const tr = document.createElement('tr');
        tr.dataset.account = item.account;
        if (state.selectedAccount === item.account) {
            tr.className = 'selected-row';
        }
        
        const isLech = item.status === 'Lệch số';
        const badgeClass = isLech ? 'badge-error' : 'badge-success';
        const badgeText = isLech ? 'Lệch số' : 'Khớp';
        
        tr.innerHTML = `
            <td style="font-weight: 700; color: var(--navy-medium);">${item.account}</td>
            <td class="text-right">${formatMoney(item.debit)}</td>
            <td class="text-right">${formatMoney(item.credit)}</td>
            <td class="text-right font-semibold ${isLech ? 'text-error' : 'text-success'}">${formatMoney(item.diff)}</td>
            <td class="text-center">
                <span class="badge ${badgeClass}">${badgeText}</span>
            </td>
        `;
        
        // Thêm sự kiện click dòng
        tr.addEventListener('click', () => {
            selectAccount(item.account);
        });
        
        tbody.appendChild(tr);
    });
    
    // Cập nhật dòng Tổng cộng Bảng A
    el.totalADebit.innerText = formatMoney(sumDebit);
    el.totalACredit.innerText = formatMoney(sumCredit);
    el.totalADiff.innerText = formatMoney(sumDiff);
    
    const hasOverallDiff = Math.abs(round(sumDiff, 2)) !== 0;
    el.totalADiff.className = `text-right font-bold ${hasOverallDiff ? 'text-error' : 'text-success'}`;
    el.totalAStatus.innerHTML = hasOverallDiff 
        ? '<span class="badge badge-error">Lệch tổng</span>' 
        : '<span class="badge badge-success">Khớp tổng</span>';
}

// Logic chọn tài khoản
function selectAccount(account) {
    state.selectedAccount = account;
    state.selectedCreator = ''; // Reset người tạo khi đổi tài khoản
    
    // Sync dropdown
    const select = window.appElements.filterAccountSelect;
    select.value = account;
    
    // Cập nhật class selected cho Bảng A
    const rows = window.appElements.tbodyTableA.querySelectorAll('tr');
    rows.forEach(r => {
        if (r.dataset.account === account) {
            r.classList.add('selected-row');
        } else {
            r.classList.remove('selected-row');
        }
    });
    
    // Render Bảng B
    renderTableB(account);
    
    // Ẩn Bảng C cho đến khi được chọn
    window.appElements.tableCSection.style.display = 'none';
}

// --- CẤP ĐỘ 2: RENDER BẢNG B ---
function renderTableB(account) {
    const el = window.appElements;
    const tbody = el.tbodyTableB;
    tbody.innerHTML = '';
    
    if (!state.data || !state.data.doi_chieu_bang_b) return;
    
    const creators = state.data.doi_chieu_bang_b[account] || [];
    el.selectedAccountBTitle.innerText = account;
    
    let sumDebit = 0;
    let sumCredit = 0;
    let sumDiff = 0;
    
    creators.forEach(item => {
        sumDebit += item.debit;
        sumCredit += item.credit;
        sumDiff += item.diff;
        
        const tr = document.createElement('tr');
        tr.dataset.creator = item.creator;
        if (state.selectedCreator === item.creator) {
            tr.className = 'selected-row';
        }
        
        const isLech = item.status === 'Lệch số';
        const hasData = item.debit > 0 || item.credit > 0;
        
        let statusBadge = '';
        if (isLech) {
            statusBadge = '<span class="badge badge-error">Lệch số</span>';
        } else if (hasData) {
            statusBadge = '<span class="badge badge-success">Khớp</span>';
        }
        
        // Chênh lệch được hiển thị dạng Link Clickable nếu lệch thực tế khác 0
        let diffTdContent = '';
        if (isLech) {
            diffTdContent = `<span class="clickable-diff" title="Click để xem chi tiết giao dịch lệch">${formatMoney(item.diff)}</span>`;
        } else {
            diffTdContent = formatMoney(item.diff);
        }
        
        tr.innerHTML = `
            <td style="font-weight: 600; color: var(--navy-medium);">${item.creator}</td>
            <td class="text-right">${formatMoney(item.debit)}</td>
            <td class="text-right">${formatMoney(item.credit)}</td>
            <td class="text-right font-semibold ${isLech ? 'text-error' : 'text-success'}">${diffTdContent}</td>
            <td class="text-center">${statusBadge}</td>
        `;
        
        // Gắn sự kiện click
        const diffSpan = tr.querySelector('.clickable-diff');
        if (diffSpan) {
            diffSpan.addEventListener('click', (e) => {
                e.stopPropagation(); // Tránh kích hoạt click dòng nếu cần phân tách sâu
                selectCreator(item.creator);
            });
        }
        
        // Click cả dòng cũng hiển thị chi tiết (để thân thiện người dùng)
        tr.addEventListener('click', () => {
            selectCreator(item.creator);
        });
        
        tbody.appendChild(tr);
    });
    
    // Cập nhật dòng Tổng cộng Bảng B
    el.totalBDebit.innerText = formatMoney(sumDebit);
    el.totalBCredit.innerText = formatMoney(sumCredit);
    
    // Thường ô Tổng cộng chênh lệch Bảng B sẽ bằng chênh lệch Bảng A của tài khoản đó
    el.totalBDiff.innerText = formatMoney(sumDiff);
    
    const hasOverallDiff = Math.abs(round(sumDiff, 2)) !== 0;
    el.totalBDiff.className = `text-right font-bold clickable-diff ${hasOverallDiff ? 'text-error' : 'text-success'}`;
    el.totalBStatus.innerHTML = hasOverallDiff
        ? '<span class="badge badge-error">Lệch số</span>'
        : (sumDebit > 0 || sumCredit > 0 ? '<span class="badge badge-success">Khớp</span>' : '');
        
    // Hiển thị Card Bảng B
    el.tableBSection.style.display = 'block';
    
    // Cuộn nhẹ xuống Bảng B
    el.tableBSection.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
}

// Logic chọn người tạo để xem bảng C
function selectCreator(creator) {
    state.selectedCreator = creator;
    state.currentPage = 1;
    
    // Cập nhật highlight trong Bảng B
    const rows = window.appElements.tbodyTableB.querySelectorAll('tr');
    rows.forEach(r => {
        if (r.dataset.creator === creator) {
            r.classList.add('selected-row');
        } else {
            r.classList.remove('selected-row');
        }
    });
    
    // Render Bảng C
    renderTableC();
}

// --- CẤP ĐỘ 3: RENDER BẢNG C (CHI TIẾT) ---
function renderTableC() {
    const el = window.appElements;
    const accCode = state.selectedAccount;
    const creator = state.selectedCreator;
    
    if (!state.data) return;
    
    // Tiêu đề Bảng C
    if (!creator) {
        el.selectedCreatorTitle = "TẤT CẢ NGƯỜI TẠO";
        el.selectedCreatorCTitle.innerHTML = `BẢNG CHI TIẾT CÁC GIAO DỊCH CHƯA ĐỐI ỨNG (TẤT CẢ NGƯỜI TẠO) - TK: <span class="text-accent">${accCode}</span>`;
    } else {
        el.selectedCreatorCTitle.innerHTML = `BẢNG CHI TIẾT CÁC GIAO DỊCH CHƯA ĐỐI ỨNG CỦA: <span class="text-accent">${creator}</span> - TK: <span class="text-accent">${accCode}</span>`;
    }
    
    // Thu thập các dòng chi tiết cho tài khoản
    const allTxs = state.data.chi_tiet_giao_dich[accCode] || [];
    
    // Lưu vào state để lọc nhanh
    state.filteredChitietRows = creator 
        ? allTxs.filter(tx => tx.creator === creator)
        : allTxs;
        
    // Reset ô tìm kiếm về rỗng khi chọn Creator mới
    el.searchInput.value = '';
    
    // Hiển thị Card Bảng C
    el.tableCSection.style.display = 'block';
    
    filterAndPaginateChitiet();
    
    // Cuộn mượt xuống Bảng C
    el.tableCSection.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
}

// Lọc nhanh đa cột trong Bảng C và Phân trang
function filterAndPaginateChitiet() {
    const el = window.appElements;
    const accCode = state.selectedAccount;
    const creator = state.selectedCreator;
    const query = el.searchInput.value.toLowerCase().trim();
    
    // Lọc lại dữ liệu gốc tương ứng
    const allTxs = state.data.chi_tiet_giao_dich[accCode] || [];
    let rows = creator 
        ? allTxs.filter(tx => tx.creator === creator)
        : allTxs;
        
    if (query !== '') {
        rows = rows.filter(r => {
            return (
                r.creator.toLowerCase().includes(query) ||
                (r.updater && r.updater.toLowerCase().includes(query)) ||
                r.trans_num.toLowerCase().includes(query) ||
                r.desc.toLowerCase().includes(query) ||
                r.date.includes(query)
            );
        });
    }
    
    // Sắp xếp theo dòng gốc trong Excel
    rows.sort((a, b) => a.row_num - b.row_num);
    
    state.filteredChitietRows = rows;
    renderChitietTable();
}

// Kết xuất dữ liệu phân trang 15 dòng cho Bảng C
function renderChitietTable() {
    const el = window.appElements;
    const tbody = el.tbodyChitiet;
    tbody.innerHTML = '';
    
    const rows = state.filteredChitietRows;
    const totalRows = rows.length;
    el.totalRowsCount.innerText = totalRows;
    
    if (totalRows === 0) {
        tbody.innerHTML = `
            <tr>
                <td colspan="9" class="text-center text-muted py-5" style="font-style: italic;">
                    Không tìm thấy giao dịch lệch nào phù hợp với điều kiện hạch toán.
                </td>
            </tr>
        `;
        
        el.pageStartRow.innerText = "0";
        el.pageEndRow.innerText = "0";
        el.currentPageNum.innerText = "1";
        el.totalPagesNum.innerText = "1";
        el.btnPrevPage.disabled = true;
        el.btnNextPage.disabled = true;
        return;
    }
    
    const totalPages = Math.ceil(totalRows / state.rowsPerPage);
    if (state.currentPage > totalPages) state.currentPage = totalPages;
    if (state.currentPage < 1) state.currentPage = 1;
    
    const startIdx = (state.currentPage - 1) * state.rowsPerPage;
    const endIdx = Math.min(startIdx + state.rowsPerPage, totalRows);
    
    const pageRows = rows.slice(startIdx, endIdx);
    
    pageRows.forEach(item => {
        const tr = document.createElement('tr');
        
        tr.innerHTML = `
            <td class="text-center text-muted" style="font-weight: 600;">${item.row_num}</td>
            <td style="font-weight: 600; color: var(--navy-light);">${item.account}</td>
            <td style="font-family: monospace; font-size: 12px; font-weight: 600;">${item.trans_num}</td>
            <td style="white-space: nowrap;">${item.date}</td>
            <td class="text-right font-semibold" style="color: var(--color-error-fill);">${item.deb > 0 ? formatMoney(item.deb) : '-'}</td>
            <td class="text-right font-semibold" style="color: var(--color-success-fill);">${item.cred > 0 ? formatMoney(item.cred) : '-'}</td>
            <td style="font-weight: 600; color: var(--navy-medium);">${item.creator}</td>
            <td class="text-muted">${item.updater || ''}</td>
            <td class="text-muted" style="max-width: 380px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;" title="${item.desc}">${item.desc}</td>
        `;
        
        tbody.appendChild(tr);
    });
    
    // Cập nhật thông tin phân trang trên UI
    el.pageStartRow.innerText = startIdx + 1;
    el.pageEndRow.innerText = endIdx;
    el.currentPageNum.innerText = state.currentPage;
    el.totalPagesNum.innerText = totalPages;
    
    // Bật tắt nút phân trang
    el.btnPrevPage.disabled = state.currentPage === 1;
    el.btnNextPage.disabled = state.currentPage === totalPages;
}
