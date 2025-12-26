// 全域變數
let currentDate = new Date();
let openEmployees = [];
let closeEmployees = [];
let customHolidays = [];

// 假日設定 (可根據需要調整)
const holidays = [
    '2024-01-01', // 元旦
    '2024-02-10', '2024-02-11', '2024-02-12', // 春節
    '2024-04-04', '2024-04-05', // 清明節
    '2024-05-01', // 勞動節
    '2024-06-10', // 端午節
    '2024-09-17', // 中秋節
    '2024-10-10', // 國慶日
    '2025-01-01', // 元旦
    '2025-01-29', '2025-01-30', '2025-01-31', // 春節
    '2025-04-04', '2025-04-05', // 清明節
    '2025-05-01', // 勞動節
    '2025-05-31', // 端午節
    '2025-10-06', // 中秋節
    '2025-10-10'  // 國慶日
];

// DOM 元素
const monthYearElement = document.getElementById('monthYear');
const calendarBody = document.getElementById('calendarBody');
const prevMonthBtn = document.getElementById('prevMonth');
const nextMonthBtn = document.getElementById('nextMonth');
const openEmployeeNameInput = document.getElementById('openEmployeeName');
const addOpenEmployeeBtn = document.getElementById('addOpenEmployee');
const openEmployeeListContainer = document.getElementById('openEmployeeListContainer');
const closeEmployeeNameInput = document.getElementById('closeEmployeeName');
const addCloseEmployeeBtn = document.getElementById('addCloseEmployee');
const closeEmployeeListContainer = document.getElementById('closeEmployeeListContainer');
const holidayDateInput = document.getElementById('holidayDate');
const addHolidayBtn = document.getElementById('addHoliday');
const holidayListContainer = document.getElementById('holidayListContainer');
const downloadExcelBtn = document.getElementById('downloadExcel');
const darkModeToggleBtn = document.getElementById('darkModeToggle');

// 初始化
document.addEventListener('DOMContentLoaded', function() {
    loadEmployees();
    loadCustomHolidays();
    loadTheme();
    displayCalendar();
    updateEmployeeList();
    updateHolidayList();
    
    // 事件監聽器
    prevMonthBtn.addEventListener('click', previousMonth);
    nextMonthBtn.addEventListener('click', nextMonth);
    addOpenEmployeeBtn.addEventListener('click', addOpenEmployee);
    addCloseEmployeeBtn.addEventListener('click', addCloseEmployee);
    addHolidayBtn.addEventListener('click', addCustomHoliday);
    downloadExcelBtn.addEventListener('click', downloadExcel);
    darkModeToggleBtn.addEventListener('click', toggleDarkMode);
});

// 顯示行事曆
function displayCalendar() {
    const year = currentDate.getFullYear();
    const month = currentDate.getMonth();
    
    // 更新月份年份顯示
    monthYearElement.textContent = `${year}年 ${month + 1}月`;
    
    // 清空行事曆
    calendarBody.innerHTML = '';
    
    // 獲得本月第一天和最後一天
    const firstDay = new Date(year, month, 1);
    const lastDay = new Date(year, month + 1, 0);
    const startDate = new Date(firstDay);
    startDate.setDate(startDate.getDate() - firstDay.getDay());
    
    // 生成6週的日曆
    for (let week = 0; week < 6; week++) {
        const row = document.createElement('tr');
        
        for (let day = 0; day < 7; day++) {
            const cell = document.createElement('td');
            const cellDate = new Date(startDate);
            cellDate.setDate(startDate.getDate() + (week * 7) + day);
            
            // 日期數字
            const dateNumber = document.createElement('div');
            dateNumber.className = 'date-number';
            dateNumber.textContent = cellDate.getDate();
            
            // 檢查是否為假日
            const dateString = formatDate(cellDate);
            const isNationalHoliday = holidays.includes(dateString);
            const isCustomHoliday = customHolidays.includes(dateString);
            const isWeekend = cellDate.getDay() === 0 || cellDate.getDay() === 6;
            
            // 設定樣式
            if (cellDate.getMonth() !== month) {
                cell.className = 'other-month';
            } else if (isNationalHoliday || isCustomHoliday) {
                cell.className = 'holiday';
            } else if (isWeekend) {
                cell.className = 'weekend';
            }
            
            cell.appendChild(dateNumber);
            
            // 只在當前月份顯示員工排班
            if (cellDate.getMonth() === month) {
                // 添加點擊編輯功能
                cell.style.cursor = 'pointer';
                cell.onclick = () => showDayEditModal(cellDate);
                
                // 顯示該日期的開門員工名單
                const dayOpenEmployees = getOpenEmployeesForDate(cellDate);
                dayOpenEmployees.forEach(emp => {
                    const empElement = document.createElement('div');
                    empElement.className = 'employee-name';
                    empElement.textContent = `開:${emp}`;
                    cell.appendChild(empElement);
                });
                
                // 顯示該日期的關門員工名單
                const dayCloseEmployees = getCloseEmployeesForDate(cellDate);
                dayCloseEmployees.forEach(emp => {
                    const empElement = document.createElement('div');
                    empElement.className = 'close-employee-name';
                    empElement.textContent = `關:${emp}`;
                    cell.appendChild(empElement);
                });
            }
            
            row.appendChild(cell);
        }
        
        calendarBody.appendChild(row);
    }
}

// 上個月
function previousMonth() {
    currentDate.setMonth(currentDate.getMonth() - 1);
    displayCalendar();
}

// 下個月
function nextMonth() {
    currentDate.setMonth(currentDate.getMonth() + 1);
    displayCalendar();
}

// 新增開門員工
function addOpenEmployee() {
    const employeeNames = openEmployeeNameInput.value.trim();
    
    if (!employeeNames) {
        alert('請填入員工姓名');
        return;
    }
    
    // 按行分割，過濾空行
    const names = employeeNames.split('\n').map(name => name.trim()).filter(name => name.length > 0);
    
    if (names.length === 0) {
        alert('請填入員工姓名');
        return;
    }
    
    // 批次新增員工
    names.forEach(name => {
        const employee = {
            id: Date.now() + Math.random(), // 避免ID重複
            name: name,
            order: openEmployees.length
        };
        openEmployees.push(employee);
    });
    
    saveEmployees();
    updateEmployeeList();
    displayCalendar();
    
    // 清空輸入欄位
    openEmployeeNameInput.value = '';
}

// 新增關門員工
function addCloseEmployee() {
    const employeeNames = closeEmployeeNameInput.value.trim();
    
    if (!employeeNames) {
        alert('請填入員工姓名');
        return;
    }
    
    // 按行分割，過濾空行
    const names = employeeNames.split('\n').map(name => name.trim()).filter(name => name.length > 0);
    
    if (names.length === 0) {
        alert('請填入員工姓名');
        return;
    }
    
    // 批次新增員工
    names.forEach(name => {
        const employee = {
            id: Date.now() + Math.random(), // 避免ID重複
            name: name,
            order: closeEmployees.length
        };
        closeEmployees.push(employee);
    });
    
    saveEmployees();
    updateEmployeeList();
    displayCalendar();
    
    // 清空輸入欄位
    closeEmployeeNameInput.value = '';
}

// 刪除開門員工
function deleteOpenEmployee(id) {
    openEmployees = openEmployees.filter(emp => emp.id !== id);
    saveEmployees();
    updateEmployeeList();
    displayCalendar();
}

// 刪除關門員工
function deleteCloseEmployee(id) {
    closeEmployees = closeEmployees.filter(emp => emp.id !== id);
    saveEmployees();
    updateEmployeeList();
    displayCalendar();
}

// 更新員工名單顯示
function updateEmployeeList() {
    // 更新開門員工名單
    openEmployeeListContainer.innerHTML = '';
    openEmployees.forEach((emp, index) => {
        const empItem = document.createElement('div');
        empItem.className = 'employee-item';
        
        empItem.innerHTML = `
            <span>${index + 1}. ${emp.name}</span>
            <button class="delete-btn" onclick="deleteOpenEmployee(${emp.id})">刪除</button>
        `;
        
        openEmployeeListContainer.appendChild(empItem);
    });
    
    // 更新關門員工名單
    closeEmployeeListContainer.innerHTML = '';
    closeEmployees.forEach((emp, index) => {
        const empItem = document.createElement('div');
        empItem.className = 'employee-item';
        
        empItem.innerHTML = `
            <span>${index + 1}. ${emp.name}</span>
            <button class="delete-btn" onclick="deleteCloseEmployee(${emp.id})">刪除</button>
        `;
        
        closeEmployeeListContainer.appendChild(empItem);
    });
}

// 獲得指定日期的開門員工名單
function getOpenEmployeesForDate(date) {
    if (openEmployees.length === 0) return [];
    
    // 檢查當天是否為假日或週末，如果是則不排班
    if (isHolidayOrWeekend(date)) {
        return [];
    }
    
    // 檢查是否有自訂排班
    const dateString = formatDate(date);
    const customSchedules = JSON.parse(localStorage.getItem('customSchedules') || '{}');
    if (customSchedules[dateString] && customSchedules[dateString].openEmployee) {
        return [customSchedules[dateString].openEmployee];
    }
    
    // 計算從每月1號開始到今天的工作日數量
    const monthStart = new Date(date.getFullYear(), date.getMonth(), 1);
    const workDaysCount = getWorkDaysBetween(monthStart, date);
    
    // 用工作日數量除以員工數量取餘數，得到當天值班的員工索引
    const employeeIndex = workDaysCount % openEmployees.length;
    
    return [openEmployees[employeeIndex].name];
}

// 獲得指定日期的關門員工名單
function getCloseEmployeesForDate(date) {
    if (closeEmployees.length === 0) return [];
    
    // 檢查當天是否為假日或週末，如果是則不排班
    if (isHolidayOrWeekend(date)) {
        return [];
    }
    
    // 檢查是否有自訂排班
    const dateString = formatDate(date);
    const customSchedules = JSON.parse(localStorage.getItem('customSchedules') || '{}');
    if (customSchedules[dateString] && customSchedules[dateString].closeEmployee) {
        return [customSchedules[dateString].closeEmployee];
    }
    
    // 計算從每月1號開始到今天的工作日數量
    const monthStart = new Date(date.getFullYear(), date.getMonth(), 1);
    const workDaysCount = getWorkDaysBetween(monthStart, date);
    
    // 用工作日數量除以員工數量取餘數，得到當天值班的員工索引
    const employeeIndex = workDaysCount % closeEmployees.length;
    
    return [closeEmployees[employeeIndex].name];
}

// 計算兩個日期之間的工作日數量
function getWorkDaysBetween(startDate, endDate) {
    let count = 0;
    let current = new Date(startDate);
    
    while (current <= endDate) {
        if (!isHolidayOrWeekend(current)) {
            count++;
        }
        current.setDate(current.getDate() + 1);
    }
    
    return count - 1; // 減1因為不包含當前日期
}

// 新增自訂假日
function addCustomHoliday() {
    const holidayDate = holidayDateInput.value;
    
    if (!holidayDate) {
        alert('請選擇假日日期');
        return;
    }
    
    if (customHolidays.includes(holidayDate)) {
        alert('此日期已經是假日了');
        return;
    }
    
    customHolidays.push(holidayDate);
    saveCustomHolidays();
    updateHolidayList();
    displayCalendar();
    
    // 清空輸入欄位
    holidayDateInput.value = '';
}

// 刪除自訂假日
function deleteCustomHoliday(holidayDate) {
    customHolidays = customHolidays.filter(date => date !== holidayDate);
    saveCustomHolidays();
    updateHolidayList();
    displayCalendar();
}

// 更新自訂假日列表
function updateHolidayList() {
    holidayListContainer.innerHTML = '';
    
    customHolidays.forEach(holidayDate => {
        const holidayItem = document.createElement('div');
        holidayItem.className = 'holiday-item';
        
        const date = new Date(holidayDate);
        const formattedDate = `${date.getFullYear()}/${(date.getMonth() + 1).toString().padStart(2, '0')}/${date.getDate().toString().padStart(2, '0')}`;
        
        holidayItem.innerHTML = `
            <span>${formattedDate}</span>
            <button class="delete-btn" onclick="deleteCustomHoliday('${holidayDate}')">刪除</button>
        `;
        
        holidayListContainer.appendChild(holidayItem);
    });
}

// 檢查是否為假日或週末
function isHolidayOrWeekend(date) {
    const dateString = formatDate(date);
    const dayOfWeek = date.getDay();
    
    // 檢查是否為週末 (週六=6, 週日=0)
    if (dayOfWeek === 0 || dayOfWeek === 6) {
        return true;
    }
    
    // 檢查是否為國定假日
    if (holidays.includes(dateString)) {
        return true;
    }
    
    // 檢查是否為自訂假日
    return customHolidays.includes(dateString);
}

// 格式化日期為 YYYY-MM-DD
function formatDate(date) {
    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const day = String(date.getDate()).padStart(2, '0');
    return `${year}-${month}-${day}`;
}

// 儲存員工資料到 localStorage
function saveEmployees() {
    localStorage.setItem('openEmployees', JSON.stringify(openEmployees));
    localStorage.setItem('closeEmployees', JSON.stringify(closeEmployees));
}

// 儲存自訂假日到 localStorage
function saveCustomHolidays() {
    localStorage.setItem('customHolidays', JSON.stringify(customHolidays));
}

// 從 localStorage 載入員工資料
function loadEmployees() {
    const savedOpen = localStorage.getItem('openEmployees');
    const savedClose = localStorage.getItem('closeEmployees');
    
    if (savedOpen) {
        openEmployees = JSON.parse(savedOpen);
    }
    
    if (savedClose) {
        closeEmployees = JSON.parse(savedClose);
    }
}

// 從 localStorage 載入自訂假日資料
function loadCustomHolidays() {
    const saved = localStorage.getItem('customHolidays');
    if (saved) {
        customHolidays = JSON.parse(saved);
    }
}

// 顯示員工彈出視窗
function showEmployeeModal(type) {
    const modal = document.getElementById('modal');
    const modalTitle = document.getElementById('modalTitle');
    const modalEmployeeList = document.getElementById('modalEmployeeList');
    
    if (type === 'open') {
        modalTitle.textContent = '開門名單管理';
        renderModalEmployeeList(openEmployees, 'open', modalEmployeeList);
    } else if (type === 'close') {
        modalTitle.textContent = '關門名單管理';
        renderModalEmployeeList(closeEmployees, 'close', modalEmployeeList);
    }
    
    modal.style.display = 'block';
    document.body.style.overflow = 'hidden'; // 防止背景滾動
}

// 顯示假日彈出視窗
function showHolidayModal() {
    const modal = document.getElementById('modal');
    const modalTitle = document.getElementById('modalTitle');
    const modalEmployeeList = document.getElementById('modalEmployeeList');
    
    modalTitle.textContent = '自訂假日列表';
    renderModalHolidayList(modalEmployeeList);
    
    modal.style.display = 'block';
    document.body.style.overflow = 'hidden';
}

// 關閉彈出視窗
function closeModal() {
    const modal = document.getElementById('modal');
    modal.style.display = 'none';
    document.body.style.overflow = 'auto'; // 恢復背景滾動
}

// 渲染彈出視窗中的員工列表
function renderModalEmployeeList(employees, type, container) {
    container.innerHTML = '';
    
    if (employees.length === 0) {
        container.innerHTML = '<p style="text-align: center; color: #85c1e9; font-style: italic;">目前沒有員工</p>';
        return;
    }
    
    // 添加排序說明
    const sortInfo = document.createElement('div');
    sortInfo.className = 'sort-info';
    sortInfo.innerHTML = '💡 拖拽員工項目來重新排序排班順序';
    container.appendChild(sortInfo);
    
    // 創建可排序的容器
    const sortableContainer = document.createElement('div');
    sortableContainer.className = 'sortable-container';
    sortableContainer.id = `sortable-${type}`;
    
    employees.forEach((emp, index) => {
        const empItem = document.createElement('div');
        empItem.className = 'employee-item sortable-item';
        empItem.draggable = true;
        empItem.dataset.empId = emp.id;
        
        empItem.innerHTML = `
            <span class="drag-handle">⋮⋮</span>
            <span class="emp-info"><strong>${index + 1}.</strong> ${emp.name}</span>
            <button class="delete-btn" onclick="deleteEmployee('${type}', ${emp.id})">刪除</button>
        `;
        
        sortableContainer.appendChild(empItem);
    });
    
    container.appendChild(sortableContainer);
    
    // 啟用拖拽排序
    enableSortable(sortableContainer, type);
}

// 渲染彈出視窗中的假日列表
function renderModalHolidayList(container) {
    container.innerHTML = '';
    
    if (customHolidays.length === 0) {
        container.innerHTML = '<p style="text-align: center; color: #85c1e9; font-style: italic;">目前沒有自訂假日</p>';
        return;
    }
    
    customHolidays.forEach(holidayDate => {
        const holidayItem = document.createElement('div');
        holidayItem.className = 'holiday-item';
        
        const date = new Date(holidayDate);
        const formattedDate = `${date.getFullYear()}/${(date.getMonth() + 1).toString().padStart(2, '0')}/${date.getDate().toString().padStart(2, '0')}`;
        
        holidayItem.innerHTML = `
            <span><strong>📅</strong> ${formattedDate}</span>
            <button class="delete-btn" onclick="deleteHolidayFromModal('${holidayDate}')">刪除</button>
        `;
        
        container.appendChild(holidayItem);
    });
}

// 從彈出視窗刪除員工
function deleteEmployee(type, id) {
    if (type === 'open') {
        deleteOpenEmployee(id);
        showEmployeeModal('open'); // 重新渲染彈出視窗
    } else if (type === 'close') {
        deleteCloseEmployee(id);
        showEmployeeModal('close'); // 重新渲染彈出視窗
    }
}

// 從彈出視窗刪除假日
function deleteHolidayFromModal(holidayDate) {
    deleteCustomHoliday(holidayDate);
    showHolidayModal(); // 重新渲染彈出視窗
}

// 點擊彈出視窗背景關閉
window.onclick = function(event) {
    const modal = document.getElementById('modal');
    if (event.target === modal) {
        closeModal();
    }
}

// ESC鍵關閉彈出視窗
document.addEventListener('keydown', function(event) {
    if (event.key === 'Escape') {
        closeModal();
    }
});

// 下載Excel檔案
function downloadExcel() {
    const year = currentDate.getFullYear();
    const month = currentDate.getMonth() + 1;
    const monthName = `${year}年${month}月`;
    
    // 獲取該月的天數
    const daysInMonth = new Date(year, month, 0).getDate();
    
    // 建立Excel表格，符合您要求的格式
    let xmlContent = `<?xml version="1.0"?>
<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet" xmlns:html="http://www.w3.org/TR/REC-html40">
<Styles>
<Style ss:ID="Default">
<Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
<Font ss:FontName="新細明體" ss:Size="11"/>
<Borders>
<Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
<Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
<Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
<Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
</Borders>
</Style>
<Style ss:ID="Title">
<Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
<Font ss:FontName="新細明體" ss:Size="14" ss:Bold="1"/>
<Borders>
<Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
<Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"/>
<Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
</Borders>
</Style>
<Style ss:ID="TitleBlue">
<Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
<Font ss:FontName="新細明體" ss:Size="14" ss:Bold="1" ss:Color="#0000FF"/>
<Borders>
<Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
<Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"/>
<Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
<Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
</Borders>
</Style>
<Style ss:ID="Header">
<Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>
<Font ss:FontName="新細明體" ss:Size="10" ss:Bold="1"/>
<Borders>
<Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
<Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
<Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
<Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
</Borders>
</Style>
<Style ss:ID="SubHeader">
<Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
<Font ss:FontName="新細明體" ss:Size="9" ss:Bold="1"/>
<Borders>
<Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
<Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
<Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
<Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
</Borders>
</Style>
<Style ss:ID="Holiday">
<Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
<Font ss:FontName="新細明體" ss:Size="11" ss:Color="#FF0000" ss:Bold="1"/>
<Borders>
<Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
<Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
<Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
<Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
</Borders>
</Style>
</Styles>
<Worksheet ss:Name="${monthName}排班表">
<Table>
<Column ss:Width="40"/>
<Column ss:Width="80"/>
<Column ss:Width="80"/>
<Column ss:Width="100"/>
<Column ss:Width="80"/>
<Column ss:Width="80"/>
<Column ss:Width="100"/>
<Column ss:Width="120"/>
<Column ss:Width="120"/>
<Column ss:Width="200"/>

<!-- 標題行 -->
<Row ss:Height="50">
<Cell ss:StyleID="Title" ss:MergeAcross="9"><Data ss:Type="String">台北富邦銀行民生分行 ${year - 1911} 年 ${month} 月保全排班/設定/解除登記表</Data></Cell>
</Row>

<!-- 第二層欄位標題行 -->
<Row ss:Height="20">
<Cell ss:StyleID="Header" ss:MergeDown="2"><Data ss:Type="String">日期</Data></Cell>
<Cell ss:StyleID="Header" ss:MergeAcross="1"><Data ss:Type="String">解除人員</Data></Cell>
<Cell ss:StyleID="Header"><Data ss:Type="String">解除</Data></Cell>
<Cell ss:StyleID="Header" ss:MergeAcross="1"><Data ss:Type="String">關門人員</Data></Cell>
<Cell ss:StyleID="Header"><Data ss:Type="String">設定</Data></Cell>
<Cell ss:StyleID="Header" ss:MergeDown="2"><Data ss:Type="String">主管蓋章&#10;確認設定/&#10;簽退時間&#10;相符(註1)</Data></Cell>
<Cell ss:StyleID="Header" ss:MergeDown="2"><Data ss:Type="String">單位主管/&#10;分行個金&#10;主管簽章</Data></Cell>
<Cell ss:StyleID="Header" ss:MergeDown="2"><Data ss:Type="String">備註 (實際解除或關門人員與原排班&#10;人員不同時，應予註明原&#10;因，並經主管簽核)</Data></Cell>
</Row>

<!-- 第三層欄位標題行 -->
<Row ss:Height="55">
<Cell ss:Index="2" ss:StyleID="SubHeader" ss:MergeAcross="1"><Data ss:Type="String">簽章</Data></Cell>
<Cell ss:StyleID="SubHeader"><Data ss:Type="String">時間</Data></Cell>
<Cell ss:StyleID="SubHeader" ss:MergeAcross="1"><Data ss:Type="String">簽章</Data></Cell>
<Cell ss:StyleID="SubHeader"><Data ss:Type="String">時間</Data></Cell>
</Row>

<!-- 第四層欄位標題行 -->
<Row ss:Height="20">
<Cell ss:Index="2" ss:StyleID="SubHeader"><Data ss:Type="String">排班</Data></Cell>
<Cell ss:StyleID="SubHeader"><Data ss:Type="String">實際</Data></Cell>
<Cell ss:StyleID="SubHeader"><Data ss:Type="String">(時/分)</Data></Cell>
<Cell ss:StyleID="SubHeader"><Data ss:Type="String">排班</Data></Cell>
<Cell ss:StyleID="SubHeader"><Data ss:Type="String">實際</Data></Cell>
<Cell ss:StyleID="SubHeader"><Data ss:Type="String">(時/分)</Data></Cell>
</Row>`;

    // 每一天的資料行
    for (let day = 1; day <= daysInMonth; day++) {
        const date = new Date(year, month - 1, day);
        const dateString = formatDate(date);
        
        let openValue = '';
        let closeValue = '';
        let isHoliday = false;
        
        // 檢查是否為假日
        if (customHolidays.includes(dateString) || holidays.includes(dateString)) {
            openValue = '休假';
            closeValue = '休假';
            isHoliday = true;
        } else if (date.getDay() === 0) {
            openValue = '日';
            closeValue = '日';
            isHoliday = true;
        } else if (date.getDay() === 6) {
            openValue = '六';
            closeValue = '六';
            isHoliday = true;
        } else {
            // 取得開門和關門員工
            const openEmployees = getOpenEmployeesForDate(date);
            const closeEmployees = getCloseEmployeesForDate(date);
            openValue = openEmployees.length > 0 ? openEmployees[0] : '';
            closeValue = closeEmployees.length > 0 ? closeEmployees[0] : '';
        }
        
        const styleID = isHoliday ? 'Holiday' : 'Default';
        
        xmlContent += `
<!-- 第${day}天 -->
<Row ss:Height="25">
<Cell ss:StyleID="Default"><Data ss:Type="Number">${day}</Data></Cell>
<Cell ss:StyleID="${styleID}"><Data ss:Type="String">${openValue}</Data></Cell>
<Cell ss:StyleID="Default"><Data ss:Type="String"></Data></Cell>
<Cell ss:StyleID="Default"><Data ss:Type="String"></Data></Cell>
<Cell ss:StyleID="${styleID}"><Data ss:Type="String">${closeValue}</Data></Cell>
<Cell ss:StyleID="Default"><Data ss:Type="String"></Data></Cell>
<Cell ss:StyleID="Default"><Data ss:Type="String"></Data></Cell>
<Cell ss:StyleID="Default"><Data ss:Type="String"></Data></Cell>
<Cell ss:StyleID="Default"><Data ss:Type="String"></Data></Cell>
</Row>`;
    }
    
    xmlContent += `
</Table>
</Worksheet>
</Workbook>`;
    
    // 下載檔案
    const blob = new Blob([xmlContent], { type: 'application/vnd.ms-excel;charset=utf-8' });
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = `台北富邦銀行民生分行${year - 1911}年${month}月排班表.xls`;
    link.click();
}

// 黑暗模式功能
function toggleDarkMode() {
    const currentTheme = document.documentElement.getAttribute('data-theme');
    const newTheme = currentTheme === 'dark' ? 'light' : 'dark';
    
    document.documentElement.setAttribute('data-theme', newTheme);
    updateDarkModeButton(newTheme);
    saveTheme(newTheme);
}

function updateDarkModeButton(theme) {
    if (theme === 'dark') {
        darkModeToggleBtn.innerHTML = '<span class="theme-icon">◐</span> 淺色';
    } else {
        darkModeToggleBtn.innerHTML = '<span class="theme-icon">◑</span> 黑暗';
    }
}

function saveTheme(theme) {
    localStorage.setItem('theme', theme);
}

function loadTheme() {
    const savedTheme = localStorage.getItem('theme') || 'light';
    document.documentElement.setAttribute('data-theme', savedTheme);
    updateDarkModeButton(savedTheme);
}

// 顯示日期編輯彈出視窗
function showDayEditModal(date) {
    const modal = document.getElementById('modal');
    const modalTitle = document.getElementById('modalTitle');
    const modalEmployeeList = document.getElementById('modalEmployeeList');
    
    const formattedDate = `${date.getFullYear()}年${date.getMonth() + 1}月${date.getDate()}日`;
    modalTitle.textContent = `編輯 ${formattedDate} 排班`;
    
    // 檢查是否為假日
    const dateString = formatDate(date);
    const isNationalHoliday = holidays.includes(dateString);
    const isCustomHoliday = customHolidays.includes(dateString);
    const isWeekend = date.getDay() === 0 || date.getDay() === 6;
    const isHoliday = isNationalHoliday || isCustomHoliday || isWeekend;
    
    // 獲取當前排班
    const dayOpenEmployees = isHoliday ? [] : getOpenEmployeesForDate(date);
    const dayCloseEmployees = isHoliday ? [] : getCloseEmployeesForDate(date);
    
    renderDayEditModal(date, dayOpenEmployees, dayCloseEmployees, isHoliday, modalEmployeeList);
    
    modal.style.display = 'block';
    document.body.style.overflow = 'hidden';
}

// 渲染日期編輯彈出視窗
function renderDayEditModal(date, openEmps, closeEmps, isHoliday, container) {
    container.innerHTML = '';
    
    if (isHoliday) {
        container.innerHTML = `
            <div style="text-align: center; padding: 20px; color: #dc3545; font-size: 16px;">
                <strong>🏖️ 此日為假日，不安排排班</strong>
                <p style="margin-top: 10px; color: #6c757d; font-size: 14px;">假日會自動跳過，使用下一個工作日的排班順序</p>
            </div>
        `;
        return;
    }
    
    // 創建編輯界面
    const editContainer = document.createElement('div');
    editContainer.className = 'day-edit-container';
    
    editContainer.innerHTML = `
        <div class="day-edit-section">
            <h4>🔓 開門人員</h4>
            <div class="current-assignment">
                <strong>當前排班：</strong> ${openEmps.length > 0 ? openEmps[0] : '無'}
            </div>
            <div class="employee-selector">
                <label>選擇開門人員：</label>
                <select id="openEmployeeSelect" class="employee-select">
                    <option value="">-- 選擇員工 --</option>
                    ${openEmployees.map(emp => 
                        `<option value="${emp.name}" ${openEmps.includes(emp.name) ? 'selected' : ''}>${emp.name}</option>`
                    ).join('')}
                </select>
            </div>
        </div>
        
        <div class="day-edit-section">
            <h4>🔒 關門人員</h4>
            <div class="current-assignment">
                <strong>當前排班：</strong> ${closeEmps.length > 0 ? closeEmps[0] : '無'}
            </div>
            <div class="employee-selector">
                <label>選擇關門人員：</label>
                <select id="closeEmployeeSelect" class="employee-select">
                    <option value="">-- 選擇員工 --</option>
                    ${closeEmployees.map(emp => 
                        `<option value="${emp.name}" ${closeEmps.includes(emp.name) ? 'selected' : ''}>${emp.name}</option>`
                    ).join('')}
                </select>
            </div>
        </div>
        
        <div class="day-edit-actions">
            <button class="save-btn" onclick="saveDayEdit('${formatDate(date)}')">💾 儲存變更</button>
            <button class="reset-btn" onclick="resetDayEdit('${formatDate(date)}')">🔄 重置為自動排班</button>
        </div>
    `;
    
    container.appendChild(editContainer);
}

// 儲存日期編輯
function saveDayEdit(dateString) {
    const openSelect = document.getElementById('openEmployeeSelect');
    const closeSelect = document.getElementById('closeEmployeeSelect');
    
    const customSchedules = JSON.parse(localStorage.getItem('customSchedules') || '{}');
    
    customSchedules[dateString] = {
        openEmployee: openSelect.value,
        closeEmployee: closeSelect.value
    };
    
    localStorage.setItem('customSchedules', JSON.stringify(customSchedules));
    
    closeModal();
    displayCalendar();
    
    // 顯示成功訊息
    showNotification('排班已更新！', 'success');
}

// 重置日期編輯
function resetDayEdit(dateString) {
    const customSchedules = JSON.parse(localStorage.getItem('customSchedules') || '{}');
    delete customSchedules[dateString];
    localStorage.setItem('customSchedules', JSON.stringify(customSchedules));
    
    closeModal();
    displayCalendar();
    
    showNotification('已重置為自動排班！', 'info');
}

// 顯示通知
function showNotification(message, type = 'info') {
    const notification = document.createElement('div');
    notification.className = `notification ${type}`;
    notification.textContent = message;
    notification.style.cssText = `
        position: fixed;
        top: 20px;
        right: 20px;
        padding: 12px 20px;
        border-radius: 8px;
        color: white;
        font-weight: bold;
        z-index: 10000;
    `;
    
    if (type === 'success') notification.style.background = '#28a745';
    if (type === 'info') notification.style.background = '#17a2b8';
    if (type === 'error') notification.style.background = '#dc3545';
    
    document.body.appendChild(notification);
    
    setTimeout(() => notification.remove(), 3000);
}

// 啟用拖拽排序功能
function enableSortable(container, type) {
    let draggedElement = null;
    
    container.addEventListener('dragstart', function(e) {
        draggedElement = e.target;
        e.target.style.opacity = '0.5';
    });
    
    container.addEventListener('dragend', function(e) {
        e.target.style.opacity = '';
        draggedElement = null;
        
        // 更新排序
        updateEmployeeOrder(container, type);
    });
    
    container.addEventListener('dragover', function(e) {
        e.preventDefault();
    });
    
    container.addEventListener('drop', function(e) {
        e.preventDefault();
        
        if (draggedElement !== e.target && e.target.classList.contains('sortable-item')) {
            const allItems = Array.from(container.querySelectorAll('.sortable-item'));
            const draggedIndex = allItems.indexOf(draggedElement);
            const targetIndex = allItems.indexOf(e.target);
            
            if (draggedIndex < targetIndex) {
                e.target.parentNode.insertBefore(draggedElement, e.target.nextSibling);
            } else {
                e.target.parentNode.insertBefore(draggedElement, e.target);
            }
        }
    });
}

// 更新員工順序
function updateEmployeeOrder(container, type) {
    const items = container.querySelectorAll('.sortable-item');
    const newOrder = Array.from(items).map(item => {
        const empId = parseFloat(item.dataset.empId);
        const employees = type === 'open' ? openEmployees : closeEmployees;
        return employees.find(emp => emp.id === empId);
    });
    
    if (type === 'open') {
        openEmployees = newOrder;
    } else {
        closeEmployees = newOrder;
    }
    
    // 更新序號顯示
    items.forEach((item, index) => {
        const empInfo = item.querySelector('.emp-info');
        const empName = empInfo.textContent.split('. ')[1];
        empInfo.innerHTML = `<strong>${index + 1}.</strong> ${empName}`;
    });
    
    saveEmployees();
    updateEmployeeList();
    displayCalendar();
    
    showNotification('排序已更新！', 'success');
}