/**
 * 主应用模块
 * 整合所有功能模块，处理页面路由和用户交互
 */

// 应用状态
const AppState = {
    currentView: 'templates',
    currentTemplate: null,
    currentRecord: null,
    editingTemplate: null,
    formCollector: null,
    tablesCollector: null,
    fieldConfigManager: null,
    tableConfigManager: null
};

// ==========================================
// Toast 提示
// ==========================================

/**
 * 显示 Toast 提示
 * @param {string} message - 消息内容
 * @param {string} type - 类型: 'success' | 'error' | 'warning'
 * @param {number} duration - 显示时长（毫秒）
 */
function showToast(message, type = 'success', duration = 3000) {
    const container = document.getElementById('toast-container');
    
    const toast = document.createElement('div');
    toast.className = `toast ${type}`;
    
    const iconSvg = type === 'success' 
        ? '<svg width="14" height="14" viewBox="0 0 14 14" fill="white"><path d="M5 7l2 2 4-4" stroke="white" stroke-width="2" fill="none"/></svg>'
        : type === 'error'
        ? '<svg width="14" height="14" viewBox="0 0 14 14" fill="white"><path d="M4 4l6 6M10 4l-6 6" stroke="white" stroke-width="2"/></svg>'
        : '<svg width="14" height="14" viewBox="0 0 14 14" fill="white"><path d="M7 4v4M7 10v1" stroke="white" stroke-width="2"/></svg>';
    
    toast.innerHTML = `
        <div class="toast-icon">${iconSvg}</div>
        <div class="toast-content">
            <div class="toast-message">${message}</div>
        </div>
    `;
    
    container.appendChild(toast);
    
    // 自动移除
    setTimeout(() => {
        toast.classList.add('hide');
        setTimeout(() => toast.remove(), 200);
    }, duration);
}

// ==========================================
// 模态框
// ==========================================

/**
 * 显示模态框
 * @param {Object} options - 模态框选项
 */
function showModal(options) {
    const { title, content, buttons = [], onClose = null } = options;
    
    const overlay = document.getElementById('modal-overlay');
    const modal = document.getElementById('modal');
    const modalTitle = document.getElementById('modal-title');
    const modalBody = document.getElementById('modal-body');
    const modalFooter = document.getElementById('modal-footer');
    
    modalTitle.textContent = title;
    modalBody.innerHTML = typeof content === 'string' ? content : '';
    if (typeof content !== 'string') {
        modalBody.innerHTML = '';
        modalBody.appendChild(content);
    }
    
    // 渲染按钮
    modalFooter.innerHTML = '';
    buttons.forEach(btn => {
        const button = document.createElement('button');
        button.className = `btn ${btn.class || 'btn-secondary'}`;
        button.textContent = btn.text;
        button.onclick = () => {
            if (btn.onClick) {
                btn.onClick();
            }
            if (btn.closeModal !== false) {
                hideModal();
            }
        };
        modalFooter.appendChild(button);
    });
    
    // 关闭按钮事件
    document.getElementById('modal-close').onclick = () => {
        if (onClose) onClose();
        hideModal();
    };
    
    overlay.classList.add('show');
}

/**
 * 隐藏模态框
 */
function hideModal() {
    const overlay = document.getElementById('modal-overlay');
    overlay.classList.remove('show');
}

// ==========================================
// 视图切换
// ==========================================

/**
 * 切换视图
 * @param {string} viewName - 视图名称
 */
function switchView(viewName) {
    // 隐藏所有视图
    document.querySelectorAll('.view').forEach(view => {
        view.classList.remove('active');
    });
    
    // 显示目标视图
    const targetView = document.getElementById(`view-${viewName}`);
    if (targetView) {
        targetView.classList.add('active');
    }
    
    // 更新导航状态
    document.querySelectorAll('.nav-item').forEach(item => {
        item.classList.remove('active');
        if (item.dataset.view === viewName) {
            item.classList.add('active');
        }
    });
    
    AppState.currentView = viewName;
    
    // 根据视图加载数据
    if (viewName === 'templates') {
        loadTemplateList();
    } else if (viewName === 'history') {
        loadHistoryList();
    }
}

// ==========================================
// 移动端侧边栏
// ==========================================

function setSidebarOpen(isOpen) {
    const sidebar = document.querySelector('.sidebar');
    const backdrop = document.getElementById('sidebar-backdrop');
    if (!sidebar || !backdrop) return;
    if (isOpen) {
        sidebar.classList.add('show');
        backdrop.classList.add('show');
    } else {
        sidebar.classList.remove('show');
        backdrop.classList.remove('show');
    }
}

// ==========================================
// 模板列表视图
// ==========================================

/**
 * 加载模板列表
 */
async function loadTemplateList() {
    const grid = document.getElementById('template-grid');
    const emptyState = document.getElementById('empty-templates');
    
    try {
        const templates = await StorageAdapter.getAllTemplates();
        
        // 清空现有内容（保留空状态元素）
        const cards = grid.querySelectorAll('.template-card');
        cards.forEach(card => card.remove());
        
        if (templates.length === 0) {
            emptyState.style.display = 'flex';
            return;
        }
        
        emptyState.style.display = 'none';
        
        // 渲染模板卡片
        for (const template of templates) {
            const stats = await Template.getTemplateStats(template.id);
            const card = createTemplateCard(template, stats);
            grid.insertBefore(card, emptyState);
        }
        
    } catch (error) {
        console.error('加载模板列表失败:', error);
        showToast('加载模板列表失败', 'error');
    }
}

/**
 * 创建模板卡片
 * @param {Object} template - 模板数据
 * @param {Object} stats - 模板统计信息
 * @returns {HTMLElement}
 */
function createTemplateCard(template, stats) {
    const card = document.createElement('div');
    card.className = 'template-card';
    card.dataset.templateId = template.id;
    
    const lastUsedText = stats.lastUsedAt 
        ? `最近使用: ${formatDateTime(stats.lastUsedAt)}`
        : '暂未使用';
    
    card.innerHTML = `
        <div class="template-card-header">
            <div class="template-card-icon">
                <svg width="24" height="24" viewBox="0 0 24 24" fill="currentColor">
                    <path d="M14 2H6a2 2 0 00-2 2v16a2 2 0 002 2h12a2 2 0 002-2V8l-6-6z"/>
                    <path d="M14 2v6h6" fill="white"/>
                </svg>
            </div>
            <div class="template-card-menu">
                <button class="template-card-menu-btn" title="更多操作">
                    <svg width="16" height="16" viewBox="0 0 16 16" fill="currentColor">
                        <circle cx="8" cy="3" r="1.5"/>
                        <circle cx="8" cy="8" r="1.5"/>
                        <circle cx="8" cy="13" r="1.5"/>
                    </svg>
                </button>
                <div class="template-card-dropdown">
                    <button class="template-card-dropdown-item" data-action="edit">
                        <svg width="14" height="14" viewBox="0 0 14 14" fill="currentColor">
                            <path d="M10 2l2 2-8 8H2v-2l8-8z"/>
                        </svg>
                        编辑模板
                    </button>
                    <button class="template-card-dropdown-item" data-action="export">
                        <svg width="14" height="14" viewBox="0 0 14 14" fill="currentColor">
                            <path d="M7 1l5 5h-3v5H5V6H2l5-5zM1 12v1h12v-1H1z"/>
                        </svg>
                        导出模板包
                    </button>
                    <button class="template-card-dropdown-item" data-action="history">
                        <svg width="14" height="14" viewBox="0 0 14 14" fill="currentColor">
                            <circle cx="7" cy="7" r="6" stroke="currentColor" fill="none"/>
                            <path d="M7 4v4l3 1.5" stroke="currentColor" fill="none"/>
                        </svg>
                        查看历史
                    </button>
                    <button class="template-card-dropdown-item danger" data-action="delete">
                        <svg width="14" height="14" viewBox="0 0 14 14" fill="currentColor">
                            <path d="M3 4h8M5 4V3a1 1 0 011-1h2a1 1 0 011 1v1M4 4l.5 7a1 1 0 001 1h3a1 1 0 001-1l.5-7"/>
                        </svg>
                        删除模板
                    </button>
                </div>
            </div>
        </div>
        <div class="template-card-title">${escapeHtml(template.name)}</div>
        <div class="template-card-desc">${escapeHtml(template.description || '暂无描述')}</div>
        <div class="template-card-meta">
            <span class="template-card-meta-item">
                <svg width="12" height="12" viewBox="0 0 12 12" fill="currentColor">
                    <path d="M4 3h4v1H4V3zm0 2h4v1H4V5zm0 2h2v1H4V7z"/>
                    <rect x="2" y="1" width="8" height="10" rx="1" stroke="currentColor" fill="none"/>
                </svg>
                ${stats.recordCount} 条记录
            </span>
            <span class="template-card-meta-item">${lastUsedText}</span>
        </div>
        <div class="template-card-action">
            <button class="btn btn-primary btn-start-fill">开始填写</button>
        </div>
    `;
    
    // 绑定事件
    const menuBtn = card.querySelector('.template-card-menu-btn');
    const dropdown = card.querySelector('.template-card-dropdown');
    
    menuBtn.addEventListener('click', (e) => {
        e.stopPropagation();
        // 关闭其他下拉菜单
        document.querySelectorAll('.template-card-dropdown.show').forEach(d => {
            if (d !== dropdown) d.classList.remove('show');
        });
        dropdown.classList.toggle('show');
    });
    
    // 下拉菜单操作
    dropdown.querySelectorAll('.template-card-dropdown-item').forEach(item => {
        item.addEventListener('click', (e) => {
            e.stopPropagation();
            dropdown.classList.remove('show');
            const action = item.dataset.action;
            handleTemplateAction(template.id, action);
        });
    });
    
    // 开始填写按钮
    card.querySelector('.btn-start-fill').addEventListener('click', (e) => {
        e.stopPropagation();
        startFillForm(template.id);
    });
    
    return card;
}

/**
 * 处理模板操作
 * @param {string} templateId - 模板 ID
 * @param {string} action - 操作类型
 */
async function handleTemplateAction(templateId, action) {
    switch (action) {
        case 'edit':
            editTemplate(templateId);
            break;
        case 'export':
            try {
                await Export.exportTemplatePack(templateId);
                showToast('模板包导出成功', 'success');
            } catch (error) {
                showToast(error.message, 'error');
            }
            break;
        case 'history':
            // 切换到历史视图并筛选该模板
            document.getElementById('history-template-filter').value = templateId;
            switchView('history');
            // 注意：switchView 会自动调用 loadHistoryList，不需要再手动调用
            break;
        case 'delete':
            confirmDeleteTemplate(templateId);
            break;
    }
}

/**
 * 确认删除模板
 * @param {string} templateId - 模板 ID
 */
async function confirmDeleteTemplate(templateId) {
    const template = await StorageAdapter.getTemplateById(templateId);
    const stats = await Template.getTemplateStats(templateId);
    
    showModal({
        title: '确认删除',
        content: `
            <p>确定要删除模板 "<strong>${escapeHtml(template.name)}</strong>" 吗？</p>
            ${stats.recordCount > 0 ? `<p class="text-danger">该模板下有 ${stats.recordCount} 条历史记录，删除后将一并删除！</p>` : ''}
            <p>此操作无法撤销。</p>
        `,
        buttons: [
            { text: '取消', class: 'btn-secondary' },
            { 
                text: '删除', 
                class: 'btn-danger',
                onClick: async () => {
                    try {
                        await Storage.deleteRecordsByTemplateId(templateId);
                        await StorageAdapter.deleteTemplate(templateId);
                        showToast('模板已删除', 'success');
                        loadTemplateList();
                    } catch (error) {
                        showToast('删除失败: ' + error.message, 'error');
                    }
                }
            }
        ]
    });
}

// ==========================================
// 模板编辑视图
// ==========================================

/**
 * 新建模板
 */
function newTemplate() {
    AppState.editingTemplate = Template.createEmptyTemplate();
    renderTemplateEditView();
    switchView('template-edit');
    document.getElementById('edit-title').textContent = '新建模板';
}

/**
 * 编辑模板
 * @param {string} templateId - 模板 ID
 */
async function editTemplate(templateId) {
    try {
        const template = await StorageAdapter.getTemplateById(templateId);
        if (!template) {
            showToast('模板不存在', 'error');
            return;
        }
        
        AppState.editingTemplate = { ...template };
        renderTemplateEditView();
        switchView('template-edit');
        document.getElementById('edit-title').textContent = '编辑模板';
        
    } catch (error) {
        showToast('加载模板失败', 'error');
    }
}

/**
 * 渲染模板编辑视图
 */
function renderTemplateEditView() {
    const template = AppState.editingTemplate;
    
    // 基本信息
    document.getElementById('template-name').value = template.name || '';
    document.getElementById('template-desc').value = template.description || '';
    
    // Word 文件
    const fileInfo = document.getElementById('word-file-info');
    const uploadContent = document.querySelector('.file-upload-content');
    
    if (template.wordFile && template.wordFile.name) {
        fileInfo.style.display = 'flex';
        uploadContent.style.display = 'none';
        fileInfo.querySelector('.file-name').textContent = template.wordFile.name;
    } else {
        fileInfo.style.display = 'none';
        uploadContent.style.display = 'block';
    }
    
    // 字段配置
    const fieldList = document.getElementById('field-list');
    AppState.fieldConfigManager = Form.renderFieldConfigList(
        fieldList,
        template.fields || [],
        (fields) => {
            AppState.editingTemplate.fields = fields;
        }
    );
    
    // 明细表配置
    const tableConfigList = document.getElementById('table-config-list');
    AppState.tableConfigManager = Form.renderTableConfigList(
        tableConfigList,
        template.detailTables || [],
        (tables) => {
            AppState.editingTemplate.detailTables = tables;
        }
    );
}

/**
 * 保存模板
 */
async function saveTemplate() {
    const template = AppState.editingTemplate;
    
    // 收集数据
    template.name = document.getElementById('template-name').value.trim();
    template.description = document.getElementById('template-desc').value.trim();
    
    if (AppState.fieldConfigManager) {
        template.fields = AppState.fieldConfigManager.getFields();
    }
    
    if (AppState.tableConfigManager) {
        template.detailTables = AppState.tableConfigManager.getTables();
    }
    
    // 验证
    const validation = Template.validateTemplate(template);
    if (!validation.valid) {
        showModal({
            title: '验证失败',
            content: `
                <p>请修正以下问题：</p>
                <ul style="color: var(--color-danger); margin-top: 12px;">
                    ${validation.errors.map(e => `<li>${escapeHtml(e)}</li>`).join('')}
                </ul>
            `,
            buttons: [{ text: '确定', class: 'btn-primary' }]
        });
        return;
    }
    
    try {
        const saved = await StorageAdapter.saveTemplate(template);
        AppState.editingTemplate = saved;
        
        showToast('模板保存成功', 'success');
        switchView('templates');
        
    } catch (error) {
        showToast('保存失败: ' + error.message, 'error');
    }
}

/**
 * 模板自检
 */
async function checkTemplate() {
    const template = AppState.editingTemplate;
    
    // 先收集当前数据
    template.name = document.getElementById('template-name').value.trim();
    
    if (AppState.fieldConfigManager) {
        template.fields = AppState.fieldConfigManager.getFields();
    }
    
    if (AppState.tableConfigManager) {
        template.detailTables = AppState.tableConfigManager.getTables();
    }
    
    if (!template.wordFile || !template.wordFile.data) {
        showToast('请先上传 Word 模板文件', 'warning');
        return;
    }
    
    try {
        const result = await Template.checkTemplateConsistency(template);
        
        let content = '';
        
        if (result.valid) {
            content = `
                <div style="text-align: center; padding: 20px;">
                    <svg width="48" height="48" viewBox="0 0 48 48" fill="#34C759">
                        <circle cx="24" cy="24" r="20" fill="#34C759"/>
                        <path d="M16 24l6 6 12-12" stroke="white" stroke-width="3" fill="none"/>
                    </svg>
                    <p style="margin-top: 16px; font-size: 16px; font-weight: 500;">模板配置正确</p>
                    <p style="color: var(--color-text-secondary); margin-top: 8px;">所有字段都已正确映射到 Word 模板</p>
                </div>
            `;
        } else {
            content = '<div style="max-height: 300px; overflow-y: auto;">';
            
            if (result.matched.length > 0) {
                content += `
                    <div style="margin-bottom: 16px;">
                        <p style="font-weight: 500; color: #34C759;">✓ 已匹配的占位符 (${result.matched.length})</p>
                        <p style="color: var(--color-text-secondary); font-size: 13px;">${result.matched.join(', ')}</p>
                    </div>
                `;
            }
            
            if (result.missingInWord.length > 0) {
                content += `
                    <div style="margin-bottom: 16px;">
                        <p style="font-weight: 500; color: #FF9500;">⚠ 配置了但 Word 中未找到 (${result.missingInWord.length})</p>
                        <p style="color: var(--color-text-secondary); font-size: 13px;">${result.missingInWord.join(', ')}</p>
                        <p style="color: var(--color-text-tertiary); font-size: 12px; margin-top: 4px;">这些字段在 Word 模板中没有对应的占位符</p>
                    </div>
                `;
            }
            
            if (result.missingInConfig.length > 0) {
                content += `
                    <div style="margin-bottom: 16px;">
                        <p style="font-weight: 500; color: #FF3B30;">✗ Word 中有但未配置 (${result.missingInConfig.length})</p>
                        <p style="color: var(--color-text-secondary); font-size: 13px;">${result.missingInConfig.join(', ')}</p>
                        <p style="color: var(--color-text-tertiary); font-size: 12px; margin-top: 4px;">请添加这些字段的配置，或从 Word 模板中移除</p>
                    </div>
                `;
            }
            
            content += '</div>';
        }
        
        showModal({
            title: '模板自检结果',
            content,
            buttons: [{ text: '确定', class: 'btn-primary' }]
        });
        
    } catch (error) {
        showToast('模板自检失败: ' + error.message, 'error');
    }
}

// ==========================================
// 表单填写视图
// ==========================================

/**
 * 开始填写表单
 * @param {string} templateId - 模板 ID
 * @param {Object} recordData - 已有记录数据（可选）
 */
async function startFillForm(templateId, recordData = null) {
    try {
        const template = await StorageAdapter.getTemplateById(templateId);
        if (!template) {
            showToast('模板不存在', 'error');
            return;
        }
        
        AppState.currentTemplate = template;
        AppState.currentRecord = recordData;
        
        // 更新标题
        document.getElementById('form-title').textContent = template.name;
        
        // 填充文件名
        const fileNameInput = document.getElementById('output-filename');
        fileNameInput.value = recordData?.fileName || '';
        
        // 渲染表单
        const formContainer = document.getElementById('dynamic-form');
        const tablesContainer = document.getElementById('detail-tables-container');
        
        const fieldData = recordData?.data || {};
        const tableData = recordData?.tables || {};
        
        AppState.formCollector = Form.renderDynamicForm(
            formContainer,
            template.fields,
            fieldData
        );
        
        AppState.tablesCollector = Form.renderDetailTables(
            tablesContainer,
            template.detailTables,
            tableData
        );
        
        switchView('form-fill');
        
    } catch (error) {
        console.error('加载表单失败:', error);
        showToast('加载表单失败', 'error');
    }
}

/**
 * 保存记录
 */
async function saveRecord() {
    if (!AppState.currentTemplate || !AppState.formCollector) {
        showToast('请先选择模板', 'error');
        return;
    }
    
    // 获取文件名
    const fileNameInput = document.getElementById('output-filename');
    const fileName = fileNameInput.value.trim();
    
    if (!fileName) {
        showToast('请输入文件名', 'warning');
        fileNameInput.focus();
        return;
    }
    
    // 验证表单
    const validation = AppState.formCollector.validate();
    if (!validation.valid) {
        showToast(validation.errors[0], 'warning');
        return;
    }
    
    try {
        const fieldData = AppState.formCollector.getData();
        const tableData = AppState.tablesCollector?.getData() || {};
        
        const record = {
            id: AppState.currentRecord?.id || null,
            templateId: AppState.currentTemplate.id,
            fileName: fileName,
            data: fieldData,
            tables: tableData
        };
        
        const saved = await StorageAdapter.saveRecord(record);
        AppState.currentRecord = saved;
        
        showToast('记录保存成功', 'success');
        
    } catch (error) {
        showToast('保存失败: ' + error.message, 'error');
    }
}

/**
 * 保存记录（带文件名，内部使用）
 */
async function saveRecordWithFileName(fileName) {
    if (!AppState.currentTemplate || !AppState.formCollector) {
        return;
    }
    
    try {
        const fieldData = AppState.formCollector.getData();
        const tableData = AppState.tablesCollector?.getData() || {};
        
        const record = {
            id: AppState.currentRecord?.id || null,
            templateId: AppState.currentTemplate.id,
            fileName: fileName,
            data: fieldData,
            tables: tableData
        };
        
        const saved = await StorageAdapter.saveRecord(record);
        AppState.currentRecord = saved;
        
    } catch (error) {
        console.error('保存记录失败:', error);
    }
}

/**
 * 生成 Word 文档
 */
async function generateWordDoc() {
    if (!AppState.currentTemplate || !AppState.formCollector) {
        showToast('请先选择模板', 'error');
        return;
    }
    
    // 获取文件名
    const fileNameInput = document.getElementById('output-filename');
    const fileName = fileNameInput.value.trim();
    
    if (!fileName) {
        showToast('请输入文件名', 'warning');
        fileNameInput.focus();
        return;
    }
    
    // 验证表单
    const validation = AppState.formCollector.validate();
    if (!validation.valid) {
        showToast(validation.errors[0], 'warning');
        return;
    }
    
    try {
        const fieldData = AppState.formCollector.getData();
        const tableData = AppState.tablesCollector?.getData() || {};
        
        // 生成完整文件名
        const fullFileName = fileName.endsWith('.docx') ? fileName : `${fileName}.docx`;
        
        await Word.generateAndDownloadWord(
            AppState.currentTemplate,
            fieldData,
            tableData,
            fullFileName
        );
        
        showToast('Word 文档生成成功', 'success');
        
        // 自动保存记录（带文件名）
        await saveRecordWithFileName(fileName);
        
    } catch (error) {
        console.error('生成 Word 失败:', error);
        showToast(error.message, 'error');
    }
}

/**
 * 从历史载入
 */
async function loadFromHistory() {
    if (!AppState.currentTemplate) {
        showToast('请先选择模板', 'error');
        return;
    }
    
    try {
        const records = await Storage.getRecordsByTemplateId(AppState.currentTemplate.id);
        
        if (records.length === 0) {
            showToast('暂无历史记录', 'warning');
            return;
        }
        
        // 创建历史列表
        const listHtml = records.map(record => {
            const titleValue = record.fileName || '未命名';
            return `
                <div class="history-item" data-record-id="${record.id}" style="cursor: pointer; margin-bottom: 8px;">
                    <div class="history-item-info">
                        <div class="history-item-title">${escapeHtml(titleValue)}</div>
                        <div class="history-item-meta">
                            <span>${formatDateTime(record.createdAt)}</span>
                        </div>
                    </div>
                </div>
            `;
        }).join('');
        
        const content = document.createElement('div');
        content.style.maxHeight = '400px';
        content.style.overflowY = 'auto';
        content.innerHTML = listHtml;
        
        // 绑定点击事件
        content.querySelectorAll('.history-item').forEach(item => {
            item.addEventListener('click', async () => {
                const recordId = item.dataset.recordId;
                const record = await StorageAdapter.getRecordById(recordId);
                if (record) {
                    AppState.currentRecord = record;
                    // 填充文件名
                    document.getElementById('output-filename').value = record.fileName || '';
                    AppState.formCollector.setData(record.data);
                    AppState.tablesCollector?.setData(record.tables || {});
                    hideModal();
                    showToast('已载入历史记录', 'success');
                }
            });
        });
        
        showModal({
            title: '选择历史记录',
            content,
            buttons: [{ text: '取消', class: 'btn-secondary' }]
        });
        
    } catch (error) {
        showToast('加载历史失败: ' + error.message, 'error');
    }
}

// ==========================================
// 历史记录视图
// ==========================================

/**
 * 加载历史记录列表
 * @param {string} templateId - 模板 ID（可选）
 */
async function loadHistoryList(templateId = null) {
    const list = document.getElementById('history-list');
    const emptyState = document.getElementById('empty-history');
    const filterSelect = document.getElementById('history-template-filter');
    
    try {
        // 更新模板筛选下拉
        const templates = await StorageAdapter.getAllTemplates();
        filterSelect.innerHTML = '<option value="">全部模板</option>' + 
            templates.map(t => `<option value="${t.id}" ${t.id === templateId ? 'selected' : ''}>${escapeHtml(t.name)}</option>`).join('');
        
        // 获取筛选条件
        const selectedTemplateId = templateId || filterSelect.value || null;
        const startDate = document.getElementById('history-date-start').value || null;
        const endDate = document.getElementById('history-date-end').value || null;
        
        // 获取记录
        const records = await StorageAdapter.getRecordsByDateRange(selectedTemplateId, startDate, endDate);
        
        // 清空列表
        list.querySelectorAll('.history-item').forEach(item => item.remove());
        
        if (records.length === 0) {
            emptyState.style.display = 'flex';
            return;
        }
        
        emptyState.style.display = 'none';
        
        // 渲染记录
        for (const record of records) {
            const template = await StorageAdapter.getTemplateById(record.templateId);
            const item = createHistoryItem(record, template);
            list.insertBefore(item, emptyState);
        }
        
    } catch (error) {
        console.error('加载历史记录失败:', error);
        showToast('加载历史记录失败', 'error');
    }
}

/**
 * 创建历史记录项
 * @param {Object} record - 记录数据
 * @param {Object} template - 模板数据
 * @returns {HTMLElement}
 */
function createHistoryItem(record, template) {
    const item = document.createElement('div');
    item.className = 'history-item';
    item.dataset.recordId = record.id;
    
    const titleValue = record.fileName || '未命名';
    
    item.innerHTML = `
        <div class="history-item-info">
            <div class="history-item-title">${escapeHtml(titleValue)}</div>
            <div class="history-item-meta">
                <span>
                    <svg width="12" height="12" viewBox="0 0 12 12" fill="currentColor">
                        <rect x="2" y="1" width="8" height="10" rx="1" stroke="currentColor" fill="none"/>
                    </svg>
                    ${escapeHtml(template?.name || '未知模板')}
                </span>
                <span>
                    <svg width="12" height="12" viewBox="0 0 12 12" fill="currentColor">
                        <circle cx="6" cy="6" r="5" stroke="currentColor" fill="none"/>
                        <path d="M6 3v4l2.5 1.5" stroke="currentColor" fill="none"/>
                    </svg>
                    创建: ${formatDateTime(record.createdAt)}
                </span>
                <span>修改: ${formatDateTime(record.updatedAt)}</span>
            </div>
        </div>
        <div class="history-item-actions">
            <button class="btn btn-small btn-secondary btn-open">打开</button>
            <button class="btn btn-small btn-secondary btn-copy">复制</button>
            <button class="btn btn-small btn-text text-danger btn-delete">删除</button>
        </div>
    `;
    
    // 绑定事件
    item.querySelector('.btn-open').addEventListener('click', () => {
        if (template) {
            startFillForm(template.id, record);
        } else {
            showToast('模板不存在', 'error');
        }
    });
    
    item.querySelector('.btn-copy').addEventListener('click', async () => {
        if (template) {
            // 复制为新记录
            const newRecord = {
                ...record,
                id: null,
                createdAt: null,
                updatedAt: null
            };
            startFillForm(template.id, newRecord);
            showToast('已复制为新记录', 'success');
        }
    });
    
    item.querySelector('.btn-delete').addEventListener('click', () => {
        confirmDeleteRecord(record.id, titleValue);
    });
    
    return item;
}

/**
 * 确认删除记录
 * @param {string} recordId - 记录 ID
 * @param {string} title - 记录标题
 */
function confirmDeleteRecord(recordId, title) {
    showModal({
        title: '确认删除',
        content: `<p>确定要删除记录 "<strong>${escapeHtml(title)}</strong>" 吗？</p><p>此操作无法撤销。</p>`,
        buttons: [
            { text: '取消', class: 'btn-secondary' },
            { 
                text: '删除', 
                class: 'btn-danger',
                onClick: async () => {
                    try {
                        await StorageAdapter.deleteRecord(recordId);
                        showToast('记录已删除', 'success');
                        loadHistoryList();
                    } catch (error) {
                        showToast('删除失败', 'error');
                    }
                }
            }
        ]
    });
}

// ==========================================
// 导入导出
// ==========================================

/**
 * 导出数据包
 */
async function exportRecords() {
    const templateId = document.getElementById('history-template-filter').value;
    
    if (!templateId) {
        showToast('请先选择一个模板', 'warning');
        return;
    }
    
    const startDate = document.getElementById('history-date-start').value || null;
    const endDate = document.getElementById('history-date-end').value || null;
    
    try {
        await Export.exportRecordsPack(templateId, startDate, endDate);
        showToast('数据包导出成功', 'success');
    } catch (error) {
        showToast(error.message, 'error');
    }
}

// ==========================================
// 工具函数
// ==========================================

/**
 * HTML 转义
 * @param {string} text - 原始文本
 * @returns {string}
 */
function escapeHtml(text) {
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
}

/**
 * 格式化日期时间
 * @param {string} isoString - ISO 日期字符串
 * @returns {string}
 */
function formatDateTime(isoString) {
    if (!isoString) return '-';
    const date = new Date(isoString);
    return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}-${String(date.getDate()).padStart(2, '0')} ${String(date.getHours()).padStart(2, '0')}:${String(date.getMinutes()).padStart(2, '0')}`;
}

// ==========================================
// 初始化
// ==========================================

/**
 * 初始化应用
 */
async function initApp() {
    try {
        // 初始化存储（自动选择文件存储或浏览器存储）
        await StorageAdapter.init();
        
        // 显示存储模式提示
        const storageMode = StorageAdapter.getStorageMode();
        if (storageMode === 'file') {
            showToast('✓ 使用本地文件存储模式', 'success', 2000);
        } else {
            showToast('使用浏览器存储模式，建议切换到文件存储', 'warning', 3000);
        }
        
        // 绑定导航事件
        document.querySelectorAll('.nav-item').forEach(item => {
            item.addEventListener('click', (e) => {
                e.preventDefault();
                const view = item.dataset.view;
                if (view) {
                    switchView(view);
                    setSidebarOpen(false);
                }
            });
        });

        // 移动端侧边栏开关
        document.querySelectorAll('.btn-nav-toggle').forEach(btn => {
            btn.addEventListener('click', () => {
                const sidebar = document.querySelector('.sidebar');
                const isOpen = sidebar && sidebar.classList.contains('show');
                setSidebarOpen(!isOpen);
            });
        });

        const sidebarBackdrop = document.getElementById('sidebar-backdrop');
        if (sidebarBackdrop) {
            sidebarBackdrop.addEventListener('click', () => setSidebarOpen(false));
        }
        window.addEventListener('resize', () => {
            if (window.innerWidth > 768) {
                setSidebarOpen(false);
            }
        });
        
        // 绑定模板列表视图事件
        document.getElementById('btn-new-template').addEventListener('click', newTemplate);
        document.getElementById('btn-import-template').addEventListener('click', () => {
            document.getElementById('import-template-input').click();
        });
        
        // 模板导入
        document.getElementById('import-template-input').addEventListener('change', async (e) => {
            const file = e.target.files[0];
            if (file) {
                try {
                    // 显示冲突处理选项
                    showModal({
                        title: '导入模板',
                        content: `
                            <p>如果存在同名模板，您希望如何处理？</p>
                            <div style="margin-top: 16px;">
                                <label style="display: flex; align-items: center; gap: 8px; margin-bottom: 12px; cursor: pointer;">
                                    <input type="radio" name="conflict" value="rename" checked>
                                    <span>重命名（保留两者）</span>
                                </label>
                                <label style="display: flex; align-items: center; gap: 8px; margin-bottom: 12px; cursor: pointer;">
                                    <input type="radio" name="conflict" value="overwrite">
                                    <span>覆盖（替换现有模板）</span>
                                </label>
                                <label style="display: flex; align-items: center; gap: 8px; cursor: pointer;">
                                    <input type="radio" name="conflict" value="skip">
                                    <span>跳过（不导入）</span>
                                </label>
                            </div>
                        `,
                        buttons: [
                            { text: '取消', class: 'btn-secondary' },
                            {
                                text: '导入',
                                class: 'btn-primary',
                                onClick: async () => {
                                    const strategy = document.querySelector('input[name="conflict"]:checked').value;
                                    try {
                                        const result = await Export.importTemplatePack(file, { conflictStrategy: strategy });
                                        showToast(result.message, 'success');
                                        loadTemplateList();
                                    } catch (error) {
                                        showToast(error.message, 'error');
                                    }
                                }
                            }
                        ]
                    });
                } catch (error) {
                    showToast(error.message, 'error');
                }
                e.target.value = '';
            }
        });
        
        // 绑定模板编辑视图事件
        document.getElementById('btn-back-from-edit').addEventListener('click', () => {
            switchView('templates');
        });
        document.getElementById('btn-save-template').addEventListener('click', saveTemplate);
        document.getElementById('btn-check-template').addEventListener('click', checkTemplate);
        document.getElementById('btn-add-field').addEventListener('click', () => {
            if (AppState.fieldConfigManager) {
                AppState.fieldConfigManager.addField();
            }
        });
        document.getElementById('btn-add-table').addEventListener('click', () => {
            if (AppState.tableConfigManager) {
                if (!AppState.tableConfigManager.canAddMore()) {
                    showToast('最多只能添加 3 张明细表', 'warning');
                    return;
                }
                AppState.tableConfigManager.addTable();
            }
        });
        
        // Word 文件上传
        const wordUploadArea = document.getElementById('word-upload-area');
        const wordFileInput = document.getElementById('word-file-input');
        
        wordUploadArea.addEventListener('click', () => {
            wordFileInput.click();
        });
        
        wordUploadArea.addEventListener('dragover', (e) => {
            e.preventDefault();
            wordUploadArea.classList.add('dragover');
        });
        
        wordUploadArea.addEventListener('dragleave', () => {
            wordUploadArea.classList.remove('dragover');
        });
        
        wordUploadArea.addEventListener('drop', async (e) => {
            e.preventDefault();
            wordUploadArea.classList.remove('dragover');
            const file = e.dataTransfer.files[0];
            if (file && file.name.endsWith('.docx')) {
                await handleWordFileSelect(file);
            } else {
                showToast('请上传 .docx 格式的文件', 'warning');
            }
        });
        
        wordFileInput.addEventListener('change', async (e) => {
            const file = e.target.files[0];
            if (file) {
                await handleWordFileSelect(file);
            }
        });
        
        document.querySelector('.btn-remove-file').addEventListener('click', (e) => {
            e.stopPropagation();
            if (AppState.editingTemplate) {
                AppState.editingTemplate.wordFile = null;
            }
            document.getElementById('word-file-info').style.display = 'none';
            document.querySelector('.file-upload-content').style.display = 'block';
            wordFileInput.value = '';
        });
        
        // 绑定表单填写视图事件
        document.getElementById('btn-back-from-form').addEventListener('click', () => {
            switchView('templates');
        });
        document.getElementById('btn-save-record').addEventListener('click', saveRecord);
        document.getElementById('btn-generate-word').addEventListener('click', generateWordDoc);
        document.getElementById('btn-load-history').addEventListener('click', loadFromHistory);
        
        // 绑定历史记录视图事件
        document.getElementById('btn-filter-history').addEventListener('click', () => {
            loadHistoryList();
        });
        document.getElementById('btn-clear-filter').addEventListener('click', () => {
            document.getElementById('history-template-filter').value = '';
            document.getElementById('history-date-start').value = '';
            document.getElementById('history-date-end').value = '';
            loadHistoryList();
        });
        document.getElementById('btn-export-records').addEventListener('click', exportRecords);
        document.getElementById('btn-import-records').addEventListener('click', () => {
            document.getElementById('import-records-input').click();
        });
        
        // 数据包导入
        document.getElementById('import-records-input').addEventListener('change', async (e) => {
            const file = e.target.files[0];
            if (file) {
                try {
                    const result = await Export.importRecordsPack(file);
                    showToast(result.message, 'success');
                    loadHistoryList();
                } catch (error) {
                    showToast(error.message, 'error');
                }
                e.target.value = '';
            }
        });
        
        // 设置工作区按钮
        document.getElementById('btn-workspace').addEventListener('click', async () => {
            try {
                await StorageAdapter.switchToFileStorage();
                showToast('✓ 工作区设置成功！数据将保存为本地文件', 'success');
                // 重新加载模板列表
                await loadTemplateList();
            } catch (error) {
                if (error.message.includes('不支持')) {
                    showToast('您的浏览器不支持文件存储，请使用最新版Chrome或Edge', 'error', 5000);
                } else if (!error.message.includes('取消')) {
                    showToast('设置失败: ' + error.message, 'error');
                }
            }
        });
        
        // 备份/恢复按钮
        document.getElementById('btn-backup').addEventListener('click', async () => {
            try {
                const fileName = await StorageAdapter.downloadBackupFile();
                showToast(`备份已保存: ${fileName}`, 'success');
            } catch (error) {
                showToast('备份失败: ' + error.message, 'error');
            }
        });
        
        document.getElementById('btn-restore').addEventListener('click', () => {
            const input = document.createElement('input');
            input.type = 'file';
            input.accept = '.json,.backup';
            input.onchange = async (e) => {
                const file = e.target.files[0];
                if (file) {
                    try {
                        const result = await StorageAdapter.importBackupFile(file);
                        showToast(`恢复成功: ${result.imported.templates} 个模板, ${result.imported.records} 条记录`, 'success');
                        loadTemplateList();
                    } catch (error) {
                        showToast('恢复失败: ' + error.message, 'error');
                    }
                }
            };
            input.click();
        });
        
        // 点击空白处关闭下拉菜单
        document.addEventListener('click', () => {
            document.querySelectorAll('.template-card-dropdown.show').forEach(d => {
                d.classList.remove('show');
            });
        });
        
        // 加载模板列表
        await loadTemplateList();
        
    } catch (error) {
        console.error('应用初始化失败:', error);
        showToast('应用初始化失败', 'error');
    }
}

/**
 * 处理 Word 文件选择
 * @param {File} file - 文件对象
 */
async function handleWordFileSelect(file) {
    try {
        const fileData = await Template.readFileAsBase64(file);
        
        // 验证文件
        const validation = Word.validateWordTemplate(fileData.data);
        if (!validation.valid) {
            // 显示详细错误信息
            let errorContent = '<p style="margin-bottom: 12px; font-weight: 500;">Word 模板存在以下问题：</p>';
            
            if (validation.errors && validation.errors.length > 0) {
                errorContent += '<div style="max-height: 300px; overflow-y: auto; background: #fff3cd; padding: 12px; border-radius: 6px; margin-bottom: 16px;">';
                errorContent += '<ul style="color: #856404; margin: 0; padding-left: 20px;">';
                validation.errors.forEach(err => {
                    errorContent += `<li style="margin-bottom: 8px; line-height: 1.5;">${escapeHtml(err)}</li>`;
                });
                errorContent += '</ul></div>';
                
                // 检查是否是循环标签问题
                const hasLoopError = validation.errors.some(e => 
                    e.includes('循环未开始') || e.includes('循环未闭合') || e.includes('unopened_loop') || e.includes('unclosed_loop')
                );
                
                if (hasLoopError) {
                    errorContent += '<div style="background: #e7f3ff; padding: 12px; border-radius: 6px; margin-bottom: 12px;">';
                    errorContent += '<p style="margin: 0 0 8px 0; font-weight: 500; color: #004085;">💡 循环标签使用说明：</p>';
                    errorContent += '<p style="margin: 0; color: #004085; font-size: 13px; line-height: 1.6;">明细表需要使用循环标签，格式为：</p>';
                    errorContent += '<div style="background: white; padding: 8px; border-radius: 4px; margin-top: 8px; font-family: monospace; font-size: 12px;">';
                    errorContent += '{#baseStationDetails}<br>';
                    errorContent += '{序号} {基站名称} {地址}<br>';
                    errorContent += '{/baseStationDetails}';
                    errorContent += '</div>';
                    errorContent += '<p style="margin: 8px 0 0 0; color: #004085; font-size: 12px;">确保开始标签 {#表名} 和结束标签 {/表名} 成对出现！</p>';
                    errorContent += '</div>';
                }
                
                errorContent += '<p style="margin-top: 12px; color: var(--color-text-secondary); font-size: 13px; font-weight: 500;">常见问题检查：</p>';
                errorContent += '<ul style="color: var(--color-text-secondary); font-size: 13px; margin: 8px 0 0 0; padding-left: 20px; line-height: 1.8;">';
                errorContent += '<li>检查是否有单独的结束标签 {/表名} 但没有对应的开始标签 {#表名}</li>';
                errorContent += '<li>检查是否有单独的开始标签 {#表名} 但没有对应的结束标签 {/表名}</li>';
                errorContent += '<li>检查标签名称是否完全一致（区分大小写）</li>';
                errorContent += '<li>检查是否误删除了某个标签</li>';
                errorContent += '</ul>';
            } else {
                errorContent += `<p style="color: var(--color-danger);">${escapeHtml(validation.error)}</p>`;
            }
            
            showModal({
                title: '❌ Word 模板验证失败',
                content: errorContent,
                buttons: [{ text: '我知道了', class: 'btn-primary' }]
            });
            
            return;
        }
        
        if (AppState.editingTemplate) {
            AppState.editingTemplate.wordFile = fileData;
        }
        
        document.getElementById('word-file-info').style.display = 'flex';
        document.querySelector('.file-upload-content').style.display = 'none';
        document.querySelector('.file-name').textContent = file.name;
        
        showToast('Word 模板上传成功', 'success');
        
    } catch (error) {
        showToast('读取文件失败', 'error');
    }
}

// 页面加载完成后初始化
document.addEventListener('DOMContentLoaded', initApp);
