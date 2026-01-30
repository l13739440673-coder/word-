/**
 * 动态表单模块
 * 提供根据模板配置动态生成表单的功能
 */

/**
 * 渲染动态表单
 * @param {HTMLElement} container - 容器元素
 * @param {Array} fields - 字段配置列表
 * @param {Object} data - 已有数据（可选）
 * @returns {Object} - 表单数据收集器
 */
function renderDynamicForm(container, fields, data = {}) {
    container.innerHTML = '';
    
    // 按 order 排序
    const sortedFields = [...fields].sort((a, b) => (a.order || 0) - (b.order || 0));
    
    sortedFields.forEach(field => {
        const formGroup = document.createElement('div');
        formGroup.className = 'form-group';
        
        // 多行文本占满整行
        if (field.type === 'textarea') {
            formGroup.classList.add('full-width');
        }
        
        // 标签
        const label = document.createElement('label');
        label.setAttribute('for', `field-${field.key}`);
        label.innerHTML = field.label;
        if (field.required) {
            label.innerHTML += ' <span class="required">*</span>';
        }
        formGroup.appendChild(label);
        
        // 输入框
        let input;
        if (field.type === 'textarea') {
            input = document.createElement('textarea');
            input.rows = 3;
        } else {
            input = document.createElement('input');
            input.type = field.type === 'number' ? 'number' : 'text';
            
            if (field.type === 'number') {
                input.step = 'any';
            }
        }
        
        input.id = `field-${field.key}`;
        input.name = field.key;
        input.placeholder = `请输入${field.label}`;
        
        if (field.required) {
            input.required = true;
        }
        
        // 设置已有值
        if (data[field.key] !== undefined) {
            input.value = data[field.key];
        }
        
        formGroup.appendChild(input);
        container.appendChild(formGroup);
    });
    
    // 返回数据收集器
    return {
        getData: () => {
            const formData = {};
            sortedFields.forEach(field => {
                const input = container.querySelector(`[name="${field.key}"]`);
                if (input) {
                    let value = input.value;
                    if (field.type === 'number' && value) {
                        value = parseFloat(value);
                    }
                    formData[field.key] = value;
                }
            });
            return formData;
        },
        setData: (newData) => {
            sortedFields.forEach(field => {
                const input = container.querySelector(`[name="${field.key}"]`);
                if (input && newData[field.key] !== undefined) {
                    input.value = newData[field.key];
                }
            });
        },
        validate: () => {
            const errors = [];
            sortedFields.forEach(field => {
                if (field.required) {
                    const input = container.querySelector(`[name="${field.key}"]`);
                    if (!input || !input.value.trim()) {
                        errors.push(`"${field.label}" 不能为空`);
                    }
                }
            });
            return { valid: errors.length === 0, errors };
        },
        clear: () => {
            sortedFields.forEach(field => {
                const input = container.querySelector(`[name="${field.key}"]`);
                if (input) {
                    input.value = '';
                }
            });
        }
    };
}

/**
 * 渲染明细表
 * @param {HTMLElement} container - 容器元素
 * @param {Object} tableConfig - 明细表配置
 * @param {Array} data - 已有数据（可选）
 * @returns {Object} - 明细表数据收集器
 */
function renderDetailTable(container, tableConfig, data = []) {
    const section = document.createElement('section');
    section.className = 'detail-table-section';
    section.dataset.tableName = tableConfig.name;
    
    // 表头
    const header = document.createElement('div');
    header.className = 'detail-table-header';
    header.innerHTML = `
        <h3>${tableConfig.name}</h3>
        <div class="detail-table-actions">
            <button class="btn btn-small btn-secondary btn-paste-data" type="button" title="从 Excel 粘贴数据">
                <svg width="14" height="14" viewBox="0 0 14 14" fill="currentColor">
                    <rect x="4" y="2" width="8" height="10" rx="1" stroke="currentColor" fill="none"/>
                    <path d="M6 2V1h4v1M4 5h6M4 7h6M4 9h6" stroke="currentColor" fill="none"/>
                </svg>
                粘贴数据
            </button>
            <button class="btn btn-small btn-secondary btn-add-row" type="button">
                <svg width="14" height="14" viewBox="0 0 14 14" fill="currentColor">
                    <path d="M7 1v12M1 7h12" stroke="currentColor" stroke-width="2" fill="none"/>
                </svg>
                添加行
            </button>
        </div>
    `;
    section.appendChild(header);
    
    // 表格容器
    const tableWrapper = document.createElement('div');
    tableWrapper.className = 'detail-table-wrapper';
    
    const table = document.createElement('table');
    table.className = 'detail-table';
    
    // 表头行
    const thead = document.createElement('thead');
    const headerRow = document.createElement('tr');
    
    tableConfig.columns.forEach(col => {
        const th = document.createElement('th');
        th.textContent = col.label;
        headerRow.appendChild(th);
    });
    
    // 操作列
    const actionTh = document.createElement('th');
    actionTh.style.width = '80px';
    actionTh.textContent = '操作';
    headerRow.appendChild(actionTh);
    
    thead.appendChild(headerRow);
    table.appendChild(thead);
    
    // 表体
    const tbody = document.createElement('tbody');
    table.appendChild(tbody);
    
    tableWrapper.appendChild(table);
    section.appendChild(tableWrapper);
    container.appendChild(section);
    
    // 行管理
    let rows = [];
    
    /**
     * 创建一行
     * @param {Object} rowData - 行数据
     * @returns {HTMLElement}
     */
    function createRow(rowData = {}) {
        const tr = document.createElement('tr');
        const rowId = Storage.generateId();
        tr.dataset.rowId = rowId;
        
        tableConfig.columns.forEach(col => {
            const td = document.createElement('td');
            const input = document.createElement('input');
            input.type = col.type === 'number' ? 'number' : 'text';
            input.name = col.key;
            input.placeholder = col.label;
            
            if (col.type === 'number') {
                input.step = 'any';
            }
            
            if (rowData[col.key] !== undefined) {
                input.value = rowData[col.key];
            }
            
            td.appendChild(input);
            tr.appendChild(td);
        });
        
        // 操作列
        const actionTd = document.createElement('td');
        actionTd.className = 'row-actions';
        actionTd.innerHTML = `
            <button class="btn-row-action copy" title="复制行" type="button">
                <svg width="14" height="14" viewBox="0 0 14 14" fill="currentColor">
                    <rect x="4" y="4" width="8" height="8" rx="1" stroke="currentColor" fill="none"/>
                    <path d="M2 10V2h8" stroke="currentColor" fill="none"/>
                </svg>
            </button>
            <button class="btn-row-action delete" title="删除行" type="button">
                <svg width="14" height="14" viewBox="0 0 14 14" fill="currentColor">
                    <path d="M3 4h8M5 4V3a1 1 0 011-1h2a1 1 0 011 1v1M6 7v3M8 7v3M4 4l.5 7a1 1 0 001 1h3a1 1 0 001-1l.5-7" stroke="currentColor" fill="none"/>
                </svg>
            </button>
        `;
        tr.appendChild(actionTd);
        
        return tr;
    }
    
    /**
     * 添加行
     * @param {Object} rowData - 行数据
     */
    function addRow(rowData = {}) {
        const tr = createRow(rowData);
        tbody.appendChild(tr);
        rows.push(tr);
        
        // 绑定行操作事件
        const copyBtn = tr.querySelector('.btn-row-action.copy');
        const deleteBtn = tr.querySelector('.btn-row-action.delete');
        
        copyBtn.addEventListener('click', () => {
            const data = getRowData(tr);
            addRow(data);
        });
        
        deleteBtn.addEventListener('click', () => {
            if (rows.length > 1) {
                const index = rows.indexOf(tr);
                if (index > -1) {
                    rows.splice(index, 1);
                }
                tr.remove();
            } else {
                // 如果只有一行，清空数据而不是删除
                clearRowData(tr);
            }
        });
    }
    
    /**
     * 获取行数据
     * @param {HTMLElement} tr - 行元素
     * @returns {Object}
     */
    function getRowData(tr) {
        const rowData = {};
        tableConfig.columns.forEach(col => {
            const input = tr.querySelector(`input[name="${col.key}"]`);
            if (input) {
                let value = input.value;
                if (col.type === 'number' && value) {
                    value = parseFloat(value);
                }
                rowData[col.key] = value;
            }
        });
        return rowData;
    }
    
    /**
     * 清空行数据
     * @param {HTMLElement} tr - 行元素
     */
    function clearRowData(tr) {
        tableConfig.columns.forEach(col => {
            const input = tr.querySelector(`input[name="${col.key}"]`);
            if (input) {
                input.value = '';
            }
        });
    }
    
    // 绑定添加行按钮事件
    const addRowBtn = section.querySelector('.btn-add-row');
    addRowBtn.addEventListener('click', () => addRow());
    
    // 绑定粘贴数据按钮事件
    const pasteBtn = section.querySelector('.btn-paste-data');
    pasteBtn.addEventListener('click', () => {
        showPasteDialog();
    });
    
    /**
     * 显示粘贴对话框
     */
    function showPasteDialog() {
        const modalContent = document.createElement('div');
        modalContent.innerHTML = `
            <p style="margin-bottom: 12px;">请从 Excel 或其他表格软件中复制数据，然后粘贴到下方文本框：</p>
            <textarea id="paste-textarea" style="width: 100%; height: 200px; font-family: monospace; font-size: 12px; padding: 8px; border: 1px solid var(--color-border); border-radius: 4px;" placeholder="在这里粘贴表格数据...&#10;&#10;支持格式：&#10;- Excel 复制粘贴（Tab 分隔）&#10;- 纯文本（Tab 或逗号分隔）"></textarea>
            <p style="margin-top: 12px; color: var(--color-text-secondary); font-size: 13px;">
                <strong>提示：</strong>
            </p>
            <ul style="margin: 8px 0 0 20px; padding: 0; color: var(--color-text-secondary); font-size: 13px; line-height: 1.6;">
                <li>从 Excel 选中并复制单元格（Ctrl+C）</li>
                <li>数据按列顺序排列：${tableConfig.columns.map(c => c.label).join(' | ')}</li>
                <li>可以包含或不包含表头行</li>
                <li>粘贴后会替换现有数据</li>
            </ul>
        `;
        
        window.showModal({
            title: '粘贴表格数据',
            content: modalContent,
            buttons: [
                { text: '取消', class: 'btn-secondary' },
                {
                    text: '确定',
                    class: 'btn-primary',
                    onClick: () => {
                        const textarea = document.getElementById('paste-textarea');
                        const text = textarea.value.trim();
                        if (text) {
                            try {
                                parseAndFillData(text);
                                window.showToast('数据粘贴成功', 'success');
                            } catch (error) {
                                window.showToast('数据解析失败: ' + error.message, 'error');
                            }
                        }
                    }
                }
            ]
        });
        
        // 聚焦文本框
        setTimeout(() => {
            const textarea = document.getElementById('paste-textarea');
            if (textarea) textarea.focus();
        }, 100);
    }
    
    /**
     * 解析并填充粘贴的数据
     * @param {string} text - 粘贴的文本
     */
    function parseAndFillData(text) {
        // 按行分割
        const lines = text.split(/\r?\n/).filter(line => line.trim());
        
        if (lines.length === 0) {
            throw new Error('没有有效数据');
        }
        
        // 检测分隔符（Tab 或逗号）
        const firstLine = lines[0];
        const delimiter = firstLine.includes('\t') ? '\t' : ',';
        
        // 解析每一行
        const parsedData = [];
        let startIndex = 0;
        
        // 检查第一行是否为表头
        const firstRowCells = lines[0].split(delimiter).map(s => s.trim());
        const hasHeader = tableConfig.columns.some(col => 
            firstRowCells.some(cell => 
                cell.toLowerCase() === col.label.toLowerCase() || 
                cell.toLowerCase() === col.key.toLowerCase()
            )
        );
        
        if (hasHeader) {
            startIndex = 1; // 跳过表头
        }
        
        for (let i = startIndex; i < lines.length; i++) {
            const cells = lines[i].split(delimiter).map(s => s.trim());
            
            if (cells.length === 0 || cells.every(c => !c)) {
                continue; // 跳过空行
            }
            
            const rowData = {};
            tableConfig.columns.forEach((col, index) => {
                if (index < cells.length) {
                    let value = cells[index];
                    // 数字类型转换
                    if (col.type === 'number' && value) {
                        const num = parseFloat(value);
                        if (!isNaN(num)) {
                            value = num;
                        }
                    }
                    rowData[col.key] = value;
                }
            });
            
            parsedData.push(rowData);
        }
        
        if (parsedData.length === 0) {
            throw new Error('没有解析到有效数据');
        }
        
        // 清空现有数据并填充新数据
        tbody.innerHTML = '';
        rows = [];
        
        parsedData.forEach(rowData => addRow(rowData));
    }
    
    // 初始化数据
    if (data && data.length > 0) {
        data.forEach(rowData => addRow(rowData));
    } else {
        // 默认添加一行
        addRow();
    }
    
    // 返回数据收集器
    return {
        getData: () => {
            return rows.map(tr => getRowData(tr));
        },
        setData: (newData) => {
            // 清空现有行
            tbody.innerHTML = '';
            rows = [];
            
            if (newData && newData.length > 0) {
                newData.forEach(rowData => addRow(rowData));
            } else {
                addRow();
            }
        },
        addRow,
        clear: () => {
            tbody.innerHTML = '';
            rows = [];
            addRow();
        }
    };
}

/**
 * 渲染所有明细表
 * @param {HTMLElement} container - 容器元素
 * @param {Array} tableConfigs - 明细表配置列表
 * @param {Object} data - 已有数据（按表名索引）
 * @returns {Object} - 所有明细表的数据收集器
 */
function renderDetailTables(container, tableConfigs, data = {}) {
    container.innerHTML = '';
    
    const tableCollectors = {};
    
    if (tableConfigs && tableConfigs.length > 0) {
        tableConfigs.forEach(tableConfig => {
            const tableData = data[tableConfig.name] || [];
            tableCollectors[tableConfig.name] = renderDetailTable(container, tableConfig, tableData);
        });
    }
    
    return {
        getData: () => {
            const allData = {};
            Object.keys(tableCollectors).forEach(name => {
                allData[name] = tableCollectors[name].getData();
            });
            return allData;
        },
        setData: (newData) => {
            Object.keys(tableCollectors).forEach(name => {
                if (newData[name]) {
                    tableCollectors[name].setData(newData[name]);
                }
            });
        },
        getTable: (name) => tableCollectors[name],
        clear: () => {
            Object.keys(tableCollectors).forEach(name => {
                tableCollectors[name].clear();
            });
        }
    };
}

/**
 * 渲染字段配置列表（用于模板编辑）
 * @param {HTMLElement} container - 容器元素
 * @param {Array} fields - 字段列表
 * @param {Function} onFieldsChange - 字段变更回调
 * @returns {Object} - 字段管理器
 */
function renderFieldConfigList(container, fields = [], onFieldsChange = null) {
    container.innerHTML = '';
    
    let currentFields = [...fields];
    
    function renderFields() {
        container.innerHTML = '';
        
        if (currentFields.length === 0) {
            container.innerHTML = '<div class="empty-hint"><p>暂无字段，点击"添加字段"创建</p></div>';
            return;
        }
        
        currentFields.forEach((field, index) => {
            const item = document.createElement('div');
            item.className = 'field-item';
            item.dataset.fieldId = field.id;
            
            const isFirst = index === 0;
            const isLast = index === currentFields.length - 1;
            
            item.innerHTML = `
                <div class="field-drag-handle" title="拖拽排序">
                    <svg width="16" height="16" viewBox="0 0 16 16" fill="currentColor">
                        <circle cx="5" cy="4" r="1.5"/>
                        <circle cx="11" cy="4" r="1.5"/>
                        <circle cx="5" cy="8" r="1.5"/>
                        <circle cx="11" cy="8" r="1.5"/>
                        <circle cx="5" cy="12" r="1.5"/>
                        <circle cx="11" cy="12" r="1.5"/>
                    </svg>
                </div>
                <div class="field-move-buttons">
                    <button class="btn-field-move up" title="上移" type="button" ${isFirst ? 'disabled' : ''}>
                        <svg width="12" height="12" viewBox="0 0 12 12" fill="currentColor">
                            <path d="M6 3l4 4H2l4-4z"/>
                        </svg>
                    </button>
                    <button class="btn-field-move down" title="下移" type="button" ${isLast ? 'disabled' : ''}>
                        <svg width="12" height="12" viewBox="0 0 12 12" fill="currentColor">
                            <path d="M6 9l4-4H2l4 4z"/>
                        </svg>
                    </button>
                </div>
                <div class="field-inputs">
                    <input type="text" name="key" placeholder="占位符 key" value="${field.key || ''}" />
                    <input type="text" name="label" placeholder="显示名称" value="${field.label || ''}" />
                    <select name="type">
                        ${Template.getFieldTypeOptions().map(opt => 
                            `<option value="${opt.value}" ${field.type === opt.value ? 'selected' : ''}>${opt.label}</option>`
                        ).join('')}
                    </select>
                    <div class="checkbox-wrapper">
                        <input type="checkbox" name="required" id="required-${field.id}" ${field.required ? 'checked' : ''} />
                        <label for="required-${field.id}">必填</label>
                    </div>
                </div>
                <button class="field-remove" title="删除字段" type="button">
                    <svg width="16" height="16" viewBox="0 0 16 16" fill="currentColor">
                        <path d="M4 4l8 8M12 4l-8 8" stroke="currentColor" stroke-width="2"/>
                    </svg>
                </button>
            `;
            
            container.appendChild(item);
            
            // 绑定输入事件
            const inputs = item.querySelectorAll('input, select');
            inputs.forEach(input => {
                input.addEventListener('change', () => {
                    updateFieldFromInputs(field.id, item);
                });
                input.addEventListener('input', () => {
                    updateFieldFromInputs(field.id, item);
                });
            });
            
            // 绑定删除事件
            const removeBtn = item.querySelector('.field-remove');
            removeBtn.addEventListener('click', () => {
                removeField(field.id);
            });
            
            // 绑定上移事件
            const upBtn = item.querySelector('.btn-field-move.up');
            upBtn.addEventListener('click', () => {
                moveField(field.id, 'up');
            });
            
            // 绑定下移事件
            const downBtn = item.querySelector('.btn-field-move.down');
            downBtn.addEventListener('click', () => {
                moveField(field.id, 'down');
            });
        });
    }
    
    function updateFieldFromInputs(fieldId, itemElement) {
        const field = currentFields.find(f => f.id === fieldId);
        if (field) {
            field.key = itemElement.querySelector('input[name="key"]').value.trim();
            field.label = itemElement.querySelector('input[name="label"]').value.trim();
            field.type = itemElement.querySelector('select[name="type"]').value;
            field.required = itemElement.querySelector('input[name="required"]').checked;
            
            if (onFieldsChange) {
                onFieldsChange(currentFields);
            }
        }
    }
    
    function addField() {
        const newField = Template.createEmptyField();
        newField.order = currentFields.length;
        currentFields.push(newField);
        renderFields();
        
        if (onFieldsChange) {
            onFieldsChange(currentFields);
        }
    }
    
    function removeField(fieldId) {
        currentFields = currentFields.filter(f => f.id !== fieldId);
        // 更新 order
        currentFields.forEach((f, i) => f.order = i);
        renderFields();
        
        if (onFieldsChange) {
            onFieldsChange(currentFields);
        }
    }
    
    function moveField(fieldId, direction) {
        const index = currentFields.findIndex(f => f.id === fieldId);
        if (index === -1) return;
        
        if (direction === 'up' && index > 0) {
            // 上移：与上一个交换位置
            [currentFields[index], currentFields[index - 1]] = [currentFields[index - 1], currentFields[index]];
        } else if (direction === 'down' && index < currentFields.length - 1) {
            // 下移：与下一个交换位置
            [currentFields[index], currentFields[index + 1]] = [currentFields[index + 1], currentFields[index]];
        } else {
            return; // 不需要移动
        }
        
        // 更新 order
        currentFields.forEach((f, i) => f.order = i);
        renderFields();
        
        if (onFieldsChange) {
            onFieldsChange(currentFields);
        }
    }
    
    // 初始渲染
    renderFields();
    
    return {
        getFields: () => currentFields,
        setFields: (newFields) => {
            currentFields = [...newFields];
            renderFields();
        },
        addField,
        removeField,
        moveField,
        refresh: renderFields
    };
}

/**
 * 渲染明细表配置列表（用于模板编辑）
 * @param {HTMLElement} container - 容器元素
 * @param {Array} tables - 明细表列表
 * @param {Function} onTablesChange - 明细表变更回调
 * @returns {Object} - 明细表管理器
 */
function renderTableConfigList(container, tables = [], onTablesChange = null) {
    container.innerHTML = '';
    
    let currentTables = [...tables];
    
    function renderTables() {
        container.innerHTML = '';
        
        if (currentTables.length === 0) {
            container.innerHTML = '<div class="empty-hint" id="empty-tables"><p>暂无明细表，点击"添加明细表"创建（最多3张）</p></div>';
            return;
        }
        
        currentTables.forEach((table, tableIndex) => {
            const item = document.createElement('div');
            item.className = 'table-config-item';
            item.dataset.tableId = table.id;
            
            // 表头
            const header = document.createElement('div');
            header.className = 'table-config-header';
            header.innerHTML = `
                <input type="text" name="tableName" placeholder="明细表名称（英文）" value="${table.name || ''}" />
                <div style="display: flex; gap: 8px;">
                    <button class="btn btn-small btn-secondary btn-add-column" type="button">添加列</button>
                    <button class="btn btn-small btn-text text-danger btn-remove-table" type="button">删除表</button>
                </div>
            `;
            item.appendChild(header);
            
            // 表体（列配置）
            const body = document.createElement('div');
            body.className = 'table-config-body';

            const columnsContainer = document.createElement('div');
            columnsContainer.className = 'table-columns';
            
            if (table.columns && table.columns.length > 0) {
                table.columns.forEach((col, colIndex) => {
                    const colItem = createColumnItem(table.id, col);
                    columnsContainer.appendChild(colItem);
                });
            } else {
                columnsContainer.innerHTML = '<div class="empty-hint"><p>暂无列，点击"添加列"</p></div>';
            }
            
            body.appendChild(columnsContainer);
            item.appendChild(body);
            container.appendChild(item);
            
            // 绑定表名输入事件
            const nameInput = header.querySelector('input[name="tableName"]');
            nameInput.addEventListener('change', () => {
                table.name = nameInput.value.trim();
                if (onTablesChange) {
                    onTablesChange(currentTables);
                }
            });
            
            // 绑定添加列事件
            const addColBtn = header.querySelector('.btn-add-column');
            addColBtn.addEventListener('click', () => {
                addColumn(table.id);
            });

            // 绑定删除表事件
            const removeTableBtn = header.querySelector('.btn-remove-table');
            removeTableBtn.addEventListener('click', () => {
                removeTable(table.id);
            });
        });
    }
    
    function createColumnItem(tableId, column) {
        const colItem = document.createElement('div');
        colItem.className = 'table-column-item';
        colItem.dataset.columnId = column.id;
        
        colItem.innerHTML = `
            <input type="text" name="colKey" placeholder="列占位符" value="${column.key || ''}" />
            <input type="text" name="colLabel" placeholder="列显示名称" value="${column.label || ''}" />
            <select name="colType">
                ${Template.getFieldTypeOptions().map(opt => 
                    `<option value="${opt.value}" ${column.type === opt.value ? 'selected' : ''}>${opt.label}</option>`
                ).join('')}
            </select>
            <button class="field-remove" title="删除列" type="button">
                <svg width="16" height="16" viewBox="0 0 16 16" fill="currentColor">
                    <path d="M4 4l8 8M12 4l-8 8" stroke="currentColor" stroke-width="2"/>
                </svg>
            </button>
        `;
        
        // 绑定输入事件
        const inputs = colItem.querySelectorAll('input, select');
        inputs.forEach(input => {
            const updateColumn = () => {
                const table = currentTables.find(t => t.id === tableId);
                const col = table?.columns?.find(c => c.id === column.id);
                if (col) {
                    col.key = colItem.querySelector('input[name="colKey"]').value.trim();
                    col.label = colItem.querySelector('input[name="colLabel"]').value.trim();
                    col.type = colItem.querySelector('select[name="colType"]').value;
                    
                    if (onTablesChange) {
                        onTablesChange(currentTables);
                    }
                }
            };
            input.addEventListener('change', updateColumn);
            input.addEventListener('input', updateColumn);
        });
        
        // 绑定删除列事件
        const removeBtn = colItem.querySelector('.field-remove');
        removeBtn.addEventListener('click', () => {
            removeColumn(tableId, column.id);
        });
        
        return colItem;
    }
    
    function addTable() {
        if (currentTables.length >= 3) {
            return false;
        }
        
        const newTable = Template.createEmptyDetailTable();
        currentTables.push(newTable);
        renderTables();
        
        if (onTablesChange) {
            onTablesChange(currentTables);
        }
        return true;
    }
    
    function removeTable(tableId) {
        currentTables = currentTables.filter(t => t.id !== tableId);
        renderTables();
        
        if (onTablesChange) {
            onTablesChange(currentTables);
        }
    }
    
    function addColumn(tableId) {
        const table = currentTables.find(t => t.id === tableId);
        if (table) {
            if (!table.columns) {
                table.columns = [];
            }
            table.columns.push(Template.createEmptyColumn());
            renderTables();
            
            if (onTablesChange) {
                onTablesChange(currentTables);
            }
        }
    }
    
    function removeColumn(tableId, columnId) {
        const table = currentTables.find(t => t.id === tableId);
        if (table && table.columns) {
            table.columns = table.columns.filter(c => c.id !== columnId);
            renderTables();
            
            if (onTablesChange) {
                onTablesChange(currentTables);
            }
        }
    }
    
    // 初始渲染
    renderTables();
    
    return {
        getTables: () => currentTables,
        setTables: (newTables) => {
            currentTables = [...newTables];
            renderTables();
        },
        addTable,
        removeTable,
        canAddMore: () => currentTables.length < 3,
        refresh: renderTables
    };
}

/**
 * 更新标题字段下拉选项
 * @param {HTMLSelectElement} selectElement - 下拉选择框
 * @param {Array} fields - 字段列表
 * @param {string} selectedKey - 当前选中的 key
 */
function updateTitleFieldSelect(selectElement, fields, selectedKey = '') {
    selectElement.innerHTML = '';
    
    if (!fields || fields.length === 0) {
        selectElement.innerHTML = '<option value="">请先添加字段</option>';
        return;
    }
    
    const defaultOption = document.createElement('option');
    defaultOption.value = '';
    defaultOption.textContent = '请选择标题字段';
    selectElement.appendChild(defaultOption);
    
    fields.forEach(field => {
        if (field.key && field.key.trim()) {
            const option = document.createElement('option');
            option.value = field.key;
            option.textContent = `${field.label || field.key} (${field.key})`;
            if (field.key === selectedKey) {
                option.selected = true;
            }
            selectElement.appendChild(option);
        }
    });
}

// 导出 Form 对象
window.Form = {
    renderDynamicForm,
    renderDetailTable,
    renderDetailTables,
    renderFieldConfigList,
    renderTableConfigList,
    updateTitleFieldSelect
};
