/**
 * 模板管理模块
 * 提供模板的创建、编辑、删除、验证等功能
 */

// 字段类型定义
const FIELD_TYPES = {
    TEXT: { value: 'text', label: '文本' },
    NUMBER: { value: 'number', label: '数字' }
};

// 获取字段类型列表
function getFieldTypeOptions() {
    return Object.values(FIELD_TYPES);
}

/**
 * 创建空白模板对象
 * @returns {Object}
 */
function createEmptyTemplate() {
    return {
        id: null,
        name: '',
        description: '',
        titleFieldKey: '',
        wordFile: null,
        fields: [],
        detailTables: [],
        createdAt: null,
        updatedAt: null
    };
}

/**
 * 创建空白字段对象
 * @returns {Object}
 */
function createEmptyField() {
    return {
        id: StorageAdapter.generateId(),
        key: '',
        label: '',
        type: 'text',
        required: false,
        order: 0
    };
}

/**
 * 创建空白明细表对象
 * @returns {Object}
 */
function createEmptyDetailTable() {
    return {
        id: Storage.generateId(),
        name: '',
        columns: []
    };
}

/**
 * 创建空白明细表列对象
 * @returns {Object}
 */
function createEmptyColumn() {
    return {
        id: StorageAdapter.generateId(),
        key: '',
        label: '',
        type: 'text'
    };
}

/**
 * 验证模板数据
 * @param {Object} template - 模板数据
 * @returns {Object} - { valid: boolean, errors: string[] }
 */
function validateTemplate(template) {
    const errors = [];

    // 验证模板名称
    if (!template.name || !template.name.trim()) {
        errors.push('模板名称不能为空');
    }

    // 验证 Word 文件
    if (!template.wordFile || !template.wordFile.data) {
        errors.push('请上传 Word 模板文件');
    }

    // 验证字段（允许纯明细表模板）
    const hasFields = template.fields && template.fields.length > 0;
    if (hasFields) {
        // 检查字段 key 是否有效
        const fieldKeys = new Set();
        template.fields.forEach((field, index) => {
            if (!field.key || !field.key.trim()) {
                errors.push(`字段 ${index + 1} 的占位符 key 不能为空`);
            } else if (!/^[a-zA-Z_][a-zA-Z0-9_]*$/.test(field.key)) {
                errors.push(`字段 "${field.key}" 的占位符格式不正确（只能包含字母、数字和下划线，且不能以数字开头）`);
            } else if (fieldKeys.has(field.key)) {
                errors.push(`字段占位符 "${field.key}" 重复`);
            } else {
                fieldKeys.add(field.key);
            }

            if (!field.label || !field.label.trim()) {
                errors.push(`字段 ${index + 1} 的显示名称不能为空`);
            }
        });
    }

    // 验证明细表（如果有的话）
    const hasDetailTables = template.detailTables && template.detailTables.length > 0;
    if (hasDetailTables) {
        if (template.detailTables.length > 3) {
            errors.push('明细表最多只能添加 3 张');
        }

        const tableNames = new Set();
        template.detailTables.forEach((table, tableIndex) => {
            if (!table.name || !table.name.trim()) {
                errors.push(`明细表 ${tableIndex + 1} 的名称不能为空`);
            } else if (!/^[a-zA-Z_][a-zA-Z0-9_]*$/.test(table.name)) {
                errors.push(`明细表 "${table.name}" 的名称格式不正确`);
            } else if (tableNames.has(table.name)) {
                errors.push(`明细表名称 "${table.name}" 重复`);
            } else {
                tableNames.add(table.name);
            }

            if (!table.columns || table.columns.length === 0) {
                errors.push(`明细表 "${table.name || tableIndex + 1}" 请至少添加一列`);
            } else {
                const columnKeys = new Set();
                table.columns.forEach((col, colIndex) => {
                    if (!col.key || !col.key.trim()) {
                        errors.push(`明细表 "${table.name}" 的列 ${colIndex + 1} 占位符不能为空`);
                    } else if (columnKeys.has(col.key)) {
                        errors.push(`明细表 "${table.name}" 的列占位符 "${col.key}" 重复`);
                    } else {
                        columnKeys.add(col.key);
                    }

                    if (!col.label || !col.label.trim()) {
                        errors.push(`明细表 "${table.name}" 的列 ${colIndex + 1} 显示名称不能为空`);
                    }
                });

            }
        });
    }

    if (!hasFields && !hasDetailTables) {
        errors.push('请至少添加一个字段或明细表');
    }

    return {
        valid: errors.length === 0,
        errors
    };
}

/**
 * 从 Word 文件中提取占位符
 * @param {ArrayBuffer} fileData - Word 文件数据
 * @returns {Set<string>} - 占位符集合
 */
function extractPlaceholdersFromWord(fileData) {
    const placeholders = new Set();
    
    try {
        const zip = new window.PizZip(fileData);
        const doc = new window.docxtemplater(zip, {
            paragraphLoop: true,
            linebreaks: true,
            delimiters: { start: '{', end: '}' }
        });
        
        // 获取文档中的所有标签
        const tags = doc.getFullText();
        
        // 使用正则表达式提取占位符
        const regex = /\{([^{}]+)\}/g;
        let match;
        
        while ((match = regex.exec(tags)) !== null) {
            const tag = match[1].trim();
            // 过滤掉循环标签（以 # 或 / 开头的）
            if (!tag.startsWith('#') && !tag.startsWith('/')) {
                placeholders.add(tag);
            }
        }
    } catch (error) {
        console.error('提取占位符失败:', error);
    }
    
    return placeholders;
}

/**
 * 模板自检 - 检查配置的字段与 Word 模板中的占位符是否匹配
 * @param {Object} template - 模板数据
 * @returns {Object} - { matched: string[], missingInWord: string[], missingInConfig: string[] }
 */
async function checkTemplateConsistency(template) {
    const result = {
        matched: [],           // 配置与 Word 都有的占位符
        missingInWord: [],     // 配置有但 Word 没有
        missingInConfig: [],   // Word 有但配置没有
        valid: true
    };

    if (!template.wordFile || !template.wordFile.data) {
        result.valid = false;
        result.error = '请先上传 Word 模板文件';
        return result;
    }

    try {
        // 将 base64 转换为 ArrayBuffer
        const base64 = template.wordFile.data.split(',')[1];
        const binaryString = atob(base64);
        const bytes = new Uint8Array(binaryString.length);
        for (let i = 0; i < binaryString.length; i++) {
            bytes[i] = binaryString.charCodeAt(i);
        }
        
        // 提取 Word 中的占位符
        const wordPlaceholders = extractPlaceholdersFromWord(bytes.buffer);
        
        // 获取配置的字段 key
        const configKeys = new Set();
        template.fields.forEach(f => configKeys.add(f.key));
        
        // 获取明细表列的 key
        if (template.detailTables) {
            template.detailTables.forEach(table => {
                if (table.columns) {
                    table.columns.forEach(col => configKeys.add(col.key));
                }
            });
        }
        
        // 比较
        configKeys.forEach(key => {
            if (wordPlaceholders.has(key)) {
                result.matched.push(key);
            } else {
                result.missingInWord.push(key);
            }
        });
        
        wordPlaceholders.forEach(placeholder => {
            if (!configKeys.has(placeholder)) {
                result.missingInConfig.push(placeholder);
            }
        });
        
        result.valid = result.missingInWord.length === 0 && result.missingInConfig.length === 0;
        
    } catch (error) {
        console.error('模板自检失败:', error);
        result.valid = false;
        result.error = '模板自检失败，请检查 Word 文件格式是否正确';
    }

    return result;
}

/**
 * 读取文件为 Base64
 * @param {File} file - 文件对象
 * @returns {Promise<Object>} - { name, data }
 */
function readFileAsBase64(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        
        reader.onload = () => {
            resolve({
                name: file.name,
                data: reader.result
            });
        };
        
        reader.onerror = () => {
            reject(new Error('读取文件失败'));
        };
        
        reader.readAsDataURL(file);
    });
}

/**
 * 将 Base64 数据转换为 Blob
 * @param {string} base64Data - Base64 数据
 * @param {string} mimeType - MIME 类型
 * @returns {Blob}
 */
function base64ToBlob(base64Data, mimeType = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document') {
    const base64 = base64Data.split(',')[1] || base64Data;
    const binaryString = atob(base64);
    const bytes = new Uint8Array(binaryString.length);
    
    for (let i = 0; i < binaryString.length; i++) {
        bytes[i] = binaryString.charCodeAt(i);
    }
    
    return new Blob([bytes], { type: mimeType });
}

/**
 * 获取模板统计信息
 * @param {string} templateId - 模板 ID
 * @returns {Promise<Object>}
 */
async function getTemplateStats(templateId) {
    const records = await StorageAdapter.getRecordsByTemplateId(templateId);
    
    let lastUsedAt = null;
    if (records.length > 0) {
        lastUsedAt = records.reduce((latest, record) => {
            const recordDate = new Date(record.updatedAt);
            return !latest || recordDate > latest ? recordDate : latest;
        }, null);
    }
    
    return {
        recordCount: records.length,
        lastUsedAt: lastUsedAt ? lastUsedAt.toISOString() : null
    };
}

/**
 * 复制模板（创建副本）
 * @param {string} templateId - 原模板 ID
 * @returns {Promise<Object>}
 */
async function duplicateTemplate(templateId) {
    const original = await StorageAdapter.getTemplateById(templateId);
    
    if (!original) {
        throw new Error('模板不存在');
    }
    
    const copy = {
        ...original,
        id: null,
        name: `${original.name} (副本)`,
        createdAt: null,
        updatedAt: null
    };
    
    // 重新生成字段 ID
    copy.fields = copy.fields.map(f => ({
        ...f,
        id: StorageAdapter.generateId()
    }));
    
    // 重新生成明细表 ID
    if (copy.detailTables) {
        copy.detailTables = copy.detailTables.map(t => ({
            ...t,
            id: StorageAdapter.generateId(),
            columns: t.columns.map(c => ({
                ...c,
                id: StorageAdapter.generateId()
            }))
        }));
    }
    
    return await StorageAdapter.saveTemplate(copy);
}

// 导出 Template 对象
window.Template = {
    // 类型定义
    FIELD_TYPES,
    getFieldTypeOptions,
    
    // 创建函数
    createEmptyTemplate,
    createEmptyField,
    createEmptyDetailTable,
    createEmptyColumn,
    
    // 验证函数
    validateTemplate,
    checkTemplateConsistency,
    extractPlaceholdersFromWord,
    
    // 文件处理
    readFileAsBase64,
    base64ToBlob,
    
    // 其他
    getTemplateStats,
    duplicateTemplate
};

