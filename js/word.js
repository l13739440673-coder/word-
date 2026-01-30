/**
 * Word 文档生成模块
 * 使用 docxtemplater 根据模板和数据生成 Word 文档
 */

/**
 * 将 Base64 字符串转换为 ArrayBuffer
 * @param {string} base64 - Base64 编码的字符串
 * @returns {ArrayBuffer}
 */
function base64ToArrayBuffer(base64) {
    // 移除 data URL 前缀（如果有）
    const base64Data = base64.includes(',') ? base64.split(',')[1] : base64;
    const binaryString = atob(base64Data);
    const bytes = new Uint8Array(binaryString.length);
    
    for (let i = 0; i < binaryString.length; i++) {
        bytes[i] = binaryString.charCodeAt(i);
    }
    
    return bytes.buffer;
}

/**
 * 准备数据用于 Word 模板填充
 * 处理明细表数据格式，确保符合 docxtemplater 的要求
 * @param {Object} fieldData - 普通字段数据
 * @param {Object} tableData - 明细表数据
 * @param {Object} template - 模板配置
 * @returns {Object}
 */
function prepareDataForWord(fieldData, tableData, template) {
    const data = { ...fieldData };
    
    // 处理明细表数据
    if (template.detailTables && tableData) {
        template.detailTables.forEach(tableConfig => {
            const tableName = tableConfig.name;
            const rows = tableData[tableName] || [];
            // docxtemplater 需要数组形式的数据来处理循环
            data[tableName] = rows.map(row => {
                const processedRow = {};
                tableConfig.columns.forEach(col => {
                    let value = row[col.key];
                    // 处理空值
                    if (value === undefined || value === null || value === '') {
                        value = col.type === 'number' ? '' : '';
                    }
                    processedRow[col.key] = value;
                });
                return processedRow;
            });
        });
    }
    
    // 处理普通字段的空值
    template.fields.forEach(field => {
        if (data[field.key] === undefined || data[field.key] === null) {
            data[field.key] = '';
        }
    });
    
    return data;
}

/**
 * 生成 Word 文档
 * @param {Object} template - 模板数据（包含 wordFile）
 * @param {Object} fieldData - 普通字段数据
 * @param {Object} tableData - 明细表数据
 * @returns {Promise<Blob>} - 生成的 Word 文档 Blob
 */
async function generateWord(template, fieldData, tableData = {}) {
    if (!template.wordFile || !template.wordFile.data) {
        throw new Error('模板缺少 Word 文件');
    }
    
    try {
        // 将 Base64 转换为 ArrayBuffer
        const arrayBuffer = base64ToArrayBuffer(template.wordFile.data);
        
        // 加载 Word 模板
        const zip = new window.PizZip(arrayBuffer);
        
        // 创建 docxtemplater 实例
        const doc = new window.docxtemplater(zip, {
            paragraphLoop: true,
            linebreaks: true,
            delimiters: { start: '{', end: '}' }
        });
        
        // 准备数据
        const data = prepareDataForWord(fieldData, tableData, template);
        
        // 设置数据
        doc.setData(data);
        
        // 渲染文档
        doc.render();
        
        // 生成输出
        const output = doc.getZip().generate({
            type: 'blob',
            mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        });
        
        return output;
        
    } catch (error) {
        console.error('生成 Word 文档失败:', error);
        
        // 提供更友好的错误信息
        if (error.properties && error.properties.errors) {
            const templateErrors = error.properties.errors;
            const errorMessages = templateErrors.map(e => {
                if (e.properties && e.properties.id) {
                    return `占位符 "${e.properties.id}" 处理失败`;
                }
                return e.message || '未知错误';
            });
            throw new Error(`Word 模板错误：${errorMessages.join('; ')}`);
        }
        
        throw new Error('生成 Word 文档失败，请检查模板格式是否正确');
    }
}

/**
 * 生成并下载 Word 文档
 * @param {Object} template - 模板数据
 * @param {Object} fieldData - 普通字段数据
 * @param {Object} tableData - 明细表数据
 * @param {string} fileName - 文件名（可选）
 */
async function generateAndDownloadWord(template, fieldData, tableData = {}, fileName = null) {
    try {
        const blob = await generateWord(template, fieldData, tableData);
        
        // 生成文件名
        let outputFileName = fileName || template.name;
        
        // 确保文件名以 .docx 结尾
        if (!outputFileName.endsWith('.docx')) {
            outputFileName += '.docx';
        }
        
        // 使用原生方式下载（更可靠）
        downloadBlob(blob, outputFileName);
        
        return true;
        
    } catch (error) {
        console.error('下载 Word 文档失败:', error);
        throw error;
    }
}

/**
 * 通过原生方式下载 Blob 文件
 * @param {Blob} blob - 文件 Blob
 * @param {string} fileName - 文件名
 */
function downloadBlob(blob, fileName) {
    // 创建下载链接
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = fileName;
    
    // 触发下载
    document.body.appendChild(link);
    link.click();
    
    // 清理
    document.body.removeChild(link);
    setTimeout(() => URL.revokeObjectURL(url), 100);
}

/**
 * 预览 Word 文档数据（用于调试）
 * @param {Object} template - 模板数据
 * @param {Object} fieldData - 普通字段数据
 * @param {Object} tableData - 明细表数据
 * @returns {Object} - 准备好的数据
 */
function previewWordData(template, fieldData, tableData = {}) {
    return prepareDataForWord(fieldData, tableData, template);
}

/**
 * 格式化日期
 * @param {Date} date - 日期对象
 * @param {string} format - 格式字符串
 * @returns {string}
 */
function formatDate(date, format = 'YYYY-MM-DD') {
    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const day = String(date.getDate()).padStart(2, '0');
    const hours = String(date.getHours()).padStart(2, '0');
    const minutes = String(date.getMinutes()).padStart(2, '0');
    const seconds = String(date.getSeconds()).padStart(2, '0');
    
    return format
        .replace('YYYY', year)
        .replace('MM', month)
        .replace('DD', day)
        .replace('HH', hours)
        .replace('mm', minutes)
        .replace('ss', seconds);
}

/**
 * 验证 Word 模板是否有效
 * @param {string} base64Data - Base64 编码的 Word 文件数据
 * @returns {Object} - { valid: boolean, error?: string, placeholders?: Set, errors?: Array }
 */
function validateWordTemplate(base64Data) {
    try {
        const arrayBuffer = base64ToArrayBuffer(base64Data);
        const zip = new window.PizZip(arrayBuffer);
        
        // 尝试创建 docxtemplater 实例
        const doc = new window.docxtemplater(zip, {
            paragraphLoop: true,
            linebreaks: true,
            delimiters: { start: '{', end: '}' }
        });
        
        // 提取占位符
        const fullText = doc.getFullText();
        const placeholders = new Set();
        const regex = /\{([^{}]+)\}/g;
        let match;
        
        while ((match = regex.exec(fullText)) !== null) {
            const tag = match[1].trim();
            // 过滤掉循环标签
            if (!tag.startsWith('#') && !tag.startsWith('/')) {
                placeholders.add(tag);
            }
        }
        
        return {
            valid: true,
            placeholders
        };
        
    } catch (error) {
        console.error('验证 Word 模板失败:', error);
        
        // 处理 docxtemplater 的多错误情况
        if (error.properties && error.properties.errors) {
            const detailedErrors = error.properties.errors.map(e => {
                let errorMsg = '';
                
                if (e.properties) {
                    const props = e.properties;
                    
                    // 错误类型
                    if (props.id === 'unopened_tag') {
                        errorMsg = `未开始的结束标签: "{${props.xtag}}" (位置: ${props.offset})`;
                    } else if (props.id === 'unopened_loop') {
                        errorMsg = `循环未开始: 找到结束标签 "{/${props.xtag}}" 但缺少开始标签 "{#${props.xtag}}" (位置: ${props.offset})`;
                    } else if (props.id === 'unclosed_tag') {
                        errorMsg = `未闭合的开始标签: "{${props.xtag}}" (位置: ${props.offset})`;
                    } else if (props.id === 'unclosed_loop') {
                        errorMsg = `循环未闭合: 找到开始标签 "{#${props.xtag}}" 但缺少结束标签 "{/${props.xtag}}" (位置: ${props.offset})`;
                    } else if (props.id === 'closing_tag_does_not_match_opening_tag') {
                        errorMsg = `开始标签和结束标签不匹配: 开始 "{#${props.openingtag}", 结束 "{/${props.closingtag}}"`;
                    } else if (props.id === 'raw_xml_tag_should_be_only_text_in_paragraph') {
                        errorMsg = `原始 XML 标签必须单独在一个段落中`;
                    } else if (props.id) {
                        errorMsg = `错误类型: ${props.id}`;
                        if (props.xtag) errorMsg += `, 标签: "${props.xtag}"`;
                        if (props.offset) errorMsg += `, 位置: ${props.offset}`;
                    }
                }
                
                if (!errorMsg) {
                    errorMsg = e.message || '未知错误';
                }
                
                return errorMsg;
            });
            
            return {
                valid: false,
                error: `Word 模板有 ${detailedErrors.length} 个错误`,
                errors: detailedErrors
            };
        }
        
        return {
            valid: false,
            error: error.message || '无法解析 Word 文件'
        };
    }
}

/**
 * 获取 Word 模板中的所有占位符
 * @param {string} base64Data - Base64 编码的 Word 文件数据
 * @returns {Array<string>} - 占位符列表
 */
function getWordPlaceholders(base64Data) {
    const result = validateWordTemplate(base64Data);
    if (result.valid && result.placeholders) {
        return Array.from(result.placeholders);
    }
    return [];
}

// 导出 Word 对象
window.Word = {
    generateWord,
    generateAndDownloadWord,
    previewWordData,
    validateWordTemplate,
    getWordPlaceholders,
    formatDate,
    base64ToArrayBuffer
};

