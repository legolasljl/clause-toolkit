#!/usr/bin/env node
/**
 * 对公版本构建脚本
 * 读取 clause_toolkit_offline.html → 混淆关键代码 + 更换授权密钥 + 反调试 → 输出 clause_toolkit_corporate.html
 */

const fs = require('fs');
const path = require('path');
const JavaScriptObfuscator = require('javascript-obfuscator');

const BASE_DIR = __dirname;
const INPUT_FILE = path.join(BASE_DIR, 'clause_toolkit_offline.html');
const OUTPUT_FILE = path.join(BASE_DIR, 'clause_toolkit_corporate.html');

// ======== 新密钥配置（对公版专用，与旧密钥不同） ========
const NEW_SECRET_KEY = 'CNexusCorp@2026#Ent!Secure';
const NEW_STORAGE_KEY = 'cnexus_corp_license';
const NEW_USED_CODES_KEY = 'cnexus_corp_used_codes';
const NEW_DEVICE_ID_KEY = 'cnexus_corp_device_id';
const NEW_PERMANENT_CODE = 'CORPACTIVATEFOREVER';  // 新的永久码（19字符，旧的16字符码不匹配）

// ======== 反调试 + F12屏蔽代码（允许复制） ========
const ANTI_DEBUG_CODE = `
// === 对公版保护 ===
(function(){
    // 屏蔽F12、Ctrl+Shift+I/J/C、Ctrl+U
    document.addEventListener('keydown',function(e){
        if(e.key==='F12'||(e.ctrlKey&&e.shiftKey&&['I','i','J','j','C','c'].indexOf(e.key)!==-1)||(e.ctrlKey&&(e.key==='u'||e.key==='U'))){
            e.preventDefault();e.stopPropagation();return false;
        }
    },true);
    // 屏蔽右键菜单（但保留文本选择和Ctrl+C/V复制粘贴功能）
    document.addEventListener('contextmenu',function(e){e.preventDefault();return false;},true);
    // 检测DevTools打开（基于窗口尺寸差异）
    var _dc=0;
    setInterval(function(){
        var w=window.outerWidth-window.innerWidth>160;
        var h=window.outerHeight-window.innerHeight>160;
        if(w||h){_dc++;if(_dc>3){document.body.innerHTML='<div style="display:flex;align-items:center;justify-content:center;height:100vh;font-size:24px;color:#c75050;font-family:sans-serif;">⚠️ 请关闭开发者工具后刷新页面</div>';}}
        else{_dc=0;}
    },1500);
    // debugger陷阱（定时触发）
    var _dd=function(){try{(function(){return false;})['constructor']('debugger')();}catch(e){}};
    setInterval(_dd,3000);
})();
`;

// ======== 混淆配置 ========
const OBFUSCATOR_OPTIONS = {
    compact: true,
    controlFlowFlattening: true,
    controlFlowFlatteningThreshold: 0.5,
    deadCodeInjection: true,
    deadCodeInjectionThreshold: 0.2,
    identifierNamesGenerator: 'hexadecimal',
    renameGlobals: false,  // 不重命名全局变量（HTML onclick引用）
    rotateStringArray: true,
    selfDefending: false,
    stringArray: true,
    stringArrayEncoding: ['base64'],
    stringArrayThreshold: 0.6,
    transformObjectKeys: false,
    unicodeEscapeSequence: false,
    // 保留HTML引用的函数名
    reservedNames: [
        // 授权系统
        'formatLicenseCode', 'verifyAndActivate', 'checkSavedLicense',
        'LICENSE_CONFIG', 'verifyLicenseCode',
        // Tab切换
        'switchMainTab',
        // 条款提取
        'handleFileSelect', 'processAllClauses', 'downloadAllResults',
        'toggleClause', 'downloadClause', 'downloadAllJson',
        'showFullClause', 'closeFullClause', 'copyClause',
        // 条款对比
        'compareClauses', 'handleDragOver', 'handleDragLeave',
        'handleDrop', 'removeFile', 'loadStandardData',
        'showCompareDetail', 'closeCompareDetail', 'downloadCompareJson',
        'downloadExcel',
        // 主险计算
        'mc_onProductChange', 'mc_onVersionChange', 'mc_onMethodChange',
        'mc_onTermChange', 'mc_calculate', 'mc_reset', 'mc_sendToAdditional',
        'mc_showIndustryLookup', 'mc_hideIndustryLookup', 'mc_filterIndustry',
        'mc_selectIndustryClass', 'mc_triggerJsonUpload', 'mc_handleJsonFile',
        'mc_triggerDocxUpload', 'mc_handleDocxFile', 'mc_confirmDocxImport',
        'mc_cancelDocxImport', 'mc_selectCoeffRow', 'mc_onSliderChange',
        'mc_onDisabilityTabClick', 'mc_onDisabilityTableChange',
        'mc_onAddonPromptAnswer', 'mc_showAddonModal', 'mc_hideAddonModal',
        'mc_selectAddonColumn', 'mc_selectDisabilityOption',
        // 附加险计算
        'rc_triggerJsonUpload', 'rc_handleJsonFile', 'rc_triggerFolderUpload',
        'rc_handleFolderFiles', 'rc_onSearch', 'rc_selectClause',
        'rc_selectCoeffRow', 'rc_calcSingle', 'rc_resetCoeff',
        'rc_handleInquiryImport', 'rc_batchCalculate', 'rc_downloadExcel',
        'rc_onSliderChange',
        // 条款查询
        'showQueryModal', 'closeQueryModal', 'queryClauseContent',
        'downloadQueryResult',
        // 捐赠
        'showDonateModal', 'closeDonateModal',
        // mc_updateParamsVisibility等内部函数保持原名以兼容
        'mc_updateParamsVisibility', 'mc_getProductType', 'mc_renderCoefficients',
        'mc_renderDisabilityCoeffSection', 'mc_renderAddonTable',
    ],
    reservedStrings: []
};

// 实际使用的混淆配置（平衡强度和兼容性）
const EFFECTIVE_OBFUSCATOR_OPTIONS = {
    compact: true,
    controlFlowFlattening: true,
    controlFlowFlatteningThreshold: 0.4,
    deadCodeInjection: false,
    identifierNamesGenerator: 'hexadecimal',
    renameGlobals: false,
    rotateStringArray: true,
    selfDefending: false,
    stringArray: true,
    stringArrayEncoding: ['base64'],
    stringArrayThreshold: 0.9,
    stringArrayWrappersCount: 2,
    transformObjectKeys: true,
    unicodeEscapeSequence: false,
    numbersToExpressions: true,
    reservedNames: OBFUSCATOR_OPTIONS.reservedNames,
};

function main() {
    console.log('📦 开始构建对公版本...');

    // 1. 读取源文件
    const html = fs.readFileSync(INPUT_FILE, 'utf-8');
    console.log('  ✅ 读取源文件完成');

    // 2. 提取最后一个内联script（主应用代码）
    // 定位最后一个 <script> ... </script> 对（不含src属性的）
    const lastScriptOpen = html.lastIndexOf('    <script>\n');
    const lastScriptClose = html.lastIndexOf('\n    </script>');
    if (lastScriptOpen === -1 || lastScriptClose === -1 || lastScriptClose <= lastScriptOpen) {
        console.error('❌ 无法找到内联script标签');
        process.exit(1);
    }

    const scriptTagEnd = lastScriptOpen + '    <script>\n'.length;
    const htmlBefore = html.substring(0, scriptTagEnd);
    let jsCode = html.substring(scriptTagEnd, lastScriptClose);
    const htmlAfter = html.substring(lastScriptClose);

    console.log('  ✅ 提取JS代码: ' + jsCode.length + ' 字符');

    // 3. 替换密钥和存储键
    jsCode = jsCode.replace(
        "SECRET_KEY: 'ClauseNexus2026SecretKey!@#'",
        "SECRET_KEY: '" + NEW_SECRET_KEY + "'"
    );
    jsCode = jsCode.replace(
        "STORAGE_KEY: 'clausenexus_license'",
        "STORAGE_KEY: '" + NEW_STORAGE_KEY + "'"
    );
    jsCode = jsCode.replace(
        "USED_CODES_KEY: 'clausenexus_used_codes'",
        "USED_CODES_KEY: '" + NEW_USED_CODES_KEY + "'"
    );
    jsCode = jsCode.replace(
        "DEVICE_ID_KEY: 'clausenexus_device_id'",
        "DEVICE_ID_KEY: '" + NEW_DEVICE_ID_KEY + "'"
    );
    // 替换永久码
    jsCode = jsCode.replace(
        "code === 'ALEXLOVESIVYMORE'",
        "code === '" + NEW_PERMANENT_CODE + "'"
    );
    console.log('  ✅ 密钥替换完成');

    // 4. 更新版本标识
    jsCode = jsCode.replace(/v18\.8/g, 'v19.0 企业版');
    jsCode = jsCode.replace(/v18\.11/g, 'v19.0');

    // 5. 分段混淆关键代码
    // 将JS分为多个区段，对关键区段使用强混淆，其他使用轻混淆
    console.log('  🔄 开始代码混淆（这可能需要一些时间）...');

    // 识别关键代码段的标记
    const criticalPatterns = [
        // 授权系统
        { start: '// 激活码验证系统', end: '// 全局变量和Tab切换' },
        // 费率数据和计算
        { start: 'const MC_PRODUCTS = {', end: '// --- 主险计算器状态 ---' },
        // 伤残赔偿数据
        { start: '// --- 伤残赔偿比例附表数据 ---', end: 'let mc_selectedDisabilityTable' },
        // 计算函数
        { start: '// --- 计算主险保费 ---', end: '// --- 渲染结果 ---' },
    ];

    // 整体混淆（使用轻量配置以保持函数名兼容）
    let obfuscatedJs;
    try {
        const result = JavaScriptObfuscator.obfuscate(jsCode, EFFECTIVE_OBFUSCATOR_OPTIONS);
        obfuscatedJs = result.getObfuscatedCode();
        console.log('  ✅ 代码混淆完成: ' + obfuscatedJs.length + ' 字符');
    } catch (err) {
        console.error('❌ 混淆失败:', err.message);
        console.log('  ⚠️ 回退为仅替换密钥（不混淆）');
        obfuscatedJs = jsCode;
    }

    // 6. 组装最终HTML
    // 修改HTML部分：更新版本号、添加反调试标识
    let finalHtmlBefore = htmlBefore;
    finalHtmlBefore = finalHtmlBefore.replace(/v18\.8/g, 'v19.0 企业版');
    finalHtmlBefore = finalHtmlBefore.replace(
        '<small style="color:#aaa;">v18.8 · ClauseNexus</small>',
        '<small style="color:#aaa;">v19.0 · ClauseNexus 企业版</small>'
    );
    // 移除捐赠按钮（对公版不需要）
    finalHtmlBefore = finalHtmlBefore.replace(
        /<button[^>]*class="donate-btn"[^>]*>[^<]*<\/button>/g,
        ''
    );
    // 移除打赏模态框整体（从注释到结束）
    finalHtmlBefore = finalHtmlBefore.replace(
        /<!-- 打赏模态框 -->[\s\S]*?(?=\n    <!-- (?!打赏))/g,
        ''
    );
    // 替换复用donate-close-btn类名的按钮为内联样式
    finalHtmlBefore = finalHtmlBefore.replace(
        /class="donate-close-btn"/g,
        'style="display:block;width:100%;padding:10px;background:var(--text-primary);color:var(--bg-primary);border:none;border-radius:8px;font-size:14px;cursor:pointer;"'
    );

    let finalHtmlAfter = htmlAfter;
    finalHtmlAfter = finalHtmlAfter.replace(/v18\.8/g, 'v19.0 企业版');

    // 组装
    let finalHtml = finalHtmlBefore + ANTI_DEBUG_CODE + '\n' + obfuscatedJs + finalHtmlAfter;

    // 7. 清理所有donate相关CSS残留（逐个匹配donate样式块）
    // 匹配所有以.donate-开头的CSS规则
    finalHtml = finalHtml.replace(/\s*\.donate-[a-z-]+(?:\.[a-z-]+)*\s*\{[^}]*\}/g, '');
    // 匹配@keyframes donate-glow
    finalHtml = finalHtml.replace(/\s*@keyframes\s+donate-glow\s*\{[^}]*\{[^}]*\}[^}]*\}/g, '');
    // 移除残留的donate按钮
    finalHtml = finalHtml.replace(/<button[^>]*onclick="showDonateModal\(\)"[^>]*>[^<]*<\/button>/g, '');

    // 8. 写入输出文件
    fs.writeFileSync(OUTPUT_FILE, finalHtml, 'utf-8');
    console.log('  ✅ 输出文件: ' + OUTPUT_FILE);
    console.log('  📊 文件大小: ' + (finalHtml.length / 1024).toFixed(1) + ' KB');
    console.log('');
    console.log('🎉 对公版本构建完成！');
    console.log('');
    console.log('📋 新密钥配置:');
    console.log('   SECRET_KEY: ' + NEW_SECRET_KEY);
    console.log('   永久激活码: ' + NEW_PERMANENT_CODE);
    console.log('   存储键前缀: cnexus_corp_');
}

main();
