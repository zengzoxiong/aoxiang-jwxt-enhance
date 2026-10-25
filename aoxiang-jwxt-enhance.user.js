// ==UserScript==
// @name         翱翔教务功能加强
// @namespace    http://tampermonkey.net/
// @version      1.8.4
// @description  1.提供GPA分析报告；2. 导出课程成绩与教学班排名；3.更好的“学生画像”显示；4.选课助手；5.课程关注与后台同步；6.一键自动评教；7.人员信息检索
// @author       leamloli
// @match        https://jwxt.nwpu.edu.cn/*
// @match        https://jwxt.nwpu.edu.cn/student/for-std/course-select/some-page*
// @match        https://ecampus.nwpu.edu.cn/*
// @match        https://teacher.nwpu.edu.cn/*
// @match        https://yjsjy.nwpu.edu.cn/*
// @grant        GM_xmlhttpRequest
// @grant        GM_setValue
// @grant        GM_getValue
// @grant        unsafeWindow
// @connect      electronic-signature.nwpu.edu.cn
// @connect      jwxt.nwpu.edu.cn
// @connect      gitee.com
// @connect      raw.giteeusercontent.com
// @connect      raw.githubusercontent.com
// @license      MIT
// @icon         https://www.google.com/s2/favicons?sz=64&domain=nwpu.edu.cn
// @run-at       document-start
// @require      https://cdnjs.cloudflare.com/ajax/libs/exceljs/4.3.0/exceljs.min.js
// @require      https://cdnjs.cloudflare.com/ajax/libs/FileSaver.js/2.0.5/FileSaver.min.js
// @homepage     https://github.com/zengzoxiong/aoxiang-jwxt-enhance
// @downloadURL https://gitee.com/leamloli/aoxiang-jwxt-enhance/raw/master/aoxiang-jwxt-enhance.user.js
// @updateURL https://gitee.com/leamloli/aoxiang-jwxt-enhance/raw/master/aoxiang-jwxt-enhance.meta.js
// ==/UserScript==
(function () {
    'use strict';

    // ============================================================
    // 【脚本启动流程总览】（环境守卫 → 路由分发 → 页面初始化）
    //   document-start 加载期（IIFE 顶层顺序执行）：
    //     0.0 拦截异常图片请求（优化加载速度）
    //     0.x 定义常量 / 工具 / 存储（ConfigManager、GmDialog 等）
    //     2.x 各功能模块定义（悬浮球、同步、存档、开课查询、选课助手等）
    //     TextbookInfoModule.installHook()：必须抢在页面自身 XHR 之前挂载拦截器，
    //       故保持在加载期无条件执行（内部有防重复挂载守卫）
    //   DOMContentLoaded（见文件末尾）：
    //     runMainFeatures() 统一路由分发，所有页面入口仅此一处：
    //       ① window.name === Worker 名      → 后台同步 Worker，直接返回
    //       ② ecampus.nwpu.edu.cn            → 门户同步 + 门户主页功能，返回
    //       ③ jwxt 顶层窗口（非画像页）       → 延迟 6s 预取画像数据（加速画像页）
    //       ④ 评教页    → initEvaluationHelper（load/hashchange/兜底三重触发，内部有幂等守卫）
    //       ⑤ 开课查询  → LessonSearchEnhancer.init()（2.5 同步面板）+ initLessonSearchPage()（2.6 课程关注）
    //       ⑥ 学生画像  → enhancePortraitPage()
    //       ⑦ 培养方案  → initProgramPageEnhancement()
    //       ⑧ 我的课表  → TextbookInfoModule.initUI()（UI 阶段；数据阶段由 installHook 缓存兜底）
    //       ⑨ 教师搜索  → 跨标签页自动搜索注入（teacher.nwpu.edu.cn）
    //       ⑩ 顶层主页 / 研究生系统 → 延迟启动后台同步控制器（研究生系统另含培养方案增强）
    //   例外：/course-selection 选课助手（2.7）为历史遗留的独立 IIFE，
    //     自带 main() 与 readyState 等待，待后续版本迁移至统一路由。
    // ============================================================

    // =============== 0.0 拦截浏览器的异常请求，优化网页加载速度 ===============
    try {
        const BAD_KEY = 'burp';

        // 1. 劫持 HTMLImageElement 原型链上的 src 属性
        const imageProto = HTMLImageElement.prototype;
        const originalSrcDescriptor = Object.getOwnPropertyDescriptor(imageProto, 'src');

        if (originalSrcDescriptor) {
            Object.defineProperty(imageProto, 'src', {
                get: function () {
                    return originalSrcDescriptor.get.call(this);
                },
                set: function (value) {
                    if (value && typeof value === 'string' && value.indexOf(BAD_KEY) !== -1) {
                        return;
                    }
                    originalSrcDescriptor.set.call(this, value);
                },
                configurable: true,
                enumerable: true
            });
        }

        // 2. 劫持 setAttribute 方法
        const originalSetAttribute = Element.prototype.setAttribute;
        Element.prototype.setAttribute = function (name, value) {
            if (this instanceof HTMLImageElement && name === 'src' && value && value.indexOf(BAD_KEY) !== -1) {
                return;
            }
            return originalSetAttribute.apply(this, arguments);
        };

    } catch (e) {
        console.error('[NWPU-Enhanced] 拦截器初始化异常', e);
    }

    // =-=-=-=-=-=-=-=-=-=-=-=-= 0. 基础工具与日志系统 =-=-=-=-=-=-=-=-=-=-=-=-=

    // --- 全局常量定义 ---
    const CONSTANTS = {
        CACHE_KEY: 'jwxtEnhancedDataCache',
        FOLLOWED_COURSES_KEY: 'jwxt_followed_courses_list',
        BACKGROUND_SYNC_KEY: 'jwxt_background_sync_data',
        LAST_SYNC_TIME_KEY: 'jwxt_last_bg_sync_time',
        HISTORY_STORAGE_KEY: 'course_enrollment_history_auto_sync',
        // 学号在 localStorage 的兼容键（选课助手 / 画像页 / 诊断报告共用）
        STUDENT_ID_LS_KEY: 'cs-course-select-student-id',
        // 跨标签页人员检索：姓名的临时传递键
        CROSS_SEARCH_NAME_KEY: 'gm_cross_search_name',
        // 首次运行「脚本使用说明」弹窗是否已展示
        INTRO_SHOWN_KEY: 'jwxt_enhanced_v162_intro_shown',
        // 统一认证身份（本科生 / 研究生），用于选课时间提醒
        USER_IDENTITY_KEY: 'gm_user_identity_type',
        USER_IDENTITY_NAME_KEY: 'gm_user_identity_name',
        // 选课时间表远程配置缓存
        SCHEDULE_CACHE_KEY: 'gm_schedule_config_cache_v1',
        SYNC_COOLDOWN_MS: 1 * 60 * 60 * 1000, // 1小时冷却
        // 中文等级制成绩到百分制分数的映射（用于等级制成绩折算百分制 GPA）
        GRADE_MAPPING_CONFIG: {
            '优秀': 93,
            '良好': 80,
            '中等': 70,
            '及格': 60,
            '不及格': 0
        },
        // 学生画像专用：记录见过的 StudentID（可能多个），用于后台预取加速画像页加载
        PORTRAIT_STUDENT_IDS_KEY: 'jwxt_portrait_student_ids',
        // 学生画像专用：跨标签页预取缓存 { studentId, data, time }
        PORTRAIT_PREFETCH_KEY: 'jwxt_portrait_prefetch_data'
    };

    /**
     * 统一日志输出工具
     * @description 所有控制台输出统一带有 [NWPU-Enhanced] 前缀
     * 注意：不使用 %c 彩色格式（Tampermonkey 日志视图无法解析该占位符，会导致样式字符串被原样拼接输出）
     */
    const Logger = {
        _print: (module, msg, type = 'log', args = []) => {
            const prefix = `[NWPU-Enhanced][${module}]`;
            if (args.length > 0) {
                console[type](prefix, msg, ...args);
            } else {
                console[type](prefix, msg);
            }
        },
        log: (module, msg, ...args) => Logger._print(module, msg, 'log', args),
        warn: (module, msg, ...args) => Logger._print(module, msg, 'warn', args),
        error: (module, msg, ...args) => Logger._print(module, msg, 'error', args),
        info: (module, msg, ...args) => Logger._print(module, msg, 'info', args)
    };

    // HTML 转义工具（弹窗文案 / 存档区存档名 / 重命名输入等共用）
    const esc = str => String(str).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

    // 统一样式元素获取：按 id 复用已存在的 <style>，否则新建一个（textContent 设置与挂载由调用方负责）
    function getOrCreateStyleEl(id) {
        let el = document.getElementById(id);
        if (!el) { el = document.createElement('style'); el.id = id; }
        return el;
    }

    // ============================================================
    // 0.1 统一视觉系统（设计令牌 + 组件覆盖）
    // 说明：只收敛颜色、字号、圆角、阴影与控件状态，不改布局和交互逻辑。
    // 使用 html body 前缀提高权重，覆盖各功能模块后注入的同名样式。
    // ============================================================
    function injectUnifiedDesignSystem() {
        if (document.getElementById('gm-unified-design-system')) return;
        const style = document.createElement('style');
        style.id = 'gm-unified-design-system';
        style.textContent = `
        :root {
            --gm-primary: #409eff;
            --gm-primary-hover: #66b1ff;
            --gm-primary-active: #337ecc;
            --gm-primary-text: #1a6fc4;
            --gm-primary-light: #ecf5ff;
            --gm-primary-border: #b3d8ff;
            --gm-success: #67c23a;
            --gm-success-bg: #f0f9eb;
            --gm-success-border: #e1f3d8;
            --gm-success-text: #2e7d32;
            --gm-warning: #e6a23c;
            --gm-warning-bg: #fdf6ec;
            --gm-warning-border: #faecd8;
            --gm-warning-text: #a85b00;
            --gm-danger: #f56c6c;
            --gm-danger-bg: #fef0f0;
            --gm-danger-border: #fab6b6;
            --gm-danger-text: #c53030;
            --gm-info: #909399;
            --gm-info-bg: #f4f4f5;
            --gm-info-border: #e9e9eb;
            --gm-text-1: #0f172a;
            --gm-text-2: #334155;
            --gm-text-3: #5b6b80;
            --gm-text-4: #94a3b8;
            --gm-border: #e2e8f0;
            --gm-border-light: #f1f5f9;
            --gm-fill: #f8fafc;
            --gm-fill-strong: #f1f5f9;
            --gm-radius-sm: 6px;
            --gm-radius: 8px;
            --gm-radius-lg: 12px;
            --gm-radius-xl: 16px;
            --gm-radius-pill: 999px;
            --gm-shadow-sm: 0 1px 2px rgba(15, 23, 42, .04), 0 2px 8px rgba(15, 23, 42, .04);
            --gm-shadow: 0 8px 24px rgba(15, 23, 42, .10);
            --gm-shadow-lg: 0 16px 40px rgba(15, 23, 42, .16);
            --gm-font: "PingFang SC", "Hiragino Sans GB", "Microsoft YaHei", "Helvetica Neue", Arial, system-ui, sans-serif;
            --gm-mono: "SFMono-Regular", Consolas, "Liberation Mono", monospace;
        }

        /* ---------- 遮罩与弹窗骨架 ---------- */
        html body .gm-dialog-overlay,
        html body .gm-modal-overlay,
        html body .gpa-report-overlay,
        html body .gm-help-overlay,
        html body .gm-fl-modal-overlay,
        html body .gm-sync-overlay,
        html body .semester-bg-overlay {
            background: rgba(15, 23, 42, .48) !important;
            backdrop-filter: blur(4px) !important;
        }
        html body .gm-dialog,
        html body .gm-modal-content,
        html body .gpa-report-modal,
        html body .gm-help-modal,
        html body .gm-fl-modal,
        html body .gm-sync-modal-card,
        html body .semester-checkbox-container,
        html body #gm-lesson-helper-panel,
        html body .gm-textbook-wrapper {
            border-radius: var(--gm-radius-lg) !important;
            border: 1px solid var(--gm-border) !important;
            box-shadow: var(--gm-shadow-lg) !important;
            background: #fff !important;
            font-family: var(--gm-font) !important;
        }
        html body .gm-modal-header,
        html body .gpa-modal-header,
        html body .gm-help-header,
        html body .gm-fl-header {
            padding: 16px 24px !important;
            border-bottom: 1px solid var(--gm-border) !important;
            background: #fff !important;
            height: auto !important;
            min-height: 58px !important;
        }
        html body .gm-modal-title,
        html body .gpa-modal-title,
        html body .gm-help-title,
        html body .gm-fl-title,
        html body .gm-sync-modal-title {
            font-family: var(--gm-font) !important;
            font-size: 16px !important;
            font-weight: 700 !important;
            color: var(--gm-text-1) !important;
            gap: 8px !important;
            letter-spacing: 0 !important;
        }
        html body .gm-modal-close,
        html body .gpa-close-btn,
        html body .gm-close-btn,
        html body .gm-fl-close,
        html body .gm-help-close,
        html body #gm-sync-close,
        html body #gm-panel-close {
            width: 30px !important;
            height: 30px !important;
            min-width: 30px !important;
            border: 0 !important;
            border-radius: 50% !important;
            background: transparent !important;
            color: var(--gm-text-3) !important;
            font-family: var(--gm-font) !important;
            font-size: 18px !important;
            line-height: 1 !important;
            display: inline-flex !important;
            align-items: center !important;
            justify-content: center !important;
            padding: 0 !important;
            cursor: pointer !important;
            transition: background .18s ease, color .18s ease, transform .18s ease !important;
        }
        html body .gm-modal-close:hover,
        html body .gpa-close-btn:hover,
        html body .gm-close-btn:hover,
        html body .gm-fl-close:hover,
        html body .gm-help-close:hover,
        html body #gm-sync-close:hover,
        html body #gm-panel-close:hover {
            background: var(--gm-danger-bg) !important;
            color: var(--gm-danger-text) !important;
            transform: rotate(90deg) !important;
        }

        /* ---------- 按钮与输入 ---------- */
        html body .gm-dialog-btn,
        html body .gm-btn,
        html body .calculate-btn,
        html body .confirm-export-button,
        html body .select-all-button,
        html body .cancel-button,
        html body .gm-ps-btn,
        html body .semester-checkbox-container button {
            min-height: 34px !important;
            height: auto !important;
            padding: 0 14px !important;
            border-radius: var(--gm-radius) !important;
            font-family: var(--gm-font) !important;
            font-size: 13px !important;
            font-weight: 600 !important;
            line-height: 1 !important;
            box-sizing: border-box !important;
            transition: background .18s ease, border-color .18s ease, color .18s ease, box-shadow .18s ease, transform .18s ease !important;
        }
        html body .gm-btn-primary,
        html body .calculate-btn,
        html body .confirm-export-button,
        html body .gm-ps-btn {
            background: var(--gm-primary) !important;
            border: 1px solid var(--gm-primary) !important;
            color: #fff !important;
            box-shadow: 0 4px 10px rgba(64, 158, 255, .22) !important;
        }
        html body .gm-btn-primary:hover,
        html body .calculate-btn:hover,
        html body .confirm-export-button:hover,
        html body .gm-ps-btn:hover {
            background: var(--gm-primary-hover) !important;
            border-color: var(--gm-primary-hover) !important;
        }
        html body .gm-btn-warning {
            background: var(--gm-warning) !important;
            border-color: var(--gm-warning) !important;
            color: #fff !important;
        }
        html body .select-all-button {
            background: #fff !important;
            border: 1px solid var(--gm-border) !important;
            color: var(--gm-text-2) !important;
        }
        html body .select-all-button:hover {
            background: var(--gm-primary-light) !important;
            border-color: var(--gm-primary-border) !important;
            color: var(--gm-primary-text) !important;
        }
        html body .cancel-button {
            background: transparent !important;
            border: 1px solid transparent !important;
            color: var(--gm-text-3) !important;
        }
        html body .cancel-button:hover {
            background: var(--gm-fill-strong) !important;
            color: var(--gm-text-1) !important;
        }
        html body .gm-dialog-btn {
            border: 1px solid var(--gm-border) !important;
            color: var(--gm-text-2) !important;
        }
        html body .gm-dialog-btn:not(.gm-dialog-btn-primary):hover {
            background: var(--gm-fill) !important;
            border-color: var(--gm-border) !important;
            color: var(--gm-text-1) !important;
        }
        html body .gm-dialog-btn-primary {
            border: 0 !important;
            color: #fff !important;
        }
        html body .gm-dialog-btn-primary:hover {
            opacity: .9 !important;
            color: #fff !important;
        }
        html body .gm-ps-input,
        html body .input-group input,
        html body .gm-filter-select,
        html body .gm-score-input,
        html body #gm-semester-dd-trigger {
            border: 1px solid var(--gm-border) !important;
            border-radius: var(--gm-radius) !important;
            background: #fff !important;
            color: var(--gm-text-1) !important;
            font-family: var(--gm-font) !important;
            font-size: 13px !important;
            box-sizing: border-box !important;
            outline: none !important;
            transition: border-color .18s ease, box-shadow .18s ease !important;
        }
        html body .gm-ps-input,
        html body .input-group input {
            height: 36px !important;
            padding: 0 11px !important;
        }
        html body .gm-filter-select {
            height: 32px !important;
            padding: 0 10px !important;
            font-size: 12px !important;
        }
        html body .gm-score-input {
            width: 68px !important;
            height: 30px !important;
            padding: 0 6px !important;
            border-radius: var(--gm-radius-sm) !important;
            font-family: var(--gm-mono) !important;
            font-size: 12px !important;
            text-align: center !important;
        }
        html body #gm-semester-dd-trigger {
            height: 36px !important;
            padding: 0 30px 0 10px !important;
            font-size: 12.5px !important;
        }
        html body .gm-ps-input:focus,
        html body .input-group input:focus,
        html body .gm-filter-select:focus,
        html body .gm-score-input:focus,
        html body #gm-semester-dd-trigger:focus {
            border-color: var(--gm-primary) !important;
            box-shadow: 0 0 0 3px rgba(64, 158, 255, .18) !important;
        }
        html body .gm-checkbox,
        html body .gm-modal-content input[type="checkbox"],
        html body .semester-checkbox-container input[type="checkbox"] {
            accent-color: var(--gm-primary) !important;
        }

        /* ---------- 标签 / 状态 / 代码 ---------- */
        html body .gm-tag,
        html body .gm-arc-badge,
        html body .gm-course-status-tag {
            display: inline-flex !important;
            align-items: center !important;
            min-height: 22px !important;
            height: auto !important;
            padding: 0 8px !important;
            border-radius: var(--gm-radius-sm) !important;
            font-family: var(--gm-font) !important;
            font-size: 11.5px !important;
            font-weight: 600 !important;
            line-height: 1 !important;
            border: 1px solid transparent !important;
            box-sizing: border-box !important;
            white-space: nowrap !important;
        }
        html body .gm-tag-blue {
            background: var(--gm-primary-light) !important;
            border-color: var(--gm-primary-border) !important;
            color: var(--gm-primary-text) !important;
        }
        html body .gm-tag-red {
            background: var(--gm-danger-bg) !important;
            border-color: var(--gm-danger-border) !important;
            color: var(--gm-danger-text) !important;
        }
        html body .gm-tag-gray {
            background: var(--gm-info-bg) !important;
            border-color: var(--gm-info-border) !important;
            color: var(--gm-text-3) !important;
        }
        html body .gm-tag-done {
            background: var(--gm-success-bg) !important;
            border-color: var(--gm-success-border) !important;
            color: var(--gm-success-text) !important;
        }
        html body .gm-tag-todo {
            background: var(--gm-warning-bg) !important;
            border-color: var(--gm-warning-border) !important;
            color: var(--gm-warning-text) !important;
        }
        html body .gm-arc-badge {
            min-height: 18px !important;
            padding: 0 5px !important;
            font-size: 10px !important;
            background: var(--gm-primary-light) !important;
            border-color: var(--gm-primary-border) !important;
            color: var(--gm-primary-text) !important;
        }
        html body .gm-code-badge {
            display: inline-flex !important;
            align-items: center !important;
            padding: 2px 6px !important;
            border-radius: var(--gm-radius-sm) !important;
            background: var(--gm-fill-strong) !important;
            color: var(--gm-text-3) !important;
            font-family: var(--gm-mono) !important;
            font-size: 10.5px !important;
        }
        html body .gm-course-code-highlight {
            border: 1.5px dashed var(--gm-primary) !important;
            color: var(--gm-primary-text) !important;
            background: var(--gm-primary-light) !important;
            border-radius: var(--gm-radius-sm) !important;
            padding: 2px 6px !important;
            font-size: 13px !important;
            font-weight: 600 !important;
        }
        html body .gm-course-code-highlight:hover {
            background: var(--gm-primary-border) !important;
            color: var(--gm-primary-active) !important;
            border-color: var(--gm-primary-active) !important;
        }
        html body .gm-hist-tag {
            display: inline-flex !important;
            align-items: center !important;
            padding: 2px 6px !important;
            border: 1px solid transparent !important;
            border-radius: 4px !important;
            font-family: var(--gm-font) !important;
            font-size: 11px !important;
            font-weight: 700 !important;
            line-height: 1.4 !important;
            white-space: nowrap !important;
        }
        html body .gm-hist-tag--ok {
            background: var(--gm-success-bg) !important;
            border-color: var(--gm-success-border) !important;
            color: var(--gm-success-text) !important;
        }
        html body .gm-hist-tag--full {
            background: var(--gm-danger-bg) !important;
            border-color: var(--gm-danger-border) !important;
            color: var(--gm-danger-text) !important;
        }
        html body .gm-release-tag {
            display: inline-flex !important;
            align-items: center !important;
            gap: 5px !important;
            padding: 2px 7px !important;
            border: 1px solid var(--gm-success-border) !important;
            border-radius: 4px !important;
            background: var(--gm-success-bg) !important;
            color: var(--gm-success-text) !important;
            font-family: var(--gm-font) !important;
            font-size: 11px !important;
            font-weight: 700 !important;
            line-height: 1.4 !important;
            white-space: nowrap !important;
        }
        html body .gm-release-tag::before {
            content: "" !important;
            width: 5px !important;
            height: 5px !important;
            border-radius: 50% !important;
            background: var(--gm-success) !important;
            flex: none !important;
        }

        /* ---------- 表格 / 标签页 ---------- */
        html body .gm-course-table,
        html body .gm-timetable,
        html body .gm-sch-table,
        html body .gm-textbook-table {
            border-color: var(--gm-border) !important;
            font-family: var(--gm-font) !important;
        }
        html body .gm-course-table th {
            background: var(--gm-fill) !important;
            color: var(--gm-text-3) !important;
            font-size: 11.5px !important;
            font-weight: 700 !important;
            border-bottom: 1px solid var(--gm-border) !important;
            padding: 10px 12px !important;
        }
        html body .gm-course-table td {
            color: var(--gm-text-2) !important;
            font-size: 13px !important;
            border-bottom: 1px solid var(--gm-border-light) !important;
            padding: 11px 12px !important;
        }
        html body .gm-course-table tr:hover td {
            background: var(--gm-fill) !important;
        }
        html body .gm-timetable {
            border: 1px solid var(--gm-border) !important;
        }
        html body .gm-timetable th {
            background: var(--gm-fill) !important;
            color: var(--gm-text-3) !important;
            font-size: 11px !important;
            font-weight: 700 !important;
            border-color: var(--gm-border) !important;
        }
        html body .gm-timetable td {
            border-color: var(--gm-border) !important;
        }
        html body .gm-tt-name {
            font-size: 12px !important;
            font-weight: 700 !important;
        }
        html body .gm-tt-info {
            color: var(--gm-text-3) !important;
            font-size: 11px !important;
        }
        html body .gm-tt-tag {
            border-radius: 4px !important;
            font-size: 10.5px !important;
        }
        html body .gm-tt-course-block {
            border-radius: var(--gm-radius-sm) !important;
        }
        html body .gm-tt-footer {
            color: var(--gm-text-3) !important;
            font-size: 11.5px !important;
        }
        html body .gm-tab-item {
            font-family: var(--gm-font) !important;
            font-size: 13px !important;
            font-weight: 600 !important;
            color: var(--gm-text-3) !important;
            border-bottom-width: 2px !important;
            padding: 12px 10px 10px !important;
        }
        html body .gm-tab-item.active {
            color: var(--gm-primary-text) !important;
            border-bottom-color: var(--gm-primary) !important;
        }
        html body .gm-tabs {
            gap: 4px !important;
        }

        /* ---------- 悬浮球与菜单 ---------- */
        html body .gm-float-ball {
            width: 48px !important;
            height: 48px !important;
            background: var(--gm-primary) !important;
            box-shadow: 0 10px 24px rgba(64, 158, 255, .32) !important;
            transition: transform .18s ease, background .18s ease, box-shadow .18s ease !important;
        }
        html body .gm-float-ball:hover {
            background: var(--gm-primary-hover) !important;
            transform: scale(1.07) !important;
            box-shadow: 0 14px 30px rgba(64, 158, 255, .38) !important;
        }
        html body .gm-float-menu {
            width: 240px !important;
            border-radius: var(--gm-radius-lg) !important;
            border: 1px solid var(--gm-border) !important;
            box-shadow: var(--gm-shadow-lg) !important;
            padding: 6px !important;
            font-family: var(--gm-font) !important;
        }
        html body .gm-menu-group-title {
            font-family: var(--gm-font) !important;
            font-size: 11px !important;
            color: var(--gm-text-4) !important;
            padding: 8px 10px 4px !important;
            margin-top: 4px !important;
            border-top: 1px solid var(--gm-border-light) !important;
            letter-spacing: .08em !important;
            font-weight: 700 !important;
        }
        html body .gm-menu-item {
            font-family: var(--gm-font) !important;
            padding: 9px 10px !important;
            border-radius: var(--gm-radius) !important;
            color: var(--gm-text-2) !important;
            font-size: 13px !important;
            gap: 9px !important;
            font-weight: 400 !important;
        }
        html body .gm-menu-item:hover:not(:disabled) {
            background: var(--gm-primary-light) !important;
            color: var(--gm-primary-text) !important;
        }
        html body .gm-menu-item:disabled {
            color: var(--gm-text-4) !important;
        }
        html body .gm-menu-item .gm-icon {
            color: var(--gm-text-3) !important;
        }
        html body .gm-menu-item:hover:not(:disabled) .gm-icon {
            color: var(--gm-primary-text) !important;
        }
        html body .gm-float-menu::-webkit-scrollbar,
        html body .gm-help-body::-webkit-scrollbar,
        html body .gm-fl-body::-webkit-scrollbar,
        html body .gm-ps-list-container::-webkit-scrollbar,
        html body .gm-semester-list::-webkit-scrollbar,
        html body .gpa-modal-body::-webkit-scrollbar {
            width: 6px !important;
        }
        html body .gm-float-menu::-webkit-scrollbar-thumb,
        html body .gm-help-body::-webkit-scrollbar-thumb,
        html body .gm-fl-body::-webkit-scrollbar-thumb,
        html body .gm-ps-list-container::-webkit-scrollbar-thumb,
        html body .gm-semester-list::-webkit-scrollbar-thumb,
        html body .gpa-modal-body::-webkit-scrollbar-thumb {
            background: var(--gm-border) !important;
            border-radius: 3px !important;
        }

        /* ---------- GmDialog ---------- */
        html body .gm-dialog {
            width: min(420px, 92vw) !important;
        }
        html body .gm-dialog-body {
            padding: 22px 24px 18px !important;
            gap: 14px !important;
        }
        html body .gm-dialog-title {
            font-family: var(--gm-font) !important;
            font-size: 15px !important;
            font-weight: 700 !important;
            color: var(--gm-text-1) !important;
            margin-bottom: 4px !important;
        }
        html body .gm-dialog-text {
            font-family: var(--gm-font) !important;
            font-size: 13px !important;
            color: var(--gm-text-2) !important;
            line-height: 1.8 !important;
        }
        html body .gm-dialog-icon {
            width: 38px !important;
            height: 38px !important;
            border-radius: 50% !important;
            background: var(--gm-fill) !important;
            display: inline-flex !important;
            align-items: center !important;
            justify-content: center !important;
        }
        html body .gm-dialog-footer {
            padding: 4px 24px 20px !important;
            gap: 10px !important;
        }

        /* ---------- 友情链接 ---------- */
        html body .gm-fl-modal {
            width: min(780px, 94vw) !important;
            background: #fff !important;
        }
        html body .gm-fl-header {
            background: #fff !important;
        }
        html body .gm-fl-body {
            padding: 20px 24px !important;
            background: #fff !important;
            max-height: 62vh !important;
        }
        html body .gm-fl-category {
            margin-bottom: 20px !important;
        }
        html body .gm-fl-cat-title {
            font-size: 12.5px !important;
            font-weight: 700 !important;
            color: var(--gm-text-2) !important;
            margin-bottom: 10px !important;
            gap: 7px !important;
        }
        html body .gm-fl-grid {
            gap: 10px !important;
        }
        html body .gm-fl-card {
            padding: 12px !important;
            border-radius: var(--gm-radius) !important;
            border: 1px solid var(--gm-border) !important;
            background: #fff !important;
            transition: border-color .18s ease, box-shadow .18s ease, transform .18s ease !important;
        }
        html body .gm-fl-card:hover {
            border-color: var(--gm-primary-border) !important;
            box-shadow: var(--gm-shadow-sm) !important;
            transform: translateY(-2px) !important;
        }
        html body .gm-fl-icon-box {
            width: 34px !important;
            height: 34px !important;
            border-radius: var(--gm-radius) !important;
            flex: none !important;
        }
        html body .gm-fl-name {
            font-size: 12.5px !important;
            font-weight: 700 !important;
            color: var(--gm-text-1) !important;
        }
        html body .gm-fl-desc {
            font-size: 11px !important;
            color: var(--gm-text-3) !important;
        }

        /* ---------- 课程关注 ---------- */
        html body .gm-modal-content {
            border-radius: var(--gm-radius-lg) !important;
            border: 1px solid var(--gm-border) !important;
            box-shadow: var(--gm-shadow-lg) !important;
        }
        html body #gm-modal-overlay .gm-modal-content {
            max-width: 1040px !important;
        }
        html body .gm-course-table th:nth-child(3),
        html body .gm-course-table td:nth-child(3) {
            white-space: nowrap !important;
        }
        html body .gm-filter-bar {
            padding: 12px 20px !important;
            background: var(--gm-fill) !important;
            border-bottom: 1px solid var(--gm-border) !important;
            gap: 12px !important;
        }
        html body .gm-filter-label {
            font-size: 12px !important;
            color: var(--gm-text-3) !important;
        }
        html body .gm-filter-select {
            border-radius: var(--gm-radius-sm) !important;
        }
        html body #gm-total-credits {
            color: var(--gm-primary-text) !important;
        }
        html body .gm-btn-clear-all {
            min-height: 28px !important;
            height: auto !important;
            padding: 0 10px !important;
            border-radius: var(--gm-radius-sm) !important;
            border: 1px solid var(--gm-danger-border) !important;
            background: var(--gm-danger-bg) !important;
            color: var(--gm-danger-text) !important;
            font-size: 11.5px !important;
            font-weight: 600 !important;
        }
        html body .gm-btn-clear-all:hover {
            background: var(--gm-danger) !important;
            color: #fff !important;
        }
        html body .gm-btn-unfollow {
            min-height: 28px !important;
            height: auto !important;
            padding: 0 9px !important;
            border-radius: var(--gm-radius-sm) !important;
            border: 1px solid var(--gm-danger-border) !important;
            background: var(--gm-danger-bg) !important;
            color: var(--gm-danger-text) !important;
            font-size: 11.5px !important;
            font-weight: 600 !important;
        }
        html body .gm-btn-unfollow:hover {
            background: var(--gm-danger) !important;
            color: #fff !important;
        }
        html body .gm-empty-state {
            color: var(--gm-text-3) !important;
        }

        /* ---------- GPA ---------- */
        html body .gpa-report-modal {
            background: var(--gm-fill) !important;
        }
        html body .gpa-modal-body {
            padding: 22px 24px !important;
            background: var(--gm-fill) !important;
        }
        html body .gpa-modal-grid {
            gap: 16px !important;
        }
        html body .current-gpa-module,
        html body .gpa-calc-card,
        html body .gpa-report-modal details {
            border-radius: var(--gm-radius-lg) !important;
            border: 1px solid var(--gm-border) !important;
            background: #fff !important;
            box-shadow: var(--gm-shadow-sm) !important;
        }
        html body .gpa-report-modal h3 {
            font-family: var(--gm-font) !important;
            font-size: 14px !important;
            font-weight: 700 !important;
            color: var(--gm-text-1) !important;
        }
        html body .gpa-report-modal p,
        html body .gpa-report-modal li {
            font-family: var(--gm-font) !important;
            font-size: 12.5px !important;
            color: var(--gm-text-2) !important;
            line-height: 1.8 !important;
        }
        html body .gpa-report-modal strong {
            color: var(--gm-text-1);
            font-weight: 600;
        }
        html body .gpa-report-modal summary {
            padding: 13px 16px !important;
            font-size: 13px !important;
            color: var(--gm-text-1) !important;
        }
        html body .gpa-report-modal summary:hover {
            background: var(--gm-primary-light) !important;
            color: var(--gm-primary-text) !important;
        }
        html body .gpa-report-modal details[open] > summary {
            background: var(--gm-fill) !important;
            border-bottom-color: var(--gm-border) !important;
        }
        html body .gpa-calc-card h4 {
            font-size: 14px !important;
            color: var(--gm-text-1) !important;
        }
        html body .input-group label {
            font-size: 12px !important;
            color: var(--gm-text-2) !important;
            font-weight: 600 !important;
        }
        html body .input-group input {
            background: #fff !important;
        }
        html body .calculate-btn {
            min-height: 34px !important;
            padding: 0 14px !important;
            font-size: 13px !important;
        }
        html body .prediction-result {
            border-radius: var(--gm-radius) !important;
            background: var(--gm-success-bg) !important;
            border: 1px solid var(--gm-success-border) !important;
            color: var(--gm-success-text) !important;
            font-size: 13px !important;
            font-weight: 700 !important;
        }
        html body .gpa-report-modal .tooltip-q {
            width: 16px !important;
            height: 16px !important;
            font-size: 11px !important;
            background: var(--gm-text-4) !important;
        }
        html body .gpa-report-modal .tooltip-q:hover::after {
            border-radius: var(--gm-radius) !important;
            background: var(--gm-text-1) !important;
            box-shadow: var(--gm-shadow) !important;
            font-size: 11.5px !important;
        }
        html body .disclaimer {
            font-size: 11.5px !important;
            color: var(--gm-text-4) !important;
            border-top-color: var(--gm-border) !important;
        }

        /* ---------- 导出学期面板 ---------- */
        html body .semester-checkbox-container {
            width: min(460px, 92vw) !important;
            padding: 22px 24px !important;
        }
        html body .semester-checkbox-container h3 {
            font-family: var(--gm-font) !important;
            font-size: 16px !important;
            font-weight: 700 !important;
            color: var(--gm-text-1) !important;
            margin-bottom: 16px !important;
        }
        html body .gm-semester-list {
            gap: 8px !important;
            max-height: 270px !important;
            margin-bottom: 16px !important;
        }
        html body .gm-semester-item {
            padding: 9px 11px !important;
            border-radius: var(--gm-radius) !important;
            border: 1px solid var(--gm-border) !important;
            background: #fff !important;
            color: var(--gm-text-2) !important;
            font-size: 12.5px !important;
        }
        html body .gm-semester-item:hover {
            border-color: var(--gm-primary-border) !important;
            background: var(--gm-primary-light) !important;
            color: var(--gm-primary-text) !important;
        }
        html body .gm-semester-item input[type="checkbox"] {
            width: 16px !important;
            height: 16px !important;
            margin-right: 9px !important;
        }
        html body .button-container {
            border-top: 1px solid var(--gm-border) !important;
            padding-top: 14px !important;
            gap: 8px !important;
        }

        /* ---------- 人员检索 ---------- */
        html body .gm-ps-left {
            width: 340px !important;
            background: var(--gm-fill) !important;
            border-right: 1px solid var(--gm-border) !important;
        }
        html body .gm-ps-search-box {
            padding: 14px !important;
            background: #fff !important;
            border-bottom: 1px solid var(--gm-border) !important;
            box-shadow: none !important;
        }
        html body .gm-ps-input {
            height: 36px !important;
            border-radius: var(--gm-radius) !important;
        }
        html body .gm-ps-btn {
            height: 36px !important;
            min-height: 36px !important;
            border-radius: var(--gm-radius) !important;
        }
        html body .gm-ps-list-item {
            padding: 13px 18px !important;
            border-bottom: 1px solid var(--gm-border-light) !important;
            background: #fff !important;
        }
        html body .gm-ps-list-item:hover {
            background: var(--gm-primary-light) !important;
        }
        html body .gm-ps-list-item.active {
            background: var(--gm-primary-light) !important;
        }
        html body .gm-ps-list-item.active::before {
            width: 3px !important;
            background: var(--gm-primary) !important;
        }
        html body .gm-ps-item-name {
            font-size: 13.5px !important;
            font-weight: 700 !important;
            color: var(--gm-text-1) !important;
        }
        html body .gm-ps-item-id {
            font-family: var(--gm-mono) !important;
            font-size: 11.5px !important;
            color: var(--gm-text-3) !important;
        }
        html body .gm-ps-card {
            width: min(360px, 100%) !important;
            border-radius: var(--gm-radius-lg) !important;
            border: 1px solid var(--gm-border) !important;
            box-shadow: var(--gm-shadow) !important;
        }
        html body .gm-ps-card-header {
            height: 92px !important;
            background: linear-gradient(135deg, var(--gm-primary), #60a5fa) !important;
        }
        html body .gm-ps-avatar {
            width: 76px !important;
            height: 76px !important;
            border: 4px solid #fff !important;
            color: var(--gm-primary-text) !important;
            font-size: 30px !important;
            box-shadow: var(--gm-shadow) !important;
        }
        html body .gm-ps-card-body {
            padding: 46px 24px 22px !important;
        }
        html body .gm-ps-info-label {
            font-size: 11.5px !important;
            color: var(--gm-text-3) !important;
        }
        html body .gm-ps-info-value {
            font-size: 13px !important;
            color: var(--gm-text-1) !important;
            font-weight: 600 !important;
        }
        html body .gm-ps-info-row,
        html body .gm-ps-info-grid {
            border-bottom-color: var(--gm-border) !important;
        }
        html body .gm-ps-loader {
            color: var(--gm-text-3) !important;
            font-size: 13px !important;
        }

        /* ---------- 自动评教 ---------- */
        html body .gm-eval-body {
            padding: 16px 20px !important;
            background: var(--gm-fill) !important;
        }
        html body .gm-course-group {
            border-radius: var(--gm-radius) !important;
            border: 1px solid var(--gm-border) !important;
            box-shadow: var(--gm-shadow-sm) !important;
        }
        html body .gm-course-header {
            padding: 9px 14px !important;
            background: var(--gm-primary-light) !important;
            border-bottom: 1px solid var(--gm-border) !important;
            color: var(--gm-primary-text) !important;
            font-size: 12.5px !important;
            font-weight: 700 !important;
        }
        html body .gm-course-status-tag {
            font-weight: 600 !important;
        }
        html body .gm-teacher-row {
            padding: 10px 14px !important;
            border-bottom: 1px solid var(--gm-border-light) !important;
        }
        html body .gm-t-name {
            font-size: 12.5px !important;
            color: var(--gm-text-2) !important;
        }
        html body .gm-done-badge {
            border-radius: 4px !important;
            color: var(--gm-success-text) !important;
            border-color: var(--gm-success) !important;
        }
        html body .gm-status-pending {
            color: var(--gm-text-3) !important;
        }
        html body .gm-status-running {
            color: var(--gm-primary-text) !important;
        }
        html body .gm-status-success {
            color: var(--gm-success-text) !important;
        }
        html body .gm-status-error {
            color: var(--gm-danger-text) !important;
        }
        html body .gm-modal-footer {
            padding: 14px 24px !important;
            border-top: 1px solid var(--gm-border) !important;
            background: #fff !important;
        }
        html body .gm-row-done {
            background: #fcfcfd !important;
        }

        /* ---------- 使用说明 ---------- */
        html body .gm-help-modal {
            width: min(720px, 92vw) !important;
        }
        html body .gm-help-body {
            background: #fff !important;
        }
        html body .gm-help-section {
            padding: 16px 24px !important;
            border-bottom: 1px solid var(--gm-border-light) !important;
            background: #fff !important;
            margin: 0 !important;
        }
        html body .gm-help-sec-title {
            font-size: 14px !important;
            font-weight: 700 !important;
            color: var(--gm-text-1) !important;
            padding-left: 10px !important;
            border-left: 4px solid var(--gm-primary) !important;
            margin-bottom: 10px !important;
        }
        html body .gm-help-step {
            font-size: 12.5px !important;
            color: var(--gm-text-2) !important;
            line-height: 1.75 !important;
            margin-bottom: 8px !important;
        }
        html body .gm-help-step::before {
            color: var(--gm-text-4) !important;
        }

        /* ---------- 课表悬浮窗 ---------- */
        html body .gm-schedule-box {
            padding: 14px !important;
            border-radius: var(--gm-radius-lg) !important;
            border: 1px solid var(--gm-border) !important;
            box-shadow: var(--gm-shadow) !important;
            background: #fff !important;
            font-family: var(--gm-font) !important;
        }
        html body .gm-status-bar {
            background: var(--gm-warning-bg) !important;
            border: 1px solid var(--gm-warning-border) !important;
            border-radius: var(--gm-radius) !important;
            color: var(--gm-warning-text) !important;
            padding: 8px 10px !important;
            font-size: 12px !important;
        }
        html body .gm-sch-table {
            border: 1px solid var(--gm-border) !important;
            color: var(--gm-text-2) !important;
            font-size: 11.5px !important;
        }
        html body .gm-sch-table th {
            background: var(--gm-fill) !important;
            color: var(--gm-text-3) !important;
            font-size: 11px !important;
            font-weight: 700 !important;
            border-color: var(--gm-border) !important;
        }
        html body .gm-sch-table td {
            border-color: var(--gm-border) !important;
            color: var(--gm-text-2) !important;
            font-size: 11.5px !important;
        }
        html body .gm-sch-highlight {
            color: var(--gm-primary-text) !important;
            font-weight: 700 !important;
        }
        html body .gm-schedule-footer {
            color: var(--gm-text-3) !important;
            font-size: 11.5px !important;
        }
        html body .gm-schedule-btn {
            min-height: 30px !important;
            height: auto !important;
            padding: 0 12px !important;
            border-radius: var(--gm-radius) !important;
            font-size: 12px !important;
            font-weight: 600 !important;
        }
        html body .gm-btn-go {
            background: var(--gm-primary) !important;
        }
        html body .gm-btn-close {
            background: var(--gm-danger) !important;
        }

        /* ---------- 教材清单 ---------- */
        html body .gm-textbook-wrapper {
            padding: 18px !important;
            box-shadow: var(--gm-shadow-sm) !important;
        }
        html body .gm-textbook-title {
            font-size: 14px !important;
            font-weight: 700 !important;
            color: var(--gm-text-1) !important;
            border-left: 4px solid var(--gm-primary) !important;
            padding-left: 10px !important;
            margin-bottom: 12px !important;
        }
        html body .gm-textbook-table {
            font-size: 12.5px !important;
        }
        html body .gm-textbook-table th {
            background: var(--gm-fill) !important;
            color: var(--gm-text-3) !important;
            font-size: 11.5px !important;
            font-weight: 700 !important;
            border-color: var(--gm-border) !important;
        }
        html body .gm-textbook-table td {
            border-color: var(--gm-border-light) !important;
            color: var(--gm-text-2) !important;
        }
        html body .gm-textbook-table tr:hover {
            background: var(--gm-fill) !important;
        }
        html body .gm-textbook-empty {
            color: var(--gm-text-3) !important;
            font-size: 12.5px !important;
        }
        html body .gm-textbook-course,
        html body .gm-textbook-course a {
            color: var(--gm-primary-text) !important;
        }
        html body .gm-online-cell {
            background: var(--gm-warning-bg) !important;
            color: var(--gm-warning-text) !important;
        }
        html body .gm-online-btn {
            border-radius: var(--gm-radius-sm) !important;
            font-size: 11.5px !important;
            padding: 3px 10px !important;
            background: var(--gm-primary) !important;
        }

        /* ---------- 画像趋势 ---------- */
        html body .gm-trend-toggle-btn {
            border-radius: var(--gm-radius) !important;
            border: 1px solid var(--gm-border) !important;
            background: #fff !important;
            color: var(--gm-text-2) !important;
            font-size: 12px !important;
            font-weight: 600 !important;
        }
        html body .gm-trend-toggle-btn:hover,
        html body .gm-trend-toggle-btn.active {
            background: var(--gm-primary-light) !important;
            border-color: var(--gm-primary-border) !important;
            color: var(--gm-primary-text) !important;
        }
        html body .gm-trend-card {
            border-radius: var(--gm-radius) !important;
            border: 1px solid var(--gm-border) !important;
            background: var(--gm-fill) !important;
        }
        html body .gm-trend-card:hover {
            background: #fff !important;
            box-shadow: var(--gm-shadow-sm) !important;
        }
        html body .gm-tc-title {
            font-size: 12.5px !important;
            font-weight: 700 !important;
            color: var(--gm-text-1) !important;
        }
        html body .gm-tc-row {
            font-size: 11.5px !important;
            color: var(--gm-text-3) !important;
        }
        html body .gm-tc-val {
            font-family: var(--gm-mono) !important;
            font-size: 12px !important;
            color: var(--gm-text-1) !important;
        }
        html body .gm-tooltip-trigger:hover::after {
            border-radius: var(--gm-radius) !important;
            background: var(--gm-text-1) !important;
            font-size: 11.5px !important;
        }

        /* ---------- Toast ---------- */
        html body .gm-arc-toast {
            padding: 10px 16px !important;
            border-radius: var(--gm-radius) !important;
            box-shadow: var(--gm-shadow) !important;
            font-family: var(--gm-font) !important;
            font-size: 12.5px !important;
            border-left-width: 3px !important;
        }

        /* ---------- 选课助手：折叠按钮 / 教师链接 ---------- */
        html body #gm-collapse-toggle-btn {
            box-shadow: none !important;
            background: #fff !important;
            border: 1px solid var(--gm-border) !important;
            color: var(--gm-primary-text) !important;
            font-weight: 600 !important;
            font-size: 13px !important;
        }
        html body #gm-collapse-toggle-btn:hover {
            background: var(--gm-primary-light) !important;
            border-color: var(--gm-primary-border) !important;
        }
        html body .gm-teacher-link {
            color: var(--gm-primary-text) !important;
        }
        html body .gm-teacher-link:hover {
            color: var(--gm-primary-hover) !important;
            background: var(--gm-primary-light) !important;
            border-radius: 4px !important;
        }

        /* ---------- 选课助手：学期下拉 ---------- */
        html body .gm-lesson-panel {
            width: 200px !important;
            padding: 0 !important;
        }
        html body .gm-lesson-panel__header {
            padding: 10px 12px !important;
            background: var(--gm-fill) !important;
            border-bottom: 1px solid var(--gm-border) !important;
            border-radius: var(--gm-radius-lg) var(--gm-radius-lg) 0 0 !important;
            font-family: var(--gm-font) !important;
            font-size: 12.5px !important;
            font-weight: 700 !important;
            color: var(--gm-text-1) !important;
            display: flex !important;
            align-items: center !important;
            justify-content: space-between !important;
            cursor: move !important;
            user-select: none !important;
        }
        html body .gm-lesson-panel__body {
            padding: 12px !important;
        }
        html body .gm-semester-dd {
            position: relative !important;
            margin-bottom: 10px !important;
        }
        html body .gm-semester-dd__trigger {
            display: flex !important;
            align-items: center !important;
            justify-content: space-between !important;
            cursor: pointer !important;
            white-space: nowrap !important;
            overflow: hidden !important;
            text-overflow: ellipsis !important;
            position: relative !important;
        }
        html body .gm-semester-dd__chevron {
            position: absolute !important;
            right: 10px !important;
            top: 50% !important;
            transform: translateY(-50%) !important;
            pointer-events: none !important;
            color: var(--gm-text-3) !important;
        }
        html body .gm-semester-dd__list {
            position: absolute !important;
            top: calc(100% + 4px) !important;
            left: 0 !important;
            width: 100% !important;
            box-sizing: border-box !important;
            overflow-y: auto !important;
            background: #fff !important;
            border: 1px solid var(--gm-border) !important;
            border-radius: var(--gm-radius) !important;
            box-shadow: var(--gm-shadow) !important;
            padding: 4px !important;
            z-index: 10 !important;
        }
        html body .gm-semester-dd__item {
            height: 32px !important;
            padding: 7px 8px !important;
            border-radius: var(--gm-radius-sm) !important;
            font-family: var(--gm-font) !important;
            font-size: 12px !important;
            color: var(--gm-text-2) !important;
            line-height: 18px !important;
            cursor: pointer !important;
            white-space: nowrap !important;
            overflow: hidden !important;
            text-overflow: ellipsis !important;
            box-sizing: border-box !important;
        }
        html body .gm-semester-dd__item:hover {
            background: var(--gm-fill) !important;
        }
        html body .gm-semester-dd__item.is-selected {
            background: var(--gm-primary-light) !important;
            color: var(--gm-primary-text) !important;
            font-weight: 600 !important;
        }
        html body .gm-lesson-panel__sync {
            width: 100% !important;
            min-height: 36px !important;
            height: auto !important;
            border: 0 !important;
            border-radius: var(--gm-radius) !important;
            background: var(--gm-primary) !important;
            color: #fff !important;
            font-family: var(--gm-font) !important;
            font-size: 13px !important;
            font-weight: 600 !important;
            cursor: pointer !important;
        }
        html body .gm-lesson-panel__sync:hover {
            background: var(--gm-primary-hover) !important;
        }

        /* ---------- 同步弹窗 / 存档 ---------- */
        html body .gm-sync-modal-card {
            width: min(440px, 94vw) !important;
            padding: 0 !important;
        }
        html body .gm-sync-modal-header {
            padding: 16px 24px !important;
            border-bottom: 1px solid var(--gm-border) !important;
            display: flex !important;
            align-items: center !important;
            justify-content: space-between !important;
            margin: 0 !important;
        }
        html body .gm-sync-modal-title {
            display: flex !important;
            align-items: center !important;
            gap: 8px !important;
        }
        html body .gm-sync-modal-body {
            padding: 20px 24px 22px !important;
        }
        html body .gm-sync-actions {
            display: flex !important;
            flex-direction: column !important;
            gap: 10px !important;
        }
        html body .gm-sync-action {
            display: flex !important;
            align-items: center !important;
            gap: 12px !important;
            width: 100% !important;
            padding: 12px 14px !important;
            border: 1px solid var(--gm-border) !important;
            border-radius: var(--gm-radius) !important;
            background: var(--gm-fill) !important;
            text-align: left !important;
            transition: border-color .18s ease, background .18s ease, transform .18s ease !important;
        }
        html body .gm-sync-action:hover {
            border-color: var(--gm-primary-border) !important;
            background: var(--gm-primary-light) !important;
            transform: translateY(-1px) !important;
        }
        html body .gm-sync-action b {
            display: block !important;
            font-family: var(--gm-font) !important;
            font-size: 13px !important;
            color: var(--gm-text-1) !important;
        }
        html body .gm-sync-action small {
            display: block !important;
            margin-top: 2px !important;
            font-family: var(--gm-font) !important;
            font-size: 11.5px !important;
            color: var(--gm-text-3) !important;
        }
        html body .gm-sync-action__icon {
            width: 36px !important;
            height: 36px !important;
            border-radius: var(--gm-radius) !important;
            background: #fff !important;
            box-shadow: var(--gm-shadow-sm) !important;
            display: inline-flex !important;
            align-items: center !important;
            justify-content: center !important;
            color: var(--gm-primary-text) !important;
            flex: none !important;
        }
        html body .gm-sync-action__icon--orange {
            color: var(--gm-warning-text) !important;
        }
        html body .gm-sync-archive {
            margin-top: 18px !important;
            padding-top: 16px !important;
            border-top: 1px solid var(--gm-border) !important;
        }
        html body .gm-sync-archive__title {
            display: flex !important;
            align-items: center !important;
            gap: 7px !important;
            font-family: var(--gm-font) !important;
            font-size: 13px !important;
            font-weight: 700 !important;
            color: var(--gm-text-1) !important;
        }
        html body .gm-archive-list {
            display: flex !important;
            flex-direction: column !important;
            gap: 8px !important;
            margin-top: 10px !important;
        }
        html body .gm-archive-slot {
            border: 1px solid var(--gm-border) !important;
            border-radius: var(--gm-radius) !important;
            background: var(--gm-fill) !important;
            padding: 10px 12px !important;
            transition: border-color .18s ease, background .18s ease !important;
        }
        html body .gm-archive-slot:hover {
            border-color: var(--gm-primary-border) !important;
        }
        html body .gm-archive-slot.is-active {
            border-color: var(--gm-primary-border) !important;
            background: var(--gm-primary-light) !important;
        }
        html body .gm-archive-slot.is-empty {
            border-style: dashed !important;
            background: var(--gm-fill) !important;
            display: flex !important;
            align-items: center !important;
            gap: 10px !important;
            color: var(--gm-text-3) !important;
            font-size: 12px !important;
        }
        html body .gm-archive-slot__top {
            display: flex !important;
            align-items: center !important;
            gap: 8px !important;
        }
        html body .gm-archive-slot__icon {
            color: var(--gm-primary-text) !important;
            display: inline-flex !important;
            flex: none !important;
        }
        html body .gm-archive-slot.is-empty .gm-archive-slot__icon {
            color: var(--gm-text-4) !important;
        }
        html body .gm-archive-slot__name {
            font-family: var(--gm-font) !important;
            font-size: 12.5px !important;
            font-weight: 700 !important;
            color: var(--gm-text-1) !important;
            min-width: 0 !important;
            overflow: hidden !important;
            text-overflow: ellipsis !important;
            white-space: nowrap !important;
        }
        html body .gm-archive-slot__meta {
            margin-top: 5px !important;
            font-size: 11.5px !important;
            color: var(--gm-text-3) !important;
        }
        html body .gm-archive-actions {
            margin-left: auto !important;
            display: flex !important;
            gap: 2px !important;
            flex: none !important;
        }
        html body .gm-arc-btn {
            width: 28px !important;
            height: 28px !important;
            border-radius: var(--gm-radius-sm) !important;
            border: 0 !important;
            background: transparent !important;
            display: inline-flex !important;
            align-items: center !important;
            justify-content: center !important;
            padding: 0 !important;
            cursor: pointer !important;
        }
        html body .gm-arc-btn:hover {
            background: var(--gm-primary-light) !important;
        }
        html body .gm-arc-btn--primary { color: var(--gm-primary-text) !important; }
        html body .gm-arc-btn--success { color: var(--gm-success-text) !important; }
        html body .gm-arc-btn--success:hover { background: var(--gm-success-bg) !important; }
        html body .gm-arc-btn--warning { color: var(--gm-warning-text) !important; }
        html body .gm-arc-btn--warning:hover { background: var(--gm-warning-bg) !important; }
        html body .gm-arc-btn--danger { color: var(--gm-danger-text) !important; }
        html body .gm-arc-btn--danger:hover { background: var(--gm-danger-bg) !important; }
        html body .gm-arc-btn[disabled] { opacity: .38 !important; cursor: not-allowed !important; background: transparent !important; }
        `;
        const parent = document.head || document.documentElement;
        if (parent) parent.appendChild(style);
    }
    injectUnifiedDesignSystem();

    // 悬浮球 UI 变量
    let floatBall = null;
    let floatMenu = null;
    let menuExportBtn = null;
    let menuGpaBtn = null;
    let menuSyncBtn = null;
    let menuFollowBtn = null;

    // 功能UI变量
    let semesterCheckboxContainer = null;
    let isDataReady = false;
    let isBackgroundSyncing = false;
    // 版本检查已发现新版本（悬浮球红点与「数据未就绪」徽章共用，需独立记忆避免被数据就绪回调隐藏）
    let isUpdateAvailable = false;

    // --- 配置管理 ---
    const ConfigManager = {
        get enableExport() { return true; }, // 基础功能始终开启
        get enableGpaReport() { return true; }, // 基础功能始终开启

        get enablePortraitEnhancement() { return GM_getValue('enablePortraitEnhancement', true); },
        set enablePortraitEnhancement(val) { GM_setValue('enablePortraitEnhancement', val); },

        get enableCourseWatch() { return GM_getValue('enableCourseWatch', true); },
        set enableCourseWatch(val) { GM_setValue('enableCourseWatch', val); }
    };

    // --- 统一弹窗工具 GmDialog ---
    const GmDialog = (() => {
        const ICONS = {
            success: '<svg viewBox="0 0 24 24" width="22" height="22" stroke="#67C23A" stroke-width="2" fill="none" stroke-linecap="round" stroke-linejoin="round"><path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"></path><polyline points="22 4 12 14.01 9 11.01"></polyline></svg>',
            warn: '<svg viewBox="0 0 24 24" width="22" height="22" stroke="#E6A23C" stroke-width="2" fill="none" stroke-linecap="round" stroke-linejoin="round"><path d="M10.29 3.86L1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0z"></path><line x1="12" y1="9" x2="12" y2="13"></line><line x1="12" y1="17" x2="12.01" y2="17"></line></svg>',
            error: '<svg viewBox="0 0 24 24" width="22" height="22" stroke="#F56C6C" stroke-width="2" fill="none" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="10"></circle><line x1="15" y1="9" x2="9" y2="15"></line><line x1="9" y1="9" x2="15" y2="15"></line></svg>',
            info: '<svg viewBox="0 0 24 24" width="22" height="22" stroke="#409EFF" stroke-width="2" fill="none" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="10"></circle><line x1="12" y1="16" x2="12" y2="12"></line><line x1="12" y1="8" x2="12.01" y2="8"></line></svg>'
        };
        const THEME = { success: '#67C23A', warn: '#E6A23C', error: '#F56C6C', info: '#409EFF' };

        const ensureStyle = () => {
            if (document.getElementById('gm-dialog-style')) return;
            const style = getOrCreateStyleEl('gm-dialog-style');
            style.textContent = `
            .gm-dialog-overlay { position: fixed; top: 0; left: 0; width: 100%; height: 100%; background: rgba(0,0,0,0.55); backdrop-filter: blur(3px); z-index: 200000; display: flex; align-items: center; justify-content: center; animation: gmDialogFadeIn 0.2s ease-out; transition: opacity 0.18s; }
            @keyframes gmDialogFadeIn { from { opacity: 0; } to { opacity: 1; } }
            .gm-dialog { background: #fff; border-radius: 12px; width: 420px; max-width: 90%; box-shadow: 0 12px 32px rgba(0,0,0,0.25); font-family: "Helvetica Neue", Helvetica, "PingFang SC", "Microsoft YaHei", Arial, sans-serif; }
            .gm-dialog-body { display: flex; gap: 14px; padding: 22px 24px 18px; align-items: flex-start; }
            .gm-dialog-icon { flex-shrink: 0; margin-top: 2px; }
            .gm-dialog-text { font-size: 14px; color: #303133; line-height: 1.7; word-break: break-word; }
            .gm-dialog-title { font-size: 16px; font-weight: 700; margin-bottom: 6px; }
            .gm-dialog-footer { display: flex; justify-content: flex-end; gap: 10px; padding: 4px 24px 20px; }
            .gm-dialog-btn { padding: 8px 18px; border-radius: 6px; font-size: 14px; cursor: pointer; border: 1px solid #dcdfe6; background: #fff; color: #606266; transition: all 0.2s; }
            .gm-dialog-btn:hover { border-color: #c0c4cc; color: #303133; }
            .gm-dialog-btn-primary { color: #fff; border: none; }
            .gm-dialog-btn-primary:hover { opacity: 0.85; color: #fff; }
        `;
            document.head.appendChild(style);
        };

        function open(opts, isConfirm) {
            opts = typeof opts === 'string' ? { message: opts } : (opts || {});
            const type = ICONS[opts.type] ? opts.type : 'info';
            ensureStyle();

            return new Promise(resolve => {
                const overlay = document.createElement('div');
                overlay.className = 'gm-dialog-overlay';
                const msgHtml = esc(opts.message || '').replace(/\n/g, '<br>');
                const titleHtml = opts.title ? `<div class="gm-dialog-title">${esc(opts.title)}</div>` : '';
                overlay.innerHTML = `
                <div class="gm-dialog">
                    <div class="gm-dialog-body">
                        <span class="gm-dialog-icon">${ICONS[type]}</span>
                        <div class="gm-dialog-text">${titleHtml}${msgHtml}</div>
                    </div>
                    <div class="gm-dialog-footer">
                        ${isConfirm ? `<button class="gm-dialog-btn gm-dialog-cancel">${esc(opts.cancelText || '取消')}</button>` : ''}
                        <button class="gm-dialog-btn gm-dialog-btn-primary" style="background:${THEME[type]};">${esc(opts.confirmText || '确定')}</button>
                    </div>
                </div>`;
                document.body.appendChild(overlay);

                const close = result => {
                    overlay.style.opacity = '0';
                    setTimeout(() => overlay.remove(), 180);
                    resolve(result);
                };
                overlay.querySelector('.gm-dialog-btn-primary').onclick = () => close(true);
                const cancelBtn = overlay.querySelector('.gm-dialog-cancel');
                if (cancelBtn) cancelBtn.onclick = () => close(false);
                overlay.onclick = e => { if (e.target === overlay) close(!isConfirm); };
            });
        }

        return {
            alert: opts => open(opts, false),
            confirm: opts => open(opts, true)
        };
    })();

    // --- 常用链接导航 ---
    const FriendlyLinks = {
        categories: [
            {
                title: "友情链接快捷访问",
                titleIcon: '<svg viewBox="0 0 24 24" width="18" height="18" stroke="#409EFF" stroke-width="2" fill="none" stroke-linecap="round" stroke-linejoin="round"><path d="M10 13a5 5 0 0 0 7.54.54l3-3a5 5 0 0 0-7.07-7.07l-1.72 1.71"></path><path d="M14 11a5 5 0 0 0-7.54-.54l-3 3a5 5 0 0 0 7.07 7.07l1.71-1.71"></path></svg>',
                links: [
                    { name: "湖畔资料", url: "http://nwpushare.cn", desc: "西工大课程资料共享平台", bg: "#ecf5ff", color: "#409EFF", svg: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M12 6.5C10.3 5.2 8 4.5 5.5 4.5H5c-.8 0-1.5.7-1.5 1.5v6.5c0 .8.7 1.5 1.5 1.5h.5c2.5 0 4.8.7 6.5 2 1.7-1.3 4-2 6.5-2h.5c.8 0 1.5-.7 1.5-1.5V6c0-.8-.7-1.5-1.5-1.5h-.5C16 4.5 13.7 5.2 12 6.5z"></path><line x1="12" y1="6.5" x2="12" y2="16"></line><path d="M3.5 19.5c1.4 1 2.9 1 4.3 0s2.9-1 4.3 0 2.9 1 4.3 0 2.9-1 4.3 0"></path></svg>' },
                    { name: "瓜兵速成指南", url: "https://nwpumanual.angine.tech/", desc: "新生入学问题指南", bg: "#fdf6ec", color: "#E6A23C", svg: '<img src="data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAEAAAABACAYAAACqaXHeAAAP6klEQVR42u2be5TV1XXHP/uc333NnTsvZoYZYAYYBURUEIPiI8EopaZ5ra6mpjapSVOjMUGWiQpp7Vos2rQ2XSGmWXkRXUlXQ+wyts1qXsZQISgK8hgkiMr7ISAMMxeYx339fuec/nHuDIyoMDgkxrDX+i3uupd7z29/z3fv/d37/AbO23k7b+ftvP3hmpzuP9y+5IrYyCAzPbSu618+vXLHOw0A9WYfLllyRaxOVT2C1msQ2fiLZz73Zed+1PAHA8BeVf2JRFJ/pFgKmdg8uvJ9V141H1Nc63Z/5wvu4/c0vqNDYOH3Z9WERm1A0VYZT/KZP34/tVXVHrL/2A/ruw4Q6B8SuX+Xb//rS+84BpQidXcsrtuiyHDT5e+itroGlIPVXbCuCwI9Gq3nE7DWzZu/1M1bcN07hgH3f/ePxjsVtRtnayaNGsMnb5iDCNAXwdd2QmcJdPmrIqA1GGMQnsTJN6hLPi6LFkW/twywKvpCEFc1gdLMvnQaohQEAquycLh4wnkA5yCKADRKz0H4Cdn8U27uglvdp+7L/N4xYMGS2a1am02RMzXTWtu4ZdaNgIXjITy4E7pDUKepnlp7ZhizFeF7FMMfypIHD/xeMEBpc2ssrmq0aK6eNNlDFAisOwrZ0umdBzDGs0JkEir4MvHYBjdvwYNu3j2XvK0Z8PnFM1OpqlS7U+6isSMauX3OnyBKoGD87ne8hv5nvIoq54koh8jPcPY71FWslEWL7NuKAcnqipmiZVIUWaaNa0OCwDv8cu+psT8UcxaiEKACpW4GWU42v8LNve8Wd8896bcNAOLch3Ugkk4kmDR6DFgL1sHGYz7ZvVXrT5jOgVLvQetHKOq17q75X3CfvbfpdwrA7UuuiDm4IQwNY0bUU1OZARwcLcGuHGg1vCsb4y+lLkYHi9Gq3d01/yvuzvmTfycANOqqNhE3MYosbSObfcxq8c53h6cRzYBSoLT/dyhmjQ8PkWZ0cA8Ba91d83/g5t577W8DgGBgQ7CXx2JBwjporW/wNBVge99pEpxgrGVd9giv5HOMTCSZUVNHSukhAmH9JVKJ1h/HcIubN38ZTr51LoVVcCI+1TQEKuJxGqqqPQBFA/vzb1r68sZw/9bf8MuOQxjncMDSaVdxZV29d+hs84SIRumbsPYmunJr3NwF3yQR/VgWL+47JyEguMnGWKoq0qSTSRAH3dGb136laD+e5aeHDxIowQFTq6qZnKl+60mzHwhrQeuZaPWDc5EwFcDChQsVSIu1jpp0Gh0Env5HQ8jbN+4ZnWNUMsWYZAXWwXvq6vnKxdPIBLHhqRqDEmZ0csLc6Obe96D73IJLhyUEci3L0jESI6y1VKXSXriI9btv7BsnNucYX1HJ0mlXkQ1LTKmsQint6/65MGv8pVQTOribKLrD3TX/l4h7mNqK/5NFi0pnxYDAJSscVDoHlcmk33EBeiJwp6fq6FQFl1bXokTOnfOvTZhRCEIKrf8U1M/J5le7uxbc5e6+u3nIANjApATiAMlY7MSnOXPm8frahKcUiBANZyi8WZ5QajpafR2TaHfz5n/DzbtnxpkzQJTGoX0jV6a7A6KzvHml2Hgsy7wX2vmzDc/wcm+P7w6H1D+ID8Uz1RUDDRhNqOBzOP2Mmzd/mdvyrb9y7pGRpy+DZzUzfn3nV3V1cPeWjfREESLQ44XOqYlRpBxi7pTfWPrKbp7KHuFTLW3MHEpJ7dcTkYsxqXo2E6tmUwoPudLSJ5ZtXPfzn7VvfP7rdz61fRADImcNggEw/QsJEB+6/DXW8s09O+gzBi3CxHSGKVXVvqd4jfP7cn0UbHQq0s6xrPMwy44c5s7N61l+5NDQFKYrt/Cz68AZULrJ2vATLx7Y+6N0TK394sM3Xj0IAF2yRSAEKIYnCa7K4KxYkFSavDWktObu8ROpCGKDd1kpnjxymA9vWMU/bH8R+zr0r48nSGhF0Vr+bfc28l4cndkNhBbePQIuyvgwdrCsfSMHOrpIJIIahZk7CIBYXOWdIy8CuWLB36sDamNDjl0twj9OuoT7L5jM96deyQ0NTYPpK0I+ivj6nu3kjeHxjkPs6uvx8X6SfbBxFIJgnKOzVCJvzlAJhw4uSMNNI30J15pDnUd4duuLBFpjIgvIuxYunJUcAODFPUdyCMdFhO58rpzVHdTHIaFOXwpfQ98xqQo+PX4Cl1XXvm7sdhULvFrMExdFyVmOhqXBTLOW6+sb+fsLJzO1qoa54y6kNp44vbgyDqoD+OhoSCmwgjWGx9vXkS+VEJH+SGzMtcTqBgB4bNGLJXHukNKKY319WBOBBeringXWDb08GfOGiasmFqcmiJMzhowOGJOqONU55/hYy3h+NP1qbhkzDjmd8w7fvd48GsakPPUDzXNbX+LlA/uJaT3wuw4qAh1WDeoFnMhOrYRjfX30FQr+zZSG8WmP7DDW7spYnHvbJnJFTS3zxk+gOZl6/d3tB/B0FcCVd/9DTTC1Gkqe+q92HuFXm9pRSg2Qujz40daqxOAyaPmNiNBXLNDRfYxMZaVXdZdmYE12mAWMZc7IUcxpaPKx/1bVY2Thxka4vt6/FqFYLPLjNc+QKxZRSpNKWFrqSuw+kiS0oLS4wd1gzK6LImuNNew+fKg81nZwYSU0J4eXBSfv6lt1vmRhZh18cKS/R+cT7c/XP8eejsMopUknLFe29dEyor9VEKMjVxwEQCzNC864VwKt2XbwADYK/Y9VaLimbvgBGA4rWZhWDX8+6oSoCgJWbdnMc9u3EgQBgYLLWnNUpQx7jiQIjQAUStb2DQJg0c0re4EnY3HNgWwXBzq7/FgstDCjFlpTZy+Nz5XzU6rgY2MgpnyABwGbd+7g8Y3r0UrhrNBQFVKfiegtaPZ1xftblO5UKtV96sGI8IiNnAtNxPqd27xC62fBB5p8lnVvE+cvqYJPtEBSe3YGAdv37+O/1qzCOoeIH9BUpUxZfjiU4N93roMdV/WeAkCX6XnKGLchkYjxm7276Tya9RI0tDAlAzfU+9e/K3NllTezFj7Z6qtU2fldB/fzn0//mmIY+ra8f2RX8hUgFbe0jiiCCErUzkXlQ5lBAHz3jg0hIotVuRqs2LLphBKMHLxvJLyrxu/Ab9si5xn4oWb4yxbfp/Tv/Cv7WLpyOX2FIvqknkGJo7MnoBj692rSBqXAOTa94dlglzn+32HJrkol42zctZOt+/ZCEHj0lcBfjPGJ57cFgnV+rZYU3DEO5jR6zVBOeJt2bGfpU8vpLRS5aFTI1NbcoK86JwNRa6wQhQ7RbvUbAvDdOzaE4ux9xriCA/533RqO93T7ULDOI39rC1w3wtPRunPreDqADzbB3DaYmPZrisIBKza18+gzKymGEZmkML6hSH1VRCxwGCtkkpZLWnIkAodzcOh4HGPsgUBs+0Dv8nprP/3Tvfuve//YRCIVzDrel+Pw0SxTWscRBIG/MS0+CVUGsDsHeePZIcPguHGe7pmYB/mjoz3jBDCADujN5/if1at4+qUtiCicU1w4skBjdUi+pNhzJEFzbcj08TmqKwyIsONwkgPdKTD20S/9zcrHTjsQicfcP4cFc11FKv7ebQcP8OjTv+aj776eRCxWPsBwMKseJqThiQ7Y3A1F68EZChj9MtY6X85GJWF6DUyvhvrECSYoBUrY+spefrLuOTqOHyUZ02jlaK4tMa6hCEC2NyCdsFzWkkMrT/ttrybZdSSJUjbUWj90xjOfBd+eNU7H1fIgUOPz+RITR43mI9e+h5pMVf9TISdOjPflYP0xf5LcVfJUdeUV5DVPlDhOfLdCQ0MCLkzD5IzXGyntWWDdwCM4x7q7Wb75edbv2M6o2gJj6z3VA+VIxhwijkKoWL29krH1RdoaPSC7OxJs3l9BKqmIStFj/3Tbr28e0tDriw+99wqt+YnWalS+UKIhU80HZsxkcuvYwcPQQPzN5iI4VISDBX+kfjyEQjlXKIGkgkzgO82RCWhM+I4zXs4x/XJW+XlgXy7H+h1beealF8n29jGlJWRiUwHnBPFynsgIx3Kalw+mONoXMHNCL/WZCGdhzY5KjuVjgDumnJv5pdtWbB3y1G/BwzdeHoh9NIjpCYWCr7OXt13ArClTaayt8zy29sR5oi6DIf2pmJPYwODP7EmsEPEHrDiy3cfZuGsnG3Zuo7O7G6U0mZTj2om9aHGIQHdes6sjQXde0VfUA/l45oV91GdCwkh4dnuGvAkwkb39gdtWPHTWY8/7H37vWAcPxxJ6dlgyhGFEOpnkktZxTL9gAq0j6tHxeNmZk643mvgOTH6l/3kiunM59nQc4oV9e9nx6kF6C3m0UiilMFYYW1/kstYc1gmlSFi9vZLjOY0S0NqvVZmwXHlBL6m4pbegWbu3mnzOfu2BT6/4/Fue+y78/qxkZPTfobhPa0mGJUNkDYEOaKqppW1kM+MbR9JYXUMmlSIRBIjS5VXE08A5jDEUw5CeQp7O7m4OZrvY19nBq0ez9OTzGAuBVmXZKgOiZkZbHw1VISKw43CSLa+kqK+KaBlRIp2wCI6KhCUZsygFOzvTbN4dezh10N65aNHKaLgG3/ztQ9dfqwN5QCn1bmcdxjiMtVhrESUkY3HSiSTpZJJULE4sCFBKiIyhFEXkikV/lYqUwhBX1u5KFEoJY+pKNFSFOCf0FRWhEeoqDY0Z73whFFZty1BfGXFpS45AO5wTlHJERigazavH4mw/GPvK1kzDFx+7+TEznJN/z4aFF8fDsU23CtwbBDLJn1/agTG/xeGsK4e/Qw1ivvir/JqTUsTU1hyj60qD3uv/nnNe2z+/t4JcUXPdpB5iZZGjBA4dj7E3W0F3jmxvTt331c8++b23/Lj8aRPkktnVQWA/Hg/sZytSXOwchKFDgJi2BGXJ3p3ThFYGFuyXqSfbRaPyTGgqYGw/7f1gt2SEYqjo7AnY1xWnt6CZ1FzgolF5jBW0gp5SjGe3pSmV+EWkzILFt618YVj+XuBM7YnnL0urUvwDAe6vBTcrmVBJ8OGBc+zrTLBpX8rvpoMRmYgRlRHWCWEk1KQNzTWlkxoZ2J+Ns7sjQTESIiNEVlDiK801E3qpr7IoLYQl6/Z2Jla270l9dfFnlv+MITTtci5k/HMbZ0wJrX2/sW6OODdVBap+Z0cF2w8lUFgmNecZO6KIiBsoFA6w/TuvHAeycTbtrcAh/SLQP7Mowtj6EhMbey0i2xz8ysJj11+x4VkRhtyhnRMATrYVz81o6uwLpjy/J3m5ieTyhupw4oy23jHOSY0IySAQ1EnSWYBiKDz9ciW5omCtsyL0OeuyIrIvGeOlqyf2bErFzVpVTGy55po1ec7ln8ycgxNteWL1zFolbkRCmwZRts5ErlaJaD+chmxvYNfvSvc6a7tVTGfFuq6gynaVx3bn7bydt/N23s7bMNn/A5XlWTqBYXJAAAAAAElFTkSuQmCC" alt="">' },
                    { name: "翱翔助手", url: "https://github.com/Lorcas-Zephyr/aoxiang-assistant", desc: "更高效的手机App", bg: "#ecf5ff", color: "#2F80ED", svg: '<svg viewBox="0 0 108 108" fill="none"><path fill="#2F80ED" d="M54,6A48,48 0,1 0,54 102A48,48 0,1 0,54 6"/><path fill="#FFFFFF" d="M31,31h46v8H31zM31,47h31v8H31zM31,63h22v8H31z"/><path fill="#B9DEFF" d="M67,48h10v23H67z"/></svg>' },
                    { name: "瓜拼白车", url: "https://nwpu.codepix.top", desc: "瓜大人自己的在线免费拼车平台", bg: "#e8f1fb", color: "#004EA2", svg: '<svg viewBox="0 0 256 256" fill="none" xmlns="http://www.w3.org/2000/svg"><rect width="256" height="256" fill="#004EA2"/><path d="M143.659 195.051C148.947 192.473 153.174 188.348 155.687 183.31L176.915 141.608L200.441 95.3858C202.932 90.6315 203.735 85.2618 202.733 80.0506C201.8 75.1446 199.287 70.6155 195.528 67.0632C191.683 63.3665 186.71 60.8584 181.284 59.8794C175.858 58.9005 170.243 59.4982 165.2 61.5914L114.885 82.0646L69.5273 100.383C64.2418 102.471 59.7973 106.051 56.8206 110.617C53.8438 115.184 52.4851 120.506 52.9361 125.835C53.3639 130.987 55.5264 135.886 59.1164 139.834C63.0327 144.226 68.4897 147.191 74.5231 148.206L99.8811 153.006L104.225 177.225C105.167 182.884 108.227 188.056 112.865 191.827C117.038 195.302 122.288 197.457 127.857 197.981C133.325 198.535 138.842 197.512 143.659 195.051ZM121.612 174.313L115.98 143.401C115.767 142.223 115.173 141.132 114.279 140.273C113.384 139.414 112.231 138.828 110.971 138.592L78.01 132.352C76.0257 131.914 74.2418 130.9 72.9166 129.457C71.5914 128.014 70.7936 126.217 70.6389 124.326C70.4842 122.436 70.9805 120.549 72.056 118.94C73.1315 117.331 74.7304 116.083 76.621 115.376L121.98 96.976L172.427 76.5043C174.188 75.7668 176.151 75.5533 178.049 75.8928C179.947 76.2324 181.686 77.1086 183.03 78.4016C184.374 79.6946 185.257 81.3413 185.558 83.1169C185.859 84.8924 185.563 86.7101 184.711 88.3215L161.14 134.544L139.913 176.246C139.14 178.046 137.754 179.561 135.968 180.559C134.182 181.557 132.093 181.984 130.019 181.774C127.946 181.564 126.002 180.73 124.484 179.397C122.966 178.065 121.957 176.308 121.612 174.395L121.612 174.313Z" fill="white"/><path d="M124.046 120.104L141.851 103.924C143.533 102.412 145.793 101.584 148.135 101.623C150.476 101.661 152.706 102.563 154.335 104.13C155.964 105.697 156.858 107.801 156.821 109.979C156.783 112.156 155.817 114.229 154.135 115.741L136.33 131.921C134.648 133.434 132.387 134.262 130.046 134.223C127.705 134.184 125.474 133.282 123.845 131.715C122.217 130.148 121.323 128.044 121.36 125.867C121.398 123.689 122.364 121.616 124.046 120.104Z" fill="white"/></svg>' }
                ]
            },
            {
                title: "校内网站快捷访问",
                titleIcon: '<svg viewBox="0 0 24 24" width="18" height="18" stroke="#67C23A" stroke-width="2" fill="none" stroke-linecap="round" stroke-linejoin="round"><path d="M3 9l9-7 9 7v11a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2z"></path><polyline points="9 22 9 12 15 12 15 22"></polyline></svg>',
                links: [
                    { name: "西北工业大学 WebVPN", url: "https://webvpn.nwpu.edu.cn/", desc: "校外免客户端访问校园网环境", bg: "#f0f9eb", color: "#67C23A", svg: '<img src="data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAACAAAAAgCAYAAABzenr0AAAJk0lEQVR42p2XaZBVxRXH/6f73nffe7MxzKjIJgwCrkFQoiWIJlHcoqUp8iFqTJUbKo6iKTFVmpjEBRNF1LiAqFFTGpeRjMaIxBhHiQpKgmYQBUSQRGZEmGHmvXtv7ycfiFE0UuD51FXdff6/7urT5xzCbhh7C4gIH324mtZu3QwvJMAEGTxGNzdjyPCxABEDABHtks+vXLXgvdUYGjGtFKVGBRqhQPsysHckxRjFaLDMBALA29dHBJSFMM75dUS0oQT+UDJ/sBebzS9yjb1DbMHAloN3DnDvu+swQLB4n+NhKTBZEx3jiSZ4ohEMNDCRZPB/t306+rwLBjGBiFkwUmLuEsydMfOrCXPHHuCVWyBNq1RoHjPmM4C7V61HrYBcZ2hcCjrTEJ3iiUYGIPpUjsBBADkBFRCyAhEEwQKAZ8QmMIhQYkZdAEoMkvy5U0rwJ3HgF4vM9w0SYUnFs7nmkFGg65avRVFS88dezMpJnO1AewUAAggSvFkSd8aEtyShswBeHxNvCUSVwbU1qI0iz2Dq1VZ8nOVIBNUYz00KYhjAB9qAcQ50sGcaGoACtoP0lZnvHRbTdYa5ElWchIM4pmr5CsuQAiGLiN8oEC8sCX6pkdy6G+vH5i9tfRUTjpy8Sw/LLWl/fWZpAvZDGm9GYUjmcVgeaJplOsGAGkCY3st4ikHLoqoK0BQSG4RkMEv42YNid9s7JqqeURtw2sT9AAATdiNaoqNO+3Roec3yDbM+rtswLrFPd6rSdamnqywg+8hJBhDBBJAgBB8QGOyDX71RieoTJ47GE7sSmsx4a8MGDGluEgLEITDv2VD/2SsfcxgA4KxF79t6GXpDEGAwLDwYQETKgSLRG6wwAVQQCA2K5Q4iM/6wEgVCU87i2ERgze1La1bcfZLBcNmDGU+vRlH4FivMFQi8roHV3Vf/sVPfcMpnIXfX4hVYogQU68hRhIjYFCIoBiDIOLjMfBR0qHjNcIr32ZZLXPP0SgDAhY93Qli/xzYtbq9a+VifpjsvG9/T8F5XioUbayGNGdWrxPyeDDP6Fc/SLFuU5R0O8EGfx2PxSAQdhgbNCMpvc9X0E19NEUUMgHkzTNgSGE0GaDkVb4n3KsPCOQ++BVameRslcyzRGQCDKHwShLe5j1EkO7KP43kGdCzAiMgvYeE2+bDj/5bmARfny0p9VDc6UEAA/yuK0MMkIGqDxR6U9bD2670KsCrst0wNH9CTCnCumyo6mqO1ONMrJmlcRznYWb1ZnEnjWvqUnKcUHRtUgDB2YSnYmV191Dd3UrwDQK4lUh3v6TS3eBXA2nXOfX9CWuQAEcdltK+u02T9m8EEOI1RuZajvTINmUtu0Vae5Q0LaV1Hic30SpXWlKxt6VfRPUqLqV4HJmPbClZfUqnypkfOPQQ05MAdAJSRMIYOdhqDgwmBjV/2/cZ/4qYfHApx7ff2x761GsL6V1mH3Gk0aEUnp7b0a2Pl2V6zENZ2lFhPr1awJjFqZF8ezcu3i4OMfSq2qjX4pGvvhAZe9du3B935yGK69fdLAAAXzF2C8/WjZDWd4A0nrH2XMHZ5wZntIQsAZTgIH5bDiJUhyIkGciYIRYBFRK4jic0FFRWvrRF2ZL+P51mI48AMSbYtkqZV+ag7ztIDP6wU7iAKjT2FxvO0wz8AoD+PcT9O38cBUwMYgvzS+lJ1nQoJsP3HBQaVgRdWDdgqrHsumACvUecVYjK2I/FqelqVa0tOtfTn0Tyt5NSgAsiYtsipVp3Jbsr1/qmO7s1y8e08x0FKhVGZYlz0q79ia1aGNvJEp2lUMN6Rdc9s7BlgmuN8x2w47WevQQocUtWF5wKLvaVwHUWppxtfWFOWdkTmC/MsR8eDAUmuLZG61bqoWwSzvxWlBQHRJCKEgnQPDCibHxtH/Sb3EBQGZlzzjOdokiTXWZeoE5hp02O/mPzZDQBAbWQxorGvUzr7HBm9VNr0QpXJNUWvR1SyeJ7K5fFeMWBMW+Tz1jwV3SHL91MmuddqOSno4IUzD5RFflU1RX+kFLor9chV4XSrxESvA2Dd42+u3WdTSbr/Xw9Mu2Ixikk0NlcGLtSsjkUYmdpkng3RVACQZNviSLU6V+gmp8Z4Ub4vUHwUAB9L+2BtSV9pvextnzMFp1zyMgSF4XkoPRNYjpPCv1dTUCcFpvXtc4/+n6b4PEDbrcfD9W5enaX1q8n4Ef3V+G6dy6lBM2DMk8Jnrboqu0OWjjG2uMCZ6Kig2QtnHihSdmWeUm/7nCmYNuPPaIq3RrmKL3NKjAs6eGHt/I7lQ9Y3FN0OISq+mFwaGwfg8JZNkVLipzqXJ3gdwFa3SZ9ealLZHfJ0jDWlBc5EU7wOgay+v8DZrLwqe5+96xhc/vNF+GBTPbp6Gk82uTjHawZZ+5cE6uFJYzfiodnf2TlAbW0dbnm42QnmF2BcBdYsjEJ+qcqibpdlo42tme9MPCXoEMjpBwtIf6Iz2vb8A9/Cpde0YcXKAvYoVQ5SeXS90zSAje2SQd/Q0xv3jNrLfCmbfgng5qsn4+TJBnUF3Z7I9MIY+UyTii7O0rHOlhd4HR3jdQhw+qGIK1eqnHpH1rwC5kvwzqo6FIXeO8viW6wSB7HxWnoz+6aL2pc0lTPcdv2Ju14Vc88QoPEjDD94EUYMdQe4UJ4fKJlMYC+EeSgupFeGEPcMK72CR5+YjUnHtYNg6lg2zfWcnAswosgsaKyvzDQuzhY/efLuleUAcO7Fz6KuRiV/f7tugdKlHwIEIvtaFPVNY5Zdr794KgDgiClPgmDqPDX/klGcAVAspVlULmfnOi+6Xln03a/UEDsDKCUJbru5rJOI7ydv13od4JQ4yFTLP4q9Lh45+WkcOvFJkLcN1gy80evkEq8Rs9VLI6SXp/1R1+C613daUcmdTb6x9Hc477xT0fb4oA+HDetd5a04IlgxmIOYbLUoeVNZKkXc5Gz9rd4n57BDBNbLBSrnK1W/qhYd+NPiG79eZ/R5O+20BWhvPxqHf/Pdw1WW3MMhHg+wFcItZEYjIz4ODCJh3hRR9fzgBr6d1PwNbyy79uu3Zl+0o464FVs2H4hC0jc++Nr5wccTQbS9NSOGlObFQjG9TGDwOxZtWPH2bbvkV+4qwMZ/L8aYkUNR7RvfXVuz7WVvsI/XYjQHdkLoR+NCtdXqZN03xq/H8y9cs8slPGE37YbZLXj4jtmoG2CbdVqeBaJtxdrsN8xJZfmqabvrbvcBPrXDD7gTdfVJBIC19n7Jigu/lp//ABqUkmhRCl8DAAAAAElFTkSuQmCC" alt="">' },
                    { name: "NOJ", url: "https://noj.nwpu.edu.cn/cpbox/", desc: "程序设计基础理论/实验 NOJ平台", bg: "#ecf5ff", color: "#409EFF", svg: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polyline points="4 17 10 11 4 5"></polyline><line x1="12" y1="19" x2="20" y2="19"></line></svg>' }
                ]
            },
            {
                title: "网课网站快速访问",
                titleIcon: '<svg viewBox="0 0 24 24" width="18" height="18" stroke="#F56C6C" stroke-width="2" fill="none" stroke-linecap="round" stroke-linejoin="round"><rect x="2" y="3" width="20" height="14" rx="2" ry="2"></rect><line x1="8" y1="21" x2="16" y2="21"></line><line x1="12" y1="17" x2="12" y2="21"></line></svg>',
                links: [
                    { name: "学堂在线", url: "https://bknwpu.yuketang.cn/", desc: "西工大专属雨课堂在线学习主页", bg: "#fef0f0", color: "#F56C6C", svg: '<img src="data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAEAAAABACAYAAACqaXHeAAAEqklEQVR42u2by28bRRjAfzuz602a96tOghM1L6IGYh5qiVohpWoV1JYD4oAEBw4IESQuOSAhwQnEPwAXxInSYyVKK5CQKBQONIc+adM0VRuURnm7TZpH3Ti2d3c4JI4a4riKg5M4u59kybPSePz9vm++me/bGU0ppQCiFpzvi3NxwGJi1sF22FEiBVQWCdrqdY60GJj64nNNKaXuzzl8/VuU0NwO03oN8RcKujpMdhcKRNTCVcoDhJYMHrVAnO+Lu0r5JyGc74sjLg5YuFUuDliIiVnHtQAmZh30taJ9VbHgnTYfuwsEPSMWpy7FtvXKULJLo8kvafILGv2SQIlgJqI48VeUvjE7aR/bAX2tH/zwkEmgRABweK9Bf8jh8r3tM112+TT2VkuCAUlzlaQ8X0sK5YN2k49PzeOsYbw1ASSUT0hFwcoBfDq8e9AkGJCE5hTfX4gyNpNZF3mmRBAMSII1kobdEqE9vU9BjkZRrsb0Y7U+AE+T460+DjQsdq+v0Og8ZPL52cj/rnRlkWBfnaStTqeqWKS3CUoBKm0AgVKR0mM2IkW5GgcbdfbX6dSWiYx6lb5d5rSmQUu1pL1Z58UaHSE2Z9yMAmiplhxtNbAd+KUnTn9odTQ2JHQ8Z9DebFCWr206+IwBqCjQ6OrIQS5ZsrlS8unpeWbnVwaj9141eaV+6xwxY47W5JfLyidWjbry1cPt27O1s1Bkck6vepbsDwh2JoBsEQ9Auh0dpf7TdhmAK4Mrl7Srg9mZVqcdgi8NWNgOS7mAw++34u4CkLB6tlreC4IeAA/AxgGYOjxbKSnN07IWgL6RQsUnx3IozNVwFJzsjtLdb7nHA462GhTmLlpeaPDWfp+7pkC+qaVsuz4IxpPMBstxEYDeUZuHT1RiQ3MOd8bt7AmCI9POikLn/bn1ZTvzMcWXP0U40KCjgO5+i5idRavAt39GebvNR0WBRu+Izd9D64/wjxYU57Z5jqCnem/21bkFbyPkAdjhktGS7P46nWOtBo6Cn6/HuDFsZ5cHlORpBEoEucb6NznVxYLOdpPaMsGecsFHh3O25MVH2h7w5ss+Xn/BACBmwxdnI+s6SlNXIVaUxqWA2lLBVNjODg947Xlj+btPQrBGuisIGjJ121sFPAAeAA+AB8AD4AHIClmw0gAwOOmsSo+zUQYnHcILav1b4W/+WOCNl3yUF2j0jdpcH7KzSvHpecXNYZsz12Lp5QIPHytOXIiSquSVqr3ZMhVW3JmwFz/jNpNhldl0+NfeOMEaSb6poYAfr8Y2VWHLgbsTNjdHbHqG7bTvPKQNYHTa4bMfIjT6BeMzigePMh8jInHFjSGbq4MWfeMO0bja2oLIfEzRk+EiR0Lpy/csbo3ZWHaWVITGk5wcDyUprc9F1PIrtoQo4PaYzYW7FteHMltOzxiAgQcOp6/EOB40UMCZa/Gkx+lPdkfpbDcxDY2psKK7P073PxZT4c0Jqtr734W3/HyXT4c8c+0z/ez00+IxC2LW1thBSBdnA1KAqCxyL4HKIoFoq9ddC6CtXkccaTHwF7rPC/yFgiMtBsLUoavDdBWExOVpU1+6PQ7uvT7/LwgCuorx9strAAAAAElFTkSuQmCC" alt="">' },
                    { name: "中国大学 MOOC 学校云", url: "https://www.icourse163.org/spoc/schoolcloud/index.htm", desc: "国家精品课程与SPOC在线学习平台", bg: "#f0f9eb", color: "#67C23A", svg: '<img src="data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAEAAAABACAYAAACqaXHeAAADsklEQVR42u2bz4sTZxjHP5NsdrO/ctB2ha7rquAqIq0KRYoWx0OtF0Hn5smTeB3/DC9FQRSZQwstCC0ulhbR7aXSWh0P7mDLuv4imqw/kk2su9ndzI9MpoeZWLXGLWqS2cz7hZdMCMzM85nnfd53nueJ5HkeUVaMiEsAEAAEgGhLOjS6zQYeAlPByAIZYBK4A2Q0RffaGcBixpnAbeAvwACuA+Oaoj+NCoB6SgN/AL8DvwETS9FT3gXAqyoCvwAXgDFN0R9HDcCrug78AJzTFH0yigBe1J/At8B3YfOMZgGoqQpcBE4C5zVFr0YNwIu6D5wATmuKPteKG9h87FeplQBqegYcB45piv7sPRsYA4aAEWADsCr4vhIYBAbDAKCmGeAo8JWm6NZbGLsM2ApsCcYmYB2QDOsUqKd7wBFN0X96k+sCG4HPgR3AdmD1UosBi+lH4LCm6LnA6I+A3cAe4Atg2VIPgovvrEqrZzPFXZcqsQ2DgXu/d3WEzeg58wPST7ZQnEnh2Faqszu2tyfVuOuFAoBHjEx+K9npYcoLtfhnNeXaLQXguN3cym4n/zSF69pNM7rlACynn5vZHRT+7sKrVgC7ZQ+hqQBct4uJrEyukMTzXKDS8unXNADp3KekH67EdR3ADU3QbTiAeXM543d3Ui7bgBO6pbahAKYKnzD5YA1e1YYoZoWnCmuCIIdIiwsAAoAAIAAIAAKAACAACAACQPhehj7GLxQM46eW1+Pn00eARLsDkOr9cGh0WzwAsRm/0FD7/PD/nvzq5D5KpXd7GersXkFPaqT5AN4AZhi/EFErSGyqN5XaEsBrgKSAzwAZ+DLwkugAeA2QgQDEbv3W/gOzs048UgB4uYb3jevMHXSsIo5VxK3Mhw5Aw3OC8UQf8UQfyb5hqq6FbeZxrAKu05KWgNamxWPxLpK9QyR7h6i6JrY5jW3mqFbK0QDwMozkcxgVp4RdzuGY03heQ3KINn6v4x38BtAHwciGojbYkeinI9GP178WxypgLzym4sy+7enywDh+U6cRHN81VNkNfXVYkmJ0JgfoTA7gVuaxFh4hSdJiTVc3+Ldh87KhypklXR5/Hjw7eulJreM/hWT/iV4ALgFXDFUutVV/QB2XvgiMAWOGKufbukEikAGMAj8DhqHKXls3SAS6BnwPjBqqnG77ZTDQBHAGOGOo8r2o7ANmAqO/NlT5WpR2gpeBU8BZQ5XNqPQHzAGngZOGKt8IZUZI/HMUkRUWAAQAASC6+gcsCmJ9lr5ltQAAAABJRU5ErkJggg==" alt="">' },
                    { name: "超星学习通", url: "http://nwpu.mooc.chaoxing.com", desc: "西工大超星泛雅网络教学平台", bg: "#ecf5ff", color: "#409EFF", svg: '<img src="data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAACAklEQVR42pWSz0vTcRzGX89ncy4RkRgSoQZiC+zULepg0KlL0MWmm+JXRcw6hUTs5GFEhMfox3JT4tv6FyI6dKmgU0SESIQHEYkuGc45v/u+O6gLUwOf6+d53jw/PmIH5nntuKqPlNOs/4YDYCNDl1BtBmscULG4CBCtv6rag9GLWa+Npt8RaRhTfn4BwMbS5zHlsNplDIgFSWARwNUPOJUBaOAs6BtBeMK8zAUbSb8i5ANQQZFzANSsuiuL/mtTT198BYYBbHTwE2atRHRRz/z32zHSe/hR/gtbBt7uig9C1K6Pd9G0tQZbHAU2cKOVcLMpCtVHbPCRZl4f6UJYzgLdDhFHxDkyXAxT7O8KcnEAGx9PHGr75tDxnRXi+0uUW0XBb4Lyko1mHhJpmiFYd9vNjzSjzSyV2i3gF41uifKOj7o+P/+FjmQCcRtIUVv/jimJFIPqVVAKNEXn6TY9fv65HgSzFswy1u+lgEAFP09HdxI0BZYAi6vol1Twu1Tw8/T0BJYazmDhALJOmeclqNhdQiaAZZxyXLtSUl9fuCf/9HSUhaUMZlmkdrBZ4rqnOqF/rI2wdgdsArGClOPMqRIrK7C2OQTKYnYSpzyR2H35+VUA7Wt6cKKNrY0pTJPI1jABtICe4CIP9HL2x56vf+hknpdggyKEFY65Sc3N/TyI9wdocMOj18wYfwAAAABJRU5ErkJggg==" alt="">' },
                    { name: "智慧树 / 知到", url: "https://www.zhihuishu.com/", desc: "国内大型学分课跨校共享平台", bg: "#fdf6ec", color: "#E6A23C", svg: '<img src="data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAABw0lEQVR42qWTTWsTURSGn5nMZJKmLdVoSCsGwRpKBX+Bv8TfIK5daJciuHIruHDhMgvdCNKtrgrSSneKJEJjsc1HMzPJZO7c103SKqTx61kduOd9Fpf3ACBpXVJDUl+/pz/ZXWcSrkvq6u/pSqozsf0rDUdSH1hmBs93X1B02ty5dZ9zCN2Z4f47lHaJkjbtsEOUWD60ZgoW3elkBcMUJMGneySHr1grVbhUrLHzJeLha7A2AxsD9tRwKujHhsdvoPXtCBvcQI5DIRdQ9BfIgM3KmKj3GZp3Udz+VdDs7TM0Xbb3oXU8xOTKOE4Oo4zMGlzHYSkYk0ZNSBqY3oBo9+OZ4OXeI76HfWqXwaQJJCc4WYzv+niuz8hAJwJjLXATG3b4+mDrTJB3F1nIB1xbgWr1Kt6V22T5NeI0Ik5DyiVRK+fJF1bAv45TXMYtBUyLpOPwUCbLdHRiNDaS7FBpGunZzhM9fb+lUTpWOJQyk8gkbdnMKDw4kCR5ABdLFQDKS9OvKeB5sH5hkzSLCTyfwAPIQ64KQGl1FYC5RUrMCEkU/OLcIm2f9xp4hXlhgLdI2viPY9r4+SIbkgZ/EBxMdusAPwBB3+9eEof0vgAAAABJRU5ErkJggg==" alt="">' }
                ]
            }
        ],

        initModal: function () {
            if (document.getElementById('gm-friendly-links-modal')) return;

            const style = getOrCreateStyleEl('gm-friendly-links-style');
            style.textContent = `
            .gm-fl-modal-overlay { position: fixed; top: 0; left: 0; right: 0; bottom: 0; background: rgba(0,0,0,0.55); backdrop-filter: blur(4px); z-index: 100005; display: flex; justify-content: center; align-items: center; animation: gmLinkFadeIn 0.2s ease-out; }
            @keyframes gmLinkFadeIn { from { opacity: 0; transform: scale(0.98); } to { opacity: 1; transform: scale(1); } }
            .gm-fl-modal { background: #f5f7fa; border-radius: 12px; width: 720px; max-width: 92%; box-shadow: 0 12px 32px rgba(0,0,0,0.25); overflow: hidden; display: flex; flex-direction: column; font-family: "Helvetica Neue", Helvetica, "PingFang SC", "Microsoft YaHei", Arial, sans-serif; }
            .gm-fl-header { padding: 18px 24px; border-bottom: 1px solid #ebeef5; display: flex; justify-content: space-between; align-items: center; background: #fff; color: #303133; }
            .gm-fl-title { font-size: 19px; font-weight: 600; color: #2c3e50; display: flex; align-items: center; gap: 10px; letter-spacing: 0.5px; }
            .gm-fl-close { cursor: pointer; color: #909399; font-size: 26px; line-height: 1; user-select: none; transition: color 0.2s; }
            .gm-fl-close:hover { color: #f56c6c; }
            .gm-fl-body { padding: 24px; max-height: 65vh; overflow-y: auto; }
            .gm-fl-category { margin-bottom: 24px; }
            .gm-fl-cat-title { font-size: 15px; font-weight: bold; color: #606266; margin-bottom: 14px; display:flex; align-items:center; gap:8px;}
            .gm-fl-grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(300px, 1fr)); gap: 14px; }
            .gm-fl-card { background: #fff; border: 1px solid #e4e7ed; border-radius: 8px; padding: 14px; text-decoration: none; display: flex; align-items: center; gap: 14px; transition: all 0.25s; cursor: pointer; }
            .gm-fl-card:hover { transform: translateY(-2px); box-shadow: 0 6px 16px rgba(0,0,0,0.08); border-color: #c6e2ff; }
            .gm-fl-icon-box { width: 36px; height: 36px; border-radius: 8px; flex-shrink: 0; display: flex; align-items: center; justify-content: center; }
            .gm-fl-icon-box svg { width: 20px; height: 20px; }
            .gm-fl-icon-box img { width: 22px; height: 22px; object-fit: contain; border-radius: 4px; }
            .gm-fl-info { flex: 1; min-width: 0; }
            .gm-fl-name { font-size: 14px; font-weight: 600; color: #303133; margin-bottom: 4px; display: flex; justify-content: space-between; align-items: center; }
            .gm-fl-arrow { font-size: 12px; color: #c0c4cc; transition: transform 0.2s, color 0.2s; font-weight: bold; }
            .gm-fl-card:hover .gm-fl-arrow { transform: translateX(3px); color: #409EFF; }
            .gm-fl-desc { font-size: 12px; color: #909399; line-height: 1.4; display: -webkit-box; -webkit-line-clamp: 2; -webkit-box-orient: vertical; overflow: hidden; }
        `;
            document.head.appendChild(style);

            const overlay = document.createElement('div');
            overlay.id = 'gm-friendly-links-modal';
            overlay.className = 'gm-fl-modal-overlay';
            overlay.style.display = 'none';

            let contentHtml = this.categories.map(cat => `
            <div class="gm-fl-category">
                <div class="gm-fl-cat-title">${cat.titleIcon} ${cat.title}</div>
                <div class="gm-fl-grid">
                    ${cat.links.map(link => `
                        <a href="${link.url}" target="_blank" class="gm-fl-card">
                            <div class="gm-fl-icon-box" style="background:${link.bg}; color:${link.color};">${link.svg}</div>
                            <div class="gm-fl-info">
                                <div class="gm-fl-name">${link.name} <span class="gm-fl-arrow"><svg viewBox="0 0 24 24" width="13" height="13" stroke="currentColor" stroke-width="2.5" fill="none" stroke-linecap="round" stroke-linejoin="round"><line x1="5" y1="12" x2="19" y2="12"></line><polyline points="12 5 19 12 12 19"></polyline></svg></span></div>
                                <div class="gm-fl-desc" title="${link.desc}">${link.desc}</div>
                            </div>
                        </a>
                    `).join('')}
                </div>
            </div>
        `).join('');

            overlay.innerHTML = `
            <div class="gm-fl-modal" onclick="event.stopPropagation()">
                <div class="gm-fl-header">
                    <div class="gm-fl-title">
                        <svg viewBox="0 0 24 24" width="22" height="22" stroke="#409EFF" stroke-width="2.5" fill="none" stroke-linecap="round" stroke-linejoin="round"><polygon points="12 2 2 7 12 12 22 7 12 2"></polygon><polyline points="2 17 12 22 22 17"></polyline><polyline points="2 12 12 17 22 12"></polyline></svg>
                        常用链接访问
                    </div>
                    <span class="gm-fl-close" onclick="document.getElementById('gm-friendly-links-modal').style.display='none'">&times;</span>
                </div>
                <div class="gm-fl-body">${contentHtml}</div>
            </div>
        `;
            overlay.onclick = () => overlay.style.display = 'none';
            document.body.appendChild(overlay);
        },
        show: function () { this.initModal(); document.getElementById('gm-friendly-links-modal').style.display = 'flex'; }
    };

    // 更新源配置：Gitee 主源 + GitHub 副源
    const UPDATE_SOURCES = {
        GITEE_REPO: 'leamloli/aoxiang-jwxt-enhance',
        GITEE_BRANCH: 'master',
        GITHUB_REPO: 'zengzoxiong/aoxiang-jwxt-enhance',
        GITHUB_BRANCH: 'main',
        USER_FILE: 'aoxiang-jwxt-enhance.user.js',
        META_FILE: 'aoxiang-jwxt-enhance.meta.js'
    };

    // --- 版本检查器 ---
    const UpdateChecker = {
        check: function (auto = false) {
            const btnText = document.getElementById('gm-update-text');
            if (!auto && btnText) btnText.innerHTML = "正在检查更新...";

            const src = UPDATE_SOURCES;
            const urlsToTry = [
                `https://gitee.com/${src.GITEE_REPO}/raw/${src.GITEE_BRANCH}/${src.META_FILE}?t=` + Date.now(),
                `https://raw.githubusercontent.com/${src.GITHUB_REPO}/${src.GITHUB_BRANCH}/${src.META_FILE}?t=` + Date.now()
            ];

            let currentTry = 0;
            const requestNext = () => {
                if (currentTry >= urlsToTry.length) {
                    if (!auto && btnText) btnText.innerHTML = `网络故障，点击重试`;
                    return;
                }
                GM_xmlhttpRequest({
                    method: 'GET',
                    url: urlsToTry[currentTry],
                    timeout: 3500,
                    onload: function (res) {
                        if (res.status === 200 && res.responseText.includes('@version')) {
                            try {
                                const match = res.responseText.match(/@version\s+([^\s]+)/);
                                if (match && match[1]) {
                                    const latestVersion = match[1];
                                    const currentVersion = typeof GM_info !== 'undefined' ? GM_info.script.version : '1.0.0';

                                    if (UpdateChecker.compareVersion(latestVersion, currentVersion) > 0) {
                                        isUpdateAvailable = true;
                                        if (btnText) btnText.innerHTML = `<span style="color:#F56C6C;font-weight:bold;display:flex;align-items:center;"><svg viewBox="0 0 24 24" width="14" height="14" fill="currentColor" style="margin-right:4px;"><path d="M12 2l3.09 6.26L22 9.27l-5 4.87 1.18 6.88L12 17.77l-6.18 3.25L7 14.14 2 9.27l6.91-1.01L12 2z"/></svg>发现新版本 v${latestVersion}</span>`;
                                        const badge = document.querySelector('.gm-float-ball .gm-badge');
                                        if (badge) badge.style.display = 'block';
                                    } else {
                                        isUpdateAvailable = false;
                                        if (!auto && btnText) btnText.innerHTML = `已是最新版本 (v${currentVersion})`;
                                        else if (auto && btnText) btnText.innerHTML = `检查版本更新 (v${currentVersion})`;
                                    }
                                    return;
                                }
                            } catch (e) { }
                        }
                        currentTry++; requestNext();
                    },
                    onerror: function () { currentTry++; requestNext(); },
                    ontimeout: function () { currentTry++; requestNext(); }
                });
            };
            requestNext();
        },
        compareVersion: function (v1, v2) {
            const p1 = v1.split('.').map(Number), p2 = v2.split('.').map(Number);
            for (let i = 0; i < Math.max(p1.length, p2.length); i++) {
                if ((p1[i] || 0) > (p2[i] || 0)) return 1;
                if ((p1[i] || 0) < (p2[i] || 0)) return -1;
            }
            return 0;
        },
        openUpdatePage: function () {
            if (document.getElementById('gm-update-modal-overlay')) return;

            const overlay = document.createElement('div');
            overlay.id = 'gm-update-modal-overlay';
            overlay.style.cssText = 'position:fixed; top:0; left:0; width:100%; height:100%; background:rgba(0,0,0,0.6); z-index:999999; display:flex; align-items:center; justify-content:center; backdrop-filter:blur(3px); animation: gmFadeIn 0.2s;';

            const modal = document.createElement('div');
            modal.style.cssText = 'background:#fff; width:440px; padding:24px; border-radius:12px; box-shadow:0 10px 30px rgba(0,0,0,0.2); font-family:"Helvetica Neue", Helvetica, "PingFang SC", "Microsoft YaHei", Arial, sans-serif; text-align:center; position:relative;';

            modal.innerHTML = `
            <h3 style="margin:0 0 15px 0; color:#303133; font-size:18px; display:flex; align-items:center; justify-content:center; gap:8px;">
                <svg viewBox="0 0 24 24" width="20" height="20" stroke="currentColor" stroke-width="2" fill="none" stroke-linecap="round" stroke-linejoin="round"><polyline points="23 4 23 10 17 10"></polyline><polyline points="1 20 1 14 7 14"></polyline><path d="M3.51 9a9 9 0 0 1 14.85-3.36L23 10M1 14l4.64 4.36A9 9 0 0 0 20.49 15"></path></svg>
                脚本更新指引
            </h3>
            <p style="color:#606266; font-size:14px; margin-bottom:20px; line-height:1.6; text-align:left;">
                点击下方按钮可从<b>Gitee / GitHub</b>下载最新版脚本，<br>下载后<b>拖拽到当前浏览器的油猴页面</b>中完成更新：
            </p>
            <div style="display:flex; flex-direction:column; gap:12px;">
                <button id="gm-update-btn-proxy" style="background:#409EFF; color:#fff; border:none; padding:12px; border-radius:6px; font-size:14px; cursor:pointer; font-weight:bold; transition:all 0.2s; display:flex; align-items:center; justify-content:center; gap:6px;">
                    <svg viewBox="0 0 24 24" width="16" height="16" stroke="currentColor" stroke-width="2" fill="none" stroke-linecap="round" stroke-linejoin="round"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path><polyline points="7 10 12 15 17 10"></polyline><line x1="12" y1="15" x2="12" y2="3"></line></svg>
                    下载最新版代码
                </button>
                <div id="gm-update-tip" style="font-size:13px; color:#E6A23C; display:none; text-align:left; background:#fdf6ec; padding:10px; border-radius:6px; border:1px solid #faecd8; line-height:1.5;">
                    <div style="display:flex; align-items:center; margin-bottom:4px; font-weight:bold; color:#67C23A;">
                        <svg viewBox="0 0 24 24" width="16" height="16" stroke="#67C23A" stroke-width="2" fill="none" stroke-linecap="round" stroke-linejoin="round" style="margin-right:4px;"><path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"></path><polyline points="22 4 12 14.01 9 11.01"></polyline></svg>代码下载成功！
                    </div>
                    请在浏览器的下载列表中找到刚刚下载的<b>“aoxiang-jwxt-enhance.user.js”</b>，将其<b>拖拽到当前浏览器的油猴页面中</b>，即可完成更新！
                </div>

                <div style="display:flex; gap:10px; margin-top:5px;">
                    <button id="gm-update-btn-direct" style="flex:1; background:#f4f4f5; color:#606266; border:1px solid #dcdfe6; padding:10px; border-radius:6px; font-size:13px; cursor:pointer; transition:all 0.2s; display:flex; align-items:center; justify-content:center; gap:6px;">
                        <svg viewBox="0 0 24 24" width="14" height="14" stroke="currentColor" stroke-width="2" fill="none" stroke-linecap="round" stroke-linejoin="round"><path d="M18 13v6a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V8a2 2 0 0 1 2-2h6"></path><polyline points="15 3 21 3 21 9"></polyline><line x1="10" y1="14" x2="21" y2="3"></line></svg>
                        访问 GitHub 仓库
                    </button>
                    <button id="gm-update-btn-close" style="flex:1; background:#fff; color:#909399; border:1px solid #e4e7ed; padding:10px; border-radius:6px; font-size:13px; cursor:pointer; transition:all 0.2s;">
                        稍后更新
                    </button>
                </div>
            </div>
        `;

            overlay.appendChild(modal);
            document.body.appendChild(overlay);

            document.getElementById('gm-update-btn-close').onclick = () => overlay.remove();
            document.getElementById('gm-update-btn-direct').onclick = () => {
                window.open('https://github.com/' + UPDATE_SOURCES.GITHUB_REPO, '_blank');
                overlay.remove();
            };

            const proxyBtn = document.getElementById('gm-update-btn-proxy');
            const tipBox = document.getElementById('gm-update-tip');

            proxyBtn.onclick = () => {
                proxyBtn.innerHTML = '<svg viewBox="0 0 24 24" width="16" height="16" stroke="currentColor" stroke-width="2" fill="none" stroke-linecap="round" stroke-linejoin="round" class="gm-spin" style="margin-right:6px;"><line x1="12" y1="2" x2="12" y2="6"></line><line x1="12" y1="18" x2="12" y2="22"></line><line x1="4.93" y1="4.93" x2="7.76" y2="7.76"></line><line x1="16.24" y1="16.24" x2="19.07" y2="19.07"></line><line x1="2" y1="12" x2="6" y2="12"></line><line x1="18" y1="12" x2="22" y2="12"></line><line x1="4.93" y1="19.07" x2="7.76" y2="16.24"></line><line x1="16.24" y1="7.76" x2="19.07" y2="4.93"></line></svg>正在拉取代码，请稍候...';

                // 等待按钮旋转动画
                if (!document.getElementById('gm-spin-style')) {
                    const style = getOrCreateStyleEl('gm-spin-style');
                    style.textContent = '@keyframes gmSpin { 100% { transform: rotate(360deg); } } .gm-spin { animation: gmSpin 1s linear infinite; }';
                    document.head.appendChild(style);
                }

                proxyBtn.disabled = true;
                proxyBtn.style.opacity = '0.7';

                const src = UPDATE_SOURCES;
                const scriptUrls = [
                    `https://gitee.com/${src.GITEE_REPO}/raw/${src.GITEE_BRANCH}/${src.USER_FILE}?t=` + Date.now(),
                    `https://raw.githubusercontent.com/${src.GITHUB_REPO}/${src.GITHUB_BRANCH}/${src.USER_FILE}?t=` + Date.now()
                ];

                const downloadFrom = (index) => {
                    if (index >= scriptUrls.length) {
                        proxyBtn.innerHTML = '<svg viewBox="0 0 24 24" width="16" height="16" stroke="currentColor" stroke-width="2" fill="none" stroke-linecap="round" stroke-linejoin="round" style="margin-right:6px;"><circle cx="12" cy="12" r="10"></circle><line x1="15" y1="9" x2="9" y2="15"></line><line x1="9" y1="9" x2="15" y2="15"></line></svg>拉取失败，请检查网络后重试';
                        proxyBtn.style.background = "#F56C6C";
                        proxyBtn.disabled = false;
                        proxyBtn.style.opacity = '1';
                        return;
                    }
                    GM_xmlhttpRequest({
                        method: 'GET',
                        url: scriptUrls[index],
                        timeout: 8000,
                        onload: (res) => {
                            if (res.status === 200 && res.responseText.includes('==/UserScript==')) {
                                const blob = new Blob([res.responseText], { type: 'text/javascript' });
                                const url = URL.createObjectURL(blob);
                                const a = document.createElement('a');
                                a.href = url;
                                a.download = 'aoxiang-jwxt-enhance.user.js';
                                document.body.appendChild(a);
                                a.click();
                                document.body.removeChild(a);
                                URL.revokeObjectURL(url);

                                proxyBtn.innerHTML = '<svg viewBox="0 0 24 24" width="16" height="16" stroke="currentColor" stroke-width="2" fill="none" stroke-linecap="round" stroke-linejoin="round" style="margin-right:6px;"><path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"></path><polyline points="22 4 12 14.01 9 11.01"></polyline></svg>下载完成！请按下方提示操作';
                                proxyBtn.style.background = "#67C23A";
                                tipBox.style.display = "block";
                            } else {
                                downloadFrom(index + 1);
                            }
                        },
                        onerror: () => downloadFrom(index + 1),
                        ontimeout: () => downloadFrom(index + 1)
                    });
                };
                downloadFrom(0);
            };
        }
    };

    // --- 关注课程数据管理 ---
    const FollowManager = {
        getList() {
            try {
                return JSON.parse(GM_getValue(CONSTANTS.FOLLOWED_COURSES_KEY, '{}'));
            } catch (e) {
                Logger.error('Follow', '关注列表数据损坏，将返回空列表', e);
                return {};
            }
        },
        add(courseId, courseData) {
            const list = this.getList();
            list[courseId] = courseData;
            GM_setValue(CONSTANTS.FOLLOWED_COURSES_KEY, JSON.stringify(list));
            Logger.log('Follow', `关注课程成功: ${courseData.name}`);
        },
        remove(courseId) {
            const list = this.getList();
            delete list[courseId];
            GM_setValue(CONSTANTS.FOLLOWED_COURSES_KEY, JSON.stringify(list));
            Logger.log('Follow', `取消关注成功: ID ${courseId}`);
        },
        has(courseId) { return !!this.getList()[courseId]; }
    };

    // --- 基础数据获取与缓存 ---

    /**
     * 获取SessionStorage缓存的数据
     */
    function getCachedData() {
        const cachedData = sessionStorage.getItem(CONSTANTS.CACHE_KEY);
        if (cachedData) {
            try { return JSON.parse(cachedData); }
            catch (error) { sessionStorage.removeItem(CONSTANTS.CACHE_KEY); return null; }
        }
        return null;
    }

    /**
     * 写入数据到SessionStorage
     */
    function setCachedData(data) {
        try { sessionStorage.setItem(CONSTANTS.CACHE_KEY, JSON.stringify(data)); }
        catch (error) { Logger.error('Core', "缓存写入失败", error); }
    }

    // 模块级 StudentID 内存缓存（生命周期：当前页面关闭前持续有效）
    let cachedStudentId = null;
    // 进行中的获取请求（用于并发去重，避免同一页面多次重复请求）
    let studentIdPromise = null;

    /**
     * 获取学生ID（带内存缓存）
     * @param {Object} [options] @param {boolean} [options.force] 跳过缓存强制重新获取
     * @returns {Promise<string|null>}
     */
    async function getStudentId(options = {}) {
        if (!options.force && cachedStudentId) {
            Logger.log('Core', `复用缓存 StudentID: ${cachedStudentId}`);
            return cachedStudentId;
        }
        if (!options.force && studentIdPromise) {
            return studentIdPromise;
        }

        Logger.log('Core', "正在获取 StudentID...");

        const fetchPromise = new Promise((resolve) => {
            const courseTableUrl = "https://jwxt.nwpu.edu.cn/student/for-std/course-table";

            GM_xmlhttpRequest({
                method: "GET",
                url: courseTableUrl,
                headers: {
                    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8"
                },
                onload: function (response) {
                    if (response.status === 200) {
                        try {
                            const html = response.responseText;
                            const dataIdMatch = html.match(/var\s+dataId\s*=\s*(\d+)\s*;/);

                            if (dataIdMatch && dataIdMatch[1]) {
                                const sid = dataIdMatch[1];
                                Logger.log('Core', `解析成功，StudentID (dataId): ${sid}`);
                                cachedStudentId = sid;
                                localStorage.setItem(CONSTANTS.STUDENT_ID_LS_KEY, sid); // 兼容选课助手
                                resolve(sid);
                            } else {
                                Logger.error('Core', "HTML 中未找到 dataId 变量");
                                resolve(null);
                            }
                        } catch (e) {
                            Logger.error('Core', "HTML 解析失败", e);
                            resolve(null);
                        }
                    } else {
                        Logger.error('Core', `页面请求失败，HTTP状态码: ${response.status}`);
                        resolve(null);
                    }
                },
                onerror: (err) => {
                    Logger.error('Core', "网络请求失败", err);
                    resolve(null);
                }
            });
        });

        if (options.force) {
            return fetchPromise;
        }

        studentIdPromise = fetchPromise;
        try {
            return await fetchPromise;
        } finally {
            studentIdPromise = null;
        }
    }

    /**
     * 从门户 JWT 中同步本科生/研究生身份，供教务页选课时间提醒使用
     * @param {string} token 门户 localStorage.token
     */
    function syncIdentityFromToken(token) {
        try {
            const part = String(token || '').split('.')[1];
            if (!part) return;
            const b64 = part.replace(/-/g, '+').replace(/_/g, '/');
            const padded = b64 + '='.repeat((4 - (b64.length % 4)) % 4);
            const binary = atob(padded);
            const bytes = Uint8Array.from(binary, c => c.charCodeAt(0));
            const payload = JSON.parse(new TextDecoder('utf-8').decode(bytes));
            const name = payload.ATTR_identityTypeName || '';
            const code = payload.ATTR_identityTypeCode || '';
            let identity = '';
            if (/研究生/.test(name) || code === 'S02') identity = 'graduate';
            else if (/本科生/.test(name) || code === 'S01') identity = 'undergraduate';
            if (!identity) return;
            GM_setValue(CONSTANTS.USER_IDENTITY_KEY, identity);
            GM_setValue(CONSTANTS.USER_IDENTITY_NAME_KEY, identity === 'graduate' ? '研究生' : '本科生');
            Logger.log('Identity', '门户身份: ' + (identity === 'graduate' ? '研究生' : '本科生'));
        } catch (e) {
            Logger.warn('Identity', '解析门户身份失败: ' + (e && e.message));
        }
    }

    /**
     * 从后端抓取所有成绩数据并缓存
     */
    async function fetchAllDataAndCache(retryCount = 0, options = {}) {
        Logger.log("Initial", "开始获取并缓存所有教务数据");
        try {
            const studentId = options.studentId || await getStudentId(options); // options.studentId 为外部已确定正确的 ID
            const [gpaRes, semRes, rankRes] = await Promise.all([
                new Promise(r => GM_xmlhttpRequest({ method: "GET", url: `https://jwxt.nwpu.edu.cn/student/for-std/student-portrait/getMyGpa?studentAssoc=${studentId}`, onload: r, onerror: () => r({ status: 500 }) })),
                new Promise(r => GM_xmlhttpRequest({ method: "GET", url: `https://jwxt.nwpu.edu.cn/student/for-std/student-portrait/getMyGrades?studentAssoc=${studentId}&semesterAssoc=`, onload: r, onerror: () => r({ status: 500 }) })),
                new Promise(r => GM_xmlhttpRequest({ method: "GET", url: `https://jwxt.nwpu.edu.cn/student/for-std/student-portrait/getMyGradesByProgram?studentAssoc=${studentId}`, onload: r, onerror: () => r({ status: 500 }) }))
            ]);

            // --- 判断 ID 是否失效 ---
            // 1. 检查 HTTP 状态码是否异常（通常 401, 403, 500 代表 ID 不匹配或过期）
            // 2. 检查返回内容是否包含登录 HTML（说明 Session 失效重定向了）
            const isInvalid = (res) => {
                return res.status !== 200 ||
                    (typeof res.responseText === 'string' && res.responseText.includes('<!DOCTYPE html>'));
            };

            if (isInvalid(gpaRes) || isInvalid(semRes)) {
                if (retryCount < 1) { // 仅允许重试一次，防止死循环
                    Logger.warn("Core", "检测到请求无效，准备重试...");
                    // 显式传入的 studentId（如预取）无效时不清理主缓存
                    if (!options.studentId) {
                        localStorage.removeItem(CONSTANTS.STUDENT_ID_LS_KEY);
                        cachedStudentId = null; // 清除内存缓存，确保重试时重新获取
                    }
                    return await fetchAllDataAndCache(retryCount + 1, options);
                } else {
                    throw new Error("多次请求均无效，请检查登录状态。");
                }
            }

            const gpaData = JSON.parse(gpaRes.responseText);
            const gpaRankData = gpaData.stdGpaRankDto || { rank: null, gpa: null };

            const semesterData = JSON.parse(semRes.responseText);
            const semesters = Array.isArray(semesterData.semesters) ? semesterData.semesters.sort((a, b) => b.id - a.id) : [];
            const semesterIds = semesters.map(s => s.id);
            const semesterNames = semesters.map(s => s.nameZh);

            const classRankData = {};

            if (rankRes.status === 200) {
                try {
                    const data = JSON.parse(rankRes.responseText);
                    // 将 data?.courseItemMap 改为 data && data.courseItemMap
                    if (data && data.courseItemMap) {
                        for (const cid in data.courseItemMap) {
                            if (Object.prototype.hasOwnProperty.call(data.courseItemMap, cid)) {
                                const c = data.courseItemMap[cid];
                                // 确保数据存在再进行赋值
                                if (c && c.stdLessonRank != null) {
                                    classRankData[cid] = c.stdLessonRank + "/" + c.stdCount;
                                }
                            }
                        }
                    }
                } catch (e) {
                    Logger.error("Initial", "解析排名数据失败", e);
                }
            }

            let allGrades = [];
            if (semesterIds.length > 0) {
                const gradePromises = semesterIds.map(semesterId =>
                    new Promise(resolve => {
                        GM_xmlhttpRequest({
                            method: "GET",
                            url: `https://jwxt.nwpu.edu.cn/student/for-std/grade/sheet/info/${studentId}?semester=${semesterId}`,
                            onload: response => { if (response.status === 200) { const data = JSON.parse(response.responseText); resolve(data.semesterId2studentGrades[semesterId] || []); } else resolve([]); },
                            onerror: () => resolve([])
                        });
                    })
                );

                const allGradesArrays = await Promise.all(gradePromises);
                allGradesArrays.forEach((grades, index) => {
                    const semesterName = semesterNames[index];
                    grades.forEach(grade => {
                        allGrades.push({
                            '课程ID': grade.course.id,
                            '课程代码': grade.course.code,
                            '课程名称': grade.course.nameZh,
                            '学分': grade.course.credits,
                            '成绩': grade.gaGrade,
                            '绩点': grade.gp,
                            '教学班排名': classRankData[grade.course.id] || "无数据",
                            '学期': semesterName,
                            '是否必修': grade.course.obligatory
                        });
                    });
                });
            }

            const finalData = { gpaRankData, allGrades, semesterNames };
            setCachedData(finalData);
            Logger.log('Initial', "数据获取完成，已写入缓存");
            return finalData;
        } catch (error) {
            Logger.error("Initial", "数据获取错误", error);
            throw error;
        }
    }

    // =-=-=-=-=-=-=-=-=-=-=-=-= 1. 主页初始化与诊断 =-=-=-=-=-=-=-=-=-=-=-=-=
    /**
     * 打印脚本初始化时的详细存储状态诊断报告
     */
    function printStorageDiagnosis() {
        // 辅助函数：计算字符串大小（KB）
        const calcSize = (str) => str ? (new Blob([str]).size / 1024).toFixed(2) + ' KB' : '0 KB';
        // 辅助函数：安全解析JSON
        const safeParse = (key, isSession = false) => {
            const raw = isSession ? sessionStorage.getItem(key) : GM_getValue(key);
            try { return raw ? JSON.parse(raw) : null; } catch (e) { return null; }
        };

        try {
            console.groupCollapsed('%c[NWPU-Enhanced]脚本环境诊断报告 (点击展开)', 'background:#007bff; color:#fff; padding:4px 8px; border-radius:4px;');

            // --- 1. 基础环境与配置 ---
            const studentId = localStorage.getItem(CONSTANTS.STUDENT_ID_LS_KEY);
            const configData = {
                '脚本版本': GM_info.script.version,
                '当前学号 (LocalStorage)': studentId || '未获取 (可能导致功能失效)',
                '功能开关: 画像增强': ConfigManager.enablePortraitEnhancement ? '开启' : 'OFF',
                '功能开关: 课程关注': ConfigManager.enableCourseWatch ? '开启' : 'OFF',
                '浏览器 UserAgent': navigator.userAgent.substring(0, 50) + '...'
            };
            console.log('%c 1. 环境与配置', 'color: #007bff; font-weight: bold;');
            console.table(configData);

            // --- 2. 成绩缓存数据 (SessionStorage) ---
            const cachedData = safeParse(CONSTANTS.CACHE_KEY, true);
            const cacheRawSize = sessionStorage.getItem(CONSTANTS.CACHE_KEY);
            console.log('%c 2. 成绩缓存数据 (SessionStorage)', 'color: #007bff; font-weight: bold;');
            if (cachedData) {
                const semesterCounts = {};
                if (cachedData.allGrades) {
                    cachedData.allGrades.forEach(g => {
                        semesterCounts[g.学期] = (semesterCounts[g.学期] || 0) + 1;
                    });
                }
                console.log(`%c 数据有效 | 占用空间: ${calcSize(cacheRawSize)}`, 'color: green');
                console.table({
                    '总课程数': cachedData.allGrades ? cachedData.allGrades.length : 0,
                    '包含学期数': cachedData.semesterNames ? cachedData.semesterNames.length : 0,
                    'GPA (Rank数据)': cachedData.gpaRankData ? cachedData.gpaRankData.gpa : '无',
                    '排名': cachedData.gpaRankData ? cachedData.gpaRankData.rank : '无'
                });
                if (Object.keys(semesterCounts).length > 0) {
                    console.log('▼ 各学期课程数量分布:');
                    console.table(semesterCounts);
                }
            } else {
                console.log('%c 未检测到成绩缓存 (正常现象，稍后会自动抓取)', 'color: orange');
            }

            // --- 3. 关注课程数据 (LocalStorage) ---
            const followed = FollowManager.getList();
            const followedRaw = GM_getValue(CONSTANTS.FOLLOWED_COURSES_KEY);
            console.log('%c 3. 关注课程列表', 'color: #007bff; font-weight: bold;');
            if (Object.keys(followed).length > 0) {
                const followStats = {
                    '关注总数': Object.keys(followed).length,
                    '数据大小': calcSize(followedRaw),
                    '最近添加': Object.values(followed).sort((a, b) => new Date(b.addedTime) - new Date(a.addedTime))[0]?.name || 'N/A'
                };
                console.table(followStats);
            } else {
                console.log('关注列表为空');
            }

            // --- 4. 选课助手/后台同步数据 ---
            const bgData = safeParse(CONSTANTS.BACKGROUND_SYNC_KEY);
            const bgRaw = GM_getValue(CONSTANTS.BACKGROUND_SYNC_KEY);
            const lastSyncTime = GM_getValue(CONSTANTS.LAST_SYNC_TIME_KEY, 0);
            const historyData = safeParse(CONSTANTS.HISTORY_STORAGE_KEY);
            const historyRaw = GM_getValue(CONSTANTS.HISTORY_STORAGE_KEY);

            console.log('%c 4. 选课助手数据', 'color: #007bff; font-weight: bold;');
            console.table({
                '全校课表缓存 (条数)': bgData ? bgData.length : 0,
                '全校课表占用': calcSize(bgRaw),
                '上次全校同步时间': lastSyncTime ? new Date(lastSyncTime).toLocaleString() : '从未同步',
                '历史余量记录 (课程数)': historyData ? Object.keys(historyData).length : 0,
                '历史记录占用': calcSize(historyRaw)
            });
            console.groupEnd();
        } catch (e) {
            console.error('[NWPU-Enhanced] 诊断报告生成失败', e);
        }
    }

    async function initializeHomePageFeatures() {
        const isEcampus = location.host === 'ecampus.nwpu.edu.cn';

        // 1. UI 初始化
        if (!isEcampus) printStorageDiagnosis();
        createFloatingMenu(isEcampus); // 传入当前环境标志

        if (isEcampus) return;

        // 主页面初次加载即预取 StudentID，供本页面所有功能复用（悬浮球同步、后台静默同步等），生命周期持续到页面关闭
        getStudentId().catch(err => Logger.error('Core', 'StudentID 预取失败', err));

        initExportUI();
        initScheduleWidget();

        // 首次运行检测
        const FIRST_RUN_KEY = CONSTANTS.INTRO_SHOWN_KEY;
        if (!GM_getValue(FIRST_RUN_KEY, false)) {
            setTimeout(() => handleHelpClick(), 1500);
            GM_setValue(FIRST_RUN_KEY, true);
        }

        // 2. 设置按钮状态为“加载中”
        updateMenuButtonsState(false);

        // 3. 【延迟执行】定义繁重的数据加载任务
        const runHeavyDataFetch = async () => {
            let cachedData = getCachedData();
            if (cachedData) {
                updateMenuButtonsState(true);
                isDataReady = true;
            } else {
                try {
                    await fetchAllDataAndCache();
                    updateMenuButtonsState(true);
                    isDataReady = true;
                } catch (error) {
                    Logger.error("Initial", "后台数据加载失败", error);
                }
            }
        };

        // 4. 使用 requestIdleCallback 在浏览器空闲时执行
        if ('requestIdleCallback' in window) {
            window.requestIdleCallback(() => {
                runHeavyDataFetch();
            }, { timeout: 3000 });
        } else {
            setTimeout(runHeavyDataFetch, 1000);
        }
    }

    function createFloatingMenu(isEcampus = false) {
        if (!document.getElementById('gm-float-menu-style')) {
            const style = getOrCreateStyleEl('gm-float-menu-style');
            style.textContent = `
            /* 悬浮球样式 */
            .gm-float-ball {
                position: fixed; top: 15%; right: 20px; width: 48px; height: 48px;
                background-color: #409EFF; color: white; border-radius: 50%;
                box-shadow: 0 4px 12px rgba(0,123,255,0.4); z-index: 100001; cursor: pointer;
                display: flex; align-items: center; justify-content: center; font-size: 26px;
                user-select: none; transition: all 0.2s; touch-action: none;
            }
            .gm-float-ball:hover { transform: scale(1.08); background-color: #66b1ff; }

            /* 菜单容器 */
            .gm-float-menu {
                position: fixed; width: 230px !important; background-color: #fff; border-radius: 8px;
                box-shadow: 0 8px 24px rgba(0,0,0,0.15); z-index: 100000;
                display: none; flex-direction: column; padding: 6px 0;
                opacity: 0; transform: translateY(10px); transition: opacity 0.2s, transform 0.2s;
                border: 1px solid #ebeef5; box-sizing: border-box !important;
                max-height: 85vh; overflow-y: auto;
            }
            .gm-float-menu.show { display: flex; opacity: 1; transform: translateY(0); }

            /* 滚动条美化 */
            .gm-float-menu::-webkit-scrollbar { width: 5px; }
            .gm-float-menu::-webkit-scrollbar-thumb { background: #dcdfe6; border-radius: 3px; }

           /* 分组标题 */
            .gm-menu-group-title {
                font-family: "Helvetica Neue", Helvetica, "PingFang SC", "Microsoft YaHei", Arial, sans-serif !important;
                font-size: 12px !important; color: #909399 !important; padding: 10px 18px 4px !important;
                margin-top: 4px !important; border-top: 1px solid #f0f2f5 !important;
                font-weight: bold !important; pointer-events: none !important; letter-spacing: 1px !important;
            }
            .gm-menu-group-title:first-child { margin-top: 0 !important; border-top: none !important; padding-top: 6px !important; }

            /* 菜单项 */
            .gm-menu-item {
                font-family: "Helvetica Neue", Helvetica, "PingFang SC", "Microsoft YaHei", Arial, sans-serif !important;
                padding: 10px 18px !important;
                cursor: pointer !important; color: #444 !important; font-size: 14px !important; text-align: left !important;
                background: transparent !important; border: none !important;
                border-radius: 0 !important; width: 100% !important; margin: 0 !important;
                transition: background 0.15s, color 0.15s !important;
                display: flex !important; align-items: center !important; gap: 10px !important;
                box-sizing: border-box !important; line-height: 1.5 !important; font-weight: normal !important;
            }
            .gm-menu-item:hover:not(:disabled) { background-color: #f0f7ff !important; color: #409EFF !important; }
            .gm-menu-item:disabled { cursor: not-allowed; color: #c0c4cc !important; }

            .gm-view-main { display: flex; flex-direction: column; width: 100%; }
            .gm-badge { position: absolute; top: -2px; right: -2px; width: 10px; height: 10px; background: #ff4d4f; border-radius: 50%; display: none; border: 2px solid #fff;}
            .gm-icon { width: 18px; text-align: center; display: inline-block; font-weight: bold; flex-shrink: 0; font-size: 15px; }
        `;
            document.head.appendChild(style);
        }

        floatBall = document.createElement('div');
        floatBall.className = 'gm-float-ball';
        floatBall.innerHTML = '<svg viewBox="0 0 24 24" width="26" height="26" stroke="currentColor" stroke-width="2" fill="none" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="3"></circle><path d="M19.4 15a1.65 1.65 0 0 0 .33 1.82l.06.06a2 2 0 0 1 0 2.83 2 2 0 0 1-2.83 0l-.06-.06a1.65 1.65 0 0 0-1.82-.33 1.65 1.65 0 0 0-1 1.51V21a2 2 0 0 1-2 2 2 2 0 0 1-2-2v-.09A1.65 1.65 0 0 0 9 19.4a1.65 1.65 0 0 0-1.82.33l-.06.06a2 2 0 0 1-2.83 0 2 2 0 0 1 0-2.83l.06-.06a1.65 1.65 0 0 0 .33-1.82 1.65 1.65 0 0 0-1.51-1H3a2 2 0 0 1-2-2 2 2 0 0 1 2-2h.09A1.65 1.65 0 0 0 4.6 9a1.65 1.65 0 0 0-.33-1.82l-.06-.06a2 2 0 0 1 0-2.83 2 2 0 0 1 2.83 0l.06.06a1.65 1.65 0 0 0 1.82.33H9a1.65 1.65 0 0 0 1-1.51V3a2 2 0 0 1 2-2 2 2 0 0 1 2 2v.09a1.65 1.65 0 0 0 1 1.51 1.65 1.65 0 0 0 1.82-.33l.06-.06a2 2 0 0 1 2.83 0 2 2 0 0 1 0 2.83l-.06.06a1.65 1.65 0 0 0-.33 1.82V9a1.65 1.65 0 0 0 1.51 1H21a2 2 0 0 1 2 2 2 2 0 0 1-2 2h-.09a1.65 1.65 0 0 0-1.51 1z"></path></svg><div class="gm-badge"></div>';
        floatBall.title = "翱翔教务功能增强设置";
        document.body.appendChild(floatBall);

        floatMenu = document.createElement('div');
        floatMenu.className = 'gm-float-menu';

        const mainView = document.createElement('div');
        mainView.className = 'gm-view-main';

        // 根据环境注入不同的 DOM
        if (isEcampus) {
            mainView.innerHTML = `
            <div class="gm-menu-group-title">快捷工具</div>
            <button class="gm-menu-item" id="gm-btn-person-search"><span class="gm-icon"><svg viewBox="0 0 24 24" width="14" height="14" stroke="currentColor" stroke-width="2" fill="none" stroke-linecap="round" stroke-linejoin="round"><circle cx="11" cy="11" r="8"></circle><line x1="21" y1="21" x2="16.65" y2="16.65"></line></svg></span> 人员信息检索</button>
            <button class="gm-menu-item" id="gm-btn-links"><span class="gm-icon"><svg viewBox="0 0 24 24" width="14" height="14" stroke="currentColor" stroke-width="2" fill="none" stroke-linecap="round" stroke-linejoin="round"><path d="M10 13a5 5 0 0 0 7.54.54l3-3a5 5 0 0 0-7.07-7.07l-1.72 1.71"></path><path d="M14 11a5 5 0 0 0-7.54-.54l-3 3a5 5 0 0 0 7.07 7.07l1.71-1.71"></path></svg></span> 常用链接访问</button>

            <div class="gm-menu-group-title">系统设置</div>
            <button class="gm-menu-item" id="gm-btn-update"><span class="gm-icon"><svg viewBox="0 0 24 24" width="14" height="14" stroke="currentColor" stroke-width="2" fill="none" stroke-linecap="round" stroke-linejoin="round"><polyline points="1 4 1 10 7 10"></polyline><polyline points="23 20 23 14 17 14"></polyline><path d="M20.49 9A9 9 0 0 0 5.64 5.64L1 10m22 4l-4.64 4.36A9 9 0 0 1 3.51 15"></path></svg></span> <span id="gm-update-text">检查版本更新</span></button>
        `;
        } else {
            mainView.innerHTML = `
            <div class="gm-menu-group-title">成绩与学业分析</div>
            <button class="gm-menu-item" id="gm-btn-gpa" disabled><span class="gm-icon"><svg viewBox="0 0 24 24" width="14" height="14" stroke="currentColor" stroke-width="2" fill="none" stroke-linecap="round" stroke-linejoin="round"><polyline points="22 12 18 12 15 21 9 3 6 12 2 12"></polyline></svg></span> GPA综合分析</button>
            <button class="gm-menu-item" id="gm-btn-export" disabled><span class="gm-icon"><svg viewBox="0 0 24 24" width="14" height="14" stroke="currentColor" stroke-width="2" fill="none" stroke-linecap="round" stroke-linejoin="round"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path><polyline points="7 10 12 15 17 10"></polyline><line x1="12" y1="15" x2="12" y2="3"></line></svg></span> 导出成绩与排名</button>

            <div class="gm-menu-group-title">课程助手</div>
            <button class="gm-menu-item" id="gm-btn-follow"><span class="gm-icon"><svg viewBox="0 0 24 24" width="14" height="14" stroke="currentColor" stroke-width="2" fill="none" stroke-linecap="round" stroke-linejoin="round"><path d="M20.84 4.61a5.5 5.5 0 0 0-7.78 0L12 5.67l-1.06-1.06a5.5 5.5 0 0 0-7.78 7.78l1.06 1.06L12 21.23l7.78-7.78 1.06-1.06a5.5 5.5 0 0 0 0-7.78z"></path></svg></span> 课程关注列表</button>
            <button class="gm-menu-item" id="gm-btn-sync-course"><span class="gm-icon"><svg viewBox="0 0 24 24" width="14" height="14" stroke="currentColor" stroke-width="2" fill="none" stroke-linecap="round" stroke-linejoin="round"><path d="M2 12h4l2-2 4 4 2-2h4"></path></svg></span> 同步最新选课数据</button>
            <button class="gm-menu-item" id="gm-btn-eval-jump"><span class="gm-icon"><svg viewBox="0 0 24 24" width="14" height="14" stroke="currentColor" stroke-width="2" fill="none" stroke-linecap="round" stroke-linejoin="round"><polygon points="14 2 18 6 7 17 3 17 3 13 14 2"></polygon><line x1="3" y1="22" x2="21" y2="22"></line></svg></span> 一键自动评教</button>

            <div class="gm-menu-group-title">偏好与系统设置</div>
            <button class="gm-menu-item" id="gm-chk-portrait-btn"><span class="gm-icon" id="icon-portrait"></span> 启用学生画像增强</button>
            <button class="gm-menu-item" id="gm-chk-watch-btn"><span class="gm-icon" id="icon-watch"></span> 启用选课辅助功能</button>
            <button class="gm-menu-item" id="gm-btn-help"><span class="gm-icon"><svg viewBox="0 0 24 24" width="14" height="14" stroke="currentColor" stroke-width="2" fill="none" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="10"></circle><path d="M9.09 9a3 3 0 0 1 5.83 1c0 2-3 3-3 3"></path><line x1="12" y1="17" x2="12.01" y2="17"></line></svg></span> 脚本使用说明</button>
            <button class="gm-menu-item" id="gm-btn-update"><span class="gm-icon"><svg viewBox="0 0 24 24" width="14" height="14" stroke="currentColor" stroke-width="2" fill="none" stroke-linecap="round" stroke-linejoin="round"><polyline points="1 4 1 10 7 10"></polyline><polyline points="23 20 23 14 17 14"></polyline><path d="M20.49 9A9 9 0 0 0 5.64 5.64L1 10m22 4l-4.64 4.36A9 9 0 0 1 3.51 15"></path></svg></span> <span id="gm-update-text">检查版本更新</span></button>
        `;
        }

        floatMenu.appendChild(mainView);
        document.body.appendChild(floatMenu);

        // 绑定共有事件
        const linksBtn = document.getElementById('gm-btn-links');
        if (linksBtn) {
            linksBtn.onclick = () => {
                hideMenu(); // 先隐藏悬浮球菜单
                FriendlyLinks.show(); // 弹出友情链接弹窗
            };
        }
        const psBtn = document.getElementById('gm-btn-person-search');
        if (psBtn) {
            psBtn.onclick = () => { hideMenu(); PersonnelSearch.openModal(); };
        }

        // 绑定教务系统专属事件
        if (!isEcampus) {
            menuExportBtn = document.getElementById('gm-btn-export');
            menuGpaBtn = document.getElementById('gm-btn-gpa');
            menuSyncBtn = document.getElementById('gm-btn-sync-course');
            menuFollowBtn = document.getElementById('gm-btn-follow');
            const menuHelpBtn = document.getElementById('gm-btn-help');

            menuExportBtn.onclick = handleExportClick;
            menuGpaBtn.onclick = handleGpaClick;
            menuSyncBtn.onclick = handleSyncCourseClick;
            menuFollowBtn.onclick = handleShowFollowedClick;
            menuHelpBtn.onclick = () => handleHelpClick();
            document.getElementById('gm-btn-eval-jump').onclick = handleJumpToEvaluation;

            const updateToggleUI = () => {
                const isPortrait = ConfigManager.enablePortraitEnhancement;
                const isWatch = ConfigManager.enableCourseWatch;
                const svgChecked = '<svg viewBox="0 0 24 24" width="15" height="15" stroke="currentColor" stroke-width="2" fill="none" stroke-linecap="round" stroke-linejoin="round"><polyline points="9 11 12 14 22 4"></polyline><path d="M21 12v7a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11"></path></svg>';
                const svgUnchecked = '<svg viewBox="0 0 24 24" width="15" height="15" stroke="currentColor" stroke-width="2" fill="none" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="3" width="18" height="18" rx="2" ry="2"></rect></svg>';
                document.getElementById('icon-portrait').innerHTML = isPortrait ? svgChecked : svgUnchecked;
                document.getElementById('icon-watch').innerHTML = isWatch ? svgChecked : svgUnchecked;

                document.getElementById('gm-chk-portrait-btn').style.color = isPortrait ? '#333' : '#999';
                document.getElementById('gm-chk-watch-btn').style.color = isWatch ? '#333' : '#999';
            };

            document.getElementById('gm-chk-portrait-btn').onclick = () => {
                ConfigManager.enablePortraitEnhancement = !ConfigManager.enablePortraitEnhancement;
                updateToggleUI();
                if (window.location.href.includes('student-portrait')) {
                    if (confirm("修改画像增强设置需要刷新页面生效，是否刷新？")) window.location.reload();
                }
            };

            document.getElementById('gm-chk-watch-btn').onclick = () => {
                ConfigManager.enableCourseWatch = !ConfigManager.enableCourseWatch;
                updateToggleUI();
                if (window.location.href.includes('lesson-search')) {
                    alert("课程关注设置已更新，将在下次进入页面或翻页时生效。");
                }
            };
            updateToggleUI();
        }

        // 处理悬浮球全局事件 (防点击空白处、拖拽等)
        document.addEventListener('click', (e) => { if (!floatMenu.contains(e.target) && !floatBall.contains(e.target)) hideMenu(); });

        let isDragging = false, hasMoved = false, startX, startY, initialLeft, initialTop;
        floatBall.addEventListener('mousedown', (e) => {
            isDragging = true; hasMoved = false; startX = e.clientX; startY = e.clientY;
            const rect = floatBall.getBoundingClientRect(); initialLeft = rect.left; initialTop = rect.top;
            floatBall.style.transition = 'none';
        });
        document.addEventListener('mousemove', (e) => {
            if (!isDragging) return;
            const deltaX = e.clientX - startX, deltaY = e.clientY - startY;
            if (Math.abs(deltaX) > 3 || Math.abs(deltaY) > 3) { hasMoved = true; hideMenu(); }
            floatBall.style.left = Math.min(Math.max(0, initialLeft + deltaX), window.innerWidth - 50) + 'px';
            floatBall.style.top = Math.min(Math.max(0, initialTop + deltaY), window.innerHeight - 50) + 'px';
            floatBall.style.bottom = 'auto'; floatBall.style.right = 'auto';
        });
        document.addEventListener('mouseup', () => { if (isDragging) { isDragging = false; floatBall.style.transition = 'all 0.2s'; } });

        floatBall.addEventListener('click', (e) => {
            e.stopPropagation(); if (hasMoved) return;
            if (floatMenu.classList.contains('show')) hideMenu();
            else {
                const rect = floatBall.getBoundingClientRect();
                // 菜单宽度由 230 调整为 240，留 10px 间距避免压住悬浮球
                let left = rect.left - 240 - 10;
                if (left < 10) left = rect.right + 10;

                floatMenu.style.left = left + 'px';
                floatMenu.style.top = rect.top + 'px';
                showMenu();
            }
        });

        // 绑定更新按钮点击事件
        document.getElementById('gm-btn-update').onclick = () => {
            const btnText = document.getElementById('gm-update-text').innerText;
            if (btnText.includes('发现新版本')) {
                UpdateChecker.openUpdatePage();
            } else {
                UpdateChecker.check(false);
            }
        };
        // 首次加载自动静默检查更新
        UpdateChecker.check(true);
    }

    function showMenu() { floatMenu.style.display = 'flex'; floatMenu.offsetHeight; floatMenu.classList.add('show'); }

    function hideMenu() { floatMenu.classList.remove('show'); setTimeout(() => { if (!floatMenu.classList.contains('show')) floatMenu.style.display = 'none'; }, 200); }

    function updateMenuButtonsState(isReady) {
        if (!menuExportBtn || !menuGpaBtn) return;
        menuExportBtn.disabled = !isReady;
        menuGpaBtn.disabled = !isReady;

        const badge = floatBall.querySelector('.gm-badge');
        if (badge) {
            badge.style.display = (!isReady || isBackgroundSyncing || isUpdateAvailable) ? 'block' : 'none';
        }
    }

    // ----------------- 功能处理函数 -----------------

    /**
     * 处理点击导出按钮
     */
    function handleExportClick() {
        hideMenu();
        if (semesterCheckboxContainer && semesterCheckboxContainer.style.display === "block") {
            semesterCheckboxContainer.style.display = "none";
            return;
        }
        const cachedData = getCachedData();
        if (isDataReady && cachedData) {
            if (!cachedData.allGrades || cachedData.allGrades.length === 0) {
                alert("当前没有可以导出的成绩数据。");
                return;
            }
            if (typeof showSemesterCheckboxes === 'function') showSemesterCheckboxes(cachedData.semesterNames);
        } else {
            alert("成绩数据仍在后台加载中，请稍候...");
        }
    }

    /**
     * 处理点击GPA分析按钮
     */
    function handleGpaClick() {
        hideMenu();
        const cachedData = getCachedData();
        if (isDataReady && cachedData) {
            if (typeof calculateAndDisplayGPA === 'function') calculateAndDisplayGPA(cachedData);
        } else {
            alert("成绩数据仍在后台加载中，请稍候...");
        }
    }

    /**
     * 处理点击同步数据按钮（直接API调用，无需跳转）
     */
    // 打开同步选课数据功能弹窗（在线同步 / 数据导出 / 本地上传）
    function handleSyncCourseClick() {
        hideMenu();
        showSyncDataModal();
    }

    // 在线同步：从教务系统抓取最新选课数据（此处对 ArchiveManager 的引用在用户点击后才执行，其定义见本函数之后）
    async function doSyncCourseData() {
        const confirmMsg = '【更新当前选课学期数据】\n\n' +
            '此操作将获取当前选课学期数据，数据将用于选课页面的”历史余量”参考。\n' +
            '同步将花费几秒到十几秒，是否开始同步？';

        if (!confirm(confirmMsg)) return;

        // --- 1. 获取学生ID ---
        let studentId = await getStudentId();
        if (!studentId) {
            alert('无法获取学生ID，请先访问过教务系统页面后再试。');
            return;
        }

        // 显示同步遮罩（z-index 高于悬浮球，同步期间不可再操作悬浮球）
        const overlay = document.createElement('div');
        overlay.id = 'gm-sync-overlay';
        overlay.style.cssText = 'position:fixed; top:0; left:0; width:100%; height:100%; background:rgba(0,0,0,0.7); z-index:100002; color:white; display:flex; flex-direction:column; align-items:center; justify-content:center;';
        overlay.innerHTML = `
            <div style="font-size:24px; font-weight:bold; margin-bottom:15px;">正在同步选课数据...</div>
            <div id="gm-overlay-target-sem" style="font-size:18px; margin-bottom:10px; color:#ffeb3b;">正在寻找最新选课学期...</div>
            <div id="gm-overlay-status" style="font-size:16px; margin-bottom:10px; color:#ddd;">正在初始化...</div>
            <div style="font-size:18px;">已获取: <span id="gm-sync-count" style="color:#4facfe; font-weight:bold;">0</span> 条</div>
            <div style="margin-top:30px; color:#aaa; font-size:14px;">请勿关闭页面，程序正在自动操作</div>
        `;
        document.body.appendChild(overlay);

        const updateStatus = (text) => {
            const el = document.getElementById('gm-overlay-status');
            if (el) el.innerText = text;
        };
        const updateCount = (count) => {
            const el = document.getElementById('gm-sync-count');
            if (el) el.innerText = count;
        };
        const updateTargetSem = (text) => {
            const el = document.getElementById('gm-overlay-target-sem');
            if (el) el.innerText = text;
        };

        try {
            // --- 3. API 配置 ---
            const API_BASE = '/student/for-std/lesson-search/semester';
            const API_FIELDS = 'course.code,minorCourse.nameZh,courseType,openDepartment,teacherAssignmentList,examMode,campus,teachLang,roomType,timeTableLayout,crossBizTypes,courseProperty';
            const PAGE_SIZE = 1000;

            const fetchCoursePage = (semesterId, page, pageSize) => {
                return new Promise((resolve, reject) => {
                    var url = API_BASE + '/' + semesterId + '/search/' + studentId +
                        '?bizTypeAssoc=1&queryPage__=' + page + ',' + pageSize +
                        '&assembleFields=' + encodeURIComponent(API_FIELDS) +
                        '&_=' + Date.now();
                    var xhr = new XMLHttpRequest();
                    xhr.open('GET', url);
                    xhr.setRequestHeader('X-Requested-With', 'XMLHttpRequest');
                    xhr.onload = function () {
                        if (xhr.status === 200) {
                            try { resolve(JSON.parse(xhr.responseText)); }
                            catch (e) { reject(new Error('API响应解析失败')); }
                        } else { reject(new Error('HTTP ' + xhr.status)); }
                    };
                    xhr.onerror = function () { reject(new Error('网络请求失败')); };
                    xhr.send();
                });
            };

            // --- 4. 通过API获取学期列表 ---
            updateStatus('正在获取学期列表...');
            const semesterOptions = await new Promise((resolve, reject) => {
                const url = '/student/for-std/lesson-search/index/' + studentId;
                const xhr = new XMLHttpRequest();
                xhr.open('GET', url);
                xhr.onload = function () {
                    if (xhr.status === 200) {
                        try {
                            const html = xhr.responseText;
                            const parser = new DOMParser();
                            const doc = parser.parseFromString(html, 'text/html');
                            const options = doc.querySelectorAll('#semester option[value]');
                            const semesterList = [];
                            options.forEach(opt => {
                                const val = opt.getAttribute('value');
                                const text = opt.textContent.trim();
                                if (val && text && val !== '') {
                                    semesterList.push({ value: val, text: text });
                                }
                            });
                            resolve(semesterList);
                        } catch (e) {
                            reject(new Error('HTML解析失败: ' + e.message));
                        }
                    } else {
                        reject(new Error('请求失败: HTTP ' + xhr.status));
                    }
                };
                xhr.onerror = function () {
                    reject(new Error('网络请求失败'));
                };
                xhr.send();
            });

            if (semesterOptions.length === 0) {
                throw new Error('无法获取学期列表');
            }

            // 按学期ID降序排列（最新的在前）
            semesterOptions.sort((a, b) => parseInt(b.value) - parseInt(a.value));

            // --- 5. 寻找最新有数据的学期 ---
            let targetSemesterId = '';
            let targetSemesterName = '';

            for (var i = 0; i < semesterOptions.length; i++) {
                var opt = semesterOptions[i];
                updateStatus('正在检测学期: ' + opt.text + '...');
                try {
                    var testData = await fetchCoursePage(opt.value, 1, 1);
                    if (testData && testData.data && testData.data.length > 0) {
                        targetSemesterId = opt.value;
                        targetSemesterName = opt.text;
                        updateTargetSem('锁定抓取学期: ' + targetSemesterName);
                        Logger.log('Sync', '找到有数据的学期: ' + opt.text);
                        break;
                    }
                } catch (e) {
                    Logger.log('Sync', '学期 ' + opt.text + ' 探测失败，跳过');
                }
            }

            if (!targetSemesterId) {
                throw new Error('未找到任何包含排课数据的学期');
            }

            // --- 6. 并行获取全部数据 ---
            updateStatus('正在获取 ' + targetSemesterName + ' 的课程数据...');
            const firstResult = await fetchCoursePage(targetSemesterId, 1, PAGE_SIZE);
            if (!firstResult || !firstResult.data || firstResult.data.length === 0) {
                throw new Error(targetSemesterName + ' 学期暂无课程数据');
            }

            const totalPages = firstResult._page_?.totalPages || 1;
            const timestamp = new Date().toLocaleString('zh-CN', { hour12: false });

            // 解析课程数据
            // 课程名在 course.course.nameZh（顶层 nameZh 是班级范围名如"全校"）
            const parseCourseData = (courses) => courses.map(course => ({
                id: String(course.id),
                code: course.code || 'N/A',
                name: course.course?.nameZh || course.nameZh || 'N/A',
                stdCount: course.stdCount || 0,
                limitCount: course.limitCount || 0,
                term: targetSemesterName,
                time: timestamp
            }));

            let allCourseData = parseCourseData(firstResult.data);
            updateCount(allCourseData.length);

            // 并行请求剩余页
            if (totalPages > 1) {
                updateStatus('正在获取剩余 ' + (totalPages - 1) + ' 页...');
                const pagePromises = [];
                for (let page = 2; page <= totalPages; page++) {
                    pagePromises.push(fetchCoursePage(targetSemesterId, page, PAGE_SIZE));
                }
                const results = await Promise.all(pagePromises);
                for (const result of results) {
                    if (result && result.data && result.data.length > 0) {
                        allCourseData.push(...parseCourseData(result.data));
                    }
                }
                updateCount(allCourseData.length);
            }

            // --- 7. 保存数据 ---
            updateStatus('正在保存数据...');
            let history = {};
            try { history = JSON.parse(GM_getValue(CONSTANTS.HISTORY_STORAGE_KEY, '{}')); }
            catch (e) { history = {}; }

            allCourseData.forEach(c => {
                if (!history[c.id]) history[c.id] = [];
                const records = history[c.id];
                const last = records[records.length - 1];
                if (!last || last.stdCount !== c.stdCount || last.limitCount !== c.limitCount) {
                    records.push(c);
                } else {
                    last.time = c.time;
                }
            });

            GM_setValue(CONSTANTS.HISTORY_STORAGE_KEY, JSON.stringify(history));

            overlay.remove();
            // 游戏式存档：在线同步的新数据静默写入「使用中」的存档（保留自定义命名）
            const savedArchiveName = ArchiveManager.updateActiveSlot(targetSemesterName);
            showToast(savedArchiveName
                ? `同步完成：${allCourseData.length} 门课程已存入存档【${savedArchiveName}】`
                : `同步完成：共获取 ${allCourseData.length} 门课程数据`, 'success');

        } catch (e) {
            Logger.error('Sync', "同步错误", e);
            overlay.remove();
            alert('同步失败: ' + e.message);
        }
    }

    // --- 选课数据存档管理（3 个固定槽位：防误操作覆盖、可回溯过往数据） ---
    const ArchiveManager = {
        STORAGE_KEY: 'jwxt_sync_archives_v1',
        SLOT_COUNT: 3,
        ICONS: {
            archive: '<svg viewBox="0 0 24 24" width="16" height="16" stroke="currentColor" stroke-width="2" fill="none" stroke-linecap="round" stroke-linejoin="round"><polyline points="21 8 21 21 3 21 3 8"></polyline><rect x="1" y="3" width="22" height="5"></rect><line x1="10" y1="12" x2="14" y2="12"></line></svg>',
            box: '<svg viewBox="0 0 24 24" width="16" height="16" stroke="#c0c4cc" stroke-width="2" fill="none" stroke-linecap="round" stroke-linejoin="round"><path d="M21 16V8a2 2 0 0 0-1-1.73l-7-4a2 2 0 0 0-2 0l-7 4A2 2 0 0 0 3 8v8a2 2 0 0 0 1 1.73l7 4a2 2 0 0 0 2 0l7-4A2 2 0 0 0 21 16z"></path></svg>',
            success: '<svg viewBox="0 0 24 24" width="22" height="22" stroke="#67C23A" stroke-width="2" fill="none" stroke-linecap="round" stroke-linejoin="round"><path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"></path><polyline points="22 4 12 14.01 9 11.01"></polyline></svg>',
            // 存档行右侧操作图标（颜色由按钮 style 控制，均为 currentColor）
            applyIcon: '<svg viewBox="0 0 24 24" width="14" height="14" stroke="currentColor" stroke-width="2" fill="none" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="10"></circle><polygon points="10 8 16 12 10 16 10 8"></polygon></svg>',
            exportIcon: '<svg viewBox="0 0 24 24" width="14" height="14" stroke="currentColor" stroke-width="2" fill="none" stroke-linecap="round" stroke-linejoin="round"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path><polyline points="7 10 12 15 17 10"></polyline><line x1="12" y1="15" x2="12" y2="3"></line></svg>',
            renameIcon: '<svg viewBox="0 0 24 24" width="14" height="14" stroke="currentColor" stroke-width="2" fill="none" stroke-linecap="round" stroke-linejoin="round"><path d="M12 20h9"></path><path d="M16.5 3.5a2.121 2.121 0 0 1 3 3L7 19l-4 1 1-4L16.5 3.5z"></path></svg>',
            removeIcon: '<svg viewBox="0 0 24 24" width="14" height="14" stroke="currentColor" stroke-width="2" fill="none" stroke-linecap="round" stroke-linejoin="round"><polyline points="3 6 5 6 21 6"></polyline><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"></path></svg>'
        },

        load() {
            let state;
            try {
                state = JSON.parse(GM_getValue(this.STORAGE_KEY, 'null'));
            } catch (e) { }
            if (!state || !Array.isArray(state.slots) || state.slots.length !== this.SLOT_COUNT) {
                state = { slots: new Array(this.SLOT_COUNT).fill(null), sourceSlot: -1 };
            }
            return this.migrateLegacyData(state);
        },

        // 旧版本升级迁移：已有活跃选课数据但从未建立过存档时，默认将其存入槽位1并设为「使用中」（幂等）
        migrateLegacyData(state) {
            if (state.sourceSlot === -1 && state.slots.every(s => !s)) {
                const history = this.readActiveHistory();
                if (history && Object.keys(history).length > 0) {
                    state.slots[0] = {
                        name: this.buildAutoName(),
                        semester: '历史数据',
                        savedAt: new Date().toLocaleString('zh-CN', { hour12: false }),
                        count: Object.keys(history).length,
                        data: history,
                        customNamed: false
                    };
                    state.sourceSlot = 0;
                    this.save(state);
                }
            }
            return state;
        },

        save(state) { GM_setValue(this.STORAGE_KEY, JSON.stringify(state)); },

        // —— 存档定位辅助（游戏式：始终使用三个存档之一） ——

        // 新数据的落槽规则：优先写入「使用中」的存档；全空时默认槽1；否则取第一个空槽
        targetIndexForNewData(state) {
            if (state.sourceSlot >= 0 && state.sourceSlot < this.SLOT_COUNT) return state.sourceSlot;
            const emptyIdx = state.slots.findIndex(s => !s);
            return emptyIdx === -1 ? 0 : emptyIdx;
        },

        // 与被删槽位距离最近的非空存档（先找左侧，再找右侧），供删除后自动切换
        findNearestOccupied(slots, removedIdx) {
            for (let d = 1; d < this.SLOT_COUNT; d++) {
                if (removedIdx - d >= 0 && slots[removedIdx - d]) return removedIdx - d;
                if (removedIdx + d < this.SLOT_COUNT && slots[removedIdx + d]) return removedIdx + d;
            }
            return -1;
        },

        // 生成自动存档名：存档日期时间（如 2026-08-27 14:30）
        buildAutoName() {
            const now = new Date();
            const pad = n => String(n).padStart(2, '0');
            return `${now.getFullYear()}-${pad(now.getMonth() + 1)}-${pad(now.getDate())} ${pad(now.getHours())}:${pad(now.getMinutes())}`;
        },

        // 读取当前活跃选课数据（选课助手正在使用的那份）
        readActiveHistory() {
            try { return JSON.parse(GM_getValue(CONSTANTS.HISTORY_STORAGE_KEY, '{}')); } catch (e) { return {}; }
        },

        // 统计存档数据覆盖的学期数（以每门课最新记录的 term 标记为准；旧版数据无标记则不计入）
        countTerms(dataObj) {
            const terms = new Set();
            Object.values(dataObj || {}).forEach(records => {
                const last = (records || [])[records.length - 1];
                if (last && last.term) terms.add(last.term);
            });
            return terms.size;
        },

        // 槽位规模描述：如「3个学期共15000门课程」；无学期标记的旧数据仅显示课程数
        describeSlotSize(slot) {
            const termCount = this.countTerms(slot.data);
            return termCount > 0 ? `${termCount}个学期共${slot.count}门课程` : `${slot.count}门课程`;
        },

        // 将数据对象覆盖式写入指定槽位；keepCustomName 时保留用户手动命名的名称
        writeToSlot(index, historyObj, semesterName, keepCustomName) {
            if (!historyObj || Object.keys(historyObj).length === 0) {
                GmDialog.alert({ title: '无法存档', message: '当前没有可存档的选课数据，请先执行在线同步或本地上传。', type: 'warn' });
                return false;
            }
            const state = this.load();
            const old = state.slots[index];
            const keepName = keepCustomName && old && old.customNamed;
            state.slots[index] = {
                name: keepName ? old.name : this.buildAutoName(),
                semester: semesterName || '未知学期',
                savedAt: new Date().toLocaleString('zh-CN', { hour12: false }),
                count: Object.keys(historyObj).length,
                data: historyObj,
                customNamed: !!keepName
            };
            this.save(state);
            return true;
        },

        // 在线同步后静默更新「使用中」的存档（保留自定义命名，刷新数据与元信息）；返回存档名供轻提示
        updateActiveSlot(semesterName) {
            const history = this.readActiveHistory();
            if (!history || Object.keys(history).length === 0) return null;
            const state = this.load();
            const idx = this.targetIndexForNewData(state);
            if (!this.writeToSlot(idx, history, semesterName, true)) return null;
            if (state.sourceSlot !== idx) {
                state.sourceSlot = idx;
                this.save(state);
            }
            return this.load().slots[idx].name;
        },

        // 将槽位存档写回活跃数据并设为「使用中」的存档，供选课助手/历史余量立即生效
        applyAsSource(index) {
            const state = this.load();
            const slot = state.slots[index];
            if (!slot) return;
            GM_setValue(CONSTANTS.HISTORY_STORAGE_KEY, JSON.stringify(slot.data));
            state.sourceSlot = index;
            this.save(state);
        },

        // 删除指定槽位存档；若删的是「使用中」存档，自动切换到最近的非空存档继续使用，全部为空则清空活跃数据
        removeSlot(index) {
            const state = this.load();
            state.slots[index] = null;
            let switchedTo = null;
            if (state.sourceSlot === index) {
                const next = this.findNearestOccupied(state.slots, index);
                state.sourceSlot = next;
                if (next >= 0) {
                    GM_setValue(CONSTANTS.HISTORY_STORAGE_KEY, JSON.stringify(state.slots[next].data));
                    switchedTo = state.slots[next].name;
                } else {
                    GM_setValue(CONSTANTS.HISTORY_STORAGE_KEY, '{}');
                }
            }
            this.save(state);
            return switchedTo;
        },

        // 手动重命名存档；留空则恢复自动命名
        renameSlot(index, newName) {
            const state = this.load();
            const slot = state.slots[index];
            if (!slot) return;
            const trimmed = String(newName || '').trim();
            if (trimmed) {
                slot.name = trimmed;
                slot.customNamed = true;
            } else {
                slot.name = this.buildAutoName();
                slot.customNamed = false;
            }
            this.save(state);
        },

        // 将指定槽位存档导出为 JSON 文件（分享/备份）；时间戳取缓存数据的最新同步时间（而非导出时间）
        exportSlot(index) {
            const slot = this.load().slots[index];
            if (!slot) return;
            let dataTime = null;
            Object.values(slot.data || {}).forEach(records => {
                (records || []).forEach(r => { if (r && r.time && (!dataTime || r.time > dataTime)) dataTime = r.time; });
            });
            const exportTime = dataTime || new Date().toLocaleString('zh-CN', { hour12: false });
            let ts = '';
            const m = String(dataTime || '').match(/^(\d{4})\/(\d{1,2})\/(\d{1,2})\s+(\d{1,2}):(\d{1,2})/);
            if (m) {
                ts = `${m[1]}${String(m[2]).padStart(2, '0')}${String(m[3]).padStart(2, '0')}T${String(m[4]).padStart(2, '0')}${String(m[5]).padStart(2, '0')}`;
            }
            const blob = new Blob([JSON.stringify({ exportTime, version: GM_info.script.version, data: slot.data }, null, 2)], { type: 'application/json' });
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = `翱翔教务选课数据${ts}.json`;
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            URL.revokeObjectURL(url);
        },

        // 注入存档区专用样式（图标按钮 / 使用中徽章）
        ensureStyles() {
            if (document.getElementById('gm-archive-style')) return;
            const style = getOrCreateStyleEl('gm-archive-style');
            style.textContent = `
            .gm-arc-btn { width:26px; height:26px; border-radius:6px; border:none; background:transparent; display:inline-flex; align-items:center; justify-content:center; cursor:pointer; padding:0; transition:all .15s; flex-shrink:0; }
            .gm-arc-btn:hover { background:rgba(64,158,255,.12); }
            .gm-arc-btn[disabled] { color:#c0c4cc !important; cursor:not-allowed; background:transparent !important; }
            .gm-arc-badge { font-size:10px; color:#409EFF; background:#ecf5ff; border:1px solid #d9ecff; border-radius:4px; padding:0 5px; line-height:16px; white-space:nowrap; }
        `;
            document.head.appendChild(style);
        },

        // 渲染同步弹窗内的存档槽位列表（游戏式三槽位：每行右侧为 存入/启用/导出/重命名/删除 图标），操作后通过 onChanged 刷新弹窗
        renderSlotList(container, onChanged) {
            if (!container) return;
            this.ensureStyles();
            const state = this.load();
            const btnClass = {
                apply: 'gm-arc-btn--primary',
                export: 'gm-arc-btn--success',
                rename: 'gm-arc-btn--warning',
                remove: 'gm-arc-btn--danger'
            };
            const mkBtn = (act, i, icon, tip, extra = '') =>
                `<button type="button" class="gm-arc-btn ${btnClass[act]}" data-index="${i}" data-act="${act}" title="${tip}" ${extra}>${icon}</button>`;
            const opBtns = (i, isSource) =>
                mkBtn('apply', i, this.ICONS.applyIcon, isSource ? '当前正在使用此存档' : '启用此存档作为选课助手数据源', isSource ? 'disabled' : '') +
                mkBtn('export', i, this.ICONS.exportIcon, '导出此存档为JSON文件') +
                mkBtn('rename', i, this.ICONS.renameIcon, '重命名存档') +
                mkBtn('remove', i, this.ICONS.removeIcon, '删除此存档');

            container.innerHTML = state.slots.map((slot, i) => {
                const isSource = state.sourceSlot === i;
                if (!slot) {
                    return `
                    <div class="gm-archive-slot is-empty">
                        <span class="gm-archive-slot__icon">${this.ICONS.box}</span>
                        <span>槽位 ${i + 1}：无存档</span>
                    </div>`;
                }
                return `
                <div class="gm-archive-slot ${isSource ? 'is-active' : ''}">
                    <div class="gm-archive-slot__top">
                        <span class="gm-archive-slot__icon">${this.ICONS.archive}</span>
                        <b class="gm-archive-slot__name">${esc(slot.name)}</b>
                        ${isSource ? '<span class="gm-arc-badge">使用中</span>' : ''}
                        <span class="gm-archive-actions">${opBtns(i, isSource)}</span>
                    </div>
                    <div class="gm-archive-slot__meta">${slot.count} 门课程 · 更新于 ${slot.savedAt}</div>
                </div>`;
            }).join('');

            container.onclick = (e) => {
                const btn = e.target.closest('button[data-act]');
                if (!btn || btn.disabled) return;
                const index = parseInt(btn.dataset.index);
                const act = btn.dataset.act;
                if (act === 'apply') {
                    this.applyAsSource(index);
                    onChanged && onChanged();
                    showToast(`已启用存档【${state.slots[index].name}】，选课助手正在使用该数据`, 'success');
                } else if (act === 'export') {
                    this.exportSlot(index);
                } else if (act === 'rename') {
                    this.showRenameDialog(index, state.slots[index].name, newName => {
                        this.renameSlot(index, newName);
                        onChanged && onChanged();
                        showToast(newName.trim() ? '已重命名存档' : '已恢复自动命名', 'success');
                    });
                } else if (act === 'remove') {
                    const slot = state.slots[index];
                    GmDialog.confirm({ title: '删除存档', message: `确定删除槽位 ${index + 1} 的存档【${slot ? slot.name : ''}】吗？\n删除后不可恢复。`, type: 'error', confirmText: '删除' }).then(ok => {
                        if (!ok) return;
                        const switchedTo = this.removeSlot(index);
                        onChanged && onChanged();
                        if (switchedTo) showToast(`已删除存档，自动切换至存档【${switchedTo}】`, 'info');
                        else showToast('已删除存档', 'info');
                    });
                }
            };
        },

        // 重命名输入弹窗（点击改名图标触发），提交回调 onDone(newName)
        showRenameDialog(index, curName, onDone) {
            const overlay = document.createElement('div');
            overlay.style.cssText = 'position:fixed;top:0;left:0;width:100%;height:100%;background:rgba(0,0,0,0.55);z-index:100035;display:flex;align-items:center;justify-content:center;backdrop-filter:blur(3px);';
            overlay.innerHTML = `
            <div style="background:#fff;border-radius:12px;width:360px;max-width:90%;padding:20px 22px;box-shadow:0 12px 32px rgba(0,0,0,0.25);font-family:'Helvetica Neue',Helvetica,'PingFang SC','Microsoft YaHei',Arial,sans-serif;">
                <b style="font-size:15px;color:#303133;display:block;margin-bottom:12px;">重命名存档（槽位 ${index + 1}）</b>
                <input id="gm-archive-rename-input" type="text" maxlength="40" value="${esc(curName)}"
                    style="width:100%;box-sizing:border-box;padding:9px 12px;border:1px solid #dcdfe6;border-radius:6px;font-size:13px;color:#303133;outline:none;"
                    placeholder="输入新的存档名称">
                <div style="font-size:12px;color:#909399;margin-top:6px;">留空恢复为自动命名</div>
                <div style="display:flex;justify-content:flex-end;gap:10px;margin-top:14px;">
                    <button id="gm-archive-rename-cancel" style="padding:7px 16px;border-radius:6px;border:1px solid #dcdfe6;background:#fff;color:#606266;font-size:13px;cursor:pointer;">取消</button>
                    <button id="gm-archive-rename-ok" style="padding:7px 16px;border-radius:6px;border:none;background:#409EFF;color:#fff;font-size:13px;cursor:pointer;">确定</button>
                </div>
            </div>`;
            document.body.appendChild(overlay);
            const input = overlay.querySelector('#gm-archive-rename-input');
            input.focus(); input.select();
            const close = () => overlay.remove();
            const submit = () => { onDone(input.value); close(); };
            overlay.querySelector('#gm-archive-rename-ok').onclick = submit;
            overlay.querySelector('#gm-archive-rename-cancel').onclick = close;
            overlay.onclick = e => { if (e.target === overlay) close(); };
            input.onkeydown = e => { if (e.key === 'Enter') submit(); if (e.key === 'Escape') close(); };
        },

        // 本地上传完成后：必须选择一个槽位存放导入的数据（该槽位同时成为当前使用的存档）
        promptArchive(summary, semesterName) {
            return new Promise(resolve => {
                const state = this.load();
                const overlay = document.createElement('div');
                overlay.style.cssText = 'position:fixed;top:0;left:0;width:100%;height:100%;background:rgba(0,0,0,0.55);z-index:100030;display:flex;align-items:center;justify-content:center;backdrop-filter:blur(3px);animation:gmFadeIn 0.2s;';
                overlay.innerHTML = `
                <div style="background:#fff;border-radius:12px;width:440px;max-width:92%;padding:22px;box-shadow:0 12px 32px rgba(0,0,0,0.25);font-family:'Helvetica Neue',Helvetica,'PingFang SC','Microsoft YaHei',Arial,sans-serif;">
                    <div style="display:flex;align-items:center;gap:10px;margin-bottom:10px;">
                        ${this.ICONS.success}
                        <b style="font-size:16px;color:#303133;">数据导入完成</b>
                    </div>
                    <div style="font-size:13px;color:#606266;line-height:1.6;margin-bottom:14px;">${summary}</div>
                    <div style="font-size:13px;color:#303133;font-weight:600;margin-bottom:4px;">请选择要覆盖的存档槽位：</div>
                    <div style="display:flex;flex-direction:column;gap:8px;">
                        ${state.slots.map((slot, i) => `
                        <button data-index="${i}" style="display:flex;align-items:center;gap:10px;width:100%;padding:10px 12px;border:1px solid ${state.sourceSlot === i ? '#b3d8ff' : '#ebeef5'};border-radius:8px;background:${state.sourceSlot === i ? '#f0f7ff' : '#f5f7fa'};cursor:pointer;text-align:left;transition:background 0.2s;font-size:13px;color:#303133;">
                            <span style="color:#409EFF;display:inline-flex;">${this.ICONS.archive}</span>
                            <span>槽位 ${i + 1}：${slot ? esc(slot.name) + '（' + this.describeSlotSize(slot) + '）' : '空'}</span>
                            ${state.sourceSlot === i ? '<span class="gm-arc-badge">使用中</span>' : ''}
                        </button>`).join('')}
                    </div>
                </div>`;
                document.body.appendChild(overlay);

                const finish = () => { overlay.remove(); resolve(); };
                overlay.querySelectorAll('button[data-index]').forEach(btn => {
                    btn.onclick = () => {
                        const index = parseInt(btn.dataset.index);
                        if (this.writeToSlot(index, this.readActiveHistory(), semesterName, true)) {
                            const st = this.load();
                            st.sourceSlot = index; // 选定槽位即成为当前使用的存档
                            this.save(st);
                            showToast(`导入数据已存入槽位 ${index + 1} 并设为使用中`, 'success');
                        }
                        finish();
                    };
                });
            });
        }
    };

    // 轻提示 Toast（顶部居中，短暂显示后自动消失）
    function showToast(message, type = 'success') {
        const existing = document.getElementById('gm-archive-toast');
        if (existing) existing.remove();
        if (!document.getElementById('gm-archive-toast-style')) {
            const style = getOrCreateStyleEl('gm-archive-toast-style');
            style.textContent = `
            .gm-arc-toast { position: fixed; top: 22px; left: 50%; transform: translateX(-50%) translateY(-10px); background: #fff; color: #303133; padding: 10px 20px; border-radius: 8px; font-size: 13px; font-family: "PingFang SC", "Hiragino Sans GB", "Microsoft YaHei", sans-serif; z-index: 300000; box-shadow: 0 6px 18px rgba(0,0,0,0.15); border-left: 3px solid #67C23A; opacity: 0; transition: all .25s ease-out; pointer-events: none; max-width: 80vw; }
            .gm-arc-toast.show { opacity: 1; transform: translateX(-50%) translateY(0); }
            .gm-arc-toast.cancel { border-left-color: #909399; }
            .gm-arc-toast.info { border-left-color: #409EFF; }
            .gm-arc-toast.warn { border-left-color: #E6A23C; }
            .gm-arc-toast.error { border-left-color: #F56C6C; }
        `;
            document.head.appendChild(style);
        }
        const toast = document.createElement('div');
        toast.id = 'gm-archive-toast';
        toast.className = `gm-arc-toast ${type}`;
        toast.textContent = message;
        document.body.appendChild(toast);
        requestAnimationFrame(() => toast.classList.add('show'));
        setTimeout(() => { toast.classList.remove('show'); setTimeout(() => toast.remove(), 300); }, 2400);
    }

    // 同步选课数据功能弹窗：在线同步 / 数据导出 / 本地上传 / 数据存档
    function showSyncDataModal() {
        const existing = document.getElementById('gm-sync-modal');
        if (existing) existing.remove();

        const overlay = document.createElement('div');
        overlay.id = 'gm-sync-modal';
        overlay.className = 'gm-sync-overlay';
        overlay.style.cssText = 'position:fixed; inset:0; z-index:100010; display:flex; align-items:center; justify-content:center; animation:gmFadeIn 0.2s;';
        overlay.innerHTML = `
        <div class="gm-sync-modal-card">
            <div class="gm-sync-modal-header">
                <div class="gm-sync-modal-title">
                    <svg viewBox="0 0 24 24" width="18" height="18" stroke="#409EFF" stroke-width="2" fill="none" stroke-linecap="round" stroke-linejoin="round"><polyline points="23 4 23 10 17 10"></polyline><polyline points="1 20 1 14 7 14"></polyline><path d="M3.51 9a9 9 0 0 1 14.85-3.36L23 10M1 14l4.64 4.36A9 9 0 0 0 20.49 15"></path></svg>
                    同步最新选课数据
                </div>
                <button type="button" id="gm-sync-close" class="gm-modal-close" aria-label="关闭">&times;</button>
            </div>
            <div class="gm-sync-modal-body">
                <div class="gm-sync-actions">
                    <button type="button" id="gm-sync-online" class="gm-sync-action">
                        <span class="gm-sync-action__icon"><svg viewBox="0 0 24 24" width="20" height="20" stroke="currentColor" stroke-width="2" fill="none" stroke-linecap="round" stroke-linejoin="round"><polyline points="23 4 23 10 17 10"></polyline><polyline points="1 20 1 14 7 14"></polyline><path d="M3.51 9a9 9 0 0 1 14.85-3.36L23 10M1 14l4.64 4.36A9 9 0 0 0 20.49 15"></path></svg></span>
                        <span><b>在线同步</b><small>从教务系统抓取最新选课学期数据</small></span>
                    </button>
                    <button type="button" id="gm-sync-import" class="gm-sync-action">
                        <span class="gm-sync-action__icon gm-sync-action__icon--orange"><svg viewBox="0 0 24 24" width="20" height="20" stroke="currentColor" stroke-width="2" fill="none" stroke-linecap="round" stroke-linejoin="round"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path><polyline points="17 8 12 3 7 8"></polyline><line x1="12" y1="3" x2="12" y2="15"></line></svg></span>
                        <span><b>本地上传</b><small>导入他人分享的选课数据</small></span>
                    </button>
                </div>
                <div class="gm-sync-archive">
                    <div class="gm-sync-archive__title">
                        <svg viewBox="0 0 24 24" width="15" height="15" stroke="#409EFF" stroke-width="2" fill="none" stroke-linecap="round" stroke-linejoin="round"><polyline points="21 8 21 21 3 21 3 8"></polyline><rect x="1" y="3" width="22" height="5"></rect><line x1="10" y1="12" x2="14" y2="12"></line></svg>
                        数据存档
                    </div>
                    <div id="gm-archive-list" class="gm-archive-list"></div>
                </div>
                <input type="file" id="gm-sync-file" accept=".json,application/json" style="display:none;">
            </div>
        </div>
        `;
        document.body.appendChild(overlay);
        ArchiveManager.renderSlotList(document.getElementById('gm-archive-list'), showSyncDataModal);

        const close = () => overlay.remove();
        document.getElementById('gm-sync-close').onclick = close;
        overlay.onclick = (e) => { if (e.target === overlay) close(); };
        document.getElementById('gm-sync-online').onclick = () => { close(); doSyncCourseData(); };
        document.getElementById('gm-sync-import').onclick = () => document.getElementById('gm-sync-file').click();
        document.getElementById('gm-sync-file').onchange = (e) => {
            const file = e.target.files[0];
            if (file) importSyncData(file);
            e.target.value = '';
        };
    }

    // 导入他人分享的选课数据 JSON
    function importSyncData(file) {
        const reader = new FileReader();
        reader.onload = () => {
            try {
                const parsed = JSON.parse(reader.result);
                const history = (parsed && parsed.data && typeof parsed.data === 'object') ? parsed.data : parsed;
                // 仅接受「课程ID → 记录数组」的存档结构，防止误导入任意 JSON
                if (!history || typeof history !== 'object' || Array.isArray(history)) throw new Error('格式不正确');
                const recordLists = Object.values(history);
                if (recordLists.length === 0 || !recordLists.every(v => Array.isArray(v))) throw new Error('格式不正确');
                GM_setValue(CONSTANTS.HISTORY_STORAGE_KEY, JSON.stringify(history));
                // 本地上传：由用户确认覆盖哪个存档槽位（选定槽位将成为当前使用的存档）
                ArchiveManager.promptArchive('导入成功！请选择要覆盖的存档槽位。', '导入数据').then(() => showSyncDataModal());
            } catch (e) {
                alert('文件格式不正确，请使用本脚本导出的 JSON 文件。');
            }
        };
        reader.onerror = () => alert('文件读取失败');
        reader.readAsText(file);
    }

    /**
     * 处理点击帮助按钮 - 弹窗版操作指南
     */
    function handleHelpClick() {
        hideMenu(); // 关闭悬浮菜单

        // 1. 注入弹窗专用样式
        const styleId = 'gm-help-popup-style';
        if (!document.getElementById(styleId)) {
            const style = getOrCreateStyleEl(styleId);
            style.textContent = `
            /* 遮罩层 */
            .gm-help-overlay {
                position: fixed; top: 0; left: 0; width: 100%; height: 100%;
                background-color: rgba(0, 0, 0, 0.6);
                z-index: 20000;
                display: flex; align-items: center; justify-content: center;
                backdrop-filter: blur(3px);
                animation: gmFadeIn 0.2s ease-out;
            }
            /* 弹窗主体 */
            .gm-help-modal {
                background: #fff;
                width: 650px;
                max-width: 90%;
                max-height: 85vh;
                border-radius: 12px;
                box-shadow: 0 8px 30px rgba(0, 0, 0, 0.2);
                display: flex; flex-direction: column;
                font-family: "Helvetica Neue", Helvetica, "PingFang SC", "Microsoft YaHei", Arial, sans-serif;
                overflow: hidden;
                border: 1px solid #eee;
            }
            /* 标题栏 */
            .gm-help-header {
                padding: 16px 24px;
                border-bottom: 1px solid #eee;
                display: flex; justify-content: space-between; align-items: center;
                background: #fcfcfc;
            }
            .gm-help-title { font-size: 18px; font-weight: bold; color: #333; display: flex; align-items: center; gap: 8px; }
            .gm-help-close { border: none; background: transparent; font-size: 24px; color: #999; cursor: pointer; transition: color 0.2s; }
            .gm-help-close:hover { color: #F56C6C; }

            /* 内容区 */
            .gm-help-body { padding: 0; overflow-y: auto; background-color: #fcfcfc; }
            .gm-help-section {
                background: #fff;
                margin: 0 0 12px 0;
                padding: 18px 24px;
                border-bottom: 1px solid #f0f0f0;
            }
            .gm-help-sec-title {
                font-size: 15px; font-weight: bold; color: #303133;
                margin-bottom: 12px; padding-left: 10px;
                border-left: 4px solid #409EFF;
                display: flex; align-items: center; justify-content: space-between;
            }
            .gm-help-step {
                font-size: 13.5px; color: #555; line-height: 1.7; margin-bottom: 8px;
                position: relative; padding-left: 15px;
            }
            .gm-help-step::before {
                content: "•"; position: absolute; left: 0; color: #bbb;
            }

            /* UI 标签模拟 */
            .gm-tag {
                display: inline-block; padding: 0 6px; border-radius: 4px;
                font-size: 12px; font-family: monospace; margin: 0 2px;
            }
            .gm-tag-blue { background: #ecf5ff; color: #409EFF; border: 1px solid #d9ecff; }
            .gm-tag-red  { background: #fef0f0; color: #F56C6C; border: 1px solid #fde2e2; }
            .gm-tag-gray { background: #f4f4f5; color: #909399; border: 1px solid #e9e9eb; }

            /* 动画 */
            @keyframes gmFadeIn { from { opacity: 0; } to { opacity: 1; } }
        `;
            document.head.appendChild(style);
        }

        // 2. 创建 DOM 结构
        const overlay = document.createElement('div');
        overlay.className = 'gm-help-overlay';

        const modal = document.createElement('div');
        modal.className = 'gm-help-modal';

        // 3. 构建 HTML 内容
        modal.innerHTML = `
        <div class="gm-help-header">
            <div class="gm-help-title">脚本使用说明</div>
            <button class="gm-help-close" title="关闭">×</button>
        </div>
        <div class="gm-help-body">

            <!-- 模块：成绩与画像 -->
            <div class="gm-help-section">
                <div class="gm-help-sec-title" style="border-color:#F56C6C;">
                    成绩与学业分析
                </div>
                <div class="gm-help-step">
                    点击悬浮球菜单 <span class="gm-tag gm-tag-gray"><svg viewBox="0 0 24 24" width="12" height="12" style="vertical-align:-1px;" stroke="currentColor" stroke-width="2" fill="none" stroke-linecap="round" stroke-linejoin="round"><polyline points="22 12 18 12 15 21 9 3 6 12 2 12"></polyline></svg> GPA综合分析</span>：查看加权均分(标准/百分制)、专业排名、卡绩分析及”GPA计算器”。
                </div>
                <div class="gm-help-step">
                    点击悬浮球菜单 <span class="gm-tag gm-tag-gray"><svg viewBox="0 0 24 24" width="12" height="12" style="vertical-align:-1px;" stroke="currentColor" stroke-width="2" fill="none" stroke-linecap="round" stroke-linejoin="round"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path><polyline points="7 10 12 15 17 10"></polyline><line x1="12" y1="15" x2="12" y2="3"></line></svg> 导出成绩</span>：生成包含<b>教学班排名</b>的 Excel 成绩单。
                </div>
                <div class="gm-help-step">
                    <b>学生画像增强：</b>进入”学生画像”页面，脚本会在顶部增加加权平均分的卡片，并优化底部”计划外课程”的表格显示。
                </div>
            </div>

            <!-- 模块：选课助手 -->
            <div class="gm-help-section">
                <div class="gm-help-sec-title" style="border-color:#409EFF;">
                    选课助手与数据存档
                </div>
                <div class="gm-help-step">
                    <b>第1步：同步选课数据</b><br>
                    在”全校开课查询”页面，右侧会自动弹出<b>选课助手面板</b>。在下拉框中选择学期，点击 <span class="gm-tag gm-tag-blue">存储选定学期课程信息</span> 按钮，脚本会抓取该学期的课程信息。数据会存储到<b>存档系统</b>中，该系统支持在线同步、本地上传、数据导出等功能。
                </div>
                <div class="gm-help-step">
                    <b>第2步：管理存档</b><br>
                    点击悬浮球菜单 <span class="gm-tag gm-tag-blue"><svg viewBox="0 0 24 24" width="12" height="12" style="vertical-align:-1px;" stroke="currentColor" stroke-width="2" fill="none" stroke-linecap="round" stroke-linejoin="round"><polyline points="23 4 23 10 17 10"></polyline><polyline points="1 20 1 14 7 14"></polyline><path d="M3.51 9a9 9 0 0 1 14.85-3.36L23 10M1 14l4.64 4.36A9 9 0 0 0 20.49 15"></path></svg> 同步最新选课数据</span>，可以：<br>
                    - 查看三个存档槽位的数据状态<br>
                    - 选择某个存档作为当前使用的数据源<br>
                    - 导出存档数据或上传本地数据<br>
                    - 接收其他用户分享的选课数据
                </div>
                <div class="gm-help-step">
                    <b>第3步：课程关注与排课</b><br>
                    在”全校开课查询”页面，点击课程左侧的 <span class="gm-tag gm-tag-red"><svg viewBox="0 0 24 24" width="12" height="12" style="vertical-align:-1px;" fill="#F56C6C" stroke="none"><path d="M20.84 4.61a5.5 5.5 0 0 0-7.78 0L12 5.67l-1.06-1.06a5.5 5.5 0 0 0-7.78 7.78l1.06 1.06L12 21.23l7.78-7.78 1.06-1.06a5.5 5.5 0 0 0 0-7.78z"></path></svg></span> 收藏课程。<br>
                    在”培养方案”页面，课程代码会自动高亮并显示最新学期的教学班信息，点击课程旁的 <span class="gm-tag gm-tag-red"><svg viewBox="0 0 24 24" width="12" height="12" style="vertical-align:-1px;" fill="#F56C6C" stroke="none"><path d="M20.84 4.61a5.5 5.5 0 0 0-7.78 0L12 5.67l-1.06-1.06a5.5 5.5 0 0 0-7.78 7.78l1.06 1.06L12 21.23l7.78-7.78 1.06-1.06a5.5 5.5 0 0 0 0-7.78z"></path></svg></span> 收藏课程。<br>
                    然后打开悬浮球菜单 <span class="gm-tag gm-tag-blue"><svg viewBox="0 0 24 24" width="12" height="12" style="vertical-align:-1px;" fill="#409EFF" stroke="none"><path d="M20.84 4.61a5.5 5.5 0 0 0-7.78 0L12 5.67l-1.06-1.06a5.5 5.5 0 0 0-7.78 7.78l1.06 1.06L12 21.23l7.78-7.78 1.06-1.06a5.5 5.5 0 0 0 0-7.78z"></path></svg> 课程关注列表</span>，切换右上角到 <b>”课表视图”</b>，可直观查看当前已关注课程的课表情况。
                </div>
                <div class="gm-help-step">
                    <b>第4步：正式选课</b><br>
                    进入”选课”页面：<br>
                    - <b>意愿值选课：</b>显示存档中的”历史余量/上限”数据。<br>
                    - <b>直选选课：</b>自动显示”待释放名额”。<br>
                    - <b>关注课程高亮：</b>已关注的课程背景会高亮显示，方便快速定位。
                </div>
            </div>

            <!-- 模块：实用工具 -->
            <div class="gm-help-section">
                <div class="gm-help-sec-title" style="border-color:#67C23A;">
                    实用工具
                </div>
                <div class="gm-help-step">
                    <b>一键自动评教：</b>进入评教页面，点击右上角的 <span class="gm-tag gm-tag-blue">打开自动评教</span> 按钮。
                    支持任意给分评教或指定给分评教。
                </div>
            </div>

            <div style="text-align:center; padding:15px; color:#c0c4cc; font-size:12px;">
                当前版本: ${GM_info.script.version} &nbsp;|&nbsp; 祝您学业进步
            </div>
        </div>
    `;

        // 4. 组装与事件绑定
        overlay.appendChild(modal);
        document.body.appendChild(overlay);

        // 关闭逻辑
        const closeFn = () => {
            overlay.style.opacity = '0';
            setTimeout(() => overlay.remove(), 200);
        };

        modal.querySelector('.gm-help-close').onclick = closeFn;
        overlay.onclick = (e) => {
            if (e.target === overlay) closeFn();
        };
    }

    /**
     * 处理跳转至评教界面
     */
    function handleJumpToEvaluation() {
        hideMenu(); // 确保 hideMenu 函数在此作用域内可见
        if (confirm("即将跳转至“学生总结性评教”页面，是否继续？")) {
            // ... (跳转逻辑)
            let evalLink = document.querySelector('a[onclick*="evaluation-student"]') ||
                document.querySelector('a[href*="evaluation-student"]') ||
                document.querySelector('a[data-text="学生总结性评教"]');

            // 尝试在顶层窗口查找（应对 iframe 情况）
            if (!evalLink && window.top !== window.self) {
                try {
                    evalLink = window.top.document.querySelector('a[onclick*="evaluation-student"]') ||
                        window.top.document.querySelector('a[data-text="学生总结性评教"]');
                } catch (e) { }
            }

            if (evalLink) {
                evalLink.click();
            } else {
                // 强制跳转作为后备方案
                window.location.href = 'https://jwxt.nwpu.edu.cn/evaluation-student-frontend/#/byTask';
            }
        }
    }


    // =-=-=-=-=-=-=-=-=-=-=-=-= 2.1 课程关注列表 =-=-=-=-=-=-=-=-=-=-=-=-=
    /**
     * 展示已关注课程列表
     */
    function handleShowFollowedClick() {
        hideMenu();
        Logger.log("2.1", "正在初始化课程关注列表...");

        // 模块 1: 课程数据解析工具 (CourseParser)
        const CourseParser = {
            cnToNumber(str) {
                const map = { '一': 1, '二': 2, '三': 3, '四': 4, '五': 5, '六': 6, '七': 7, '八': 8, '九': 9, '十': 10, '十一': 11, '十二': 12, '十三': 13, '十四': 14, '日': 7, '天': 7 };
                const clean = str.replace(/[^\d一二三四五六七八九十日天]/g, '');
                return map[clean] || parseInt(clean) || 0;
            },
            formatWeekSet(weekSet) {
                if (!weekSet || weekSet.size === 0) return "";
                const weeks = Array.from(weekSet).sort((a, b) => a - b);
                const ranges = [];
                let start = weeks[0], prev = weeks[0];
                for (let i = 1; i < weeks.length; i++) {
                    if (weeks[i] === prev + 1) {
                        prev = weeks[i];
                    } else {
                        ranges.push(start === prev ? `${start}` : `${start}-${prev}`);
                        start = weeks[i];
                        prev = weeks[i];
                    }
                }
                ranges.push(start === prev ? `${start}` : `${start}-${prev}`);
                return ranges.join(',') + "周";
            },
            parseActiveWeeks(weekStr) {
                const activeWeeks = new Set();
                if (!weekStr) return activeWeeks;
                let content = weekStr.replace(/周/g, '').replace(/[\(\[\{（].*?[\)\]\}）]/g, '');
                const isOdd = weekStr.includes('单');
                const isEven = weekStr.includes('双');
                content.split(/[,，]/).forEach(part => {
                    const rangeMatch = part.match(/(\d+)\s*[-~～]\s*(\d+)/);
                    if (rangeMatch) {
                        const start = parseInt(rangeMatch[1]);
                        const end = parseInt(rangeMatch[2]);
                        for (let i = start; i <= end; i++) {
                            if (isOdd && i % 2 === 0) continue;
                            if (isEven && i % 2 !== 0) continue;
                            activeWeeks.add(i);
                        }
                    } else {
                        const single = parseInt(part);
                        if (!isNaN(single)) activeWeeks.add(single);
                    }
                });
                return activeWeeks;
            },
            parseTimeAndPlace(timeStr) {
                if (!timeStr || timeStr === '-' || timeStr === '') return [];
                const results = [];
                const cleanStr = timeStr.replace(/<br\s*\/?>/gi, ';').replace(/\n/g, ';');
                cleanStr.split(/[;；]/).forEach(seg => {
                    seg = seg.trim();
                    if (!seg) return;
                    const weekMatch = seg.match(/([\d,\-~～]+)周(?:\([单双]\))?/);
                    const activeWeeks = weekMatch ? this.parseActiveWeeks(weekMatch[0]) : null;
                    const dayMatch = seg.match(/[周星][期]?[一二三四五六日天1-7]/);
                    let day = 0;
                    if (dayMatch) day = this.cnToNumber(dayMatch[0].replace(/[周星期]/g, ''));
                    const nodeMatch = seg.match(/(?:第)?(\d+|[一二三四五六七八九十]+)(?:[节\s]*[-~～][第\s]*(\d+|[一二三四五六七八九十]+))?节/);
                    if (day > 0 && nodeMatch) {
                        const startNode = this.cnToNumber(nodeMatch[1]);
                        const endNode = nodeMatch[2] ? this.cnToNumber(nodeMatch[2]) : startNode;
                        let location = seg.replace(weekMatch ? weekMatch[0] : '', '').replace(dayMatch ? dayMatch[0] : '', '').replace(nodeMatch[0], '').trim().replace(/^[\s,，]+|[\s,，]+$/g, '');
                        if (startNode > 0 && startNode <= 14) {
                            results.push({ day, startNode, endNode, activeWeeks, location, rawInfo: seg });
                        }
                    }
                });
                return results;
            }
        };

        // 模块 2: 视图渲染器 (CourseParser)
        const ViewRenderer = {
            getStyle(name) {
                const palettes = [
                    { bg: 'rgba(111, 176, 243, 0.2)', border: 'rgb(111, 176, 243)' },
                    { bg: 'rgba(154, 166, 189, 0.2)', border: 'rgb(154, 166, 189)' },
                    { bg: 'rgba(240, 200, 109, 0.2)', border: 'rgb(240, 200, 109)' },
                    { bg: 'rgba(56, 200, 180, 0.2)', border: 'rgb(56, 200, 180)' },
                    { bg: 'rgba(244, 144, 96, 0.2)', border: 'rgb(244, 144, 96)' },
                    { bg: 'rgba(121, 150, 202, 0.2)', border: 'rgb(121, 150, 202)' },
                    { bg: 'rgba(218, 196, 165, 0.2)', border: 'rgb(218, 196, 165)' },
                    { bg: 'rgba(253, 171, 154, 0.2)', border: 'rgb(253, 171, 154)' },
                    { bg: 'rgba(255, 117, 117, 0.2)', border: 'rgb(255, 117, 117)' },
                    { bg: 'rgba(169, 206, 149, 0.2)', border: 'rgb(169, 206, 149)' }
                ];
                let hash = 0;
                for (let i = 0; i < name.length; i++) hash = name.charCodeAt(i) + ((hash << 5) - hash);
                return palettes[Math.abs(hash) % palettes.length];
            },

            renderList(courses, container) {
                if (!courses.length) {
                    container.innerHTML = `<div class="gm-empty-state"><p>暂无相关课程</p></div>`;
                    return;
                }
                let html = `
                <table class="gm-course-table">
                    <thead><tr><th width="120">代码</th><th>课程名称</th><th width="12%">学期</th><th width="15%">教师</th><th width="50" align="center">学分</th><th>时间/地点</th><th width="70" align="center">操作</th></tr></thead>
                    <tbody>
            `;
                courses.forEach(c => {
                    const tpStr = c.timeAndPlace ? esc(String(c.timeAndPlace)).replace(/[;；]\s*/g, '<br>') : '-';
                    html += `
                    <tr>
                        <td><span class="gm-code-badge">${esc(c.code)}</span></td>
                        <td>${esc(c.name)}</td>
                        <td style="color:#999;font-size:12px;">${esc(c.semester || '历史')}</td>
                        <td>${esc(c.teachers)}</td>
                        <td align="center">${esc(c.credits)}</td>
                        <td style="font-size:12px;line-height:1.4;">${tpStr}</td>
                        <td align="center"><button class="gm-btn-unfollow" data-id="${esc(c.id)}">取消</button></td>
                    </tr>`;
                });
                html += `</tbody></table>`;
                container.innerHTML = html;
            },

            renderTimetable(courses, container, targetWeek) {
                const timeSlots = [
                    { range: [1, 2] }, { range: [3, 4] }, { range: [5, 6] }, { range: [7, 8] }, { range: [9, 10] }, { range: [11, 12] }, { range: [13] }
                ];

                let html = `<table class="gm-timetable"><thead><tr><th width="50" style="background:#f5f7fa;"></th><th width="13.5%">星期一</th><th width="13.5%">星期二</th><th width="13.5%">星期三</th><th width="13.5%">星期四</th><th width="13.5%">星期五</th><th width="13.5%">星期六</th><th width="13.5%">星期日</th></tr></thead><tbody>`;

                timeSlots.forEach((slot, index) => {
                    const startNode = slot.range[0];
                    let slotBg = '#f9fafc'; // 默认颜色

                    if (startNode <= 4) {
                        slotBg = '#e6f7ff'; // 1-4节 (上午): 浅蓝
                    } else if (startNode <= 6) {
                        slotBg = '#fff7e6'; // 5-6节 (下午1): 浅橙
                    } else if (startNode <= 10) {
                        slotBg = '#f6ffed'; // 7-10节 (下午2+晚1): 浅绿
                    } else {
                        slotBg = '#f4f4f5'; // 11-13节 (晚2): 浅灰
                    }

                    let periodHtml = `<div class="gm-period-wrapper">`;
                    slot.range.forEach((num, idx) => {
                        const borderStyle = (idx < slot.range.length - 1) ? 'border-bottom: 1px solid rgba(0,0,0,0.06);' : '';
                        periodHtml += `<div class="gm-period-num" style="${borderStyle}">${num}</div>`;
                    });
                    periodHtml += `</div>`;

                    html += `<tr><td class="gm-tt-period" style="background:${slotBg}">${periodHtml}</td>`;

                    for (let day = 1; day <= 7; day++) {
                        const coursesInSlotMap = new Map();

                        courses.forEach(course => {
                            const segments = CourseParser.parseTimeAndPlace(course.timeAndPlace);
                            segments.forEach(seg => {
                                if (seg.day !== day) return;
                                if (targetWeek !== 'all') {
                                    const weekNum = parseInt(targetWeek);
                                    if (seg.activeWeeks && seg.activeWeeks.size > 0 && !seg.activeWeeks.has(weekNum)) return;
                                }

                                if (seg.startNode <= slot.range[slot.range.length - 1] && seg.endNode >= slot.range[0]) {
                                    const key = course.id;
                                    if (!coursesInSlotMap.has(key)) {
                                        coursesInSlotMap.set(key, {
                                            ...course,
                                            mergedWeeks: new Set(),
                                            segLocation: seg.location,
                                            detailSegments: []
                                        });
                                    }
                                    const existing = coursesInSlotMap.get(key);
                                    if (seg.activeWeeks) seg.activeWeeks.forEach(w => existing.mergedWeeks.add(w));
                                    existing.detailSegments.push(seg.rawInfo);
                                }
                            });
                        });

                        html += `<td style="vertical-align: top; padding: 2px;">`;
                        html += `<div class="gm-tt-cell-wrapper">`;

                        if (coursesInSlotMap.size > 0) {
                            coursesInSlotMap.forEach(item => {
                                const style = this.getStyle(item.name);
                                let weekInfoStr = CourseParser.formatWeekSet(item.mergedWeeks);
                                if (weekInfoStr === "") weekInfoStr = "未知周次";
                                const uniqueDetails = [...new Set(item.detailSegments)];
                                const tooltip = `${item.name}\n${item.teachers}\n----------------\n${uniqueDetails.join('\n')}`;

                                html += `
                                <div class="gm-tt-course-block"
                                     style="background: ${style.bg}; border-left: 3px solid ${style.border};"
                                     title="${esc(tooltip)}">
                                    <div class="gm-tt-name">${esc(item.name)}</div>
                                    <div class="gm-tt-info">@${esc(item.segLocation)}</div>
                                    <div class="gm-tt-info" style="display:flex;justify-content:space-between;align-items:center;margin-top:2px;">
                                        <span style="overflow:hidden;text-overflow:ellipsis;white-space:nowrap;max-width:60%;" title="${esc(item.teachers)}">${esc(item.teachers)}</span>
                                        <span class="gm-tt-tag">${weekInfoStr}</span>
                                    </div>
                                </div>
                            `;
                            });
                        }
                        html += `</div></td>`;
                    }
                    html += `</tr>`;
                });

                html += `</tbody></table>`;
                const weekText = targetWeek === 'all' ? '全部周次' : `第 ${targetWeek} 周`;
                if (courses.length > 0) {
                    html += `<div class="gm-tt-footer">当前展示：${weekText}</div>`;
                } else {
                    html = `<div class="gm-empty-state"><p>${weekText} 暂无课程</p></div>`;
                }
                container.innerHTML = html;
            }
        };

        // 模块 3: CSS 样式注入
        const styleId = 'gm-followed-modal-style';
        if (!document.getElementById(styleId)) {
            const style = getOrCreateStyleEl(styleId);
            style.textContent = `
            .gm-modal-overlay { position: fixed; top: 0; left: 0; width: 100%; height: 100%; background-color: rgba(0, 0, 0, 0.6); z-index: 10005; display: flex; align-items: center; justify-content: center; backdrop-filter: blur(2px); }
            .gm-modal-content { background-color: #fff; border-radius: 12px; width: 95%; max-width: 1200px; height: 90vh; max-height: 950px; box-shadow: 0 10px 25px rgba(0, 0, 0, 0.2); display: flex; flex-direction: column; overflow: hidden; font-family: "Helvetica Neue", Helvetica, Arial, sans-serif; animation: gmFadeIn 0.2s ease-out; }
            @keyframes gmFadeIn { from { opacity: 0; transform: scale(0.99); } to { opacity: 1; transform: scale(1); } }
            .gm-modal-header { padding: 0 20px; border-bottom: 1px solid #eee; background: #fff; display: flex; justify-content: space-between; align-items: center; flex-shrink: 0; height: 50px; }
            .gm-modal-title { font-size: 16px; font-weight: bold; color: #333; display: flex; align-items: center; gap: 8px; }
            .gm-modal-close { border: none; background: none; font-size: 24px; color: #999; cursor: pointer; padding: 0 10px; display:flex; align-items:center; }
            .gm-tabs { display: flex; gap: 20px; margin-left: 30px; height: 100%; }
            .gm-tab-item { display: flex; align-items: center; height: 100%; font-size: 14px; color: #666; cursor: pointer; border-bottom: 3px solid transparent; transition: all 0.2s; padding: 0 5px; font-weight: 500; }
            .gm-tab-item.active { color: #409EFF; border-bottom-color: #409EFF; }
            .gm-filter-bar { padding: 10px 20px; background: #fff; border-bottom: 1px solid #f0f0f0; display: flex; align-items: center; justify-content: space-between; flex-shrink: 0; gap: 15px;}
            .gm-filter-group { display: flex; align-items: center; gap: 10px; }
            .gm-filter-label { font-size: 13px; color: #606266; }
            .gm-filter-select { padding: 5px 10px; font-size: 13px; border: 1px solid #dcdfe6; border-radius: 3px; color: #606266; outline: none; background-color: white; }
            .gm-btn-clear-all { padding: 5px 12px; font-size: 13px; border-radius: 3px; border: 1px solid #f56c6c; color: #f56c6c; background: #fff; cursor: pointer; transition: all 0.2s; display: flex; align-items: center; gap: 5px; }
            .gm-btn-clear-all:hover { background: #f56c6c; color: #fff; }
            .gm-modal-body { flex: 1; overflow: hidden; position: relative; display: flex; flex-direction: column; background: #fff; }
            .gm-view-container { flex: 1; overflow-y: auto; padding: 0 20px 20px 20px; display: none; height: 100%; }
            .gm-view-container.active { display: block; }
            .gm-course-table { width: 100%; border-collapse: separate; border-spacing: 0; margin-top: 10px; }
            .gm-course-table th { position: sticky; top: 0; background: #fff; z-index: 10; padding: 12px 10px; text-align: left; font-size: 13px; color: #909399; font-weight: bold; border-bottom: 2px solid #f0f0f0; }
            .gm-course-table td { padding: 10px; font-size: 13px; color: #606266; border-bottom: 1px solid #ebeef5; vertical-align: middle; line-height: 1.4; }
            .gm-code-badge { background: #f4f4f5; color: #909399; padding: 2px 6px; border-radius: 3px; font-size: 12px; }
            .gm-btn-unfollow { padding: 4px 10px; font-size: 12px; border-radius: 3px; border: 1px solid #fab6b6; color: #f56c6c; background: #fef0f0; cursor: pointer; }
            .gm-timetable { width: 100%; border-collapse: collapse; table-layout: fixed; margin-top: 10px; border: 1px solid #e0e0e0; font-size: 12px; }
            .gm-timetable th { background: #f8f9fa; color: #333; font-weight: bold; padding: 8px; text-align: center; border: 1px solid #ebeef5; height: 36px; }
            .gm-timetable td { border: 1px solid #ebeef5; height: auto; }
            .gm-tt-period { position: relative; padding: 0 !important; vertical-align: top; width: 50px; height: 1px; }
            .gm-period-wrapper { position: absolute; top: 0; left: 0; right: 0; bottom: 0; display: flex; flex-direction: column; }
            .gm-period-num { flex: 1; display: flex; align-items: center; justify-content: center; font-size: 13px; color: #555; font-weight: bold; width: 100%; }
            .gm-tt-cell-wrapper { width: 100%; min-height: 60px; display: flex; flex-direction: column; gap: 4px; padding: 4px; box-sizing: border-box; }
            .gm-tt-course-block { padding: 6px; font-size: 12px; line-height: 1.35; color: #333; cursor: pointer; border-radius: 0; overflow: hidden; }
            .gm-tt-course-block:hover { box-shadow: 0 4px 12px rgba(0,0,0,0.15); transform: scale(1.01); z-index: 5; transition: all 0.1s; }
            .gm-tt-name { font-weight: bold; color: #000; margin-bottom: 3px; font-size: 13px; }
            .gm-tt-info { color: #555; font-size: 12px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
            .gm-tt-tag { opacity: 0.9; font-size: 11px; color: #333; background: rgba(255,255,255,0.6); padding: 1px 4px; border-radius: 3px; flex-shrink: 0; }
            .gm-tt-footer { margin-top:10px; font-size:12px; color:#606266; text-align:right; padding-right:10px;}
            .gm-empty-state { display: flex; flex-direction: column; align-items: center; justify-content: center; height: 100%; color: #909399; text-align: center; padding-top: 60px;}
        `;
            document.head.appendChild(style);
        }

        // 模块 4: 初始化
        const followedCourses = FollowManager.getList();
        let courseList = Object.values(followedCourses);
        const allSemesters = [...new Set(courseList.map(c => c.semester || '历史关注'))].sort().reverse();

        let semesterOptions = `<option value="all">显示全部学期</option>`;
        allSemesters.forEach(sem => { semesterOptions += `<option value="${sem}">${sem}</option>`; });

        let weekOptions = `<option value="all">显示全部周次</option>`;
        for (let i = 1; i <= 20; i++) weekOptions += `<option value="${i}">第 ${i} 周</option>`;

        const modalHTML = `
        <div class="gm-modal-overlay" id="gm-modal-overlay">
            <div class="gm-modal-content">
                <div class="gm-modal-header">
                    <div style="display:flex; align-items:center;">
                        <div class="gm-modal-title"><svg viewBox="0 0 24 24" width="18" height="18" style="vertical-align:-3px; margin-right:6px;" fill="#f56c6c" stroke="none"><path d="M20.84 4.61a5.5 5.5 0 0 0-7.78 0L12 5.67l-1.06-1.06a5.5 5.5 0 0 0-7.78 7.78l1.06 1.06L12 21.23l7.78-7.78 1.06-1.06a5.5 5.5 0 0 0 0-7.78z"></path></svg>课程关注列表</div>
                        <div class="gm-tabs">
                            <div class="gm-tab-item active" data-tab="list">列表视图</div>
                            <div class="gm-tab-item" data-tab="timetable">课表视图</div>
                        </div>
                    </div>
                    <button class="gm-modal-close" id="gm-modal-close">×</button>
                </div>
                <div class="gm-modal-body">
                    <div class="gm-filter-bar">
                        <div class="gm-filter-group">
                            <span class="gm-filter-label">学期:</span>
                            <select id="gm-semester-select" class="gm-filter-select" style="min-width:140px;">${semesterOptions}</select>
                            <div id="gm-week-filter-container" style="display:none; align-items:center; margin-left:15px;">
                                <span class="gm-filter-label">周次:</span>
                                <select id="gm-week-select" class="gm-filter-select" style="min-width:100px; margin-left:5px;">${weekOptions}</select>
                            </div>
                        </div>

                        <div class="gm-right-actions" style="display:flex; align-items:center; gap:15px;">
                            <span style="font-size:13px; color:#606266; font-weight:bold;">
                                总学分: <span id="gm-total-credits" style="color:#409EFF">0</span>
                            </span>
                            <button id="gm-btn-clear-all" class="gm-btn-clear-all">清空当前</button>
                        </div>
                    </div>
                    <div class="gm-view-container active" id="gm-view-list"><div id="gm-table-wrapper"></div></div>
                    <div class="gm-view-container" id="gm-view-timetable"><div id="gm-timetable-wrapper"></div></div>
                </div>
            </div>
        </div>
    `;

        const existingOverlay = document.getElementById('gm-modal-overlay');
        if (existingOverlay) existingOverlay.remove();
        const wrapper = document.createElement('div');
        wrapper.innerHTML = modalHTML;
        document.body.appendChild(wrapper.firstElementChild);

        // 状态管理
        const state = { semester: 'all', week: 'all', currentTab: 'list' };

        const refreshView = () => {
            const filtered = courseList.filter(c => state.semester === 'all' || (c.semester || '历史关注') === state.semester);
            filtered.sort((a, b) => {
                const semA = a.semester || ''; const semB = b.semester || '';
                if (semA !== semB) return semB.localeCompare(semA);
                return a.code.localeCompare(b.code);
            });

            // 【修改点】 计算并更新总学分
            let totalCredits = 0;
            filtered.forEach(c => {
                const credit = parseFloat(c.credits);
                if (!isNaN(credit)) {
                    totalCredits += credit;
                }
            });
            const creditSpan = document.getElementById('gm-total-credits');
            if (creditSpan) {
                creditSpan.innerText = totalCredits % 1 === 0 ? totalCredits : totalCredits.toFixed(1);
            }

            if (state.currentTab === 'list') {
                ViewRenderer.renderList(filtered, document.getElementById('gm-table-wrapper'));
            } else {
                ViewRenderer.renderTimetable(filtered, document.getElementById('gm-timetable-wrapper'), state.week);
            }

            const clearBtn = document.getElementById('gm-btn-clear-all');
            const btnText = state.semester === 'all' ? '清空全部' : '清空当前学期';
            clearBtn.innerHTML = `<span style="margin-left:4px">${btnText}</span>`;
            clearBtn.style.opacity = filtered.length === 0 ? '0.5' : '1';
            clearBtn.style.pointerEvents = filtered.length === 0 ? 'none' : 'auto';
        };

        // 事件绑定
        const tabs = document.querySelectorAll('.gm-tab-item');
        const views = document.querySelectorAll('.gm-view-container');
        const weekFilterContainer = document.getElementById('gm-week-filter-container');

        tabs.forEach(tab => {
            tab.onclick = () => {
                tabs.forEach(t => t.classList.remove('active'));
                views.forEach(v => v.classList.remove('active'));
                tab.classList.add('active');
                document.getElementById(`gm-view-${tab.dataset.tab}`).classList.add('active');
                state.currentTab = tab.dataset.tab;
                weekFilterContainer.style.display = (state.currentTab === 'timetable') ? 'flex' : 'none';
                refreshView();
            };
        });

        document.getElementById('gm-semester-select').onchange = (e) => { state.semester = e.target.value; refreshView(); };
        document.getElementById('gm-week-select').onchange = (e) => { state.week = e.target.value; refreshView(); };
        const closeModal = () => document.getElementById('gm-modal-overlay').remove();
        document.getElementById('gm-modal-close').onclick = closeModal;
        document.getElementById('gm-modal-overlay').onclick = (e) => { if (e.target.id === 'gm-modal-overlay') closeModal(); };

        document.getElementById('gm-btn-clear-all').onclick = () => {
            const targetName = state.semester === 'all' ? '所有' : state.semester;
            GmDialog.confirm({
                title: '取消关注确认',
                message: `确定要取消关注【${targetName}】下的所有课程吗？`,
                type: 'warn',
                confirmText: '确定取消'
            }).then(ok => {
                if (!ok) return;
                const idsToRemove = courseList
                    .filter(c => state.semester === 'all' || (c.semester || '历史关注') === state.semester)
                    .map(c => c.id);
                idsToRemove.forEach(id => FollowManager.remove(id));
                courseList = courseList.filter(c => !idsToRemove.includes(c.id));
                if (courseList.length === 0) { closeModal(); handleShowFollowedClick(); }
                else refreshView();
            });
        };

        const btnContainer = document.querySelector('.gm-modal-body');
        btnContainer.addEventListener('click', (e) => {
            if (e.target.classList.contains('gm-btn-unfollow')) {
                const id = e.target.dataset.id;
                if (confirm('确定不再关注此课程吗？')) {
                    // 1. 从存储中移除
                    FollowManager.remove(id);

                    // 2. 重新从存储读取最新全量列表
                    courseList = Object.values(FollowManager.getList());

                    // 3. 刷新视图 (ViewRenderer 会自动处理空列表的情况)
                    refreshView();
                }
            }
        });

        refreshView();
    }

    // ----------------- 2.2 导出成绩 -----------------

    // 导出面板样式注入（独立函数，initExportUI 与 showSemesterCheckboxes 共用，幂等）
    function injectExportStyles() {
        if (document.getElementById('export-ui-styles')) return;
        const style = getOrCreateStyleEl('export-ui-styles');
        style.textContent = `
    .semester-bg-overlay { position: fixed; top: 0; left: 0; width: 100%; height: 100%; background: rgba(0,0,0,0.4); backdrop-filter: blur(4px); z-index: 10001; display: none; transition: all 0.3s; }
    .semester-checkbox-container { position: fixed; top: 50%; left: 50%; transform: translate(-50%, -50%); z-index: 10002; background-color: #ffffff; padding: 24px; border-radius: 12px; box-shadow: 0 10px 40px rgba(0, 0, 0, 0.15); display: none; width: 450px; border: 1px solid #f0f0f0; font-family: "Helvetica Neue", Helvetica, "PingFang SC", "Microsoft YaHei", Arial, sans-serif; }
    .semester-checkbox-container h3 { margin: 0 0 20px 0; font-size: 18px; color: #1f1f1f; display: flex; align-items: center; gap: 8px; font-weight: 600; }
    .gm-semester-list { display: grid; grid-template-columns: 1fr 1fr; gap: 10px; max-height: 320px; overflow-y: auto; padding: 4px; margin-bottom: 20px; }
    .gm-semester-item { display: flex; align-items: center; padding: 10px 12px; background: #f5f7fa; border-radius: 8px; cursor: pointer; transition: all 0.2s; border: 1px solid transparent; font-size: 14px; color: #444; user-select: none; }
    .gm-semester-item:hover { background: #eef5fe; border-color: #b3d8ff; color: #409EFF; }
    .gm-semester-item input[type='checkbox'] { margin-right: 10px; width: 16px; height: 16px; cursor: pointer; }
    .button-container { display: flex; justify-content: flex-end; gap: 12px; border-top: 1px solid #f0f0f0; padding-top: 20px; }
    .semester-checkbox-container button { padding: 8px 18px; border-radius: 6px; cursor: pointer; font-size: 14px; font-weight: 500; transition: all 0.2s; border: none; }
    .confirm-export-button { background-color: #409EFF; color: white; }
    .confirm-export-button:hover { background-color: #66b1ff; transform: translateY(-1px); }
    .select-all-button { background-color: #f0f2f5; color: #595959; border: 1px solid #d9d9d9 !important; }
    .cancel-button { background-color: transparent; color: #8c8c8c; }
    .gm-semester-list::-webkit-scrollbar { width: 6px; }
    .gm-semester-list::-webkit-scrollbar-thumb { background: #dcdfe6; border-radius: 3px; }
`;
        document.head.appendChild(style);
    }

    function initExportUI() {
        injectExportStyles();
        let bgOverlay = document.createElement("div"); bgOverlay.className = "semester-bg-overlay"; document.body.appendChild(bgOverlay);
        semesterCheckboxContainer = document.createElement("div"); semesterCheckboxContainer.className = "semester-checkbox-container"; document.body.appendChild(semesterCheckboxContainer);
        bgOverlay.addEventListener('click', () => { semesterCheckboxContainer.style.display = "none"; });
        Object.defineProperty(semesterCheckboxContainer.style, 'display', { set: function (val) { bgOverlay.style.display = val; this.setProperty('display', val); }, get: function () { return this.getPropertyValue('display'); } });
    }

    function showSemesterCheckboxes(semesterNames) {
        if (!document.getElementById('export-ui-styles')) {
            injectExportStyles(); // 封装样式注入逻辑
        }

        Logger.log("2.2", "开始导出成绩...");
        semesterCheckboxContainer.innerHTML = "";
        const title = document.createElement("h3");
        title.textContent = "选择要导出的学期";
        semesterCheckboxContainer.appendChild(title);

        const listDiv = document.createElement("div");
        listDiv.className = "gm-semester-list";

        semesterNames.forEach((semesterName) => {
            const label = document.createElement("label");
            label.className = "gm-semester-item";
            const checkbox = document.createElement("input");
            checkbox.type = "checkbox";
            checkbox.value = semesterName;
            checkbox.checked = true;
            label.appendChild(checkbox);
            label.appendChild(document.createTextNode(semesterName));
            listDiv.appendChild(label);
        });
        semesterCheckboxContainer.appendChild(listDiv);

        const buttonContainer = document.createElement("div");
        buttonContainer.className = "button-container";

        const selectAllButton = document.createElement("button");
        selectAllButton.textContent = "全选/反选";
        selectAllButton.className = "select-all-button";
        selectAllButton.onclick = () => {
            const checkboxes = semesterCheckboxContainer.querySelectorAll("input[type='checkbox']");
            const isAllChecked = Array.from(checkboxes).every(c => c.checked);
            checkboxes.forEach(c => { c.checked = !isAllChecked; });
        };

        const cancelButton = document.createElement("button");
        cancelButton.textContent = "取消";
        cancelButton.className = "cancel-button";
        cancelButton.onclick = () => { semesterCheckboxContainer.style.display = "none"; };

        const confirmExportButton = document.createElement("button");
        confirmExportButton.textContent = "导出至 Excel";
        confirmExportButton.className = "confirm-export-button";
        confirmExportButton.onclick = () => {
            const selectedSemesters = Array.from(semesterCheckboxContainer.querySelectorAll("input[type='checkbox']:checked")).map(c => c.value);
            const cachedData = getCachedData();
            if (cachedData) {
                const filteredGrades = cachedData.allGrades.filter(grade => selectedSemesters.includes(grade.学期));
                exportToExcel(filteredGrades);
            }
            semesterCheckboxContainer.style.display = "none";
        };

        buttonContainer.appendChild(selectAllButton);
        buttonContainer.appendChild(cancelButton);
        buttonContainer.appendChild(confirmExportButton);
        semesterCheckboxContainer.appendChild(buttonContainer);

        semesterCheckboxContainer.style.display = "block";
    }

    async function exportToExcel(filteredGrades) {
        if (!filteredGrades || filteredGrades.length === 0) { alert("没有选中任何成绩数据，已取消导出。"); return; }
        try {
            const workbook = new ExcelJS.Workbook();
            const worksheet = workbook.addWorksheet('课程成绩与排名');
            worksheet.addRow(["注意：由于教务系统将同一分数视为同一排名，故您的实际教学班排名可能会低于本数据。"]);
            worksheet.mergeCells('A1:H1');
            worksheet.getRow(1).font = { bold: true, color: { argb: 'FFFF0000' } };
            worksheet.getRow(1).alignment = { horizontal: 'left', vertical: 'middle' };
            const header = ['课程ID', '课程代码', '课程名称', '学分', '成绩', '绩点', '教学班排名', '学期'];
            worksheet.addRow(header);
            worksheet.getRow(2).font = { bold: true };
            worksheet.getRow(2).alignment = { horizontal: 'center', vertical: 'middle' };
            filteredGrades.forEach((grade) => {
                worksheet.addRow([grade['课程ID'], grade['课程代码'], grade['课程名称'], grade['学分'], grade['成绩'], grade['绩点'], grade['教学班排名'], grade['学期']]);
            });
            worksheet.columns = [{ width: 10 }, { width: 12 }, { width: 35 }, { width: 7 }, { width: 7 }, { width: 7 }, { width: 12 }, { width: 22 }];
            for (let i = 3; i <= worksheet.rowCount; i++) { worksheet.getRow(i).alignment = { horizontal: 'center', vertical: 'middle' }; }
            const buffer = await workbook.xlsx.writeBuffer();
            const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
            saveAs(blob, '课程成绩与排名.xlsx');
        } catch (error) { Logger.error('2.2', 'Excel生成失败', error); alert("导出Excel文件时发生错误，请查看控制台了解详情。"); }
    }

    // ----------------- 2.3 GPA分析 -----------------

    function calculateAndDisplayGPA(data) {
        Logger.log("2.3", "开始进行GPA及加权成绩分析...");
        const { allGrades, gpaRankData } = data;
        if (!allGrades || allGrades.length === 0) { alert("没有可供分析的成绩数据。"); return; }

        const chineseGradeMap = { '优秀': 4.0, '良好': 3.0, '中等': 2.0, '及格': 1.3, '不及格': 0.0, '通过': null, '不通过': 0.0 };
        const stuckGradesMap = { 94: 4.1, 89: 3.9, 84: 3.7, 80: 3.3, 77: 2.7, 74: 2.3, 71: 2.0, 67: 2.0, 63: 1.7, 59: 1.3 };
        const validGradesForGpa = [];
        let totalScoreCreditsNumericOnly = 0, totalCreditsNumericOnly = 0;
        let totalScoreCreditsWithMapping = 0, totalCreditsWithMapping = 0;

        allGrades.forEach(grade => {
            const credits = parseFloat(grade['学分']); const score = grade['成绩']; let gp = parseFloat(grade['绩点']);
            if (isNaN(credits) || credits <= 0 || grade['绩点'] === null || isNaN(gp)) return;
            let finalGp = gp;
            if (typeof score === 'string' && chineseGradeMap.hasOwnProperty(score)) { const mappedGp = chineseGradeMap[score]; if (mappedGp === null) return; finalGp = mappedGp; }
            validGradesForGpa.push({ ...grade, '学分': credits, '成绩': score, '绩点': finalGp });
            const numericScore = parseFloat(score);
            if (!isNaN(numericScore)) { totalScoreCreditsNumericOnly += numericScore * credits; totalCreditsNumericOnly += credits; }
            if (!isNaN(numericScore)) { totalScoreCreditsWithMapping += numericScore * credits; totalCreditsWithMapping += credits; } else if (typeof score === 'string' && CONSTANTS.GRADE_MAPPING_CONFIG.hasOwnProperty(score)) { totalScoreCreditsWithMapping += CONSTANTS.GRADE_MAPPING_CONFIG[score] * credits; totalCreditsWithMapping += credits; }
        });

        const weightedScoreNumeric = totalCreditsNumericOnly > 0 ? (totalScoreCreditsNumericOnly / totalCreditsNumericOnly) : 0;
        const weightedScoreWithMapping = totalCreditsWithMapping > 0 ? (totalScoreCreditsWithMapping / totalCreditsWithMapping) : 0;
        if (validGradesForGpa.length === 0) { alert("未找到可用于计算GPA的有效课程成绩。"); return; }

        const totalCreditPoints = validGradesForGpa.reduce((sum, g) => sum + (g['绩点'] * g['学分']), 0);
        const totalCredits = validGradesForGpa.reduce((sum, g) => sum + g['学分'], 0);
        const gpa = totalCredits > 0 ? (totalCreditPoints / totalCredits) : 0;
        const stuckCourses = validGradesForGpa.filter(g => stuckGradesMap.hasOwnProperty(parseFloat(g['成绩'])));

        let reportData = { gpa: gpa.toFixed(4), totalCredits: totalCredits.toFixed(2), totalCreditPoints: totalCreditPoints.toFixed(4), courseCount: validGradesForGpa.length, hasStuckCourses: stuckCourses.length > 0, weightedScoreNumeric: weightedScoreNumeric.toFixed(4), weightedScoreWithMapping: weightedScoreWithMapping.toFixed(4), gpaRankData: gpaRankData };
        if (reportData.hasStuckCourses) {
            const stuckCoursesCredits = stuckCourses.reduce((sum, c) => sum + c['学分'], 0);
            let hypotheticalTotalCreditPoints = validGradesForGpa.reduce((sum, g) => { const scoreNum = parseFloat(g['成绩']); return sum + ((stuckGradesMap[scoreNum] || g['绩点']) * g['学分']); }, 0);
            const hypotheticalGpa = totalCredits > 0 ? (hypotheticalTotalCreditPoints / totalCredits) : 0;
            Object.assign(reportData, { stuckCoursesCount: stuckCourses.length, stuckCoursesCredits: stuckCoursesCredits.toFixed(2), stuckCoursesList: stuckCourses, hypotheticalGpa: hypotheticalGpa.toFixed(4), hypotheticalTotalCreditPoints: hypotheticalTotalCreditPoints.toFixed(4) });
        }
        showGpaReportModal(reportData);
    }

    function showGpaReportModal(reportData) {
        const existingOverlay = document.querySelector('.gpa-report-overlay');
        if (existingOverlay) existingOverlay.remove();

        const styleId = 'gpa-report-modal-styles';
        if (!document.getElementById(styleId)) {
            const style = getOrCreateStyleEl(styleId);
            style.textContent = `
            .gpa-report-overlay { position: fixed; top: 0; left: 0; width: 100%; height: 100%; background-color: rgba(0, 0, 0, 0.55); backdrop-filter: blur(4px); z-index: 10005; display: flex; align-items: center; justify-content: center; animation: gmFadeIn 0.2s ease-out; }
            .gpa-report-modal { background-color: #f5f7fa; border-radius: 12px; width: 880px; max-width: 95%; box-shadow: 0 12px 32px rgba(0, 0, 0, 0.25); font-family: "Helvetica Neue", Helvetica, "PingFang SC", "Microsoft YaHei", Arial, sans-serif; display: flex; flex-direction: column; max-height: 85vh; overflow: hidden; }
            .gpa-modal-header { padding: 18px 24px; border-bottom: 1px solid #ebeef5; display: flex; justify-content: space-between; align-items: center; background: #fff; border-radius: 12px 12px 0 0; }
            .gpa-modal-title { font-size: 18px; font-weight: 600; color: #303133; display: flex; align-items: center; gap: 8px;}
            .gpa-close-btn { border: none; background: transparent; font-size: 26px; color: #909399; cursor: pointer; line-height: 1; transition: color 0.2s; padding: 0; }
            .gpa-close-btn:hover { color: #f56c6c; }
            .gpa-modal-body { padding: 24px; overflow-y: auto; background-color: #f5f7fa; }
            .gpa-modal-body::-webkit-scrollbar { width: 6px; }
            .gpa-modal-body::-webkit-scrollbar-thumb { background: #dcdfe6; border-radius: 3px; }

            .gpa-modal-grid { display: flex; gap: 20px; align-items: flex-start; }
            .gpa-column-left { flex: 5.5; display: flex; flex-direction: column; gap: 16px; }
            .gpa-column-right { flex: 4.5; display: flex; flex-direction: column; gap: 16px; }

            .current-gpa-module { background: #fff; border: 1px solid #ebeef5; border-radius: 8px; padding: 20px; box-shadow: 0 2px 12px 0 rgba(0,0,0,0.02); }
            .gpa-report-modal h3 { margin: 0 0 15px 0; font-size: 16px; color: #303133; display: flex; align-items: center; }
            .gpa-report-modal h3::before { content: ''; display: inline-block; width: 4px; height: 16px; background: #409EFF; margin-right: 8px; border-radius: 2px; }
            .gpa-report-modal p, .gpa-report-modal li { font-size: 14px; line-height: 1.8; color: #606266 !important; margin-bottom: 0;}
            .gpa-report-modal strong { color: #303133; font-weight: 600; }

            .gpa-report-modal details { border: 1px solid #ebeef5; border-radius: 8px; background-color: #fff; overflow: hidden; box-shadow: 0 2px 12px 0 rgba(0,0,0,0.02); }
            .gpa-report-modal summary { padding: 14px 18px; font-weight: 600; font-size: 15px; color: #303133; cursor: pointer; list-style: none; outline: none; transition: background 0.2s; user-select: none; }
            .gpa-report-modal summary:hover { background: #f0f7ff; color: #409EFF; }
            .gpa-report-modal summary::-webkit-details-marker { display: none; }
            .gpa-report-modal summary::before { content: ''; width: 12px; height: 12px; margin-right: 10px; display: inline-block; background: url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 24 24' fill='%23909399'%3E%3Cpath d='M8 5v14l11-7z'/%3E%3C/svg%3E") no-repeat center / contain; transition: transform 0.2s; }
            .gpa-report-modal details[open] > summary::before { transform: rotate(90deg); }
            .gpa-report-modal details[open] > summary { border-bottom: 1px solid #ebeef5; background: #fafafa; }

            .gpa-calc-card { background: #fff; border: 1px solid #ebeef5; border-radius: 8px; padding: 20px; box-shadow: 0 2px 12px 0 rgba(0,0,0,0.02); }
            .gpa-calc-card h4 { margin: 0 0 16px 0; font-size: 15px; color: #303133; font-weight: 600; display: flex; align-items: center; gap: 8px;}
            .input-group { display: flex; flex-direction: column; align-items: flex-start; margin-bottom: 15px; gap: 6px; }
            .input-group label { width: 100%; font-size: 13px; color: #606266; font-weight: 500;}
            .input-group input { width: 100%; box-sizing: border-box; padding: 9px 12px; border: 1px solid #dcdfe6; border-radius: 6px; font-size: 14px; transition: all 0.2s; outline: none; background: #fafafa;}
            .input-group input:focus { border-color: #409EFF; background: #fff; box-shadow: 0 0 0 2px rgba(64,158,255,0.1);}
            .calculate-btn { width: 100%; padding: 10px; background-color: #409EFF; color: white; border: none; border-radius: 6px; cursor: pointer; font-size: 14px; font-weight: 500; margin-top: 5px; transition: background 0.2s; }
            .calculate-btn:hover { background-color: #66b1ff; }
            .prediction-result { margin-top: 15px; font-weight: bold; text-align: center; font-size: 15px; min-height: 24px; color: #303133; background: #f0f9eb; padding: 12px; border-radius: 6px; display: none; }

            .gpa-report-modal .tooltip-q { display: inline-flex; align-items: center; justify-content: center; width: 16px; height: 16px; border-radius: 50%; background-color: #c0c4cc; color: white; font-size: 12px; cursor: help; margin-left: 6px; position: relative; }
            .gpa-report-modal .tooltip-q:hover::after { content: attr(data-gm-tooltip); position: absolute; left: 50%; bottom: 130%; transform: translateX(-50%); background-color: #303133; color: #fff; padding: 8px 12px; border-radius: 6px; font-size: 12px; font-weight: normal; white-space: pre-line; z-index: 10; width: max-content; max-width: 280px; line-height: 1.4; box-shadow: 0 4px 12px rgba(0,0,0,0.15); }
            .disclaimer { font-size: 12px; color: #909399; margin-top: 20px; text-align: center; padding-top: 15px; border-top: 1px dashed #ebeef5; }
            @keyframes gmFadeIn { from { opacity: 0; transform: scale(0.98); } to { opacity: 1; transform: scale(1); } }
        `;
            document.head.appendChild(style);
        }

        const mappingConfigString = Object.entries(CONSTANTS.GRADE_MAPPING_CONFIG).map(([key, value]) => `${key}: ${value}`).join(', ');
        const tooltipTextWithMapping = `使用百分制成绩和中文等级制分数进行计算\n您可以在脚本最上面配置参数，当前参数：\n${mappingConfigString}。`;

        const overlay = document.createElement('div');
        overlay.className = 'gpa-report-overlay';
        const modal = document.createElement('div');
        modal.className = 'gpa-report-modal';

        let contentHTML = `
        <div class="gpa-modal-header">
            <div class="gpa-modal-title"><svg viewBox="0 0 24 24" width="20" height="20" style="vertical-align:-3px; margin-right:8px;" stroke="#409EFF" stroke-width="2" fill="none" stroke-linecap="round" stroke-linejoin="round"><polyline points="22 12 18 12 15 21 9 3 6 12 2 12"></polyline></svg>GPA综合分析报告</div>
            <button class="gpa-close-btn" title="关闭">&times;</button>
        </div>
        <div class="gpa-modal-body">
            <div class="gpa-modal-grid">

                <div class="gpa-column-left">
                    <div class="current-gpa-module">
                        <h3>当前学业总览</h3>
                        <p style="line-height: 2;">
                            <strong>GPA：</strong> <strong style="color:#1A6FC4 !important; font-size:16px;">${reportData.gpa}</strong><br>
                            <strong>专业排名：</strong> ${reportData.gpaRankData.rank ?? '无数据'}<br>
                            <strong>前一名GPA：</strong> ${reportData.gpaRankData.beforeRankGpa ?? '无数据'}<br>
                            <strong>后一名GPA：</strong> ${reportData.gpaRankData.afterRankGpa ?? '无数据'}<br>
                            <strong>纳入计算课程：</strong> ${reportData.courseCount} 门<br>
                            <strong>总学分：</strong> ${reportData.totalCredits}<br>
                            <strong>总学分绩点：</strong> ${reportData.totalCreditPoints}<br>
                            <strong>加权百分制成绩：</strong> <strong style="color:#2E7D32 !important;">${reportData.weightedScoreNumeric}</strong> <span class="tooltip-q" data-gm-tooltip="仅计算百分制成绩，不含中文等级制成绩和PNP课程。">?</span><br>
                            <strong>加权百分制(含中文)：</strong> <strong style="color:#A85B00 !important;">${reportData.weightedScoreWithMapping}</strong> <span class="tooltip-q" data-gm-tooltip="${tooltipTextWithMapping}">?</span>
                        </p>
                    </div>
                    <details class="stuck-analysis-section">
                        <summary>卡绩分析</summary>
                        <div class="details-content" style="padding: 18px;">
    `;

        if (reportData.hasStuckCourses) {
            let stuckCoursesListHTML = '<ul style="padding-left: 20px; margin: 10px 0;">';
            reportData.stuckCoursesList.forEach(course => {
                stuckCoursesListHTML += `<li>${esc(course['课程名称'])} (成绩: ${esc(course['成绩'])}, 绩点: ${esc(course['绩点'])})</li>`;
            });
            stuckCoursesListHTML += '</ul>';
            contentHTML += `
            <p>发现 <strong style="color:#C53030 !important;">${reportData.stuckCoursesCount} 门</strong>卡绩科目，共计 <strong>${reportData.stuckCoursesCredits}</strong> 学分。</p>
            ${stuckCoursesListHTML}
            <p style="margin-top:10px; border-top:1px dashed #eee; padding-top:10px;">如果这些科目成绩均提高1分，您的GPA结果为：</p>
            <p>
                <strong>总学分绩点：</strong> ${reportData.hypotheticalTotalCreditPoints}<br>
                <strong>加权平均GPA：</strong> <strong style="color: #2E7D32 !important; font-size:16px;">${reportData.hypotheticalGpa}</strong>
            </p>
        `;
        } else {
            contentHTML += `<p style="color:#2E7D32 !important; font-weight:bold; text-align:center; padding: 10px 0;">[ 恭喜您！当前未发现卡绩科目 ]</p>`;
        }

        contentHTML += `
                        </div>
                    </details>
                </div>

                <div class="gpa-column-right">
                    <div class="gpa-calc-card">
                        <h4>
                            <svg viewBox="0 0 24 24" width="18" height="18" stroke="#409EFF" stroke-width="2" fill="none" stroke-linecap="round" stroke-linejoin="round"><polyline points="22 12 18 12 15 21 9 3 6 12 2 12"></polyline></svg>
                            预估 GPA
                        </h4>
                        <div class="input-group">
                            <label>下学期预计总学分：</label>
                            <input type="number" id="next-credits-a" placeholder="例如: 25">
                        </div>
                        <div class="input-group">
                            <label>期望达到的均绩(GPA)：</label>
                            <input type="number" id="next-gpa-a" step="0.01" placeholder="1.0 ~ 4.1">
                        </div>
                        <button id="calculate-prediction-btn-a" class="calculate-btn">计算预测结果</button>
                        <div id="predicted-gpa-result-a" class="prediction-result"></div>
                    </div>

                    <div class="gpa-calc-card">
                        <h4>
                            <svg viewBox="0 0 24 24" width="18" height="18" stroke="#67C23A" stroke-width="2" fill="none" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="10"></circle><circle cx="12" cy="12" r="6"></circle><circle cx="12" cy="12" r="2"></circle></svg>
                            达成目标所需均绩
                        </h4>
                        <div class="input-group">
                            <label>期望GPA：</label>
                            <input type="number" id="target-gpa-b" step="0.01" placeholder="例如: 3.80">
                        </div>
                        <div class="input-group">
                            <label>剩余未修总学分：</label>
                            <input type="number" id="next-credits-b" placeholder="例如: 20">
                        </div>
                        <button id="calculate-target-btn-b" class="calculate-btn">计算所需成绩</button>
                        <div id="target-gpa-result-b" class="prediction-result"></div>
                    </div>
                </div>

            </div>
            <p class="disclaimer">温馨提示：此结果仅供参考，基于所有已获取的成绩数据计算，可能与教务系统保研/评奖评优所用最终规则略有差异。</p>
        </div>
    `;

        modal.innerHTML = contentHTML;
        overlay.appendChild(modal);
        document.body.appendChild(overlay);

        const close = () => {
            overlay.style.opacity = '0';
            setTimeout(() => document.body.removeChild(overlay), 200);
        };
        overlay.querySelector('.gpa-close-btn').onclick = close;
        overlay.onclick = (e) => { if (e.target === overlay) close(); };

        const calculateBtnA = document.getElementById('calculate-prediction-btn-a');
        const nextCreditsInputA = document.getElementById('next-credits-a');
        const nextGpaInputA = document.getElementById('next-gpa-a');
        const resultDisplayA = document.getElementById('predicted-gpa-result-a');

        calculateBtnA.addEventListener('click', () => {
            const nextCredits = parseFloat(nextCreditsInputA.value);
            const nextGpa = parseFloat(nextGpaInputA.value);
            resultDisplayA.style.display = 'block';

            if (isNaN(nextCredits) || nextCredits <= 0 || isNaN(nextGpa) || nextGpa < 1.0 || nextGpa > 4.1) {
                resultDisplayA.innerHTML = '<svg viewBox="0 0 24 24" width="15" height="15" style="vertical-align:-2px; margin-right:5px;" stroke="#F56C6C" stroke-width="2" fill="none" stroke-linecap="round" stroke-linejoin="round"><path d="M10.29 3.86L1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0z"></path><line x1="12" y1="9" x2="12" y2="13"></line><line x1="12" y1="17" x2="12.01" y2="17"></line></svg>请输入有效的学分与GPA (GPA应在1.0-4.1之间)';
                resultDisplayA.style.color = '#F56C6C';
                resultDisplayA.style.background = '#fef0f0';
                return;
            }
            const currentTotalCredits = parseFloat(reportData.totalCredits);
            const currentTotalCreditPoints = parseFloat(reportData.totalCreditPoints);
            const predictedOverallGPA = (currentTotalCreditPoints + (nextCredits * nextGpa)) / (currentTotalCredits + nextCredits);

            resultDisplayA.style.color = '#303133';
            resultDisplayA.style.background = '#f0f9eb';
            resultDisplayA.innerHTML = `总GPA将变为<br><span style="color: #67C23A; font-size: 22px;">${predictedOverallGPA.toFixed(4)}</span>`;
        });

        const calculateBtnB = document.getElementById('calculate-target-btn-b');
        const targetGpaInputB = document.getElementById('target-gpa-b');
        const nextCreditsInputB = document.getElementById('next-credits-b');
        const resultDisplayB = document.getElementById('target-gpa-result-b');

        calculateBtnB.addEventListener('click', () => {
            const targetGpa = parseFloat(targetGpaInputB.value);
            const nextCredits = parseFloat(nextCreditsInputB.value);
            resultDisplayB.style.display = 'block';

            if (isNaN(targetGpa) || targetGpa < 1.0 || targetGpa > 4.1 || isNaN(nextCredits) || nextCredits <= 0) {
                resultDisplayB.innerHTML = '<svg viewBox="0 0 24 24" width="15" height="15" style="vertical-align:-2px; margin-right:5px;" stroke="#F56C6C" stroke-width="2" fill="none" stroke-linecap="round" stroke-linejoin="round"><path d="M10.29 3.86L1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0z"></path><line x1="12" y1="9" x2="12" y2="13"></line><line x1="12" y1="17" x2="12.01" y2="17"></line></svg>请输入有效的学分与期望GPA';
                resultDisplayB.style.color = '#F56C6C';
                resultDisplayB.style.background = '#fef0f0';
                return;
            }
            const currentTotalCredits = parseFloat(reportData.totalCredits);
            const currentTotalCreditPoints = parseFloat(reportData.totalCreditPoints);
            const requiredCreditPointsNext = (targetGpa * (currentTotalCredits + nextCredits)) - currentTotalCreditPoints;
            const requiredGpaNext = requiredCreditPointsNext / nextCredits;

            let resultHTML = `剩余课程均绩需达到<br><span style="font-size: 22px; color: ${requiredGpaNext > 4.1 ? '#F56C6C' : '#67C23A'};">${requiredGpaNext.toFixed(4)}</span>`;

            if (requiredGpaNext > 4.1) {
                resultHTML += '<br><span style="color: #F56C6C; font-size: 13px; font-weight:normal;">(目标过高，当前评分体系无法实现)</span>';
                resultDisplayB.style.background = '#fef0f0';
            } else if (requiredGpaNext < 1.0) {
                resultHTML += '<br><span style="color: #909399; font-size: 13px; font-weight:normal;">(闭着眼考都能实现)</span>';
                resultDisplayB.style.background = '#f4f4f5';
            } else {
                resultDisplayB.style.background = '#f0f9eb';
            }

            resultDisplayB.style.color = '#303133';
            resultDisplayB.innerHTML = resultHTML;
        });
    }

    // ----------------- 2.4 学生画像增强 -----------------

    function precomputeAllWeightedScores(allGrades) {
        if (!allGrades) return {};
        const scoresBySemester = {}; const gradesBySemester = {};
        allGrades.forEach(grade => { const semester = grade['学期']; if (!gradesBySemester[semester]) gradesBySemester[semester] = []; gradesBySemester[semester].push(grade); });
        const calculate = (grades) => {
            let totalScoreCredits = 0, totalCredits = 0;
            grades.forEach(grade => {
                const credits = parseFloat(grade['学分']); if (isNaN(credits) || credits <= 0) return;
                const numericScore = parseFloat(grade['成绩']);
                if (!isNaN(numericScore)) { totalScoreCredits += numericScore * credits; totalCredits += credits; }
            });
            return totalCredits > 0 ? (totalScoreCredits / totalCredits).toFixed(4) : 'N/A';
        };
        for (const semesterName in gradesBySemester) { scoresBySemester[semesterName] = { weightedScore: calculate(gradesBySemester[semesterName]), tooltipText: `当前学期加权百分制成绩\n(不含PNP和中文等级制成绩)` }; }
        scoresBySemester['全部'] = { weightedScore: calculate(allGrades), tooltipText: `所有学期加权百分制成绩\n(不含PNP和中文等级制成绩)` };
        return scoresBySemester;
    }

    function injectTooltipStylesForPortrait() {
        const styleId = 'gm-tooltip-styles-portrait'; if (document.getElementById(styleId)) return;
        const style = getOrCreateStyleEl(styleId);
        style.textContent = `
        .gm-tooltip-trigger { position: relative; cursor: help; font-family: "iconfont" !important; font-size: 14px; font-style: normal; }
        .gm-tooltip-trigger:hover::after { content: attr(data-gm-tooltip); position: absolute; bottom: 125%; left: 50%; transform: translateX(-50%); background-color: #303133; color: #fff; padding: 8px 12px; border-radius: 4px; font-size: 12px; line-height: 1.4; white-space: pre-line; z-index: 10001; display: inline-block; width: max-content; max-width: 280px; box-shadow: 0 2px 12px 0 rgba(0, 0, 0, 0.1); pointer-events: none; }
        .el-card__body .header { display: flex; justify-content: space-between; align-items: center; }
        .gm-trend-toggle-btn { padding: 5px 12px; font-size: 13px; border-radius: 4px; border: 1px solid #dcdfe6; background: #fff; color: #606266; cursor: pointer; transition: all 0.3s; outline: none; }
        .gm-trend-toggle-btn:hover { color: #409EFF; border-color: #c6e2ff; background-color: #ecf5ff; }
        .gm-trend-toggle-btn.active { color: #409EFF; border-color: #409EFF; background-color: #ecf5ff; font-weight: bold; }
        .gm-trend-detail-panel { overflow: hidden; transition: all 0.4s cubic-bezier(0.4, 0, 0.2, 1); max-height: 0; opacity: 0; margin-top: 0; padding-top: 0; border-top: 1px dashed transparent; }
        .gm-trend-detail-panel.show { max-height: 800px; opacity: 1; margin-top: 15px; padding-top: 15px; border-top-color: #ebeef5; }
        .gm-trend-grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(160px, 1fr)); gap: 12px; }
        .gm-trend-card { background: #fafafa; border: 1px solid #ebeef5; border-radius: 6px; padding: 10px 12px; transition: all 0.3s; position: relative; }
        .gm-trend-card:hover { box-shadow: 0 2px 12px 0 rgba(0,0,0,0.05); background: #fff; }
        .gm-tc-close { position: absolute; top: 4px; right: 6px; font-size: 16px; color: #c0c4cc; cursor: pointer; line-height: 1; transition: color 0.2s, transform 0.2s; user-select: none; }
        .gm-tc-close:hover { color: #f56c6c; transform: scale(1.2); }
        .gm-tc-title { font-size: 13px; font-weight: bold; color: #303133; margin-bottom: 8px; text-align: center; padding: 0 10px;}
        .gm-tc-row { display: flex; align-items: center; font-size: 12px; color: #666; margin-top: 4px; line-height: 1; }
        .gm-tc-dot { display: inline-block; width: 8px; height: 8px; border-radius: 50%; margin-right: 6px; flex-shrink: 0; }
        .gm-tc-val { margin-left: auto; font-weight: 900; font-family: Consolas, monospace; font-size: 13px; color: #333; }
    `;
        document.head.appendChild(style);
    }

    function updateSummaryTilesForPortrait(data, scoreContentElement, weightedScores) {
        if (!scoreContentElement || !ConfigManager.enablePortraitEnhancement) return;

        const infoDivs = Array.from(scoreContentElement.querySelectorAll('.info'));
        const avgScoreLabel = infoDivs.find(el => el.textContent.includes("平均分") || el.textContent.includes("加权分") || el.dataset.originalHtml);
        if (!avgScoreLabel) return;

        const semInput = document.querySelector('.myScore .el-select .el-input__inner');
        const currentSem = semInput ? (semInput.value || '全部') : '全部';
        const currentScoreData = weightedScores[currentSem] || {
            weightedScore: data ? 'N/A' : '计算中...',
            tooltipText: data ? '未找到成绩' : '正在拉取成绩与计算中...'
        };

        const weightedScoreTileId = 'gm-weighted-score-tile';
        let weightedScoreTile = document.getElementById(weightedScoreTileId);

        if (scoreContentElement.dataset.gmEnhancedSummary === 'true') {
            if (weightedScoreTile) {
                const scoreValDiv = weightedScoreTile.querySelector('.score');
                if (scoreValDiv && scoreValDiv.textContent !== currentScoreData.weightedScore) {
                    scoreValDiv.textContent = currentScoreData.weightedScore;
                }
                const tooltipTrigger = weightedScoreTile.querySelector('.gm-tooltip-trigger');
                if (tooltipTrigger) tooltipTrigger.setAttribute('data-gm-tooltip', currentScoreData.tooltipText);
            }
            return;
        }

        if (!weightedScoreTile) {
            weightedScoreTile = document.createElement('li');
            weightedScoreTile.id = weightedScoreTileId;
            weightedScoreTile.className = 'score-item';
            weightedScoreTile.style.background = '#409EFF';
            weightedScoreTile.innerHTML = `<div class="icon-img"><i class="icon-shuben"></i></div><div class="score-info"><div class="score">${currentScoreData.weightedScore}</div><div class="info">加权百分制分数 <i class="iconfont icon-bangzhu gm-tooltip-trigger" data-gm-tooltip="${currentScoreData.tooltipText}"></i></div>`;
            scoreContentElement.appendChild(weightedScoreTile);
        }

        scoreContentElement.dataset.gmEnhancedSummary = 'true';
    }

    function enhanceScoreTrendChart(data, weightedScores, forceRebuild = false) {
        if (!ConfigManager.enablePortraitEnhancement) return;

        const chartContainer = document.getElementById('semesterScoreLine');
        if (!chartContainer) return;

        const cardBody = chartContainer.closest('.el-card__body');
        const header = cardBody ? cardBody.querySelector('.header') : null;
        if (!header) return;

        const existingBtn = header.querySelector('.gm-trend-toggle-btn');
        const existingPanel = chartContainer.parentNode.querySelector('.gm-trend-detail-panel');

        // 正常观察模式下，如果有缓存且没要求强制重绘，直接跳过
        if (!forceRebuild && existingBtn && existingPanel && header.dataset.gmEnhancedTrend === 'true') {
            return;
        }

        // 强制重绘时，彻底清理旧组件防止双按钮
        if (existingBtn) existingBtn.remove();
        if (existingPanel) existingPanel.remove();

        const btn = document.createElement('button');
        btn.className = 'gm-trend-toggle-btn';
        btn.innerHTML = '展开数据';
        header.appendChild(btn);
        header.dataset.gmEnhancedTrend = 'true';

        const panel = document.createElement('div');
        panel.className = 'gm-trend-detail-panel';

        let html = '';
        // 如果还没获取到数据，渲染占位符
        if (!data) {
            html = '<div style="padding: 25px; text-align: center; color: #909399; font-size: 13px;">正在获取历史成绩详情，请稍候...</div>';
        } else {
            const semesterMap = {};
            data.allGrades.forEach(g => {
                const sem = g['学期'];
                if (!semesterMap[sem]) semesterMap[sem] = { credits: 0, points: 0 };
                const credit = parseFloat(g['学分']) || 0;
                const gp = parseFloat(g['绩点']);
                if (credit > 0 && !isNaN(gp)) {
                    semesterMap[sem].credits += credit;
                    semesterMap[sem].points += gp * credit;
                }
            });

            html = '<div class="gm-trend-grid">';
            const sortedSemesters = data.semesterNames ? [...data.semesterNames].reverse() : Object.keys(semesterMap).sort();

            sortedSemesters.forEach(sem => {
                const stats = semesterMap[sem];
                if (!stats) return;

                const gpa = stats.credits > 0 ? (stats.points / stats.credits).toFixed(2) : '-';
                const avgData = weightedScores[sem];
                const avg = avgData && avgData.weightedScore !== 'N/A' ? avgData.weightedScore : '-';

                html += `
                <div class="gm-trend-card">
                    <span class="gm-tc-close" title="关闭此项">&times;</span>
                    <div class="gm-tc-title">${sem}</div>
                    <div class="gm-tc-row">
                        <span class="gm-tc-dot" style="background:#F6DF49"></span>
                        <span>加权成绩</span>
                        <span class="gm-tc-val">${avg}</span>
                    </div>
                    <div class="gm-tc-row">
                        <span class="gm-tc-dot" style="background:#5B9BD5"></span>
                        <span>GPA</span>
                        <span class="gm-tc-val">${gpa}</span>
                    </div>
                </div>
            `;
            });
            html += '</div>';
        }

        panel.innerHTML = html;
        chartContainer.parentNode.insertBefore(panel, chartContainer.nextSibling);

        panel.addEventListener('click', (e) => {
            if (e.target.classList.contains('gm-tc-close')) {
                const card = e.target.closest('.gm-trend-card');
                if (card) {
                    card.style.transform = 'scale(0.8)'; card.style.opacity = '0';
                    setTimeout(() => { card.style.display = 'none'; }, 200);
                }
            }
        });

        btn.onclick = () => {
            const isHidden = !panel.classList.contains('show');
            if (isHidden) {
                panel.querySelectorAll('.gm-trend-card').forEach(c => {
                    c.style.display = ''; c.style.opacity = '1'; c.style.transform = 'scale(1)';
                });
                panel.classList.add('show');
                btn.classList.add('active');
                btn.innerHTML = '收起数据';
            } else {
                panel.classList.remove('show');
                btn.classList.remove('active');
                btn.innerHTML = '展开数据';
            }
        };
    }

    function getPassStatus(score) {
        const passingGrades = ['优秀', '良好', '中等', '及格', '通过', 'P'];
        const failingGrades = ['不及格', '不通过'];
        if (passingGrades.includes(score)) return '<span class="value">通过</span>';
        if (failingGrades.includes(score)) return '<span class="value" style="color: #F56C6C">不通过</span>';
        const numericScore = parseFloat(score);
        if (!isNaN(numericScore)) return numericScore >= 60 ? '<span class="value">通过</span>' : '<span class="value" style="color: #F56C6C">不通过</span>';
        return '';
    }

    function createEnhancedOutOfPlanTableForPortrait(data, originalTableContainer, forceRebuild = false) {
        if (!data) return; // 无数据时不处理

        const enhancedId = 'gm-enhanced-table-wrapper';
        let enhancedContainer = document.getElementById(enhancedId);

        if (!ConfigManager.enablePortraitEnhancement) {
            if (enhancedContainer) enhancedContainer.remove();
            originalTableContainer.style.display = '';
            originalTableContainer.removeAttribute('data-gm-enhanced');
            return;
        }

        // 正常状态跳过
        if (!forceRebuild && originalTableContainer.dataset.gmEnhanced === 'true' && enhancedContainer) {
            return;
        }

        // 热更新时，暴力清空旧表格
        if (enhancedContainer) enhancedContainer.remove();

        const outOfPlanCourseCodes = new Set();
        const rows = originalTableContainer.querySelectorAll('.el-table__body-wrapper tbody tr');
        const headerCells = Array.from(originalTableContainer.querySelectorAll('.el-table__header-wrapper th'));
        let codeIndex = headerCells.findIndex(th => th.textContent.trim().includes('课程代码'));
        if (codeIndex === -1) return;

        rows.forEach(row => {
            const cells = row.querySelectorAll('td');
            if (cells[codeIndex]) outOfPlanCourseCodes.add(cells[codeIndex].textContent.trim());
        });

        if (outOfPlanCourseCodes.size === 0) return;

        const outOfPlanGrades = data.allGrades.filter(grade => outOfPlanCourseCodes.has(grade['课程代码']));
        const classRankMap = new Map(data.allGrades.map(g => [g['课程代码'], g['教学班排名']]));

        const totalCredits = outOfPlanGrades.reduce((sum, g) => sum + parseFloat(g['学分'] || 0), 0);
        const passedCredits = outOfPlanGrades.reduce((sum, g) => {
            const statusHtml = getPassStatus(g['成绩']);
            return (statusHtml.includes('通过') && !statusHtml.includes('不')) ? sum + parseFloat(g['学分'] || 0) : sum;
        }, 0);
        const failedCredits = totalCredits - passedCredits;

        const originalHandler = originalTableContainer.querySelector('.node-handler');
        let paddingLeft = '20px';
        if (originalHandler && originalHandler.style.paddingLeft) paddingLeft = originalHandler.style.paddingLeft;

        enhancedContainer = document.createElement('div');
        enhancedContainer.id = enhancedId;
        enhancedContainer.className = 'node-wrapper courseTreeNode marginBottom';
        originalTableContainer.insertAdjacentElement('afterend', enhancedContainer);

        const colGroupHTML = `<colgroup><col width="48"><col width="200"><col width="100"><col width="120"><col width="80"><col width="60"><col width="60"><col width="60"><col width="100"><col width="80"></colgroup>`;
        const headerHTML = `<div class="el-table__header-wrapper"><table cellspacing="0" cellpadding="0" border="0" class="el-table__header" style="width: 100%;">${colGroupHTML}<thead class="has-gutter"><tr class="table-header"><th class="is-leaf" width="50"><div class="cell">序号</div></th><th class="is-leaf"><div class="cell">课程名称</div></th><th class="is-leaf" width="100"><div class="cell">课程代码</div></th><th class="is-leaf" width="120"><div class="cell">学年学期</div></th><th class="is-leaf" width="80"><div class="cell">是否必修</div></th><th class="is-leaf" width="60"><div class="cell">学分</div></th><th class="is-leaf" width="60"><div class="cell">成绩</div></th><th class="is-leaf" width="60"><div class="cell">绩点</div></th><th class="is-leaf" width="100"><div class="cell">教学班排名</div></th><th class="is-leaf" width="80"><div class="cell">是否通过</div></th></tr></thead></table></div>`;

        const tableBodyRows = outOfPlanGrades.map((grade, index) => {
            const score = grade['成绩'];
            const isFail = parseFloat(score) < 60 && !isNaN(parseFloat(score));
            const scoreStyle = isFail ? 'color: #F56C6C; font-weight: bold;' : '';
            const passStatus = getPassStatus(score);
            return `<tr class="el-table__row"><td class="cell-style"><div class="cell">${index + 1}</div></td><td class="cell-style"><div class="cell el-tooltip"><span class="value">${esc(grade['课程名称'] || '')}</span></div></td><td class="cell-style"><div class="cell el-tooltip">${esc(grade['课程代码'] || '')}</div></td><td class="cell-style"><div class="cell el-tooltip">${esc(grade['学期'] || '')}</div></td><td class="cell-style"><div class="cell el-tooltip"><span class="value">${grade['是否必修'] ? '是' : '否'}</span></div></td><td class="cell-style"><div class="cell el-tooltip">${esc(grade['学分'] || '')}</div></td><td class="cell-style"><div class="cell el-tooltip" style="${scoreStyle}">${esc(grade['成绩'] || '')}</div></td><td class="cell-style"><div class="cell el-tooltip">${esc(grade['绩点'] ?? '')}</div></td><td class="cell-style"><div class="cell el-tooltip"><span class="value">${esc(classRankMap.get(grade['课程代码']) || '-')}</span></div></td><td class="cell-style"><div class="cell el-tooltip">${passStatus}</div></td></tr>`;
        }).join('');

        const bodyHTML = `<div class="el-table__body-wrapper is-scrolling-left"><table cellspacing="0" cellpadding="0" border="0" class="el-table__body" style="width: 100%;">${colGroupHTML}<tbody>${tableBodyRows}</tbody></table></div>`;

        enhancedContainer.innerHTML = `<div class="node-handler background" style="padding-left: ${paddingLeft}; cursor: pointer;"><div class="arrow"></div><div class="title"><div class="course-name">计划外课程</div><div class="require-item"><span class="score">学分：</span><span class="con">共 ${totalCredits} | 已通过 ${passedCredits} | 未通过 </span><span class="unpassed">${failedCredits}</span></div></div></div><div class="node-child-wrapper none"><div class="node-child"><div class="child"><div class="el-table el-table--fit el-table--enable-row-hover el-table--enable-row-transition el-table--small" style="width: 100%;">${headerHTML}${bodyHTML}</div></div></div></div>`;

        const handler = enhancedContainer.querySelector('.node-handler');
        const wrapper = enhancedContainer.querySelector('.node-child-wrapper');
        const arrow = enhancedContainer.querySelector('.arrow');
        handler.addEventListener('click', () => {
            if (wrapper.classList.contains('none')) { wrapper.classList.remove('none'); arrow.classList.add('up'); }
            else { wrapper.classList.add('none'); arrow.classList.remove('up'); }
        });

        originalTableContainer.style.display = 'none';
        originalTableContainer.dataset.gmEnhanced = 'true';
    }

    /** 记录见过的 StudentID（去重），供后台预取 */
    function recordPortraitStudentId(sid) {
        if (!sid) return;
        try {
            const list = JSON.parse(GM_getValue(CONSTANTS.PORTRAIT_STUDENT_IDS_KEY, '[]'));
            if (!Array.isArray(list)) return;
            const str = String(sid);
            if (list.indexOf(str) === -1) {
                list.push(str);
                GM_setValue(CONSTANTS.PORTRAIT_STUDENT_IDS_KEY, JSON.stringify(list));
            }
        } catch (e) { }
    }

    /** 用记录的 StudentID 提前拉取画像数据缓存；无权限的 ID 自动失败忽略 */
    async function prefetchPortraitData() {
        try {
            const list = JSON.parse(GM_getValue(CONSTANTS.PORTRAIT_STUDENT_IDS_KEY, '[]'));
            if (!Array.isArray(list) || list.length === 0) return;
            for (const id of list) {
                if (!id) continue;
                try {
                    const data = await fetchAllDataAndCache(0, { force: true, studentId: id });
                    if (data && data.allGrades && data.allGrades.length > 0) {
                        GM_setValue(CONSTANTS.PORTRAIT_PREFETCH_KEY, JSON.stringify({ studentId: String(id), data: data, time: Date.now() }));
                        Logger.log("2.4", `预取学生画像数据成功: studentId=${id}`);
                        break;
                    }
                } catch (e) {
                    Logger.log("2.4", `预取 studentId=${id} 失败(可能无权限)，忽略`);
                }
            }
        } catch (e) { }
    }

    /** 获取画像页当前身份 tab 对应的 StudentID（tab-id 数字即 cultivateTypeAssoc） */
    function getPortraitStudentId() {
        return new Promise((resolve) => {
            // tab id 形如 tab-1(主修)/tab-6(研究生)，数字即 cultivateTypeAssoc
            let cultivateTypeAssoc = '1';
            try {
                const activeTab = document.querySelector('.el-tabs__item.is-active');
                if (activeTab) {
                    const m = activeTab.id.match(/tab-(\d+)/);
                    if (m) cultivateTypeAssoc = m[1];
                }
            } catch (e) { }

            GM_xmlhttpRequest({
                method: 'GET',
                url: `https://jwxt.nwpu.edu.cn/student/for-std/student-portrait/getStdInfo?bizTypeAssoc=2&cultivateTypeAssoc=${cultivateTypeAssoc}`,
                onload: (res) => {
                    if (res.status === 200) {
                        try {
                            const data = JSON.parse(res.responseText);
                            const sid = data.student && data.student.id;
                            if (sid) {
                                Logger.log('Core', `画像页获取正确 StudentID: ${sid} (cultivateTypeAssoc=${cultivateTypeAssoc})`);
                                cachedStudentId = sid;
                                localStorage.setItem(CONSTANTS.STUDENT_ID_LS_KEY, sid);
                                recordPortraitStudentId(sid);
                                resolve(sid);
                            } else {
                                Logger.error('Core', "getStdInfo 响应中未找到 student.id");
                                resolve(null);
                            }
                        } catch (e) {
                            Logger.error('Core', "getStdInfo JSON 解析失败", e);
                            resolve(null);
                        }
                    } else {
                        Logger.error('Core', `getStdInfo 请求失败，HTTP状态码: ${res.status}`);
                        resolve(null);
                    }
                },
                onerror: () => resolve(null)
            });
        });
    }

    async function enhancePortraitPage() {
        // 最多等待 30 秒，避免页面结构变更后永久轮询
        for (let waitedMs = 0; !document.body || !document.querySelector(".score-content"); waitedMs += 50) {
            if (waitedMs >= 30000) { Logger.warn("2.4", "等待 .score-content 超时，画像增强未启动"); return; }
            await new Promise(resolve => setTimeout(resolve, 50));
        }
        Logger.log("2.4", "脚本已在学生画像页激活");
        injectTooltipStylesForPortrait();

        // 优先读预取缓存，否则回退 sessionStorage
        let currentData = null;
        let prefetchedStudentId = null;
        try {
            const pf = JSON.parse(GM_getValue(CONSTANTS.PORTRAIT_PREFETCH_KEY, 'null'));
            if (pf && pf.data && pf.data.allGrades && pf.data.allGrades.length > 0 && pf.studentId) {
                currentData = pf.data;
                prefetchedStudentId = String(pf.studentId);
            }
        } catch (e) { }
        if (!currentData) currentData = getCachedData();
        let currentWeightedScores = currentData ? precomputeAllWeightedScores(currentData.allGrades) : {};

        // 渲染方法 (常规调配)
        const renderAllComponents = () => {
            if (!ConfigManager.enablePortraitEnhancement) return;
            const scoreContent = document.querySelector(".score-content");
            if (scoreContent) updateSummaryTilesForPortrait(currentData, scoreContent, currentWeightedScores);

            const outOfPlanTable = document.querySelector('.outPlanTable');
            if (outOfPlanTable && outOfPlanTable.querySelector('.el-table__body-wrapper tbody tr')) {
                createEnhancedOutOfPlanTableForPortrait(currentData, outOfPlanTable);
            }

            enhanceScoreTrendChart(currentData, currentWeightedScores);
        };

        // 热更新方法 (数据到达后，强制推翻重建)
        const triggerHotUpdate = () => {
            Logger.log("2.4", "执行热更新...");
            const scoreContent = document.querySelector(".score-content");
            if (scoreContent) updateSummaryTilesForPortrait(currentData, scoreContent, currentWeightedScores);

            const outOfPlanTable = document.querySelector('.outPlanTable');
            if (outOfPlanTable && outOfPlanTable.querySelector('.el-table__body-wrapper tbody tr')) {
                createEnhancedOutOfPlanTableForPortrait(currentData, outOfPlanTable, true); // 传递 true 强制更新
            }

            enhanceScoreTrendChart(currentData, currentWeightedScores, true); // 传递 true 强制更新
        };

        // ============ 执行流程 ============

        // 1. 立即挂载（无论有无缓存均立即展示界面骨架）
        renderAllComponents();

        // 2. 先获取当前身份 tab 对应的正确 StudentID（本研/研究生双身份），再拉取成绩数据
        const loadPortraitData = () => {
            return getPortraitStudentId().then(correctId => {
                // 预取 ID 与当前身份匹配 → 即刻用预取数据更新，随后后台刷新
                if (correctId && prefetchedStudentId && String(correctId) === prefetchedStudentId && currentData && currentData.allGrades) {
                    currentWeightedScores = precomputeAllWeightedScores(currentData.allGrades);
                    setTimeout(triggerHotUpdate, 100);
                }
                return fetchAllDataAndCache(0, { force: true, studentId: correctId });
            });
        };
        const initTab = document.querySelector('.el-tabs__item.is-active');
        window._gm_last_portrait_tab = initTab ? initTab.id : null; // 记录初始 tab，避免首次 observer 误触发
        loadPortraitData().then(freshData => {
            currentData = freshData;
            currentWeightedScores = precomputeAllWeightedScores(freshData.allGrades);
            setTimeout(triggerHotUpdate, 150);
        }).catch(err => {
            Logger.error("2.4", "静默获取成绩失败:", err);
        });

        window._gm_last_portrait_sem = null;
        const observer = new MutationObserver(() => {
            if (!ConfigManager.enablePortraitEnhancement) return;

            // 身份 tab 切换 → 重新拉取对应身份数据
            const activeTab = document.querySelector('.el-tabs__item.is-active');
            const tabId = activeTab ? activeTab.id : null;
            if (tabId && tabId !== window._gm_last_portrait_tab) {
                window._gm_last_portrait_tab = tabId;
                Logger.log("2.4", "检测到身份切换: " + ((activeTab.textContent || '').trim()));
                loadPortraitData().then(freshData => {
                    currentData = freshData;
                    currentWeightedScores = precomputeAllWeightedScores(freshData.allGrades);
                    setTimeout(triggerHotUpdate, 150);
                }).catch(err => {
                    Logger.error("2.4", "身份切换后数据拉取失败:", err);
                });
            }

            // 学期下拉框变更 → 更新卡片
            const semInput = document.querySelector('.myScore .el-select .el-input__inner');
            if (semInput) {
                const currentSem = semInput.value || '全部';
                if (currentSem !== window._gm_last_portrait_sem) {
                    window._gm_last_portrait_sem = currentSem;
                    updateSummaryTilesForPortrait(currentData, document.querySelector(".score-content"), currentWeightedScores);
                }
            }

            // 常规容错：如果 Vue 刷新掉了组件，自动补回
            renderAllComponents();
        });

        observer.observe(document.body, { childList: true, subtree: true });
    }

    // =-=-=-=-=-=-=-=-=-=-=-=-= 2.5 全校开课查询页选课记录 =-=-=-=-=-=-=-=-=-=-=-=-=
    const LessonSearchEnhancer = {
        // 配置常量
        CONFIG: {
            HISTORY_KEY: CONSTANTS.HISTORY_STORAGE_KEY,
            // 从 URL 中提取学生ID，如 /lesson-search/index/277933
            STUDENT_ID: window.location.pathname.match(/lesson-search\/index\/(\d+)/)?.[1] || '',
            API_BASE: '/student/for-std/lesson-search/semester',
            API_FIELDS: 'course.code,minorCourse.nameZh,courseType,openDepartment,teacherAssignmentList,examMode,campus,teachLang,roomType,timeTableLayout,crossBizTypes,courseProperty',
            PAGE_SIZE: 1000, // 使用 queryPage__ 参数格式，服务端支持最大1000条/页
            MAX_VISIBLE_SEMESTERS: 5, // 学期下拉列表每屏最多显示条数（超出可滚动查看）
            SEMESTER_ITEM_HEIGHT: 32, // 下拉列表单项高度(px)，用于计算列表最大高度
            // 保留 DOM 选择器用于渲染历史标签
            TABLE_ROWS: '#table tbody tr'
        },

        init() {
            // 幂等守卫：防止路由分发与历史遗留入口造成重复初始化（重复注入控制面板）
            if (this._initialized) return;
            if (!window.location.href.includes('/student/for-std/lesson-search')) return;
            if (!document.querySelector('.page-config') || !document.querySelector('#table')) {
                // 最多重试 50 次（约 15 秒），避免目标容器缺失时无限轮询
                this._initRetries = (this._initRetries || 0) + 1;
                if (this._initRetries > 50) { Logger.warn("2.5", "等待页面容器超时，选课助手面板未注入"); return; }
                setTimeout(() => this.init(), 300);
                return;
            }
            this._initialized = true;

            Logger.log("2.5", "初始化选课记录模块...");
            this.injectControlPanel();
            this.renderHistoryTags();

            // 通过API获取学期列表并渲染到自定义下拉框（全部学期，滚动查看）
            if (this.CONFIG.STUDENT_ID) {
                this.fetchSemesterList().then(semesterOptions => {
                    if (semesterOptions.length > 0) {
                        semesterOptions.sort((a, b) => parseInt(b.value) - parseInt(a.value));
                        this.renderSemesterDropdown(semesterOptions);
                        Logger.log("2.5", `已加载 ${semesterOptions.length} 个学期`);
                        this.syncSemesterSelect();
                    } else {
                        this.setSemesterPlaceholder('未找到学期数据');
                    }
                }).catch(e => {
                    Logger.error("2.5", "API获取学期列表失败: " + e.message);
                    this.setSemesterPlaceholder('加载失败，请刷新页面重试');
                });
            }

            // 启动观察者，页面数据更新时同步历史标签和下拉框学期
            const observer = new MutationObserver(() => {
                this.renderHistoryTags();
                this.syncSemesterSelect();
            });
            const target = document.querySelector('#table') || document.body;
            observer.observe(target, { childList: true, subtree: true });
        },

        fetchSemesterList() {
            return new Promise((resolve, reject) => {
                const url = '/student/for-std/lesson-search/index/' + this.CONFIG.STUDENT_ID;
                const xhr = new XMLHttpRequest();
                xhr.open('GET', url);
                xhr.onload = function () {
                    if (xhr.status === 200) {
                        try {
                            const html = xhr.responseText;
                            const parser = new DOMParser();
                            const doc = parser.parseFromString(html, 'text/html');
                            const options = doc.querySelectorAll('#semester option[value]');
                            const semesterList = [];
                            options.forEach(opt => {
                                const val = opt.getAttribute('value');
                                const text = opt.textContent.trim();
                                if (val && text && val !== '') {
                                    semesterList.push({ value: val, text: text });
                                }
                            });
                            resolve(semesterList);
                        } catch (e) {
                            reject(new Error('HTML解析失败: ' + e.message));
                        }
                    } else {
                        reject(new Error('请求失败: HTTP ' + xhr.status));
                    }
                };
                xhr.onerror = function () {
                    reject(new Error('网络请求失败'));
                };
                xhr.send();
            });
        },

        syncSemesterSelect() {
            const originalSemester = document.querySelector('.selectize-control.semester .item');
            const list = document.getElementById('gm-semester-dd-list');
            // 用户正在展开下拉列表手动选择时不覆盖
            if (originalSemester && this.semesterOptions && this.semesterOptions.length > 0 && (!list || list.style.display === 'none')) {
                const semName = originalSemester.innerText.trim();
                if (this.selectedSemester && this.selectedSemester.text === semName) return;
                const match = this.semesterOptions.find(o => o.text === semName);
                if (match) this.selectSemester(match);
            }
        },

        // --- 自定义学期下拉框（全部学期，列表每屏最多显示 MAX_VISIBLE_SEMESTERS 项，可滚动） ---
        renderSemesterDropdown(semesterOptions) {
            this.semesterOptions = semesterOptions;
            const list = document.getElementById('gm-semester-dd-list');
            if (!list) return;
            list.innerHTML = '';
            semesterOptions.forEach(opt => {
                const item = document.createElement('div');
                item.className = 'gm-semester-dd__item';
                item.dataset.value = opt.value;
                item.textContent = opt.text;
                item.title = opt.text;
                item.onclick = (e) => {
                    e.stopPropagation();
                    this.selectSemester(opt);
                    this.toggleSemesterList(false);
                };
                list.appendChild(item);
            });
            if (semesterOptions.length > 0) this.selectSemester(semesterOptions[0]);
        },

        selectSemester(opt) {
            this.selectedSemester = opt;
            const text = document.getElementById('gm-semester-dd-text');
            if (text) {
                text.textContent = opt.text;
                text.style.color = '#303133';
            }
            const list = document.getElementById('gm-semester-dd-list');
            if (list) {
                Array.from(list.children).forEach(el => {
                    const isSel = el.dataset.value === opt.value;
                    el.dataset.selected = isSel ? '1' : '0';
                    el.classList.toggle('is-selected', isSel);
                });
            }
        },

        toggleSemesterList(force) {
            const list = document.getElementById('gm-semester-dd-list');
            if (!list) return;
            const show = (force !== undefined) ? force : (list.style.display === 'none');
            list.style.display = show ? 'block' : 'none';
        },

        setSemesterPlaceholder(msg) {
            const text = document.getElementById('gm-semester-dd-text');
            if (text) {
                text.textContent = msg;
                text.style.color = '#c0c4cc';
            }
        },

        injectControlPanel() {
            if (document.getElementById('gm-lesson-helper-panel')) return;

            const panel = document.createElement('div');
            panel.id = 'gm-lesson-helper-panel';
            panel.className = 'gm-lesson-panel';
            panel.innerHTML = `
            <div class="gm-lesson-panel__header" id="gm-panel-header">
                选课助手
                <button type="button" id="gm-panel-close" class="gm-modal-close" title="关闭面板">&times;</button>
            </div>
            <div class="gm-lesson-panel__body">
                <div id="gm-semester-dd" class="gm-semester-dd">
                    <div id="gm-semester-dd-trigger" class="gm-semester-dd__trigger">
                        <span id="gm-semester-dd-text">加载学期列表中...</span>
                        <svg class="gm-semester-dd__chevron" viewBox="0 0 24 24" width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polyline points="6 9 12 15 18 9"></polyline></svg>
                    </div>
                    <div id="gm-semester-dd-list" class="gm-semester-dd__list" style="display:none; max-height:${this.CONFIG.MAX_VISIBLE_SEMESTERS * this.CONFIG.SEMESTER_ITEM_HEIGHT}px;"></div>
                </div>
                <button type="button" id="gm-btn-sync-start" class="gm-lesson-panel__sync">存储选定学期课程信息</button>
            </div>
        `;
            panel.style.cssText = 'position:fixed; top:120px; right:30px; z-index:99999;';
            document.body.appendChild(panel);

            const btnSync = document.getElementById('gm-btn-sync-start');
            const btnClose = document.getElementById('gm-panel-close');
            const semesterDd = document.getElementById('gm-semester-dd');
            const semesterTrigger = document.getElementById('gm-semester-dd-trigger');

            // 关闭后保留一个「选课助手」小按钮，点击即可重新打开面板（无需刷新页面）
            btnClose.onclick = () => {
                panel.style.display = 'none';
                if (document.getElementById('gm-lesson-helper-reopen')) return;
                const reopenBtn = document.createElement('button');
                reopenBtn.id = 'gm-lesson-helper-reopen';
                reopenBtn.type = 'button';
                reopenBtn.textContent = '选课助手';
                reopenBtn.title = '重新打开选课助手面板';
                reopenBtn.style.cssText = 'position:fixed; top:120px; right:30px; z-index:99999; padding:8px 14px; border:1px solid var(--gm-border); border-radius:8px; background:#fff; color:var(--gm-primary-text); font-size:13px; font-weight:600; cursor:pointer; box-shadow:0 4px 12px rgba(0,0,0,0.12);';
                reopenBtn.onclick = () => { reopenBtn.remove(); panel.style.display = 'block'; };
                document.body.appendChild(reopenBtn);
            };
            btnSync.onclick = () => this.startSyncProcess();

            semesterTrigger.onclick = (e) => { e.stopPropagation(); this.toggleSemesterList(); };
            document.addEventListener('click', (e) => {
                if (semesterDd && !semesterDd.contains(e.target)) this.toggleSemesterList(false);
            });

            const header = document.getElementById('gm-panel-header');
            let isDragging = false, startX, startY, initialLeft, initialTop;
            // 拖拽期间才挂载 move/up 监听，避免覆盖页面自身的 document 事件处理器
            const onDragMove = (e) => {
                if (!isDragging) return;
                e.preventDefault();
                panel.style.left = (initialLeft + e.clientX - startX) + 'px';
                panel.style.top = (initialTop + e.clientY - startY) + 'px';
                panel.style.right = 'auto';
            };
            const onDragEnd = () => {
                if (!isDragging) return;
                isDragging = false;
                document.removeEventListener('mousemove', onDragMove);
                document.removeEventListener('mouseup', onDragEnd);
            };
            header.addEventListener('mousedown', (e) => {
                if (e.target === btnClose) return;
                isDragging = true; startX = e.clientX; startY = e.clientY;
                const rect = panel.getBoundingClientRect(); initialLeft = rect.left; initialTop = rect.top;
                document.addEventListener('mousemove', onDragMove);
                document.addEventListener('mouseup', onDragEnd);
            });
        },

        // --- 2. Core: 同步逻辑（API 直接获取） ---
        async startSyncProcess() {
            if (!this.selectedSemester || !this.selectedSemester.value) {
                alert("学期列表仍在加载中或加载失败，请稍等后再试。");
                return;
            }

            if (!this.CONFIG.STUDENT_ID) {
                alert("无法从URL中提取学生ID，请确认页面地址格式正确。");
                return;
            }

            const targetSemesterId = this.selectedSemester.value;
            const targetSemesterName = this.selectedSemester.text;

            if (!confirm(`即将抓取【${targetSemesterName}】的数据。\n过程可能需要几秒到十几秒，请勿关闭页面。`)) return;

            let overlay = null;

            try {
                overlay = this.showOverlay(targetSemesterName);

                // --- 通过 API 分页获取全部数据（并行请求） ---
                this.updateOverlayStatus("正在获取课程数据...");
                const pageSize = this.CONFIG.PAGE_SIZE;

                // 第一次请求：获取 totalPages 信息
                this.updateOverlayStatus("正在获取第1页...");
                const firstResult = await this.fetchCoursePage(targetSemesterId, 1, pageSize);
                if (!firstResult || !firstResult.data || firstResult.data.length === 0) {
                    alert(`【${targetSemesterName}】学期暂无课程数据。`);
                    overlay.remove();
                    return;
                }

                const totalPages = firstResult._page_?.totalPages || 1;
                const allCourseData = [...this.parseCourseData(firstResult.data, targetSemesterName)];
                this.updateOverlay(allCourseData.length);

                // 如果有更多页，并行请求剩余所有页
                if (totalPages > 1) {
                    this.updateOverlayStatus(`正在获取剩余 ${totalPages - 1} 页...`);
                    const pagePromises = [];
                    for (let page = 2; page <= totalPages; page++) {
                        pagePromises.push(this.fetchCoursePage(targetSemesterId, page, pageSize));
                    }

                    const results = await Promise.all(pagePromises);
                    for (const result of results) {
                        if (result && result.data && result.data.length > 0) {
                            const pageData = this.parseCourseData(result.data, targetSemesterName);
                            allCourseData.push(...pageData);
                        }
                    }
                    this.updateOverlay(allCourseData.length);
                }

                if (allCourseData.length > 0) {
                    this.saveToHistory(allCourseData);
                    overlay.remove();
                    // 游戏式存档：在线同步的新数据静默写入「使用中」的存档（保留自定义命名），随后刷新页面
                    ArchiveManager.updateActiveSlot(targetSemesterName);
                    window.location.reload();
                } else {
                    alert(`【${targetSemesterName}】学期暂无课程数据。`);
                    overlay.remove();
                }

            } catch (e) {
                Logger.error("2.5", "同步中断", e);
                alert('同步中断: ' + e.message);
                if (overlay) overlay.remove();
            }
        },

        // --- API 调用：获取指定学期指定页的课程数据 ---
        fetchCoursePage(semesterId, page, pageSize) {
            return new Promise((resolve, reject) => {
                const timestamp = Date.now();
                const url = `${this.CONFIG.API_BASE}/${semesterId}/search/${this.CONFIG.STUDENT_ID}` +
                    `?bizTypeAssoc=1&queryPage__=${page},${pageSize}&assembleFields=${encodeURIComponent(this.CONFIG.API_FIELDS)}` +
                    `&_=${timestamp}`;

                const xhr = new XMLHttpRequest();
                xhr.open('GET', url);
                xhr.setRequestHeader('X-Requested-With', 'XMLHttpRequest');
                xhr.onload = function () {
                    if (xhr.status === 200) {
                        try {
                            resolve(JSON.parse(xhr.responseText));
                        } catch (e) {
                            reject(new Error('API响应解析失败'));
                        }
                    } else {
                        reject(new Error(`API请求失败: HTTP ${xhr.status}`));
                    }
                };
                xhr.onerror = function () {
                    reject(new Error('网络请求失败'));
                };
                xhr.send();
            });
        },

        // --- 解析 API 返回的课程数据 ---
        parseCourseData(courses, semesterName) {
            const timestamp = new Date().toLocaleString('zh-CN', { hour12: false });
            // 课程名在 course.course.nameZh（顶层 nameZh 是班级范围名如"全校"）
            return courses.map(course => ({
                id: String(course.id),
                code: course.code || 'N/A',
                name: course.course?.nameZh || course.nameZh || 'N/A',
                stdCount: course.stdCount || 0,
                limitCount: course.limitCount || 0,
                term: semesterName,
                time: timestamp
            }));
        },

        saveToHistory(courseData) {
            let history = {};
            try {
                history = JSON.parse(GM_getValue(this.CONFIG.HISTORY_KEY, '{}'));
            } catch (e) {
                history = {};
            }

            courseData.forEach(c => {
                if (!history[c.id]) history[c.id] = [];
                const records = history[c.id];
                const last = records[records.length - 1];
                if (!last || last.stdCount !== c.stdCount || last.limitCount !== c.limitCount) {
                    records.push(c);
                } else {
                    last.time = c.time;
                }
            });

            GM_setValue(this.CONFIG.HISTORY_KEY, JSON.stringify(history));
            this.renderHistoryTags();
        },

        // --- 4. UI: 渲染历史标签 ---
        renderHistoryTags() {
            let history = {};
            try {
                history = JSON.parse(GM_getValue(this.CONFIG.HISTORY_KEY, '{}'));
            } catch (e) {
                GM_setValue(this.CONFIG.HISTORY_KEY, '{}');
                return;
            }

            const rows = document.querySelectorAll(this.CONFIG.TABLE_ROWS);

            rows.forEach(row => {
                const idInput = row.querySelector('input[name="model_id"]');
                if (!idInput) return;

                // DOM 复用检测：比对当前 lesson ID 与上次处理的 ID，若一致则跳过
                if (row.dataset.gmProcessed === idInput.value) return;

                // 清理旧的历史标签（DOM 被复用时）
                if (row.dataset.gmProcessed) {
                    row.querySelectorAll('.gm-hist-tag').forEach(el => el.remove());
                }

                const records = history[idInput.value];
                if (records && records.length > 0) {
                    const last = records[records.length - 1];
                    const countSpan = Array.from(row.querySelectorAll('span')).find(s => s.getAttribute('data-original-title') === '实际/上限人数');

                    if (countSpan && !countSpan.parentNode.querySelector('.gm-hist-tag')) {
                        const tag = document.createElement('span');
                        const isFull = last.stdCount >= last.limitCount;
                        tag.className = `gm-hist-tag ${isFull ? 'gm-hist-tag--full' : 'gm-hist-tag--ok'}`;
                        tag.innerText = `记录:${last.stdCount}/${last.limitCount}`;
                        tag.title = `上次同步时间: ${last.time}`;
                        countSpan.parentNode.appendChild(tag);
                    }
                }
                row.dataset.gmProcessed = idInput.value;
            });
        },

        showOverlay(semesterName = '当前学期') {
            const div = document.createElement('div');
            div.id = 'gm-sync-overlay';
            div.style.cssText = 'position:fixed; top:0; left:0; width:100%; height:100%; background:rgba(0,0,0,0.7); z-index:100002; color:white; display:flex; flex-direction:column; align-items:center; justify-content:center;';
            div.innerHTML = `
            <div style="font-size:24px; font-weight:bold; margin-bottom:15px;">正在同步课程数据...</div>
            <div id="gm-overlay-target-sem" style="font-size:18px; margin-bottom:10px; color:#ffeb3b;">锁定抓取学期: ${semesterName}</div>
            <div id="gm-overlay-status" style="font-size:16px; margin-bottom:10px; color:#ddd;">正在初始化...</div>
            <div style="font-size:18px;">已抓取: <span id="gm-sync-count" style="color:#4facfe; font-weight:bold;">0</span> 条</div>
            <div style="margin-top:30px; color:#aaa; font-size:14px;">请勿关闭页面，程序正在自动操作</div>
        `;
            document.body.appendChild(div);
            return div;
        },

        updateOverlay(count) {
            const el = document.getElementById('gm-sync-count');
            if (el) el.innerText = count;
        },

        updateOverlayStatus(text) {
            const el = document.getElementById('gm-overlay-status');
            if (el) el.innerText = text;
        }
    };

    // 注：开课查询页的初始化入口已统一收归 runMainFeatures 路由分发（见文件末尾），此处不再单独绑定。
    // =-=-=-=-=-=-=-=-=-=-=-=-= 2.6 课程关注 =-=-=-=-=-=-=-=-=-=-=-=-=
    /** 教师名可点击跳转教师主页【旧版：开课查询，直接绑定 onclick】 */
    function enhanceTeacherNames(row) {
        if (!document.getElementById('gm-teacher-link-style')) {
            const style = getOrCreateStyleEl('gm-teacher-link-style');
            style.textContent = `
            .gm-teacher-link { color: #409EFF !important; text-decoration: none; font-weight: bold; transition: all 0.2s; cursor: pointer; }
            .gm-teacher-link:hover { color: #0056b3 !important; text-decoration: underline; background-color: rgba(64,158,255,0.1); border-radius: 4px; padding: 2px 4px; }
        `;
            document.head.appendChild(style);
        }

        const teacherEl = row.querySelector('.course-teacher');
        if (!teacherEl || teacherEl.dataset.gmLinked === "true") return;

        const rawText = teacherEl.innerText.trim();
        if (!rawText || rawText === '待定' || rawText === '-') return;

        teacherEl.dataset.gmLinked = "true";
        teacherEl.innerHTML = '';
        // 按空白/逗号/分号拆分多教师（\s 兼容换行）
        const names = rawText.split(/[\s,，;；]+/);
        names.forEach((name, idx) => {
            if (!name) return;
            const a = document.createElement('a');
            a.className = 'gm-teacher-link';
            a.textContent = name;
            a.title = `点击搜索 ${name} 的教师主页`;
            a.href = "javascript:void(0);";
            a.onclick = (e) => {
                e.preventDefault();
                e.stopPropagation();
                GM_setValue(CONSTANTS.CROSS_SEARCH_NAME_KEY, name);
                window.open('https://teacher.nwpu.edu.cn/search/syss/.html', '_blank');
            };
            teacherEl.appendChild(a);
            if (idx < names.length - 1) teacherEl.appendChild(document.createTextNode(' '));
        });
    }

    /**
     * 在开课查询页面的表格中注入关注按钮
     */
    function injectFollowButtons() {
        if (!ConfigManager.enableCourseWatch) return;

        // --- 1. 获取学期信息 ---
        let currentSemester = "未知学期";
        try {
            const semesterEl = document.querySelector('.selectize-control.semester .item') ||
                document.querySelector('.semester-name') ||
                document.querySelector('.selectize-input .item');
            if (semesterEl) {
                currentSemester = semesterEl.innerText.trim();
            }
        } catch (e) { Logger.warn("2.6", "无法自动获取学期名称", e); }

        // --- 3. 获取表格容器 ---
        const scrollBodyTable = document.querySelector('.dataTables_scrollBody table#table');
        if (!scrollBodyTable) return;

        const rows = scrollBodyTable.querySelectorAll('tbody tr');
        rows.forEach(row => {
            if (row.querySelector('.dataTables_empty')) return;

            enhanceTeacherNames(row);

            const firstTd = row.querySelector('td:first-child');
            const checkbox = row.querySelector('input[name="model_id"]');
            if (!firstTd || !checkbox) return;
            if (firstTd.querySelector('.gm-follow-btn')) return;

            // --- 数据提取 ---
            const storageId = checkbox.value;
            const lessonCodeDiv = row.querySelector('.lesson-code');
            const displayCode = lessonCodeDiv ? lessonCodeDiv.innerText.trim() : '未知编号';

            const nameEl = row.querySelector('.course-name a');
            const name = nameEl ? nameEl.innerText.trim() : '未知课程';
            const teacherEl = row.querySelector('.course-teacher');
            const teachers = teacherEl ? teacherEl.innerText.trim() : '';
            const creditEl = row.children[3];
            const credits = creditEl ? creditEl.innerText.trim() : '';
            const placeEl = row.querySelector('.course-datetime-place');
            const timeAndPlace = placeEl ? placeEl.innerText.trim() : '';

            // --- 样式布局 ---
            firstTd.style.display = 'flex';
            firstTd.style.flexDirection = 'column';
            firstTd.style.alignItems = 'center';
            firstTd.style.justifyContent = 'center';
            firstTd.style.padding = '8px 0';
            firstTd.style.height = '100%';
            checkbox.style.margin = '0';

            // --- 创建按钮 ---
            const btn = document.createElement('div');
            btn.className = 'gm-follow-btn';
            btn.innerHTML = '<svg viewBox="0 0 24 24" width="20" height="20" fill="currentColor" stroke="none"><path d="M20.84 4.61a5.5 5.5 0 0 0-7.78 0L12 5.67l-1.06-1.06a5.5 5.5 0 0 0-7.78 7.78l1.06 1.06L12 21.23l7.78-7.78 1.06-1.06a5.5 5.5 0 0 0 0-7.78z"></path></svg>';
            btn.style.cssText = `cursor: pointer; font-size: 20px; margin-top: 4px; line-height: 1; user-select: none; transition: all 0.2s; font-family: sans-serif;`;

            const updateState = () => {
                if (FollowManager.has(storageId)) {
                    btn.title = '点击取消关注';
                    btn.style.color = '#f56c6c';
                    btn.style.textShadow = 'none';
                    btn.style.opacity = '1';
                    btn.style.transform = 'scale(1.08)';
                } else {
                    btn.title = '点击关注课程';
                    btn.style.color = '#dcdfe6';
                    btn.style.textShadow = 'none';
                    btn.style.opacity = '1';
                    btn.style.transform = 'scale(1)';
                }
            };
            updateState();

            btn.onmouseenter = () => { if (!FollowManager.has(storageId)) { btn.style.color = '#fbc4c4'; btn.style.transform = 'scale(1.1)'; } };
            btn.onmouseleave = () => updateState();

            btn.onclick = (e) => {
                e.stopPropagation(); e.preventDefault();
                btn.style.transform = 'scale(0.8)';
                setTimeout(() => updateState(), 150);

                if (FollowManager.has(storageId)) {
                    FollowManager.remove(storageId);
                    showToast(`已取消关注 ${displayCode}`, 'cancel');
                } else {
                    FollowManager.add(storageId, {
                        id: storageId,
                        code: displayCode,
                        name, teachers, credits, timeAndPlace,
                        semester: currentSemester, // 保存当前学期
                        addedTime: new Date().toLocaleString()
                    });
                    showToast(`已加入关注列表 ${displayCode}`, 'success');
                }
            };
            firstTd.appendChild(btn);
        });
    }

    /**
     * 启动开课查询页面的监听器
     */
    function initLessonSearchPage() {
        if (!ConfigManager.enableCourseWatch) return;
        Logger.log("2.6", "已进入全校开课查询页面 (Iframe)");

        // 初始执行一次
        injectFollowButtons();

        // 使用 MutationObserver 监听表格变化（翻页、搜索时触发，100ms 防抖合并连续变更）
        let debounceTimer = null;
        const observer = new MutationObserver(() => {
            clearTimeout(debounceTimer);
            debounceTimer = setTimeout(injectFollowButtons, 100);
        });

        const tableContainer = document.getElementById('e-content-area') || document.body;
        observer.observe(tableContainer, {
            childList: true,
            subtree: true
        });
    }

    // =-=-=-=-=-=-=-=-=-=-=-=-= 2.7 选课助手 =-=-=-=-=-=-=-=-=-=-=-=-=
    if (window.location.href.includes('/course-selection')) {
        (function () {
            'use strict';

            if (unsafeWindow.courseHelperInitialized) return;
            unsafeWindow.courseHelperInitialized = true;

            // ==============================================================================
            // [1. 配置与核心变量]
            // ==============================================================================
            const API_URL_TEMPLATE = "https://jwxt.nwpu.edu.cn/student/for-std-lessons/info/";
            const TARGET_CELL_SELECTOR = "td div.el-progress";
            const UI_ELEMENT_CLASS = 'course-helper-ui-element';
            const HISTORY_STORAGE_KEY = CONSTANTS.HISTORY_STORAGE_KEY;
            // 拦截 query-lesson 请求时强制放大的分页条数（fetch 与 XHR 通道统一取值）
            const FORCED_PAGE_LIMIT = 100;

            let courseCodeToLessonIdMap = null;

            const originalFetch = unsafeWindow.fetch;
            const originalXhrSend = unsafeWindow.XMLHttpRequest.prototype.send;
            const originalXhrOpen = unsafeWindow.XMLHttpRequest.prototype.open;

            // ==============================================================================
            // [2. 网络拦截与数据解析]
            // ==============================================================================

            function cleanupAndReset() {
                // 数据重置时，清空映射表
                courseCodeToLessonIdMap = null;
            }

            function forceUpdateUI() {
                if (!courseCodeToLessonIdMap) return;
                const tables = document.querySelectorAll('.el-table__body');
                const currentMode = getSelectionMode();
                tables.forEach(tableBody => {
                    if (tableBody.rows.length > 0) {
                        tableBody.querySelectorAll('tr.el-table__row').forEach(row => {
                            processRowWithCode(row, currentMode);
                        });
                    }
                });
            }

            function processApiResponse(responseText) {
                cleanupAndReset();
                try {
                    const data = JSON.parse(responseText);
                    if (data && data.data && data.data.lessons) {
                        courseCodeToLessonIdMap = new Map(data.data.lessons.map(lesson => [lesson.code, lesson.id]));
                    }
                } catch (e) {
                    Logger.error('2.7', '解析课程列表JSON时出错', e);
                }
                // 数据准备好后，通知 UI 刷新
                setTimeout(forceUpdateUI, 500);
            }

            // 解析已选课程列表响应（课程数组直接位于 data 下），重建映射表以支持"已选课程"标签页
            function processSelectedLessonsResponse(responseText) {
                cleanupAndReset();
                try {
                    const data = JSON.parse(responseText);
                    if (data && Array.isArray(data.data)) {
                        courseCodeToLessonIdMap = new Map(
                            data.data
                                .filter(lesson => lesson && lesson.code != null && lesson.id != null)
                                .map(lesson => [lesson.code, lesson.id])
                        );
                    }
                } catch (e) {
                    Logger.error('2.7', '解析已选课程列表JSON时出错', e);
                }
                setTimeout(forceUpdateUI, 500);
            }

            // --- 1. 拦截 Fetch  ---
            unsafeWindow.fetch = function (...args) {
                let [resource, config] = args;
                // 兼容 resource 是 Request 对象的情况
                const requestUrl = resource instanceof Request ? resource.url : resource;

                // 检查是否是查询请求
                if (requestUrl && requestUrl.includes('/query-lesson/')) {
                    // A. 尝试修改请求参数
                    if (config && config.body && typeof config.body === 'string') {
                        try {
                            const data = JSON.parse(config.body);
                            if (data.limit || data.pageSize) {
                                const TARGET = FORCED_PAGE_LIMIT;
                                if (data.limit) data.limit = TARGET;
                                if (data.pageSize) data.pageSize = TARGET;
                                config.body = JSON.stringify(data);
                            }
                        } catch (e) { }
                    }

                    // B. 监听响应
                    return originalFetch.apply(this, args).then(response => {
                        const cloned = response.clone();
                        cloned.text().then(text => processApiResponse(text));
                        return response;
                    });
                }

                // 已选课程列表：仅观察响应以建立课程代码→lessonId 映射，不改动请求
                if (requestUrl && requestUrl.includes('/selected-lessons/')) {
                    return originalFetch.apply(this, args).then(response => {
                        const cloned = response.clone();
                        cloned.text().then(text => processSelectedLessonsResponse(text));
                        return response;
                    });
                }
                return originalFetch.apply(this, args);
            };

            // --- 2. 拦截 XHR ---
            unsafeWindow.XMLHttpRequest.prototype.open = function (method, url) {
                this._gm_url = url; // 保存 URL 供 send 使用
                return originalXhrOpen.apply(this, arguments);
            };

            // 拦截 Send 修改分页参数
            unsafeWindow.XMLHttpRequest.prototype.send = function (data) {
                this.addEventListener('load', function () {
                    if (this.responseURL && this.responseURL.includes('/query-lesson/')) {
                        processApiResponse(this.responseText);
                    } else if (this.responseURL && this.responseURL.includes('/selected-lessons/')) {
                        processSelectedLessonsResponse(this.responseText);
                    }
                }, { once: true });

                if (this._gm_url && this._gm_url.includes('/query-lesson/')) {
                    try {
                        if (typeof data === 'string') {
                            let jsonData = JSON.parse(data);
                            // 强制修改 limit / pageSize
                            if (jsonData.hasOwnProperty('limit') || jsonData.hasOwnProperty('pageSize')) {
                                const TARGET_LIMIT = FORCED_PAGE_LIMIT;

                                if (jsonData.limit) jsonData.limit = TARGET_LIMIT;
                                if (jsonData.pageSize) jsonData.pageSize = TARGET_LIMIT;

                                // 重新打包数据
                                data = JSON.stringify(jsonData);
                            }
                        }
                    } catch (e) {
                        // 静默失败
                    }
                }

                return originalXhrSend.apply(this, [data]);
            };

            // ==============================================================================
            // [3. 辅助功能函数]
            // ==============================================================================

            function getSelectionMode() {
                const semesterSpan = document.querySelector('div.course-select-semester > span');
                if (!semesterSpan) return 'unknown';
                const modeText = semesterSpan.textContent || '';
                if (modeText.includes('直选')) return 'direct';
                else return 'wishlist';
            }

            function injectDirectSelectionUI(row, lessonId) {
                fetch(`${API_URL_TEMPLATE}${lessonId}`)
                    .then(response => response.ok ? response.text() : Promise.reject(`HTTP error! status: ${response.status}`))
                    .then(htmlString => {
                        const parser = new DOMParser();
                        const doc = parser.parseFromString(htmlString, "text/html");
                        let releaseCount = -1;
                        const ths = doc.querySelectorAll('th');
                        for (const th of ths) {
                            if (th.textContent.trim() === '待释放保留人数') {
                                const td = th.nextElementSibling;
                                if (td) { releaseCount = parseInt(td.textContent.trim(), 10) || 0; }
                                break;
                            }
                        }
                        const targetCell = row.querySelector(TARGET_CELL_SELECTOR);
                        if (targetCell && targetCell.parentElement) {
                            const existingElement = targetCell.parentElement.querySelector(`.${UI_ELEMENT_CLASS}`);
                            if (existingElement) existingElement.remove();
                            if (releaseCount > 0) {
                                const displayElement = document.createElement('div');
                                displayElement.className = `${UI_ELEMENT_CLASS} gm-release-tag`;
                                displayElement.textContent = `待释放保留人数: ${releaseCount}`;
                                targetCell.parentElement.appendChild(displayElement);
                            }
                        }
                    })
                    .catch(error => { /* 静默 */ });
            }

            async function injectWishlistUI(row, lessonId) {
                // 读取存储的历史数据
                const historyJSON = await GM_getValue(HISTORY_STORAGE_KEY, '{}');
                let history = {};
                try {
                    history = JSON.parse(historyJSON);
                } catch (e) { history = {}; }

                const courseHistory = history[lessonId];
                const targetContainer = row.querySelector('td:nth-child(5) > .cell'); // 适配 ElementUI 表格列
                if (targetContainer) {
                    const existingElement = targetContainer.querySelector(`.${UI_ELEMENT_CLASS}`);
                    if (existingElement) existingElement.remove();
                    if (courseHistory && courseHistory.length > 0) {
                        const latestRecord = courseHistory[courseHistory.length - 1];
                        const { stdCount, limitCount, time } = latestRecord;
                        const isFull = stdCount >= limitCount;
                        const displayElement = document.createElement('span');
                        displayElement.className = `${UI_ELEMENT_CLASS} gm-hist-tag ${isFull ? 'gm-hist-tag--full' : 'gm-hist-tag--ok'}`;
                        displayElement.textContent = `记录:${stdCount}/${limitCount}`;
                        displayElement.title = `上次同步时间: ${time}`;
                        targetContainer.appendChild(displayElement);
                    }
                }
            }

            function injectCollapseControl() {
                const cols = document.querySelectorAll('.el-col.el-col-24');
                let targetContainer = null;

                for (const col of cols) {
                    if (col.innerText.includes('【主修】')) {
                        targetContainer = col;
                        break;
                    }
                }

                if (!targetContainer) return; // 未找到目标
                if (document.getElementById('gm-collapse-toggle-btn')) return; // 防止重复

                const btn = document.createElement('button');
                btn.id = 'gm-collapse-toggle-btn';

                // 样式调整：
                btn.className = 'el-button el-button--primary el-button--small';

                btn.innerHTML = '<i class="el-icon-s-operation"></i> 全部展开/折叠';

                // CSS调整：右浮动 + 阴影 + 字体加粗
                btn.style.cssText = `
                 float: right;               /* 靠最右侧 */
                 margin-right: 5px;          /* 右侧留一点缝隙 */
                 margin-top: 5px;           /* 微调垂直位置，使其垂直居中 */
                 font-weight: bold;          /* 字体加粗 */
                 font-size: 14px;            /* 字体加大 */
                 box-shadow: 0 4px 12px rgba(64, 158, 255, 0.5); /* 添加蓝色光晕阴影，增加显眼度 */
                 transition: all 0.3s;
            `;

                // 4. 绑定点击逻辑
                let isExpanded = true;
                btn.onclick = (e) => {
                    e.stopPropagation();
                    isExpanded = !isExpanded;

                    // 添加点击动画效果
                    btn.style.transform = 'scale(0.95)';
                    setTimeout(() => btn.style.transform = 'scale(1)', 150);

                    document.querySelectorAll('.course-module').forEach(mod => {
                        const icon = mod.querySelector('i');
                        if (!icon) return;
                        const isOpen = icon.classList.contains('el-icon-caret-bottom');

                        if ((isExpanded && !isOpen) || (!isExpanded && isOpen)) {
                            mod.click();
                        }
                    });

                    btn.innerHTML = isExpanded ? '<i class="el-icon-folder-opened"></i> 全部折叠' : '<i class="el-icon-folder"></i> 全部展开';
                    btn.blur();
                };

                // 5. 插入 DOM
                targetContainer.appendChild(btn);
            }

            // ==============================================================================
            // [4. 核心逻辑]
            // ==============================================================================

            function processRowWithCode(row, mode) {
                let courseCode = null;
                // 1. 尝试获取课程代码
                const accurateCodeElement = row.querySelector('div.lesson-code > a.link-url');
                if (accurateCodeElement) {
                    courseCode = accurateCodeElement.textContent.trim();
                } else {
                    const fallbackCodeElement = row.querySelector('td:first-child span.el-tooltip');
                    if (fallbackCodeElement) courseCode = fallbackCodeElement.textContent.trim();
                }

                // 2. [Diff 检查]：防止重复渲染导致的闪烁
                // 如果当前行已经标记了代码，且代码未变，说明是同一行，仅更新状态颜色，不重绘 DOM
                if (row.dataset.gmCurrentCode === courseCode) {
                    const existingBtn = row.querySelector('.gm-follow-btn');
                    // 如果按钮存在且挂载了 updateState 方法，直接调用更新颜色
                    if (existingBtn && existingBtn.updateState) {
                        existingBtn.updateState();
                    }
                    return;
                }

                // 3. [清理旧状态]：如果代码变了（说明翻页了，DOM 被复用），清除旧的样式和元素
                if (row.dataset.gmCurrentCode) {
                    row.style.backgroundColor = '';
                    row.style.boxShadow = '';
                    row.style.transition = '';
                    const nameEl = row.querySelector('.course-name');
                    if (nameEl) {
                        nameEl.style.fontWeight = '';
                        nameEl.style.color = '';
                    }
                    row.querySelectorAll('.gm-follow-btn, .course-helper-ui-element').forEach(el => el.remove());
                    delete row.dataset.gmCurrentCode;
                }

                // 4. [注入新状态]
                if (courseCode && courseCodeToLessonIdMap && courseCodeToLessonIdMap.has(courseCode)) {
                    // 标记当前行归属
                    row.dataset.gmCurrentCode = courseCode;
                    const lessonId = courseCodeToLessonIdMap.get(courseCode);
                    const nameEl = row.querySelector('.course-name');

                    // --- 注入交互式关注按钮 ---
                    if (nameEl) {
                        const btn = document.createElement('span');
                        btn.className = 'gm-follow-btn';
                        btn.innerHTML = '<svg viewBox="0 0 24 24" width="18" height="18" fill="currentColor" stroke="none"><path d="M20.84 4.61a5.5 5.5 0 0 0-7.78 0L12 5.67l-1.06-1.06a5.5 5.5 0 0 0-7.78 7.78l1.06 1.06L12 21.23l7.78-7.78 1.06-1.06a5.5 5.5 0 0 0 0-7.78z"></path></svg>';
                        btn.style.cssText = `cursor: pointer; font-size: 18px; margin-left: 8px; line-height: 1; user-select: none; transition: all 0.2s; display: inline-block; vertical-align: middle;`;
                        btn.title = "点击关注课程";

                        // 挂载状态更新函数
                        btn.updateState = () => {
                            if (typeof FollowManager !== 'undefined' && FollowManager.has(lessonId)) {
                                // 已关注样式：深红、加粗、粉背景、内阴影
                                btn.title = '点击取消关注';
                                btn.style.color = '#f56c6c';
                                btn.style.textShadow = 'none';
                                btn.style.transform = 'scale(1.08)';
                                nameEl.style.fontWeight = 'bold';
                                nameEl.style.color = '#c53030';
                                row.style.backgroundColor = '#fff1f2';
                                row.style.boxShadow = 'inset 3px 0 0 #f56c6c';
                            } else {
                                // 未关注样式：浅灰
                                btn.title = '点击关注课程';
                                btn.style.color = '#e4e7ed';
                                btn.style.textShadow = 'none';
                                btn.style.transform = 'scale(1)';
                                nameEl.style.fontWeight = '';
                                nameEl.style.color = '';
                                row.style.backgroundColor = '';
                                row.style.boxShadow = '';
                            }
                        };
                        btn.updateState(); // 初始化调用

                        // 绑定点击事件
                        btn.onclick = (e) => {
                            e.stopPropagation();
                            if (typeof FollowManager === 'undefined') { alert('功能未加载'); return; }

                            if (FollowManager.has(lessonId)) {
                                FollowManager.remove(lessonId);
                                showToast('已取消关注', 'cancel');
                            } else {
                                // --- 数据抓取 ---
                                let teachers = '待定', credits = '-', timeAndPlace = '-';
                                try {
                                    // 教师：第3列
                                    const teacherEl = row.querySelector('td:nth-child(3) .course-teacher');
                                    if (teacherEl) teachers = teacherEl.innerText.replace(/[\r\n]+/g, ' ').trim();

                                    // 时间地点：第4列
                                    const placeEl = row.querySelector('td:nth-child(4) .dateTimePlace');
                                    if (placeEl) {
                                        const tooltipDiv = placeEl.querySelector('.tooltip-dateTimePlace span');
                                        timeAndPlace = (tooltipDiv ? tooltipDiv.innerText : placeEl.innerText).replace(/[\r\n]+/g, '; ').trim();
                                    }

                                    // 学分：第1列下方
                                    const infoEl = row.querySelector('td:nth-child(1) .text-color-6');
                                    if (infoEl) {
                                        const creditMatch = infoEl.innerText.match(/([\d\.]+)学分/);
                                        if (creditMatch) credits = creditMatch[1];
                                    }
                                } catch (err) { }

                                // --- 学期提取 (从页面标题) ---
                                let targetSemester = '选课页面关注';
                                try {
                                    const semesterEl = document.querySelector('span[title*="选课"]');
                                    if (semesterEl) {
                                        const rawText = semesterEl.getAttribute('title') || semesterEl.innerText;
                                        const match = rawText.match(/(\d{4}-\d{4}[春夏秋冬])/);
                                        if (match) targetSemester = match[1];
                                    }
                                } catch (e) { }

                                FollowManager.add(lessonId, {
                                    id: lessonId, code: courseCode, name: nameEl.innerText.trim(),
                                    teachers, credits, timeAndPlace, semester: targetSemester, addedTime: new Date().toLocaleString()
                                });
                                showToast(`已关注 ${courseCode}`, 'success');
                            }
                            btn.updateState();
                        };

                        btn.onmouseenter = () => { if (!FollowManager.has(lessonId)) btn.style.color = '#fbc4c4'; };
                        btn.onmouseleave = () => { if (!FollowManager.has(lessonId)) btn.style.color = '#e4e7ed'; };

                        nameEl.appendChild(btn);
                    }

                    // --- 注入其他辅助信息 ---
                    if (mode === 'direct') {
                        injectDirectSelectionUI(row, lessonId);
                    } else {
                        injectWishlistUI(row, lessonId);

                    }
                }
            }

            // ==============================================================================
            // [5. 初始化]
            // ==============================================================================

            function main() {
                let debounceTimer = null;
                const mainObserver = new MutationObserver(() => {
                    if (debounceTimer) clearTimeout(debounceTimer);
                    // 50ms 防抖，检测到 DOM 变动停止后执行 UI 更新
                    debounceTimer = setTimeout(() => {
                        if (courseCodeToLessonIdMap) forceUpdateUI();
                        injectCollapseControl();
                    }, 50);
                });
                const container = document.getElementById('app-content') || document.body;
                mainObserver.observe(container, { childList: true, subtree: true });
            }

            if (document.readyState === 'loading') {
                document.addEventListener('DOMContentLoaded', main, { once: true });
            } else {
                main();
            }
        })();
    }

    // =-=-=-=-=-=-=-=-=-=-=-=-= 2.8 教务系统静默登录（CAS SSO） =-=-=-=-=-=-=-=-=-=-=-=-=
    // 门户与教务同属统一身份认证，门户会话有效时访问 /student/sso-login 可无感换票；
    // 直接访问受保护页面只会落到教务自己的 /student/login，并不会触发 CAS。
    const JwxtSso = {
        AUTH_ENTRY: 'https://jwxt.nwpu.edu.cn/student/sso-login',
        AUTH_FRAME_ID: 'gm_jwxt_sso_frame',
        _sessionPromise: null,

        // 轻量探测：course-table 需要登录，未登录时返回登录页 HTML，不含 dataId
        probe() {
            return new Promise(resolve => {
                GM_xmlhttpRequest({
                    method: 'GET',
                    url: 'https://jwxt.nwpu.edu.cn/student/for-std/course-table',
                    timeout: 6000,
                    onload: res => resolve(res.status === 200 && /var\s+dataId\s*=\s*(\d+)\s*;/.test(res.responseText || '')),
                    onerror: () => resolve(false),
                    ontimeout: () => resolve(false)
                });
            });
        },

        preflight(timeout = 8000) {
            return new Promise(resolve => {
                const old = document.getElementById(this.AUTH_FRAME_ID);
                if (old) old.remove();

                const frame = document.createElement('iframe');
                frame.id = this.AUTH_FRAME_ID;
                frame.style.cssText = 'position: fixed; top: 0; left: -15000px; width: 1280px; height: 800px; border: none; visibility: visible; z-index: -100;';

                let settled = false;
                let debounce = null;
                const done = () => {
                    if (settled) return;
                    settled = true;
                    clearTimeout(debounce);
                    clearTimeout(cap);
                    setTimeout(() => { if (frame.parentNode) frame.remove(); }, 1500);
                    resolve();
                };
                const cap = setTimeout(done, timeout);
                // CAS 换票会连续触发多次 load，等最后一次 load 后再留 800ms 给 Cookie 落盘
                frame.addEventListener('load', () => {
                    clearTimeout(debounce);
                    debounce = setTimeout(done, 800);
                });

                frame.src = this.AUTH_ENTRY;
                document.body.appendChild(frame);
            });
        },

        // 返回 true 表示教务会话可用；在教务页面内天然可用，无需预检
        async ensureSession() {
            if (location.host === 'jwxt.nwpu.edu.cn') return true;
            if (this._sessionPromise) return this._sessionPromise;

            this._sessionPromise = (async () => {
                Logger.log("2.8", "开始探测教务会话...");
                const first = await this.probe();
                Logger.log("2.8", "教务会话探测结果: " + first);
                if (first) return true;

                Logger.log("2.8", "开始统一认证静默预检...");
                await this.preflight();
                // GM 探测可能因 WAF 挑战误判，预检完成后交给 Worker 的真实请求判定；
                // 若仍未登录，Worker 会回传 AUTH_REQUIRED 再提示用户
                const second = await this.probe();
                Logger.log("2.8", "预检后会话探测结果: " + second);
                return true;
            })().then(ok => {
                if (!ok) this._sessionPromise = null;
                return ok;
            });
            return this._sessionPromise;
        }
    };

    // =-=-=-=-=-=-=-=-=-=-=-=-= 2.8 后台静默数据同步 =-=-=-=-=-=-=-=-=-=-=-=-=
    const BackgroundSyncSystem = {
        WORKER_NAME: 'gm_bg_sync_worker_frame',

        /**
         * 数据迁移：清理旧版字段映射错误的后台同步数据
         * @description 旧版 parseCourseData 把班级范围名（"全校"）存进了 name、把 scheduleText 对象存进了 timeAndPlace。
         * 通过检测 timeAndPlace 是否为对象来判断是否为旧版脏数据，清除后本次页面会自动重新同步。
         */
        migrateStaleBackgroundData() {
            try {
                const bgRaw = GM_getValue(CONSTANTS.BACKGROUND_SYNC_KEY, '');
                if (!bgRaw) return;
                const bg = JSON.parse(bgRaw);
                // 旧版数据特征：timeAndPlace 为对象（scheduleText 对象），或 name 为班级范围名"全校"
                const isStale = Array.isArray(bg) && bg.some(c =>
                    (c && typeof c.timeAndPlace === 'object') || (c && c.name === '全校')
                );
                if (isStale) {
                    Logger.log("2.8", "检测到旧版后台同步数据字段错位，已清除，等待重新同步");
                    GM_setValue(CONSTANTS.BACKGROUND_SYNC_KEY, '');
                    GM_setValue(CONSTANTS.LAST_SYNC_TIME_KEY, 0);
                }
            } catch (e) {
                Logger.log("2.8", "后台同步数据迁移检查失败，忽略: " + (e && e.message));
            }
        },

        // 主控逻辑（登录失效时静默跳过，不弹窗打扰）
        async initController() {
            // 旧版字段映射错误的数据迁移
            // 检测到异常数据则清除后台同步缓存并重置时间戳，触发本次页面立即重新同步
            this.migrateStaleBackgroundData();

            const lastSync = GM_getValue(CONSTANTS.LAST_SYNC_TIME_KEY, 0);
            const now = Date.now();

            if (now - lastSync < CONSTANTS.SYNC_COOLDOWN_MS) {
                const remainingMs = CONSTANTS.SYNC_COOLDOWN_MS - (now - lastSync);
                const remainingMins = Math.ceil(remainingMs / 1000 / 60);
                Logger.log("2.8", `处于冷却期，下次自动同步需等待 ${remainingMins} 分钟`);
                return;
            }

            if (isBackgroundSyncing) {
                Logger.log("2.8", "已有后台同步任务进行中，跳过本次触发");
                return;
            }

            Logger.log("2.8", "检查教务登录状态...");
            isBackgroundSyncing = true;
            updateMenuButtonsState(isDataReady);

            // 门户页面与教务是跨域站点，直接请求受保护页面只会落到教务自己的登录页；
            // 先走 /student/sso-login 利用统一认证会话无感换票
            let authOk = false;
            try {
                authOk = await JwxtSso.ensureSession();
            } catch (e) {
                Logger.error("2.8", "教务登录预检失败", e);
            }
            if (!authOk) {
                isBackgroundSyncing = false;
                updateMenuButtonsState(isDataReady);
                Logger.log("2.8", "教务登录已失效，静默跳过本次后台同步");
                return;
            }

            Logger.log("2.8", "准备创建后台 Iframe...");
            const oldFrame = document.getElementById('gm_bg_sync_frame');
            if (oldFrame) oldFrame.remove();

            // 复用主页面缓存的 StudentID，通过 URL 参数传给 worker，避免 worker 内重复请求
            const sid = (await getStudentId().catch(() => null)) || '';

            const iframe = document.createElement('iframe');
            iframe.id = 'gm_bg_sync_frame';
            iframe.name = this.WORKER_NAME;
            iframe.src = `https://jwxt.nwpu.edu.cn/student/for-std/lesson-search?gmStudentId=${sid}`;
            iframe.style.cssText = `position: fixed; top: 0; left: -15000px; width: 1440px; height: 900px; border: none; visibility: visible; z-index: -100;`;
            document.body.appendChild(iframe);

            let watchDog = null;
            const cleanup = (delay = 0) => {
                isBackgroundSyncing = false;
                updateMenuButtonsState(isDataReady);
                if (watchDog) clearTimeout(watchDog);
                window.removeEventListener('message', messageHandler);
                setTimeout(() => {
                    const frame = document.getElementById('gm_bg_sync_frame');
                    if (frame) frame.remove();
                }, delay);
            };

            const messageHandler = (event) => {
                if (!event.data) return;
                if (event.data.type === 'GM_BG_SYNC_AUTH_REQUIRED') {
                    Logger.log("2.8", "教务会话已失效，后台同步静默中止");
                    JwxtSso._sessionPromise = null;
                    cleanup(0);
                    return;
                }
                if (event.data.type === 'GM_BG_SYNC_COMPLETE') {
                    Logger.log("2.8", `后台同步完成。抓取: ${event.data.count}`);
                    if (event.data.count > 0) GM_setValue(CONSTANTS.LAST_SYNC_TIME_KEY, Date.now());
                    cleanup(2000);
                }
            };
            window.addEventListener('message', messageHandler);

            watchDog = setTimeout(() => {
                Logger.log("2.8", "后台同步超时（90秒），已中止");
                cleanup(0);
            }, 90000);
        },

        // Worker 逻辑（API 直接获取）
        startWorker() {
            Logger.info("Sync-Worker", "启动");

            // 未登录时 iframe 会落到 /student/login 或 sso-login，这里不再继续请求接口
            if (!location.pathname.includes('/student/for-std/lesson-search')) {
                Logger.log("Sync-Worker", "当前不在开课查询页，判定为教务登录失效: " + location.pathname);
                window.top.postMessage({ type: 'GM_BG_SYNC_AUTH_REQUIRED' }, '*');
                return;
            }

            const sleep = (ms) => new Promise(r => setTimeout(r, ms));
            const API_BASE = '/student/for-std/lesson-search/semester';
            const API_FIELDS = 'course.code,minorCourse.nameZh,courseType,openDepartment,teacherAssignmentList,examMode,campus,teachLang,roomType,timeTableLayout,crossBizTypes,courseProperty';
            const PAGE_SIZE = 1000; // 使用 queryPage__ 参数格式，服务端支持最大1000条/页

            // 解析 API 返回的课程数据
            // 课程名在 course.course.nameZh（顶层 nameZh 是班级范围名如"全校"）
            const parseCourseData = (courses, semesterName) => {
                return courses.map(course => ({
                    id: String(course.id),
                    code: course.code || '',
                    name: course.course?.nameZh || course.nameZh || course.nameEn || '',
                    teachers: course.teacherAssignmentStr || '待定',
                    credits: course.course?.credits || course.course?.credit || '',
                    // scheduleText 是对象，取 dateTimePlaceText.textZh（时间+地点）
                    timeAndPlace: course.scheduleText?.dateTimePlaceText?.textZh
                        || course.scheduleText?.dateTimePlacePersonText?.textZh
                        || '详见课表',
                    stdCount: course.stdCount || 0,
                    limitCount: course.limitCount || 0,
                    semester: semesterName,
                    updateTime: Date.now()
                }));
            };

            // 自动化执行流程
            const runAutomation = async () => {
                try {
                    // 优先复用主页面传入的 StudentID（URL 参数），否则重新获取
                    const sidFromUrl = new URLSearchParams(window.location.search).get('gmStudentId');
                    const studentId = sidFromUrl || await getStudentId();
                    if (!studentId) throw new Error('无法获取学生ID');

                    // API 请求封装（需要 studentId）
                    const fetchCoursePage = (semesterId, page, pageSize) => {
                        return new Promise((resolve, reject) => {
                            const timestamp = Date.now();
                            const url = API_BASE + '/' + semesterId + '/search/' + studentId +
                                '?bizTypeAssoc=1&queryPage__=' + page + ',' + pageSize +
                                '&assembleFields=' + encodeURIComponent(API_FIELDS) +
                                '&_=' + timestamp;

                            const xhr = new XMLHttpRequest();
                            xhr.open('GET', url);
                            xhr.setRequestHeader('X-Requested-With', 'XMLHttpRequest');
                            xhr.onload = function () {
                                if (xhr.status === 200) {
                                    try { resolve(JSON.parse(xhr.responseText)); }
                                    catch (e) { reject(new Error('API响应解析失败')); }
                                } else {
                                    reject(new Error('API请求失败: HTTP ' + xhr.status));
                                }
                            };
                            xhr.onerror = function () { reject(new Error('网络请求失败')); };
                            xhr.send();
                        });
                    };

                    // 通过API获取学期列表（替代DOM操作）
                    Logger.log('Sync-Worker', '正在通过API获取学期列表...');
                    let semesterOptions = [];

                    const fetchSemesterList = () => {
                        return new Promise((resolve, reject) => {
                            const url = '/student/for-std/lesson-search/index/' + studentId;
                            const xhr = new XMLHttpRequest();
                            xhr.open('GET', url);
                            xhr.onload = function () {
                                if (xhr.status === 200) {
                                    try {
                                        const html = xhr.responseText;
                                        const parser = new DOMParser();
                                        const doc = parser.parseFromString(html, 'text/html');

                                        // 解析学期下拉选项
                                        const options = doc.querySelectorAll('#semester option[value]');
                                        const semesterList = [];
                                        options.forEach(opt => {
                                            const val = opt.getAttribute('value');
                                            const text = opt.textContent.trim();
                                            if (val && text && val !== '') {
                                                semesterList.push({ value: val, text: text });
                                            }
                                        });

                                        resolve(semesterList);
                                    } catch (e) {
                                        reject(new Error('HTML解析失败: ' + e.message));
                                    }
                                } else {
                                    reject(new Error('请求失败: HTTP ' + xhr.status));
                                }
                            };
                            xhr.onerror = function () {
                                reject(new Error('网络请求失败'));
                            };
                            xhr.send();
                        });
                    };

                    try {
                        semesterOptions = await fetchSemesterList();
                        Logger.log('Sync-Worker', '获取到 ' + semesterOptions.length + ' 个学期选项');
                    } catch (e) {
                        Logger.error('Sync-Worker', 'API获取学期列表失败: ' + e.message);
                        throw new Error('无法获取学期列表: ' + e.message);
                    }

                    if (semesterOptions.length === 0) {
                        // 登录页 / WAF 挑战页都没有 #semester，按登录失效处理
                        const authErr = new Error('未找到任何学期选项');
                        authErr.authRequired = true;
                        throw authErr;
                    }

                    // ================== 1. 通过API寻找最新有数据的学期 ==================
                    let activeSemesterName = '未知学期';
                    let activeSemesterId = '';
                    let foundValidSemester = false;

                    for (let i = 0; i < semesterOptions.length; i++) {
                        const opt = semesterOptions[i];
                        Logger.log('Sync-Worker', 'API探测学期: ' + opt.text);

                        try {
                            const testData = await fetchCoursePage(opt.value, 1, 1);
                            if (testData && testData.data && testData.data.length > 0) {
                                activeSemesterName = opt.text;
                                activeSemesterId = opt.value;
                                foundValidSemester = true;
                                Logger.log('Sync-Worker', '寻址成功: ' + activeSemesterName);
                                break;
                            }
                        } catch (e) {
                            Logger.log('Sync-Worker', '学期 ' + opt.text + ' 探测失败，跳过');
                        }
                    }

                    if (!foundValidSemester) {
                        throw new Error('遍历了所有学期均未找到排课数据');
                    }

                    // ================== 2. 通过API并行获取全部数据 ==================
                    // 第一次请求：获取 totalPages 信息
                    const firstResult = await fetchCoursePage(activeSemesterId, 1, PAGE_SIZE);
                    if (!firstResult || !firstResult.data || firstResult.data.length === 0) {
                        throw new Error(activeSemesterName + ' 学期暂无课程数据');
                    }

                    const totalPages = firstResult._page_?.totalPages || 1;
                    let allCourseData = parseCourseData(firstResult.data, activeSemesterName);
                    Logger.log('Sync-Worker', '已获取 ' + allCourseData.length + ' 条 (第1/' + totalPages + '页)');

                    // 如果有更多页，并行请求剩余所有页
                    if (totalPages > 1) {
                        Logger.log('Sync-Worker', '正在获取剩余 ' + (totalPages - 1) + ' 页...');
                        const pagePromises = [];
                        for (let page = 2; page <= totalPages; page++) {
                            pagePromises.push(fetchCoursePage(activeSemesterId, page, PAGE_SIZE));
                        }

                        const results = await Promise.all(pagePromises);
                        for (const result of results) {
                            if (result && result.data && result.data.length > 0) {
                                allCourseData = allCourseData.concat(parseCourseData(result.data, activeSemesterName));
                            }
                        }
                        Logger.log('Sync-Worker', '获取完成，共 ' + allCourseData.length + ' 条');
                    }

                    Logger.log('2.8', '全部完成! 存储 ' + allCourseData.length + ' 条。');
                    GM_setValue(CONSTANTS.BACKGROUND_SYNC_KEY, JSON.stringify(allCourseData));
                    window.top.postMessage({ type: 'GM_BG_SYNC_COMPLETE', count: allCourseData.length }, '*');

                } catch (err) {
                    Logger.error("2.8", "后台同步 Worker 异常", err);
                    if (err && err.authRequired) {
                        window.top.postMessage({ type: 'GM_BG_SYNC_AUTH_REQUIRED' }, '*');
                    } else {
                        window.top.postMessage({ type: 'GM_BG_SYNC_COMPLETE', count: 0 }, '*');
                    }
                }
            };

            setTimeout(runAutomation, 1500);
        }
    };

    // =-=-=-=-=-=-=-=-=-=-=-=-= 2.9 培养方案课程代码智能预览 =-=-=-=-=-=-=-=-=-=-=-=-=
    function initProgramPageEnhancement() {
        // 检查功能开关
        if (!ConfigManager.enableCourseWatch) {
            return;
        }
        Logger.log('2.9', "初始化培养方案课程预览");

        // 1. 数据准备
        const bgDataStr = GM_getValue(CONSTANTS.BACKGROUND_SYNC_KEY);
        if (!bgDataStr) return;

        let courseDB;
        try { courseDB = JSON.parse(bgDataStr); } catch (e) { return; }
        if (!courseDB || courseDB.length === 0) return;

        // 构建索引 (Parent Code -> List of Courses)
        const courseMap = new Map();
        courseDB.forEach(c => {
            if (!c.code) return;
            // 提取课程代码前缀 (例如 U14M11003.01 -> U14M11003)
            const parentCode = c.code.trim().split('.')[0];
            if (!courseMap.has(parentCode)) courseMap.set(parentCode, []);
            courseMap.get(parentCode).push(c);
        });

        // 定义高清 SVG 图标
        const svgs = {
            book: `<svg viewBox="0 0 24 24" width="18" height="18" style="fill:none;stroke:#409EFF;stroke-width:2;stroke-linecap:round;stroke-linejoin:round;vertical-align:-4px;"><path d="M4 19.5A2.5 2.5 0 0 1 6.5 17H20"></path><path d="M6.5 2H20v20H6.5A2.5 2.5 0 0 1 4 19.5v-15A2.5 2.5 0 0 1 6.5 2z"></path></svg>`,
            user: `<svg viewBox="0 0 24 24" width="14" height="14" style="fill:none;stroke:#5B6B80;stroke-width:2;stroke-linecap:round;stroke-linejoin:round;margin-right:6px;flex:none;"><path d="M20 21v-2a4 4 0 0 0-4-4H8a4 4 0 0 0-4 4v2"></path><circle cx="12" cy="7" r="4"></circle></svg>`,
            pin: `<svg viewBox="0 0 24 24" width="14" height="14" style="fill:none;stroke:#5B6B80;stroke-width:2;stroke-linecap:round;stroke-linejoin:round;margin-right:6px;flex:none;"><path d="M21 10c0 7-9 13-9 13s-9-6-9-13a9 9 0 0 1 18 0z"></path><circle cx="12" cy="10" r="3"></circle></svg>`
        };

        // 2. 注入美化后的 CSS
        if (!document.getElementById('gm-program-tooltip-style')) {
            const style = getOrCreateStyleEl('gm-program-tooltip-style');
            style.textContent = `
            /* 课程代码高亮样式：全包围虚线 */
            .gm-course-code-highlight {
                border: 1.5px dashed var(--gm-primary);
                color: var(--gm-primary-text);
                font-weight: 600;
                font-size: 13px;
                cursor: pointer;
                background-color: var(--gm-primary-light);
                padding: 2px 6px;
                border-radius: var(--gm-radius-sm);
                transition: all 0.2s;
            }
            .gm-course-code-highlight:hover {
                background-color: var(--gm-primary-border);
                color: var(--gm-primary-active);
                border-color: var(--gm-primary-active);
            }

            /* 弹窗容器 - 磨砂玻璃质感 */
            .gm-program-tooltip {
                position: fixed; z-index: 100001;
                background: #fff;
                border: 1px solid var(--gm-border);
                box-shadow: var(--gm-shadow-lg);
                border-radius: var(--gm-radius-lg); padding: 0;
                width: 440px;
                font-family: var(--gm-font);
                display: none; opacity: 0; transition: opacity 0.2s ease, transform 0.2s ease;
                pointer-events: auto;
                transform: translateY(5px);
            }
            .gm-program-tooltip.show { display: block; opacity: 1; transform: translateY(0); }

            /* 头部样式 */
            .gm-pt-header {
                background: var(--gm-fill);
                padding: 14px 18px;
                border-bottom: 1px solid var(--gm-border);
                font-weight: 700; color: var(--gm-text-1); font-size: 14px;
                display:flex; justify-content:space-between; align-items: center;
                border-radius: var(--gm-radius-lg) var(--gm-radius-lg) 0 0;
                letter-spacing: 0;
            }
            .gm-pt-badge {
                font-weight: 600; color: var(--gm-primary-text); font-size: 11.5px;
                background: var(--gm-primary-light);
                border: 1px solid var(--gm-primary-border);
                padding: 2px 8px; border-radius: var(--gm-radius-sm);
                white-space: nowrap;
            }

            /* 列表区域 */
            .gm-pt-list { max-height: 420px; overflow-y: auto; padding: 0; }

            /* 滚动条美化 */
            .gm-pt-list::-webkit-scrollbar { width: 6px; }
            .gm-pt-list::-webkit-scrollbar-track { background: transparent; }
            .gm-pt-list::-webkit-scrollbar-thumb { background-color: var(--gm-border); border-radius: 3px; }

            /* 单个课程卡片 */
            .gm-pt-item {
                padding: 14px 20px;
                border-bottom: 1px solid #f2f4f7;
                transition: background-color 0.2s;
            }
            .gm-pt-item:last-child { border-bottom: none; }
            .gm-pt-item:hover { background-color: var(--gm-primary-light); }

            /* 标题行：课程名 + 人数标签 */
            .gm-pt-title-row {
                display: flex; justify-content: space-between; align-items: center;
                margin-bottom: 4px;
            }
            .gm-pt-title {
                font-weight: 700; font-size: 14px; color: var(--gm-text-1);
                white-space: nowrap; overflow: hidden; text-overflow: ellipsis;
                flex: 1; min-width: 0; padding-right: 10px;
            }

            /* 代码行：课程代码 + 关注按钮 */
            .gm-pt-code-row {
                display: flex; justify-content: space-between; align-items: center;
                margin-bottom: 8px;
            }
            .gm-pt-code { font-size: 13.5px; font-weight: 600; color: var(--gm-text-3); font-family: var(--gm-mono); }

            /* 元信息区：教师、地点 */
            .gm-pt-meta { display: flex; flex-direction: column; gap: 5px; color: var(--gm-text-2); font-size: 12.5px; line-height: 1.55; }
            .gm-pt-row { display: flex; align-items: center; }
            .gm-pt-row--place { align-items: flex-start; }
            .gm-pt-row--place svg { margin-top: 2px; }

            /* 人数标签：与界面同一字体，数字等宽对齐 */
            .gm-pt-stat {
                font-family: var(--gm-font); font-size: 12.5px; font-weight: 700;
                font-variant-numeric: tabular-nums;
                padding: 2px 7px; border-radius: var(--gm-radius-sm); flex-shrink: 0;
            }
            .gm-tag-full { color: var(--gm-danger-text); background: var(--gm-danger-bg); border: 1px solid var(--gm-danger-border); }
            .gm-tag-avail { color: var(--gm-success-text); background: var(--gm-success-bg); border: 1px solid var(--gm-success-border); }

            /* 关注按钮 */
            .gm-pt-btn {
                cursor: pointer; font-size: 22px; color: #dcdfe6; line-height: 1;
                transition: all 0.2s cubic-bezier(0.175, 0.885, 0.32, 1.275);
                display: flex; align-items: center; justify-content: center;
                width: 32px; height: 32px; border-radius: 50%;
            }
            .gm-pt-btn:hover { transform: scale(1.08); background-color: #fff; box-shadow: var(--gm-shadow-sm); }
            .gm-pt-btn.is-active { color: #f56c6c !important; text-shadow: none; }
        `;
            document.head.appendChild(style);
        }

        let tooltip = document.querySelector('.gm-program-tooltip');
        if (!tooltip) {
            tooltip = document.createElement('div');
            tooltip.className = 'gm-program-tooltip';
            document.body.appendChild(tooltip);
        }

        // 3. 全局事件委托 (处理悬停和点击)
        let hideTimer = null;

        document.body.addEventListener('mouseover', function (e) {
            if (e.target.classList.contains('gm-course-code-highlight')) {
                if (hideTimer) clearTimeout(hideTimer);
                const code = e.target.getAttribute('data-code');
                showTooltip(e.target, code);
            }
            else if (e.target.closest('.gm-program-tooltip')) {
                if (hideTimer) clearTimeout(hideTimer);
            }
        });

        document.body.addEventListener('mouseout', function (e) {
            if (e.target.classList.contains('gm-course-code-highlight') || e.target.closest('.gm-program-tooltip')) {
                hideTimer = setTimeout(() => { tooltip.classList.remove('show'); }, 300);
            }
        });

        document.body.addEventListener('dblclick', function (e) {
            // 检查是否点击了高亮的代码块
            if (e.target.classList.contains('gm-course-code-highlight')) {
                const code = e.target.getAttribute('data-code');
                // 使用剪贴板 API
                navigator.clipboard.writeText(code).then(() => {
                    // 视觉反馈：变为绿色并闪烁一下
                    const originalTransition = e.target.style.transition;
                    const originalBg = e.target.style.backgroundColor;
                    const originalColor = e.target.style.color;

                    e.target.style.transition = 'all 0.1s';
                    e.target.style.backgroundColor = '#f0f9eb';
                    e.target.style.color = '#67C23A';
                    e.target.textContent = '已复制!'; // 临时改变文字提示

                    setTimeout(() => {
                        e.target.textContent = code; // 恢复文字
                        e.target.style.backgroundColor = originalBg;
                        e.target.style.color = originalColor;
                        e.target.style.transition = originalTransition;
                    }, 800);
                }).catch(err => {
                    Logger.error('2.9', '复制失败', err);
                    alert('复制失败，请手动复制');
                });

                // 阻止选中文本的默认行为
                e.preventDefault();
                window.getSelection().removeAllRanges();
            }
        });

        // 处理关注按钮点击
        document.body.addEventListener('click', function (e) {
            const btn = e.target.closest('.gm-pt-btn');
            if (btn) {
                e.stopPropagation();
                handleFollowClick(btn);
            }
        });

        function handleFollowClick(btn) {
            const id = btn.dataset.id.toString();
            const semester = btn.dataset.semester && btn.dataset.semester !== "undefined"
                ? btn.dataset.semester
                : '从培养方案页关注';

            const data = {
                id: id,
                code: btn.dataset.code,
                name: btn.dataset.name,
                semester: semester,
                teachers: btn.dataset.teachers,
                credits: btn.dataset.credits || '-',
                timeAndPlace: btn.dataset.place,
                addedTime: new Date().toLocaleString()
            };

            if (FollowManager.has(id)) {
                FollowManager.remove(id);
                btn.classList.remove('is-active');
                btn.style.color = '#dcdfe6';
            } else {
                FollowManager.add(id, data);
                btn.classList.add('is-active');
                btn.style.color = '#f56c6c';
            }
        }

        // 4. DOM 扫描 (将普通文本转换为高亮节点)
        function processCells() {
            const cells = document.querySelectorAll('td');
            cells.forEach(td => {
                if (td.dataset.gmProcessed) return;
                const rawText = td.textContent;
                if (!rawText) return;
                const text = rawText.trim();
                // 简单的正则匹配课程代码 (大写字母开头，包含数字，长度适中)
                if (text.length >= 5 && text.length <= 15 && /^[A-Z][A-Z0-9]+$/.test(text)) {
                    if (courseMap.has(text)) {
                        td.dataset.gmProcessed = "true";
                        td.innerHTML = `<span class="gm-course-code-highlight" data-code="${text}" title="双击复制课程代码">${text}</span>`;
                    }
                }
            });
        }

        // 5. 显示浮层 (生成HTML)
        function showTooltip(targetEl, code) {
            const courses = courseMap.get(code) || [];
            const rect = targetEl.getBoundingClientRect();

            let contentHTML = '';
            if (courses.length === 0) {
                contentHTML = '<div style="padding:30px;text-align:center;color:#909399;font-size:13px;">本学期暂无开课记录</div>';
            } else {
                contentHTML = `<div class="gm-pt-list">`;
                courses.forEach(c => {
                    const isFull = c.stdCount >= c.limitCount;
                    const countClass = isFull ? 'gm-tag-full' : 'gm-tag-avail';
                    const isFollowed = FollowManager.has(c.id);
                    const activeClass = isFollowed ? 'is-active' : '';
                    const initColor = isFollowed ? '#f56c6c' : '#dcdfe6';

                    const teacherText = c.teachers || '待定';
                    // 兼容旧版脏数据：timeAndPlace 若为对象（scheduleText），提取时间+地点文本
                    let placeText = c.timeAndPlace || '详见课表';
                    if (typeof placeText === 'object') {
                        placeText = placeText?.dateTimePlaceText?.textZh
                            || placeText?.dateTimePlacePersonText?.textZh
                            || '详见课表';
                    }
                    const teacherHtml = esc(teacherText);
                    // 分号分隔多个时间段，逐段换行显示
                    const placeHtml = String(placeText)
                        .split(/[;；]/)
                        .map(s => esc(s.trim()))
                        .filter(Boolean)
                        .join('<br>');

                    contentHTML += `
                    <div class="gm-pt-item">
                        <div class="gm-pt-title-row">
                            <div class="gm-pt-title" title="${esc(c.name)}">${esc(c.name)}</div>
                            <div class="gm-pt-stat ${countClass}">${c.stdCount}/${c.limitCount}</div>
                        </div>
                        <div class="gm-pt-code-row">
                            <div class="gm-pt-code">${esc(c.code)}</div>
                            <div class="gm-pt-btn ${activeClass}" style="color:${initColor}"
                                 data-id="${c.id}" data-code="${esc(c.code)}" data-name="${esc(c.name)}"
                                 data-teachers="${teacherHtml}" data-place="${esc(placeText)}"
                                 data-credits="${esc(c.credits || '')}"
                                 data-semester="${esc(c.semester || '')}"
                                 title="${isFollowed ? '取消关注' : '关注此班级'}"><svg viewBox="0 0 24 24" width="16" height="16" fill="currentColor" stroke="none"><path d="M20.84 4.61a5.5 5.5 0 0 0-7.78 0L12 5.67l-1.06-1.06a5.5 5.5 0 0 0-7.78 7.78l1.06 1.06L12 21.23l7.78-7.78 1.06-1.06a5.5 5.5 0 0 0 0-7.78z"></path></svg></div>
                        </div>
                        <div class="gm-pt-meta">
                            <div class="gm-pt-row">${svgs.user} <span>${teacherHtml}</span></div>
                            <div class="gm-pt-row gm-pt-row--place">${svgs.pin} <span>${placeHtml}</span></div>
                        </div>
                    </div>
                `;
                });
                contentHTML += `</div>`;
            }

            tooltip.innerHTML = `
            <div class="gm-pt-header">
                <span style="display:flex;align-items:center;gap:8px">${svgs.book} <span style="font-family:var(--gm-mono);font-size:17px;font-weight:700;">${esc(code)}</span></span>
                <span class="gm-pt-badge">本学期 ${courses.length} 个班级</span>
            </div>
            ${contentHTML}
        `;

            // 智能定位
            const viewportHeight = window.innerHeight;
            const tooltipHeight = Math.min(500, courses.length * 90 + 125); // 估算高度
            let top = rect.bottom + 8;

            // 如果底部放不下，就放上面
            if (rect.bottom + tooltipHeight > viewportHeight) {
                top = rect.top - tooltipHeight - 10;
                if (top < 10) top = 10; // 防止溢出顶部
            }

            // 水平定位
            let left = rect.left + 80;
            if (left + 440 > window.innerWidth) {
                left = window.innerWidth - 450;
            }

            tooltip.style.left = left + 'px';
            tooltip.style.top = top + 'px';
            tooltip.classList.add('show');
        }

        // 观察页面变化，动态处理新加载的内容
        const observer = new MutationObserver(() => {
            if (window.gm_program_timer) clearTimeout(window.gm_program_timer);
            window.gm_program_timer = setTimeout(processCells, 200);
        });
        const targetNode = document.querySelector('.main-content') || document.body;
        observer.observe(targetNode, { childList: true, subtree: true });

        // 初始执行
        setTimeout(processCells, 500);
        setTimeout(processCells, 1500);
    }

    // =-=-=-=-=-=-=-=-=-=-=-=-= 2.10 选课时间提醒 =-=-=-=-=-=-=-=-=-=-=-=-=
    // 时间表放在仓库 schedule.json 中，按身份和学期维护；脚本运行时拉取并缓存 1 小时
    const SCHEDULE_CONFIG_URLS = [
        'https://gitee.com/leamloli/aoxiang-jwxt-enhance/raw/master/schedule.json',
        'https://raw.githubusercontent.com/zengzoxiong/aoxiang-jwxt-enhance/main/schedule.json'
    ];
    const SCHEDULE_CACHE_TTL = 60 * 60 * 1000;

    function gmGetText(url, timeout = 8000) {
        return new Promise((resolve, reject) => {
            GM_xmlhttpRequest({
                method: 'GET',
                url,
                timeout,
                onload: res => {
                    if (res.status >= 200 && res.status < 300) resolve(res.responseText || '');
                    else reject(new Error('HTTP ' + res.status));
                },
                onerror: () => reject(new Error('网络请求失败')),
                ontimeout: () => reject(new Error('请求超时'))
            });
        });
    }

    function getScheduleRange(schedule) {
        let start = Infinity;
        let end = -Infinity;
        (schedule && schedule.groups ? schedule.groups : []).forEach(group => {
            (group.phases || []).forEach(phase => {
                const s = new Date(phase.start).getTime();
                const e = new Date(phase.end).getTime();
                if (!isNaN(s)) start = Math.min(start, s);
                if (!isNaN(e)) end = Math.max(end, e);
            });
        });
        return { start, end };
    }

    function normalizeSchedule(schedule) {
        if (!schedule || !Array.isArray(schedule.groups)) return schedule;
        const range = getScheduleRange(schedule);
        if (!schedule.expires && isFinite(range.end)) {
            schedule.expires = new Date(range.end + 7 * 24 * 3600 * 1000).toISOString();
        }
        return schedule;
    }

    async function loadScheduleConfig() {
        let cache = null;
        try { cache = JSON.parse(GM_getValue(CONSTANTS.SCHEDULE_CACHE_KEY, 'null')); } catch (e) { cache = null; }

        if (cache && cache.data && Array.isArray(cache.data.schedules) && Date.now() - cache.fetchedAt < SCHEDULE_CACHE_TTL) {
            cache.data.__source = 'cache';
            return cache.data;
        }

        for (const url of SCHEDULE_CONFIG_URLS) {
            try {
                const text = (await gmGetText(url)).replace(/^\uFEFF/, '');
                const data = JSON.parse(text);
                if (data && Array.isArray(data.schedules)) {
                    data.schedules.forEach(normalizeSchedule);
                    GM_setValue(CONSTANTS.SCHEDULE_CACHE_KEY, JSON.stringify({ fetchedAt: Date.now(), data }));
                    data.__source = url.includes('gitee') ? 'gitee' : 'github';
                    return data;
                }
                Logger.warn('2.10', '选课时间表格式不正确: ' + url);
            } catch (e) {
                Logger.warn('2.10', '选课时间表加载失败: ' + url + ' ' + (e && e.message));
            }
        }

        if (cache && cache.data && Array.isArray(cache.data.schedules)) {
            cache.data.__source = 'stale-cache';
            return cache.data;
        }
        Logger.error('2.10', '选课时间表不可用，已跳过选课提醒');
        return { schedules: [], __source: 'none' };
    }

    async function detectUserIdentity() {
        const stored = GM_getValue(CONSTANTS.USER_IDENTITY_KEY, '');
        if (stored === 'graduate' || stored === 'undergraduate') {
            Logger.log('2.10', '使用已同步身份: ' + (stored === 'graduate' ? '研究生' : '本科生'));
            return stored;
        }

        // 1. 页面用户信息区域
        const selectors = ['.user-info', '.user-name', '.student-info', '.avatar-wrapper', '.user-detail', '.header', '.el-header'];
        const headerText = selectors.map(sel => {
            const el = document.querySelector(sel);
            return el ? (el.innerText || '') : '';
        }).join(' ');
        if (/研究生/.test(headerText)) {
            GM_setValue(CONSTANTS.USER_IDENTITY_KEY, 'graduate');
            return 'graduate';
        }
        if (/本科生/.test(headerText)) {
            GM_setValue(CONSTANTS.USER_IDENTITY_KEY, 'undergraduate');
            return 'undergraduate';
        }

        // 2. 画像页当前身份 tab（tab-6 研究生 / tab-1 主修）
        const activeTab = document.querySelector('.el-tabs__item.is-active');
        if (activeTab && activeTab.id === 'tab-6') {
            GM_setValue(CONSTANTS.USER_IDENTITY_KEY, 'graduate');
            return 'graduate';
        }
        if (activeTab && activeTab.id === 'tab-1') {
            GM_setValue(CONSTANTS.USER_IDENTITY_KEY, 'undergraduate');
            return 'undergraduate';
        }

        // 3. getStdInfo 探测（研究生 / 本科生各请求一次）
        const detectFromText = (text) => {
            if (!text) return '';
            if (/研究生/.test(text)) return 'graduate';
            if (/本科生/.test(text)) return 'undergraduate';
            try {
                const str = JSON.stringify(JSON.parse(text));
                if (/研究生/.test(str)) return 'graduate';
                if (/本科生/.test(str)) return 'undergraduate';
            } catch (e) { /* ignore */ }
            return '';
        };
        for (const assoc of [6, 1]) {
            try {
                const text = await gmGetText(`https://jwxt.nwpu.edu.cn/student/for-std/student-portrait/getStdInfo?bizTypeAssoc=2&cultivateTypeAssoc=${assoc}`, 5000);
                const identity = detectFromText(text);
                if (identity) {
                    GM_setValue(CONSTANTS.USER_IDENTITY_KEY, identity);
                    return identity;
                }
            } catch (e) { /* 探测失败按下一个 */ }
        }

        Logger.log('2.10', '未识别到身份，按本科生处理');
        GM_setValue(CONSTANTS.USER_IDENTITY_KEY, 'undergraduate');
        return 'undergraduate';
    }

    function pickSchedule(data, identity) {
        const list = (data.schedules || []).filter(s => s && s.identity === identity && Array.isArray(s.groups) && s.groups.length);
        if (!list.length) return null;

        const now = Date.now();
        const scored = list.map(schedule => ({ schedule, ...getScheduleRange(schedule) }))
            .filter(x => isFinite(x.start) && isFinite(x.end));
        if (!scored.length) return list[0];

        const active = scored.find(x => now >= x.start - 16 * 3600 * 1000 && now <= x.end);
        if (active) return active.schedule;

        const upcoming = scored.filter(x => x.start > now).sort((a, b) => a.start - b.start)[0];
        if (upcoming) return upcoming.schedule;

        const past = scored.filter(x => x.end < now).sort((a, b) => b.end - a.end)[0];
        return (past || scored[0]).schedule;
    }

    async function initScheduleWidget() {
        const remote = await loadScheduleConfig();
        const identity = await detectUserIdentity();
        const identityName = identity === 'graduate' ? '研究生' : '本科生';
        const schedule = pickSchedule(remote, identity);
        if (!schedule) {
            Logger.log('2.10', `未找到${identityName}的选课时间配置，已跳过选课提醒`);
            return;
        }

        const SCHEDULE_CONFIG = {
            EXPIRATION_DATE: schedule.expires || '',
            STORAGE_KEY: schedule.storageKey || `jwxt_schedule_table_closed_${schedule.term || 'term'}_${identity}`,
            COURSE_URL: schedule.courseUrl || remote.courseUrl || 'https://jwxt.nwpu.edu.cn/student/for-std/course-select',
            PRE_NOTIFY_HOURS: Number(schedule.preNotifyHours ?? remote.preNotifyHours ?? 16),
            GROUPS: schedule.groups
        };
        Logger.log('2.10', `选课提醒身份: ${identityName}，时间表: ${schedule.term || '未标注'}（来源: ${remote.__source || 'json'}）`);

        const showWidget = () => {
            if (GM_getValue(SCHEDULE_CONFIG.STORAGE_KEY, false) === true) return;
            // 避免单次页面刷新内重复关闭后弹出
            if (window.gm_schedule_manually_closed) return;

            const now = Date.now();
            const expirationTime = new Date(SCHEDULE_CONFIG.EXPIRATION_DATE).getTime();

            if (isNaN(expirationTime) || now > expirationTime) return;
            if (document.querySelector('.gm-schedule-box')) return;

            // --- 1. 计算当前状态 & 构建表格行 ---
            let statusHtml = '<span style="color: #909399;">当前未处于选课时段</span>';
            let showPreSyncLink = false; // 是否显示同步链接
            let tableRowsHtml = '';

            // 扁平化遍历所有阶段以检查时间
            let activePhaseFound = false;

            SCHEDULE_CONFIG.GROUPS.forEach((group, gIndex) => {
                group.phases.forEach((phase, pIndex) => {
                    const startTime = new Date(phase.start).getTime();
                    const endTime = new Date(phase.end).getTime();
                    const preStartTime = startTime - (SCHEDULE_CONFIG.PRE_NOTIFY_HOURS * 60 * 60 * 1000);

                    // A. 检查状态: 进行中
                    if (!activePhaseFound && now >= startTime && now <= endTime) {
                        statusHtml = `当前处于 <span style="color: #f56c6c; font-weight: bold; border-bottom: 2px solid #f56c6c;">${group.groupName} - ${phase.name}</span>`;
                        activePhaseFound = true;
                    }
                    // B. 检查状态: 即将开始 (正选前N小时提示)
                    else if (!activePhaseFound && phase.type === 'positive' && now >= preStartTime && now < startTime) {
                        const hoursLeft = Math.ceil((startTime - now) / 3600000);
                        statusHtml = `<span style="color: #E65100; font-weight:bold;">${group.groupName}${phase.name}</span> 将于 ${hoursLeft} 小时后开始。` +
                            `<span id="gm-sch-pre-sync" style="color:#409EFF; cursor:pointer; text-decoration:underline; font-weight:bold; margin-left:10px;">[建议提前点击此处同步课程余量数据]</span>`;
                        showPreSyncLink = true;
                        activePhaseFound = true;
                    }

                    // C. 构建表格行
                    // 格式化时间显示 (移除年份，保留 月-日 时:分)
                    const formatTime = (isoStr) => {
                        const d = new Date(isoStr);
                        return `${d.getMonth() + 1}月${d.getDate()}日 ${d.getHours()}点`;
                    };
                    const timeStr = `${formatTime(phase.start)} 至 ${formatTime(phase.end)}`;

                    tableRowsHtml += `<tr>`;
                    // 处理第一列的 Rowspan (合并单元格)
                    if (pIndex === 0) {
                        const borderStyle = gIndex > 0 ? 'border-top:2px solid #ebeef5;' : '';
                        tableRowsHtml += `<td rowspan="${group.phases.length}" style="font-weight:bold; ${borderStyle}">${group.groupName}</td>`;
                    }

                    // 高亮选课方式
                    const methodClass = phase.method.includes('意愿值') || phase.method.includes('直选') ? 'gm-sch-highlight' : '';

                    tableRowsHtml += `
                    <td>${phase.name}</td>
                    <td>${timeStr}</td>
                    <td class="${methodClass}">${phase.method}</td>
                    <td>${phase.scope}</td>
                </tr>`;
                });
            });

            // --- 2. 注入样式 ---
            if (!document.getElementById('gm-schedule-table-style')) {
                const style = getOrCreateStyleEl('gm-schedule-table-style');
                style.textContent = `
                .gm-schedule-box {
                    position: fixed; left: 20px; bottom: 20px; z-index: 9999;
                    background: #fff; padding: 12px; border-radius: 8px;
                    box-shadow: 0 4px 25px rgba(0,0,0,0.15);
                    border: 1px solid #dcdfe6;
                    font-family: "Helvetica Neue", Helvetica, "PingFang SC", "Microsoft YaHei", Arial, sans-serif;
                    animation: gmSlideUp 0.5s ease-out;
                    width: auto; max-width: 680px;
                }
                .gm-status-bar {
                    text-align: center; background: #fdf6ec; padding: 8px;
                    border-radius: 4px; margin-bottom: 10px; font-size: 13px;
                    border: 1px inset #faecd8; line-height: 1.5;
                }
                .gm-sch-table {
                    width: 100%; border-collapse: collapse; font-size: 12px; color: #333;
                    margin-bottom: 10px; border: 1px solid #ebeef5;
                }
                .gm-sch-table th, .gm-sch-table td {
                    border: 1px solid #ebeef5; padding: 6px 8px; text-align: center; vertical-align: middle;
                }
                .gm-sch-table th { background-color: #f5f7fa; font-weight: bold; color: #606266; }
                .gm-sch-highlight { color: #409EFF; font-weight: bold; }
                .gm-schedule-footer {
                    display: flex; justify-content: space-between; align-items: center;
                    font-size: 12px; color: #909399; margin-top: 8px;
                }
                .gm-sch-btn-group { display: flex; gap: 10px; }
                .gm-schedule-btn {
                    border: none; padding: 5px 12px; border-radius: 4px; cursor: pointer;
                    font-size: 12px; transition: opacity 0.2s; color: white;
                }
                .gm-btn-close { background: #f56c6c; }
                .gm-btn-go { background: #409EFF; }
                .gm-schedule-btn:hover { opacity: 0.8; }
                #gm-sch-pre-sync:hover { color: #66b1ff; }
                @keyframes gmSlideUp { from { transform: translateY(30px); opacity: 0; } to { transform: translateY(0); opacity: 1; } }
            `;
                document.head.appendChild(style);
            }

            // --- 3. 构建容器 HTML ---
            const div = document.createElement('div');
            div.className = 'gm-schedule-box';
            div.innerHTML = `
            <div style="font-weight:bold; font-size:14px; margin-bottom:8px; color:#303133; text-align:center;">
                选课时间安排表
                <span style="font-size:11.5px; font-weight:600; color:var(--gm-primary-text); background:var(--gm-primary-light); border:1px solid var(--gm-primary-border); padding:1px 6px; border-radius:6px; margin-left:6px; vertical-align:1px;">${identityName}</span>
            </div>
            <div class="gm-status-bar">${statusHtml}</div>
            <table class="gm-sch-table">
                <thead>
                    <tr><th>选课阶段</th><th>选课轮次</th><th>时间安排</th><th>选课方式</th><th>课程范围</th></tr>
                </thead>
                <tbody>
                    ${tableRowsHtml}
                </tbody>
            </table>
            <div class="gm-schedule-footer">
                <label style="cursor:pointer; display:flex; align-items:center; user-select:none;">
                    <input type="checkbox" id="gm-schedule-check" style="margin-right:6px;">
                    不再显示此安排
                </label>
                <div class="gm-sch-btn-group">
                    <button class="gm-schedule-btn gm-btn-go" id="gm-schedule-go-btn">进入选课</button>
                    <button class="gm-schedule-btn gm-btn-close" id="gm-schedule-close-btn">关闭</button>
                </div>
            </div>
        `;
            document.body.appendChild(div);

            // --- 4. 事件绑定 ---
            // 绑定同步数据的点击事件
            if (showPreSyncLink) {
                const syncLink = document.getElementById('gm-sch-pre-sync');
                if (syncLink) {
                    syncLink.onclick = () => {
                        // 调用全局定义的同步函数
                        if (typeof handleSyncCourseClick === 'function') {
                            handleSyncCourseClick();
                        } else {
                            alert("同步功能初始化中，请稍后再试。");
                        }
                    };
                }
            }

            // 跳转选课页面
            document.getElementById('gm-schedule-go-btn').onclick = () => {
                window.location.href = SCHEDULE_CONFIG.COURSE_URL;
            };

            // 关闭
            document.getElementById('gm-schedule-close-btn').onclick = () => {
                if (document.getElementById('gm-schedule-check').checked) {
                    GM_setValue(SCHEDULE_CONFIG.STORAGE_KEY, true);
                }
                window.gm_schedule_manually_closed = true;
                div.remove();
            };
        };

        const hideWidget = () => {
            const box = document.querySelector('.gm-schedule-box');
            if (box) box.remove();
        };

        // 监控页面变化
        setInterval(() => {
            const iframes = document.querySelectorAll('iframe');
            let hasActiveSubPage = false;
            for (let f of iframes) {
                // 忽略插件自己创建的 iframe
                if (f.id && (f.id.startsWith('gm_') || f.style.visibility === 'hidden')) continue;
                // 检测是否有可见的大型iframe覆盖
                if (f.offsetParent !== null && f.offsetHeight > 300 && f.offsetWidth > 300) {
                    hasActiveSubPage = true;
                    break;
                }
            }
            if (window.location.href.includes('/student/home') && !hasActiveSubPage) {
                showWidget();
            } else {
                hideWidget();
            }
        }, 1000); // 稍微放宽检查间隔

        // 首次立即检查
        if (window.location.href.includes('/student/home')) showWidget();
    }

    // =-=-=-=-=-=-=-=-=-=-=-=-= 2.11 自动评教模块 =-=-=-=-=-=-=-=-=-=-=-=-=
    function initEvaluationHelper() {
        if (window.gm_eval_observer_started) return;
        window.gm_eval_observer_started = true;

        // --- 基础工具 ---
        const waitForElement = (selector, timeout = 5000) => {
            return new Promise((resolve) => {
                if (document.querySelector(selector)) return resolve(document.querySelector(selector));
                const observer = new MutationObserver(() => {
                    if (document.querySelector(selector)) {
                        observer.disconnect();
                        resolve(document.querySelector(selector));
                    }
                });
                observer.observe(document.body, { childList: true, subtree: true });
                setTimeout(() => { observer.disconnect(); resolve(null); }, timeout);
            });
        };
        const sleep = (ms) => new Promise(r => setTimeout(r, ms));

        // 模拟输入事件，确保Vue响应
        const triggerInputEvent = (element, value) => {
            if (!element) return;
            element.focus();
            element.value = value;
            element.dispatchEvent(new Event('input', { bubbles: true }));
            element.dispatchEvent(new Event('change', { bubbles: true }));
            element.blur();
        };

        // --- 1. 注入 CSS ---
        if (!document.getElementById('gm-eval-style')) {
            const style = getOrCreateStyleEl('gm-eval-style');
            style.textContent = `
            .gm-modal-overlay { position: fixed; top: 0; left: 0; width: 100%; height: 100%; background: rgba(0,0,0,0.5); z-index: 20000; display: flex; align-items: center; justify-content: center; backdrop-filter: blur(4px); }
            .gm-modal-content { background: #fff; border-radius: 12px; width: 720px; max-height: 85vh; display: flex; flex-direction: column; box-shadow: 0 12px 40px rgba(0,0,0,0.25); font-family: "Helvetica Neue", Helvetica, "PingFang SC", "Microsoft YaHei", Arial, sans-serif; animation: gmFadeIn 0.25s ease-out; border: 1px solid #ebeef5; }
            .gm-modal-header { padding: 18px 24px; border-bottom: 1px solid #eee; display: flex; justify-content: space-between; align-items: center; background: #fff; border-radius: 12px 12px 0 0; }
            .gm-modal-title { font-size: 18px; font-weight: 700; color: #303133; letter-spacing: 0.5px; }
            .gm-close-btn { width: 30px; height: 30px; border-radius: 50%; border: none; background: transparent; display: flex; align-items: center; justify-content: center; cursor: pointer; color: #909399; transition: all 0.2s; }
            .gm-close-btn:hover { background-color: #f56c6c; color: #fff; transform: rotate(90deg); }
            .gm-close-btn svg { width: 16px; height: 16px; fill: currentColor; }
            .gm-eval-body { padding: 20px; overflow-y: auto; flex: 1; background: #f5f7fa; }
            .gm-course-group { background: #fff; border: 1px solid #ebeef5; border-radius: 8px; margin-bottom: 15px; box-shadow: 0 2px 8px rgba(0,0,0,0.03); overflow: hidden; }
            .gm-course-header { background: #eef5fe; padding: 10px 15px; border-bottom: 1px solid #ebeef5; font-weight: bold; color: #409EFF; font-size: 14px; display: flex; align-items: center; gap: 8px; justify-content: space-between;}
            .gm-course-status-tag { font-size: 12px; font-weight: normal; padding: 2px 8px; border-radius: 10px; }
            .gm-tag-done { background: #f0f9eb; color: #67C23A; border: 1px solid #e1f3d8; }
            .gm-tag-todo { background: #fdf6ec; color: #E6A23C; border: 1px solid #faecd8; }
            .gm-teacher-row { display: flex; align-items: center; padding: 12px 15px; border-bottom: 1px solid #f2f2f2; transition: background 0.2s; }
            .gm-teacher-row:last-child { border-bottom: none; }
            .gm-teacher-row:hover { background: #fafafa; }
            .gm-teacher-row.gm-row-done { background: #fcfcfc; color: #999; }
            .gm-t-name { flex: 1; font-size: 14px; color: #606266; margin-left: 10px; font-weight: 500; display: flex; align-items: center; gap: 5px; }
            .gm-row-done .gm-t-name { color: #a8abb2; text-decoration: line-through; }
            .gm-done-badge { font-size: 12px; color: #67C23A; border: 1px solid #67C23A; padding: 0 4px; border-radius: 3px; transform: scale(0.9); text-decoration: none; display: inline-block;}
            .gm-score-input { width: 80px; padding: 6px 8px; border: 1px solid #dcdfe6; border-radius: 4px; text-align: center; font-family: Consolas, monospace; transition: 0.2s; margin-right: 15px; }
            .gm-score-input:focus { border-color: #409EFF; outline: none; box-shadow: 0 0 0 2px rgba(64,158,255,0.2); }
            .gm-score-input:disabled { background: #f5f7fa; color: #c0c4cc; cursor: not-allowed; border-color: #e4e7ed; }
            .gm-checkbox { cursor: pointer; width: 16px; height: 16px; accent-color: #409EFF; }
            .gm-checkbox:disabled { cursor: not-allowed; opacity: 0.5; }
            .gm-status-box { width: 70px; text-align: right; font-size: 12px; }
            .gm-modal-footer { padding: 16px 24px; border-top: 1px solid #eee; background: #fff; border-radius: 0 0 12px 12px; display: flex; justify-content: space-between; align-items: center; gap: 15px; }
            .gm-btn { padding: 9px 20px; border-radius: 6px; border: none; font-size: 14px; cursor: pointer; font-weight: 500; transition: 0.2s; display: inline-flex; align-items: center; gap: 6px; }
            .gm-btn-primary { background: #409EFF; color: white; }
            .gm-btn-primary:hover { background: #66b1ff; }
            .gm-btn-warning { background: #E6A23C; color: white; }
            .gm-btn-warning:hover { background: #ebb563; }
            .gm-btn:disabled { opacity: 0.6; cursor: not-allowed; background: #e4e7ed; color: #909399; }
            .gm-status-pending { color: #909399; }
            .gm-status-running { color: #409EFF; font-weight: bold; }
            .gm-status-success { color: #67C23A; font-weight: bold; }
            .gm-status-error { color: #F56C6C; }
            @keyframes gmFadeIn { from { opacity: 0; transform: scale(0.98); } to { opacity: 1; transform: scale(1); } }
        `;
            document.head.appendChild(style);
        }

        // --- 2. 抓取任务 ---
        function scrapeTasks() {
            const tasks = [];
            let idCounter = 0;
            const rows = document.querySelectorAll('.el-table__body-wrapper tbody tr');

            rows.forEach(row => {
                const courseNameEl = row.querySelector('.coursename .name') || row.querySelector('td:nth-child(2)');
                const courseName = courseNameEl ? courseNameEl.innerText.replace(/\s+/g, ' ').trim() : '未知课程';
                const successTag = row.querySelector('.el-tag--success');
                const isRowComplete = successTag && successTag.innerText.includes('已完成');

                const links = row.querySelectorAll('a');

                links.forEach(link => {
                    const isSubmitted = link.innerText.includes('已评') || link.classList.contains('submitted');
                    const isDisabled = link.classList.contains('is-disabled');

                    if (link.innerText.length > 1 && !isDisabled) {
                        tasks.push({
                            id: ++idCounter,
                            course: courseName,
                            teacher: link.innerText.trim(),
                            element: link,
                            isDone: isSubmitted,
                            courseIsDone: isRowComplete,
                            status: isSubmitted ? 'done' : 'pending'
                        });
                    }
                });
            });
            return tasks;
        }

        // --- 3. 显示主面板 ---
        const showEvalModal = () => {
            if (document.getElementById('gm-eval-modal')) return;
            const taskList = scrapeTasks();
            const courseGroups = {};
            taskList.forEach(task => {
                if (!courseGroups[task.course]) courseGroups[task.course] = [];
                courseGroups[task.course].push(task);
            });

            const overlay = document.createElement('div');
            overlay.id = 'gm-eval-modal';
            overlay.className = 'gm-modal-overlay';
            const pendingCount = taskList.filter(t => !t.isDone).length;

            overlay.innerHTML = `
            <div class="gm-modal-content">
                <div class="gm-modal-header">
                    <div class="gm-modal-title">自动评教功能 <span style="font-size:12px;font-weight:normal;color:#999;margin-left:10px;">待评任务: ${pendingCount}</span></div>
                    <button class="gm-close-btn" id="gm-eval-close" title="关闭">
                        <svg viewBox="0 0 1024 1024"><path d="M512 64C264.6 64 64 264.6 64 512s200.6 448 448 448 448-200.6 448-448S759.4 64 512 64zm165.4 618.2l-66-.3L512 563.4l-99.3 118.4-66.1.3c-4.4 0-8-3.5-8-8 0-1.9.7-3.7 1.9-5.2l130.1-155L340.5 359a8.32 8.32 0 0 1-1.9-5.2c0-4.4 3.6-8 8-8l66.1.3L512 464.6l99.3-118.4 66-.3c4.4 0 8 3.5 8 8 0 1.9-.7 3.7-1.9 5.2L553.5 514l130 155c1.2 1.5 1.9 3.3 1.9 5.2 0 4.4-3.6 8-8 8z"></path></svg>
                    </button>
                </div>
                <div class="gm-eval-body" id="gm-eval-container">
                    <div style="margin-bottom:10px;display:flex;justify-content:flex-end;">
                        <label style="font-size:13px;color:#606266;cursor:pointer;display:flex;align-items:center;">
                            <input type="checkbox" id="gm-check-all-available" style="margin-right:5px;"> 全选所有待评任务
                        </label>
                    </div>
                </div>
                <div class="gm-modal-footer">
                    <div style="flex:1;"></div>
                    <div style="display:flex; gap:10px;">
                        <button id="gm-btn-min-eval" class="gm-btn gm-btn-warning" title="跳过已完成课程，未完成课程只评第一个">
                            <svg viewBox="0 0 24 24" width="14" height="14" style="vertical-align:-2px; margin-right:4px;" stroke="currentColor" stroke-width="2" fill="none" stroke-linecap="round" stroke-linejoin="round"><polygon points="13 2 3 14 12 14 11 22 21 10 12 10 13 2"></polygon></svg>自动完成最低评教
                        </button>
                        <button id="gm-btn-run-selected" class="gm-btn gm-btn-primary">
                            <svg viewBox="0 0 24 24" width="14" height="14" style="vertical-align:-2px; margin-right:4px;" fill="currentColor" stroke="none"><path d="M8 5v14l11-7z"></path></svg>开始评教
                        </button>
                    </div>
                </div>
            </div>
        `;
            document.body.appendChild(overlay);

            const container = document.getElementById('gm-eval-container');
            if (Object.keys(courseGroups).length === 0) {
                container.innerHTML = '<div style="text-align:center;padding:50px;color:#999;">当前没有评教任务</div>';
                document.getElementById('gm-btn-min-eval').disabled = true;
                document.getElementById('gm-btn-run-selected').disabled = true;
            } else {
                for (const [courseName, teachers] of Object.entries(courseGroups)) {
                    const hasDone = teachers.some(t => t.isDone) || (teachers.length > 0 && teachers[0].courseIsDone);
                    const statusTag = hasDone
                        ? `<span class="gm-course-status-tag gm-tag-done">最低要求已达成</span>`
                        : `<span class="gm-course-status-tag gm-tag-todo">未完成</span>`;
                    const groupDiv = document.createElement('div');
                    groupDiv.className = 'gm-course-group';
                    let teachersHtml = '';
                    teachers.forEach(t => {
                        const rowClass = t.isDone ? 'gm-teacher-row gm-row-done' : 'gm-teacher-row';
                        const nameBadge = t.isDone ? '<span class="gm-done-badge">已完成</span>' : '';
                        const statusText = t.isDone ? '<span class="gm-status-success">已提交</span>' : '<span class="gm-status-pending">待评</span>';
                        const disabledAttr = t.isDone ? 'disabled' : '';
                        const inputPlaceholder = t.isDone ? '-' : '分数';
                        teachersHtml += `
                        <div class="${rowClass}">
                            <input type="checkbox" class="gm-item-check gm-checkbox" data-id="${t.id}" ${disabledAttr}>
                            <div class="gm-t-name">${t.teacher} ${nameBadge}</div>
                            <input type="number" class="gm-score-input" data-id="${t.id}" id="score-${t.id}" placeholder="${inputPlaceholder}" min="0" max="100" ${disabledAttr}>
                            <div class="gm-status-box"><span id="status-${t.id}">${statusText}</span></div>
                        </div>
                    `;
                    });
                    groupDiv.innerHTML = `
                    <div class="gm-course-header">
                        <span>${courseName}</span>
                        ${statusTag}
                    </div>
                    <div class="gm-teacher-list">
                        ${teachersHtml}
                    </div>
                `;
                    container.appendChild(groupDiv);
                }
            }

            const btnMin = document.getElementById('gm-btn-min-eval');
            const btnRun = document.getElementById('gm-btn-run-selected');
            const checkAll = document.getElementById('gm-check-all-available');

            document.getElementById('gm-eval-close').onclick = () => overlay.remove();

            checkAll.onchange = (e) => {
                document.querySelectorAll('.gm-item-check:not(:disabled)').forEach(cb => cb.checked = e.target.checked);
            };
            document.querySelectorAll('.gm-score-input:not(:disabled)').forEach(input => {
                input.oninput = function () {
                    const id = this.getAttribute('data-id');
                    const cb = document.querySelector(`.gm-item-check[data-id="${id}"]`);
                    if (cb) cb.checked = true;
                };
            });

            const fillFormExact = (targetScore) => {
                const groups = document.querySelectorAll('.el-radio-group');
                const questions = [];
                let maxTotalScore = 0;

                // 1. 扫描题目结构
                groups.forEach((group, index) => {
                    const options = group.querySelectorAll('.el-radio');
                    if (options.length === 0) return;

                    const text = options[0].innerText || "";
                    let maxPoints = 5;
                    let step = 1;

                    if (text.includes("10分")) {
                        maxPoints = 10;
                        step = 2;
                    }

                    maxTotalScore += maxPoints;
                    questions.push({
                        domOptions: options,
                        maxPoints: maxPoints,
                        step: step,
                        currentIdx: 0
                    });
                });

                if (targetScore > maxTotalScore) targetScore = maxTotalScore;
                if (targetScore < 0) targetScore = 0;

                let pointsToLose = maxTotalScore - targetScore;

                // 2. 算法扣分
                // Phase A: 扣除奇数分 (找5分题)
                if (pointsToLose % 2 !== 0) {
                    const q5 = questions.find(q => q.step === 1);
                    if (q5) {
                        q5.currentIdx = 1;
                        pointsToLose -= 1;
                    }
                }

                // Phase B: 扣除偶数分 (优先10分题)
                for (let q of questions) {
                    if (pointsToLose <= 0) break;
                    const remainingSteps = (q.domOptions.length - 1) - q.currentIdx;
                    const maxDeductable = remainingSteps * q.step;

                    if (maxDeductable > 0) {
                        let deduct = Math.min(pointsToLose, maxDeductable);
                        const stepsToMove = deduct / q.step;
                        q.currentIdx += stepsToMove;
                        pointsToLose -= deduct;
                    }
                }

                // 3.深度点击执行
                questions.forEach(q => {
                    const targetOption = q.domOptions[q.currentIdx] || q.domOptions[q.domOptions.length - 1];
                    if (targetOption) {
                        // 尝试找到内部真正的 input 元素
                        const internalInput = targetOption.querySelector('input.el-radio__original');
                        if (internalInput) {
                            internalInput.click(); // 原生点击
                            // 双重保险：手动派发变更事件，确保 Vue Model 更新
                            internalInput.checked = true;
                            internalInput.dispatchEvent(new Event('change', { bubbles: true }));
                        } else {
                            // 降级：点击 Label
                            targetOption.click();
                        }
                    }
                });

                // 4. 填星星
                document.querySelectorAll('.el-rate').forEach(group => {
                    const stars = group.querySelectorAll('.el-rate__item');
                    let starIdx = stars.length - 1;
                    if (targetScore < 90) starIdx = Math.max(0, stars.length - 2);
                    if (stars[starIdx]) stars[starIdx].click();
                });

                // 5. 填评语
                const comments = ["老师授课认真，重点突出。", "教学严谨，对学生负责。", "课堂氛围好，讲解生动。", "深入浅出，受益匪浅。", "理论联系实际，收获很大。"];
                document.querySelectorAll('textarea').forEach(area => {
                    const randomComment = comments[Math.floor(Math.random() * comments.length)];
                    triggerInputEvent(area, randomComment);
                });
            };

            // --- 核心执行函数 ---
            const executeTasks = async (tasksToRun) => {
                if (tasksToRun.length === 0) {
                    alert("没有选中任何任务！");
                    return;
                }

                btnMin.disabled = true;
                btnRun.disabled = true;
                document.querySelectorAll('input').forEach(i => i.disabled = true);
                let downgradedCourses = [];

                for (let i = 0; i < tasksToRun.length; i++) {
                    const task = tasksToRun[i];
                    if (task.isDone) continue;

                    const statusEl = document.getElementById(`status-${task.id}`);
                    const inputVal = document.getElementById(`score-${task.id}`).value;
                    let scoreVal = inputVal ? parseInt(inputVal) : 95;

                    statusEl.className = 'gm-status-running';
                    statusEl.innerText = '准备进入...';
                    const rowEl = statusEl.closest('.gm-teacher-row');
                    if (rowEl) rowEl.scrollIntoView({ behavior: 'smooth', block: 'center' });

                    try {
                        // ★ 重新寻找DOM (Fix Stale Element)
                        let activeLink = null;
                        const allRows = document.querySelectorAll('.el-table__body-wrapper tbody tr');
                        for (let tr of allRows) {
                            const courseText = tr.innerText;
                            if (courseText.includes(task.course) && courseText.includes(task.teacher)) {
                                const links = tr.querySelectorAll('a');
                                for (let link of links) {
                                    if (link.innerText.includes(task.teacher) && !link.classList.contains('is-disabled')) {
                                        activeLink = link;
                                        break;
                                    }
                                }
                            }
                            if (activeLink) break;
                        }
                        if (!activeLink) activeLink = task.element;

                        activeLink.click();

                        statusEl.innerText = '加载表单...';
                        const formReady = await waitForElement('.el-radio-group', 15000);
                        if (!formReady) {
                            // 重试点击
                            activeLink.click();
                            const retryReady = await waitForElement('.el-radio-group', 10000);
                            if (!retryReady) throw new Error("表单加载超时");
                        }
                        await sleep(1000);

                        // 1. 填表
                        statusEl.innerText = '正在填表...';
                        fillFormExact(scoreVal);
                        await sleep(1500);

                        // 2. 提交
                        let submitBtn = null;
                        const btnGroup = document.getElementById('btn-group');
                        if (btnGroup) {
                            const btns = btnGroup.querySelectorAll('button');
                            for (let btn of btns) {
                                if (btn.textContent.includes('提交') && !btn.textContent.includes('匿名')) {
                                    submitBtn = btn;
                                    break;
                                }
                            }
                        }

                        if (submitBtn) {
                            // 如果按钮还禁用，重试填表
                            if (submitBtn.disabled || submitBtn.classList.contains('is-disabled')) {
                                fillFormExact(scoreVal);
                                await sleep(1000);
                            }

                            statusEl.innerText = '提交中...';
                            submitBtn.click();

                            const msgBox = await waitForElement('.el-message-box', 5000);

                            // 检查是否有错误提示 (500 Error 会弹 toast 或 message-box)
                            const errorToast = document.querySelector('.el-message--error');
                            if (errorToast) {
                                throw new Error("服务器返回错误(500)，可能是提交过快");
                            }

                            if (msgBox) {
                                const text = msgBox.innerText || "";
                                const confirmBtn = msgBox.querySelector('.el-button--primary');

                                // 场景 A：20% 限制
                                if (text.includes('20%') || text.includes('不得超过') || text.includes('优秀')) {
                                    statusEl.innerText = '限制触发, 降分...';
                                    if (confirmBtn) confirmBtn.click();
                                    await sleep(1000);

                                    scoreVal = 89;
                                    downgradedCourses.push(`${task.course}`);
                                    fillFormExact(89);
                                    await sleep(1500);

                                    if (!submitBtn.disabled) {
                                        submitBtn.click();
                                        const confirmBox2 = await waitForElement('.el-message-box__btns', 5000);
                                        if (confirmBox2) {
                                            const finalOk = confirmBox2.querySelector('.el-button--primary');
                                            if (finalOk) finalOk.click();
                                        }
                                    }
                                }
                                // 场景 B：普通确认
                                else {
                                    if (confirmBtn) confirmBtn.click();
                                }
                            }

                            // 等待返回列表
                            statusEl.innerText = '等待返回...';
                            await waitForElement('.el-table__body-wrapper', 15000);
                            await sleep(1500);

                            statusEl.className = 'gm-status-success';
                            statusEl.innerText = `完成(${scoreVal})`;
                        } else {
                            throw new Error("未找到提交按钮");
                        }

                    } catch (e) {
                        Logger.error('2.11', '评教任务执行异常', e);
                        statusEl.className = 'gm-status-error';
                        statusEl.innerText = '失败';
                        const backBtn = Array.from(document.querySelectorAll('button')).find(b => b.innerText.includes('取消') || b.innerText.includes('返回'));
                        if (backBtn) {
                            backBtn.click();
                            await waitForElement('.el-table__body-wrapper', 5000);
                        }
                        await sleep(2000);
                    }
                }

                btnMin.innerText = "流程结束";
                btnRun.innerText = "流程结束";

                let finishMsg = "所有任务处理完成！";
                if (downgradedCourses.length > 0) {
                    finishMsg += `\n\n检测到优秀率限制，以下课程已自动降为 89 分：\n` + downgradedCourses.join('\n');
                }
                finishMsg += "\n\n建议刷新页面更新状态。是否刷新？";

                GmDialog.confirm({ title: '评教流程完成', message: finishMsg, type: downgradedCourses.length > 0 ? 'warn' : 'success', confirmText: '刷新页面' }).then(ok => {
                    if (ok) window.location.reload();
                });
            };

            // --- 功能 A: 自动完成最低评教 ---
            btnMin.onclick = () => {
                document.querySelectorAll('.gm-item-check').forEach(c => c.checked = false);
                document.querySelectorAll('.gm-score-input').forEach(i => {
                    if (!i.disabled) i.value = '';
                });

                const itemsToRun = [];
                let skippedCourses = 0;

                for (const [courseName, teachers] of Object.entries(courseGroups)) {
                    const alreadyDone = teachers.some(t => t.isDone) || (teachers.length > 0 && teachers[0].courseIsDone);
                    if (alreadyDone) {
                        skippedCourses++;
                        continue;
                    }
                    if (teachers.length > 0) {
                        const target = teachers[0];
                        if (target.isDone) continue;

                        const checkbox = document.querySelector(`.gm-item-check[data-id="${target.id}"]`);
                        const scoreInput = document.getElementById(`score-${target.id}`);

                        if (checkbox && scoreInput && !checkbox.disabled) {
                            checkbox.checked = true;
                            // 随机 80 - 89 分
                            scoreInput.value = Math.floor(Math.random() * 10) + 80;
                            itemsToRun.push(target);
                        }
                    }
                }

                if (itemsToRun.length === 0) {
                    alert(`没有待处理的最低评教任务。\n\n已跳过 ${skippedCourses} 门已完成(或部分完成)的课程。`);
                    return;
                }

                if (confirm(`即将对 ${itemsToRun.length} 门课程进行最低标准评教（每门课评1人，随机80-89分）。\n\n是否开始？`)) {
                    executeTasks(itemsToRun);
                }
            };

            // --- 功能 B: 开始评教---
            btnRun.onclick = () => {
                const selectedIds = Array.from(document.querySelectorAll('.gm-item-check:checked'))
                    .filter(cb => !cb.disabled)
                    .map(cb => parseInt(cb.dataset.id));

                const itemsToRun = taskList.filter(t => selectedIds.includes(t.id));

                if (itemsToRun.length === 0) {
                    alert("请至少勾选一个待评任务！");
                    return;
                }

                let hasEmptyScore = false;
                itemsToRun.forEach(t => {
                    const val = document.getElementById(`score-${t.id}`).value;
                    if (!val) hasEmptyScore = true;
                });

                let msg = `即将对 ${itemsToRun.length} 位教师进行评教。`;
                if (hasEmptyScore) msg += `\n\n注意：部分未填分，默认按 95分 (优秀) 处理。`;
                msg += `\n\n是否开始？`;

                GmDialog.confirm({ title: '开始评教', message: msg, type: hasEmptyScore ? 'warn' : 'info', confirmText: '开始' }).then(ok => {
                    if (ok) executeTasks(itemsToRun);
                });
            };
        };

        // --- 4. 入口按钮 ---
        const injectPageButton = () => {
            const targetContainer = document.querySelector('.el-tab-pane .el-select') || document.querySelector('.el-form');
            if (!targetContainer || document.getElementById('gm-page-eval-btn')) return;

            const btn = document.createElement('button');
            btn.id = 'gm-page-eval-btn';
            btn.className = 'el-button el-button--primary el-button--small';
            btn.innerHTML = `<i class="el-icon-s-cooperation"></i> 打开自动评教`;
            btn.style.cssText = 'margin-left: 15px; vertical-align: top; height: 32px; font-weight: bold; box-shadow: 0 2px 6px rgba(64,158,255, 0.3);';

            if (targetContainer.parentNode) targetContainer.parentNode.insertBefore(btn, targetContainer.nextSibling);
            else targetContainer.appendChild(btn);

            btn.onclick = showEvalModal;
        };

        const startObserve = () => {
            const observer = new MutationObserver(() => {
                if (location.href.includes('evaluation-student-frontend')) injectPageButton();
            });
            observer.observe(document.body, { childList: true, subtree: true });
            injectPageButton();
        };

        if (document.body) startObserve();
        else window.addEventListener('load', startObserve);
    }

    // =-=-=-=-=-=-=-=-=-=-=-=-= 2.12 人员信息检索 =-=-=-=-=-=-=-=-=-=-=-=-=
    const PersonnelSearch = {

        STORAGE_KEY: "nwpu_synced_token",
        API_BASE: "https://electronic-signature.nwpu.edu.cn/api/local-user/page",
        state: { page: 1, loading: false, hasMore: true, keyword: "", items: [] },

        // 名片动态渐变色系库
        THEME_COLORS: [
            { bg: 'linear-gradient(135deg, #4facfe 0%, #00f2fe 100%)', text: '#0084ff' }, // 经典蓝
            { bg: 'linear-gradient(135deg, #43e97b 0%, #38f9d7 100%)', text: '#21b25b' }, // 清新绿
            { bg: 'linear-gradient(135deg, #fa709a 0%, #fee140 100%)', text: '#e64e7c' }, // 晚霞粉
            { bg: 'linear-gradient(135deg, #a18cd1 0%, #fbc2eb 100%)', text: '#8c6ad1' }, // 梦幻紫
            { bg: 'linear-gradient(135deg, #f6d365 0%, #fda085 100%)', text: '#e8753a' }  // 活力橙
        ],

        // 中国大陆及港澳台省级行政区划代码字典 (前2位)
        PROVINCE_MAP: {
            11: "北京", 12: "天津", 13: "河北", 14: "山西", 15: "内蒙古",
            21: "辽宁", 22: "吉林", 23: "黑龙江", 31: "上海", 32: "江苏",
            33: "浙江", 34: "安徽", 35: "福建", 36: "江西", 37: "山东",
            41: "河南", 42: "湖北", 43: "湖南", 44: "广东", 45: "广西",
            46: "海南", 50: "重庆", 51: "四川", 52: "贵州", 53: "云南",
            54: "西藏", 61: "陕西", 62: "甘肃", 63: "青海", 64: "宁夏",
            65: "新疆", 71: "台湾", 81: "香港", 82: "澳门", 83: "台湾"
        },

        _inferIdentity(id) {
            if (!id) return { short: '', long: '' };

            // 规则1：检查是否为 10 位字符（前4位数字年份，5-6位类型码允许字母，后4位数字编号）
            if (id.length === 10 && /^\d{4}[A-Z0-9]{2}\d{4}$/i.test(id)) {
                const year = id.substring(0, 4);
                const typeCode = id.substring(4, 6).toUpperCase();

                let shortLabel = '';
                let longLabel = '';
                let yearSuffix = '级';

                switch (typeCode) {
                    case '10': shortLabel = '博士生'; longLabel = '学术型 博士研究生'; break;
                    case '11': shortLabel = '博士生'; longLabel = '专项计划 博士研究生'; break;
                    case '12': shortLabel = '博士'; longLabel = '同等学力申请博士学位'; yearSuffix = '年申请'; break;
                    case '16': shortLabel = '博士生'; longLabel = '专业型 博士研究生'; break;
                    case '18': shortLabel = '留学生'; longLabel = '国际留学生 (博士)'; break;
                    case '20': shortLabel = '学硕'; longLabel = '全日制 学术型 硕士研究生'; break;
                    case '21': shortLabel = '专硕'; longLabel = '非全日制 专业型 硕士研究生'; break;
                    case '22': shortLabel = '硕士'; longLabel = '同等学力申请硕士学位'; yearSuffix = '年申请'; break;
                    case '24': shortLabel = ''; longLabel = ''; break;
                    case '25': shortLabel = ''; longLabel = ''; break;
                    case '26': shortLabel = '专硕'; longLabel = '全日制 专业型 硕士研究生'; break;
                    case '28': shortLabel = '留学生'; longLabel = '国际留学生 (硕士)'; break;
                    case '30': shortLabel = '本科生'; longLabel = '本科生'; break;
                    case '32': shortLabel = '第二学位'; longLabel = '第二学士学位本科生'; break;
                    case '37': shortLabel = '交换生'; longLabel = '本科交换生'; break;
                    case '38': shortLabel = '留学生'; longLabel = '国际留学生 (本科)'; break;
                    case '70': shortLabel = '预科生'; longLabel = '预科生'; break;
                    case '00':
                        shortLabel = '职工';
                        longLabel = '职工';
                        yearSuffix = '年入职';
                        break;
                    case '01':
                        shortLabel = '职工';
                        longLabel = '职工';
                        yearSuffix = '年入职';
                        break;
                    case '02':
                        shortLabel = '职工';
                        longLabel = '职工';
                        yearSuffix = '年入职';
                        break;
                    case '03':
                        shortLabel = '职工';
                        longLabel = '职工';
                        yearSuffix = '年入职';
                        break;
                    case '05':
                        shortLabel = '职工';
                        longLabel = '职工';
                        yearSuffix = '年入职';
                        break;
                    case '06':
                        shortLabel = '职工';
                        longLabel = '职工';
                        yearSuffix = '年入职';
                        break;
                    case '07':
                        shortLabel = '职工';
                        longLabel = '职工';
                        yearSuffix = '年入职';
                        break;
                    case '08':
                        shortLabel = '职工';
                        longLabel = '职工';
                        yearSuffix = '年入职';
                        break;
                    case '0P':
                        shortLabel = '职工';
                        longLabel = '职工';
                        yearSuffix = '年入职';
                        break;
                    case '1K':
                        shortLabel = '博士';
                        longLabel = '翱翔快响专项计划录取博士生';
                        break;
                    case '1P':
                        shortLabel = '职工';
                        longLabel = '后勤部职工';
                        yearSuffix = '年入职';
                        break;
                    default:
                        return { short: '', long: '' };
                }

                return {
                    short: shortLabel,
                    long: `${year}${yearSuffix} ${longLabel}`
                };
            }

            // 规则2：检查是否为 8 位纯数字（退休）
            if (id.length === 8 && /^\d{8}$/.test(id)) {
                return {
                    short: '离/退休',
                };
            }

            // 规则3：检查是否为 6 位纯数字（早期学号）
            if (id.length === 6 && /^\d{6}$/.test(id)) {
                const shortYear = '20' + id.substring(0, 2);
                return {
                    short: '本科生',
                    long: `${shortYear}级 本科生`
                };
            }

            return { short: '', long: '' };
        },

        _parseIDCard(sfzjh) {
            if (!sfzjh) return null;

            // 1. 大陆身份证 及 港澳台居民居住证（18位脱敏，第17位包含性别）
            // (居住证前两位：81香港，82澳门，83台湾)
            if (sfzjh.length === 18 && /^[1-9]\d\*{14}\d[\dXx]$/i.test(sfzjh)) {
                const provCode = sfzjh.substring(0, 2);
                const genderNum = parseInt(sfzjh.charAt(16), 10);
                return {
                    type: '身份证/居住证',
                    province: this.PROVINCE_MAP[provCode] || '未知区域',
                    gender: (genderNum % 2 === 1) ? '男' : '女'
                };
            }

            // 2. 港澳居民来往内地通行证 (回乡证，9位字符脱敏)
            if (sfzjh.length === 9 && /^[HMhm]\d\*{5}\d{2}$/.test(sfzjh)) {
                const firstLetter = sfzjh.charAt(0).toUpperCase();
                return {
                    type: '港澳通行证',
                    province: firstLetter === 'H' ? '香港' : '澳门',
                    gender: '未知' // 通行证号码不包含性别信息
                };
            }

            // 3. 台湾居民来往大陆通行证 (台胞证，8位数字脱敏)
            if (sfzjh.length === 8 && /^\d{2}\*{4}\d{2}$/.test(sfzjh)) {
                return {
                    type: '台湾通行证',
                    province: '台湾',
                    gender: '未知' // 台胞证号码不包含性别信息
                };
            }
            return null;
        },

        syncToken() {
            if (location.host !== 'ecampus.nwpu.edu.cn') return;
            const checkAndSave = () => {
                const token = localStorage.getItem('token');
                if (token) {
                    GM_setValue(this.STORAGE_KEY, token);
                    syncIdentityFromToken(token);
                }
            };
            checkAndSave();
            setTimeout(checkAndSave, 500);
            setTimeout(checkAndSave, 2000);
        },

        openModal() {
            Logger.log('2.12', "初始化人员信息检索");
            const token = GM_getValue(this.STORAGE_KEY);
            if (token) {
                if (document.getElementById('gm-person-search-overlay')) return;
                this.injectStyles();
                this.createUI();
                this.resetState();
                return;
            }
            this._startSilentSync();
        },

        _startSilentSync() {
            this._showToast("正在后台获取授权，请稍候...");
            const iframe = document.createElement('iframe');
            iframe.src = 'https://ecampus.nwpu.edu.cn';
            iframe.style.display = 'none';
            iframe.id = 'gm-sync-iframe-worker';
            document.body.appendChild(iframe);

            let attempts = 0;
            const maxAttempts = 15;
            const timer = setInterval(() => {
                const newToken = GM_getValue(this.STORAGE_KEY);
                if (newToken) {
                    clearInterval(timer);
                    this._cleanupSync();
                    this._showToast("授权成功！正在打开界面...", 1000);
                    setTimeout(() => this.openModal(), 500);
                } else {
                    attempts++;
                    if (attempts >= maxAttempts) {
                        clearInterval(timer);
                        this._cleanupSync();
                        this._removeToast();
                        if (confirm("后台自动同步超时（可能是您未登录翱翔门户）。\n\n是否打开新窗口手动登录？")) {
                            window.open('https://ecampus.nwpu.edu.cn', '_blank');
                        }
                    }
                }
            }, 500);
        },

        _cleanupSync() {
            const frame = document.getElementById('gm-sync-iframe-worker');
            if (frame) frame.remove();
        },

        _showToast(msg, duration = 0) {
            let toast = document.getElementById('gm-search-toast');
            if (!toast) {
                toast = document.createElement('div');
                toast.id = 'gm-search-toast';
                toast.style.cssText = 'position:fixed; top:20%; left:50%; transform:translateX(-50%); background:rgba(0,0,0,0.75); color:white; padding:12px 24px; border-radius:30px; font-size:14px; z-index:100020; transition:opacity 0.3s; box-shadow:0 4px 15px rgba(0,0,0,0.2); pointer-events:none; font-family: "PingFang SC", "Microsoft YaHei", sans-serif;';
                document.body.appendChild(toast);
            }
            toast.innerText = msg;
            toast.style.opacity = '1';
            if (duration > 0) {
                setTimeout(() => {
                    toast.style.opacity = '0';
                    setTimeout(() => toast.remove(), 300);
                }, duration);
            }
        },

        _removeToast() {
            const toast = document.getElementById('gm-search-toast');
            if (toast) toast.remove();
        },

        injectStyles() {
            if (document.getElementById('gm-person-search-style')) return;
            const style = getOrCreateStyleEl('gm-person-search-style');
            style.textContent = `
            .gm-modal-overlay { position: fixed; top: 0; left: 0; width: 100%; height: 100%; background-color: rgba(0, 0, 0, 0.55); backdrop-filter: blur(4px); z-index: 10005; display: flex; align-items: center; justify-content: center; animation: gmFadeIn 0.2s ease-out;}
            .gm-modal-content { background-color: #fff; border-radius: 12px; width: 880px; max-width: 95%; height: 75vh; max-height: 850px; box-shadow: 0 16px 40px rgba(0, 0, 0, 0.2); display: flex; flex-direction: column; overflow: hidden; font-family: "PingFang SC", "Hiragino Sans GB", "Microsoft YaHei", "Helvetica Neue", Helvetica, Arial, sans-serif; }
            .gm-modal-header { padding: 18px 24px; border-bottom: 1px solid #ebeef5; background: #fff; display: flex; justify-content: space-between; align-items: center; flex-shrink: 0; }
            .gm-modal-title { font-size: 19px; font-weight: 600; color: #2c3e50; display: flex; align-items: center; gap: 10px; letter-spacing: 0.5px;}
            .gm-modal-close { border: none; background: transparent; font-size: 26px; color: #909399; cursor: pointer; padding: 0; line-height: 1; transition: color 0.2s; font-family: Arial, sans-serif;}
            .gm-modal-close:hover { color: #f56c6c; }

            .gm-ps-body { display: flex; flex-direction: row; height: 100%; overflow: hidden; padding: 0 !important; background: #fff;}
            .gm-ps-left { width: 340px; display: flex; flex-direction: column; border-right: 1px solid #ebeef5; background: #fafafa; flex-shrink: 0;}
            .gm-ps-right { flex: 1; display: flex; align-items: center; justify-content: center; background: #f5f7fa; position: relative; overflow: hidden;}

            .gm-ps-search-box { padding: 16px; background: #fff; border-bottom: 1px solid #ebeef5; box-shadow: 0 2px 10px rgba(0,0,0,0.02); z-index: 2;}
            .gm-ps-search-bar { display: flex; gap: 10px; }
            .gm-ps-input { flex: 1; padding: 10px 14px; border: 1px solid #dcdfe6; border-radius: 6px; outline: none; font-size: 14px; transition: all 0.2s; color: #303133; width: 100%; box-sizing: border-box;}
            .gm-ps-input:focus { border-color: #409EFF; box-shadow: 0 0 0 2px rgba(64,158,255,0.1);}
            .gm-ps-btn { padding: 0 18px; background: #409EFF; color: white; border: none; border-radius: 6px; cursor: pointer; font-size: 14px; font-weight: 500; transition: background 0.2s; white-space: nowrap;}
            .gm-ps-btn:hover { background: #66b1ff; }

            .gm-ps-list-container { flex: 1; overflow-y: auto; padding: 0; position: relative; }
            .gm-ps-list-container::-webkit-scrollbar { width: 6px; }
            .gm-ps-list-container::-webkit-scrollbar-thumb { background: #dcdfe6; border-radius: 3px; }

            .gm-ps-list-item { padding: 15px 20px; cursor: pointer; border-bottom: 1px solid #ebeef5; transition: all 0.2s; background: #fff; display: flex; flex-direction: column; gap: 8px; position: relative;}
            .gm-ps-list-item:hover { background: #f0f7ff; }
            .gm-ps-list-item.active { background: #ecf5ff; }
            .gm-ps-list-item.active::before { content: ''; position: absolute; left: 0; top: 0; bottom: 0; width: 4px; background: #409EFF; }
            .gm-ps-item-name { font-size: 16px; font-weight: 600; color: #2c3e50; display: flex; justify-content: space-between; align-items: center; letter-spacing: 0.5px;}
            .gm-ps-item-id { font-family: "Helvetica Neue", Helvetica, Arial, sans-serif; font-size: 13.5px; color: #8c939d; letter-spacing: 0.8px;}
            .gm-ps-loader { padding: 24px; text-align: center; color: #909399; font-size: 14px; }

            .gm-ps-card { background: #fff; width: 380px; border-radius: 12px; box-shadow: 0 16px 36px rgba(0,0,0,0.08); overflow: hidden; display: none; flex-direction: column; border: 1px solid #ebeef5;}
            .gm-ps-card.show { display: flex; animation: gmFadeInUp 0.35s cubic-bezier(0.18, 0.89, 0.32, 1.28); }
            .gm-ps-card-header { background: linear-gradient(135deg, #409EFF, #73bfff); height: 100px; position: relative; transition: background 0.5s ease;}
            .gm-ps-avatar { width: 80px; height: 80px; border-radius: 50%; background: #fff; border: 4px solid #fff; position: absolute; bottom: -40px; left: 28px; box-shadow: 0 6px 16px rgba(0,0,0,0.12); display: flex; align-items: center; justify-content: center; font-size: 34px; color: #409EFF; font-weight: bold; font-family: "PingFang SC", "Microsoft YaHei", sans-serif; transition: color 0.5s ease;}
            .gm-ps-card-body { padding: 56px 28px 32px 28px; }

            .gm-ps-info-row { display: flex; flex-direction: column; gap: 8px; margin-bottom: 20px; padding-bottom: 18px; border-bottom: 1px dashed #ebeef5; }
            .gm-ps-info-row:last-child { margin-bottom: 0; padding-bottom: 0; border-bottom: none; }

            /* 新增：两列网格布局，用于展示性别和籍贯等简短信息 */
            .gm-ps-info-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 16px; margin-bottom: 20px; padding-bottom: 18px; border-bottom: 1px dashed #ebeef5; }
            .gm-ps-info-col { display: flex; flex-direction: column; gap: 8px; }

            .gm-ps-info-label { font-size: 13px; color: #909399; font-weight: 500;}
            .gm-ps-info-value { font-size: 15px; color: #303133; font-weight: 500; letter-spacing: 0.5px;}

            @keyframes gmFadeIn { from { opacity: 0; transform: scale(0.98); } to { opacity: 1; transform: scale(1); } }
            @keyframes gmFadeInUp { from { opacity: 0; transform: translateY(20px); } to { opacity: 1; transform: translateY(0); } }
        `;
            document.head.appendChild(style);
        },

        createUI() {
            const overlay = document.createElement('div');
            overlay.id = 'gm-person-search-overlay';
            overlay.className = 'gm-modal-overlay';

            overlay.innerHTML = `
            <div class="gm-modal-content">
                <div class="gm-modal-header">
                    <div class="gm-modal-title">
                        <svg viewBox="0 0 24 24" width="22" height="22" stroke="#409EFF" stroke-width="2.5" fill="none" stroke-linecap="round" stroke-linejoin="round">
                            <circle cx="10" cy="8" r="5"></circle>
                            <path d="M16 21v-2a4 4 0 0 0-4-4H8a4 4 0 0 0-4 4v2"></path>
                            <line x1="21" y1="21" x2="15" y2="15"></line>
                        </svg>
                        人员信息检索
                    </div>
                    <button class="gm-modal-close" id="gm-ps-close">&times;</button>
                </div>
                <div class="gm-modal-body gm-ps-body">

                    <div class="gm-ps-left">
                        <div class="gm-ps-search-box">
                            <div class="gm-ps-search-bar">
                                <input type="text" id="gm-ps-input" class="gm-ps-input" placeholder="输入姓名、学号或工号">
                                <button id="gm-ps-btn" class="gm-ps-btn">检索</button>
                            </div>
                        </div>
                        <div class="gm-ps-list-container" id="gm-ps-scroll-area">
                            <div id="gm-ps-list"></div>
                            <div id="gm-ps-loader" class="gm-ps-loader">请输入关键词开始搜索</div>
                        </div>
                    </div>

                    <div class="gm-ps-right">
                        <div id="gm-ps-empty-tip" style="color:#a8abb2; font-size:14px; display:flex; flex-direction:column; align-items:center; gap:16px;">
                            <svg viewBox="0 0 24 24" width="56" height="56" stroke="#dcdfe6" stroke-width="1.5" fill="none" stroke-linecap="round" stroke-linejoin="round">
                                <rect x="3" y="3" width="18" height="18" rx="2" ry="2"></rect>
                                <circle cx="8.5" cy="8.5" r="1.5"></circle>
                                <polyline points="21 15 16 10 5 21"></polyline>
                            </svg>
                            <span>点击左侧列表查看名片详情</span>
                        </div>

                        <div id="gm-ps-detail-card" class="gm-ps-card">
                            <div class="gm-ps-card-header" id="gm-ps-card-header-bg">
                                <div class="gm-ps-avatar" id="gm-ps-card-avatar">A</div>
                            </div>
                            <div class="gm-ps-card-body">
                                <div id="gm-ps-card-name" style="font-size: 24px; font-weight: 600; color: #2c3e50; margin-bottom: 8px; letter-spacing: 1px;">-</div>
                                <div id="gm-ps-card-id" style="font-family: 'Helvetica Neue', Helvetica, Arial, sans-serif; color: #8c939d; font-size: 15px; margin-bottom: 28px; letter-spacing: 0.8px;">-</div>

                                <!-- 提取的籍贯和性别（双列网格） -->
                                <div class="gm-ps-info-grid" id="gm-ps-extra-info">
                                    <div class="gm-ps-info-col">
                                        <span class="gm-ps-info-label">性别</span>
                                        <span class="gm-ps-info-value" id="gm-ps-card-gender">-</span>
                                    </div>
                                    <div class="gm-ps-info-col">
                                        <span class="gm-ps-info-label">籍贯</span>
                                        <span class="gm-ps-info-value" id="gm-ps-card-prov">-</span>
                                    </div>
                                </div>

                                <div class="gm-ps-info-row">
                                    <span class="gm-ps-info-label">所在院系/部门</span>
                                    <span class="gm-ps-info-value" id="gm-ps-card-dept">-</span>
                                </div>
                                <div class="gm-ps-info-row" id="gm-ps-card-type-row">
                                    <span class="gm-ps-info-label">身份特征</span>
                                    <span class="gm-ps-info-value" id="gm-ps-card-type">-</span>
                                </div>
                            </div>
                        </div>
                    </div>

                </div>
            </div>
        `;
            document.body.appendChild(overlay);

            const closeFn = () => overlay.remove();
            document.getElementById('gm-ps-close').onclick = closeFn;
            overlay.onclick = (e) => { if (e.target === overlay) closeFn(); };

            const doSearch = () => {
                const val = document.getElementById('gm-ps-input').value.trim();
                if (val) {
                    this.state.keyword = val;
                    this.resetState();
                    this.fetchData();
                }
            };
            document.getElementById('gm-ps-btn').onclick = doSearch;
            document.getElementById('gm-ps-input').onkeypress = (e) => { if (e.key === 'Enter') doSearch(); };

            const scrollArea = document.getElementById('gm-ps-scroll-area');
            scrollArea.onscroll = () => {
                if (scrollArea.scrollTop + scrollArea.clientHeight >= scrollArea.scrollHeight - 30) {
                    if (!this.state.loading && this.state.hasMore) this.fetchData();
                }
            };

            document.getElementById('gm-ps-list').addEventListener('click', (e) => {
                const itemDiv = e.target.closest('.gm-ps-list-item');
                if (!itemDiv) return;

                document.querySelectorAll('.gm-ps-list-item').forEach(el => el.classList.remove('active'));
                itemDiv.classList.add('active');

                const idx = itemDiv.getAttribute('data-idx');
                const data = this.state.items[idx];
                if (data) this.showCardDetail(data);
            });
        },

        resetState() {
            this.state = { page: 1, loading: false, hasMore: true, keyword: this.state.keyword, items: [] };
            const listDiv = document.getElementById('gm-ps-list');
            if (listDiv) listDiv.innerHTML = '';
            const loader = document.getElementById('gm-ps-loader');
            if (loader) {
                loader.style.display = 'block';
                loader.innerText = this.state.keyword ? '正在检索数据库...' : '请输入关键词';
            }
            document.getElementById('gm-ps-empty-tip').style.display = 'flex';
            document.getElementById('gm-ps-detail-card').classList.remove('show');
        },

        fetchData() {
            const token = GM_getValue(this.STORAGE_KEY);
            if (!token || !this.state.keyword) return;

            this.state.loading = true;
            const loader = document.getElementById('gm-ps-loader');
            if (loader) loader.innerText = "正在加载数据...";

            GM_xmlhttpRequest({
                method: "GET",
                url: `${this.API_BASE}?current=${this.state.page}&size=20&keyword=${encodeURIComponent(this.state.keyword)}`,
                headers: { "X-Id-Token": token, "X-Requested-With": "XMLHttpRequest" },
                onload: (res) => {
                    this.state.loading = false;
                    try {
                        const resp = JSON.parse(res.responseText);
                        if (resp.success && resp.data.records) {
                            this.renderRows(resp.data.records);
                            const total = resp.data.total;

                            if (resp.data.records.length < 20 || this.state.page * 20 >= total) {
                                this.state.hasMore = false;
                                if (loader) loader.innerText = `— 已到底部 (共 ${total} 人) —`;
                            } else {
                                this.state.page++;
                                if (loader) loader.innerText = "向下滚动加载更多...";
                            }

                            if (total === 0 && this.state.page === 1) {
                                if (loader) loader.innerText = "没有找到匹配的人员";
                            }
                        } else {
                            if (loader) loader.innerText = "授权验证已过期，正在自动刷新...";
                            GM_setValue(this.STORAGE_KEY, "");
                            setTimeout(() => this._startSilentSync(), 1000);
                        }
                    } catch (e) {
                        if (loader) loader.innerText = "解析数据出现异常";
                    }
                },
                onerror: () => {
                    this.state.loading = false;
                    if (loader) loader.innerText = "网络请求失败，请检查网络设置";
                }
            });
        },

        renderRows(newRecords) {
            const listDiv = document.getElementById('gm-ps-list');
            if (!listDiv) return;

            const startIdx = this.state.items.length;
            this.state.items = this.state.items.concat(newRecords);

            let html = '';
            newRecords.forEach((item, i) => {
                const actualIdx = startIdx + i;
                const name = item.xm || '未知姓名';
                const id = item.gh || '未知学号/工号';
                const dept = item.yxmc || '未知单位';

                const identity = this._inferIdentity(id);
                const tagHtml = identity.short ? `<span style="font-size:12px; font-weight:normal; color:#909399; background:#f0f2f5; padding:3px 8px; border-radius:4px; letter-spacing:0.5px;">${identity.short}</span>` : '';

                html += `
                <div class="gm-ps-list-item" data-idx="${actualIdx}">
                    <div class="gm-ps-item-name">
                        <span>${esc(name)}</span>
                        ${tagHtml}
                    </div>
                    <div class="gm-ps-item-id">ID: ${esc(id)}</div>
                    <div style="font-size:13px; color:#a8abb2; white-space:nowrap; overflow:hidden; text-overflow:ellipsis;">
                        ${esc(dept)}
                    </div>
                </div>
            `;
            });

            listDiv.insertAdjacentHTML('beforeend', html);
        },

        showCardDetail(data) {
            const name = data.xm || '未知';
            const id = data.gh || '未知';
            const dept = data.yxmc || '未知所属机构';
            const sfzjh = data.sfzjh || ''; // 获取身份证号

            const identity = this._inferIdentity(id);
            const parsedIDCard = this._parseIDCard(sfzjh);

            const firstChar = name.substring(0, 1).toUpperCase();

            let hash = 0;
            for (let i = 0; i < name.length; i++) {
                hash = name.charCodeAt(i) + ((hash << 5) - hash);
            }
            const colorTheme = this.THEME_COLORS[Math.abs(hash) % this.THEME_COLORS.length];

            const avatarEl = document.getElementById('gm-ps-card-avatar');
            avatarEl.innerText = firstChar;
            avatarEl.style.color = colorTheme.text;

            const headerEl = document.getElementById('gm-ps-card-header-bg');
            headerEl.style.background = colorTheme.bg;

            document.getElementById('gm-ps-card-name').innerText = name;
            document.getElementById('gm-ps-card-id').innerText = `ID: ${id}`;
            document.getElementById('gm-ps-card-dept').innerText = dept;

            // 动态隐藏/显示“身份证提取信息”行及性别动态排版
            const extraInfoEl = document.getElementById('gm-ps-extra-info');
            if (parsedIDCard) {
                const genderEl = document.getElementById('gm-ps-card-gender');
                const provEl = document.getElementById('gm-ps-card-prov');

                provEl.innerText = parsedIDCard.province;

                // 核心逻辑：判断性别是否已知
                if (parsedIDCard.gender && parsedIDCard.gender !== '未知') {
                    // 性别已知：显示双列网格
                    genderEl.innerText = parsedIDCard.gender;
                    genderEl.parentElement.style.display = 'flex';
                    extraInfoEl.style.gridTemplateColumns = '1fr 1fr';
                } else {
                    // 性别未知（如港澳台通行证）：隐藏性别列，籍贯列占满整行
                    genderEl.parentElement.style.display = 'none';
                    extraInfoEl.style.gridTemplateColumns = '1fr';
                }

                extraInfoEl.style.display = 'grid'; // 恢复显示整体区块
            } else {
                extraInfoEl.style.display = 'none'; // 解析失败或无数据时整体隐藏
            }

            // 动态隐藏/显示“身份特征”行
            const typeRowEl = document.getElementById('gm-ps-card-type-row');
            if (identity.long) {
                document.getElementById('gm-ps-card-type').innerText = identity.long;
                typeRowEl.style.display = 'flex';
            } else {
                typeRowEl.style.display = 'none';
            }

            document.getElementById('gm-ps-empty-tip').style.display = 'none';

            const card = document.getElementById('gm-ps-detail-card');
            card.classList.remove('show');
            void card.offsetWidth;
            card.classList.add('show');
        }
    };

    // =-=-=-=-=-=-=-=-=-=-=-=-= 2.13 我的课表教材信息显示 =-=-=-=-=-=-=-=-=-=-=-=-=
    const TextbookInfoModule = {
        // 网课平台地址配置
        PLATFORM_MAP: {
            '学堂在线': 'https://bknwpu.yuketang.cn/',
            '中国大学MOOC': 'https://www.icourse163.org/spoc/schoolcloud/index.htm',
            '超星': 'http://nwpu.mooc.chaoxing.com',
            'MOOC': 'https://www.icourse163.org/spoc/schoolcloud/index.htm',
            '智慧树': 'https://www.zhihuishu.com/',
            '知到': 'https://www.zhihuishu.com/',
            '学习通': 'https://cx.chaoxing.com/'
        },

        _cachedData: null,

        // 【阶段1】立即执行：挂载网络拦截器 (解决首次加载抓不到包的问题)
        installHook() {
            if (unsafeWindow.XMLHttpRequest.prototype._gm_textbook_hooked) return;
            unsafeWindow.XMLHttpRequest.prototype._gm_textbook_hooked = true;

            const _send = unsafeWindow.XMLHttpRequest.prototype.send;
            const _open = unsafeWindow.XMLHttpRequest.prototype.open;
            const that = this;

            unsafeWindow.XMLHttpRequest.prototype.open = function (method, url) {
                this._gm_textbook_url = url;
                return _open.apply(this, arguments);
            };

            unsafeWindow.XMLHttpRequest.prototype.send = function (data) {
                this.addEventListener('load', function () {
                    if (this._gm_textbook_url && this._gm_textbook_url.includes('/print-data/')) {
                        try {
                            const responseJson = JSON.parse(this.responseText);
                            // 如果此时 UI 还没初始化（DOM还没ready），就先把数据存起来
                            if (document.readyState === 'loading' || !document.body) {
                                Logger.log('2.13', '拦截到数据，页面未就绪，已缓存');
                                that._cachedData = responseJson;
                            } else {
                                // 否则直接处理
                                that.processData(responseJson);
                            }
                        } catch (e) {
                            Logger.error('2.13', '解析课表 print-data 失败', e);
                        }
                    }
                }, { once: true });
                return _send.apply(this, arguments);
            };
            Logger.log('2.13', 'XHR拦截器已立即挂载');
        },

        // 【阶段2】延迟执行：初始化界面
        initUI() {
            if (!window.location.href.includes('/student/for-std/course-table')) return;
            Logger.log('2.13', 'UI模块初始化');

            this.injectStyles();

            // 检查是否有缓存的数据待处理
            if (this._cachedData) {
                Logger.log('2.13', '发现缓存数据，开始渲染...');
                this.processData(this._cachedData);
                this._cachedData = null; // 清空缓存
            }
        },

        // 递归提取课程信息
        processData(jsonData) {
            const courseMap = new Map();

            const findCourses = (obj) => {
                if (Array.isArray(obj)) {
                    obj.forEach(item => findCourses(item));
                } else if (obj !== null && typeof obj === 'object') {
                    if (obj.course && obj.course.id && obj.course.nameZh) {
                        courseMap.set(obj.course.id, obj.course.nameZh);
                    }
                    for (let key in obj) {
                        findCourses(obj[key]);
                    }
                }
            };

            findCourses(jsonData);

            if (courseMap.size > 0) {
                Logger.log('2.13', `提取到 ${courseMap.size} 门课程，准备获取教材信息`);
                this.fetchTextbooks(courseMap);
            } else {
                Logger.log('2.13', `未能提取到课程信息`);
                if (window.location.href.includes('/student/for-std/course-table')) {
                    this.renderContainer('当前课表数据为空，未提取到本学期的课程信息。');
                }
            }
        },

        // 并发获取教材详情页面并解析
        async fetchTextbooks(courseMap) {
            this.renderContainer('正在努力获取本学期课程的教材与网课信息，请稍候...');

            const allTextbooks = [];
            const promises = [];

            for (const [courseId, courseName] of courseMap.entries()) {
                const p = fetch(`https://jwxt.nwpu.edu.cn/student/for-std/lesson-search/info/${courseId}`)
                    .then(res => res.text())
                    .then(html => {
                        const parser = new DOMParser();
                        const doc = parser.parseFromString(html, "text/html");

                        let isOnlineCourse = false;
                        let platformName = '';

                        const baseItems = doc.querySelectorAll('.base-info-item');
                        baseItems.forEach(item => {
                            const labelEl = item.querySelector('.base-info-label');
                            const valueEl = item.querySelector('.base-info-value');
                            if (labelEl && valueEl) {
                                const label = labelEl.innerText.trim();
                                const value = valueEl.innerText.trim();
                                if (label === '课程负责人' && value.includes('在线开放课程')) isOnlineCourse = true;
                                if (label === '平台链接') platformName = value;
                            }
                        });

                        if (isOnlineCourse) {
                            allTextbooks.push({
                                courseId: courseId, courseName: courseName, isOnline: true,
                                platformName: platformName || '未知平台',
                                name: '-', author: '-', isbn: '-', publisher: '-', edition: '-', pubDate: '-'
                            });
                            return;
                        }

                        const rows = doc.querySelectorAll('.textbook-table tbody tr');
                        rows.forEach(row => {
                            const tds = row.querySelectorAll('td');
                            if (tds.length === 0) return;
                            const offset = tds.length >= 8 ? 2 : 1;
                            if (tds.length < 6) return;

                            allTextbooks.push({
                                courseId: courseId, courseName: courseName, isOnline: false,
                                name: tds[offset] ? tds[offset].innerText.trim() : '-',
                                author: tds[offset + 1] ? tds[offset + 1].innerText.trim() : '-',
                                isbn: tds[offset + 2] ? tds[offset + 2].innerText.trim() : '-',
                                publisher: tds[offset + 3] ? tds[offset + 3].innerText.trim() : '-',
                                edition: tds[offset + 4] ? tds[offset + 4].innerText.trim() : '-',
                                pubDate: tds[offset + 5] ? tds[offset + 5].innerText.trim() : '-'
                            });
                        });
                    })
                    .catch(err => {
                        Logger.warn('2.13', `获取 ${courseName} 信息失败`, err);
                    });
                promises.push(p);
            }

            await Promise.allSettled(promises);
            this.renderTable(allTextbooks);
        },

        // 注入UI样式
        injectStyles() {
            if (document.getElementById('gm-textbook-style')) return;
            const style = getOrCreateStyleEl('gm-textbook-style');
            style.textContent = `
            .gm-textbook-wrapper {
                margin: 20px; padding: 20px; background: #fff;
                border-radius: 8px; box-shadow: 0 4px 12px rgba(0,0,0,0.1);
                font-family: "Helvetica Neue", Helvetica, "PingFang SC", "Microsoft YaHei", Arial, sans-serif;
            }
            .gm-textbook-title {
                font-size: 16px; font-weight: bold; color: #303133; margin-bottom: 15px;
                border-left: 4px solid #409EFF; padding-left: 10px;
            }
            .gm-textbook-table {
                width: 100%; border-collapse: collapse; font-size: 13px;
            }
            .gm-textbook-table th, .gm-textbook-table td {
                border: 1px solid #ebeef5; padding: 10px 15px; text-align: center; vertical-align: middle;color: #606266;
            }
            .gm-textbook-table th {
                background: #f5f7fa; font-weight: bold; color: #333;
            }
            .gm-textbook-table tr:hover { background-color: #f5f7fa; }
            .gm-textbook-empty { text-align: center; color: #909399; padding: 30px; }
            .gm-textbook-course { font-weight: bold; color: #409EFF; }
            .gm-textbook-course a { color: #409EFF; text-decoration: none; transition: color 0.2s; }
            .gm-textbook-course a:hover { color: #66b1ff; text-decoration: underline; }
            .gm-online-cell { background: #fdf6ec; color: #e6a23c !important; font-weight: 500; }
            .gm-online-btn {
                display: inline-block; padding: 4px 12px; margin-left: 10px;
                background-color: #409EFF; color: #fff; border-radius: 4px;
                text-decoration: none; font-size: 12px; transition: background 0.3s;
            }
            .gm-online-btn:hover { background-color: #66b1ff; color: #fff; text-decoration: none;}
        `;
            document.head.appendChild(style);
        },

        // 渲染基础容器
        renderContainer(msg) {
            let container = document.getElementById('gm-textbook-container');
            if (!container) {
                container = document.createElement('div');
                container.id = 'gm-textbook-container';
                container.className = 'gm-textbook-wrapper';
                const target = document.querySelector('.main-content') || document.querySelector('#app') || document.body;
                target.appendChild(container);
            }
            container.innerHTML = `
            <div class="gm-textbook-title">本学期课程教材与网课清单</div>
            <div class="gm-textbook-empty">${msg}</div>
        `;
        },

        // 渲染最终表格
        renderTable(dataList) {
            if (dataList.length === 0) {
                this.renderContainer('本学期的所有课程目前均未在教务系统中登记教材信息。');
                return;
            }

            const uniqueKeys = new Set();
            const finalData = [];
            dataList.forEach(item => {
                const key = item.isOnline ? `ONLINE-${item.courseId}` : `${item.courseName}-${item.isbn}-${item.name}`;
                if (!uniqueKeys.has(key)) {
                    uniqueKeys.add(key);
                    finalData.push(item);
                }
            });

            finalData.sort((a, b) => a.courseName.localeCompare(b.courseName));

            const courseCountMap = {};
            finalData.forEach(item => {
                courseCountMap[item.courseName] = (courseCountMap[item.courseName] || 0) + 1;
            });

            let rowsHtml = '';
            let currentCourse = '';

            finalData.forEach(tb => {
                rowsHtml += `<tr>`;

                if (tb.courseName !== currentCourse) {
                    currentCourse = tb.courseName;
                    const courseUrl = `https://jwxt.nwpu.edu.cn/student/for-std/lesson-search/info/${tb.courseId}`;
                    rowsHtml += `<td rowspan="${courseCountMap[currentCourse]}" class="gm-textbook-course">
                                <a href="${courseUrl}" target="_blank" title="查看课程详情">${esc(tb.courseName)}</a>
                             </td>`;
                }

                if (tb.isOnline) {
                    let targetUrl = '';
                    for (const [key, url] of Object.entries(this.PLATFORM_MAP)) {
                        if (tb.platformName.includes(key)) { targetUrl = url; break; }
                    }

                    const platformDisplay = tb.platformName || "在线平台";
                    const linkHtml = targetUrl
                        ? `<a href="${targetUrl}" target="_blank" class="gm-online-btn">跳转至 ${esc(platformDisplay)}</a>`
                        : `<span style="margin-left:10px; color:#999; font-size:12px;">(暂无跳转链接)</span>`;

                    rowsHtml += `
                    <td colspan="6" class="gm-online-cell">
                        <span style="margin-right:8px;"><svg viewBox="0 0 24 24" width="14" height="14" style="vertical-align:-2px; margin-right:3px;" stroke="#909399" stroke-width="2" fill="none" stroke-linecap="round" stroke-linejoin="round"><path d="M18 10h-1.26A8 8 0 1 0 9 20h9a5 5 0 0 0 0-10z"></path></svg>在线开放课程</span>
                        <span>平台：${esc(platformDisplay)}</span>
                        ${linkHtml}
                    </td>
                </tr>`;
                } else {
                    rowsHtml += `
                    <td>${esc(tb.name)}</td>
                    <td>${esc(tb.author)}</td>
                    <td>${esc(tb.publisher)}</td>
                    <td>${esc(tb.isbn)}</td>
                    <td>${esc(tb.edition)}</td>
                    <td>${esc(tb.pubDate)}</td>
                </tr>`;
                }
            });

            const container = document.getElementById('gm-textbook-container');
            if (container) {
                container.innerHTML = `
                <div class="gm-textbook-title">本学期课程教材与网课清单</div>
                <table class="gm-textbook-table">
                    <thead>
                        <tr>
                            <th width="20%">课程名称</th>
                            <th width="20%">教材名称 / 平台信息</th>
                            <th width="17%">作者</th>
                            <th width="15%">出版社</th>
                            <th width="10%">ISBN/编号</th>
                            <th width="5%">版次</th>
                            <th width="8%">出版年月</th>
                        </tr>
                    </thead>
                    <tbody>
                        ${rowsHtml}
                    </tbody>
                </table>
            `;
            }
        }
    };

    // 教材模块【阶段1】：必须在页面自身 XHR 发出前挂载拦截器，故在加载期执行（内部防重复）
    try {
        TextbookInfoModule.installHook();
    } catch (e) { Logger.error('2.13', '教材拦截器挂载失败', e); }

    // --- 3. 脚本主入口（路由分发，全部页面初始化的唯一入口） ---

    function runMainFeatures() {
        const href = window.location.href;

        // 0. 【最高优先级】后台 Worker
        if (window.name === BackgroundSyncSystem.WORKER_NAME) {
            BackgroundSyncSystem.startWorker();
            return;
        }

        // 门户(ecampus) 挂载与同步
        if (location.host === 'ecampus.nwpu.edu.cn') {
            PersonnelSearch.syncToken();
            if (window.top === window.self && window.location.pathname === '/main.html') {
                initializeHomePageFeatures();
                // 门户已通过统一认证，可静默换票后在后台同步教务数据；
                // 登录失效时静默失败，不弹窗打断门户使用
                setTimeout(() => {
                    BackgroundSyncSystem.initController();
                }, 5000);
            }
            return;
        }

        // 后台预取学生画像数据，加速画像页加载
        if (location.host === 'jwxt.nwpu.edu.cn' && !href.includes('/student/for-std/student-portrait') && window.top === window.self) {
            setTimeout(() => prefetchPortraitData(), 6000);
        }

        // 1. 评教页面检测
        if (href.includes('evaluation-student-frontend')) {
            window.addEventListener('load', initEvaluationHelper);
            window.addEventListener('hashchange', () => {
                if (window.location.hash.includes('byTask')) initEvaluationHelper();
            });
            setTimeout(initEvaluationHelper, 2000); // 兜底
        }

        // 2. 开课查询页面（统一入口：2.5 选课记录同步面板 + 2.6 课程关注按钮）
        else if (href.includes('/student/for-std/lesson-search')) {
            LessonSearchEnhancer.init();
            if (document.body) initLessonSearchPage();
        }

        // 3. 学生画像页面
        else if (href.includes('/student/for-std/student-portrait')) {
            if (ConfigManager.enablePortraitEnhancement) {
                enhancePortraitPage(); // 功能3
            }
        }

        // 4. 培养方案页面
        else if (href.includes('/student/for-std/program/info/') ||
            href.includes('/student/for-std/program-completion-preview/info/') ||
            href.includes('/student/for-std/majorPrograms/info/')) {
            initProgramPageEnhancement(); // 功能8
        }

        // 4.1 研究生培养方案页面（研究生教育综合系统，含 /pyxx/ 等模块）
        else if (location.host === 'yjsjy.nwpu.edu.cn') {
            initProgramPageEnhancement(); // 功能8
            // 研究生系统顶层页面也触发同一套后台同步（数据源仍为本科教务 jwxt）
            if (window.top === window.self) {
                setTimeout(() => {
                    BackgroundSyncSystem.initController();
                }, 5000);
            }
        }

        // 5. 我的课表页面
        else if (href.includes('/student/for-std/course-table')) {
            TextbookInfoModule.initUI(); // 仅执行 UI 初始化
        }

        // 6. 教师主页跨标签页自动触发搜索
        else if (href.includes('teacher.nwpu.edu.cn/search')) {
            const searchName = GM_getValue(CONSTANTS.CROSS_SEARCH_NAME_KEY);
            if (searchName) {
                const autoSearch = () => {
                    const input = document.getElementById('sea');
                    const btn = document.querySelector('.dyym2_btn');
                    if (input && btn) {
                        input.value = searchName;

                        // 派发事件，防止网页使用Vue等框架导致直接赋值无效
                        input.dispatchEvent(new Event('input', { bubbles: true }));
                        input.dispatchEvent(new Event('change', { bubbles: true }));

                        btn.click();

                        // 只有真正成功点击搜索后，才清空缓存
                        GM_setValue(CONSTANTS.CROSS_SEARCH_NAME_KEY, '');
                    }
                };

                let retryCount = 0;
                const timer = setInterval(() => {
                    // 确保输入框和搜索按钮都渲染出来再执行
                    if (document.getElementById('sea') && document.querySelector('.dyym2_btn')) {
                        clearInterval(timer);
                        // 稍微延迟 300ms 执行，避免填入的值被网页原生初始化的 JS 覆盖清空
                        setTimeout(autoSearch, 300);
                    }
                    // 延长超时时间至 15 秒(150次)，适配首次无缓存加载较慢的情况
                    else if (retryCount > 150) {
                        clearInterval(timer);
                        GM_setValue(CONSTANTS.CROSS_SEARCH_NAME_KEY, ''); // 超时兜底清空
                    }
                    retryCount++;
                }, 100);
            }
        }

        // 7. 顶层主页
        else if (window.top === window.self) {
            initializeHomePageFeatures();
            // 延迟启动后台控制器
            setTimeout(() => {
                BackgroundSyncSystem.initController();
            }, 5000);
        }
    }

    // 统一等待 DOM 就绪后再做路由分发（document-start 注入时 readyState 必为 loading）
    if (document.readyState === 'loading') {
        window.addEventListener('DOMContentLoaded', runMainFeatures, { once: true });
    }
    else {
        runMainFeatures();
    }

})();