// ==UserScript==
// @name         翱翔教务功能加强
// @namespace    http://tampermonkey.net/
// @version      1.8.1
// @description  1.提供GPA分析报告；2. 导出课程成绩与教学班排名；3.更好的“学生画像”显示；4.选课助手；5.课程关注与后台同步；6.一键自动评教；7.人员信息检索
// @author       leamloli
// @match        https://jwxt.nwpu.edu.cn/*
// @match        https://jwxt.nwpu.edu.cn/student/for-std/course-select/some-page*
// @match        https://ecampus.nwpu.edu.cn/*
// @match        https://teacher.nwpu.edu.cn/*
// @grant        GM_xmlhttpRequest
// @grant        GM_setValue
// @grant        GM_getValue
// @grant        GM_registerMenuCommand
// @grant        unsafeWindow
// @connect      electronic-signature.nwpu.edu.cn
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
