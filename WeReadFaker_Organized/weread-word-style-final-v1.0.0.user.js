// ==UserScript==
// @name         微信读书 - Word 2019专业增强风格 
// @namespace    http://tampermonkey.net/
// @version      1.0
// @description  完美复刻 Word 2019，修复undefined图标缺失，实现标尺精准刻度对齐与无缝选项卡
// @author       AI助手 & 优化
// @match        https://weread.qq.com/*
// @icon         https://res.wx.qq.com/a/wx_fed/weread/res/static/res/images/icon/icon_weread_logo_round.2c1c8.png
// @grant        none
// @license      MIT
// @run-at       document-start
// @compatible   chrome
// @compatible   firefox
// @compatible   edge
// ==/UserScript==

(function() {
    'use strict';

    let isStyleApplied = false;
    let isScriptKilled = false;
    let mainObserver = null;

    // --- SVG 图标库 (补齐所有丢失的图标) ---
    const SVG_ICONS = {
        WORD_LOGO: '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24"><path fill="white" d="M21.5,17.5h-19C1.78,17.5,1,16.72,1,15.75V3.75C1,2.78,1.78,2,2.5,2h19C22.22,2,23,2.78,23,3.75v12 C23,16.72,22.22,17.5,21.5,17.5z M5,13.25h7V6.75H5V13.25z M12,13.25h7v-1.5h-7V13.25z M12,9.75h7V8.25h-7V9.75z M12,6.75h7V5.25h-7V6.75z"/></svg>',
        SEARCH_DARK: '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24"><path fill="#666" d="M15.5 14h-.79l-.28-.27C15.41 12.59 16 11.11 16 9.5 16 5.91 13.09 3 9.5 3S3 5.91 3 9.5 5.91 16 9.5 16c1.61 0 3.09-.59 4.23-1.57l.27.28v.79l5 4.99L20.49 19l-4.99-5zm-6 0C7.01 14 5 11.99 5 9.5S7.01 5 9.5 5 14 7.01 14 9.5 11.99 14 9.5 14z"/></svg>',
        LIGHTBULB: '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24"><path fill="#666" d="M12 2C8.13 2 5 5.13 5 9c0 2.38 1.19 4.47 3 5.74V17c0 .55.45 1 1 1h6c.55 0 1-.45 1-1v-2.26c1.81-1.27 3-3.36 3-5.74 0-3.87-3.13-7-7-7zm2.85 11.1l-.85.6V16h-4v-2.3l-.85-.6A4.997 4.997 0 017 9c0-2.76 2.24-5 5-5s5 2.24 5 5c0 1.63-.8 3.16-2.15 4.1zM10 20h4v1h-4zm-1-2h6v1H9z"/></svg>',
        MINIMIZE: '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 16 16"><path fill="white" d="M14 8v1H3V8h11z"/></svg>',
        MAXIMIZE: '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 16 16"><path fill="white" d="M3 3v10h10V3H3zm9 9H4V4h8v8z"/></svg>',
        CLOSE: '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 16 16"><path fill="white" d="M9.06 8l3.47-3.47a.75.75 0 00-1.06-1.06L8 6.94 4.53 3.47a.75.75 0 00-1.06 1.06L6.94 8 3.47 11.47a.75.75 0 001.06 1.06L8 9.06l3.47 3.47a.75.75 0 001.06-1.06L9.06 8z"/></svg>',
        CLOSE_DARK: '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 16 16"><path fill="#666" d="M9.06 8l3.47-3.47a.75.75 0 00-1.06-1.06L8 6.94 4.53 3.47a.75.75 0 00-1.06 1.06L6.94 8 3.47 11.47a.75.75 0 001.06 1.06L8 9.06l3.47 3.47a.75.75 0 001.06-1.06L9.06 8z"/></svg>',
        SAVE: '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 20 20"><path fill="white" d="M4 2c-1.103 0-2 .897-2 2v12c0 1.103.897 2 2 2h12c1.103 0 2-.897 2-2V4c0-1.103-.897-2-2-2H4zm0 14V4h12l.002 12H4z"/><path fill="white" d="M9 5h2v4h4v2H9v-6z"/></svg>',
        UNDO: '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 20 20"><path fill="white" d="M10 2c-4.411 0-8 3.589-8 8s3.589 8 8 8c3.212 0 6.007-1.897 7.391-4.636l-1.884-.79c-1.037 2.019-3.216 3.426-5.507 3.426-3.313 0-6-2.687-6-6s2.687-6 6-6c1.865 0 3.518.847-4.671 2.181l-3.671-.181V10h7V3h-1.08c-1.527-1.129-3.483-2-5.498-2z"/></svg>',
        REDO: '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 20 20"><path fill="white" d="M10 2c4.411 0 8 3.589 8 8s-3.589 8-8 8c-3.212 0-6.007-1.897-7.391-4.636l1.884-.79c1.037 2.019 3.216 3.426 5.507 3.426 3.313 0 6-2.687 6-6s-2.687-6-6-6c-1.865 0-3.518.847-4.671 2.181l3.671-.181V10H3V3h1.08c1.527-1.129 3.483-2 5.498-2z"/></svg>',
        TOUCH_MOUSE: '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 20 20"><path fill="white" d="M10 2c-4.411 0-8 3.589-8 8s3.589 8 8 8c3.212 0 6.007-1.897 7.391-4.636l-1.884-.79c-1.037 2.019-3.216 3.426-5.507 3.426-3.313 0-6-2.687-6-6s2.687-6 6-6c1.865 0 3.518.847-4.671 2.181l-3.671-.181V10h7V3h-1.08c-1.527-1.129-3.483-2-5.498-2z"/></svg>',
        PASTE: '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24"><path fill="currentColor" d="M19 2h-4.5c-.32 0-.6.1-.85.3L12.5 3.5h-1c-.26-.2-.53-.3-.85-.3H5c-1.1 0-2 .9-2 2v14c0 1.1.9 2 2 2h14c1.1 0 2-.9 2-2V4c0-1.1-.9-2-2-2zm-6 16H8v-2h5v2zm3-4H8v-2h8v2zm0-4H8V8h8v2z"/></svg>',
        CUT: '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24"><path fill="currentColor" d="M9 4c-1.1 0-2 .9-2 2s.9 2 2 2 2-.9 2-2-.9-2-2-2zM4 16c-1.1 0-2 .9-2 2s.9 2 2 2 2-.9 2-2-.9-2-2-2zm5-12c-1.1 0-2 .9-2 2s.9 2 2 2 2-.9 2-2-.9-2-2-2zm5 12c-1.1 0-2 .9-2 2s.9 2 2 2 2-.9 2-2-.9-2-2-2zM9 2v3h6V2H9zm-5 13v3h6v-3H4zm10 0v3h6v-3h-6zm0-11.23L15.23 3H19v4.23L17.77 8H14v-.23z"/></svg>',
        COPY: '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24"><path fill="currentColor" d="M16 1H4c-1.1 0-2 .9-2 2v14h2V3h12V1zm3 4H8c-1.1 0-2 .9-2 2v14c0 1.1.9 2 2 2h11c1.1 0 2-.9 2-2V7c0-1.1-.9-2-2-2zm0 16H8V7h11v14z"/></svg>',
        FORMAT_PAINTER: '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24"><path fill="currentColor" d="M18 4H6c-1.1 0-2 .9-2 2v12c0 1.1.9 2 2 2h12c1.1 0 2-.9 2-2V6c0-1.1-.9-2-2-2zm0 14H6V6h12v12z"/><path fill="currentColor" d="M13 14H9v-2h4v2zM15 10H9V8h6v2z"/></svg>',
        BOLD: '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24"><path fill="currentColor" d="M15.6 10.79c.97-1.1 1.6-2.58 1.6-4.29 0-3.53-2.73-5.21-6.1-5.21H3v18h10.42c3.47 0 6.09-1.74 6.09-5.32 0-2.45-1.16-4.04-2.91-4.38zM8.5 4.5h3c1.24 0 2.25.96 2.25 2.2s-1.01 2.2-2.25 2.2h-3V4.5zm4.5 13H8.5V14.5h4.5c1.24 0 2.25.96 2.25 2.2s-1.01 2.2-2.25 2.2z"/></svg>',
        ITALIC: '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24"><path fill="currentColor" d="M10 4.5h4v2h-2.5l-3.5 12h2.5v2h-4v-2h2.5L12 6.5H9.5v-2z"/></svg>',
        UNDERLINE: '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24"><path fill="currentColor" d="M12 17c3.31 0 6-2.69 6-6V3h-2v8c0 2.21-1.79 4-4 4s-4-1.79-4-4V3H6v8c0 3.31 2.69 6 6 6zM5 19h14v2H5v-2z"/></svg>',
        STRIKETHROUGH: '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24"><path fill="currentColor" d="M13.5 10.5H19v2h-5.5zm-8 0H11v2H5.5zM3 14h18v2H3zm-2-7h22v2H1z"/></svg>',
        TEXT_COLOR: '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24"><path fill="currentColor" d="M0 20h24v4H0z"/><path fill="currentColor" d="M11 3L5.5 17h2.25l1.12-3h6.25l1.12 3h2.25L13 3zm-1.38 9L12 6.16 14.38 12H9.62z"/></svg>',
        HIGHLIGHT: '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24"><path fill="currentColor" d="M6 14l3 3l6-6l3 3V8l-6-6-6 6zm10.5-9.5l3 3L12 18H9l3-3 4.5-4.5z"/></svg>',
        BULLET_LIST: '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24"><path fill="currentColor" d="M4 10.5c-.83 0-1.5.67-1.5 1.5s.67 1.5 1.5 1.5 1.5-.67 1.5-1.5-.67-1.5-1.5-1.5zM4 4.5c-.83 0-1.5.67-1.5 1.5S3.17 7.5 4 7.5s1.5-.67 1.5-1.5S4.83 4.5 4 4.5zm0 6c-.83 0-1.5.67-1.5 1.5s.67 1.5 1.5 1.5 1.5-.67 1.5-1.5-.67-1.5-1.5-1.5zM6 12h14v-2H6v2zm0-8v2h14V4H6zm0 8h14v-2H6v2z"/></svg>',
        NUMBER_LIST: '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24"><path fill="currentColor" d="M2 17h2v.5H3v1h1v.5H2v1h3v-4H2v1zm1-9h1V4H2v1h1v3zm-1 9h1V4H2v1h1v3zm1 9h1V4H2v1h1v3zm1-9h1V4H2v1h1v3zm-1 9h1V4H2v1h1v3zm-1 9h1V4H2v1h1v3zM6 6h14v2H6V6zm0 12h14v2H6v-2zm0-6h14v2H6v-2z"/></svg>',
        ALIGN_LEFT: '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24"><path fill="currentColor" d="M15 15H3v2h12v-2zm0-8H3v2h12V7zM3 13h18v-2H3v2zm0 8h18v-2H3v2zM3 3v2h18V3H3z"/></svg>',
        ALIGN_CENTER: '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24"><path fill="currentColor" d="M7 15v2h10v-2H7zm-4 6h18v-2H3v2zm0-8h18v-2H3v2zm4-6v2h10V7H7zM3 3v2h18V3H3z"/></svg>',
        ALIGN_RIGHT: '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24"><path fill="currentColor" d="M3 15h12v2H3zm0-8h12v2H3zm6 4h12V9H9zm0 8h12v-2H9zm-6-4h18v-2H3z"/></svg>',
        JUSTIFY: '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24"><path fill="currentColor" d="M3 21h18v-2H3v2zm0-4h18v-2H3v2zm0-4h18V7H3v2zm0-4h18V3H3v2z"/></svg>',
        LINE_SPACING: '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24"><path fill="currentColor" d="M16 11H8v2h8v-2zm-2 4H8v2h6v-2zm-4-8H8v2h2V7zm10 0h-8v2h8V7zm0 8h-8v2h8v-2z"/></svg>',
        HEADING1: '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24"><path fill="currentColor" d="M17 11h-4V4H7v7H3v2h4v7h4v-7h4v7h4v-7h4v-2z"/></svg>',
        HEADING2: '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24"><path fill="currentColor" d="M17 11h-4V4H7v7H3v2h4v7h4v-7h4v7h4v-7h4v-2z"/></svg>',
        NORMAL_TEXT: '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24"><path fill="currentColor" d="M14 17H4v-2h10v2zm0-4H4v-2h10v2zm0-4H4V7h10v2zm6 0v2h-2V9h2z"/></svg>',
        QUOTE: '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24"><path fill="currentColor" d="M6 17h3l2-4V7H5v6h3zm8 0h3l2-4V7h-6v6h3z"/></svg>',
        FIND: '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24"><path fill="currentColor" d="M15.5 14h-.79l-.28-.27C15.41 12.59 16 11.11 16 9.5 16 5.91 13.09 3 9.5 3S3 5.91 3 9.5 5.91 16 9.5 16c1.61 0 3.09-.59 4.23-1.57l.27.28v.79l5 4.99L20.49 19l-4.99-5zm-6 0C7.01 14 5 11.99 5 9.5S7.01 5 9.5 5 14 7.01 14 9.5 11.99 14 9.5 14z"/></svg>',
        REPLACE: '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24"><path fill="currentColor" d="M16 10c0-1.1-.9-2-2-2h-3V5c0-1.1-.9-2-2-2s-2 .9-2 2v3H5c-1.1 0-2 .9-2 2s.9 2 2 2h3v3c0 1.1.9 2 2 2s2-.9 2-2v-3h3c1.1 0 2-.9 2-2zm-6 0v-2h2v2h-2zM9 13.5l1.5-1.5L9 10.5V13.5zm7.5 1.5L15 13.5 16.5 12v3z"/></svg>',
        SELECT: '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24"><path fill="currentColor" d="M3 13h18v-2H3v2zm0 4h18v-2H3v2zm0-8h18V7H3v2zm0-4h18V3H3v2z"/></svg>',
        ARROW_LEFT: '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="white"><path d="M15.41 7.41L14 6l-6 6 6 6 1.41-1.41L10.83 12z"/></svg>',
        ARROW_RIGHT: '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="white"><path d="M10 6L8.59 7.41 13.17 12l-4.58 4.59L10 18l6-6z"/></svg>',

        // 滚动条与状态栏图标
        SCROLL_UP: '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 16 16"><path fill="#606060" d="M8 5l4 4H4z"/></svg>',
        SCROLL_DOWN: '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 16 16"><path fill="#606060" d="M8 11L4 7h8z"/></svg>',
        SPELL_CHECK: '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24"><path fill="#555" d="M18 2H6c-1.1 0-2 .9-2 2v16c0 1.1.9 2 2 2h12c1.1 0 2-.9 2-2V4c0-1.1-.9-2-2-2zm0 18H6V4h12v16zm-4.3-5.3l-1.4-1.4-1.4 1.4-1.4-1.4 1.4-1.4-1.4-1.4 1.4-1.4 1.4 1.4 1.4-1.4 1.4 1.4-1.4 1.4 1.4 1.4z"/></svg>',
        MACRO: '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24"><path fill="#555" d="M20 3H4c-1.1 0-2 .9-2 2v11c0 1.1.9 2 2 2h3l-1 1v2h12v-2l-1-1h3c1.1 0 2-.9 2-2V5c0-1.1-.9-2-2-2zm0 13H4V5h16v11zM13 8h-2v2h2V8zm0 4h-2v2h2v-2z"/></svg>',
        VIEW_READ: '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24"><path fill="#555" d="M21 4H3c-1.1 0-2 .9-2 2v13c0 1.1.9 2 2 2h18c1.1 0 2-.9 2-2V6c0-1.1-.9-2-2-2zM3 19V6h8v13H3zm18 0h-8V6h8v13z"/></svg>',
        VIEW_PRINT: '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24"><path fill="#555" d="M14 2H6c-1.1 0-1.99.9-1.99 2L4 20c0 1.1.89 2 1.99 2H18c1.1 0 2-.9 2-2V8l-6-6zm2 16H8v-2h8v2zm0-4H8v-2h8v2zm-3-5V3.5L18.5 9H13z"/></svg>',
        VIEW_WEB: '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24"><path fill="#555" d="M12 2C6.48 2 2 6.48 2 12s4.48 10 10 10 10-4.48 10-10S17.52 2 12 2zm-1 17.93c-3.95-.49-7-3.85-7-7.93 0-.62.08-1.21.21-1.79L9 15v1c0 1.1.9 2 2 2v1.93zm6.9-2.54c-.26-.81-1-1.39-1.9-1.39h-1v-3c0-.55-.45-1-1-1H8v-2h2c.55 0 1-.45 1-1V7h2c1.1 0 2-.9 2-2v-.41c2.93 1.19 5 4.06 5 7.41 0 2.08-.8 3.97-2.1 5.39z"/></svg>'
    };

    const WORD_STYLE_CSS = `
        /* ========= 核心顶栏与 Ribbons ========= */
        #word-ribbon-container {
            position: fixed !important; top: 0 !important; left: 0 !important; width: 100% !important;
            background-color: white !important; z-index: 10001 !important;
            font-family: 'Segoe UI', 'Microsoft YaHei', sans-serif !important;
            box-shadow: 0 1px 4px rgba(0,0,0,0.1) !important; border-bottom: 1px solid #d0d0d0 !important;
        }
        #word-top-bar {
            background-color: #2b579a !important; height: 32px !important; width: 100% !important; display: flex !important;
            align-items: center !important; padding: 0 8px !important; color: white !important; font-size: 12px !important; box-sizing: border-box !important;
            position: relative !important;
        }
        .word-top-left { display: flex !important; align-items: center !important; margin-left: 8px !important; }
        .word-top-btn { width: 24px !important; height: 24px !important; display: flex !important; align-items: center !important; justify-content: center !important; border-radius: 3px !important; cursor: pointer !important; margin-right: 4px !important; }
        .word-top-btn:hover { background-color: rgba(255,255,255,0.15) !important; }
        .word-top-btn svg { width: 16px !important; height: 16px !important; fill: white !important; }
        .word-logo-img { width: 20px !important; height: 20px !important; margin-right: 12px !important; }
        .word-doc-title { position: absolute !important; left: 50% !important; transform: translateX(-50%) !important; font-weight: 400 !important; font-size: 13px !important; }

        .word-top-right { display: flex !important; align-items: center !important; margin-left: auto !important; }
        .word-user-info { display: flex !important; align-items: center !important; margin-right: 16px !important; cursor: pointer !important; }
        .word-user-avatar { width: 24px !important; height: 24px !important; border-radius: 50% !important; background-color: #8e918f !important; color: white !important; display: flex !important; align-items: center !important; justify-content: center !important; font-size: 14px !important; margin-right: 8px !important; }

        .word-window-controls { display: flex !important; align-items: center !important; height: 100% !important; }
        .word-window-btn { width: 45px !important; height: 100% !important; display: flex !important; align-items: center !important; justify-content: center !important; cursor: pointer !important; }
        .word-window-btn:hover { background-color: rgba(255,255,255,0.15) !important; }
        .word-window-btn.close-btn:hover { background-color: #e81123 !important; }
        .word-window-btn.close-btn:hover svg { fill: white !important; }
        .word-window-btn svg { width: 16px !important; height: 16px !important; fill: white !important; }

        /* ========= 选项卡 (Tabs) 无缝融合 ========= */
        #word-ribbon-tabs { display: flex !important; padding: 0 8px !important; height: 30px !important; background-color: #f3f2f1 !important; border-bottom: 1px solid #d0d0d0 !important; }
        .word-ribbon-tab { padding: 0 12px !important; height: 100% !important; display: flex !important; align-items: center !important; font-size: 13px !important; color: #444 !important; cursor: pointer !important; border: 1px solid transparent !important; margin-right: 2px !important; transition: background-color 0.2s, border-color 0.2s !important; }
        .word-ribbon-tab.tab-file { background-color: #2b579a !important; color: white !important; }
        .word-ribbon-tab.tab-file:hover { background-color: #1a365d !important; }
        .word-ribbon-tab.tab-tell-me { color: #444 !important; font-size: 12px !important; }
        .word-ribbon-tab.tab-tell-me svg { width: 14px !important; height: 14px !important; margin-right: 4px !important; }
        .word-ribbon-tab:hover:not(.tab-file) { background-color: #e5e5e5 !important; }

        /* 选中状态：覆盖底边框，实现无缝连接下方工具栏 */
        .word-ribbon-tab.active {
            background-color: white !important;
            border: 1px solid #d0d0d0 !important;
            border-bottom: 1px solid white !important;
            margin-bottom: -1px !important;
            z-index: 2 !important;
            color: #2b579a !important;
            font-weight: 600 !important;
        }

        #word-ribbon-content { height: 80px !important; padding: 5px 10px !important; display: flex !important; background-color: white !important; align-items: flex-start !important; overflow-x: auto !important; scrollbar-width: none !important; -ms-overflow-style: none !important; border-bottom: 1px solid #d0d0d0 !important; border-top: 1px solid transparent !important; }
        #word-ribbon-content::-webkit-scrollbar { display: none !important; }
        .word-ribbon-group { display: flex !important; flex-direction: column !important; border-right: 1px solid #d0d0d0 !important; padding: 0px 5px !important; min-width: 60px !important; height: 100% !important; box-sizing: border-box !important; justify-content: flex-end !important; }
        .word-ribbon-group:last-child { border-right: none !important; }
        .word-ribbon-group-title { font-size: 10px !important; color: #666 !important; text-align: center !important; margin-top: auto !important; padding-bottom: 2px !important; }
        .word-ribbon-group-content { display: flex !important; flex-wrap: wrap !important; justify-content: center !important; align-items: center !important; flex-grow: 1 !important; }
        .word-ribbon-item { height: 70px !important; min-width: 40px !important; text-align: center !important; margin: 2px 1px !important; padding: 3px !important; border-radius: 3px !important; display: flex !important; flex-direction: column !important; align-items: center !important; justify-content: center !important; cursor: pointer !important; font-size: 11px !important; color: #444 !important; box-sizing: border-box !important; }
        .word-ribbon-item.large { height: 70px !important; width: 60px !important; justify-content: center !important; padding-top: 10px !important; }
        .word-ribbon-item:hover { background-color: #e5e5e5 !important; }
        .word-ribbon-item-icon { font-size: 16px !important; height: 24px !important; display: flex !important; align-items: center !important; justify-content: center !important; margin-bottom: 4px !important; }
        .word-ribbon-item-icon svg { width: 20px !important; height: 20px !important; fill: currentColor !important; }
        .word-ribbon-item-label { font-size: 10px !important; line-height: 1.2 !important; }

        /* ========= 导航窗格 (Navigation Pane) ========= */
        #word-nav-pane {
            position: fixed !important; top: 142px !important; left: 0 !important; width: 260px !important; bottom: 22px !important;
            background-color: #f3f2f1 !important; border-right: 1px solid #d0d0d0 !important; z-index: 10000 !important;
            box-sizing: border-box !important; display: flex !important; flex-direction: column !important;
            font-family: 'Segoe UI', 'Microsoft YaHei', sans-serif !important;
        }
        .word-nav-title-bar { display: flex !important; justify-content: space-between !important; align-items: center !important; padding: 12px 16px 8px !important; font-size: 14px !important; font-weight: 600 !important; color: #333 !important; }
        .word-nav-close { cursor: pointer !important; width: 16px !important; height: 16px !important; opacity: 0.6 !important; }
        .word-nav-close:hover { opacity: 1 !important; }
        .word-nav-search-box { margin: 0 16px 12px !important; position: relative !important; }
        .word-nav-search-box input { width: 100% !important; padding: 4px 8px 4px 28px !important; border: 1px solid #7a7a7a !important; background: white !important; box-sizing: border-box !important; height: 28px !important; outline: none !important; font-size: 12px !important; }
        .word-nav-search-icon { position: absolute !important; left: 8px !important; top: 7px !important; width: 14px !important; height: 14px !important; }
        .word-nav-tabs { display: flex !important; border-bottom: 1px solid #d0d0d0 !important; margin: 0 16px !important; }
        .word-nav-tab { flex: 1 !important; text-align: center !important; padding: 4px 0 !important; font-size: 12px !important; cursor: pointer !important; color: #666 !important; border-bottom: 2px solid transparent !important; margin-bottom: -1px !important; }
        .word-nav-tab.active { color: #2b579a !important; border-bottom: 2px solid #2b579a !important; font-weight: 600 !important; }
        .word-nav-list { flex: 1 !important; overflow-y: auto !important; padding: 16px !important; background-color: #f3f2f1 !important; color: #555 !important; }
        .empty-text-main { font-size: 12px !important; margin-bottom: 12px !important; line-height: 1.5 !important; }
        .empty-text-sub { font-size: 12px !important; line-height: 1.5 !important; }

        /* ========= 精细定位标尺 (Rulers) ========= */
        #word-top-ruler {
            position: fixed !important; top: 142px !important; left: 260px !important; right: 16px !important; height: 14px !important;
            background-color: white !important; border-bottom: 1px solid #d0d0d0 !important; z-index: 10000 !important;
            display: flex !important; overflow: hidden !important; font-family: 'Segoe UI', Arial, sans-serif !important;
        }
        .ruler-margin { background-color: #f0f0f0 !important; height: 100% !important; flex-shrink: 0 !important; position: relative !important; border-right: 1px solid #c0c0c0 !important; }
        .ruler-margin.right-margin { border-right: none !important; border-left: 1px solid #c0c0c0 !important; }
        .ruler-main { background-color: #fff !important; height: 100% !important; flex: 1 !important; position: relative !important; display: flex !important; white-space: nowrap !important; }

        /* 标尺刻度容器宽度设定，控制数字对齐 */
        .ruler-cm-block { display: inline-block !important; width: 40px !important; height: 100% !important; position: relative !important; flex-shrink: 0 !important; }
        .ruler-num { position: absolute !important; top: 0px !important; left: 0px !important; transform: translateX(-50%) !important; font-size: 9px !important; color: #555 !important; line-height: 1 !important; }
        .ruler-ticks { position: absolute !important; bottom: 0 !important; width: 100% !important; height: 40% !important; display: flex !important; justify-content: space-between !important; }
        .ruler-tick { width: 1px !important; background-color: #888 !important; height: 40% !important; align-self: flex-end !important; }
        .ruler-tick.mid { height: 80% !important; }
        .ruler-tick.start { height: 100% !important; background-color: #555 !important; }

        .indent-marker-top { position: absolute !important; left: 0 !important; top: 0 !important; width: 0 !important; height: 0 !important; border-left: 4px solid transparent !important; border-right: 4px solid transparent !important; border-top: 5px solid #666 !important; transform: translateX(-50%) !important; z-index: 2 !important;}
        .indent-marker-bottom { position: absolute !important; left: 0 !important; bottom: 3px !important; width: 0 !important; height: 0 !important; border-left: 4px solid transparent !important; border-right: 4px solid transparent !important; border-bottom: 5px solid #666 !important; transform: translateX(-50%) !important; z-index: 2 !important;}
        .indent-marker-rect { position: absolute !important; left: 0 !important; bottom: 0 !important; width: 8px !important; height: 3px !important; background-color: #666 !important; transform: translateX(-50%) !important; z-index: 2 !important;}

        #word-left-ruler {
            position: fixed !important; top: 156px !important; left: 260px !important; width: 14px !important; bottom: 22px !important;
            background-color: white !important; border-right: 1px solid #d0d0d0 !important; z-index: 10000 !important;
            display: flex !important; flex-direction: column !important; overflow: hidden !important;
        }
        .ruler-v-margin { background-color: #f0f0f0 !important; width: 100% !important; flex-shrink: 0 !important; position: relative !important; border-bottom: 1px solid #c0c0c0 !important; }
        .ruler-v-margin.bottom-margin { border-bottom: none !important; border-top: 1px solid #c0c0c0 !important; }
        .ruler-v-main { background-color: #fff !important; width: 100% !important; flex: 1 !important; position: relative !important; display: flex !important; flex-direction: column !important; }
        .ruler-v-cm-block { width: 100% !important; height: 40px !important; position: relative !important; flex-shrink: 0 !important; }
        .ruler-v-num { position: absolute !important; top: 0px !important; right: 1px !important; transform: translateY(-50%) !important; font-size: 9px !important; color: #555 !important; line-height: 1 !important; }
        .ruler-v-ticks { position: absolute !important; right: 0 !important; height: 100% !important; width: 40% !important; display: flex !important; flex-direction: column !important; justify-content: space-between !important; }
        .ruler-v-tick { height: 1px !important; background-color: #888 !important; width: 40% !important; align-self: flex-end !important; }
        .ruler-v-tick.mid { width: 80% !important; }
        .ruler-v-tick.start { width: 100% !important; background-color: #555 !important; }

        /* ========= 底部状态栏 (Status Bar) ========= */
        #word-status-bar {
            position: fixed !important; bottom: 0 !important; left: 0 !important; width: 100% !important; height: 22px !important;
            background-color: #f3f2f1 !important; z-index: 10002 !important; border-top: 1px solid #d0d0d0 !important;
            display: flex !important; justify-content: space-between !important; align-items: center !important;
            font-family: 'Segoe UI', 'Microsoft YaHei', sans-serif !important; font-size: 11px !important; color: #444 !important;
            padding: 0 12px !important; box-sizing: border-box !important;
        }
        .status-left, .status-right { display: flex !important; align-items: center !important; height: 100% !important; }
        .status-item { margin-right: 16px !important; display: flex !important; align-items: center !important; cursor: default !important; }
        .status-item:hover { background-color: #e5e5e5 !important; }
        .status-icon { width: 14px !important; height: 14px !important; margin-right: 4px !important; display: flex !important; align-items: center !important; justify-content: center !important; }
        .status-icon svg { width: 100% !important; height: 100% !important; }
        .status-btn { padding: 0 6px !important; height: 100% !important; display: flex !important; align-items: center !important; justify-content: center !important; cursor: pointer !important; margin-right: 2px !important;}
        .status-btn:hover { background-color: #e5e5e5 !important; }
        .status-btn.active { background-color: #cfcfcf !important; border: 1px solid #b0b0b0 !important; box-sizing: border-box !important; }

        .status-zoom { display: flex !important; align-items: center !important; margin-left: 8px !important; }
        .zoom-btn { cursor: pointer !important; padding: 0 4px !important; font-size: 14px !important; font-weight: bold !important; color: #666 !important;}
        .zoom-slider-track { width: 100px !important; height: 2px !important; background-color: #a0a0a0 !important; margin: 0 8px !important; position: relative !important; cursor: pointer !important; }
        .zoom-slider-thumb { width: 4px !important; height: 10px !important; background-color: #555 !important; position: absolute !important; top: -4px !important; left: 60% !important; }
        .zoom-text { width: 36px !important; text-align: right !important; margin-left: 4px !important; }

        /* ========= 右侧仿真滚动条 (Fake Scrollbar) ========= */
        #word-fake-scrollbar {
            position: fixed !important; right: 0 !important; top: 142px !important; bottom: 22px !important; width: 16px !important;
            background-color: #f0f0f0 !important; border-left: 1px solid #e0e0e0 !important; z-index: 10000 !important;
            display: flex !important; flex-direction: column !important; box-sizing: border-box !important;
        }
        .scroll-btn { width: 100% !important; height: 16px !important; background-color: #f0f0f0 !important; display: flex !important; align-items: center !important; justify-content: center !important; cursor: default !important; }
        .scroll-btn:hover { background-color: #dadada !important; }
        .scroll-btn svg { width: 10px !important; height: 10px !important; }
        .scroll-track { flex: 1 !important; position: relative !important; background-color: #f0f0f0 !important; }
        .scroll-thumb { position: absolute !important; top: 5% !important; right: 1px !important; left: 1px !important; height: 15% !important; background-color: #cdcdcd !important; border: 1px solid #f0f0f0 !important; box-sizing: border-box !important; }
        .scroll-thumb:hover { background-color: #a6a6a6 !important; }

        /* ========= 页面布局与翻页适配 ========= */
        body { padding-top: 156px !important; padding-left: 274px !important; padding-right: 16px !important; padding-bottom: 22px !important; }
        .readerTopBar, .readerControls, .readerFooter, .navBar, .footerBar, .app_content.reader { display: none !important; }
        #readerContent { margin-top: 0 !important; padding-top: 0 !important; transform: none !important; transition: opacity 0.3s ease; }
        #root, .wr_whiteTheme, .app_content { margin-top: 0 !important; transition: opacity 0.3s ease; }

        .word-page-turn-btn {
            position: fixed !important; top: 50% !important; transform: translateY(-50%) !important; width: 48px !important;
            height: 90px !important; background: linear-gradient(135deg, rgba(43,87,154,0.85) 60%, rgba(43,87,154,0.6) 100%) !important;
            color: white !important; display: flex !important; align-items: center !important; justify-content: center !important;
            cursor: pointer !important; z-index: 10002 !important; opacity: 0 !important; transition: opacity 0.25s, background 0.25s, box-shadow 0.25s !important;
            border-radius: 24px !important; box-shadow: 0 4px 16px 0 rgba(43,87,154,0.12), 0 1.5px 6px 0 rgba(0,0,0,0.08) !important;
            border: 1.5px solid rgba(255,255,255,0.18) !important; pointer-events: none;
        }
        .word-page-turn-btn.visible { opacity: 1 !important; pointer-events: auto; }
        .word-page-turn-btn svg { width: 36px !important; height: 36px !important; filter: drop-shadow(0 1px 2px rgba(0,0,0,0.12)); }
        #word-page-turn-prev { left: 290px !important; }
        #word-page-turn-next { right: 34px !important; }
        .word-page-turn-btn:hover { background: linear-gradient(135deg, #2b579a 80%, #185abd 100%) !important; box-shadow: 0 8px 24px 0 rgba(43,87,154,0.22), 0 2px 8px 0 rgba(0,0,0,0.12) !important; border: 2px solid #fff !important; }
    `;

    function addStyles() {
        if (document.getElementById('word-style-css')) return;
        const styleElement = document.createElement('style');
        styleElement.id = 'word-style-css';
        styleElement.textContent = WORD_STYLE_CSS;
        if (document.head) document.head.appendChild(styleElement);
        else {
            const observer = new MutationObserver((mutations, obs) => {
                if (document.head) { document.head.appendChild(styleElement); obs.disconnect(); }
            });
            observer.observe(document.documentElement, { childList: true, subtree: true });
        }
    }

    function closeWordUI() {
        if (!isStyleApplied) return;
        isScriptKilled = true;
        isStyleApplied = false;

        const idsToRemove = ['word-style-css', 'word-ribbon-container', 'word-nav-pane', 'word-top-ruler', 'word-left-ruler', 'word-status-bar', 'word-fake-scrollbar', 'word-page-turn-prev', 'word-page-turn-next'];
        idsToRemove.forEach(id => { const el = document.getElementById(id); if (el) el.remove(); });

        const elementsToRestore = ['.readerTopBar', '.readerControls', '.readerFooter', '.navBar', '.footerBar', '.app_content.reader'];
        elementsToRestore.forEach(selector => { document.querySelectorAll(selector).forEach(el => { el.style.display = ''; }); });

        document.body.style.paddingTop = ''; document.body.style.paddingLeft = ''; document.body.style.paddingRight = ''; document.body.style.paddingBottom = '';
        const contentAreas = document.querySelectorAll('#root, .app_content, #readerContent');
        contentAreas.forEach(el => { if (el) { el.style.marginTop = ''; el.style.paddingTop = ''; el.style.transform = ''; el.style.opacity = '1'; el.style.pointerEvents = 'auto'; } });
        if (mainObserver) mainObserver.disconnect();
    }

    document.addEventListener('keydown', (e) => {
        if (e.altKey && e.key.toLowerCase() === 'w') {
            e.preventDefault();
            if (isStyleApplied) closeWordUI();
            else { isScriptKilled = false; init(); }
        }
    });

    // 动态生成标尺刻度 HTML (完美对齐算法)
    function generateRulerHTML(isHorizontal) {
        let html = '';
        if (isHorizontal) {
            html += '<div class="ruler-margin" style="width: 80px;">';
            for (let i = 2; i >= 1; i--) { html += `<div class="ruler-cm-block"><span class="ruler-num">${i}</span><div class="ruler-ticks"><div class="ruler-tick start"></div>${'<div class="ruler-tick"></div>'.repeat(4)}<div class="ruler-tick mid"></div>${'<div class="ruler-tick"></div>'.repeat(4)}</div></div>`; }
            html += '</div>';

            html += '<div class="ruler-main">';
            html += `<div class="indent-marker-top"></div><div class="indent-marker-bottom"></div><div class="indent-marker-rect"></div>`;
            for (let i = 1; i <= 30; i++) { html += `<div class="ruler-cm-block"><span class="ruler-num">${i}</span><div class="ruler-ticks"><div class="ruler-tick start"></div>${'<div class="ruler-tick"></div>'.repeat(4)}<div class="ruler-tick mid"></div>${'<div class="ruler-tick"></div>'.repeat(4)}</div></div>`; }
            html += '</div>';

            html += '<div class="ruler-margin right-margin" style="width: 80px; position:absolute; right:0; top:0;"></div>';
        } else {
            html += '<div class="ruler-v-margin" style="height: 60px;"></div>';
            html += '<div class="ruler-v-main">';
            for (let i = 1; i <= 20; i++) { html += `<div class="ruler-v-cm-block"><span class="ruler-v-num">${i}</span><div class="ruler-v-ticks"><div class="ruler-v-tick start"></div>${'<div class="ruler-v-tick"></div>'.repeat(4)}<div class="ruler-v-tick mid"></div>${'<div class="ruler-v-tick"></div>'.repeat(4)}</div></div>`; }
            html += '</div>';
            html += '<div class="ruler-v-margin bottom-margin" style="height: 60px; position:absolute; bottom:0; left:0;"></div>';
        }
        return html;
    }

    function createWordInterface() {
        if (document.getElementById('word-ribbon-container')) return;

        // --- 1. 创建顶部 Ribbon ---
        const container = document.createElement('div'); container.id = 'word-ribbon-container';
        const topBar = document.createElement('div'); topBar.id = 'word-top-bar';

        const leftArea = document.createElement('div'); leftArea.className = 'word-top-left';
        const logo = document.createElement('img'); logo.className = 'word-logo-img';
        logo.src = 'data:image/svg+xml;base64,PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHZpZXdCb3g9IjAgMCAyNCAyNCI+PHBhdGggZmlsbD0id2hpdGUiIGQ9Ik0yMS41LDE3LjVoLTE5QzEuNzgsMTcuNSwxLDE2LjcyLDEsMTUuNzVWMy43NUMxLDIuNzgsMS43OCwyLDIuNSwyaDE5QzIyLjIyLDIsMjMsMi43OCwyMywzLjc1djEyIEMyMywxNi43MiwyMi4yMiwxNy41LDIxLjUsMTcuNXogTTUsMTMuMjVoN1Y2Ljc1SDVWMTMuMjV6IE0xMiwxMy4yNWg3di0xLjVoLTdWMTMuMjV6IE0xMiw5Ljc1aDdWOC4yNWgtN1Y5Ljc1eiBNMTIsNi43NWg3VjUuMjVoLTdWNi43NXoiLz48L3N2Zz4=';
        leftArea.appendChild(logo);
        ['SAVE', 'UNDO', 'REDO', 'TOUCH_MOUSE'].forEach(iconKey => {
            const btn = document.createElement('div'); btn.className = 'word-top-btn'; btn.innerHTML = SVG_ICONS[iconKey] || ''; leftArea.appendChild(btn);
        });

        const title = document.createElement('div'); title.className = 'word-doc-title'; title.textContent = '文档1 - Word';

        const rightArea = document.createElement('div'); rightArea.className = 'word-top-right';
        const userInfo = document.createElement('div'); userInfo.className = 'word-user-info';
        userInfo.innerHTML = `<div class="word-user-avatar">作</div><span>作者</span>`;
        rightArea.appendChild(userInfo);

        const windowControls = document.createElement('div'); windowControls.className = 'word-window-controls';
        const minBtn = document.createElement('div'); minBtn.className = 'word-window-btn'; minBtn.innerHTML = SVG_ICONS.MINIMIZE; minBtn.title = '最小化(隐藏正文，保留框架)';
        const maxBtn = document.createElement('div'); maxBtn.className = 'word-window-btn'; maxBtn.innerHTML = SVG_ICONS.MAXIMIZE; maxBtn.title = '最大化(恢复正文)';
        const closeBtn = document.createElement('div'); closeBtn.className = 'word-window-btn close-btn'; closeBtn.innerHTML = SVG_ICONS.CLOSE; closeBtn.title = '关闭伪装 (快捷键 Alt+W)';

        minBtn.addEventListener('click', () => {
            const contentAreas = document.querySelectorAll('#root, .app_content');
            contentAreas.forEach(el => { if (el) { el.style.opacity = '0'; el.style.pointerEvents = 'none'; } });
            const prevBtn = document.getElementById('word-page-turn-prev'); const nextBtn = document.getElementById('word-page-turn-next');
            if (prevBtn) prevBtn.style.display = 'none'; if (nextBtn) nextBtn.style.display = 'none';
        });
        maxBtn.addEventListener('click', () => {
            const contentAreas = document.querySelectorAll('#root, .app_content');
            contentAreas.forEach(el => { if (el) { el.style.opacity = '1'; el.style.pointerEvents = 'auto'; } });
            const prevBtn = document.getElementById('word-page-turn-prev'); const nextBtn = document.getElementById('word-page-turn-next');
            if (prevBtn) prevBtn.style.display = ''; if (nextBtn) nextBtn.style.display = '';
        });
        closeBtn.addEventListener('click', closeWordUI);

        windowControls.appendChild(minBtn); windowControls.appendChild(maxBtn); windowControls.appendChild(closeBtn);
        rightArea.appendChild(windowControls);
        topBar.appendChild(leftArea); topBar.appendChild(title); topBar.appendChild(rightArea);
        container.appendChild(topBar);

        const tabs = document.createElement('div'); tabs.id = 'word-ribbon-tabs';
        ['文件', '开始', '插入', '设计', '布局', '引用', '邮件', '审阅', '视图', '帮助', '操作说明搜索'].forEach((name) => {
            const tab = document.createElement('div'); tab.className = 'word-ribbon-tab';
            if (name === '文件') { tab.classList.add('tab-file'); tab.textContent = name; }
            else if (name === '开始') { tab.classList.add('active'); tab.textContent = name; }
            else if (name === '操作说明搜索') { tab.classList.add('tab-tell-me'); tab.innerHTML = SVG_ICONS.LIGHTBULB + name; }
            else { tab.textContent = name; }
            tabs.appendChild(tab);
        });
        container.appendChild(tabs);

        const content = document.createElement('div'); content.id = 'word-ribbon-content';
        const clipboardGroup = createRibbonGroup('剪贴板', [ createRibbonItem('粘贴', SVG_ICONS.PASTE, 'large'), createRibbonItem('剪切', SVG_ICONS.CUT), createRibbonItem('复制', SVG_ICONS.COPY), createRibbonItem('格式刷', SVG_ICONS.FORMAT_PAINTER) ]);
        const fontGroup = createRibbonGroup('字体', [ createRibbonItem('加粗', SVG_ICONS.BOLD), createRibbonItem('斜体', SVG_ICONS.ITALIC), createRibbonItem('下划线', SVG_ICONS.UNDERLINE), createRibbonItem('删除线', SVG_ICONS.STRIKETHROUGH), createRibbonItem('文本颜色', SVG_ICONS.TEXT_COLOR), createRibbonItem('高亮', SVG_ICONS.HIGHLIGHT) ]);
        const paragraphGroup = createRibbonGroup('段落', [ createRibbonItem('项目符号', SVG_ICONS.BULLET_LIST), createRibbonItem('编号', SVG_ICONS.NUMBER_LIST), createRibbonItem('左对齐', SVG_ICONS.ALIGN_LEFT), createRibbonItem('居中', SVG_ICONS.ALIGN_CENTER), createRibbonItem('右对齐', SVG_ICONS.ALIGN_RIGHT), createRibbonItem('两端对齐', SVG_ICONS.JUSTIFY), createRibbonItem('行距', SVG_ICONS.LINE_SPACING) ]);
        const stylesGroup = createRibbonGroup('样式', [ createRibbonItem('标题1', SVG_ICONS.HEADING1), createRibbonItem('标题2', SVG_ICONS.HEADING2), createRibbonItem('正文', SVG_ICONS.NORMAL_TEXT), createRibbonItem('引用', SVG_ICONS.QUOTE) ]);
        const editingGroup = createRibbonGroup('编辑', [ createRibbonItem('查找', SVG_ICONS.FIND), createRibbonItem('替换', SVG_ICONS.REPLACE), createRibbonItem('选择', SVG_ICONS.SELECT) ]);
        content.appendChild(clipboardGroup); content.appendChild(fontGroup); content.appendChild(paragraphGroup); content.appendChild(stylesGroup); content.appendChild(editingGroup);
        container.appendChild(content);

        // --- 2. 创建左侧导航窗格 ---
        const navPane = document.createElement('div'); navPane.id = 'word-nav-pane';
        navPane.innerHTML = `
            <div class="word-nav-title-bar"><span>导航</span><div class="word-nav-close">${SVG_ICONS.CLOSE_DARK}</div></div>
            <div class="word-nav-search-box"><div class="word-nav-search-icon">${SVG_ICONS.SEARCH_DARK}</div><input type="text" placeholder="在文档中搜索 (Alt+W 切换伪装)"></div>
            <div class="word-nav-tabs"><div class="word-nav-tab">标题</div><div class="word-nav-tab">页面</div><div class="word-nav-tab active">结果</div></div>
            <div class="word-nav-list">
                <div class="empty-text-main">文本、批注、图片...Word可以查找您的文档中的任何内容。</div>
                <div class="empty-text-sub">请使用搜索框查找文本或者放大镜查找任何其他内容。</div>
            </div>
        `;

        // --- 3. 创建动态标尺 ---
        const topRuler = document.createElement('div'); topRuler.id = 'word-top-ruler';
        topRuler.innerHTML = generateRulerHTML(true);
        const leftRuler = document.createElement('div'); leftRuler.id = 'word-left-ruler';
        leftRuler.innerHTML = generateRulerHTML(false);

        // --- 4. 创建底部状态栏 ---
        const statusBar = document.createElement('div'); statusBar.id = 'word-status-bar';
        statusBar.innerHTML = `
            <div class="status-left">
                <div class="status-item">第 1 页，共 4 页</div>
                <div class="status-item">1815 个字</div>
                <div class="status-item"><div class="status-icon">${SVG_ICONS.SPELL_CHECK}</div></div>
                <div class="status-item">中文(中国)</div>
            </div>
            <div class="status-right">
                <div class="status-item"><div class="status-icon">${SVG_ICONS.MACRO}</div>显示器设置</div>
                <div class="status-btn"><div class="status-icon" style="margin:0;">${SVG_ICONS.VIEW_READ}</div></div>
                <div class="status-btn active"><div class="status-icon" style="margin:0;">${SVG_ICONS.VIEW_PRINT}</div></div>
                <div class="status-btn"><div class="status-icon" style="margin:0;">${SVG_ICONS.VIEW_WEB}</div></div>
                <div class="status-zoom">
                    <span class="zoom-btn">-</span>
                    <div class="zoom-slider-track"><div class="zoom-slider-thumb"></div></div>
                    <span class="zoom-btn">+</span>
                    <span class="zoom-text">120%</span>
                </div>
            </div>
        `;

        // --- 5. 创建右侧仿真滚动条 (使用SVG箭头) ---
        const fakeScrollbar = document.createElement('div'); fakeScrollbar.id = 'word-fake-scrollbar';
        fakeScrollbar.innerHTML = `
            <div class="scroll-btn">${SVG_ICONS.SCROLL_UP}</div>
            <div class="scroll-track"><div class="scroll-thumb"></div></div>
            <div class="scroll-btn">${SVG_ICONS.SCROLL_DOWN}</div>
        `;

        if (document.body) {
            document.body.insertBefore(container, document.body.firstChild);
            document.body.appendChild(navPane); document.body.appendChild(topRuler); document.body.appendChild(leftRuler);
            document.body.appendChild(statusBar); document.body.appendChild(fakeScrollbar);
            addTabsInteractivity();
        } else {
            const observer = new MutationObserver((mutations, obs) => {
                if (document.body) {
                    document.body.insertBefore(container, document.body.firstChild);
                    document.body.appendChild(navPane); document.body.appendChild(topRuler); document.body.appendChild(leftRuler);
                    document.body.appendChild(statusBar); document.body.appendChild(fakeScrollbar);
                    addTabsInteractivity(); obs.disconnect();
                }
            });
            observer.observe(document.documentElement, { childList: true, subtree: true });
        }
    }

    function createRibbonGroup(title, items) {
        const group = document.createElement('div'); group.className = 'word-ribbon-group';
        const content = document.createElement('div'); content.className = 'word-ribbon-group-content';
        items.forEach(item => content.appendChild(item));
        const titleElement = document.createElement('div'); titleElement.className = 'word-ribbon-group-title'; titleElement.textContent = title;
        group.appendChild(content); group.appendChild(titleElement); return group;
    }

    function createRibbonItem(label, iconSVG, sizeClass = '') {
        const item = document.createElement('div'); item.className = 'word-ribbon-item'; if (sizeClass) item.classList.add(sizeClass);
        const iconElement = document.createElement('div'); iconElement.className = 'word-ribbon-item-icon'; iconElement.innerHTML = iconSVG || '';
        const labelElement = document.createElement('div'); labelElement.className = 'word-ribbon-item-label'; labelElement.textContent = label;
        item.appendChild(iconElement); item.appendChild(labelElement); return item;
    }

    function addTabsInteractivity() {
        const tabs = document.querySelectorAll('.word-ribbon-tab:not(.tab-file):not(.tab-tell-me)');
        if (!tabs || tabs.length === 0) return;
        tabs.forEach(tab => { tab.addEventListener('click', () => { tabs.forEach(t => t.classList.remove('active')); tab.classList.add('active'); }); });
    }

    function hideOriginalElements() {
        const elementsToHide = ['.readerTopBar', '.readerControls', '.readerFooter', '.navBar', '.footerBar', '.app_content.reader'];
        elementsToHide.forEach(selector => { document.querySelectorAll(selector).forEach(el => { if (el) el.style.display = 'none'; }); });
    }

    function fixContentArea() {
        if (document.body) { document.body.style.paddingTop = '156px'; document.body.style.paddingLeft = '274px'; document.body.style.paddingRight = '16px'; document.body.style.paddingBottom = '22px'; }
        const readerContent = document.querySelector('#readerContent');
        if (readerContent) { readerContent.style.marginTop = '0'; readerContent.style.paddingTop = '0'; readerContent.style.transform = 'none'; }
    }

    function createPageTurnButtons() {
        if (document.getElementById('word-page-turn-prev')) return;
        const prevBtn = document.createElement('div'); prevBtn.id = 'word-page-turn-prev'; prevBtn.className = 'word-page-turn-btn'; prevBtn.innerHTML = SVG_ICONS.ARROW_LEFT;
        prevBtn.addEventListener('click', () => {
            const originalPrevBtn = document.querySelector('.readerChapter_button_prev, .wr_readerPrevPage, .wr_pagePrev, .wr-btn-prev');
            if (originalPrevBtn) { originalPrevBtn.click(); } else {
                const evt = new KeyboardEvent('keydown', { key: 'ArrowLeft', keyCode: 37, which: 37, bubbles: true }); document.dispatchEvent(evt);
            }
        });
        const nextBtn = document.createElement('div'); nextBtn.id = 'word-page-turn-next'; nextBtn.className = 'word-page-turn-btn'; nextBtn.innerHTML = SVG_ICONS.ARROW_RIGHT;
        nextBtn.addEventListener('click', () => {
            const originalNextBtn = document.querySelector('.readerChapter_button_next, .wr_readerNextPage, .wr_pageNext, .wr-btn-next');
            if (originalNextBtn) { originalNextBtn.click(); } else {
                const evt = new KeyboardEvent('keydown', { key: 'ArrowRight', keyCode: 39, which: 39, bubbles: true }); document.dispatchEvent(evt);
            }
        });
        const mainContent = document.querySelector('.app_content, .wr_whiteTheme, .readerContent, #readerContent, .readerChapterContent');
        if (mainContent) { mainContent.appendChild(prevBtn); mainContent.appendChild(nextBtn); }
        else { document.body.appendChild(prevBtn); document.body.appendChild(nextBtn); }

        let isHovering = false;
        const showButtons = () => { prevBtn.classList.add('visible'); nextBtn.classList.add('visible'); };
        const hideButtons = () => { if (!isHovering) { prevBtn.classList.remove('visible'); nextBtn.classList.remove('visible'); } };

        [prevBtn, nextBtn].forEach(btn => {
            btn.addEventListener('mouseenter', () => { isHovering = true; showButtons(); });
            btn.addEventListener('mouseleave', () => { isHovering = false; hideButtons(); });
        });

        document.addEventListener('mousemove', (e) => {
            const w = window.innerWidth;
            if (e.clientX > 260 && e.clientX <= 380) { prevBtn.classList.add('visible'); } else { if (!isHovering) prevBtn.classList.remove('visible'); }
            if (e.clientX >= w - 120) { nextBtn.classList.add('visible'); } else { if (!isHovering) nextBtn.classList.remove('visible'); }
        });
    }

    function init() {
        if (isStyleApplied || isScriptKilled) return;
        try {
            addStyles(); createWordInterface(); createPageTurnButtons(); hideOriginalElements(); fixContentArea();
            isStyleApplied = true;
            console.log('微信读书 - Word 2019增强风格 v2.1.0 初始化成功 (已修复所有图标，精细化视觉优化完成)');
        } catch (error) { console.error('初始化Word风格时出错:', error); }
    }

    addStyles();
    document.addEventListener('DOMContentLoaded', init);
    window.addEventListener('load', init);
    if (document.readyState === 'complete' || document.readyState === 'interactive') { init(); }

    mainObserver = new MutationObserver((mutations) => {
        if (document.querySelector('#readerContent') && !isStyleApplied && !isScriptKilled) { init(); }
    });
    mainObserver.observe(document.documentElement, { childList: true, subtree: true });

    setTimeout(init, 1000); setTimeout(init, 2000);
})();
