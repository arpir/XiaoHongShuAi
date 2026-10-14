// ==UserScript==
// @name         小红书全能AI助手
// @namespace    http://tampermonkey.net/
// @version      2.6.1
// @description  采用API拦截技术，支持自动滚动获取全部笔记，生成带xsec_token的永久有效链接，支持导出Excel/CSV/JSON。新增AI创作模块，内置多种写作模版，支持自定义模版和AI生成人设。提升创作效率，助力内容变现！新增excel带图片导出模式，方便直观查看封面图。新增资源下载功能，支持高清图片/视频批量下载。
// @author       Coriander
// @icon         data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAACAAAAAgCAYAAABzenr0AAAAAXNSR0IArs4c6QAAAx9JREFUWEfNl09MU0EQxn/beFTDDRI41FAMcNGbBw62oPEGid6UULxg1EhEEzExgdBEEzRqlKDxZCHgDZJ6U8TWAyaQGIsHMQLSA0S8VYQT2NXp9tnX0vKnpi2TNH1vd3bmm5lv9+0o0kQ73SXsc7QCx1EcjU9rnOl6O3pXRNAqCjqCIsB6LKQioYh9rbK/6MMnWojFHgElO3KwWyUBBD1q9q3fWvoPgHY1dIHu2a3N3PRVt5ob98naOABdVd+K5nluxnJc5dBe9TU4qHS128lvRzDnOufoH4iyETukihJ9EnSH0i5PAFRj7oH8z0r9UmlXw0fQZrsVWhQRKcFCEepvQo0DcNXrQgeechDtbQAVpbCyBiurqUmqqYSD+2FyOnPyZE50ln7A4vKWCc5egvIyCA3DzV4YeZ00UlEGQ/eN88670HsjOTczZ8bbvXCiDqbC8HkeBkahuhLE5sBICqDdAzh9yjh1n4OlZZgdTxqcDEPfIAw9SI1aMjg1DVrDpe5tAIRewOJ36LyXzIAgv+IFz1ljXN5FJAOjrwwIcd583YwfO2L0JHvW2qqGjKXYnAExJkYfDyYBaGWibmyDGhe0t/z9bikDSMQO4NZlEO5YJTggfHCBf8SUIo0TqQCEPB8C0Ddg6m5xQIj4xAcXu+DLPASHjY5/1BDUDkAyWF6amXjCkcYLW5Sg1gWBZ3C7H6Y+mWdJ48y35LiQ0HvGGLHzIFsJLAJLSSQzssYmmzMg0TVfM9vMqqMYkcwIejEiv59rhliy3URP2H6n3/zXJsbsO+ipz+huCUCQSb2E3eJQRNL+ZsIQS/a1ALQIKDtCxu0i4EUs8GPvk7YEXFPbNrvAmj5ZJ3dB49wSYbTlUIgqANJFzoFfq4aE8izBiC0h49iEmctagszUyevoHvgYFf1zXEwA6PBeuJLVXwUe5pVp2Yyr2HmVaMUW8tYNZXWuI6xrT6IxcbeiHYVtTCT62ZDf1pp5ekB1FaYU2qfmgvGLQWpzKi0adOfxlhxF0ZGxObUiT7RqbjRNoJ0oVZIzINMNy5Eehtg7NvCrSChqz/IfgUZkW/BhLsQAAAAASUVORK5CYII=
// @match        https://creator.xiaohongshu.com/publish/*
// @match        https://www.xiaohongshu.com/*
// @connect      *
// @grant        GM_xmlhttpRequest
// @grant        GM_setValue
// @grant        GM_getValue
// @grant        GM_addStyle
// @grant        GM_registerMenuCommand
// @grant        GM_download
// @require      https://cdn.jsdelivr.net/npm/exceljs@4.4.0/dist/exceljs.min.js
// @require      https://cdn.jsdelivr.net/npm/jszip@3.10.1/dist/jszip.min.js
// @license      MIT
// ==/UserScript==

(function () {
  "use strict";

  // ==========================================
  // 0. 全局数据存储 (核心)
  // ==========================================
  const GLOBAL_DATA = new Map();
  let isAutoScrolling = false;
  let currentPageUrl = location.href; // 记录当前页面URL

  // 全局视频URL存储 - 页面加载时捕获的真实视频链接
  let CACHED_VIDEO_URL = null;

  // 已捕获的笔记卡片数据（来自页面自身加载笔记时的 feed/note API 响应）
  // key = noteId。页面渲染用的就是这份数据，提取时直接复用最可靠
  const CACHED_NOTE_CARDS = new Map();

  function cacheNoteCard(note, fallbackId) {
    if (!note || typeof note !== "object") return;
    const id = String(note.noteId || note.note_id || fallbackId || "");
    if (!id) return;
    if (CACHED_NOTE_CARDS.size >= 20) {
      CACHED_NOTE_CARDS.delete(CACHED_NOTE_CARDS.keys().next().value);
    }
    CACHED_NOTE_CARDS.set(id, note);
  }

  // ==========================================
  // 0.1 XHR 拦截 - 捕获实际视频流URL
  // ==========================================
  (function () {
    const originalOpen = XMLHttpRequest.prototype.open;
    const originalSend = XMLHttpRequest.prototype.send;

    XMLHttpRequest.prototype.open = function (method, url, ...rest) {
      this._xhrUrl = url;
      // 【新增】直接捕获 stream 类型的 mp4 链接 (最高优先级)
      if (
        url &&
        typeof url === "string" &&
        url.includes("xhscdn.com") &&
        url.includes("/stream/") &&
        url.includes(".mp4")
      ) {
        CACHED_VIDEO_URL = url;
        console.log("[XHS-Stream] 捕获无水印流地址:", url);
      }
      return originalOpen.apply(this, [method, url, ...rest]);
    };

    XMLHttpRequest.prototype.send = function (data, ...rest) {
      const self = this;
      const origStateChange = this.onreadystatechange;

      this.onreadystatechange = function () {
        if (self.readyState === 4 && self.status === 200 && self._xhrUrl) {
          try {
            // 【关键】从笔记详情 API 提取无水印视频链接
            if (
              self._xhrUrl.includes("/api/sns/") &&
              (self._xhrUrl.includes("/feed") || self._xhrUrl.includes("/note"))
            ) {
              const resp = JSON.parse(self.responseText);
              if (resp && resp.data) {
                // 遍历所有笔记数据
                const processNote = (note) => {
                  if (note && note.type === "video" && note.video) {
                    const v = note.video;

                    // 【最优】尝试 origin_video_key (真正无水印原始版本)
                    // 修改：不再主动覆盖 CACHED_VIDEO_URL，仅做日志记录，由 renderResourceGrid 决定是否使用
                    if (v.consumer && v.consumer.origin_video_key) {
                      const url = `https://sns-video-bd.xhscdn.com/${v.consumer.origin_video_key}`;
                      // CACHED_VIDEO_URL = url; // 禁用：避免覆盖真实抓取的 Stream 链接
                      console.log(
                        "[XHS-无水印] 发现 origin_video_key:",
                        url.substring(0, 60),
                      );
                      return;
                    }

                    // 【备选】H.264 master URL (通常带水印)
                    if (
                      v.media &&
                      v.media.stream &&
                      v.media.stream.h264 &&
                      v.media.stream.h264[0]
                    ) {
                      const url =
                        v.media.stream.h264[0].master_url ||
                        v.media.stream.h264[0].masterUrl;
                      // if (url && url.startsWith("http") && !CACHED_VIDEO_URL) {
                      //   CACHED_VIDEO_URL = url; // 禁用
                      // }
                    }
                    // 备选 H.265
                    else if (
                      v.media &&
                      v.media.stream &&
                      v.media.stream.h265 &&
                      v.media.stream.h265[0]
                    ) {
                      const url =
                        v.media.stream.h265[0].master_url ||
                        v.media.stream.h265[0].masterUrl;
                      // if (url && url.startsWith("http")) {
                      //   CACHED_VIDEO_URL = url; // 禁用
                      // }
                    }
                  }
                };

                if (resp.data.items) {
                  resp.data.items.forEach((item) => {
                    const note = item.note_card || item.note || item;
                    cacheNoteCard(note, item.id);
                    processNote(note);
                  });
                }
                if (resp.data.note) {
                  cacheNoteCard(resp.data.note);
                }
              }
            }
          } catch (e) {
            console.warn("[XHS] API 拦截异常:", e.message);
          }
        }
        if (origStateChange) origStateChange.call(this);
      };
      return originalSend.apply(this, [data, ...rest]);
    };
  })();

  // ==========================================
  // 0.1.1 Fetch 拦截 - 补充捕获
  // ==========================================
  (function () {
    const originalFetch = window.fetch;
    window.fetch = function (input, init) {
      let url = input;
      if (input instanceof Request) {
        url = input.url;
      }
      const result = originalFetch.apply(this, arguments);

      // 【新增】捕获页面通过 fetch 加载的笔记数据（与 XHR 拦截同源）
      if (
        url &&
        typeof url === "string" &&
        url.includes("/api/sns/") &&
        (url.includes("/feed") || url.includes("/note"))
      ) {
        result
          .then((res) => {
            try {
              res
                .clone()
                .text()
                .then((t) => {
                  try {
                    const resp = JSON.parse(t);
                    if (resp && resp.data) {
                      (resp.data.items || []).forEach((item) => {
                        const note = item.note_card || item.note || item;
                        cacheNoteCard(note, item.id);
                      });
                      if (resp.data.note) cacheNoteCard(resp.data.note);
                    }
                  } catch (e) {}
                })
                .catch(() => {});
            } catch (e) {}
          })
          .catch(() => {});
      }

      if (
        url &&
        typeof url === "string" &&
        url.includes("xhscdn.com") &&
        url.includes("/stream/") &&
        url.includes(".mp4")
      ) {
        CACHED_VIDEO_URL = url;
        console.log("[XHS-Stream-Fetch] 捕获无水印流地址:", url);
      }
      return result;
    };
  })();

  // ==========================================
  // 0.2 页面切换监听 - 保持数据清洁性
  // ==========================================
  function getPageKey(url) {
    try {
      const u = new URL(url);
      // 用户主页: /user/profile/xxx
      // 收藏夹: /user/profile/xxx/collect 或 /board/xxx
      // 搜索: /search_result
      // 首页: /
      return u.pathname;
    } catch {
      return url;
    }
  }

  function checkPageChange() {
    const newUrl = location.href;
    const oldKey = getPageKey(currentPageUrl);
    const newKey = getPageKey(newUrl);

    if (oldKey !== newKey) {
      console.log(`[XHS助手] 页面切换: ${oldKey} -> ${newKey}，清空数据`);
      GLOBAL_DATA.clear();
      updateCountUI();
      // 停止自动滚动（如果正在进行）
      if (isAutoScrolling) {
        const btn = document.getElementById("auto-scroll-btn");
        if (btn) btn.click(); // 触发停止
      }
      currentPageUrl = newUrl;

      // 检查是否为详情页，提示资源下载
      if (document.getElementById("xhs-ai-helper")) {
        checkResourcePage(newUrl);
      }
    }
  }

  function checkResourcePage(url) {
    const isDetail = url.includes("/explore/");
    const input = document.getElementById("res-url-input");
    const tab = document.querySelector('.ai-tab-item[data-tab="download"]');

    if (input && isDetail) {
      input.value = url;
      // 高亮提示
      if (tab) {
        const originText = tab.innerText;
        tab.style.color = "#e60023";
        tab.innerText = "资源下载";
        setTimeout(() => {
          tab.style.color = "";
          tab.innerText = "资源下载";
        }, 8000);
      }
    } else if (input) {
      input.value = "";
    }
  }

  // 监听 URL 变化（小红书是 SPA，用 History API）
  function setupPageChangeListener() {
    // 监听 popstate（浏览器前进/后退）
    window.addEventListener("popstate", checkPageChange);

    // 拦截 pushState 和 replaceState
    const originalPushState = history.pushState;
    const originalReplaceState = history.replaceState;

    history.pushState = function (...args) {
      originalPushState.apply(this, args);
      setTimeout(checkPageChange, 100);
    };

    history.replaceState = function (...args) {
      originalReplaceState.apply(this, args);
      setTimeout(checkPageChange, 100);
    };

    // 兜底：定时检查（防止某些情况遗漏）
    setInterval(checkPageChange, 2000);
  }

  // 启动页面切换监听
  setupPageChangeListener();

  // ==========================
  // 屏蔽登录弹窗逻辑
  // ==========================
  const SHIELD_LOGIN_KEY = "pc-shield-login-dialog";
  let _shieldLoginObserver = null;

  function addShieldLoginObserver() {
    if (_shieldLoginObserver) return;
    try {
      console.log("[XHS助手] 启用屏蔽登录弹窗");
      _shieldLoginObserver = new MutationObserver(() => {
        const closeBtn = document.querySelector(
          ".login-container .icon-btn-wrapper",
        );
        if (closeBtn) {
          try {
            closeBtn.click();
            console.log("[XHS助手] 登录弹窗出现，已自动关闭");
          } catch (e) {
            console.warn("[XHS助手] 关闭登录弹窗失败", e);
          }
        }
      });
      const attachObserver = () => {
        try {
          if (document.body)
            _shieldLoginObserver.observe(document.body, {
              childList: true,
              subtree: true,
            });
        } catch (e) {
          console.warn("[XHS助手] attach observer 失败", e);
        }
      };
      if (document.body) {
        attachObserver();
      } else {
        window.addEventListener("DOMContentLoaded", attachObserver, {
          once: true,
        });
      }
      // 立即尝试隐藏已存在的节点
      const loginNode = document.querySelector(".login-container");
      if (loginNode) {
        try {
          loginNode.style.display = "none";
        } catch (e) {}
      }
    } catch (e) {
      console.warn("[XHS助手] 添加屏蔽登录弹窗 observer 失败", e);
    }
  }

  function removeShieldLoginObserver() {
    if (_shieldLoginObserver) {
      try {
        _shieldLoginObserver.disconnect();
      } catch (e) {}
      _shieldLoginObserver = null;
      console.log("[XHS助手] 已停止屏蔽登录弹窗");
    }
  }

  // 读取设置并初始化
  function initShieldLoginFromSetting() {
    try {
      const enable = GM_getValue ? GM_getValue(SHIELD_LOGIN_KEY, false) : false;
      const $cb = document.getElementById("pc-shield-login-dialog");
      if ($cb) $cb.checked = !!enable;
      if (enable) addShieldLoginObserver();
      else removeShieldLoginObserver();
    } catch (e) {
      console.warn("[XHS助手] 读取屏蔽登录设置失败", e);
    }
  }

  // 立即在脚本启动时应用用户设置（若用户此前已开启，则生效）
  try {
    initShieldLoginFromSetting();
  } catch (e) {
    console.warn("[XHS助手] 启动时初始化屏蔽登录设置失败", e);
  }

  // 绑定设置面板复选框变化
  function bindShieldLoginSetting() {
    const $cb = document.getElementById("pc-shield-login-dialog");
    if (!$cb) return;
    $cb.addEventListener("change", (e) => {
      const checked = !!e.target.checked;
      try {
        if (GM_setValue) GM_setValue(SHIELD_LOGIN_KEY, checked);
      } catch (err) {
        console.warn("[XHS助手] 保存屏蔽登录设置失败", err);
      }
      if (checked) addShieldLoginObserver();
      else removeShieldLoginObserver();
    });
  }

  // ==========================================
  // 1. API 拦截器 (Hook XHR)
  // ==========================================
  function hookXHR() {
    const originalOpen = XMLHttpRequest.prototype.open;
    const originalSend = XMLHttpRequest.prototype.send;

    XMLHttpRequest.prototype.open = function (method, url) {
      this._url = url;
      return originalOpen.apply(this, arguments);
    };

    XMLHttpRequest.prototype.send = function (body) {
      this.addEventListener("load", function () {
        // 监听特定 API 接口
        // /api/sns/web/v1/user_posted -> 用户主页（笔记）
        // /api/sns/web/v2/note/collect/page -> 收藏夹
        // /api/sns/web/v1/note/like/page -> 点赞列表
        // /api/sns/web/v1/user/like -> 点赞列表（备选）
        // /api/sns/web/v1/homefeed -> 首页推荐/搜索
        // /api/sns/web/v1/search/notes -> 搜索结果
        if (
          this._url &&
          (this._url.includes("/api/sns/web/v1/user_posted") ||
            this._url.includes("/api/sns/web/v2/note/collect/page") ||
            this._url.includes("/api/sns/web/v1/note/like/page") ||
            this._url.includes("/api/sns/web/v1/user/like") ||
            this._url.includes("/api/sns/web/v1/homefeed") ||
            this._url.includes("/api/sns/web/v1/search/notes"))
        ) {
          try {
            const res = JSON.parse(this.responseText);
            if (res.data && (res.data.notes || res.data.items)) {
              const list = res.data.notes || res.data.items || [];
              processNotes(list);
            }
          } catch (e) {
            console.error("小红书助手: 解析API数据失败", e);
          }
        }
      });
      return originalSend.apply(this, arguments);
    };
  }

  // 处理并存储笔记数据
  function processNotes(notes) {
    let newCount = 0;
    // 兼容搜索结果页结构（如 /api/sns/web/v1/search/notes）及 首页推荐（/api/sns/web/v1/homefeed）
    // notes 可能是 [{note: {...}}, ...] 或直接 [{...}]
    notes.forEach((raw) => {
      let note = raw;
      // 搜索结果页结构：{note: {...}, ...}
      if (raw && raw.note && typeof raw.note === "object") {
        note = raw.note;
      }
      // 兼容 note_card 层
      if (note && note.note_card) {
        note = note.note_card;
      }
      // 兼容 note_info 层
      if (note && note.note_info) {
        note = note.note_info;
      }
      // 兼容 item 层（部分搜索结果）
      if (note && note.item) {
        note = note.item;
      }
      // 兼容 feed_note 层（部分推荐/搜索结果）
      if (note && note.feed_note) {
        note = note.feed_note;
      }

      // 统一提取字段
      // 优先从深层对象取，取不到尝试从原始对象取（防止层级下钻导致外层属性丢失）
      const id =
        note.id || note.note_id || note.noteId || raw.id || raw.note_id;
      if (!id) return;

      const token = note.xsec_token || raw.xsec_token || "";
      let link = `https://www.xiaohongshu.com/explore/${id}`;
      if (token) {
        link += `?xsec_token=${token}&xsec_source=pc_user`;
      }

      const title =
        note.title ||
        note.display_title ||
        note.desc ||
        raw.title ||
        raw.display_title ||
        "无标题";

      const type = note.type || raw.type || "normal";

      const user = note.user || raw.user || {};
      const authorName =
        user.nickname ||
        user.name ||
        note.author ||
        raw.author ||
        (raw.user && raw.user.nickname) ||
        "未知作者";

      const likes =
        note.likes ||
        note.liked_count ||
        note.like_count ||
        (note.interact_info && note.interact_info.liked_count) ||
        raw.likes ||
        raw.liked_count ||
        (raw.interact_info && raw.interact_info.liked_count) ||
        0;

      const coverUrl =
        (note.cover && note.cover.url_default) ||
        (note.images_list && note.images_list[0] && note.images_list[0].url) ||
        note.cover_url ||
        (raw.cover && raw.cover.url_default) ||
        raw.cover_url ||
        "";

      if (!GLOBAL_DATA.has(id)) {
        GLOBAL_DATA.set(id, {
          笔记ID: id,
          标题: title,
          链接: link,
          作者: authorName,
          点赞数: likes,
          封面图: coverUrl,
          类型: type,
          xsec_token: token,
        });
        newCount++;
      }
    });

    // 更新 UI 计数
    updateCountUI();
  }

  // 启动 Hook
  hookXHR();

  // ==========================================
  // 2. 默认模板 (AI部分 - 保持不变)
  // ==========================================
  const DEFAULT_TEMPLATES = [
    {
      id: "novel_default",
      name: "小说推文 (情绪爆款)",
      desc1: "小说名称",
      desc2: "精彩片段/剧情",
      placeholder1: "例如：重生之将门毒后",
      placeholder2: "粘贴这一章的剧情...",
      system:
        "你是一个小红书推文博主，风格非常情绪化、激动，喜欢用'啊啊啊'、'高开暴走'、'Top1'等词汇。",
      prompt: `请模仿以下风格推荐小说《{{title}}》。\n【参考风格】："啊啊啊高开暴走！！..."\n【小说内容】：{{content}}\n要求：标题带悬念含Emoji，正文情绪化分段，JSON格式输出 {"title": "...", "content": "..."}`,
    },
    {
      id: "product_default",
      name: "好物种草 (痛点直击)",
      desc1: "产品名称",
      desc2: "核心卖点/痛点",
      placeholder1: "例如：戴森吹风机",
      placeholder2: "痛点：头发干枯... 体验：柔顺...",
      system:
        "你是一个小红书金牌种草官，擅长挖掘痛点和场景，说话就像闺蜜聊天。",
      prompt: `请为产品【{{title}}】写一篇种草笔记。\n信息：{{content}}\n要求：痛点+场景+Emoji，JSON格式输出 {"title": "...", "content": "..."}`,
    },
  ];

  // ==========================================
  // 3. 核心样式
  // ==========================================
  /* ==========================================
   * 设计系统：Pinterest Marketing System
   * - Pinterest Red (#e60023) 仅用于主 CTA / 激活态 / 品牌标识
   * - 圆角体系：16px 组件 / 32px 大卡片 / 9999px 胶囊
   * - 字体：系统字体栈（Pin Sans 不可用，不引入第三方字体）
   * - 奶油中性色 Chrome，内容卡片零阴影（仅面板模态层有阴影）
   * ========================================== */
  const UI_CSS = `
        /* 覆盖全局样式 */
        #xhs-ai-helper * { box-sizing: border-box; }
        @keyframes slideIn { from { opacity:0; transform: translateY(10px); } to { opacity:1; transform: translateY(0); } }
        @keyframes pulseInk { 0% { box-shadow: 0 0 0 0 rgba(0,0,0,0.25); } 70% { box-shadow: 0 0 0 10px rgba(0,0,0,0); } 100% { box-shadow: 0 0 0 0 rgba(0,0,0,0); } }
        /* 展开/收起动画：只动合成属性（opacity/transform），不触发布局重排，避免卡顿 */
        @keyframes panelIn { from { opacity: 0; transform: translateX(16px); } to { opacity: 1; transform: translateX(0); } }
        @keyframes dotIn { from { opacity: 0; transform: scale(0.5); } to { opacity: 1; transform: scale(1); } }

        #xhs-ai-helper {
            /* ---- 色彩 Token ---- */
            --red: #e60023;               /* Pinterest Red：主CTA/激活态/品牌标识专用 */
            --red-pressed: #cc001f;
            --canvas: #ffffff;
            --surface-soft: #fbfbf9;
            --surface-card: #f6f6f3;      /* 暖奶油卡片面 */
            --secondary-bg: #e5e5e0;
            --secondary-pressed: #c8c8c1;
            --hairline: #dadad3;
            --ink: #000000;
            --body: #33332e;
            --mute: #62625b;
            --ash: #91918c;
            --error: #9e0a0a;
            --focus-outer: #435ee5;
            --font: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, "PingFang SC", "Hiragino Sans GB", "Microsoft YaHei", sans-serif;

            position: fixed; top: 0; right: 0; width: 400px; height: 100vh;
            background: var(--canvas);
            border-left: 1px solid var(--hairline);
            border-radius: 32px 0 0 32px;
            box-shadow: 0 16px 32px rgba(0,0,0,0.12);
            z-index: 999999;
            font-family: var(--font);
            display: flex; flex-direction: column;
            color: var(--body); font-size: 14px;
        }
        /* 展开入场：淡入 + 右滑入（纯合成动画） */
        #xhs-ai-helper:not(.minimized) { animation: panelIn 0.2s cubic-bezier(0.2, 0.8, 0.2, 1) both; }
        /* ---- 标题栏（heading-lg）---- */
        .drag-handle { padding: 14px 20px; cursor: default; display: flex; justify-content: space-between; align-items: center; border-bottom: 1px solid var(--hairline); user-select: none; }
        .ai-brand { font-weight: 600; font-size: 22px; color: var(--ink); }
        /* ---- 最小化：品牌红圆形标识 ---- */
        #xhs-ai-helper.minimized { width: 36px; height: 36px; border-radius: 9999px; overflow: hidden; cursor: pointer; background: var(--red); box-shadow: 0 4px 12px rgba(230,0,35,0.35); border-left: none; }
        /* 展开态隐藏最小化图标，避免左上角泄漏 */
        #xhs-ai-helper .minimized-icon { display: none; }
        #xhs-ai-helper.minimized .minimized-icon { display: flex; width: 100%; height: 100%; align-items: center; justify-content: center; color: #fff; font-size: 20px; font-weight: 700; animation: dotIn 0.18s ease both; }
        #xhs-ai-helper.minimized .ai-main-wrapper { display: none; }
        /* 展开态：主区域撑满面板高度，作为flex列容器让内容区可内部滚动 */
        #xhs-ai-helper .ai-main-wrapper { display: flex; flex-direction: column; flex: 1; min-height: 0; }

        /* ---- 图标按钮（button-icon-circular）---- */
        .ai-icon-btn { width: 40px; height: 40px; border-radius: 9999px; border: none; background: var(--surface-card); color: var(--ink); font-size: 18px; font-weight: 700; cursor: pointer; display: flex; align-items: center; justify-content: center; flex-shrink: 0; transition: background 0.2s; }
        .ai-icon-btn:hover { background: var(--secondary-bg); }

        /* ---- Tab 筹码行（filter-chip / filter-chip-active）---- */
        .ai-tabs {
            display: flex; gap: 8px; padding: 12px 20px;
            background: var(--canvas);
            border-bottom: 1px solid var(--hairline); user-select: none;
            overflow-x: auto; white-space: nowrap;
            scrollbar-width: none; -ms-overflow-style: none;
        }
        .ai-tabs::-webkit-scrollbar { display: none; }
        .ai-tab-item {
            padding: 8px 14px; font-size: 12px; font-weight: 700; color: var(--ink);
            cursor: pointer; border-radius: 9999px; transition: background 0.2s, color 0.2s;
            background: var(--surface-card); flex-shrink: 0;
        }
        .ai-tab-item:hover { background: var(--secondary-bg); }
        .ai-tab-item.active { background: var(--ink); color: var(--canvas); }

        /* ---- 内容区 ---- */
        .ai-content-body { padding: 16px; flex: 1; min-height: 0; overflow-y: auto; background: var(--canvas); }
        .ai-content-body::-webkit-scrollbar { width: 8px; }
        .ai-content-body::-webkit-scrollbar-thumb { background: var(--secondary-pressed); border-radius: 9999px; }
        .ai-content-body::-webkit-scrollbar-track { background: transparent; }
        .tab-panel { display: none; }
        .tab-panel.active { display: block; animation: slideIn 0.2s; }

        /* ---- 输入（text-input / text-input-focused 双环聚焦）---- */
        .ai-input, .ai-textarea, .ai-select {
            width: 100%; padding: 11px 15px; border: 1px solid var(--ash);
            border-radius: 16px; margin-bottom: 10px; background: var(--canvas);
            color: var(--ink); font-family: var(--font); font-size: 14px;
            transition: border-color 0.2s, box-shadow 0.2s;
        }
        .ai-input::placeholder, .ai-textarea::placeholder { color: var(--ash); }
        .ai-input:focus, .ai-textarea:focus, .ai-select:focus {
            outline: none; border-color: var(--ink);
            box-shadow: 0 0 0 4px var(--focus-outer);
        }
        .ai-textarea { height: 80px; resize: vertical; line-height: 1.5; }

        /* ---- 按钮（button-primary / button-secondary）---- */
        .ai-btn {
            width: 100%; padding: 10px 14px; min-height: 40px;
            background: var(--red); color: #fff; border: none; border-radius: 16px;
            cursor: pointer; font-family: var(--font); font-size: 14px; font-weight: 700; line-height: 1.2;
            margin-top: 5px; transition: background 0.2s;
        }
        .ai-btn:hover { background: var(--red-pressed); }
        .ai-btn:disabled { background: var(--surface-card); color: var(--ash); cursor: default; }
        .ai-btn.secondary { background: var(--secondary-bg); color: var(--ink); }
        .ai-btn.secondary:hover { background: var(--secondary-pressed); }
        .ai-btn.secondary:disabled { background: var(--surface-card); color: var(--ash); }
        .ai-btn.danger { background: var(--secondary-bg); color: var(--error); }
        .ai-btn.danger:hover { background: var(--secondary-pressed); }
        .ai-btn.scrolling { background: var(--ink); color: #fff; animation: pulseInk 2s infinite; }

        /* ---- 卡片（扁平奶油面，零阴影）---- */
        .data-card { background: var(--surface-card); padding: 16px; border-radius: 16px; margin-bottom: 12px; color: var(--body); }
        .export-tip { font-size: 12px; color: var(--mute); margin-top: 12px; line-height: 1.6; background: var(--surface-card); padding: 12px 16px; border-radius: 16px; }

        /* ---- 文件上传区 ---- */
        .file-upload-label {
            display: flex; align-items: center; justify-content: center; gap: 8px;
            background: var(--surface-soft); border: 1px dashed var(--hairline); border-radius: 16px;
            padding: 15px; cursor: pointer; color: var(--mute); font-size: 13px;
            transition: border-color 0.2s, color 0.2s, background 0.2s; margin-bottom: 5px;
        }
        .file-upload-label:hover { border-color: var(--ink); color: var(--ink); }
        .analysis-result-box {
            margin-top: 10px; font-size: 13px; line-height: 1.6;
            background: var(--canvas); padding: 12px 16px; border-radius: 16px;
            border: 1px solid var(--hairline); max-height: 250px; overflow-y: auto;
            color: var(--body); white-space: pre-wrap; display: none;
        }
        .ai-compact-box { background: var(--surface-card); padding: 12px 16px; border-radius: 16px; }

        /* ---- 资源网格（pin-card：图即卡片）---- */
        .res-grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(80px, 1fr)); gap: 8px; margin-top: 10px; }
        .res-item { position: relative; aspect-ratio: 1; border-radius: 16px; overflow: hidden; background: var(--surface-card); border: 2px solid transparent; cursor: pointer; }
        .res-item img, .res-item video { width: 100%; height: 100%; object-fit: cover; display: block; }
        .res-item .res-type { position: absolute; top: 8px; right: 8px; font-size: 12px; font-weight: 700; background: var(--canvas); color: var(--ink); padding: 4px 10px; border-radius: 9999px; }
        .res-item:hover::after { content: ''; position: absolute; top: 0; left: 0; right: 0; bottom: 0; background: rgba(0,0,0,0.06); }
        .res-item.selected { border-color: var(--ink); }
        .res-item .res-check { position: absolute; top: 6px; left: 6px; z-index: 2; width: 20px; height: 20px; background: var(--canvas); border: 1px solid var(--hairline); border-radius: 9999px; display: flex; align-items: center; justify-content: center; opacity: 0.9; }
        .res-item.selected .res-check { background: var(--ink); border-color: var(--ink); opacity: 1; }

        /* ---- 导出字段筹码（选中反转）---- */
        .export-field-container { display: grid; grid-template-columns: repeat(4, 1fr); gap: 8px; font-size: 12px; }
        .export-field-item { background: var(--canvas); border: 1px solid var(--hairline); border-radius: 9999px; padding: 8px 4px; text-align: center; cursor: pointer; user-select: none; transition: background 0.2s, color 0.2s, border-color 0.2s; color: var(--body); font-weight: 500; }
        .export-field-item:hover { border-color: var(--ink); }
        .export-field-item.selected { background: var(--ink); border-color: var(--ink); color: #fff; font-weight: 700; }

        /* ---- 开关 ---- */
        .ai-switch { position: relative; display: inline-block; width: 44px; height: 24px; flex-shrink: 0; }
        .ai-switch input { opacity: 0; width: 0; height: 0; }
        .ai-slider { position: absolute; cursor: pointer; top: 0; left: 0; right: 0; bottom: 0; background-color: var(--secondary-bg); transition: .3s; border-radius: 9999px; }
        .ai-slider:before { position: absolute; content: ""; height: 18px; width: 18px; left: 3px; bottom: 3px; background-color: white; transition: .3s; border-radius: 9999px; box-shadow: 0 2px 4px rgba(0,0,0,0.2); }
        input:checked + .ai-slider { background-color: var(--red); }
        input:checked + .ai-slider:before { transform: translateX(20px); }

        /* 小屏幕/副屏适配 */
        @media screen and (max-width: 500px), screen and (max-height: 600px) {
            #xhs-ai-helper {
                width: calc(100vw - 20px) !important;
                max-width: 360px;
                right: 0 !important;
                left: auto !important;
                top: 0 !important;
                height: 100vh !important;
                max-height: 100vh;
            }
            #xhs-ai-helper .ai-content-body { max-height: none; }
            /* 最小化图标不受媒体查询宽度影响 */
            #xhs-ai-helper.minimized {
                width: 36px !important;
                height: 36px !important;
            }
        }
        @media screen and (max-width: 400px) {
            #xhs-ai-helper {
                width: calc(100vw - 10px) !important;
                right: 0 !important;
                font-size: 13px;
            }
            #xhs-ai-helper.minimized {
                width: 36px !important;
                height: 36px !important;
            }
            #xhs-ai-helper .ai-tabs { padding: 12px 16px; }
            #xhs-ai-helper .ai-tab-item { padding: 8px 12px; font-size: 12px; }
            #xhs-ai-helper .ai-content-body { padding: 12px; }
            #xhs-ai-helper .drag-handle { padding: 12px 16px; }
            #xhs-ai-helper .ai-brand { font-size: 18px; }
        }
    `;
  GM_addStyle(UI_CSS);

  // ==========================================
  // 4. UI 构建
  // ==========================================
  let templates = [];
  function loadTemplates() {
    const stored = GM_getValue("user_templates", null);
    templates = stored
      ? JSON.parse(stored)
      : JSON.parse(JSON.stringify(DEFAULT_TEMPLATES));
  }

  function saveTemplates() {
    GM_setValue("user_templates", JSON.stringify(templates));
  }

  function createUI() {
    if (document.getElementById("xhs-ai-helper")) return;
    loadTemplates();

    const div = document.createElement("div");
    div.id = "xhs-ai-helper";
    div.classList.add("minimized"); // 默认不展开，仅显示可拖拽的悬浮图标
    div.innerHTML = `
            <div class="minimized-icon"></div>
            <div class="ai-main-wrapper">
                <div class="drag-handle">
                    <div class="ai-brand">小红书助手</div>
                    <button id="minimize-btn" class="ai-icon-btn" title="收起面板">_</button>
                </div>
                <div class="ai-tabs">
                    <div class="ai-tab-item active" data-tab="data">数据导出</div>
                    <div class="ai-tab-item" data-tab="write">AI创作</div>
                    <!-- <div class="ai-tab-item" data-tab="marketing">营销工具</div> -->
                    <div class="ai-tab-item" data-tab="analysis">AI分析</div>
                    <div class="ai-tab-item" data-tab="download">资源下载</div>
                    <div class="ai-tab-item" data-tab="settings">设置</div>
                </div>
                <div class="ai-content-body">
                    <div id="panel-data" class="tab-panel active">
                        <div class="data-card">
                            <div style="display:flex;justify-content:space-between;font-size:12px;margin-bottom:5px;">
                                <span>当前已捕获：<b id="page-obj-count" style="color:#000000;font-size:14px;font-weight:700;">0</b> 条</span>
                                <span id="scroll-status" style="color:#91918c;">静止中</span>
                            </div>
                            <div style="font-size:11px;color:#62625b;margin-bottom:8px;">提示：向下滑动页面，数据会自动增加。</div>

                            <button id="auto-scroll-btn" class="ai-btn secondary" style="padding:8px;">自动滚动加载全部</button>
                        </div>

                        <div style="display:flex;gap:5px; flex-wrap:wrap; align-items: center;">
                            <select id="export-format" class="ai-select" style="margin-bottom:0; width:auto; flex:1;">
                                <option value="xlsx_embed">Excel (永久嵌入图)</option>
                                <option value="csv">CSV (纯文本)</option>
                                <option value="json">JSON</option>
                                <option value="md">Markdown (MD)</option>
                                <option value="html">HTML (网页还原)</option>
                            </select>
                            <button id="clean-data-btn" class="ai-btn secondary" style="width:auto;margin-top:0;">清空</button>
                        </div>
                        
                        <div id="split-setting-area" style="display:flex; font-size:12px; margin-top:5px; align-items:center; gap:8px; background:#f6f6f3; padding:8px 12px; border-radius:16px; border:1px solid #dadad3;">
                                <span style="font-weight:700;color:#33332e;">分包设置：</span>
                                <label style="cursor:pointer;"><input type="checkbox" id="enable-split" checked> 启用分包</label>
                                <div style="display:flex;align-items:center;">
                                    每包 <input type="number" id="split-size" value="200" style="width:50px; margin:0 4px; padding:4px; text-align:center; border:1px solid #dadad3; border-radius:8px;"> 条
                                </div>
                        </div>
                        
                        <div class="ai-compact-box" style="margin-top:8px; padding: 6px; background:#f6f6f3; border-radius:16px;">
                             <label style="font-size:12px; font-weight:bold; display:block; margin-bottom:6px; color:#33332e;">导出字段 (点击切换):</label>
                             <div id="export-field-container" class="export-field-container">
                                <div class="export-field-item selected" data-value="笔记ID">笔记ID</div>
                                <div class="export-field-item selected" data-value="链接">链接</div>
                                <div class="export-field-item selected" data-value="封面图">封面</div>
                                <div class="export-field-item selected" data-value="标题">标题</div>
                                <div class="export-field-item selected" data-value="作者">作者</div>
                                <div class="export-field-item selected" data-value="点赞数">点赞数</div>
                                <div class="export-field-item selected" data-value="类型">类型</div>
                             </div>
                        </div>

                        <button id="export-btn" class="ai-btn" style="margin-top:10px;">导出所有捕获数据</button>

                        <div class="export-tip">
                            <b>已启用 API 拦截模式</b><br>
                            导出的链接将包含 <code>xsec_token</code>，确保能在浏览器中直接访问，不会出现"笔记不存在"。<br>
                            支持收藏夹、个人主页、搜索结果页。
                        </div>
                    </div>


                    <div id="panel-write" class="tab-panel">
                        <div class="ai-input-group">
                             <select id="template-select" class="ai-select"></select>
                             <input id="input-1" class="ai-input" placeholder="主题">
                             <textarea id="input-2" class="ai-textarea" placeholder="内容详情"></textarea>

                            <div style="display:flex;gap:6px;margin-bottom:10px;">
                              <button id="template-manage-toggle" class="ai-btn secondary" style="margin:0;">管理模版</button>
                              <button id="persona-gen-btn" class="ai-btn secondary" style="margin:0;">🧠 AI生成人设</button>
                            </div>

                            <div id="template-manage-panel" style="display:none;padding:10px;border:1px solid #dadad3;border-radius:16px;background:#f6f6f3;">
                              <div style="display:flex;gap:8px;align-items:center;margin-bottom:8px;">
                               <select id="template-manage-select" class="ai-select" style="margin:0;"></select>
                               <button id="template-new-btn" class="ai-btn secondary" style="margin:0;width:120px;">新增模版</button>
                              </div>
                              <input id="tpl-name" class="ai-input" placeholder="模版名称 (例：好物种草)" />
                              <input id="tpl-desc1" class="ai-input" placeholder="第一个输入提示 (例：产品名称)" />
                              <input id="tpl-desc2" class="ai-input" placeholder="第二个输入提示 (例：核心卖点)" />
                              <textarea id="tpl-system" class="ai-textarea" placeholder="System 提示 (AI角色)" style="height:70px;"></textarea>
                              <textarea id="tpl-prompt" class="ai-textarea" placeholder="Prompt 模版，使用 {{title}} / {{content}} 占位符" style="height:90px;"></textarea>
                              <textarea id="tpl-persona" class="ai-textarea" placeholder="输入模版人设/要求，让AI自动生成 system 与 prompt" style="height:60px;"></textarea>
                              <div style="display:flex;gap:8px;">
                                <button id="template-save-btn" class="ai-btn" style="margin:0;">保存修改</button>
                                <button id="template-save-new-btn" class="ai-btn secondary" style="margin:0;">另存为新模版</button>
                                <button id="template-delete-btn" class="ai-btn secondary danger" style="margin:0;">删除</button>
                              </div>
                              <div id="template-manage-status" style="margin-top:6px;font-size:12px;color:#62625b;"></div>
                            </div>
                            
                            <button id="ai-gen-btn" class="ai-btn" style="margin-top:10px;">生成文案</button>
                            <div id="ai-status" style="text-align:center;font-size:12px;margin-top:5px;color:#91918c;"></div>
                        </div>
                    </div>

                    <!--
                    <div id="panel-marketing" class="tab-panel">
                        <div class="data-card">
                             <div style="font-size: 14px; font-weight: bold; margin-bottom: 5px; color:#000000;">关键词搜索</div>
                             <div style="display:flex; gap:5px;">
                                <input id="mkt-keyword" class="ai-input" placeholder="输入搜索关键词" style="margin-bottom:0;">
                                <button id="mkt-search-btn" class="ai-btn" style="width:auto; margin:0; background:#ff2442;">搜索</button>
                             </div>
                             <div style="font-size: 11px; color:#91918c; margin-top:5px;">将会跳转到搜索页并自动准备抓取数据。</div>
                        </div>

                        <div class="data-card" style="border-left: 4px solid #9c27b0;">
                             <div style="font-size: 14px; font-weight: bold; margin-bottom: 5px; color:#000000;">批量评论导出</div>
                             <div style="font-size: 11px; color:#62625b; margin-bottom:10px;">基于当前已捕获的笔记列表，批量抓取评论。</div>
                             
                             <div class="ai-compact-box" style="background:#f6f6f3; border-color:#e1d6e8;">
                                <div style="font-size:12px; margin-bottom:5px;">
                                    当前待处理笔记：<span id="mkt-note-count" style="font-weight:bold; color:#9c27b0;">0</span> 篇
                                </div>
                                
                                <label style="display:block; font-size:12px; margin-bottom:3px; color:#33332e;">点赞数过滤 (仅导出点赞大于 N 的评论):</label>
                                <input id="mkt-filter-likes" class="ai-input" type="number" value="10" placeholder="0 表示不限" style="height:30px; margin-bottom:8px;">
                                
                                <label style="display:block; font-size:12px; margin-bottom:3px; color:#33332e;">抓取范围 (每篇笔记):</label>
                                <select id="mkt-crawl-scope" class="ai-select" style="height:30px; padding:2px 8px; margin-bottom:8px;">
                                    <option value="initial">仅首屏热评 (速度快，免验证)</option>
                                </select>
                                
                                <button id="mkt-run-btn" class="ai-btn" style="background:#9c27b0;">开始批量抓取评论</button>
                             </div>

                             <div id="mkt-progress-area" style="display:none; margin-top:10px;">
                                <div style="background:#eee; height:6px; border-radius:3px; overflow:hidden;">
                                    <div id="mkt-progress-bar" style="width:0%; height:100%; background:#9c27b0; transition: width 0.3s;"></div>
                                </div>
                                <div id="mkt-status-text" style="font-size:11px; color:#62625b; text-align:center; margin-top:4px;">正在准备...</div>
                             </div>
                             
                             <div id="mkt-result-area" style="display:none; margin-top:10px; border-top:1px dashed #ddd; padding-top:10px;">
                                <div style="font-size:12px; margin-bottom:5px;">
                                    已采集评论：<b id="mkt-result-count" style="color:#d32f2f;">0</b> 条
                                </div>
                                <button id="mkt-export-btn" class="ai-btn" style="background:#00b85c;">导出评论数据 (Excel)</button>
                             </div>
                        </div>
                    </div>
                    -->
                    
                    <div id="panel-analysis" class="tab-panel">
                        <!-- Section 1: 智能分析 -->
                        <div class="data-card">
                           <div style="font-size: 14px; font-weight: bold; margin-bottom: 5px; color:#000000;">智能总结与分析</div>
                           <div style="font-size: 11px; color:#62625b; margin-bottom:10px;">上传导出的 CSV/JSON，让 AI 分析趋势。</div>
                           
                           <input type="file" id="analysis-file-input" accept=".csv,.json" style="display:none;" />
                           <label for="analysis-file-input" class="file-upload-label" id="analysis-file-label">
                               点击选择或拖拽文件 (CSV/JSON)
                           </label>
                           <div id="analysis-file-name" style="font-size:11px; color:#000000; margin:5px 0; padding-left: 5px; display:none;"></div>

                           <button id="analysis-summary-btn" class="ai-btn" style="margin-top:8px;">🧠 开始智能分析</button>
                           <div id="analysis-result" class="analysis-result-box"></div>
                        </div>

                        <!-- Section 2: 智能分类 -->
                        <div class="data-card" style="margin-top:15px;">
                           <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom: 10px;">
                              <div style="font-size: 14px; font-weight: bold; color:#000000;">智能分类重构</div>
                              <button id="analysis-config-btn" class="ai-btn secondary" style="width:auto; padding:4px 8px; font-size:12px; margin:0;" title="API 设置">配置 API</button>
                           </div>
                           <div style="font-size: 11px; color:#62625b; margin: 0 0 10px;">AI 自动分类数据并生成新文件。</div>
                           
                           <div class="ai-compact-box">
                               <label style="font-size:12px;color:#91918c;display:block;margin-bottom:4px;">自定义分类 (可选, 逗号分隔)</label>
                               <textarea id="analysis-categories" class="ai-textarea" style="height:40px; margin-bottom:8px; border:1px solid #dadad3; background:#ffffff; padding:5px; font-size: 12px;" placeholder="美妆, 穿搭, 美食..."></textarea>
                               
                               <div style="display:flex; justify-content:space-between; align-items:center;">
                                   <div style="font-size:12px;color:#62625b; display:flex; align-items:center; gap:5px;">
                                      导出格式: 
                                      <select id="analysis-export-format" class="ai-select" style="width:auto; padding:3px 6px; margin:0; height:auto; background:#fff; border-color:#dadad3;">
                                          <option value="xls">Excel</option> // 默认Excel
                                          <option value="csv">CSV</option>
                                          <option value="json">JSON</option>
                                      </select>
                                   </div>
                                   <button id="analysis-classify-btn" class="ai-btn" style="width:auto; padding:8px 18px; margin:0;">开始分类</button>
                               </div>
                           </div>
                           
                           <div id="classify-status" style="margin-top:8px; font-size:12px; color:#62625b;"></div>
                           <div id="api-limit-tip" style="display:none; margin-top:8px; font-size:11px; color:#9e0a0a; background:#f6f6f3; padding:8px; border-radius:16px; line-height: 1.4;">
                              检测到未分类数据。这通常是因为 API 速率限制或额度不足。<br>建议：<br>1. 检查 API Key 额度。<br>2. 更换更稳定的模型 (如 gpt-3.5/4)。<br>3. 每次处理需要一定时间，请耐心等待。
                           </div>
                        </div>
                    </div>
                
                    <div id="panel-download" class="tab-panel">
                        <div class="data-card">
                             <div style="font-size: 14px; font-weight: bold; margin-bottom: 5px; color:#000000;">笔记资源提取</div>
                             <div style="font-size: 11px; color:#62625b; margin-bottom:10px;">输入笔记链接，或者在详情页直接使用。（支持无水印图片/视频下载）</div>
                             
                             <input id="res-url-input" class="ai-input" placeholder="粘贴小红书笔记链接 (或者留空自动检测当前页)" />
                             <button id="res-fetch-btn" class="ai-btn">提取资源</button>
                        </div>

                        <div id="res-result-area" class="data-card" style="display:none;">
                             <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:8px;">
                                <div style="font-size: 13px; font-weight: bold;">资源列表 <span id="res-count" style="color:#62625b;font-weight:normal;font-size:11px;"></span></div>
                                <select id="res-quality-select" class="ai-select" style="width:auto;margin:0;padding:2px 5px;height:24px;font-size:11px;">
                                    <option value="no_wm">高清无水印 (推荐)</option>
                                    <option value="best_wm">最高画质 (可能带水印)</option>
                                </select>
                             </div>
                             
                             <div class="ai-compact-box" style="margin-bottom:8px;padding:5px;background:#f6f6f3;">
                                <div style="font-size: 11px; display:flex; justify-content:space-between; align-items:center;">
                                   <label style="cursor:pointer;"><input type="checkbox" id="res-select-all" checked> 全选</label>
                                   <span style="color:#62625b;">点击图片可预览/取消</span>
                                </div>
                             </div>
                             
                             <div id="res-grid" class="res-grid"></div>
                             
                             <button id="res-download-btn" class="ai-btn" style="margin-top:10px;">批量下载选中资源</button>
                             <div style="margin-top:4px; font-size:11px; color:#91918c; text-align:center;">资源将保存到以笔记标题命名的文件夹中</div>
                             <div id="res-status" style="margin-top:5px; font-size:11px; color:#62625b; text-align:center;"></div>
                        </div>
                    </div>
                
                    <div id="panel-settings" class="tab-panel">
                        <div class="data-card">
                             <div style="font-size: 13px; font-weight: bold; margin-bottom: 8px;">全局 API 设置</div>
                             <div style="font-size: 11px; color:#62625b; margin-bottom:8px;">此处的配置将应用于 AI 创作、分析和分类功能。</div>

                             <div id="config-area" style="margin-bottom:10px;background:#ffffff;padding:12px;border-radius:16px;border:1px solid #dadad3;">
                                <div style="display:flex;gap:5px;margin-bottom:8px;">
                                    <select id="api-config-select" class="ai-select" style="margin:0;flex:1;"></select>
                                    <button id="api-config-add" class="ai-btn secondary" style="width:40px;height:40px;min-height:40px;margin:0;padding:0;border-radius:9999px;font-size:16px;" title="新增配置">+</button>
                                    <button id="api-config-del" class="ai-btn secondary" style="width:40px;height:40px;min-height:40px;margin:0;padding:0;border-radius:9999px;font-size:16px;" title="删除配置">×</button>
                                </div>
                                
                                <input id="api-base-url" class="ai-input" placeholder="Base URL" style="margin-bottom:5px;">
                                <input id="api-key" type="password" class="ai-input" placeholder="API Key" style="margin-bottom:5px;">
                                
                                <div style="display:flex;gap:5px;">
                                    <input id="api-model" class="ai-input" placeholder="Model (e.g. gpt-3.5-turbo)" style="margin-bottom:0;flex:1;">
                                    <button id="api-model-fetch-btn" class="ai-btn secondary" style="width:auto;margin:0;padding:4px 8px;font-size:12px;" title="尝试获取模型列表">获取</button>
                                </div>
                                <select id="api-model-select" class="ai-select" style="display:none;margin-top:5px;"></select>
                             </div>
                               <div style="margin-top:8px;padding:12px;background:#fff;border-radius:16px;border:1px solid #dadad3; display:flex; justify-content:space-between; align-items:center;">
                                  <div>
                                     <div style="font-size:14px; font-weight:600; color:#000000; margin-bottom:2px;">屏蔽登录弹窗</div>
                                     <div style="font-size:11px; color:#91918c;">自动关闭烦人的登录提示，享受纯净浏览</div>
                                  </div>
                                  <label class="ai-switch" title="点击开启/关闭">
                                    <input type="checkbox" id="pc-shield-login-dialog">
                                    <span class="ai-slider"></span>
                                  </label>
                               </div>
                        </div>
                    </div>
                </div>
            </div>
        `;
    document.body.appendChild(div);

    // 绑定事件
    bindDrag(div);

    // 最小化按钮：收起为悬浮图标，并恢复上次拖拽保存的位置
    // （点击图标展开面板的逻辑在 bindDrag 中通过"点击/拖拽"区分实现）
    div.querySelector("#minimize-btn").onclick = (e) => {
      e.stopPropagation();
      togglePanel(div, "minimize");
    };

    const tabs = div.querySelectorAll(".ai-tab-item");
    const tabsContainer = div.querySelector(".ai-tabs");
    tabs.forEach(
      (t) =>
        (t.onclick = () => {
          tabs.forEach((x) => x.classList.remove("active"));
          t.classList.add("active");
          div
            .querySelectorAll(".tab-panel")
            .forEach((p) => p.classList.remove("active"));
          div.querySelector("#panel-" + t.dataset.tab).classList.add("active");

          // 选中 Tab 自动居中滚动逻辑
          if (tabsContainer) {
            const targetScroll =
              t.offsetLeft - tabsContainer.offsetWidth / 2 + t.offsetWidth / 2;
            tabsContainer.scrollTo({
              left: targetScroll,
              behavior: "smooth",
            });
          }
        }),
    );

    // ==========================
    // 数据功能绑定
    // ==========================
    div.querySelector("#auto-scroll-btn").onclick = toggleAutoScroll;
    div.querySelector("#clean-data-btn").onclick = () => {
      if (confirm("确定清空已捕获的数据吗？")) {
        GLOBAL_DATA.clear();
        updateCountUI();
      }
    };
    div.querySelector("#export-btn").onclick = exportData;

    // 监听导出格式切换
    const fmtSel = div.querySelector("#export-format");
    const splitArea = div.querySelector("#split-setting-area");
    // 全类型支持分包，不再隐藏 splitArea
    // if (fmtSel && splitArea) { ... }

    // 资源下载绑定
    div.querySelector("#res-fetch-btn").onclick = handleFetchResources;
    div.querySelector("#res-select-all").onchange = (e) => {
      const checked = e.target.checked;
      div.querySelectorAll(".res-item").forEach((item) => {
        if (checked) item.classList.add("selected");
        else item.classList.remove("selected");
      });
    };
    div.querySelector("#res-quality-select").onchange = () => {
      if (CURRENT_NOTE_RAW) {
        renderResourceGrid(CURRENT_NOTE_RAW);
      }
    };
    div.querySelector("#res-download-btn").onclick = handleBatchDownload;

    // AI功能绑定
    // div.querySelector("#config-toggle").onclick = ... // Removed

    // 初始化并绑定屏蔽登录弹窗设置
    try {
      initShieldLoginFromSetting();
      bindShieldLoginSetting();
    } catch (e) {
      console.warn("[XHS助手] 初始化屏蔽登录设置失败", e);
    }

    // ============================
    // Config Manager Logic
    // ============================
    const DEFAULT_CONFIGS = [
      {
        id: "moonshot",
        name: "Moonshot (Kimi)",
        baseUrl: "https://api.moonshot.cn/v1/chat/completions",
        key: "",
        model: "moonshot-v1-8k",
        builtIn: true,
      },
      {
        id: "deepseek",
        name: "DeepSeek",
        baseUrl: "https://api.deepseek.com/chat/completions",
        key: "",
        model: "deepseek-chat",
        builtIn: true,
      },
      {
        id: "openai",
        name: "OpenAI (GPT)",
        baseUrl: "https://api.openai.com/v1/chat/completions",
        key: "",
        model: "gpt-3.5-turbo",
        builtIn: true,
      },
      {
        id: "custom",
        name: "自定义配置",
        baseUrl: "",
        key: "",
        model: "",
        builtIn: false,
      },
    ];

    let apiConfigs = [];
    try {
      apiConfigs = JSON.parse(GM_getValue("api_configs", "[]"));
      if (!Array.isArray(apiConfigs) || apiConfigs.length === 0)
        apiConfigs = JSON.parse(JSON.stringify(DEFAULT_CONFIGS));
    } catch (e) {
      apiConfigs = JSON.parse(JSON.stringify(DEFAULT_CONFIGS));
    }

    let currentConfigId = GM_getValue("current_api_config_id", "moonshot");

    function renderConfigSelect() {
      const sel = document.getElementById("api-config-select");
      sel.innerHTML = "";
      apiConfigs.forEach((c) => {
        const opt = document.createElement("option");
        opt.value = c.id;
        opt.innerText = c.name;
        sel.appendChild(opt);
      });

      // Ensure current ID exists
      if (!apiConfigs.some((c) => c.id === currentConfigId)) {
        currentConfigId = apiConfigs[0].id;
      }

      sel.value = currentConfigId;
      loadConfigToUI(currentConfigId);
    }

    function loadConfigToUI(id) {
      const config = apiConfigs.find((c) => c.id === id) || apiConfigs[0];
      currentConfigId = config.id;
      GM_setValue("current_api_config_id", currentConfigId);

      const baseUrlInput = document.getElementById("api-base-url");
      const keyInput = document.getElementById("api-key");
      const modelInput = document.getElementById("api-model");
      const delBtn = document.getElementById("api-config-del");
      const modelSelect = document.getElementById("api-model-select");

      baseUrlInput.value = config.baseUrl;
      keyInput.value = config.key;
      modelInput.value = config.model;

      // Hide model select on config switch
      modelSelect.style.display = "none";

      if (config.builtIn) {
        delBtn.style.display = "none";
      } else {
        delBtn.style.display = "block";
      }
    }

    function saveCurrentConfigFromUI() {
      const configIndex = apiConfigs.findIndex((c) => c.id === currentConfigId);
      if (configIndex !== -1) {
        apiConfigs[configIndex].baseUrl =
          document.getElementById("api-base-url").value;
        apiConfigs[configIndex].key = document.getElementById("api-key").value;
        apiConfigs[configIndex].model =
          document.getElementById("api-model").value;
        GM_setValue("api_configs", JSON.stringify(apiConfigs));

        // Sync legacy values for compatibility if needed elsewhere
        GM_setValue("api_base_url", apiConfigs[configIndex].baseUrl);
        GM_setValue("api_key", apiConfigs[configIndex].key);
        GM_setValue("api_model", apiConfigs[configIndex].model);
      }
    }

    // Bind Config Events
    document.getElementById("api-config-select").onchange = (e) =>
      loadConfigToUI(e.target.value);

    ["api-base-url", "api-key", "api-model"].forEach((id) => {
      const el = document.getElementById(id);
      el.onchange = saveCurrentConfigFromUI;
      el.oninput = saveCurrentConfigFromUI;
    });

    document.getElementById("api-config-add").onclick = () => {
      const name = prompt(
        "请输入新配置名称",
        "我的配置 " + (apiConfigs.length + 1),
      );
      if (name) {
        const newId = "custom_" + Date.now();
        apiConfigs.push({
          id: newId,
          name: name,
          baseUrl: "https://",
          key: "",
          model: "",
          builtIn: false,
        });
        GM_setValue("api_configs", JSON.stringify(apiConfigs));
        renderConfigSelect();
        // Select new
        document.getElementById("api-config-select").value = newId;
        loadConfigToUI(newId);
      }
    };

    document.getElementById("api-config-del").onclick = () => {
      if (confirm("确定删除当前配置吗？")) {
        apiConfigs = apiConfigs.filter((c) => c.id !== currentConfigId);
        if (apiConfigs.length === 0)
          apiConfigs = JSON.parse(JSON.stringify(DEFAULT_CONFIGS));
        GM_setValue("api_configs", JSON.stringify(apiConfigs));
        currentConfigId = apiConfigs[0].id;
        renderConfigSelect();
      }
    };

    document.getElementById("api-model-fetch-btn").onclick = async () => {
      const btn = document.getElementById("api-model-fetch-btn");
      const modelSel = document.getElementById("api-model-select");
      const originalText = btn.innerText;
      btn.innerText = "...";
      btn.disabled = true;

      const baseUrl = document.getElementById("api-base-url").value;
      const key = document.getElementById("api-key").value;

      if (!baseUrl) {
        alert("请先填写 Base URL");
        btn.innerText = originalText;
        btn.disabled = false;
        return;
      }
      if (!key) {
        alert("请先填写 API Key");
        btn.innerText = originalText;
        btn.disabled = false;
        return;
      }

      // Try to deduce models endpoint
      // Standard: .../v1/chat/completions -> .../v1/models
      let modelsUrl = baseUrl;
      if (modelsUrl.endsWith("/chat/completions")) {
        modelsUrl = modelsUrl.replace("/chat/completions", "/models");
      } else if (modelsUrl.endsWith("/")) {
        modelsUrl = modelsUrl + "models";
      } else {
        // If base url is just host, add /v1/models? Not sure, user usually puts chat/completions
        // Let's try replacing last segment if it is not models
        // Fallback: assume user put full chat url
        modelsUrl = modelsUrl.replace(/\/chat\/completions\/?$/, "/models");
      }

      console.log("[AI] Fetching models from:", modelsUrl);

      try {
        const res = await new Promise((resolve, reject) => {
          GM_xmlhttpRequest({
            method: "GET",
            url: modelsUrl,
            headers: {
              Authorization: `Bearer ${key}`,
              "Content-Type": "application/json",
            },
            onload: (r) => {
              if (r.status === 200) resolve(JSON.parse(r.responseText));
              else
                reject(
                  "Status " +
                    r.status +
                    "\n" +
                    r.responseText.substring(0, 100),
                );
            },
            onerror: (e) => reject("Network Error"),
          });
        });

        let models = [];
        if (res && Array.isArray(res.data)) {
          models = res.data.map((m) => m.id);
        } else if (Array.isArray(res)) {
          models = res.map((m) => m.id || m); // Some APIs return array directly
        }

        if (models.length > 0) {
          modelSel.innerHTML =
            '<option value="" disabled selected>请选择模型 (API获取成功)</option>';

          // 简单规则：包含特定关键词的模型被认为支持 AI 分析/分类 (即支持长文本或通用能力强)
          const capableKeywords = [
            "moonshot",
            "gpt-4",
            "claude-3",
            "deepseek",
            "128k",
            "32k",
            "200k",
            "pro",
          ];

          models.forEach((m) => {
            const opt = document.createElement("option");
            opt.value = m;
            let label = m;
            if (capableKeywords.some((k) => m.toLowerCase().includes(k))) {
              label += " (支持AI分析)";
            }
            opt.innerText = label;
            modelSel.appendChild(opt);
          });
          modelSel.style.display = "block";
          modelSel.onchange = () => {
            document.getElementById("api-model").value = modelSel.value;
            saveCurrentConfigFromUI();
          };
          // Auto expand?
          modelSel.click();
        } else {
          alert(
            "API请求成功，但未能解析出模型列表。请手动输入。\n" +
              JSON.stringify(res).slice(0, 100),
          );
        }
      } catch (e) {
        alert("获取模型失败: " + e + "\n尝试URL: " + modelsUrl);
      } finally {
        btn.innerText = originalText;
        btn.disabled = false;
      }
    };

    // Initialize
    renderConfigSelect();

    div.querySelector("#ai-gen-btn").onclick = handleAI;

    // 模版管理
    const managePanel = document.getElementById("template-manage-panel");
    document.getElementById("template-manage-toggle").onclick = () => {
      managePanel.style.display =
        managePanel.style.display === "none" ? "block" : "none";
    };
    document.getElementById("template-manage-select").onchange = () => {
      const t = templates.find(
        (x) => x.id === document.getElementById("template-manage-select").value,
      );
      if (t) fillTemplateForm(t);
    };
    document.getElementById("template-new-btn").onclick = () => {
      const empty = {
        id: "tmp_" + Date.now(),
        name: "新模版",
        desc1: "输入1",
        desc2: "输入2",
        placeholder1: "",
        placeholder2: "",
        system: "你是一个专业的创作者助手。",
        prompt: "请基于{{title}}和{{content}}生成创作内容。",
      };
      templates.push(empty);
      saveTemplates();
      refreshTemplates();
      refreshManageSelect(empty.id);
      fillTemplateForm(empty);
      toastManage("已创建空白模版，可编辑后保存");
    };
    document.getElementById("template-save-btn").onclick = () => {
      const id = document.getElementById("template-manage-select").value;
      const idx = templates.findIndex((x) => x.id === id);
      if (idx === -1) return alert("请选择要保存的模版");
      templates[idx] = collectTemplateForm(id);
      saveTemplates();
      refreshTemplates();
      refreshManageSelect(id);
      fillTemplateForm(templates[idx]);
      toastManage("已保存模版");
    };
    document.getElementById("template-save-new-btn").onclick = () => {
      const id = "tpl_" + Date.now();
      const t = collectTemplateForm(id);
      templates.push(t);
      saveTemplates();
      refreshTemplates();
      refreshManageSelect(id);
      fillTemplateForm(t);
      toastManage("已另存为新模版");
    };
    document.getElementById("template-delete-btn").onclick = () => {
      const id = document.getElementById("template-manage-select").value;
      if (!confirm("确认删除该模版吗？")) return;
      templates = templates.filter((x) => x.id !== id);
      if (!templates.length)
        templates = JSON.parse(JSON.stringify(DEFAULT_TEMPLATES));
      saveTemplates();
      refreshTemplates();
      const first = templates[0];
      refreshManageSelect(first.id);
      fillTemplateForm(first);
      toastManage("模版已删除");
    };
    document.getElementById("persona-gen-btn").onclick = handlePersonaGenerate;

    // AI分析功能绑定
    div.querySelector("#analysis-file-input").onchange = function () {
      const file = this.files[0];
      const label = div.querySelector("#analysis-file-label");
      const nameDisplay = div.querySelector("#analysis-file-name");
      if (file) {
        nameDisplay.textContent = "已选择: " + file.name;
        nameDisplay.style.display = "block";
        label.style.borderColor = "#000000";
        label.style.background = "#f6f6f3";
        label.textContent = "更换文件";
      }
    };

    div.querySelector("#analysis-summary-btn").onclick = handleAnalysisSummary;
    div.querySelector("#analysis-classify-btn").onclick =
      handleAnalysisClassify;
    div.querySelector("#analysis-config-btn").onclick = () => {
      // Switch to Settings tab
      const settingsTab = div.querySelector(
        '.ai-tab-item[data-tab="settings"]',
      );
      if (settingsTab) settingsTab.click();
    };

    // 营销工具绑定（营销面板当前已注释停用，元素不存在时跳过，避免异常中断后续初始化）
    const mktSearchBtn = div.querySelector("#mkt-search-btn");
    if (mktSearchBtn) {
      mktSearchBtn.onclick = () => {
        const kw = div.querySelector("#mkt-keyword").value.trim();
        if (kw) {
          const url = `https://www.xiaohongshu.com/search_result?keyword=${encodeURIComponent(kw)}&source=web_search_result_notes`;
          if (
            confirm(
              `即将跳转到搜索页：${kw}\n\n跳转后请点击“自动滚动”以加载更多笔记，然后再回到这里抓取评论。`,
            )
          ) {
            location.href = url;
          }
        }
      };
    }

    // 更新营销工具面板的数据计数
    // Add an observer or hook to update count when tab is switched
    tabs.forEach((t) => {
      if (t.dataset.tab === "marketing") {
        t.addEventListener("click", () => {
          const cnt = document.getElementById("mkt-note-count");
          if (cnt) cnt.innerText = GLOBAL_DATA.size;
        });
      }
    });

    const mktRunBtn = div.querySelector("#mkt-run-btn");
    if (mktRunBtn) mktRunBtn.onclick = handleBatchCommentCrawl;
    const mktExportBtn = div.querySelector("#mkt-export-btn");
    if (mktExportBtn) mktExportBtn.onclick = exportCommentData;

    // 绑定导出字段点击事件
    const exportContainer = div.querySelector("#export-field-container");
    if (exportContainer) {
      exportContainer.addEventListener("click", (e) => {
        const item = e.target.closest(".export-field-item");
        if (item) {
          item.classList.toggle("selected");
        }
      });
    }

    refreshManageSelect();
    if (templates[0]) fillTemplateForm(templates[0]);

    refreshTemplates();

    // 初始化时，如果页面已经有数据（SSR），尝试简单抓取一下当前DOM补充（作为兜底）
    setTimeout(scanInitialDOM, 2000);
  }

  // ==========================================
  // 5. 自动滚动与数据逻辑
  // ==========================================

  function updateCountUI() {
    const el = document.getElementById("page-obj-count");
    if (el) el.innerText = GLOBAL_DATA.size;
  }

  // 兜底策略：扫描当前DOM (针对页面刚打开时已经存在的数据)
  function scanInitialDOM() {
    const cards = document.querySelectorAll("section.note-item, .feed-card");
    let count = 0;
    cards.forEach((card) => {
      // 尝试获取ID和链接
      // fix: 优先找 a.cover，因为它通常包含带 xsec_token 的完整链接，且覆盖在卡片上
      // 其次才是普通的 explore 链接
      let linkEl =
        card.querySelector("a.cover") ||
        card.querySelector('a[href*="/explore/"]') ||
        card.querySelector('a[href*="/user/profile/"]');

      let href = "";
      if (linkEl) href = linkEl.href; // .href 获取的是绝对路径

      if (href) {
        // 提取 ID
        // 匹配逻辑：获取 URL path 的最后一段作为 ID
        // 例如: /explore/66... 或 /user/profile/xxx/66...
        let id = "";
        try {
          const urlObj = new URL(href);
          const pathParts = urlObj.pathname.split("/").filter((p) => p);
          // 假设最后一部分是ID (通常是24位ObjectId)
          const lastPart = pathParts[pathParts.length - 1];
          // 简单的校验：ID通常由字母数字组成，长度24位左右
          if (lastPart && /^[a-fA-F0-9]{24}$/.test(lastPart)) {
            id = lastPart;
          } else if (href.includes("/explore/")) {
            // 兼容旧的 explore 提取方式
            const m = href.match(/\/explore\/(\w+)/);
            if (m) id = m[1];
          }
        } catch (e) {}

        if (id && !GLOBAL_DATA.has(id)) {
          const title =
            (
              card.querySelector(".title span") ||
              card.querySelector(".title") ||
              {}
            ).innerText || "未获取";
          const author =
            (card.querySelector(".author") || {}).innerText || "未获取";

          // 尝试获取封面
          let coverUrl = "";
          const coverDiv = card.querySelector(".cover");
          if (coverDiv) {
            const style = coverDiv.getAttribute("style");
            const bgMatch = style && style.match(/url\("?(.+?)"?\)/);
            if (bgMatch) coverUrl = bgMatch[1];
          }
          if (!coverUrl) {
            const img = card.querySelector("img");
            if (img) coverUrl = img.src;
          }

          GLOBAL_DATA.set(id, {
            笔记ID: id,
            标题: title,
            链接: href, // 使用从 DOM 获取的完整 href（包含 token）
            作者: author,
            点赞数: (card.querySelector(".count") || {}).innerText || "0",
            封面图: coverUrl,
            类型: "dom_scan",
          });
          count++;
        }
      }
    });
    if (count > 0) updateCountUI();
    console.log(`[XHS助手] 初始DOM扫描发现 ${count} 条数据`);
  }

  // 自动滚动逻辑
  let scrollInterval;
  function toggleAutoScroll() {
    const btn = document.getElementById("auto-scroll-btn");
    const status = document.getElementById("scroll-status");

    if (isAutoScrolling) {
      // 停止
      isAutoScrolling = false;
      clearInterval(scrollInterval);
      btn.innerText = "自动滚动加载全部";
      btn.classList.remove("scrolling");
      status.innerText = "已停止";
    } else {
      // 开始
      isAutoScrolling = true;
      btn.innerText = "停止滚动 (抓取中...)";
      btn.classList.add("scrolling");
      status.innerText = "滚动中...";

      let lastHeight = 0;
      let sameHeightCount = 0;

      scrollInterval = setInterval(() => {
        window.scrollTo(0, document.body.scrollHeight);

        const currentHeight = document.body.scrollHeight;
        if (currentHeight === lastHeight) {
          sameHeightCount++;
          if (sameHeightCount > 10) {
            // 连续10次高度不变（约10-15秒），认为到底了
            toggleAutoScroll(); // 自动停止
            alert(
              `滚动结束！共捕获 ${GLOBAL_DATA.size} 条数据。\n请点击导出。`,
            );
          }
        } else {
          sameHeightCount = 0;
          lastHeight = currentHeight;
        }
      }, 1200); // 间隔1.2秒滚动一次，给接口加载留时间
    }
  }

  // ==========================================
  // 新增核心功能：ExcelJS 导出与图片嵌入
  // ==========================================
  function fetchImageBuffer(url) {
    return new Promise((resolve) => {
      if (!url || !url.startsWith("http")) {
        resolve(null);
        return;
      }
      GM_xmlhttpRequest({
        method: "GET",
        url: url,
        responseType: "arraybuffer",
        timeout: 15000,
        onload: (res) => {
          if (res.status === 200) {
            resolve(res.response);
          } else {
            resolve(null);
          }
        },
        onerror: () => resolve(null),
        ontimeout: () => resolve(null),
      });
    });
  }

  async function exportAdvanced(dataList, selectedCols, splitSize = 200) {
    const statusDiv = document.createElement("div");
    statusDiv.id = "xhs-export-status";
    statusDiv.style.cssText =
      "position:fixed;top:20px;left:50%;transform:translateX(-50%);background:#ffffff;color:#000000;padding:14px 20px;border-radius:16px;border:1px solid #dadad3;z-index:9999999;font-size:14px;line-height:1.5;box-shadow:0 16px 32px rgba(0,0,0,0.16);max-width:320px;";
    document.body.appendChild(statusDiv);

    const updateStatus = (msg) => {
      if (statusDiv) statusDiv.innerHTML = msg;
    };
    const closeStatus = () => {
      setTimeout(() => {
        if (statusDiv) statusDiv.remove();
      }, 3000);
    };

    try {
      updateStatus("正在初始化数据...");

      // 1. 数据分包
      const chunks = [];
      for (let i = 0; i < dataList.length; i += splitSize) {
        chunks.push(dataList.slice(i, i + splitSize));
      }

      // 并发控制辅助函数
      const processImagesInBatch = async (tasks, limit = 5) => {
        const results = [];
        const executing = [];
        for (const task of tasks) {
          const p = Promise.resolve().then(() => task());
          results.push(p);
          if (limit <= tasks.length) {
            const e = p.then(() => executing.splice(executing.indexOf(e), 1));
            executing.push(e);
            if (executing.length >= limit) {
              await Promise.race(executing);
            }
          }
        }
        return Promise.all(results);
      };

      for (let i = 0; i < chunks.length; i++) {
        const chunk = chunks[i];
        const wb = new ExcelJS.Workbook();
        const sheet = wb.addWorksheet("小红书数据");

        // 设置列
        sheet.columns = selectedCols.map((key) => {
          let w = 25;
          if (key === "封面图") w = 35; // 适配 250px 宽度
          if (key === "链接" || key === "标题") w = 40;
          if (key === "内容" || key === "描述") w = 60;
          if (key === "笔记ID") w = 30;
          return { header: key, key: key, width: w };
        });

        // 收集图片任务
        const imageTasks = [];

        // 填充数据文本
        chunk.forEach((rowData) => {
          const row = sheet.addRow(rowData);
          const rowNum = row.number;

          // 如果有封面图列，准备下载
          if (selectedCols.includes("封面图")) {
            const url = rowData["封面图"];
            if (url) {
              row.height = 250; // 适配 334px 高度 (approx 250 points)
              // 记录任务
              imageTasks.push(async () => {
                const buffer = await fetchImageBuffer(url);
                // 获取图片尺寸用于按比例缩放
                let w = 0,
                  h = 0;
                if (buffer) {
                  try {
                    const blob = new Blob([buffer]);
                    const u = URL.createObjectURL(blob);
                    const img = new Image();
                    await new Promise((resolve) => {
                      img.onload = () => {
                        w = img.naturalWidth;
                        h = img.naturalHeight;
                        resolve();
                      };
                      img.onerror = resolve;
                      img.src = u;
                    });
                    URL.revokeObjectURL(u);
                  } catch (e) {}
                }
                return {
                  buffer,
                  rowNum,
                  colIndex: selectedCols.indexOf("封面图"),
                  w,
                  h,
                };
              });
            } else {
              row.height = 20;
            }
          } else {
            row.height = 20;
          }

          // 处理链接样式
          if (selectedCols.includes("链接")) {
            const idx = selectedCols.indexOf("链接");
            const cell = row.getCell(idx + 1);
            if (cell.value && String(cell.value).startsWith("http")) {
              cell.value = { text: "点击跳转", hyperlink: rowData["链接"] };
              cell.font = { color: { argb: "FF0000FF" }, underline: true };
            }
          }
        });

        // 执行图片下载任务
        if (imageTasks.length > 0) {
          // 分批下载提示
          const totalTasks = imageTasks.length;
          updateStatus(
            `处理分包 ${i + 1}/${chunks.length}<br>正在下载并嵌入 ${totalTasks} 张图片 (并发5)...<br><span style='font-size:12px;opacity:0.8'>请勿关闭页面</span>`,
          );

          const results = await processImagesInBatch(imageTasks, 5);

          // 嵌入图片
          results.forEach((res) => {
            if (res && res.buffer) {
              const imgId = wb.addImage({
                buffer: res.buffer,
                extension: "png",
              });

              // 按比例缩放 (适配 334px 高度)
              let finalW = 250;
              let finalH = 334;
              if (res.w && res.h) {
                finalW = finalH * (res.w / res.h);
              }

              sheet.addImage(imgId, {
                tl: { col: res.colIndex, row: res.rowNum - 1 }, // row is 0-based for image pos
                ext: { width: finalW, height: finalH },
                editAs: "oneCell",
              });
              // 清空该单元格文本
              sheet.getRow(res.rowNum).getCell(res.colIndex + 1).value = "";
            }
          });
        }

        const buffer = await wb.xlsx.writeBuffer();
        const fileName =
          chunks.length > 1
            ? `xhs_data_part_${i + 1}.xlsx`
            : `xhs_data_full.xlsx`;

        updateStatus(`导出分包 ${i + 1}/${chunks.length}，正在下载...`);
        const blob = new Blob([buffer], {
          type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        });
        const url = URL.createObjectURL(blob);
        const a = document.createElement("a");
        a.href = url;
        a.download = fileName;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);

        setTimeout(() => URL.revokeObjectURL(url), 10000);

        // 如果是多包下载，稍作停顿，避免浏览器拦截或卡顿
        if (chunks.length > 1) {
          await new Promise((r) => setTimeout(r, 1000));
        }
      }

      closeStatus();
    } catch (e) {
      console.error(e);
      updateStatus("导出出错: " + e.message);
      setTimeout(() => closeStatus(), 5000);
    }
  }

  async function exportList(
    dataList,
    format,
    baseName,
    selectedCols,
    splitSize = 999999,
  ) {
    if (!dataList || dataList.length === 0) return;

    // 默认全选
    let headers = Object.keys(dataList[0]);
    if (selectedCols && selectedCols.length > 0) {
      headers = selectedCols;
    }

    // 分包处理
    const chunks = [];
    for (let i = 0; i < dataList.length; i += splitSize) {
      chunks.push(dataList.slice(i, i + splitSize));
    }

    // 简易提示
    const statusDiv = document.createElement("div");
    if (chunks.length > 1) {
      statusDiv.style.cssText =
        "position:fixed;top:20px;left:50%;transform:translateX(-50%);background:#ffffff;color:#000000;padding:14px 20px;border-radius:16px;border:1px solid #dadad3;z-index:999999;font-size:14px;line-height:1.5;box-shadow:0 16px 32px rgba(0,0,0,0.16);max-width:320px;";
      document.body.appendChild(statusDiv);
    }

    for (let i = 0; i < chunks.length; i++) {
      const chunk = chunks[i];
      const currentBaseName =
        chunks.length > 1 ? `${baseName}_part_${i + 1}` : baseName;

      if (chunks.length > 1) {
        statusDiv.innerHTML = `正在导出 ${format.toUpperCase()} 分包 ${i + 1}/${chunks.length}...`;
      }

      let content = "";
      let type = "";
      let ext = "";

      if (format === "json") {
        const filteredList = chunk.map((row) => {
          const newRow = {};
          headers.forEach((h) => (newRow[h] = row[h]));
          return newRow;
        });
        content = JSON.stringify(filteredList, null, 2);
        ext = "json";
        type = "application/json";
      } else if (format === "md") {
        content =
          "# 小红书采集数据导出\n\n> 导出时间: " +
          new Date().toLocaleString() +
          "\n\n---\n\n";
        chunk.forEach((row, index) => {
          const rowIdx = i * splitSize + index + 1;
          content += `### ${rowIdx}. ${row["标题"] || "无标题"}\n\n`;
          content += `**作者**: ${row["作者"] || "未知"}  |  **点赞**: ${row["点赞数"] || 0}  |  [原文链接](${row["链接"]})\n\n`;
          if (row["封面图"]) {
            content += `![封面图](${row["封面图"]})\n\n`;
          }
          if (row["内容"] || row["描述"]) {
            content += `**内容描述**:\n\n${(row["内容"] || row["描述"]).replace(/\n/g, "\n\n")}\n\n`;
          }
          content += `---\n\n`;
        });
        ext = "md";
        type = "text/markdown;charset=utf-8";
      } else if (format === "html") {
        // HTML: 尝试下载图片并转 Base64 用于本地保存
        if (chunks.length <= 1) {
          // 如果仅一包但也需要处理图片，必须显示提示 (因为下载图片较慢)
          statusDiv.style.cssText =
            "position:fixed;top:20px;left:50%;transform:translateX(-50%);background:#ffffff;color:#000000;padding:14px 20px;border-radius:16px;border:1px solid #dadad3;z-index:999999;font-size:14px;line-height:1.5;box-shadow:0 16px 32px rgba(0,0,0,0.16);max-width:320px;";
          if (!statusDiv.parentNode) document.body.appendChild(statusDiv);
        }
        statusDiv.innerHTML = `正在导出 HTML (Part ${i + 1}/${chunks.length})<br>正在下载并将图片转为Base64以实现本地保存...`;

        // 并发控制函数
        const processImagesInBatch = async (tasks, limit = 5) => {
          const results = [];
          const executing = [];
          for (const task of tasks) {
            const p = Promise.resolve().then(() => task());
            results.push(p);
            if (limit <= tasks.length) {
              const e = p.then(() => executing.splice(executing.indexOf(e), 1));
              executing.push(e);
              if (executing.length >= limit) {
                await Promise.race(executing);
              }
            }
          }
          return Promise.all(results);
        };

        const imageTasks = chunk.map((row, idx) => async () => {
          const url = row["封面图"];
          if (!url) return { index: idx, base64: "" };
          try {
            const buffer = await fetchImageBuffer(url);
            if (buffer) {
              const blob = new Blob([buffer]);
              return new Promise((resolve) => {
                const reader = new FileReader();
                reader.onloadend = () =>
                  resolve({ index: idx, base64: reader.result });
                reader.readAsDataURL(blob);
              });
            }
          } catch (e) {
            console.warn("Image fetch failed", e);
          }
          return { index: idx, base64: url }; // 失败则使用原链接
        });

        // 执行图片下载任务
        const imageResults = await processImagesInBatch(imageTasks, 5);
        const base64Map = {};
        imageResults.forEach((r) => {
          if (r) base64Map[r.index] = r.base64;
        });

        let cardsHtml = "";
        chunk.forEach((row, idx) => {
          const cover = base64Map[idx] || row["封面图"] || "";
          const title = row["标题"] || "无标题";
          const author = row["作者"] || "未知";
          const likes = row["点赞数"] || "0";
          const link = row["链接"] || "#";
          const desc = row["内容"] || row["描述"] || "";

          cardsHtml += `
               <div class="note-card">
                  <a href="${link}" target="_blank" class="card-link">
                      <div class="cover-wrapper">
                          <img src="${cover}" class="cover-img" loading="lazy" alt="${title}">
                      </div>
                      <div class="card-body">
                          <div class="title">${title}</div>
                          <div class="author-info">
                              <div class="author">
                                 <span>${author}</span>
                              </div>
                              <div class="likes">
                                 <span>${likes}</span>
                              </div>
                          </div>
                           <div class="desc">${desc}</div>
                      </div>
                  </a>
               </div>`;
        });

        content = `
            <!DOCTYPE html>
            <html lang="zh-CN">
            <head>
              <meta charset="utf-8">
              <meta name="viewport" content="width=device-width, initial-scale=1.0">
              <title>小红书数据导出 - Part ${i + 1}</title>
              <style>
                  body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif; background: #f4f6f8; padding: 20px; margin: 0; }
                  .header { text-align: center; margin-bottom: 30px; color: #000000; }
                  .container { max-width: 1400px; margin: 0 auto; display: grid; grid-template-columns: repeat(auto-fill, minmax(280px, 1fr)); gap: 20px; }
                  .note-card { background: white; border-radius: 12px; overflow: hidden; box-shadow: 0 2px 10px rgba(0,0,0,0.05); transition: transform 0.2s, box-shadow 0.2s; position: relative; }
                  .note-card:hover { transform: translateY(-5px); box-shadow: 0 8px 20px rgba(0,0,0,0.12); }
                  .card-link { text-decoration: none; color: inherit; display: block; height: 100%; }
                  .cover-wrapper { position: relative; padding-top: 133.33%; background: #f0f0f0; }
                  .cover-img { position: absolute; top: 0; left: 0; width: 100%; height: 100%; object-fit: cover; }
                  .card-body { padding: 12px; display: flex; flex-direction: column; justify-content: space-between; }
                  .title { font-weight: 600; font-size: 15px; line-height: 1.4; color: #000000; margin-bottom: 8px; display: -webkit-box; -webkit-line-clamp: 2; -webkit-box-orient: vertical; overflow: hidden; text-overflow: ellipsis; min-height: 42px; }
                  .author-info { display: flex; align-items: center; justify-content: space-between; font-size: 12px; color: #91918c; margin-bottom: 8px; }
                  .desc { font-size: 13px; color: #62625b; line-height: 1.5; display: -webkit-box; -webkit-line-clamp: 3; -webkit-box-orient: vertical; overflow: hidden; text-overflow: ellipsis; margin-top: 5px; }
                  @media (max-width: 600px) {
                      .container { grid-template-columns: repeat(2, 1fr); gap: 10px; }
                      body { padding: 10px; }
                  }
              </style>
            </head>
            <body>
                <div class="header">
                    <h2>小红书采集数据 (Part ${i + 1}/${chunks.length}, 共${chunk.length}条)</h2>
                    <p style="font-size:12px;color:#62625b;">导出时间: ${new Date().toLocaleString()}</p>
                </div>
                <div class="container">
                    ${cardsHtml}
                </div>
            </body>
            </html>`;
        ext = "html";
        type = "text/html;charset=utf-8";
      } else if (format === "xls") {
        // Excel (HTML Table)
        let html = `
            <html xmlns:o="urn:schemas-microsoft-com:office:office" 
                  xmlns:x="urn:schemas-microsoft-com:office:excel" 
                  xmlns="http://www.w3.org/TR/REC-html40">
            <head>
               <meta charset="utf-8">
               <style>
                 td { vertical-align: middle; text-align: center; font-size: 11pt; }
                 .text { mso-number-format:"\@"; }
               </style>
            </head>
            <body>
            <table border="1" style="border-collapse: collapse; width: 100%;">`;
        html += "<thead><tr style='background-color:#f2f2f2; height:40px;'>";
        headers.forEach(
          (h) =>
            (html += `<th style="padding:10px; border:1px solid #dadad3;">${h}</th>`),
        );
        html += "</tr></thead><tbody>";

        chunk.forEach((row) => {
          html += "<tr style='height:110px;'>";
          headers.forEach((h) => {
            const val = row[h] || "";
            if (h === "封面图" && val) {
              html += `<td style="width:120px; text-align:center;"><img src="${val}" width="100" height="100" /></td>`;
            } else if (h === "链接" && val) {
              html += `<td><a href="${val}" target="_blank">点击跳转</a></td>`;
            } else {
              html += `<td class="text" style="max-width:300px; overflow:hidden;">${val}</td>`;
            }
          });
          html += "</tr>";
        });
        html += "</tbody></table></body></html>";

        content = html;
        ext = "xls";
        type = "application/vnd.ms-excel";
      } else {
        // CSV
        const csvBody = chunk
          .map((row) =>
            headers
              .map((h) => {
                let v = row[h] || "";
                v = String(v).replace(/"/g, '""');
                return `"${v}"`;
              })
              .join(","),
          )
          .join("\n");
        content = "\ufeff" + headers.join(",") + "\n" + csvBody;
        ext = "csv";
        type = "text/csv;charset=utf-8";
      }

      download(content, `${currentBaseName}.${ext}`, type);

      // 延时防拦截
      if (chunks.length > 1) {
        await new Promise((r) => setTimeout(r, 1000));
      }
    }

    if (chunks.length > 1) {
      statusDiv.innerHTML = "所有分包导出完成";
      setTimeout(() => statusDiv.remove(), 2000);
    }
  }

  function exportData() {
    if (GLOBAL_DATA.size === 0) {
      alert("数据为空！请先点击「自动滚动」或手动浏览页面。");
      return;
    }
    const format = document.getElementById("export-format").value;

    // 获取勾选的列
    let selectedCols = [];
    const exportContainer = document.getElementById("export-field-container");

    if (exportContainer) {
      // 新版：表格选择
      const selectedItems = exportContainer.querySelectorAll(
        ".export-field-item.selected",
      );
      selectedCols = Array.from(selectedItems).map((el) =>
        el.getAttribute("data-value"),
      );
    } else {
      // ... fallback ...
      const selectEl = document.getElementById("export-col-select");
      if (selectEl) {
        selectedCols = Array.from(selectEl.selectedOptions).map(
          (opt) => opt.value,
        );
      } else {
        const checks = document.querySelectorAll(".export-col-check");
        if (checks.length > 0) {
          selectedCols = Array.from(checks)
            .filter((c) => c.checked)
            .map((c) => c.value);
        }
      }
    }

    // 如果用户什么都没选，提醒一下（或者默认全选）
    if (selectedCols.length === 0) {
      alert("请至少选择一个导出字段！");
      return;
    }

    const dataList = Array.from(GLOBAL_DATA.values());

    // 统一获取分包设置
    const enableSplit = document.getElementById("enable-split").checked;
    const splitSizeInput =
      parseInt(document.getElementById("split-size").value) || 200;
    const splitSize = enableSplit ? splitSizeInput : 999999;

    if (format === "xlsx_embed") {
      exportAdvanced(dataList, selectedCols, splitSize);
    } else {
      exportList(dataList, format, "xhs_data_full", selectedCols, splitSize);
    }
  }

  function download(content, name, type) {
    const blob = new Blob([content], { type: type });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = name;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
  }

  // ==========================================
  // 6. 辅助功能 (拖拽等)
  // ==========================================
  // 面板展开/收起切换
  // state: "expand" | "minimize" | 不传则自动取反
  function togglePanel(div, state) {
    const isMin = div.classList.contains("minimized");
    const toMinimize =
      state === "minimize" ? true : state === "expand" ? false : !isMin;

    if (toMinimize) {
      div.classList.add("minimized");
      // 恢复图标上次拖拽保存的位置
      const pos = GM_getValue("pos");
      if (pos) {
        div.style.left = pos.l;
        div.style.top = pos.t;
        div.style.right = "auto";
      } else {
        div.style.left = "";
        div.style.top = "";
        div.style.right = "";
      }
    } else {
      // 展开：清除内联定位，回到 CSS 控制的右侧固定停靠
      // (position:fixed; top:0; right:0; height:100vh)，面板不可移动
      div.classList.remove("minimized");
      div.style.left = "";
      div.style.top = "";
      div.style.right = "";
    }
  }

  function bindDrag(div) {
    // 仅"最小化(图标)"状态可拖拽；展开状态面板固定停靠右侧，禁止拖拽移动
    let isDragging = false;
    let hasMoved = false;
    let startX, startY, initL, initT;

    const isMinimized = () => div.classList.contains("minimized");

    div.addEventListener("mousedown", (e) => {
      if (!isMinimized()) return; // 展开状态：面板固定，不响应拖拽
      isDragging = true;
      hasMoved = false;
      startX = e.clientX;
      startY = e.clientY;
      const rect = div.getBoundingClientRect();
      initL = rect.left;
      initT = rect.top;

      // 优化：拖拽开始时禁用过渡动画，防止延迟感
      div.style.transition = "none";
      // 防止由拖拽引起的文本选中
      document.body.style.userSelect = "none";
      e.preventDefault();
    });

    document.addEventListener("mousemove", (e) => {
      if (!isDragging) return;
      const dx = e.clientX - startX;
      const dy = e.clientY - startY;
      // 位移小于阈值视为点击，不进入拖拽
      if (!hasMoved && Math.abs(dx) + Math.abs(dy) < 5) return;
      hasMoved = true;

      let newLeft = initL + dx;
      let newTop = initT + dy;
      // 边界保护：确保图标不会拖出屏幕
      const maxLeft = window.innerWidth - 36;
      const maxTop = window.innerHeight - 36;
      newLeft = Math.max(0, Math.min(newLeft, maxLeft));
      newTop = Math.max(0, Math.min(newTop, maxTop));

      div.style.left = newLeft + "px";
      div.style.top = newTop + "px";
      div.style.right = "auto";
    });

    document.addEventListener("mouseup", () => {
      if (!isDragging) return;
      isDragging = false;
      // 拖拽结束，恢复过渡动画和选择
      div.style.transition = "";
      document.body.style.userSelect = "";

      if (hasMoved) {
        // 记录图标位置
        GM_setValue("pos", { l: div.style.left, t: div.style.top });
      } else {
        // 未移动 视为点击图标 在浏览器右侧固定位置展开面板
        togglePanel(div, "expand");
      }
    });

    // 恢复上次保存的图标位置（默认最小化时生效）
    if (isMinimized()) {
      const pos = GM_getValue("pos");
      if (pos) {
        const savedLeft = Math.max(
          0,
          Math.min(parseInt(pos.l) || 0, window.innerWidth - 100),
        );
        const savedTop = Math.max(
          0,
          Math.min(parseInt(pos.t) || 0, window.innerHeight - 100),
        );
        div.style.left = savedLeft + "px";
        div.style.top = savedTop + "px";
        div.style.right = "auto";
      }
    }
  }

  // ==========================================
  // 7. AI 逻辑 (保持原样)
  // ==========================================
  function refreshTemplates() {
    const sel = document.getElementById("template-select");
    sel.innerHTML = "";
    templates.forEach((t) => {
      const opt = document.createElement("option");
      opt.value = t.id;
      opt.innerText = t.name;
      sel.appendChild(opt);
    });
    sel.onchange = () => {
      const t = templates.find((x) => x.id === sel.value);
      document.getElementById("input-1").placeholder =
        t.placeholder1 || t.desc1;
      document.getElementById("input-2").placeholder =
        t.placeholder2 || t.desc2;
    };
    sel.onchange();
  }

  function refreshManageSelect(selectId) {
    const sel = document.getElementById("template-manage-select");
    if (!sel) return;
    sel.innerHTML = "";
    templates.forEach((t) => {
      const opt = document.createElement("option");
      opt.value = t.id;
      opt.innerText = t.name;
      sel.appendChild(opt);
    });
    if (selectId && templates.some((t) => t.id === selectId))
      sel.value = selectId;
    else if (templates[0]) sel.value = templates[0].id;
  }

  function fillTemplateForm(t) {
    document.getElementById("template-manage-select").value = t.id;
    document.getElementById("tpl-name").value = t.name || "";
    document.getElementById("tpl-desc1").value = t.desc1 || "";
    document.getElementById("tpl-desc2").value = t.desc2 || "";
    document.getElementById("tpl-system").value = t.system || "";
    document.getElementById("tpl-prompt").value = t.prompt || "";
    document.getElementById("tpl-persona").value = t.persona || "";
  }

  function collectTemplateForm(id) {
    return {
      id,
      name: document.getElementById("tpl-name").value || "未命名模版",
      desc1: document.getElementById("tpl-desc1").value || "输入1",
      desc2: document.getElementById("tpl-desc2").value || "输入2",
      placeholder1: document.getElementById("tpl-desc1").value || "",
      placeholder2: document.getElementById("tpl-desc2").value || "",
      system: document.getElementById("tpl-system").value || "",
      prompt:
        document.getElementById("tpl-prompt").value ||
        "请基于{{title}}和{{content}}生成创作内容。",
    };
  }

  function toastManage(msg) {
    const el = document.getElementById("template-manage-status");
    if (el) {
      el.innerText = msg;
      setTimeout(() => {
        if (el.innerText === msg) el.innerText = "";
      }, 2000);
    }
  }

  async function handlePersonaGenerate() {
    const persona = document.getElementById("tpl-persona").value.trim();
    const status = document.getElementById("template-manage-status");
    if (!persona) return alert("请先填写模版人设/要求");

    status.innerText = "AI 生成人设中...";
    try {
      const res = await new Promise((resolve, reject) => {
        GM_xmlhttpRequest({
          method: "POST",
          url: document.getElementById("api-base-url").value,
          headers: {
            "Content-Type": "application/json",
            Authorization: `Bearer ${document.getElementById("api-key").value}`,
          },
          data: JSON.stringify({
            model: document.getElementById("api-model").value,
            messages: [
              {
                role: "system",
                content:
                  '你是提示词工程师，请根据用户的人设描述，返回JSON {"system":"...","prompt":"..."}，prompt内使用 {{title}} 和 {{content}} 作为占位符。',
              },
              {
                role: "user",
                content: persona,
              },
            ],
          }),
          onload: (r) => resolve(JSON.parse(r.responseText)),
          onerror: reject,
        });
      });

      const content = res.choices?.[0]?.message?.content || "";
      let json = {};
      try {
        json = JSON.parse(content.replace(/```json|```/g, "").trim());
      } catch (e) {
        json = {};
      }

      if (json.system)
        document.getElementById("tpl-system").value = json.system;
      if (json.prompt)
        document.getElementById("tpl-prompt").value = json.prompt;
      toastManage("AI 已生成模版提示");
    } catch (e) {
      console.error(e);
      status.innerText = "AI 生成失败";
      alert("AI生成失败: " + e);
    }
  }

  async function handleAI() {
    const btn = document.getElementById("ai-gen-btn");
    const status = document.getElementById("ai-status");
    const tId = document.getElementById("template-select").value;
    const v1 = document.getElementById("input-1").value;
    const v2 = document.getElementById("input-2").value;

    if (!v1 && !v2) return alert("请输入内容");

    btn.disabled = true;
    btn.innerText = "生成中...";
    try {
      const t = templates.find((x) => x.id === tId);
      const prompt = t.prompt
        .replace("{{title}}", v1)
        .replace("{{content}}", v2);

      const res = await new Promise((resolve, reject) => {
        GM_xmlhttpRequest({
          method: "POST",
          url: document.getElementById("api-base-url").value,
          headers: {
            "Content-Type": "application/json",
            Authorization: `Bearer ${document.getElementById("api-key").value}`,
          },
          data: JSON.stringify({
            model: document.getElementById("api-model").value,
            messages: [
              { role: "system", content: t.system },
              { role: "user", content: prompt },
            ],
          }),
          onload: (r) => resolve(JSON.parse(r.responseText)),
          onerror: reject,
        });
      });

      const content = res.choices[0].message.content;
      let json = {};
      try {
        json = JSON.parse(content.replace(/```json|```/g, "").trim());
      } catch (e) {
        json = { title: "Error", content: content };
      }

      const titleInput = document.querySelector('input[placeholder*="标题"]');
      if (titleInput) {
        titleInput.value = json.title;
        titleInput.dispatchEvent(new Event("input", { bubbles: true }));
      }

      const editor = document.getElementById("post-textarea");
      if (editor) {
        editor.innerText = json.content;
        editor.dispatchEvent(new Event("input", { bubbles: true }));
      } else {
        navigator.clipboard.writeText(json.content);
        alert("已复制到剪贴板");
      }

      status.innerText = "完成";
    } catch (e) {
      console.error(e);
      status.innerText = "失败";
      alert("API错误: " + e);
    } finally {
      btn.disabled = false;
      btn.innerText = "生成文案";
    }
  }

  // ==========================================
  // 8. AI 分析与分类逻辑
  // ==========================================

  function readFileData(file) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = (e) => resolve(e.target.result);
      reader.onerror = reject;
      reader.readAsText(file);
    });
  }

  function parseUploadedData(text, type) {
    if (type.includes("json")) {
      return JSON.parse(text);
    } else {
      // Simple CSV Parser
      const lines = text.split("\n").filter((l) => l.trim());
      if (lines.length < 2) return [];
      const headers = lines[0].split(","); // 简单分割，未处理复杂CSV
      // 这里的CSV解析较为简陋，建议使用JSON上传
      const result = [];
      for (let i = 1; i < lines.length; i++) {
        const row = lines[i].split(","); // 同样简单分割
        if (row.length === headers.length) {
          let obj = {};
          headers.forEach((h, idx) => {
            obj[h.replace(/"/g, "").trim()] = row[idx]
              ? row[idx].replace(/"/g, "").trim()
              : "";
          });
          result.push(obj);
        }
      }
      return result;
    }
  }

  async function callAI(messages) {
    return new Promise((resolve, reject) => {
      GM_xmlhttpRequest({
        method: "POST",
        url: document.getElementById("api-base-url").value,
        headers: {
          "Content-Type": "application/json",
          Authorization: `Bearer ${document.getElementById("api-key").value}`,
        },
        data: JSON.stringify({
          model: document.getElementById("api-model").value,
          messages: messages,
        }),
        onload: (r) => {
          try {
            resolve(JSON.parse(r.responseText));
          } catch (e) {
            reject(e);
          }
        },
        onerror: reject,
      });
    });
  }

  async function handleAnalysisSummary() {
    const fileInput = document.getElementById("analysis-file-input");
    const resultDiv = document.getElementById("analysis-result");
    const btn = document.getElementById("analysis-summary-btn");

    if (!fileInput.files.length) return alert("请先上传文件");
    const file = fileInput.files[0];

    btn.disabled = true;
    btn.innerText = "分析中...";
    resultDiv.style.display = "block";
    resultDiv.innerText = "正在读取文件并进行 AI 分析，请稍候...";

    try {
      const text = await readFileData(file);
      let data = parseUploadedData(text, file.name);

      if (!data || data.length === 0) throw new Error("解析数据失败或数据为空");

      // 截取前 50 条数据，防止 Token 溢出
      const sample = data.slice(0, 50).map((item) => ({
        标题: item.标题 || item.display_title,
        点赞: item.点赞数 || item.likes,
        作者: item.作者 || item.user?.nickname,
      }));

      const prompt = `请分析以下小红书笔记数据（仅展示前${sample.length}条）：\n${JSON.stringify(
        sample,
      )}\n\n请给出：\n1. 热门话题/关键词总结\n2. 高赞笔记的共同特点\n3. 内容创作建议`;

      const res = await callAI([
        {
          role: "system",
          content: "你是一个资深的小红书数据分析师，擅长洞察趋势。",
        },
        { role: "user", content: prompt },
      ]);

      const content = res.choices?.[0]?.message?.content || "AI 未返回内容";
      resultDiv.innerText = content;
    } catch (e) {
      console.error(e);
      resultDiv.innerText = "分析失败: " + e.message;
    } finally {
      btn.disabled = false;
      btn.innerText = "🧠 生成总结分析";
    }
  }

  async function handleAnalysisClassify() {
    const fileInput = document.getElementById("analysis-file-input");
    const statusDiv = document.getElementById("classify-status");
    const btn = document.getElementById("analysis-classify-btn");
    const tipDiv = document.getElementById("api-limit-tip");
    const format = document.getElementById("analysis-export-format").value;
    const categories = document
      .getElementById("analysis-categories")
      .value.trim();

    if (!fileInput.files.length) return alert("请先上传文件");
    const file = fileInput.files[0];

    btn.disabled = true;
    btn.innerText = "分类中...";
    statusDiv.innerText = "正在读取文件...";
    if (tipDiv) tipDiv.style.display = "none";

    try {
      const text = await readFileData(file);
      let data = parseUploadedData(text, file.name);

      if (!data || data.length === 0) throw new Error("解析数据失败或数据为空");

      // 分批处理逻辑
      const BATCH_SIZE = 20;
      const totalItems = data.length;
      const batches = Math.ceil(totalItems / BATCH_SIZE);
      const classifiedMap = new Map();

      // 自动生成的默认分类
      const defaultCats = "教程,好物分享,情感,新闻,搞笑,其他";
      const targetCats = categories || defaultCats;

      for (let i = 0; i < batches; i++) {
        const start = i * BATCH_SIZE;
        const end = Math.min((i + 1) * BATCH_SIZE, totalItems);
        const batchData = data.slice(start, end);

        statusDiv.innerText = `正在分类第 ${i + 1}/${batches} 批数据 (${start + 1}-${end})...`;

        // 构建请求数据，仅包含必要字段以节省 Token
        const sample = batchData.map((item) => ({
          id: item.笔记ID || item.note_id || item.id,
          title: item.标题 || item.display_title,
          desc: (item.desc || "").substring(0, 50),
        }));

        const prompt = `请将以下笔记归类到这些类别中：[${targetCats}]。\n输入数据：\n${JSON.stringify(
          sample,
        )}\n\n要求：返回 JSON 格式，数组包含对象 { "id": "...", "category": "..." }。`;

        try {
          const res = await callAI([
            {
              role: "system",
              content:
                "你是一个数据分类助手。只返回纯 JSON，不要 Markdown 格式。",
            },
            { role: "user", content: prompt },
          ]);

          let content = res.choices?.[0]?.message?.content || "[]";
          content = content.replace(/```json|```/g, "").trim();
          let batchResult = [];
          try {
            batchResult = JSON.parse(content);
          } catch (e) {
            console.warn(`Batch ${i + 1} JSON parse failed`, content);
          }

          if (Array.isArray(batchResult)) {
            batchResult.forEach((item) => {
              if (item && item.id) {
                classifiedMap.set(String(item.id), item.category);
              }
            });
          }
        } catch (apiError) {
          console.error(`Batch ${i + 1} API failed`, apiError);
          if (tipDiv) tipDiv.style.display = "block";
        }

        // 简单延时，避免 QPS 限制
        if (i < batches - 1) {
          await new Promise((r) => setTimeout(r, 1000));
        }
      }

      data.forEach((item) => {
        const id = item.笔记ID || item.note_id || item.id;
        if (id && classifiedMap.has(String(id))) {
          item["智能分类"] = classifiedMap.get(String(id));
        } else {
          item["智能分类"] = "未分类/API限制";
        }
      });

      const unclassifiedCount = data.filter(
        (x) => x["智能分类"] === "未分类/API限制",
      ).length;
      if (unclassifiedCount > 0 && tipDiv) {
        tipDiv.innerHTML = `完成，但有 ${unclassifiedCount} 条数据未成功分类。<br>可能是API超时或额度限制，建议检查API设置。`;
        tipDiv.style.display = "block";
      }

      // 导出新文件
      exportList(data, format, "classified_xhs_data");

      statusDiv.innerText = "分类完成！已自动下载新文件。";
    } catch (e) {
      console.error(e);
      statusDiv.innerText = "分类失败: " + e.message;
    } finally {
      btn.disabled = false;
      btn.innerText = "智能分类并导出";
    }
  }

  // ==========================================
  // 9. 营销工具 - 批量评论抓取逻辑
  // ==========================================
  let COLLECTED_COMMENTS = [];

  async function handleBatchCommentCrawl() {
    const btn = document.getElementById("mkt-run-btn");
    const progressArea = document.getElementById("mkt-progress-area");
    const progressBar = document.getElementById("mkt-progress-bar");
    const statusText = document.getElementById("mkt-status-text");
    const resultArea = document.getElementById("mkt-result-area");
    const resultCount = document.getElementById("mkt-result-count");

    // Configs
    const likeFilter =
      parseInt(document.getElementById("mkt-filter-likes").value) || 0;
    // const scope = document.getElementById('mkt-crawl-scope').value; // Currently only 'initial' supported

    if (GLOBAL_DATA.size === 0)
      return alert(
        "当前没有捕获到任何笔记。\n请先进行搜索或浏览页面，确保“数据导出”面板有数据。",
      );

    btn.disabled = true;
    progressArea.style.display = "block";
    resultArea.style.display = "none";
    COLLECTED_COMMENTS = [];

    const notes = Array.from(GLOBAL_DATA.values());
    const total = notes.length;
    let processed = 0;
    let successCount = 0;

    for (let i = 0; i < total; i++) {
      const note = notes[i];
      const percent = Math.round(((i + 1) / total) * 100);
      progressBar.style.width = percent + "%";
      statusText.innerText = `正在处理 (${i + 1}/${total}): ${note.标题.substring(0, 10)}...`;

      try {
        let noteId = note.笔记ID;
        let xsecToken = "";
        let targetUrl = note.链接 || "";

        // 尝试从链接提取 token
        if (targetUrl) {
          const match = targetUrl.match(/xsec_token=([^&]+)/);
          if (match) xsecToken = match[1];
        }

        // 如果笔记ID缺失，尝试从链接提取
        if (!noteId && targetUrl) {
          const idMatch = targetUrl.match(/\/explore\/([a-zA-Z0-9]+)/);
          if (idMatch) noteId = idMatch[1];
        }

        if (noteId) {
          let apiUrl = `https://edith.xiaohongshu.com/api/sns/web/v2/comment/page?note_id=${noteId}&cursor=&top_comment_id=&image_formats=jpg,webp,avif`;
          if (xsecToken) {
            apiUrl += `&xsec_token=${xsecToken}`;
          }

          const responseText = await new Promise((resolve) => {
            GM_xmlhttpRequest({
              method: "GET",
              url: apiUrl,
              headers: {
                "User-Agent": navigator.userAgent,
                Referer: "https://www.xiaohongshu.com/",
                Origin: "https://www.xiaohongshu.com",
              },
              onload: (res) => resolve(res.responseText),
              onerror: () => resolve(null),
            });
          });

          if (responseText) {
            const json = JSON.parse(responseText);
            // API 结构: data.comments [...]
            if (
              json.data &&
              json.data.comments &&
              Array.isArray(json.data.comments)
            ) {
              const comments = json.data.comments;

              const validComments = comments.filter((c) => {
                const likes = parseInt(c.like_count) || 0;
                return likes >= likeFilter;
              });

              validComments.forEach((c) => {
                const u = c.user_info || c.user || {};
                COLLECTED_COMMENTS.push({
                  noteId: noteId,
                  noteTitle: note.标题,
                  noteLink: targetUrl,
                  commentContent: c.content,
                  commentLikes: c.like_count || 0,
                  commentTime: c.create_time || 0,
                  userNickname: u.nickname || "未知",
                  userId: u.user_id || u.id || "",
                  userImage: u.image || "",
                });
              });
              successCount++;
            }
          }
        }
      } catch (e) {
        console.warn(`[MKTool] Failed to fetch note ${note.笔记ID}`, e);
      }

      // Random delay to avoid block (500ms - 1000ms) - API 请求频率限制
      await new Promise((r) => setTimeout(r, 500 + Math.random() * 500));
      processed++;
    }

    statusText.innerText = `完成！成功处理 ${successCount} 篇笔记，共提取 ${COLLECTED_COMMENTS.length} 条评论`;
    resultCount.innerText = COLLECTED_COMMENTS.length;
    resultArea.style.display = "block";
    btn.disabled = false;
  }
  function exportCommentData() {
    if (COLLECTED_COMMENTS.length === 0) return alert("没有可导出的评论数据");

    // Excel Format
    let html = `
            <html xmlns:o="urn:schemas-microsoft-com:office:office" 
                  xmlns:x="urn:schemas-microsoft-com:office:excel" 
                  xmlns="http://www.w3.org/TR/REC-html40">
            <head><meta charset="utf-8"></head><body>
            <table border="1">
            <thead>
                <tr style="background:#f0f0f0;">
                    <th>笔记ID</th>
                    <th>笔记标题</th>
                    <th>笔记链接</th>
                    <th>评论内容</th>
                    <th>评论点赞</th>
                    <th>评论时间</th>
                    <th>用户昵称</th>
                    <th>用户ID</th>
                </tr>
            </thead><tbody>`;

    COLLECTED_COMMENTS.forEach((row) => {
      html += `<tr>
            <td>${row.noteId}</td>
            <td>${row.noteTitle}</td>
            <td><a href="${row.noteLink}" target="_blank">链接</a></td>
            <td>${row.commentContent}</td>
            <td>${row.commentLikes}</td>
            <td>${new Date(parseInt(row.commentTime)).toLocaleString()}</td>
            <td>${row.userNickname}</td>
            <td>${row.userId}</td>
          </tr>`;
    });
    html += "</tbody></table></body></html>";

    download(
      html,
      `xhs_comments_${new Date().getTime()}.xls`,
      "application/vnd.ms-excel",
    );
  }

  // ==========================================
  // 5. 资源下载模块逻辑
  // ==========================================
  let CURRENT_NOTE_RAW = null; // 存储当前笔记原始数据，用于切换画质

  async function handleFetchResources() {
    const input = document.getElementById("res-url-input");
    const btn = document.getElementById("res-fetch-btn");
    const resultArea = document.getElementById("res-result-area");
    const grid = document.getElementById("res-grid");
    const countEl = document.getElementById("res-count");

    let url = input.value.trim();
    if (!url) {
      url = location.href;
    }

    btn.disabled = true;
    btn.innerText = "提取中... (请稍候)";
    resultArea.style.display = "none";
    grid.innerHTML = "";
    CURRENT_NOTE_RAW = null;
    // CACHED_VIDEO_URL = null; // 修改：不要清空缓存，保留用户浏览期间抓取的真实链接

    try {
      let noteData = null;

      let targetUrl = url;

      // 策略0（最优）: 页面自身加载笔记时已通过签名 API 获取过完整数据，
      // 直接复用，避免重新请求页面（SSR 现已不内嵌笔记数据，且受风控影响）
      const preId = getNoteIdFromUrl(targetUrl);
      if (preId && CACHED_NOTE_CARDS.has(preId)) {
        noteData = CACHED_NOTE_CARDS.get(preId);
        console.log("[资源下载] 命中页面已加载的笔记数据:", preId);
      }

      // 策略1: fetch 页面源码，正则提取 state（兼容性兜底）
      if (!noteData) {
        const html = await fetchHtml(targetUrl);
        noteData = extractNoteFromHtml(html);
      }

      // 策略2: 使用 noteId 调用接口兜底
      if (!noteData) {
        const noteId = getNoteIdFromUrl(targetUrl);
        noteData = await fetchNoteDetailViaApi(noteId);
      }

      if (!noteData)
        throw new Error(
          "无法提取笔记数据，请确认链接有效且为公开笔记。提示：请先在浏览器中正常打开该笔记页面，再点击提取资源",
        );

      // 补充步骤: 如果是当前页，尝试从 DOM 提取额外视频链接 (参考 Demo)
      if (
        noteData.type === "video" &&
        (targetUrl.includes(location.pathname) || targetUrl === location.href)
      ) {
        const domVideo = await extractVideoFromDOM();
        if (domVideo) {
          if (!noteData.video) noteData.video = {};
          // 将 DOM 提取的 URL 存入 temporary field
          noteData.video.dom_url = domVideo;
        }
      }

      // 等待 100ms 让 XHR 拦截器有时间捕获缓存
      await new Promise((r) => setTimeout(r, 100));

      CURRENT_NOTE_RAW = noteData;
      renderResourceGrid(noteData);
      resultArea.style.display = "block";
    } catch (e) {
      alert("提取失败: " + e.message);
      console.error(e);
    } finally {
      btn.disabled = false;
      btn.innerText = "提取资源";
    }
  }

  function fetchHtml(url) {
    return new Promise((resolve, reject) => {
      GM_xmlhttpRequest({
        method: "GET",
        url: url,
        headers: {
          "User-Agent": navigator.userAgent,
          Referer: "https://www.xiaohongshu.com/",
        },
        onload: (res) => resolve(res.responseText),
        onerror: (err) => reject(new Error("网络请求失败")),
      });
    });
  }

  function extractNoteFromHtml(html) {
    // 尝试多种写法的 __INITIAL_STATE__，避免因前端调整导致解析失败
    const candidates = [];

    // 1) 直接对象形式 window.__INITIAL_STATE__ = {...}</script>
    const directMatch = html.match(
      /__INITIAL_STATE__=({[\s\S]*?})\s*<\/script>/,
    );
    if (directMatch && directMatch[1]) {
      candidates.push(directMatch[1]);
    }

    // 2) JSON.parse(decodeURIComponent("...")) 形式
    const decodeMatch = html.match(
      /__INITIAL_STATE__=JSON\.parse\(decodeURIComponent\("([^"]+)"\)\)/,
    );
    if (decodeMatch && decodeMatch[1]) {
      try {
        candidates.push(decodeURIComponent(decodeMatch[1]));
      } catch (e) {
        console.warn("decodeURIComponent 失败", e);
      }
    }

    // 逐个尝试解析
    for (const raw of candidates) {
      try {
        const jsonStr = raw.replace(/undefined/g, "null");
        const state = JSON.parse(jsonStr);
        if (state && state.note && state.note.noteDetailMap) {
          const keys = Object.keys(state.note.noteDetailMap);
          if (keys.length > 0) return state.note.noteDetailMap[keys[0]];
        }
        if (state && state.note && state.note.firstNoteId) {
          return state.note.noteDetailMap[state.note.firstNoteId];
        }
      } catch (e) {
        console.warn("__INITIAL_STATE__ 解析失败，尝试下一个候选", e);
      }
    }

    return null;
  }

  function getNoteIdFromUrl(url) {
    const match = url.match(/explore\/([a-zA-Z0-9_-]+)/);
    return match ? match[1] : null;
  }

  async function fetchNoteDetailViaApi(noteId) {
    if (!noteId) return null;
    // 使用 web feed 接口兜底（需登录态，但与页面一致，成功率更高）
    const apiUrl = `https://edith.xiaohongshu.com/api/sns/web/v1/feed?source=explore_note&note_id=${noteId}`;
    try {
      const res = await new Promise((resolve, reject) => {
        GM_xmlhttpRequest({
          method: "GET",
          url: apiUrl,
          headers: {
            "Content-Type": "application/json",
            Referer: "https://www.xiaohongshu.com/",
            "User-Agent": navigator.userAgent,
          },
          onload: (r) => {
            try {
              resolve(JSON.parse(r.responseText));
            } catch (e) {
              reject(e);
            }
          },
          onerror: () => reject(new Error("接口请求失败")),
        });
      });

      if (res && res.data && res.data.items && res.data.items[0]) {
        const card = res.data.items[0].note_card || res.data.items[0].note;
        return card || null;
      }
    } catch (e) {
      console.warn("接口兜底获取失败", e);
    }
    return null;
  }

  // 参考视频解析demo.js: 尝试从 DOM 提取视频链接
  function extractVideoFromDOM() {
    return new Promise((resolve) => {
      const video = document.querySelector("video");
      if (!video) return resolve(null);

      // 1. 检查 src
      if (video.src && video.src.startsWith("http")) {
        return resolve(video.src);
      }

      // 2. 检查 source 标签
      const sources = video.querySelectorAll("source");
      for (let s of sources) {
        if (s.src && s.src.startsWith("http")) return resolve(s.src);
      }

      // 3. Demo 策略: 轮询 source (针对动态注入)
      let checks = 0;
      const timer = setInterval(() => {
        checks++;
        const dynamicSources = video.querySelectorAll("source");
        for (let s of dynamicSources) {
          if (s.src && s.src.startsWith("http")) {
            clearInterval(timer);
            resolve(s.src);
            return;
          }
        }
        if (checks > 20) {
          // 2秒超时
          clearInterval(timer);
          resolve(null);
        }
      }, 100);
    });
  }

  function renderResourceGrid(data) {
    const note = data.note || data;
    const grid = document.getElementById("res-grid");
    const countEl = document.getElementById("res-count");
    const qualityMode = document.getElementById("res-quality-select").value; // 'no_wm' or 'best_wm'

    grid.innerHTML = ""; // 清空旧数据
    const resources = [];
    const noteId = note.noteId || note.id || "unknown";

    // 常量定义 (参考 Demo)
    const imageServer = [
      "https://sns-img-hw.xhscdn.net/",
      "https://sns-img-bd.xhscdn.com/",
      "https://sns-img-qc.xhscdn.com/",
      "https://ci.xiaohongshu.com/",
    ];
    // 正则提取图片 Key
    const keyReg = /(?<=\/)(spectrum\/)?[a-z0-9A-Z\-]+(?=!)/;

    const videoServer = [
      "https://sns-video-hw.xhscdn.com/",
      "https://sns-video-bd.xhscdn.com/",
      "https://sns-video-al.xhscdn.com/",
    ];

    // 1. 视频处理
    if (note.type === "video" && note.video) {
      let videoUrl = "";
      let coverUrl = "";
      // 兼容 imageList / images_list
      const imgs = note.images_list || note.imageList || [];
      if (imgs.length > 0 && imgs[0])
        coverUrl = imgs[0].url || imgs[0].url_default || "";

      // 【新增】策略 Special: 优先使用捕获到的 Stream 无水印 MP4
      if (
        CACHED_VIDEO_URL &&
        CACHED_VIDEO_URL.includes("/stream/") &&
        CACHED_VIDEO_URL.includes(".mp4")
      ) {
        videoUrl = CACHED_VIDEO_URL;
        console.log("[资源下载] 命中 Stream 无水印直链(VIP):", videoUrl);
      }

      // 策略 A: origin_video_key (构造无水印原片链接)
      // 注意：原片(origin)往往是H.265编码，在部分Windows设备上可能只有声音无画面
      if (
        !videoUrl &&
        note.video.consumer &&
        note.video.consumer.origin_video_key
      ) {
        // 如果用户明确想要“无水印”且接受可能的不兼容，可以使用 origin
        // 即使有可能无画面，但只要是 "no_wm" 模式，我们应该优先保证无水印。
        // 使用 https 协议，并默认使用 bd 节点
        videoUrl = `${videoServer[1]}${note.video.consumer.origin_video_key}`;
      }

      // 策略 B: 优先使用缓存的真实视频URL (来自 XHR 拦截，无水印)
      if (!videoUrl && CACHED_VIDEO_URL) {
        videoUrl = CACHED_VIDEO_URL;
        console.log(
          "[资源下载] 使用XHR捕获的无水印视频:",
          CACHED_VIDEO_URL.substring(0, 70),
        );
      }

      // 策略 C: DOM 嗅探提取 (从页面<video>标签获取)
      if (!videoUrl && note.video.dom_url) {
        videoUrl = note.video.dom_url;
      }

      // 策略 D: 从 media.stream 获取 (支持 h265/h264 master_url / masterUrl)
      // 兼容：不同 API 可能返回 snake_case 或 camelCase
      if (!videoUrl && note.video.media && note.video.media.stream) {
        const stream = note.video.media.stream;
        let h264Url = "";
        let h265Url = "";

        const getUrl = (list) => {
          if (list && list.length > 0) {
            return list[0].master_url || list[0].masterUrl || "";
          }
          return "";
        };

        h264Url = getUrl(stream.h264);
        h265Url = getUrl(stream.h265);

        // 选择逻辑：
        if (!videoUrl) {
          // 如果 origin (策略A) 没有找到链接
          // 一般来说优先 H.264 以保证兼容性（但这往往带水印）
          // 如果用户选了 no_wm 但 origin 没找到，那也没办法，只能给一个能播的。
          if (h264Url) {
            videoUrl = h264Url;
          } else if (h265Url) {
            videoUrl = h265Url;
          }
        }
      }

      // 策略 C: 再次尝试 origin_video_key (兜底)
      if (
        !videoUrl &&
        note.video.consumer &&
        note.video.consumer.origin_video_key
      ) {
        videoUrl = `http://sns-video-bd.xhscdn.com/${note.video.consumer.origin_video_key}`;
      }

      if (videoUrl) {
        resources.push({
          type: "video",
          url: videoUrl,
          cover: coverUrl,
          name: `video_${noteId}.mp4`,
        });
      }
    }

    // 2. 图片处理 (含 Live Photo)
    const images = note.images_list || note.imageList || [];
    if (images && images.length > 0) {
      images.forEach((img, index) => {
        let targetUrl = img.url_default || img.url;

        // --- 图片无水印/高清优化 ---
        if (qualityMode === "no_wm") {
          try {
            // 尝试从 URL 提取 key 并拼接原图服务器 (Demo 方案)
            // 原 URL 通常类似 http://sns-webpic-qc.xhscdn.net/2024/.../...!...
            const keyMatch = targetUrl.match(keyReg);
            if (keyMatch) {
              targetUrl = imageServer[1] + keyMatch[0]; // 使用 sns-img-bd
            } else if (img.infoList && img.infoList.length > 0) {
              // 备用方案: 寻找 WB_DFT
              const noWmBest = img.infoList.find(
                (i) => i.imageScene === "WB_DFT",
              );
              if (noWmBest && noWmBest.url) targetUrl = noWmBest.url;
            }
          } catch (e) {
            console.error("图片解析优化失败", e);
          }
        } else {
          // 此处维持原有的最高画质寻找逻辑 (通常 infoList 里会有水印大图)
          if (img.infoList && img.infoList.length > 0) {
            const best = img.infoList.find(
              (i) => i.imageScene === "CRD_WM_WEBP",
            ); // 往往最大
            if (best && best.url) targetUrl = best.url;
            else {
              const largest = img.infoList.reduce(
                (p, c) => ((p.size || 0) > (c.size || 0) ? p : c),
                img.infoList[0],
              );
              if (largest && largest.url) targetUrl = largest.url;
            }
          }
        }

        resources.push({
          type: "image",
          url: targetUrl,
          cover: targetUrl,
          name: `image_${noteId}_${index + 1}.jpg`,
        });

        // --- Live Photo (实况图) ---
        // 检查 stream 字段 (h264/h265 视频流)
        if (img.live_photo && img.stream) {
          // live_photo 可能是 boolean
          let liveUrl = "";
          let h264Url = "";
          let h265Url = "";

          // 优先使用 H.264 (兼容性好)，其次 H.265
          // 兼容：同时尝试 master_url 和 masterUrl
          const getLiveUrl = (list) => {
            if (list && list.length > 0) {
              return list[0].master_url || list[0].masterUrl || "";
            }
            return "";
          };

          h264Url = getLiveUrl(img.stream.h264);
          h265Url = getLiveUrl(img.stream.h265);

          liveUrl = h264Url || h265Url;

          if (liveUrl) {
            resources.push({
              type: "video",
              subType: "live_photo",
              url: liveUrl,
              cover: targetUrl,
              name: `live_photo_${noteId}_${index + 1}.mp4`,
            });
          }
        }
      });
    }

    countEl.innerText = `(${resources.length} 个文件)`;

    resources.forEach((res, idx) => {
      const div = document.createElement("div");
      div.className = "selected res-item"; // 默认选中
      div.dataset.url = res.url;
      div.dataset.name = res.name;
      div.setAttribute("title", "点击选中/取消");

      let inner = "";
      if (res.type === "video") {
        const label = res.subType === "live_photo" ? "实况" : "视频";
        inner = `<img src="${res.cover}"><div class="res-type">${label}</div>`;
      } else {
        inner = `<img src="${res.cover}"><div class="res-type">图片</div>`;
      }

      div.innerHTML = `
             <div class="res-check"></div>
             ${inner}
          `;

      div.onclick = (e) => {
        // 切换选中状态
        if (div.classList.contains("selected")) {
          div.classList.remove("selected");
        } else {
          div.classList.add("selected");
        }
      };

      grid.appendChild(div);
    });
  }

  // 清理文件夹名称：去除文件系统非法字符并限制长度
  function sanitizeFolderName(name) {
    let s = String(name || "")
      .replace(/[\\/:*?"<>|]/g, "_") // 文件系统非法字符
      .replace(/[\u0000-\u001f]/g, "") // 控制字符
      .replace(/\s+/g, " ")
      .trim();
    s = s.replace(/^\.+/, "").slice(0, 50).trim();
    return s;
  }

  // 获取当前笔记的文件夹名称（优先使用笔记标题，其次笔记ID）
  function getNoteFolderName() {
    const note = CURRENT_NOTE_RAW
      ? CURRENT_NOTE_RAW.note || CURRENT_NOTE_RAW
      : {};
    const byTitle = sanitizeFolderName(note.title || note.display_title);
    if (byTitle) return byTitle;
    const byId = sanitizeFolderName(note.noteId || note.id);
    return byId || "小红书资源";
  }

  async function handleBatchDownload() {
    const items = document.querySelectorAll(".res-item.selected");
    const btn = document.getElementById("res-download-btn");
    const status = document.getElementById("res-status");

    if (items.length === 0) return alert("请至少选择一个资源");

    // 所有选中资源统一放到以笔记标题命名的文件夹中
    const folder = getNoteFolderName();

    btn.disabled = true;
    const originalText = btn.innerText;
    btn.innerText = "下载中...";
    status.innerText = "正在初始化下载...";

    let success = 0;
    for (let i = 0; i < items.length; i++) {
      const url = items[i].dataset.url;
      const name = items[i].dataset.name;
      status.innerText = `正在下载 (${i + 1}/${items.length}) ${folder}/ ...`;

      try {
        // name 带目录路径，GM_download 会在下载目录下自动创建对应文件夹
        await downloadFile(url, `${folder}/${name}`);
        success++;
        // 冷却 300ms
        await new Promise((r) => setTimeout(r, 300));
      } catch (e) {
        console.error("Download fail", e);
      }
    }

    status.innerText = `完成！成功下载 ${success} 个文件（${folder}）`;
    btn.disabled = false;
    btn.innerText = originalText;
  }

  function downloadFile(url, filename) {
    return new Promise((resolve, reject) => {
      // 优先使用 GM_download (支持跨域和大文件)
      if (typeof GM_download !== "undefined") {
        GM_download({
          url: url,
          name: filename,
          saveAs: false, // 设为 true 可弹出保存框，但批量下载建议 false
          onload: () => resolve(),
          onerror: (e) =>
            reject(new Error("GM_download error: " + (e.error || "unknown"))),
          ontimeout: () => reject(new Error("GM_download timeout")),
        });
        return;
      }

      // 降级方案: Blob (可能受 CORS 限制)
      // 注意: a.download 不支持目录路径，仅保留文件名
      const baseName = filename.split("/").pop();
      GM_xmlhttpRequest({
        method: "GET",
        url: url,
        responseType: "blob",
        onload: (res) => {
          try {
            if (res.status !== 200) throw new Error("Status " + res.status);
            const blob = res.response;
            const u = URL.createObjectURL(blob);
            const a = document.createElement("a");
            a.href = u;
            a.download = baseName;
            a.style.display = "none";
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            setTimeout(() => URL.revokeObjectURL(u), 10000);
            resolve();
          } catch (e) {
            reject(e);
          }
        },
        onerror: reject,
      });
    });
  }

  setTimeout(createUI, 1500);
})();
