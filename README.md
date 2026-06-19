# SUSTech TIS 课表导出 SOP

## 目标

从 `全校课表（本研）` 页面抓到**完整课程 JSON**，再生成 Excel。

核心原则：**先抓真实接口 JSON，不要先抓 DOM 表格。**

## 用户做什么

1. 在校园网或 VPN 环境里登录教务系统。
2. 打开 `全校课表（本研）`，切到目标学期，例如 `2026秋季`。
3. 打开浏览器 Console，运行下面脚本。
4. 手动点击一次 `查询`。
5. 等待浏览器自动下载合并后的 JSON。
6. 把 JSON 上传给 agent。

## agent 做什么

1. 检查 JSON 是否为全量：
   - 看 `total`
   - 看 `mergedCount`
   - 看 `list.length`
2. 如果不是全量，让用户重新抓。
3. 如果是全量，用 JSON 生成 Excel。
4. 按需要输出：
   - 全量版：`exports/sustech_2026_fall_all_courses_full.xlsx`
   - 本科版：`exports/sustech_2026_fall_undergrad_courses.xlsx`
5. 交付前检查：
   - 全量版行数正确
   - 本科版没有 `研 / 研究生`

## 抓取 JSON 的 Console 代码

把下面整段粘到浏览器 Console 里，再点一次 `查询`：

```javascript
(() => {
  const PAGE_KEYS = ['pageNum', 'page', 'pageNo', 'current'];
  const SIZE_KEYS = ['pageSize', 'limit', 'size', 'rows'];

  const looksLikeCourseResp = text => {
    try {
      const j = JSON.parse(text);
      return !!(j && typeof j === 'object' && j.rwList && Array.isArray(j.rwList.list) && typeof j.total !== 'undefined');
    } catch {
      return false;
    }
  };

  const parseJSON = text => {
    try { return JSON.parse(text); } catch { return null; }
  };

  const cloneHeaders = headersObj => {
    const out = {};
    for (const [k, v] of Object.entries(headersObj || {})) {
      const key = String(k).toLowerCase();
      if (['content-length', 'host', 'origin', 'referer'].includes(key)) continue;
      out[k] = v;
    }
    return out;
  };

  const bodyToText = body => {
    if (body == null) return null;
    if (typeof body === 'string') return body;
    if (body instanceof URLSearchParams) return body.toString();
    if (body instanceof FormData) {
      const usp = new URLSearchParams();
      for (const [k, v] of body.entries()) usp.append(k, v);
      return usp.toString();
    }
    try { return JSON.stringify(body); } catch { return null; }
  };

  const setPagingInUrl = (urlText, pageNum, pageSize) => {
    const u = new URL(urlText, location.origin);
    for (const k of PAGE_KEYS) {
      if (u.searchParams.has(k)) u.searchParams.set(k, String(pageNum));
    }
    for (const k of SIZE_KEYS) {
      if (u.searchParams.has(k)) u.searchParams.set(k, String(pageSize));
    }
    return u.toString();
  };

  const setPagingInBody = (bodyText, headers, pageNum, pageSize) => {
    if (bodyText == null) return null;
    const ct = String(headers['Content-Type'] || headers['content-type'] || '').toLowerCase();

    if (ct.includes('application/json') || bodyText.trim().startsWith('{')) {
      try {
        const obj = JSON.parse(bodyText);
        for (const k of PAGE_KEYS) obj[k] = pageNum;
        for (const k of SIZE_KEYS) obj[k] = pageSize;
        return JSON.stringify(obj);
      } catch {}
    }

    try {
      const usp = new URLSearchParams(bodyText);
      for (const k of PAGE_KEYS) usp.set(k, String(pageNum));
      for (const k of SIZE_KEYS) usp.set(k, String(pageSize));
      return usp.toString();
    } catch {}

    return bodyText;
  };

  async function replayAllPages(meta, firstJson) {
    const total = Number(firstJson.total || firstJson.rwList?.total || 0);
    const pageSize = Number(firstJson.pageSize || firstJson.rwList?.pageSize || 500);
    const pages = Number(firstJson.rwList?.pages || Math.ceil(total / pageSize) || 1);

    const all = [];
    for (let p = 1; p <= pages; p++) {
      const headers = cloneHeaders(meta.headers);
      const finalUrl = setPagingInUrl(meta.url, p, pageSize);
      const finalBody = meta.method === 'GET' ? undefined : setPagingInBody(meta.bodyText, headers, p, pageSize);

      const resp = await fetch(finalUrl, {
        method: meta.method,
        headers,
        body: finalBody,
        credentials: 'include'
      });

      const text = await resp.text();
      const json = parseJSON(text);
      if (!json || !json.rwList || !Array.isArray(json.rwList.list)) {
        throw new Error(`第 ${p} 页返回不是预期课程 JSON`);
      }
      all.push(...json.rwList.list);
    }

    const merged = {
      total,
      pageSize,
      pages,
      mergedCount: all.length,
      list: all
    };

    const blob = new Blob(
      [JSON.stringify(merged, null, 2)],
      { type: 'application/json;charset=utf-8' }
    );
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = `tis_course_all_pages_${Date.now()}.json`;
    a.click();

    console.log(`全部完成：合并后 ${all.length} 条，已下载 JSON 文件。`);
  }

  let triggered = false;
  const origOpen = XMLHttpRequest.prototype.open;
  const origSend = XMLHttpRequest.prototype.send;
  const origSetHeader = XMLHttpRequest.prototype.setRequestHeader;

  XMLHttpRequest.prototype.open = function(method, url, ...rest) {
    this.__tis_meta = { method: method || 'GET', url, headers: {} };
    return origOpen.call(this, method, url, ...rest);
  };

  XMLHttpRequest.prototype.setRequestHeader = function(k, v) {
    if (this.__tis_meta) this.__tis_meta.headers[k] = v;
    return origSetHeader.call(this, k, v);
  };

  XMLHttpRequest.prototype.send = function(body) {
    if (this.__tis_meta) this.__tis_meta.bodyText = bodyToText(body);

    this.addEventListener('load', async function() {
      if (triggered) return;
      const text = this.responseText || '';
      if (!looksLikeCourseResp(text)) return;

      triggered = true;
      const json = JSON.parse(text);
      console.log('已捕获课程接口，开始自动抓全量...');
      try {
        await replayAllPages(this.__tis_meta, json);
      } catch (e) {
        console.error('自动抓全量失败：', e);
      }
    });

    return origSend.call(this, body);
  };

  console.log('监听已安装。现在请手动点一次页面上的“查询”。');
})();
```

## 常见坑

1. 云端 agent 登录不上：通常是校园网 / VPN 限制，不要硬试，改走用户本地浏览器抓 JSON。
2. 抓 DOM 行数不对：页面可能有克隆层或虚拟列表，直接改抓 API JSON。
3. 上传的 JSON 只有第 1 页：如果 `total` 明显大于 `list.length`，说明不是全量。
