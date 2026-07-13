# Overseas Inventory Dashboard

多语言功能使用 `free-google-translate` 思路中的 Google 网页翻译接口方案，前端直接请求 `translate.googleapis.com/translate_a/single?client=gtx`，不再使用腾讯云翻译代理。

页面会把已翻译文本缓存在浏览器本地，减少重复请求，避免频繁触发 429。
