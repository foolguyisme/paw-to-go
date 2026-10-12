// Paw-To-Go 前端設定。三個欄位都留空時，網站會以內建商品與模擬結帳運作。
window.PTG_CONFIG = {
  // Apps Script「網頁應用程式」網址（結尾是 /exec）
  APPS_SCRIPT_URL: "",
  // Google Cloud 的 OAuth 用戶端 ID（結尾是 .apps.googleusercontent.com），需與 backend/Code.gs 相同
  GOOGLE_CLIENT_ID: "",
  // 分享 QR 使用的網址；留空則使用目前頁面網址
  SHARE_URL: ""
};
