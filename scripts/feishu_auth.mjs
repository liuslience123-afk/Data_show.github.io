const AUTH_URL =
  "https://open.feishu.cn/open-apis/auth/v3/tenant_access_token/internal/";
const DEFAULT_APP_ID = "cli_a90ef5b2abf91cb3";
const DEFAULT_APP_SECRET = "iGWYy9xqu3In0QsV7OXuEgRRdXhx3Rk0";

let cachedToken = "";
let cachedExpireAt = 0;

function readEnv(name) {
  return (process.env[name] ?? "").trim();
}

async function requestTenantAccessToken(appId, appSecret) {
  const resp = await fetch(AUTH_URL, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ app_id: appId, app_secret: appSecret })
  });
  const data = await resp.json();
  if (data.code !== 0) {
    throw new Error(`获取 tenant_access_token 失败：${data.code} - ${data.msg}`);
  }
  return {
    token: data.tenant_access_token,
    expireSeconds:
      Number(data.expire) ||
      Number(data.expire_in) ||
      Number(data.expire_time) ||
      0
  };
}

export async function resolveTenantToken() {
  const directToken = readEnv("FEISHU_DIRECT_TOKEN");
  if (directToken) return directToken;

  const appId = readEnv("FEISHU_APP_ID") || DEFAULT_APP_ID;
  const appSecret = readEnv("FEISHU_APP_SECRET") || DEFAULT_APP_SECRET;

  const now = Date.now();
  if (cachedToken && cachedExpireAt && now < cachedExpireAt) {
    return cachedToken;
  }

  const { token, expireSeconds } = await requestTenantAccessToken(
    appId,
    appSecret
  );
  cachedToken = token;
  if (expireSeconds > 0) {
    const safeExpireSeconds = Math.max(expireSeconds - 30, 0);
    cachedExpireAt = Date.now() + safeExpireSeconds * 1000;
  } else {
    cachedExpireAt = 0;
  }
  return cachedToken;
}

export async function feishuFetch(url, options = {}) {
  const token = await resolveTenantToken();
  const headers = {
    Authorization: `Bearer ${token}`,
    ...(options.headers ?? {})
  };
  return fetch(url, { ...options, headers });
}
