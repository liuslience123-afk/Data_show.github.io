const AUTH_URL =
  "https://open.feishu.cn/open-apis/auth/v3/tenant_access_token/internal/";

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

  const appId = readEnv("FEISHU_APP_ID");
  const appSecret = readEnv("FEISHU_APP_SECRET");
  if (!appId || !appSecret) {
    throw new Error(
      "缺少飞书凭据：请在 GitHub 仓库 Settings -> Secrets and variables -> Actions 中配置 FEISHU_APP_ID 和 FEISHU_APP_SECRET"
    );
  }

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
