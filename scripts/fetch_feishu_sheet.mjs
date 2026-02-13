import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { feishuFetch } from "./feishu_auth.mjs";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const rootDir = path.resolve(__dirname, "..");

const {
  FEISHU_SHEET_TOKEN,
  COVER_RANGE = "bfd5ff!A1:J20",
  PAGE1_RANGE = "zBgHxt!A1:Z200",
  PAGE2_SHEET_NAME = "2",
  PAGE2_SHEET_ID = "jonA37",
  PAGE2_RANGE = "A1:Z200"
} = process.env;

if (!FEISHU_SHEET_TOKEN) {
  console.error("缺少环境变量：FEISHU_SHEET_TOKEN");
  process.exit(1);
}

const DATA_DIR = path.join(rootDir, "data");
const OUTPUT_PATH = path.join(DATA_DIR, "map_data.json");

async function fetchRange(range) {
  const url = `https://open.feishu.cn/open-apis/sheets/v2/spreadsheets/${FEISHU_SHEET_TOKEN}/values/${encodeURIComponent(
    range
  )}`;
  const resp = await feishuFetch(url);
  const json = await resp.json();
  if (json.code !== 0) {
    throw new Error(`飞书 API 返回异常：${json.code} - ${json.msg}`);
  }
  return json.data?.valueRange?.values || [];
}

let cachedSheetsMeta = null;

async function fetchSheetsMeta() {
  if (cachedSheetsMeta) return cachedSheetsMeta;
  const url = `https://open.feishu.cn/open-apis/sheets/v3/spreadsheets/${FEISHU_SHEET_TOKEN}/sheets/query`;
  const resp = await feishuFetch(url);
  const json = await resp.json();
  if (json.code !== 0) {
    throw new Error(`获取 sheet 列表失败：${json.code} - ${json.msg}`);
  }
  cachedSheetsMeta = json.data?.sheets ?? [];
  return cachedSheetsMeta;
}

async function resolveSheetRange({ sheetTitle, sheetId, cellRange }) {
  const cleanRange = (cellRange || "A1:Z200").replace(/^!/, "");
  if (sheetId) return `${sheetId}!${cleanRange}`;
  if (!sheetTitle) {
    throw new Error("需要提供 sheet 的标题或 sheet_id");
  }
  const sheets = await fetchSheetsMeta();
  const target = sheets.find(sheet => (sheet?.title ?? "").trim() === sheetTitle.trim());
  if (!target) {
    const available = sheets.map(s => s?.title ?? "").join(", ");
    throw new Error(`未找到名称为 \"${sheetTitle}\" 的 sheet，可用：${available}`);
  }
  return `${target.sheet_id}!${cleanRange}`;
}

async function resolvePageRange(pageNum) {
  const envKey = `PAGE${pageNum}_SHEET_ID`;
  const fromEnv = (process.env[envKey] ?? "").trim();
  if (fromEnv) {
    return `${fromEnv}!A1:Z300`;
  }
  const sheets = await fetchSheetsMeta();
  const titles = [String(pageNum), `Page${pageNum}`];
  const target = sheets.find(sheet => titles.includes((sheet?.title ?? "").trim()));
  if (target) return `${target.sheet_id}!A1:Z300`;
  return null;
}

async function main() {
  try {
    const coverValues = await fetchRange(COVER_RANGE);
    const page1Values = await fetchRange(PAGE1_RANGE);

    const page2Range = await resolveSheetRange({
      sheetTitle: PAGE2_SHEET_NAME,
      sheetId: PAGE2_SHEET_ID,
      cellRange: PAGE2_RANGE
    });
    const page2Values = await fetchRange(page2Range);

    const pages = {};
    for (let n = 3; n <= 13; n++) {
      const range = await resolvePageRange(n);
      if (!range) continue;
      try {
        const values = await fetchRange(range);
        pages[`page${n}`] = { range, values };
      } catch (err) {
        console.error(`读取页面 ${n} 失败：`, err.message || err);
      }
    }

    const payload = {
      spreadsheetToken: FEISHU_SHEET_TOKEN,
      cover: { range: COVER_RANGE, values: coverValues },
      page1: { range: PAGE1_RANGE, values: page1Values },
      page2: { range: page2Range, values: page2Values },
      pages,
      fetchedAt: new Date().toISOString()
    };

    await fs.mkdir(DATA_DIR, { recursive: true });
    await fs.writeFile(OUTPUT_PATH, JSON.stringify(payload, null, 2), "utf8");
    console.log(`飞书数据已保存到 ${OUTPUT_PATH}`);
  } catch (error) {
    console.error("读取飞书表格或保存数据时出错：", error.message || error);
    process.exit(1);
  }
}

main();

import fs from "node:fs/promises";
import { resolveTenantToken } from "./feishu_auth.mjs";

const {
  FEISHU_SHEET_TOKEN,
  FEISHU_RANGE = "Sheet1!A1:Z50",
  FEISHU_OUTPUT = "",
  FEISHU_LIST_SHEETS,
  FEISHU_DIRECT_TOKEN
} = process.env;

if (!FEISHU_SHEET_TOKEN) {
  console.error("缺少环境变量：FEISHU_SHEET_TOKEN");
  process.exit(1);
}

async function getSheetRange(tenantToken, sheetToken, range) {
  const url = new URL(
    `https://open.feishu.cn/open-apis/sheets/v2/spreadsheets/${sheetToken}/values_batch_get`
  );
  url.searchParams.append("ranges", range);

  const resp = await fetch(url, {
    method: "GET",
    headers: {
      Authorization: `Bearer ${tenantToken}`,
      "Content-Type": "application/json"
    }
  });
  const data = await resp.json();
  if (data.code !== 0) {
    throw new Error(`读取表格失败：${data.code} - ${data.msg}`);
  }
  return data.data?.valueRanges ?? [];
}

async function getSheetMeta(tenantToken, sheetToken) {
  const resp = await fetch(
    `https://open.feishu.cn/open-apis/sheets/v3/spreadsheets/${sheetToken}`,
    {
      method: "GET",
      headers: {
        Authorization: `Bearer ${tenantToken}`,
        "Content-Type": "application/json"
      }
    }
  );
  const data = await resp.json();
  if (data.code !== 0) {
    throw new Error(`获取表格信息失败：${data.code} - ${data.msg}`);
  }
  return data.data?.sheets ?? [];
}

async function main() {
  try {
    const tenantToken = await resolveTenantToken();

    if (FEISHU_DIRECT_TOKEN) {
      const directData = await getDirectRange(
        FEISHU_DIRECT_TOKEN,
        FEISHU_SHEET_TOKEN,
        FEISHU_RANGE
      );
      const payload = JSON.stringify(directData, null, 2);
      if (FEISHU_OUTPUT) {
        await fs.writeFile(FEISHU_OUTPUT, payload, "utf8");
        console.log(`数据已写入 ${FEISHU_OUTPUT}`);
      } else {
        console.log(payload);
      }
      return;
    }

    if (FEISHU_LIST_SHEETS) {
      const sheets = await getSheetMeta(tenantToken, FEISHU_SHEET_TOKEN);
      console.log("可用 sheet 列表：");
      sheets.forEach(sheet => {
        console.log(
          `- title: ${sheet?.title ?? "未知"}\n  sheet_id: ${sheet?.sheet_id}\n  grid_id: ${sheet?.grid_id}`
        );
      });
      return;
    }

    const ranges = await getSheetRange(tenantToken, FEISHU_SHEET_TOKEN, FEISHU_RANGE);

    if (FEISHU_OUTPUT) {
      await fs.writeFile(
        FEISHU_OUTPUT,
        JSON.stringify(ranges, null, 2),
        "utf8"
      );
      console.log(`已写入 ${FEISHU_OUTPUT}`);
    } else {
      console.log(JSON.stringify(ranges, null, 2));
    }
  } catch (error) {
    console.error(error);
    process.exit(1);
  }
}

async function getDirectRange(token, spreadsheetToken, range) {
  const encodedRange = encodeURIComponent(range);
  const url = `https://open.feishu.cn/open-apis/sheets/v2/spreadsheets/${spreadsheetToken}/values/${encodedRange}`;
  const resp = await fetch(url, {
    method: "GET",
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json"
    }
  });
  const data = await resp.json();
  if (data.code !== 0) {
    throw new Error(`读取表格失败：${data.code} - ${data.msg}`);
  }
  return data.data;
}

main();
