/**
 * GitHub Issues -> Excel (Office Scripts / TypeScript)
 * Chạy được Excel Windows + Mac + Web. Không cần VBA, không cần curl.
 *
 * Cách dùng:
 *  1. Excel -> tab Automate -> New Script
 *  2. Dán toàn bộ file này
 *  3. Sửa CONFIG bên dưới (owner, repo, token, states, maxIssues, columns)
 *  4. Bấm Run
 */

async function main(workbook: ExcelScript.Workbook) {
  // ===== CONFIG - SỬA Ở ĐÂY =====
  const owner = "octocat";
  const repo = "Hello-World";
  const token = "ghp_xxxxxxxxxxxxxxxxx"; // PAT, scope read:project
  const states: "ALL" | "OPEN" | "CLOSED" = "ALL";
  const maxIssues = 100; // > 100 sẽ tự pagination
  // columns: để [] -> xuất tất cả (11 issue fields + project fields tìm thấy)
  // Hoặc liệt kê: ["Number","Title","State","[Sprint Board] Status"]
  const columns: string[] = [];
  // ================================

  const result = await getRepoIssues(owner, repo, token, states, maxIssues);

  // Tạo / thay thế sheet theo tên repo
  const sheetName = safeSheetName(`${owner}_${repo}`);
  let sheet = workbook.getWorksheet(sheetName);
  if (sheet) {
    sheet.delete();
  }
  sheet = workbook.addWorksheet(sheetName);
  sheet.activate();

  // Xác định danh sách cột
  const issueFields = ["Number", "Title", "Url", "State", "Author",
    "CreatedAt", "UpdatedAt", "ClosedAt", "Labels", "Assignees", "Milestone"];

  let cols: string[];
  if (columns.length > 0) {
    cols = columns;
  } else {
    // Tự phát hiện project fields
    const pf = new Set<string>();
    for (const iss of result) {
      for (const k of Object.keys(iss.projectFields)) pf.add(k);
    }
    cols = issueFields.concat(Array.from(pf).sort());
  }

  // Ghi header
  const data: (string | number)[][] = [];
  data.push(cols);

  // Ghi từng dòng
  for (const iss of result) {
    const row: (string | number)[] = cols.map((c) => {
      switch (c.toUpperCase()) {
        case "NUMBER": return iss.number;
        case "TITLE": return iss.title;
        case "URL": return iss.url;
        case "STATE": return iss.state;
        case "AUTHOR": return iss.author;
        case "CREATEDAT": return iss.createdAt;
        case "UPDATEDAT": return iss.updatedAt;
        case "CLOSEDAT": return iss.closedAt;
        case "LABELS": return iss.labels;
        case "ASSIGNEES": return iss.assignees;
        case "MILESTONE": return iss.milestone;
        default: return iss.projectFields[c] ?? "";
      }
    });
    data.push(row);
  }

  // Đổ ra sheet
  const range = sheet.getRangeByIndexes(0, 0, data.length, cols.length);
  range.setValues(data);

  // Tạo Table
  const table = workbook.addTable(range, true);
  table.setName(safeTableName(`tblRepoIssues_${owner}_${repo}`));

  // AutoFit
  sheet.getUsedRange()?.getFormat().autofitColumns();

  console.log(`Đã ghi ${result.length} issues vào sheet "${sheetName}".`);
}

// ===========================================================================
// Kiểu dữ liệu issue đã chuẩn hóa
// ===========================================================================
interface Issue {
  number: number;
  title: string;
  url: string;
  state: string;
  author: string;
  createdAt: string;
  updatedAt: string;
  closedAt: string;
  labels: string;
  assignees: string;
  milestone: string;
  projectFields: { [key: string]: string };
}

// ===========================================================================
// Gọi GitHub GraphQL API (có pagination)
// ===========================================================================
async function getRepoIssues(
  owner: string,
  repo: string,
  token: string,
  states: string,
  maxIssues: number
): Promise<Issue[]> {
  const statesFilter =
    states.toUpperCase() === "OPEN" ? "[OPEN]" :
    states.toUpperCase() === "CLOSED" ? "[CLOSED]" :
    "[OPEN, CLOSED]";

  const query = `
    query($owner: String!, $name: String!, $first: Int!, $after: String) {
      repository(owner: $owner, name: $name) {
        issues(first: $first, after: $after, states: ${statesFilter}, orderBy: {field: CREATED_AT, direction: DESC}) {
          pageInfo { hasNextPage endCursor }
          nodes {
            number title url state author { login } createdAt updatedAt closedAt
            labels(first: 10) { nodes { name } }
            assignees(first: 5) { nodes { login } }
            milestone { title }
            projectItems(first: 10) { nodes { project { title } fieldValues(first: 20) { nodes {
              ... on ProjectV2ItemFieldSingleSelectValue { field { ... on ProjectV2SingleSelectField { name } } name }
              ... on ProjectV2ItemFieldTextValue { field { ... on ProjectV2Field { name } } text }
              ... on ProjectV2ItemFieldNumberValue { field { ... on ProjectV2Field { name } } number }
              ... on ProjectV2ItemFieldDateValue { field { ... on ProjectV2Field { name } } date }
              ... on ProjectV2ItemFieldIterationValue { field { ... on ProjectV2IterationField { name } } title }
            } } } }
          }
        }
      }
    }`;

  const out: Issue[] = [];
  let cursor: string | null = null;
  let hasNext = true;

  while (hasNext && out.length < maxIssues) {
    const pageSize = Math.min(100, maxIssues - out.length);
    const body = JSON.stringify({
      query,
      variables: { owner, name: repo, first: pageSize, after: cursor },
    });

    const resp = await fetch("https://api.github.com/graphql", {
      method: "POST",
      headers: {
        "Authorization": `Bearer ${token}`,
        "Content-Type": "application/json",
        "Accept": "application/json",
        "User-Agent": "OfficeScripts-GitHub-Client",
      },
      body,
    });

    if (resp.status !== 200) {
      throw new Error(`HTTP ${resp.status}: ${await resp.text()}`);
    }

    const json = await resp.json() as GraphQLResponse;
    if (json.errors && json.errors.length > 0) {
      throw new Error(`GraphQL: ${json.errors[0].message}`);
    }
    const issues = json.data?.repository?.issues;
    if (!issues) {
      throw new Error(`NOT_FOUND: ${owner}/${repo}`);
    }

    for (const node of issues.nodes) {
      out.push(mapNode(node));
    }

    hasNext = issues.pageInfo.hasNextPage;
    cursor = issues.pageInfo.endCursor;
    if (issues.nodes.length === 0) break;
  }

  return out.slice(0, maxIssues);
}

// ===========================================================================
// Chuyển 1 node JSON -> Issue
// ===========================================================================
function mapNode(node: IssueNode): Issue {
  const projectFields: { [key: string]: string } = {};
  for (const pi of node.projectItems?.nodes ?? []) {
    const projTitle = pi.project?.title ?? "";
    for (const fv of pi.fieldValues?.nodes ?? []) {
      const fname = fv.field?.name;
      if (!fname) continue;
      let val = "";
      if (fv.text !== undefined) val = fv.text;
      else if (fv.number !== undefined) val = String(fv.number);
      else if (fv.date !== undefined) val = fv.date;
      else if (fv.title !== undefined) val = fv.title;
      else if (fv.name !== undefined) val = fv.name;
      projectFields[`[${projTitle}] ${fname}`] = val;
    }
  }

  return {
    number: node.number,
    title: node.title,
    url: node.url,
    state: node.state,
    author: node.author?.login ?? "",
    createdAt: node.createdAt ?? "",
    updatedAt: node.updatedAt ?? "",
    closedAt: node.closedAt ?? "",
    labels: (node.labels?.nodes ?? []).map((l) => l.name).join(", "),
    assignees: (node.assignees?.nodes ?? []).map((a) => a.login).join(", "),
    milestone: node.milestone?.title ?? "",
    projectFields,
  };
}

// ===========================================================================
// Helpers
// ===========================================================================
function safeSheetName(raw: string): string {
  let s = raw.replace(/[:\\/?*\[\]]/g, "_");
  if (s.length > 31) s = s.substring(0, 31);
  return s.length === 0 ? "Issues" : s;
}

function safeTableName(raw: string): string {
  let s = raw.replace(/[^A-Za-z0-9_]/g, "_");
  if (/^[0-9]/.test(s)) s = "_" + s;
  return s;
}

// ===========================================================================
// Kiểu dữ liệu cho response GraphQL (giúp TypeScript hiểu cấu trúc JSON)
// ===========================================================================
interface GraphQLResponse {
  data?: { repository?: { issues: IssuesConnection } };
  errors?: { message: string; type?: string }[];
}
interface IssuesConnection {
  pageInfo: { hasNextPage: boolean; endCursor: string };
  nodes: IssueNode[];
}
interface IssueNode {
  number: number;
  title: string;
  url: string;
  state: string;
  author?: { login: string };
  createdAt?: string;
  updatedAt?: string;
  closedAt?: string;
  labels?: { nodes: { name: string }[] };
  assignees?: { nodes: { login: string }[] };
  milestone?: { title: string };
  projectItems?: { nodes: ProjectItemNode[] };
}
interface ProjectItemNode {
  project?: { title: string };
  fieldValues?: { nodes: FieldValueNode[] };
}
interface FieldValueNode {
  field?: { name: string };
  name?: string;
  text?: string;
  number?: number;
  date?: string;
  title?: string;
}
