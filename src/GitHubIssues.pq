// ============================================================================
// GitHubIssues - Power Query (M) để lấy Issues + Project Fields từ GitHub
// Chạy được cả Windows và macOS, KHÔNG cần curl, KHÔNG cần VBA.
// Dán toàn bộ nội dung này vào Advanced Editor của một Blank Query.
// ============================================================================

let
    // ===== CẤU HÌNH - SỬA Ở ĐÂY =====
    Owner     = "octocat",            // GitHub username hoặc tổ chức
    Repo      = "Hello-World",         // Tên repository
    States    = "ALL",                 // "ALL", "OPEN", hoặc "CLOSED"
    MaxIssues = 100,                   // Số issue tối đa (tự pagination nếu > 100)
    Token     = "ghp_xxxxxxxxxxxxxxxxx", // Personal Access Token (scope read:project)
    // =================================

    // --- States filter cho GraphQL ---
    StatesFilter =
        if Text.Upper(States) = "OPEN" then "[OPEN]"
        else if Text.Upper(States) = "CLOSED" then "[CLOSED]"
        else "[OPEN, CLOSED]",

    // --- GraphQL query ---
    GraphQL =
        "query($owner: String!, $name: String!, $first: Int!, $after: String) {" &
        "  repository(owner: $owner, name: $name) {" &
        "    issues(first: $first, after: $after, states: " & StatesFilter & ", orderBy: {field: CREATED_AT, direction: DESC}) {" &
        "      pageInfo { hasNextPage endCursor }" &
        "      nodes {" &
        "        number title url state author { login } createdAt updatedAt closedAt" &
        "        labels(first: 10) { nodes { name } }" &
        "        assignees(first: 5) { nodes { login } }" &
        "        milestone { title }" &
        "        projectItems(first: 10) { nodes { project { title } fieldValues(first: 20) { nodes {" &
        "          ... on ProjectV2ItemFieldSingleSelectValue { field { ... on ProjectV2SingleSelectField { name } } name }" &
        "          ... on ProjectV2ItemFieldTextValue { field { ... on ProjectV2Field { name } } text }" &
        "          ... on ProjectV2ItemFieldNumberValue { field { ... on ProjectV2Field { name } } number }" &
        "          ... on ProjectV2ItemFieldDateValue { field { ... on ProjectV2Field { name } } date }" &
        "          ... on ProjectV2ItemFieldIterationValue { field { ... on ProjectV2IterationField { name } } title }" &
        "        } } } }" &
        "      }" &
        "    }" &
        "  }" &
        "}",

    // --- Hàm gọi 1 trang (POST tới GitHub GraphQL) ---
    GetPage = (cursor as nullable text, pageSize as number) as record =>
        let
            BodyRecord = [
                query = GraphQL,
                variables = [ owner = Owner, name = Repo, first = pageSize, after = cursor ]
            ],
            Response = Web.Contents(
                "https://api.github.com/graphql",
                [
                    Headers = [
                        #"Authorization" = "Bearer " & Token,
                        #"Content-Type"  = "application/json",
                        #"Accept"        = "application/json",
                        #"User-Agent"    = "PowerQuery-GitHub-Client"
                    ],
                    Content = Json.FromValue(BodyRecord)
                ]
            ),
            Json = Json.Document(Response),
            Issues = Json[data][repository][issues]
        in
            Issues,

    // --- Pagination: lặp lấy từng trang cho đến khi đủ MaxIssues hoặc hết ---
    Pages = List.Generate(
        () =>
            let iss = GetPage(null, List.Min({100, MaxIssues}))
            in [ iss = iss,
                 cur = iss[pageInfo][endCursor],
                 cont = iss[pageInfo][hasNextPage],
                 acc = List.Count(iss[nodes]) ],
        each [iss] <> null and [acc] > 0,
        each
            if [cont] and [acc] < MaxIssues then
                let iss = GetPage([cur], List.Min({100, MaxIssues - [acc]}))
                in [ iss = iss,
                     cur = iss[pageInfo][endCursor],
                     cont = iss[pageInfo][hasNextPage],
                     acc = [acc] + List.Count(iss[nodes]) ]
            else
                [ iss = null, cur = null, cont = false, acc = [acc] ],
        each [iss]
    ),

    // --- Gộp tất cả nodes từ các trang, cắt đúng MaxIssues ---
    AllNodesRaw = List.Combine(List.Transform(Pages, each _[nodes])),
    AllNodes = List.FirstN(AllNodesRaw, MaxIssues),

    // --- Helper: lấy project fields của 1 node thành record [Project] Field = Value ---
    ProjectFieldsRecord = (node as record) as record =>
        let
            piNodes = try node[projectItems][nodes] otherwise {},
            Pairs = List.Combine(List.Transform(piNodes, (pi) =>
                let
                    ProjTitle = try pi[project][title] otherwise "",
                    fvNodes = try pi[fieldValues][nodes] otherwise {}
                in
                    List.Transform(fvNodes, (fv) =>
                        let
                            FName = try fv[field][name] otherwise null,
                            FVal =
                                if Record.HasFields(fv, "text")   then fv[text]
                                else if Record.HasFields(fv, "number") then Text.From(fv[number])
                                else if Record.HasFields(fv, "date")   then fv[date]
                                else if Record.HasFields(fv, "title")  then fv[title]
                                else if Record.HasFields(fv, "name")   then fv[name]
                                else null
                        in
                            if FName = null then null
                            else [ Name = "[" & ProjTitle & "] " & FName, Value = FVal ]
                    )
            )),
            Cleaned = List.Select(Pairs, each _ <> null),
            Result = Record.FromList(
                List.Transform(Cleaned, each _[Value]),
                List.Transform(Cleaned, each _[Name])
            )
        in
            Result,

    // --- Helper: chuyển ISO 8601 -> datetime (local), trả null nếu rỗng ---
    ToDate = (s as nullable text) as nullable datetime =>
        if s = null then null
        else try DateTimeZone.ToLocal(DateTimeZone.FromText(s)){0} otherwise
             try DateTime.FromText(s) otherwise null,

    // --- Helper: 1 node issue -> record (1 hàng) ---
    IssueRow = (node as record) as record =>
        let
            Base = [
                Number    = node[number],
                Title     = node[title],
                Url       = node[url],
                State     = node[state],
                Author    = if (try node[author][login] otherwise null) = null then "" else node[author][login],
                CreatedAt = node[createdAt],
                UpdatedAt = node[updatedAt],
                ClosedAt  = if node[closedAt] = null then "" else node[closedAt],
                Labels    = Text.Combine(List.Transform(try node[labels][nodes] otherwise {}, each [name]), ", "),
                Assignees = Text.Combine(List.Transform(try node[assignees][nodes] otherwise {}, each [login]), ", "),
                Milestone = if (try node[milestone][title] otherwise null) = null then "" else node[milestone][title]
            ],
            ProjFields = ProjectFieldsRecord(node)
        in
            Record.Combine({ Base, ProjFields }),

    Rows = List.Transform(AllNodes, each IssueRow(_)),

    // --- Gom đủ tên cột (issue fields + tất cả project fields phát hiện được) ---
    AllKeys = List.Distinct(List.Combine(List.Transform(Rows, each Record.FieldNames(_)))),

    // --- Tạo bảng ---
    OutputTable = Table.FromRecords(Rows, AllKeys, MissingField.UseNull)
in
    OutputTable
