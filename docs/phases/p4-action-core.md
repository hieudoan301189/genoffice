# P4 – Action Core hợp nhất

Ngày: 02/10/2026. Kế hoạch: [dvh-architecture-implementation-plan.md](../dvh-architecture-implementation-plan.md),
mục P4. Kiểm bằng vitest trên môi trường cloud (Linux, không có bản đóng gói Electron).

## Kết luận

- **Một catalog cho mọi ứng dụng.** Hệ thao tác sẵn có của từng ứng dụng được bọc thành action, không viết lại:
  - Docs: mỗi `OpDef` thành `Document.*`, preview là dry run của `executeOps`;
  - Sheets: mỗi thao tác của `workbookOperationSchema` thành `Spreadsheet.*`, input là schema của chính thao
    tác đó, chạy qua `planFromOps` + `applyChangePlan`;
  - Slides: mỗi op của `pptx-ops` thành `Presentation.*`, chạy qua `runTxn` / `applySessionTxn`.

  Các action Smart Data viết tay từ P1–P3 (`Data.*`, `Table.*`, `Link.*`) nằm cùng catalog. Op có tên bắt đầu
  bằng delete/remove/clear được xếp mức `destructive`.

- **Fingerprint catalog**: 12 ký tự hex trên JSON chuẩn hoá, cùng cách với `packages/cli/src/op-catalog.ts`
  (dùng hash FNV đồng bộ của `dvh-model` để renderer tính được). Kế hoạch lập trên catalog cũ bị từ chối
  (`stale_catalog`).
- **Chính sách xác nhận**: `destructive` và `external` luôn hỏi; mọi action (kể cả `write`) có preview vượt
  ngưỡng N đối tượng cũng hỏi. N nằm trong Cài đặt → Tích hợp → MCP ("Ngưỡng xác nhận", mặc định 50).
- **Transaction**: trong một ứng dụng, một action vẫn là một mục Undo. Xuyên tài liệu dùng saga
  (`runSaga`): mỗi tài liệu chụp checkpoint trước bước đầu tiên của nó; một bước lỗi thì mọi tài liệu đã chạm
  được hoàn tác về checkpoint, mới nhất trước; nhật ký ghi checkpoint, bước, lỗi, hoàn tác.
  - Checkpoint Docs: nội dung editor + model Smart Data; hoàn tác là một lần sửa Undo được.
  - Checkpoint Sheets: độ sâu ngăn Undo; hoàn tác bằng Undo về đúng độ sâu đó.
  - Checkpoint Slides: snapshot lịch sử của deck; trước khi hoàn tác đẩy trạng thái hiện tại lên ngăn Undo.
- **Mở ra ngoài**:
  - Tool MCP `dvh_list_actions`, `dvh_preview`, `dvh_execute` (`apps/shell/src/main/mcp/tools/dvh-tools.ts`).
    `dvh_execute` chạy cả kế hoạch như một saga qua các tab đang mở. Kế hoạch cần xác nhận thì shell hiện hộp
    thoại liệt kê từng bước với preview; người dùng từ chối thì không đổi gì.
  - Skill `dvh-actions` cho `agent-core` (`createDvhActionsSkill`): tool `dvh_list_actions` và `dvh_run`, gắn
    vào tác tử của Docs và Sheets. Action cần xác nhận hiện hộp thoại kèm preview.
- **Quyền**: tác tử bên ngoài (MCP) mặc định chỉ có `read`/`write`/`bulk`; `destructive`/`external` chỉ được
  cấp cho lần chạy người dùng đã xác nhận. Action không có trong catalog của tài liệu bị từ chối ngay khi
  lập kế hoạch, trước khi chạy bất cứ gì.

## Cách chạy

### Cầu nối (`packages/dvh-actions/src/bridge.ts`)

Một giao thức cho mọi ứng dụng, kiểm bằng zod ở phía nhận:

| Lệnh         | Ý nghĩa                                                                          |
| ------------ | -------------------------------------------------------------------------------- |
| `catalog`    | action, mức tác động, JSON Schema input, fingerprint, `docId`                    |
| `run`        | chạy hoặc chạy thử một action, kèm quyền được cấp và cờ "người dùng đã xác nhận" |
| `checkpoint` | chụp checkpoint, trả về token; checkpoint ở lại trong tiến trình của tài liệu    |
| `restore`    | hoàn tác về checkpoint của token                                                 |
| `release`    | bỏ checkpoint khi saga thành công                                                |

- Docs, Sheets: lệnh MCP `dvh_actions` qua cầu nối renderer sẵn có (`docs:mcp-command`, `sheets:mcp-command`).
- Slides: phiên nằm ở main process, nên endpoint dựng ngay trong shell (`apps/shell/src/main/mcp/dvh-bridge.ts`).

### Luồng `dvh_execute`

1. Phân giải từng bước tới tab (tab id hoặc đường dẫn), lấy catalog; action lạ → từ chối.
2. So fingerprint nếu được truyền; khác → từ chối.
3. Chạy thử mọi bước; bước nào lỗi → báo lỗi, không đổi gì.
4. Cần xác nhận (theo mức tác động hoặc tổng số đối tượng vượt ngưỡng) → hộp thoại; từ chối → không đổi gì.
5. Chạy saga; trả về `ok`, `txId`, số change-set, nhật ký; lỗi thì kèm bước lỗi và các tài liệu đã hoàn tác.

## Tệp chính

- `packages/dvh-actions/src`:
  - `index.ts`: nhóm `Presentation`, `fingerprint()`, `CatalogEntry`, `ConfirmPolicy`, `needsConfirmation`,
    mã lỗi `stale_catalog`;
  - `families.ts` (mới): `registerOpFamily`, `pascalCase`, `effectOfOp`;
  - `saga.ts` (mới): `runSaga`, `previewSaga`, `localParticipant`;
  - `bridge.ts` (mới): `createBridgeEndpoint`, `remoteParticipant`;
  - `agent-skill.ts` (mới): `createDvhActionsSkill`.
- Docs: `dvh-actions.ts` (adapter `Document.*`, `docsSagaParticipant`), `mcp-bridge.ts` (lệnh `dvh_actions`),
  `ai/AiPanel.tsx` (skill).
- Sheets: `dvh-actions.ts` (adapter `Spreadsheet.*`, `sheetsSagaParticipant`), `mcp-bridge.ts`, `App.tsx`
  (endpoint, skill).
- Slides: `src/main/dvh-actions.ts` (mới, adapter `Presentation.*`, `slidesSagaParticipant`); thêm phụ thuộc
  `@genoffice/dvh-actions`, `@genoffice/dvh-model` vào `dependencies` và `exclude` của `externalizeDepsPlugin`.
- Shell: `mcp/tools/dvh-tools.ts`, `mcp/dvh-bridge.ts` (mới), `mcp/app-mcp.ts` (ngưỡng trong `McpSettings`),
  `index.ts` (hộp thoại xác nhận, lưu `mcpConfirmThreshold`), preload, `McpServerSection.tsx` (ô nhập ngưỡng).
- i18n: `dvhConfirmAction` trong `i18n/dvh/*` của Docs và Sheets; 3 chuỗi hộp thoại trong main của shell;
  2 chuỗi cài đặt trong `strings.ts` của shell; đủ 20 ngôn ngữ.

## Kiểm thử (vitest)

- `packages/dvh-actions` 18 test:
  - `p4.test.ts` – **nghiệm thu Transaction**: saga 5 bước trên 2 tài liệu lỗi ở bước 4 → cả hai về checkpoint,
    nhật ký đúng thứ tự, các change-set chung `txId`; tài liệu không mở làm saga lỗi và hoàn tác; preview không
    đổi gì. **Từ chối**: action ngoài catalog, thiếu quyền, input sai, catalog cũ. Ngưỡng xác nhận.
  - `bridge.test.ts`: saga qua cầu nối (dữ liệu đi qua JSON như IPC) hoàn tác được; quyền cấp quyết định action
    destructive.
  - `agent-skill.test.ts` – **an toàn AI mức API**: action do nội dung tài liệu "yêu cầu" mà không có trong
    catalog bị từ chối; destructive chờ người dùng, từ chối thì không đổi gì; chạy thử không hỏi.
- `apps/docs/tests/dvh-actions-p4.test.ts` (2): `Document.*` (dry run không đổi gì, chạy thật là một mục Undo,
  kiểm tra của op system vẫn áp dụng); saga lỗi ở tài liệu thứ hai hoàn tác tài liệu Docs về nội dung cũ.
- `apps/sheets/tests/dvh-actions-p4.test.ts` (1): một action mỗi thao tác DSL, xếp loại destructive, preview
  rồi chạy, input sai bị từ chối.
- `apps/slides/tests/dvh-actions-p4.test.ts` (2): đủ op của `pptx-ops`; preview không đổi deck; saga lỗi khôi
  phục deck và lần khôi phục Undo được.
- `apps/shell/tests/mcp/dvh-tools.test.ts` (4): catalog + fingerprint; kế hoạch 2 tài liệu thành công và hoàn
  tác khi lỗi; action lạ, fingerprint cũ bị từ chối trước khi chạy; destructive / vượt ngưỡng phải xác nhận.
- Bộ đầy đủ: Docs 3.133/3.133; Slides 1.248/1.249 và Shell 561/565 – các ca lỗi (`slide-qc`,
  `headless-export`, `settings-integrations`) lỗi y hệt trên bản trước thay đổi; test DVH và cầu nối MCP của
  Sheets 73/73. Typecheck docs/sheets/slides/shell/dvh-actions sạch; theme-colors, english-comments, eslint,
  prettier đạt.

## Để lại cho máy Windows (chưa chạy)

1. Cấu hình một client MCP tới `http://127.0.0.1:<port>/mcp`, mở một `.docx` và một `.xlsx` có Smart Data.
   `dvh_list_actions` từng tài liệu; `dvh_execute` một kế hoạch 2 tài liệu có bước cuối lỗi (ví dụ
   `Spreadsheet.SetCell` vào sheet không tồn tại) → cả hai tài liệu trở lại như cũ, Ctrl+Z / Ctrl+Y vẫn đúng.
2. `Presentation.DeleteSlide` qua MCP → hộp thoại xác nhận; Huỷ thì deck giữ nguyên.
3. Đổi ngưỡng trong Cài đặt → kế hoạch vượt ngưỡng phải hỏi.
4. Tác tử trong Docs: yêu cầu "đặt Project.Name = X" → tác tử dùng `dvh_run Data.SetField`.

## Giới hạn còn lại

- Action `Document.*` chạy `executeOps` trực tiếp, không qua lớp chặn "chỉ số khối đã cũ" của
  `executeTool` mà `apply_ops` dùng: tác tử phải đọc lại tài liệu trước khi chỉ định khối theo chỉ số.
- Preview `Document.*` không đếm được số khối khớp mà không áp dụng: tính là 1 đối tượng/op.
- Hoàn tác Sheets bằng Undo giả định không ai sửa workbook trong lúc saga chạy (saga chạy tuần tự, vài giây).
- Recorder (P8) sẽ nghe dòng sự kiện của Action Core; các lệnh ribbon chưa đi qua action vẫn chưa ghi được.
