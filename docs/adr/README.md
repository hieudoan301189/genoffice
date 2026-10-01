# Quyết định kiến trúc (ADR)

Các quyết định D1–D10 của [kế hoạch kiến trúc](../dvh-architecture-implementation-plan.md), chốt ở P0 (02/10/2026) sau 5 spike.

| ADR                                     | Quyết định                                                          | Trạng thái   |
| --------------------------------------- | ------------------------------------------------------------------- | ------------ |
| [D1](0001-model-in-ooxml-customxml.md)  | Mô hình DVH lưu trong phần customXml của OOXML                      | Đã chấp nhận |
| [D2](0002-docx-smart-field-sdt.md)      | Smart Field trong docx là content control cấp dòng có w:dataBinding | Đã chấp nhận |
| [D3](0003-xlsx-binding-hidden-names.md) | Binding xlsx bằng defined name ẩn `_dvh.*`                          | Đã chấp nhận |
| [D4](0004-ids.md)                       | Định danh ổn định                                                   | Đã chấp nhận |
| [D5](0005-core-packages.md)             | Lõi DVH trong package riêng, móc mỏng vào ứng dụng                  | Đã chấp nhận |
| [D6](0006-wrap-existing-op-catalogs.md) | Action Core bọc các catalog op sẵn có                               | Đã chấp nhận |
| [D7](0007-changesets.md)                | Mọi hành động ghi phát ChangeSet                                    | Đã chấp nhận |
| [D8](0008-link-host-in-shell.md)        | Bộ điều phối liên kết ở main process của shell                      | Đã chấp nhận |
| [D9](0009-history-embedded.md)          | Lịch sử nhúng trong file                                            | Đã chấp nhận |
| [D10](0010-generic-smart-data.md)       | Smart Data là lớp bọc chung                                         | Đã chấp nhận |
