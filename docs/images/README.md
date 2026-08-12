# README Visual Assets

The images in this directory are project-provided Word review screenshots used
by the root README. They illustrate the document results that `md2word` is
intended to preserve and validate.

| File | README purpose |
| --- | --- |
| `目录.png` | Generated table of contents and Word navigation pane. |
| `一级标题.png` | Numbered heading hierarchy in body text. |
| `三线表.png` | Three-line table, figure placement, and captions. |
| `图.png` | Figure layout with captions and nearby explanatory text. |
| `公式1.png` | Displayed Office Math equations and equation numbering. |
| `公式2.png` | A complex multiline Office Math equation. |
| `代码.png` | Code paragraph styling alongside a rendered chart. |
| `mathtype.png` | Optional MathType editor integration in Word. |

These screenshots are visual review evidence only. They are not source assets
consumed by `examples/demo-report.md`, and they do not prove that the public
demo generated the report content shown in the screenshots. The automated test
suite verifies the package-level DOCX contract; visual acceptance remains a
separate review step.

Before a public release, confirm that the project owns or has permission to
redistribute every screenshot and its visible report content. Do not add
customer, course, institutional, competition, or third-party material without
the review described in [`../release-boundary.md`](../release-boundary.md).
