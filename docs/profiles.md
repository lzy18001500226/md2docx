# Profiles

Profiles keep template-specific choices out of the converter. Each TOML profile
has a reference DOCX and the labels used for figure and table captions.

```toml
reference_docx = "../templates/public-demo.docx"
figure_label = "图"
table_label = "表"
code_style = "Normal"
```

Paths are resolved relative to the profile file. A profile is the unit to share
with a team. It can be public only when its reference DOCX and all example
assets are redistributable.
