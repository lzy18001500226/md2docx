"""Template-governed Markdown-to-DOCX delivery."""

from .pipeline import BuildProfile, BuildResult, build_document, load_profile

__all__ = ["BuildProfile", "BuildResult", "build_document", "load_profile"]
