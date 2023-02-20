"""
OneNoteAnalyzer

Assemblyline service using the OneNoteAnalzer tool to analyze OneNote files.
"""

from __future__ import annotations

import os
from pathlib import Path
import subprocess

from assemblyline_v4_service.common.base import ServiceBase
from assemblyline_v4_service.common.request import ServiceRequest
from assemblyline_v4_service.common.result import Result, ResultImageSection, ResultSection


class OneNoteAnalyzer(ServiceBase):
    """OneNoteAnalyzer Service"""

    def __init__(self, config: dict | None) -> None:
        super().__init__(config)

    def execute(self, request: ServiceRequest) -> None:
        subprocess.run(
            ["wine", "OneNoteAnalyzer/OneNoteAnalyzer.exe", "--file", request.file_path],
            capture_output=True,
            check=False,
        )
        output_dir = Path(f"{request.file_path}_content/")
        if output_dir.exists():
            request.result = Result([section for section in self._make_results(request, output_dir) if section])
        else:
            request.result = Result()

    def _make_results(
        self, request: ServiceRequest, output_dir: Path
    ) -> tuple[ResultSection | None, ResultSection | None, ResultSection | None]:
        return (
            self._make_attachments_section(request, output_dir / "OneNoteAttachments"),
            self._make_preview_section(request, output_dir / f"ConvertImage_{os.path.basename(request.file_path)}.png"),
            self._make_images_section(request, output_dir / "OneNoteImages"),
        )

    def _make_attachments_section(self, request: ServiceRequest, attachments_dir: Path) -> ResultSection | None:
        if attachments_dir.exists():
            for file_path in attachments_dir.iterdir():
                if file_path.is_file():
                    request.add_extracted(str(file_path), file_path.name, "attachment extracted from onenote.")
        return None

    def _make_preview_section(self, request: ServiceRequest, preview_path: Path) -> ResultImageSection | None:
        if preview_path.exists():
            preview_section = ResultImageSection(request, "OneNote File Image Preview.")
            if preview_section.add_image(
                str(preview_path), name=preview_path.name, description="OneNote file converted to PNG."
            ):
                return preview_section
        return None

    def _make_images_section(self, request: ServiceRequest, images_dir: Path) -> ResultImageSection | None:
        if images_dir.exists():
            images_section = ResultImageSection(request, "OneNote Embedded Images")
            if any(
                images_section.add_image(
                    str(image_path), name=image_path.name, description="image extracted from OneNote."
                )
                for image_path in images_dir.iterdir()
                if image_path.is_file()
            ):
                return images_section
        return None
