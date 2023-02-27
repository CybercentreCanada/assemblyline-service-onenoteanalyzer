"""
OneNoteAnalyzer

Assemblyline service using the OneNoteAnalzer tool to analyze OneNote files.
"""

from __future__ import annotations

from pathlib import Path
import subprocess

from PIL import UnidentifiedImageError

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
            self._make_preview_section(request, output_dir / f"ConvertImage_{Path(request.file_path).stem}.png"),
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
            try:
                if preview_section.add_image(
                    str(preview_path), name=preview_path.name, description="OneNote file converted to PNG."
                ):
                    return preview_section
            except UnidentifiedImageError:
                request.add_supplementary(
                    str(preview_path), name=preview_path.name, description="OneNote file converted to PNG."
                )
                return ResultSection(
                    "OneNote File Image Preview.",
                    body=(
                        "Preview was generated but could not be displayed."
                        f"\nSee supplimentary file [{preview_path.name}]"
                    ),
                )

        return None

    def _make_images_section(self, request: ServiceRequest, images_dir: Path) -> ResultImageSection | None:
        def add_image(section: ResultImageSection, path: Path) -> bool:
            """Helper function for error handling ResultImageSection.add_image()"""
            try:
                return section.add_image(str(path), name=path.name, description="image extracted from OneNote.")
            except UnidentifiedImageError:
                return False

        if images_dir.exists():
            images_section = ResultImageSection(request, "OneNote Embedded Images")
            if any(
                add_image(images_section, image_path) for image_path in images_dir.iterdir() if image_path.is_file()
            ):
                return images_section
        return None
