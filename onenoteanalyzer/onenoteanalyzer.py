"""
OneNoteAnalyzer

Assemblyline service using the OneNoteAnalzer tool to analyze OneNote files.
"""

from __future__ import annotations

import os
import subprocess

from pathlib import Path
from collections import defaultdict

from PIL import UnidentifiedImageError

from assemblyline.common import forge
from assemblyline.common.str_utils import safe_str
from assemblyline_v4_service.common.balbuzard.patterns import PatternMatch
from assemblyline_v4_service.common.base import ServiceBase

from assemblyline_v4_service.common.extractor.ocr import detections
from assemblyline_v4_service.common.request import ServiceRequest
from assemblyline_v4_service.common.result import (
    Heuristic,
    KVSectionBody,
    Result,
    ResultImageSection,
    ResultSection,
)
from assemblyline_v4_service.common.utils import extract_passwords


class OneNoteAnalyzer(ServiceBase):
    """OneNoteAnalyzer Service"""

    LAUNCHABLE_EXTENSIONS = {
        ".ade",
        ".adp",
        ".as",  # Adobe ActionScript
        ".bat",  # DOS/Windows batch file
        ".chm",
        ".cmd",  # Windows command
        ".com",  # DOS command
        ".cpl",
        ".exe",  # DOS/Windows executable
        ".dll",  # Windows library
        ".hta",
        ".inf",  # Windows autorun file
        ".ins",
        ".isp",
        ".jar",  # Java JAR
        ".jse",
        ".js",  # Javascript
        ".lib",
        ".lnk",  # Windows shortcut
        ".mde",
        ".msc",
        ".msp",
        ".mst",
        ".pif",
        ".py",  # Python script
        ".scr",  # Windows screen saver
        ".sct",
        ".shb",
        ".sys",
        ".url",  # Windows URL Shortcut
        ".vb",  # VB Script
        ".vbe",  # Encrypted VB script
        ".vbs",  # VB Script
        ".vxd",
        ".wsc",
        ".wsf",
        ".wsh",
    }

    LAUNCHABLE_TYPE = {
        "code/batch",
        "code/ps1",
        "code/python",
        "code/vbs",
    }

    LAUNCHABLE_TYPE_PREFIX = {"executable", "shortcut"}

    def __init__(self, config: dict | None = None) -> None:
        super().__init__(config)
        self.identify = forge.get_identify(use_cache=os.environ.get("PRIVILEGED", "false").lower() == "true")

    def execute(self, request: ServiceRequest) -> None:
        subprocess.run(
            [
                "wine",
                "OneNoteAnalyzer/OneNoteAnalyzer.exe",
                "--file",
                request.file_path,
            ],
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
    ) -> tuple[ResultSection | None, ResultSection | None, ResultSection | None, ResultSection | None]:
        self._make_hyperlinks_section(request, output_dir / "OneNoteHyperlinks")
        return (
            self._make_attachments_section(request, output_dir / "OneNoteAttachments"),
            self._make_preview_section(request, output_dir / f"ConvertImage_{Path(request.file_path).stem}.png"),
            self._make_images_section(request, output_dir / "OneNoteImages"),
            self._make_text_section(request, output_dir / "OneNoteText"),
        )

    def _make_attachments_section(self, request: ServiceRequest, attachments_dir: Path) -> ResultSection | None:
        if not attachments_dir.exists():
            return None
        executable_attachments: list[str] = []
        attachments: list[str] = []
        for file_path in attachments_dir.iterdir():
            if not file_path.is_file():
                continue
            request.add_extracted(
                str(file_path),
                file_path.name,
                "attachment extracted from onenote.",
            )
            file_type: str = self.identify.fileinfo(str(file_path))["type"]
            if (
                file_path.suffix in self.LAUNCHABLE_EXTENSIONS
                or file_type in self.LAUNCHABLE_TYPE
                or file_type.split("/", 1)[0] in self.LAUNCHABLE_TYPE_PREFIX
            ):
                executable_attachments.append(file_path.name)
            else:
                attachments.append(file_path.name)

        if not (attachments or executable_attachments):
            return None
        return ResultSection(
            "OneNote Attachments",
            body="Executables:\n" + "\n".join(executable_attachments) + "Other:\n" + "\n".join(attachments),
            heuristic=Heuristic(1) if executable_attachments else None,
        )

    def _make_preview_section(self, request: ServiceRequest, preview_path: Path) -> ResultImageSection | None:
        if not (preview_path.exists() and preview_path.stat().st_size):
            return None
        try:
            preview_section = ResultImageSection(request, "OneNote File Image Preview.")
            preview_section.add_image(
                str(preview_path),
                name=preview_path.name,
                description="OneNote file converted to PNG.",
            )
            return preview_section
        except UnidentifiedImageError:
            request.add_supplementary(
                str(preview_path),
                name=preview_path.name,
                description="OneNote file converted to PNG.",
            )
            return ResultSection(
                "OneNote File Image Preview.",
                body=(
                    "Preview was generated but could not be displayed."
                    f"\nSee supplimentary file [{preview_path.name}]"
                ),
            )

    def _make_images_section(self, request: ServiceRequest, images_dir: Path) -> ResultImageSection | None:
        def add_image(section: ResultImageSection, path: Path) -> bool:
            """Helper function for error handling ResultImageSection.add_image()"""
            try:
                section.add_image(
                    str(path),
                    name=path.name,
                    description="image extracted from OneNote.",
                )
                return True
            except UnidentifiedImageError:
                return False

        if not images_dir.exists():
            return None
        images_section = ResultImageSection(request, "OneNote Embedded Images")
        if any(
            add_image(images_section, image_path)
            for image_path in images_dir.iterdir()
            if image_path.is_file() and image_path.stat().st_size
        ):
            return images_section
        return None

    def _make_text_section(self, request: ServiceRequest, text_dir: Path) -> ResultSection | None:
        if not text_dir.exists():
            return None
        patterns = PatternMatch()
        results: dict[str, list[str]] = defaultdict(list)
        tags: dict[str, list[str]] = defaultdict(list)
        for page in text_dir.iterdir():
            if not page.is_file():
                continue
            with page.open("r") as f:
                text = f.read()
            if page.name.startswith("1_"):  # Keep potential passwords from the first page
                passwords = extract_passwords(text)
                if "passwords" in request.temp_submission_data:
                    request.temp_submission_data["passwords"].extend(passwords)
                else:
                    request.temp_submission_data["passwords"] = passwords
            tag_type: str
            values: set[bytes]
            for tag_type, values in patterns.ioc_match(text.encode(), True, True).items():
                tags[tag_type].extend(safe_str(tag) for tag in values)
            for detection_type, indicators in detections(text):
                results[detection_type].extend(indicators)

        if not results or tags:
            return None
        text_section = ResultSection("OneNote Text")
        ResultSection(
            "Suspicious strings found in OneNote Text",
            KVSectionBody(**results),
            heuristic=Heuristic(2, signatures={f"{k}_strings": len(v) for k, v in results.items()}),
            parent=text_section,
        )
        ResultSection(
            "Network Indicators found in OneNote Text",
            KVSectionBody(**tags),
            heuristic=Heuristic(3, signatures={k.replace(".", "_"): 1 for k, _ in tags.items()}),
            tags=tags,
            parent=text_section,
        )
        return text_section

    def _make_hyperlinks_section(self, request: ServiceRequest, hyperlinks_dir: Path) -> None:
        # I have no idea what hyperlinks is supposed to be, adding as supplimentary so we can monitor it
        if not hyperlinks_dir.exists():
            return
        expected_file = hyperlinks_dir / "onenote_hyperlinks"
        if not expected_file.exists() and expected_file.is_file():
            return
        request.add_supplementary(
            expected_file, request.sha256[:8] + expected_file.name, "OneNoteAnalyzer Hyperlinks file"
        )
