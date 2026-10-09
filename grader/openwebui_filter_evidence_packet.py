"""
title: Evidence Packet
author: Reginald Finley
version: 1.1
license: MIT
description: When a chat message carries a student's files (.docx, .sav, .spv), build the evidence packet with make_packet.py and hand the model the packet instead of the raw files. Runs before the model sees the message.
requirements: python-docx, pyreadstat, pillow
"""

import base64
import glob
import os
import shutil
import sys
import tempfile
import time

from pydantic import BaseModel, Field

GRADE_INSTRUCTION = "Grade this submission using the brief. Produce Part One, then Part Two."
STUDENT_EXTS = (".docx", ".sav", ".spv", ".pdf", ".xlsx")


class Filter:
    class Valves(BaseModel):
        grader_dir: str = Field(
            default=os.environ.get("GRADER_DIR", os.path.expanduser("~/grader")),
            description="Folder on the server holding make_packet.py, the course config, the blank workbook, and the course .sav",
        )
        packets_dir: str = Field(
            default=os.environ.get("GRADER_PACKETS_DIR", os.path.expanduser("~/grader/packets")),
            description="Where generated packets and chart images are saved for the record",
        )
        template_docx: str = Field(
            default="blank_workbook.docx",
            description="Blank workbook filename inside grader_dir (separates student text from instructions)",
        )
        source_sav: str = Field(
            default="course_data.sav",
            description="Course data file inside grader_dir (detects an unchanged .sav)",
        )
        config_file: str = Field(
            default="course_config.json",
            description="Course config JSON inside grader_dir (box names, variables, codebook, true values)",
        )
        include_charts: bool = Field(
            default=False,
            description="Also send the pasted chart images to the model (slower; the model must support images)",
        )
        max_charts: int = Field(default=7, description="Upper limit on chart images sent when include_charts is on")
        add_grade_instruction: bool = Field(
            default=True,
            description="Append the standard grading instruction if the user's message is short",
        )

    def __init__(self):
        self.valves = self.Valves()

    # ------------------------------------------------------------------ helpers

    def _resolve_path(self, f):
        """Turn an Open WebUI file record into a path on disk."""
        fid = f.get("id") or (f.get("file") or {}).get("id")
        name = f.get("name") or (f.get("file") or {}).get("filename") or (f.get("file") or {}).get("meta", {}).get("name") or ""
        # 1. the file record itself may carry a path
        for cand in (f.get("path"), (f.get("file") or {}).get("path")):
            if cand and os.path.exists(cand):
                return cand, name
        # 2. ask Open WebUI's file table
        if fid:
            try:
                from open_webui.models.files import Files  # type: ignore
                rec = Files.get_file_by_id(fid)
                if rec is not None:
                    p = getattr(rec, "path", None)
                    if p and os.path.exists(p):
                        return p, name or getattr(rec, "filename", "")
                    name = name or getattr(rec, "filename", "")
            except Exception:
                pass
            # 3. the uploads folder, files are stored as <id>_<original name>
            try:
                from open_webui.config import UPLOAD_DIR  # type: ignore
                dirs = [str(UPLOAD_DIR)]
            except Exception:
                dirs = []
            data_dir = os.environ.get("DATA_DIR", os.path.expanduser("~/open-webui/data"))
            dirs.append(os.path.join(data_dir, "uploads"))
            for d in dirs:
                hits = glob.glob(os.path.join(d, fid + "_*")) + glob.glob(os.path.join(d, fid + "*"))
                if hits:
                    return hits[0], name or os.path.basename(hits[0]).split("_", 1)[-1]
        return None, name

    def _student_files(self, body):
        files = list(body.get("files") or [])
        meta = body.get("metadata") or {}
        files += [f for f in (meta.get("files") or []) if f not in files]
        found = []
        for f in files:
            if not isinstance(f, dict):
                continue
            if f.get("type") not in (None, "file"):
                continue  # knowledge collections, web pages, etc.
            path, name = self._resolve_path(f)
            label = (name or path or "").lower()
            if path and label.endswith(STUDENT_EXTS):
                found.append((path, name or os.path.basename(path)))
        return found

    def _build(self, found):
        gd = os.path.expanduser(self.valves.grader_dir)
        if gd not in sys.path:
            sys.path.insert(0, gd)
        import importlib
        mp = importlib.import_module("make_packet")
        importlib.reload(mp)
        work = tempfile.mkdtemp(prefix="packet_")
        try:
            for path, name in found:
                shutil.copy(path, os.path.join(work, os.path.basename(name)))
            template = os.path.join(gd, self.valves.template_docx)
            source = os.path.join(gd, self.valves.source_sav)
            config = os.path.join(gd, self.valves.config_file)
            out_dir = os.path.expanduser(self.valves.packets_dir)
            text, packet_path, chart_dir = mp.build_packet(
                [work], out_dir,
                template=template if os.path.exists(template) else None,
                source_sav=source if os.path.exists(source) else None,
                config=config if os.path.exists(config) else None,
            )
            charts = []
            if chart_dir and self.valves.include_charts:
                for fn in sorted(os.listdir(chart_dir)):
                    if fn.lower().endswith(".png") and len(charts) < self.valves.max_charts:
                        with open(os.path.join(chart_dir, fn), "rb") as fh:
                            charts.append((fn, base64.b64encode(fh.read()).decode("ascii")))
            return text, packet_path, charts
        finally:
            shutil.rmtree(work, ignore_errors=True)

    # ------------------------------------------------------------------ the hook

    def inlet(self, body: dict, __user__: dict = None, __event_emitter__=None) -> dict:
        found = self._student_files(body)
        if not found:
            return body  # ordinary chat, nothing to do

        messages = body.get("messages") or []
        if not messages or messages[-1].get("role") != "user":
            return body
        user_msg = messages[-1]
        original = user_msg.get("content")
        if isinstance(original, list):
            original_text = " ".join(p.get("text", "") for p in original if p.get("type") == "text")
        else:
            original_text = original or ""
        original_text = original_text.strip()

        t0 = time.time()
        try:
            packet_text, packet_path, charts = self._build(found)
        except Exception as e:
            note = ("EVIDENCE PACKET COULD NOT BE BUILT. Tell the user this exact error and stop: %s: %s"
                    % (type(e).__name__, e))
            user_msg["content"] = note
            body["files"] = []
            if isinstance(body.get("metadata"), dict):
                body["metadata"]["files"] = []
            return body

        instruction = original_text
        if self.valves.add_grade_instruction and len(original_text.split()) < 12:
            instruction = (original_text + " " + GRADE_INSTRUCTION).strip()

        header = ("The evidence packet for this submission follows. It was built by script from the files "
                  "the user attached (%s). Grade only from it.\n\n" % ", ".join(n for _, n in found))
        new_text = instruction + "\n\n" + header + packet_text

        if charts:
            parts = [{"type": "text", "text": new_text + "\n\nThe student's pasted charts follow as images, named by box: "
                      + ", ".join(fn for fn, _ in charts)}]
            for fn, b64 in charts:
                parts.append({"type": "image_url", "image_url": {"url": "data:image/png;base64," + b64}})
            user_msg["content"] = parts
        else:
            user_msg["content"] = new_text

        # Remove the raw uploads so Open WebUI does not also paste their extracted text into context.
        body["files"] = []
        if isinstance(body.get("metadata"), dict):
            body["metadata"]["files"] = []

        # Let the user see that it happened.
        if __event_emitter__ is not None:
            try:
                import asyncio
                msg = "Evidence packet built in %.1f s from %d file(s); saved to %s" % (
                    time.time() - t0, len(found), packet_path)
                coro = __event_emitter__({"type": "status", "data": {"description": msg, "done": True}})
                if asyncio.iscoroutine(coro):
                    try:
                        loop = asyncio.get_event_loop()
                        if loop.is_running():
                            asyncio.ensure_future(coro)
                        else:
                            loop.run_until_complete(coro)
                    except Exception:
                        pass
            except Exception:
                pass
        return body
