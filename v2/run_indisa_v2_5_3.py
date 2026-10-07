#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
INDISA V2.5.3 AUTÓNOMA
-----------------
Versión sin verificación exacta de médico.

Diferencias respecto a V9.1:
- NO usa Excel.
- NO lee/compara RUT de médico.
- NO selecciona timbre por identidad del médico.
- NO usa Tesseract.
- recibe UN timbre explícito por --stamp-file.
- busca una zona libre ENTRE patient_signature y doctor_signature.
- si detecta un data_stamp inferior, no inserta otro (sin verificar identidad).
- doctor_name y doctor_rut se revisan SOLO por presencia; no se valida su identidad.
- originales nunca se modifican.

Esta variante sirve si la identidad/timbre correcto se resuelve externamente
y este componente solo debe encontrar un espacio visual seguro y aplicar el PNG.
"""

from pathlib import Path
from copy import deepcopy
from datetime import datetime, timezone
import argparse, csv, json, math, shutil, hashlib, unicodedata
import numpy as np
import cv2
import pymupdf as fitz
from ultralytics import YOLO

import indisa_core_v9_1 as core

VERSION = "V2.5.3"

MIN_TEMPLATE_INLIERS = 40
SIGNATURE_MARGIN = 0.012
SAFE_REGION = (0.24, 0.66, 0.94, 0.95)
STAMP_WIDTHS = (0.18, 0.165, 0.15, 0.135)

def now_iso():
    return datetime.now(timezone.utc).astimezone().isoformat(timespec="seconds")

def best_box(component):
    if component.get("status") not in {"PRESENT", "PRESENT_GEOMETRIC_CONFIRMED"}:
        return None
    cands = component.get("candidates", []) or []
    if not cands:
        return None
    c = max(cands, key=lambda x: float(x.get("confidence", 0) or 0))
    b = c.get("bbox", {}).get("xyxy_norm")
    if not b or len(b) != 4:
        return None
    return tuple(float(x) for x in b)

def inflate(box, m):
    x1,y1,x2,y2 = box
    return (max(0,x1-m), max(0,y1-m), min(1,x2+m), min(1,y2+m))

def intersects(a,b):
    return not (a[2] <= b[0] or a[0] >= b[2] or a[3] <= b[1] or a[1] >= b[3])

def box_gap(a,b):
    dx = max(a[0]-b[2], b[0]-a[2], 0.0)
    dy = max(a[1]-b[3], b[1]-a[3], 0.0)
    return math.hypot(dx,dy)

def lower_stamp_present(result):
    cands = result.get("visual", {}).get("data_stamp", {}).get("candidates", []) or []
    for c in cands:
        b = c.get("bbox", {}).get("xyxy_norm")
        if b and len(b) == 4:
            cy = (float(b[1]) + float(b[3])) / 2.0
            if cy > 0.30:
                return True
    return False

def load_placement_zones(path=None):
    """
    Zonas NORMALIZADAS [x1,y1,x2,y2] donde el timbre puede colocarse.
    Son editables sin tocar el código.
    """
    defaults = {
        "endoscopia_alta": [0.443522, 0.838824, 0.571429, 0.923529],
        "colonoscopia": [0.438538, 0.861176, 0.579734, 0.975294],
    }

    if path is None:
        return defaults

    p = Path(path)
    if not p.exists():
        raise FileNotFoundError(f"No existe placement-zones: {p}")

    data = json.loads(p.read_text(encoding="utf-8"))
    for key in ("endoscopia_alta", "colonoscopia"):
        if key not in data or not isinstance(data[key], list) or len(data[key]) != 4:
            raise ValueError(
                f"placement_zones.json debe incluir {key}: [x1,y1,x2,y2]"
            )
        vals = [float(x) for x in data[key]]
        if not (0 <= vals[0] < vals[2] <= 1 and 0 <= vals[1] < vals[3] <= 1):
            raise ValueError(f"Zona inválida para {key}: {vals}")
        data[key] = vals
    return data


def propose_calibrated_slot_target(
    page_shape,
    patient_box,
    doctor_box,
    template_type,
    placement_zones,
):
    """
    V2.4: no requiere un PNG de timbre.

    El slot calibrado se considera el target lógico completo.
    Si ese slot invade una firma detectada (incluyendo margen de seguridad),
    queda en MANUAL_REVIEW.
    """
    zone = tuple(float(x) for x in placement_zones[template_type])

    p = inflate(patient_box, SIGNATURE_MARGIN)
    d = inflate(doctor_box, SIGNATURE_MARGIN)

    if intersects(zone, p) or intersects(zone, d):
        return None, "no_safe_target_inside_calibrated_slot"

    clearance = min(box_gap(zone, p), box_gap(zone, d))
    return zone, {
        "search_mode": "fixed_calibrated_slot_no_stamp_file",
        "allowed_zone_norm": [round(v, 6) for v in zone],
        "target_box_norm": [round(v, 6) for v in zone],
        "clearance_norm": round(clearance, 6),
    }


def draw_placement_guide(
    rectified,
    patient_box,
    doctor_box,
    target,
    allowed_zone,
    out_path,
):
    """Preview de depuración: zona permitida + firmas + target."""
    img = rectified.copy()
    h, w = img.shape[:2]

    def px(box):
        x1,y1,x2,y2 = box
        return (
            int(round(x1*w)), int(round(y1*h)),
            int(round(x2*w)), int(round(y2*h)),
        )

    # Zona permitida.
    x1,y1,x2,y2 = px(allowed_zone)
    cv2.rectangle(img, (x1,y1), (x2,y2), (255,0,0), 3)

    # Firma paciente / médico.
    for box in (patient_box, doctor_box):
        x1,y1,x2,y2 = px(box)
        cv2.rectangle(img, (x1,y1), (x2,y2), (0,0,255), 3)

    # Target final.
    if target is not None:
        x1,y1,x2,y2 = px(target)
        cv2.rectangle(img, (x1,y1), (x2,y2), (0,180,0), 4)

    core.safe_imwrite(Path(out_path), img, 95)



UNSUPPORTED_FILENAME_KEYWORDS = (
    "ENDOSONOGRAF",
    "ECOENDOSCOP",
    "CPRE",
    "ERCP",
)


def _ascii_upper(text):
    s = unicodedata.normalize("NFKD", str(text))
    return "".join(ch for ch in s if not unicodedata.combining(ch)).upper()


def filename_declares_unsupported_template(filename):
    name = _ascii_upper(filename)
    return any(k in name for k in UNSUPPORTED_FILENAME_KEYWORDS)


def page_fingerprint(raw):
    h = hashlib.sha256()
    h.update(str(raw.shape).encode("ascii"))
    h.update(raw.tobytes())
    return h.hexdigest()


def choose_template(raw, templates):
    attempts = []
    for name, tpl in templates.items():
        try:
            H, matches, inliers = core.orb_register(tpl, raw)
            attempts.append((inliers, matches, name, H))
        except Exception:
            pass

    if not attempts:
        return None

    attempts.sort(reverse=True, key=lambda x: (x[0], x[1]))
    inliers, matches, name, H = attempts[0]

    if inliers < MIN_TEMPLATE_INLIERS:
        return None

    runner_up = None
    if len(attempts) > 1:
        ri, rm, rn, _ = attempts[1]
        runner_up = {
            "template_type": rn,
            "matches": int(rm),
            "inliers": int(ri),
        }

    return {
        "template_type": name,
        "H": H,
        "matches": int(matches),
        "inliers": int(inliers),
        "min_required_inliers": MIN_TEMPLATE_INLIERS,
        "runner_up": runner_up,
    }


FIELD_GEOM_CONFIRM_THRESHOLD = 0.010


def confirm_uncertain_fields_by_template_diff(result, rectified, template_aligned, template_type):
    """
    Solo intenta rescatar campos que YA están UNCERTAIN.
    Nunca cambia un EMPTY de alta confianza.

    La evidencia proviene de tinta adicional respecto de la plantilla alineada.
    """
    h, w = rectified.shape[:2]
    cfg = core.CONFIG["templates"][template_type]

    scores = {}

    for name, field in result.get("fields", {}).items():
        roi = cfg["field_rois"].get(name)
        if roi is None:
            continue

        x1, y1, x2, y2 = core.nbox_to_px(roi, w, h)
        feat = core.diff_features(
            rectified[y1:y2, x1:x2],
            template_aligned[y1:y2, x1:x2],
        )
        score = float(feat["clean_ratio"])
        scores[name] = round(score, 6)

        if (
            field.get("status") == "UNCERTAIN"
            and score >= FIELD_GEOM_CONFIRM_THRESHOLD
        ):
            field.update({
                "status": "FILLED",
                "filled": True,
                "source": "template_subtraction_confirmation",
                "manual_review_required": False,
                "review_reason": None,
                "geometric_presence_score": round(score, 6),
                "geometric_presence_threshold": FIELD_GEOM_CONFIRM_THRESHOLD,
            })

    result["field_presence_geometric_fallback"] = {
        "mode": "confirm_UNCERTAIN_only",
        "threshold": FIELD_GEOM_CONFIRM_THRESHOLD,
        "scores": scores,
    }


def custom_review(result):
    """
    V2 autónoma:
    - reconoce/revisa TODOS los campos estructurados por presencia;
    - no compara su contenido con Excel;
    - no exige identificar qué médico es.
    """
    items = []

    for name, field in result.get("fields", {}).items():
        if field.get("manual_review_required"):
            items.append({
                "component": name,
                "reason": field.get("review_reason"),
            })

    for name, comp in result.get("visual", {}).items():
        if name == "data_stamp":
            continue
        if comp.get("manual_review_required"):
            items.append({
                "component": name,
                "reason": comp.get("review_reason"),
            })

    stamp = result.get("stamp_stage", {})
    if stamp.get("manual_review_required"):
        reason = stamp.get("review_reason")
        if not any(x.get("reason") == reason for x in items):
            items.append({
                "component": "stamp_stage",
                "reason": reason,
            })

    uniq = []
    seen = set()
    for x in items:
        key = (x.get("component"), x.get("reason"))
        if key not in seen:
            seen.add(key)
            uniq.append(x)

    result["manual_review_items"] = uniq
    result["manual_review_required"] = bool(uniq)


def preflight(args):
    checks = [
        ("input",Path(args.input)),
        ("field-model",Path(args.field_model)),
        ("visual-main",Path(args.visual_main)),
        ("doctor-model",Path(args.doctor_model)),
        ("template-alta",Path(args.template_alta)),
        ("template-colon",Path(args.template_colon)),
    ]
    missing=[f"{n}: {p}" for n,p in checks if not p.exists()]
    if missing:
        raise SystemExit("Faltan recursos:\n- " + "\n- ".join(missing))


def _component_state(result, component):
    if component in result.get("fields", {}):
        x = result["fields"][component]
        return x.get("status"), x.get("confidence")
    if component in result.get("visual", {}):
        x = result["visual"][component]
        return x.get("status"), x.get("confidence")
    if component == "stamp_stage":
        x = result.get("stamp_stage", {})
        return x.get("status"), None
    return None, None


def _alert_message(component, reason, status):
    labels = {
        "patient_name": "Nombre del paciente",
        "patient_rut": "RUT del paciente",
        "procedure_date": "Fecha",
        "doctor_name": "Nombre del médico",
        "doctor_rut": "RUT del médico",
        "signer_selected_checkbox": "Tipo de firmante",
        "patient_signature": "Firma paciente/apoderado",
        "doctor_signature": "Firma médico",
        "stamp_stage": "Aplicación del timbre",
        "pdf": "PDF",
        "document_layout": "Estructura del documento",
        "rectification": "Rectificación",
        "inference": "Inferencia",
    }
    label = labels.get(component, component)

    if reason == "low_confidence":
        return f"{label}: inferencia de baja confianza; requiere revisión."
    if reason == "missing_required_field":
        return f"{label}: no se detectó contenido en el campo requerido."
    if reason == "not_detected":
        return f"{label}: no fue detectado."
    if reason == "ambiguous_multiple_candidates":
        return f"{label}: detección ambigua; hay múltiples candidatos."
    if reason == "doctor_signature_not_safely_located":
        return "No se pudo localizar la firma médica con seguridad para posicionar el timbre."
    if reason == "prior_stamp_inconclusive":
        return "Se detectó un posible timbre previo, pero no pudo confirmarse con seguridad."
    if reason == "pdf_alignment_low_confidence":
        return "La alineación con el PDF original no fue suficientemente confiable."
    if reason == "mapped_target_outside_page":
        return "La zona calculada para el timbre quedó fuera del área válida de la página."
    if reason == "doctor_stamp_png_not_found":
        return "No se encontró la imagen de timbre correspondiente al médico."
    if reason == "doctor_stamp_png_unreadable":
        return "La imagen de timbre no pudo leerse."
    if reason == "no_safe_target_inside_calibrated_slot":
        return "El slot calibrado está invadido por una firma y requiere revisión."
    if reason == "unsupported_template":
        return "El procedimiento corresponde a una plantilla no soportada por esta versión."
    if reason == "no_supported_consent_page_recognized":
        return "No se reconoció una página de consentimiento compatible con las plantillas soportadas."
    if reason and reason.startswith("missing_signature"):
        return f"{label}: faltan firmas necesarias para validar el slot de timbre."
    return f"{label}: {reason or status or 'requiere revisión'}."


def build_alerts_from_results(results):
    alerts = []
    for r in results:
        source_pdf = r.get("source_pdf") or Path(str(r.get("source_pdf_path", ""))).name
        page = r.get("page")
        for item in r.get("manual_review_items", []) or []:
            component = item.get("component")
            reason = item.get("reason")
            status, confidence = _component_state(r, component)
            alerts.append({
                "source_pdf": source_pdf,
                "page": page,
                "component": component,
                "status": status,
                "reason": reason,
                "confidence": confidence,
                "severity": "WARNING",
                "message": _alert_message(component, reason, status),
                "final_pdf": r.get("final_pdf") or r.get("output_pdf"),
            })
    return alerts


ANNOTATION_COLORS = {
    "FILLED": (0.0, 0.65, 0.0),        # verde
    "EMPTY": (1.0, 0.0, 0.0),          # rojo
    "UNCERTAIN": (1.0, 0.55, 0.0),     # naranjo
    "CHECKBOX": (0.0, 0.65, 0.85),     # celeste
    "SIGNATURE": (0.85, 0.0, 0.85),    # magenta
    "STAMP": (0.45, 0.1, 0.75),        # morado
    "SLOT": (0.0, 0.2, 1.0),           # azul
    "TARGET": (0.0, 0.8, 0.0),         # verde fuerte
}


def _rect_px_to_norm(box_px, rectified_shape):
    h, w = rectified_shape[:2]
    x1, y1, x2, y2 = [float(v) for v in box_px]
    return (x1 / w, y1 / h, x2 / w, y2 / h)


def _norm_box_to_pdf_quad(box_norm, rectified_shape, raw_shape, H, page):
    q = core.target_quad(box_norm, rectified_shape)
    mapped = cv2.perspectiveTransform(q, H).reshape(-1, 2)

    rh, rw = raw_shape[:2]
    pr = page.rect
    pts = []
    for x, y in mapped:
        px = pr.x0 + (float(x) / rw) * pr.width
        py = pr.y0 + (float(y) / rh) * pr.height
        pts.append(fitz.Point(px, py))
    return pts


def _draw_quad_with_label(page, pts, label, color, width=1.2):
    if len(pts) != 4:
        return

    poly = [pts[0], pts[1], pts[2], pts[3], pts[0]]
    page.draw_polyline(
        poly,
        color=color,
        width=width,
        overlay=True,
    )

    # Etiqueta compacta, colocada sobre la esquina superior izquierda.
    p = pts[0]
    x = max(page.rect.x0 + 2, p.x)
    y = max(page.rect.y0 + 8, p.y - 2)
    try:
        page.insert_text(
            fitz.Point(x, y),
            label[:55],
            fontsize=5.5,
            color=color,
            overlay=True,
        )
    except Exception:
        pass


def _best_candidate_box(component):
    cands = component.get("candidates", []) or []
    valid = []
    for c in cands:
        b = c.get("bbox", {}).get("xyxy_norm")
        if b and len(b) == 4:
            valid.append(c)
    if not valid:
        return None
    c = max(valid, key=lambda z: float(z.get("confidence", 0) or 0))
    return tuple(float(v) for v in c["bbox"]["xyxy_norm"]), float(c.get("confidence", 0) or 0)


def annotate_final_pdf(
    clean_pdf_path,
    annotated_pdf_path,
    source_pdf_path,
    page_results,
    placement_zones,
    dpi,
):
    """
    Crea un PDF final de revisión con cajas VISIBLES.

    El PDF limpio continúa disponible en pdfs/ o manual_review/.
    final_pdfs/ contiene esta versión anotada.
    """
    src_pdf = Path(source_pdf_path)
    clean_pdf = Path(clean_pdf_path)
    annotated_pdf = Path(annotated_pdf_path)

    doc = fitz.open(str(clean_pdf))
    annotation_errors = []

    for r in page_results:
        try:
            page_num = int(r.get("page", 1))
            page_idx = page_num - 1
            if page_idx < 0 or page_idx >= len(doc):
                continue

            rect_path = Path(r.get("rectified_image", ""))
            rectified = core.safe_imread(rect_path, cv2.IMREAD_COLOR)
            if rectified is None:
                raise RuntimeError("rectified_image_unreadable")

            raw = core.render_pdf_page_bgr(src_pdf, page_num, dpi)
            H, hm = core.estimate_pdf_homography(rectified, raw)
            if (
                H is None
                or hm.get("inliers", 0) < core.MIN_H_INLIERS
                or hm.get("inlier_ratio", 0) < core.MIN_H_RATIO
                or hm.get("median_reprojection_error_px", 999) > core.MAX_H_MEDIAN_ERROR_PX
            ):
                raise RuntimeError("annotation_homography_low_confidence")

            page = doc[page_idx]

            # 1) Campos estructurados.
            for name, field in r.get("fields", {}).items():
                roi_px = field.get("roi_px")
                if not roi_px or len(roi_px) != 4:
                    continue
                status = field.get("status", "UNCERTAIN")
                conf = field.get("confidence")
                box = _rect_px_to_norm(roi_px, rectified.shape)
                pts = _norm_box_to_pdf_quad(box, rectified.shape, raw.shape, H, page)
                conf_txt = "" if conf is None else f" {float(conf):.2f}"
                _draw_quad_with_label(
                    page,
                    pts,
                    f"{name}: {status}{conf_txt}",
                    ANNOTATION_COLORS.get(status, ANNOTATION_COLORS["UNCERTAIN"]),
                    width=1.15,
                )

            # 2) Checkbox.
            checkbox = r.get("visual", {}).get("signer_selected_checkbox", {})
            bb = _best_candidate_box(checkbox)
            if bb:
                box, conf = bb
                pts = _norm_box_to_pdf_quad(box, rectified.shape, raw.shape, H, page)
                option = checkbox.get("selected_option") or "checkbox"
                _draw_quad_with_label(
                    page, pts, f"checkbox: {option} {conf:.2f}",
                    ANNOTATION_COLORS["CHECKBOX"], width=1.3
                )

            # 3) Firmas.
            for key, short in (
                ("patient_signature", "firma paciente"),
                ("doctor_signature", "firma medico"),
            ):
                comp = r.get("visual", {}).get(key, {})
                bb = _best_candidate_box(comp)
                if bb:
                    box, conf = bb
                    pts = _norm_box_to_pdf_quad(box, rectified.shape, raw.shape, H, page)
                    _draw_quad_with_label(
                        page, pts, f"{short}: {conf:.2f}",
                        ANNOTATION_COLORS["SIGNATURE"], width=1.5
                    )

            # 4) Timbre existente detectado, si lo hay.
            stamp_comp = r.get("visual", {}).get("data_stamp", {})
            bb = _best_candidate_box(stamp_comp)
            if bb:
                box, conf = bb
                pts = _norm_box_to_pdf_quad(box, rectified.shape, raw.shape, H, page)
                _draw_quad_with_label(
                    page, pts, f"data_stamp: {conf:.2f}",
                    ANNOTATION_COLORS["STAMP"], width=1.4
                )

            # 5) Slot calibrado SIEMPRE visible.
            tt = r.get("template_type")
            if tt in placement_zones:
                slot = tuple(float(x) for x in placement_zones[tt])
                pts = _norm_box_to_pdf_quad(slot, rectified.shape, raw.shape, H, page)
                _draw_quad_with_label(
                    page, pts, "SLOT TIMBRE PERMITIDO",
                    ANNOTATION_COLORS["SLOT"], width=2.0
                )

            # 6) Target usado, si existió.
            target = r.get("stamp_stage", {}).get("target_box_norm")
            if target and len(target) == 4:
                target = tuple(float(x) for x in target)
                pts = _norm_box_to_pdf_quad(target, rectified.shape, raw.shape, H, page)
                _draw_quad_with_label(
                    page, pts, "TARGET TIMBRE",
                    ANNOTATION_COLORS["TARGET"], width=2.2
                )

            # 7) Resumen de alertas de esta página.
            alerts = r.get("manual_review_items", []) or []
            if alerts:
                y = page.rect.y0 + 10
                x = page.rect.x0 + 8
                page.insert_text(
                    fitz.Point(x, y),
                    "REVISAR:",
                    fontsize=6.5,
                    color=(1.0, 0.0, 0.0),
                    overlay=True,
                )
                for item in alerts[:5]:
                    y += 7
                    msg = f"- {item.get('component')}: {item.get('reason')}"
                    page.insert_text(
                        fitz.Point(x, y),
                        msg[:90],
                        fontsize=5.3,
                        color=(0.8, 0.0, 0.0),
                        overlay=True,
                    )

            r["annotation_homography"] = hm

        except Exception as exc:
            annotation_errors.append({
                "page": r.get("page"),
                "error": f"{type(exc).__name__}: {exc}",
            })
            r["annotation_error"] = f"{type(exc).__name__}: {exc}"

    if annotated_pdf.exists():
        annotated_pdf.unlink()
    doc.save(str(annotated_pdf), garbage=4, deflate=True)
    doc.close()
    return annotation_errors



def _audit_field(r, name):
    return (r.get("fields", {}) or {}).get(name, {}) or {}


def _audit_visual(r, name):
    return (r.get("visual", {}) or {}).get(name, {}) or {}


def _audit_geom_score(comp):
    g = comp.get("geometric_confirmation", {}) or {}
    return g.get("score", g.get("presence_score"))


def build_detailed_audit_row(r):
    row = {
        "record_type": r.get("record_type"),
        "source_pdf": r.get("source_pdf"),
        "page": r.get("page"),
        "template_type": r.get("template_type"),
        "template_orb_inliers": (r.get("template_alignment", {}) or {}).get("orb_inliers"),
        "template_min_inliers": (r.get("template_alignment", {}) or {}).get("min_required_inliers"),
        "is_exact_duplicate": r.get("is_exact_duplicate"),
    }

    field_geom_scores = (
        (r.get("field_presence_geometric_fallback", {}) or {}).get("scores", {}) or {}
    )

    for name in ("patient_name", "patient_rut", "procedure_date", "doctor_name", "doctor_rut"):
        f = _audit_field(r, name)
        row[f"{name}_status"] = f.get("status")
        row[f"{name}_confidence"] = f.get("confidence")
        row[f"{name}_source"] = f.get("source")
        row[f"{name}_geom_score"] = f.get(
            "geometric_presence_score", field_geom_scores.get(name)
        )
        row[f"{name}_manual_review"] = f.get("manual_review_required")
        row[f"{name}_reason"] = f.get("review_reason")

    cb = _audit_visual(r, "signer_selected_checkbox")
    cbg = cb.get("geometric_confirmation", {}) or {}
    ps = _audit_visual(r, "patient_signature")
    ds = _audit_visual(r, "doctor_signature")
    stamp = _audit_visual(r, "data_stamp")

    row.update({
        "checkbox_status": cb.get("status"),
        "checkbox_option": cb.get("selected_option"),
        "checkbox_confidence": cb.get("confidence"),
        "checkbox_geom_score": cbg.get("presence_score"),
        "checkbox_geom_margin": cbg.get("option_margin"),
        "checkbox_manual_review": cb.get("manual_review_required"),

        "patient_signature_status": ps.get("status"),
        "patient_signature_confidence": ps.get("confidence"),
        "patient_signature_geom_score": _audit_geom_score(ps),
        "patient_signature_manual_review": ps.get("manual_review_required"),

        "doctor_signature_status": ds.get("status"),
        "doctor_signature_confidence": ds.get("confidence"),
        "doctor_signature_geom_score": _audit_geom_score(ds),
        "doctor_signature_manual_review": ds.get("manual_review_required"),

        "data_stamp_status": stamp.get("status"),
        "data_stamp_count": stamp.get("count"),
        "data_stamp_confidence": stamp.get("confidence"),

        "slot_status": (r.get("stamp_stage", {}) or {}).get("status"),
        "slot_reason": (r.get("stamp_stage", {}) or {}).get("reason"),

        "manual_review_required": r.get("manual_review_required"),
        "manual_review_reasons": " | ".join(
            f"{x.get('component')}:{x.get('reason')}"
            for x in (r.get("manual_review_items", []) or [])
        ),
        "output_pdf": r.get("output_pdf"),
        "final_pdf": r.get("final_pdf"),
    })
    return row


OUTPUT_SCHEMA_VERSION = "1.0"


def _simple_confirmation_from_visual(component):
    g = component.get("geometric_confirmation", {}) or {}
    if not g.get("confirmed"):
        return None

    out = {"method": "geometric_confirmation"}
    score = g.get("score", g.get("presence_score"))
    if score is not None:
        out["score"] = score
    if g.get("threshold") is not None:
        out["threshold"] = g.get("threshold")
    if g.get("option_margin") is not None:
        out["margin"] = g.get("option_margin")
    return out


def _simple_field(result, name):
    field = (result.get("fields", {}) or {}).get(name, {}) or {}
    out = {
        "status": field.get("status"),
        "confidence": field.get("confidence"),
        "source": field.get("source"),
    }

    source = field.get("source")
    if source == "template_subtraction_confirmation":
        geom_scores = (
            (result.get("field_presence_geometric_fallback", {}) or {})
            .get("scores", {}) or {}
        )
        score = field.get("geometric_presence_score", geom_scores.get(name))
        out["confirmation"] = {
            "method": "template_subtraction",
            "score": score,
        }

    # Quitar claves nulas para mantener el JSON compacto.
    return {k: v for k, v in out.items() if v is not None}


def _simple_visual(component, option_key=None):
    component = component or {}
    out = {
        "status": component.get("status"),
        "confidence": component.get("confidence"),
    }
    if option_key:
        out["option"] = component.get(option_key)

    confirmation = _simple_confirmation_from_visual(component)
    if confirmation:
        out["confirmation"] = confirmation

    return {k: v for k, v in out.items() if v is not None}


def simplify_result_document(result):
    fields = result.get("fields", {}) or {}
    visual = result.get("visual", {}) or {}
    stage = result.get("stamp_stage", {}) or {}
    alignment = result.get("template_alignment", {}) or {}

    source_pdf = result.get("source_pdf")
    if not source_pdf and result.get("source_pdf_path"):
        source_pdf = Path(result["source_pdf_path"]).name

    simple = {
        "source_pdf": source_pdf,
        "page": result.get("page"),
        "attention_id": result.get("attention_id"),
        "procedure": result.get("procedure"),
        "template": {
            "type": result.get("template_type"),
            "recognized": bool(result.get("template_type")),
            "alignment_inliers": alignment.get("orb_inliers"),
        },
        "fields": {
            name: _simple_field(result, name)
            for name in (
                "patient_name",
                "patient_rut",
                "procedure_date",
                "doctor_name",
                "doctor_rut",
            )
            if name in fields
        },
        "checkbox": _simple_visual(
            visual.get("signer_selected_checkbox", {}),
            option_key="selected_option",
        ),
        "signatures": {
            "patient": _simple_visual(visual.get("patient_signature", {})),
            "doctor": _simple_visual(visual.get("doctor_signature", {})),
        },
        "stamp": {
            "status": stage.get("status"),
            "action": stage.get("action"),
            "reason": stage.get("reason"),
            "existing_stamp_detected": stage.get("status") in {
                "EXISTING_STAMP_DETECTED",
                "EXISTING_CORRECT_STAMP",
            },
        },
        "review": {
            "required": bool(result.get("manual_review_required")),
            "reasons": [
                {
                    "component": x.get("component"),
                    "reason": x.get("reason"),
                }
                for x in (result.get("manual_review_items", []) or [])
            ],
        },
        "final_pdf": result.get("final_pdf"),
    }

    # Limpiar opcionales de nivel superior.
    return {
        k: v
        for k, v in simple.items()
        if v is not None
    }


def build_simple_results(results, pipeline_version):
    return {
        "schema_version": OUTPUT_SCHEMA_VERSION,
        "pipeline_version": pipeline_version,
        "documents": [simplify_result_document(r) for r in results],
    }

def main():
    ap=argparse.ArgumentParser(
        description="INDISA V2.5.3 AUTÓNOMA - timbrado por espacio entre firmas, sin verificar médico"
    )
    ap.add_argument("--input",required=True)
    ap.add_argument("--output",default="resultados_indisa_space")
    ap.add_argument(
        "--placement-zones",
        default="placement_zones.json",
        help="JSON con zonas permitidas normalizadas por plantilla",
    )
    ap.add_argument("--field-model",default=r"models\field_presence_v811_best.pt")
    ap.add_argument("--visual-main",default=r"models\visual_main_v813.pt")
    ap.add_argument("--doctor-model",default=r"models\doctor_signature_fallback_v811.pt")
    ap.add_argument("--template-alta",
        default=r"datos\Consentimientos\Plantillas base\Consentimiento_base_EndoscopíaDigestivaAlta_Esófago-gastr.pdf")
    ap.add_argument("--template-colon",
        default=r"datos\Consentimientos\Plantillas base\Consentimiento_base_Colonoscopía_EndoscopíaDigestivaBaja.pdf")
    ap.add_argument("--dpi",type=int,default=150)
    ap.add_argument("--device",default="0")
    ap.add_argument(
        "--no-pdf-export",
        action="store_true",
        help=(
            "No genera copias PDF en pdfs/, manual_review/ ni final_pdfs/. "
            "Se mantienen JSON, CSV, summary y previews independientes del PDF final."
        ),
    )
    ap.add_argument("--overwrite",action="store_true")
    args=ap.parse_args()

    preflight(args)
    placement_zones = load_placement_zones(args.placement_zones)
    out=Path(args.output)
    if out.exists() and any(out.iterdir()):
        if not args.overwrite:
            raise SystemExit(f"Salida no vacía: {out}. Usa otra o --overwrite.")
        shutil.rmtree(out)

    clean_dir=out/"pdfs"
    review_dir=out/"manual_review"
    final_dir=out/"final_pdfs"
    previews=out/"previews"
    rects=out/"work"/"rectified"

    # JSON/CSV y trabajo interno siempre se generan.
    for d in (previews,rects):
        d.mkdir(parents=True,exist_ok=True)

    # La exportación PDF es opcional para acelerar ejecuciones de integración.
    if not args.no_pdf_export:
        for d in (clean_dir,review_dir,final_dir):
            d.mkdir(parents=True,exist_ok=True)

    tpl_alta=core.render_pdf_page_bgr(Path(args.template_alta),1,args.dpi)
    tpl_colon=core.render_pdf_page_bgr(Path(args.template_colon),1,args.dpi)
    templates={"endoscopia_alta":tpl_alta,"colonoscopia":tpl_colon}

    print("Cargando modelos...")
    field_model=YOLO(args.field_model)
    visual_main=YOLO(args.visual_main)
    doctor_model=YOLO(args.doctor_model)
    print("Modelos cargados.")

    pdfs=core.collect_input_pdfs(args.input)
    print("PDFs de entrada:",len(pdfs))
    print("Excel: NO")
    print("Tesseract: NO")
    print("Verificación exacta de médico: NO")
    print("Inserción de PNG de timbre: NO; solo se valida y muestra el slot")
    print("Exportación PDF:", "NO" if args.no_pdf_export else "SÍ")
    print()

    results=[]
    pdf_summaries=[]
    seen_page_hashes={}

    for pi,pdf in enumerate(pdfs,1):
        print(f"[{pi}/{len(pdfs)}] {pdf.name}")

        if filename_declares_unsupported_template(pdf.name):
            review_copy=None
            final_copy=None
            if not args.no_pdf_export:
                review_copy=review_dir/pdf.name
                final_copy=final_dir/pdf.name
                shutil.copy2(pdf,review_copy)
                shutil.copy2(pdf,final_copy)

            rec={
                "record_type":"pdf_status",
                "schema_version":VERSION,
                "processed_at":now_iso(),
                "source_pdf":pdf.name,
                "source_pdf_path":str(pdf),
                "page":None,
                "template_type":None,
                "manual_review_required":True,
                "manual_review_items":[{
                    "component":"document_layout",
                    "reason":"unsupported_template"
                }],
                "stamp_stage":{
                    "status":"MANUAL_REVIEW",
                    "action":"NONE",
                    "reason":"unsupported_template",
                    "manual_review_required":True,
                    "review_reason":"unsupported_template",
                },
                "output_pdf":str(review_copy) if review_copy else None,
                "final_pdf":str(final_copy) if final_copy else None,
                "final_pdf_mode":"disabled" if args.no_pdf_export else "unchanged_unsupported_pdf",
            }
            results.append(rec)
            pdf_summaries.append({
                "source_pdf":pdf.name,
                "status":"UNSUPPORTED_TEMPLATE",
                "manual_review_required":True,
                "consent_pages":0,
                "rejected_pages":[],
                "stamps_applied":0,
                "output_pdf":str(review_copy) if review_copy else None,
            })
            print("  -> REVIEW unsupported template declared by filename")
            continue

        srcdoc=fitz.open(str(pdf))
        page_count=len(srcdoc)
        srcdoc.close()

        page_results=[]
        jobs=[]
        consent_pages=0
        rejected_pages=[]

        for page_num in range(1,page_count+1):
            try:
                raw=core.render_pdf_page_bgr(pdf,page_num,args.dpi)
                page_sha=page_fingerprint(raw)
                choice=choose_template(raw,templates)
            except Exception as exc:
                page_sha=None
                choice=None
                rejected_pages.append({
                    "page":page_num,
                    "reason":f"render_or_template_error:{type(exc).__name__}",
                })

            if choice is None:
                if not any(x.get("page")==page_num for x in rejected_pages):
                    rejected_pages.append({
                        "page":page_num,
                        "reason":"template_alignment_below_min_inliers",
                        "min_required_inliers":MIN_TEMPLATE_INLIERS,
                    })
                continue

            consent_pages += 1
            tt=choice["template_type"]
            tpl=templates[tt]
            rect=core.rectify_scan(raw,choice["H"],tpl.shape)
            rect_path=rects/f"{pdf.stem}_p{page_num}_{tt}.jpg"
            core.safe_imwrite(rect_path,rect,95)

            r=core.analyze_rectified(
                rect,tpl,tt,field_model,visual_main,doctor_model,args.device
            )

            # V2.1: evidencia geométrica SOLO para campos que quedaron UNCERTAIN.
            tmpl_aligned_fields = core.align_template_ecc(tpl, rect)
            confirm_uncertain_fields_by_template_diff(
                r, rect, tmpl_aligned_fields, tt
            )

            r.update({
                "schema_version":VERSION,
                "processed_at":now_iso(),
                "source_pdf":pdf.name,
                "source_pdf_path":str(pdf),
                "page":page_num,
                "template_type":tt,
                "rectified_image":str(rect_path),
                "doctor_identity_verification":"DISABLED",
                "structured_fields_mode":"presence_only_all_fields",
                "placement_policy":"fixed_calibrated_slot_avoiding_signatures",
                "record_type":"consent_page",
                "template_alignment":{
                    "orb_matches":choice["matches"],
                    "orb_inliers":choice["inliers"],
                    "min_required_inliers":choice.get("min_required_inliers"),
                    "runner_up":choice.get("runner_up"),
                },
                "page_sha256":page_sha,
            })

            prior=seen_page_hashes.get(page_sha)
            if prior is None:
                seen_page_hashes[page_sha]={
                    "source_pdf":pdf.name,
                    "page":page_num,
                }
                r["is_exact_duplicate"]=False
                r["exact_duplicate_of"]=None
            else:
                r["is_exact_duplicate"]=True
                r["exact_duplicate_of"]=prior

            if lower_stamp_present(r):
                r["stamp_stage"]={
                    "status":"EXISTING_STAMP_DETECTED",
                    "action":"NO_INSERT",
                    "reason":"lower_data_stamp_present_identity_not_verified",
                    "manual_review_required":False,
                    "review_reason":None,
                }
                custom_review(r)
                page_results.append(r);results.append(r)
                print(f"  p{page_num}: existing lower stamp -> no duplicate")
                continue

            pb=best_box(r["visual"]["patient_signature"])
            db=best_box(r["visual"]["doctor_signature"])

            if pb is None or db is None:
                missing=[]
                if pb is None: missing.append("patient_signature")
                if db is None: missing.append("doctor_signature")
                reason="missing_signature_for_between_space:" + ",".join(missing)
                r["stamp_stage"]={
                    "status":"MANUAL_REVIEW",
                    "action":"NONE",
                    "reason":reason,
                    "manual_review_required":True,
                    "review_reason":reason,
                }
                custom_review(r)
                page_results.append(r);results.append(r)
                print(f"  p{page_num}: REVIEW {reason}")
                continue

            target,meta=propose_calibrated_slot_target(
                rect.shape,
                pb,
                db,
                tt,
                placement_zones,
            )
            if target is None:
                r["stamp_stage"]={
                    "status":"MANUAL_REVIEW",
                    "action":"NONE",
                    "reason":meta,
                    "manual_review_required":True,
                    "review_reason":meta,
                }
                custom_review(r)
                page_results.append(r);results.append(r)
                print(f"  p{page_num}: REVIEW {meta}")
                continue

            guide_path = previews / f"{pdf.stem}_p{page_num}_placement_guide.jpg"
            draw_placement_guide(
                rect,
                pb,
                db,
                target,
                placement_zones[tt],
                guide_path,
            )

            Hpdf,hm=core.estimate_pdf_homography(rect,raw)
            if (
                Hpdf is None
                or hm.get("inliers",0) < core.MIN_H_INLIERS
                or hm.get("inlier_ratio",0) < core.MIN_H_RATIO
                or hm.get("median_reprojection_error_px",999) > core.MAX_H_MEDIAN_ERROR_PX
            ):
                r["stamp_stage"]={
                    "status":"MANUAL_REVIEW",
                    "action":"NONE",
                    "reason":"pdf_alignment_low_confidence",
                    "manual_review_required":True,
                    "review_reason":"pdf_alignment_low_confidence",
                    "homography":hm,
                }
                custom_review(r)
                page_results.append(r);results.append(r)
                print(f"  p{page_num}: REVIEW pdf_alignment_low_confidence")
                continue

            quad=cv2.perspectiveTransform(
                core.target_quad(target,rect.shape),Hpdf
            ).reshape(-1,2)
            rh,rw=raw.shape[:2]
            if (
                np.any(quad[:,0] < -5) or np.any(quad[:,0] > rw+5)
                or np.any(quad[:,1] < -5) or np.any(quad[:,1] > rh+5)
            ):
                r["stamp_stage"]={
                    "status":"MANUAL_REVIEW",
                    "action":"NONE",
                    "reason":"mapped_target_outside_page",
                    "manual_review_required":True,
                    "review_reason":"mapped_target_outside_page",
                }
                custom_review(r)
                page_results.append(r);results.append(r)
                continue

            r["stamp_stage"]={
                "status":"SLOT_READY",
                "action":"NO_INSERT",
                "reason":"safe_calibrated_slot_detected",
                "manual_review_required":False,
                "review_reason":None,
                "target_box_norm":[round(float(x),6) for x in target],
                "between_space":meta,
                "homography":hm,
            }
            custom_review(r)
            page_results.append(r);results.append(r)
            print(f"  p{page_num}: SLOT_READY | review={r['manual_review_required']}")

        if consent_pages == 0:
            review_copy=None
            final_copy=None
            if not args.no_pdf_export:
                review_copy=review_dir/pdf.name
                final_copy=final_dir/pdf.name
                shutil.copy2(pdf,review_copy)
                shutil.copy2(review_copy,final_copy)

            rec={
                "record_type":"pdf_status",
                "schema_version":VERSION,
                "processed_at":now_iso(),
                "source_pdf":pdf.name,
                "source_pdf_path":str(pdf),
                "page":None,
                "template_type":None,
                "manual_review_required":True,
                "manual_review_items":[{
                    "component":"document_layout",
                    "reason":"no_supported_consent_page_recognized"
                }],
                "stamp_stage":{
                    "status":"MANUAL_REVIEW",
                    "action":"NONE",
                    "reason":"no_supported_consent_page_recognized",
                    "manual_review_required":True,
                    "review_reason":"no_supported_consent_page_recognized",
                },
                "template_pages_rejected":rejected_pages,
                "output_pdf":str(review_copy) if review_copy else None,
                "final_pdf":str(final_copy) if final_copy else None,
                "final_pdf_mode":"disabled" if args.no_pdf_export else "unchanged_unrecognized_pdf",
            }
            results.append(rec)
            pdf_summaries.append({
                "source_pdf":pdf.name,
                "status":"NO_CONSENT_PAGE_RECOGNIZED",
                "manual_review_required":True,
                "consent_pages":0,
                "rejected_pages":rejected_pages,
                "stamps_applied":0,
                "output_pdf":str(review_copy) if review_copy else None,
            })
            print("  -> REVIEW no supported consent page recognized")
            continue

        pdf_has_review=any(r.get("manual_review_required") for r in page_results)
        applied=0

        if args.no_pdf_export:
            # Modo rápido/API: no escribir, copiar ni anotar PDFs.
            for rr in page_results:
                rr["output_pdf"] = None
                rr["final_pdf"] = None
                rr["final_pdf_mode"] = "disabled"

            pdf_summaries.append({
                "source_pdf":pdf.name,
                "status":"MANUAL_REVIEW" if pdf_has_review else "CLEAN",
                "manual_review_required":pdf_has_review,
                "consent_pages":consent_pages,
                "rejected_pages":rejected_pages,
                "stamps_applied":0,
                "output_pdf":None,
            })
        else:
            dst=(review_dir if pdf_has_review else clean_dir)/pdf.name

            doc=fitz.open(str(pdf))
            try:
                for job in jobs:
                    page=doc[job["page_idx"]]
                    page.insert_image(
                        page.rect,stream=job["overlay"],overlay=True,keep_proportion=False
                    )
                    job["result"]["stamp_stage"]["status"]="APPLIED"
                    applied += 1
                doc.save(str(dst),garbage=4,deflate=True)
            finally:
                doc.close()

            final_dst = final_dir / pdf.name
            annotation_errors = annotate_final_pdf(
                dst,
                final_dst,
                pdf,
                page_results,
                placement_zones,
                args.dpi,
            )
            for rr in page_results:
                rr["output_pdf"] = str(dst)
                rr["final_pdf"] = str(final_dst)
                rr["final_pdf_mode"] = "annotated_review_pdf"

            # Preview del PDF FINAL anotado solo cuando existe exportación PDF.
            for rr in page_results:
                try:
                    ann_preview = previews / f"{pdf.stem}_p{rr.get('page')}_final_annotated.jpg"
                    core.save_preview_from_pdf(
                        final_dst,
                        int(rr.get("page", 1)) - 1,
                        ann_preview,
                    )
                    rr["annotated_preview"] = str(ann_preview)
                except Exception as exc:
                    rr["annotated_preview_error"] = f"{type(exc).__name__}: {exc}"

            for job in jobs:
                pp=previews/f"{pdf.stem}_p{job['page_idx']+1}_after.jpg"
                try:
                    core.save_preview_from_pdf(dst,job["page_idx"],pp)
                except Exception:
                    pass
                job["result"]["output_pdf"]=str(dst)

            pdf_summaries.append({
                "source_pdf":pdf.name,
                "status":"MANUAL_REVIEW" if pdf_has_review else "CLEAN",
                "manual_review_required":pdf_has_review,
                "consent_pages":consent_pages,
                "rejected_pages":rejected_pages,
                "stamps_applied":applied,
                "output_pdf":str(dst),
            })

    debug_payload={
        "run":{
            "pipeline_version":VERSION,
            "doctor_identity_verification":False,
            "structured_fields_mode":"presence_only_all_fields",
            "excel_used":False,
            "tesseract_used":False,
            "stamp_source":"none",
            "placement":"fixed calibrated slot avoiding both signatures",
            "stamp_insertion_enabled":False,
            "final_pdf_mode":"disabled" if args.no_pdf_export else "annotated_review_pdf",
            "pdf_export_enabled":not args.no_pdf_export,
            "template_min_inliers":MIN_TEMPLATE_INLIERS,
            "patient_false_positive_guards":{
                "header_stamp_min_confidence":core.PATIENT_HEADER_STAMP_MIN_CONF,
                "patient_signature_roi_expansion":0.15,
                "checkbox_geometric_min_margin":core.CHECKBOX_GEOM_MIN_MARGIN,
            },
        },
        "documents":results,
        "pdfs":pdf_summaries,
    }

    # Fuente técnica completa para trazabilidad / depuración.
    (out/"results_debug.json").write_text(
        json.dumps(debug_payload,ensure_ascii=False,indent=2),encoding="utf-8"
    )

    # Contrato simplificado para integración.
    simple_payload=build_simple_results(results,VERSION)
    (out/"results.json").write_text(
        json.dumps(simple_payload,ensure_ascii=False,indent=2),encoding="utf-8"
    )

    alerts = build_alerts_from_results(results)
    (out/"alerts.json").write_text(
        json.dumps(
            {
                "schema_version": OUTPUT_SCHEMA_VERSION,
                "pipeline_version": VERSION,
                "alert_count": len(alerts),
                "alerts": alerts,
            },
            ensure_ascii=False,
            indent=2,
        ),
        encoding="utf-8",
    )


    audit=[build_detailed_audit_row(r) for r in results]
    if audit:
        with (out/"audit_log.csv").open("w",newline="",encoding="utf-8-sig") as f:
            wr=csv.DictWriter(f,fieldnames=audit[0].keys())
            wr.writeheader()
            wr.writerows(audit)

    consent_records=[
        r for r in results
        if r.get("record_type")=="consent_page"
    ]
    summary={
        "schema_version":OUTPUT_SCHEMA_VERSION,
        "pipeline_version":VERSION,
        "pdfs_input":len(pdfs),
        "consent_pages_recognized":len(consent_records),
        "unique_consent_pages":sum(not bool(r.get("is_exact_duplicate")) for r in consent_records),
        "exact_duplicate_pages":sum(bool(r.get("is_exact_duplicate")) for r in consent_records),
        "template_pages_rejected":sum(len(x.get("rejected_pages",[])) for x in pdf_summaries),
        "unsupported_pdfs":sum(x.get("status")=="UNSUPPORTED_TEMPLATE" for x in pdf_summaries),
        "stamps_applied":0,
        "slots_ready":sum(r.get("stamp_stage",{}).get("status")=="SLOT_READY" for r in consent_records),
        "existing_lower_stamp_detected":sum(r.get("stamp_stage",{}).get("status")=="EXISTING_STAMP_DETECTED" for r in consent_records),
        "manual_review_pages":sum(bool(r.get("manual_review_required")) for r in consent_records),
        "clean_pdfs":sum(x.get("status")=="CLEAN" for x in pdf_summaries),
        "manual_review_pdfs":sum(x.get("status")!="CLEAN" for x in pdf_summaries),
        "doctor_identity_verification":False,
        "excel_used":False,
        "originals_modified":0,
        "alerts_total":len(alerts),
        "pdf_export_enabled":not args.no_pdf_export,
    }
    (out/"summary.json").write_text(
        json.dumps(summary,ensure_ascii=False,indent=2),encoding="utf-8"
    )

    print()
    print("="*72)
    print("INDISA V2.5.3 AUTÓNOMA - RESUMEN")
    print("="*72)
    for k,v in summary.items():
        print(f"{k}: {v}")
    print("Resultados simples:",out/"results.json")
    print("Resultados debug:",out/"results_debug.json")
    print("Alertas:",out/"alerts.json")
    print("PDF limpio:",out/"pdfs","o",out/"manual_review")
    print("PDF FINAL CON CAJAS:",out/"final_pdfs")
    print("Auditoría:",out/"audit_log.csv")

if __name__=="__main__":
    main()
