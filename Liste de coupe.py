"""
Parcourt le composant actif : pour chaque corps (madrier), calcule la coupe
transversale en intersectant un plan perpendiculaire à l'arête la plus longue.
Détecte les coupes en angle aux bouts (angle, face concernée, parallélisme).
Regroupe par section puis par (longueur, note), tri décroissant.
"""

import math
import os
import time
import html as _html_mod
import traceback
from pathlib import Path

import adsk.core, adsk.fusion

CM_TO_IN = 1.0 / 2.54
SQUARE_TOL_DEG = 0.5   # en-dessous de cette valeur → coupe considérée droite

BOARD_SIZES_IN = [96, 120, 144, 168, 192]          # 8, 10, 12, 14, 16 pieds
BOARD_LABELS   = {96: "8 pi", 120: "10 pi", 144: "12 pi", 168: "14 pi", 192: "16 pi"}


def _min_board_size_in(length_in):
    """Retourne la taille minimale de madrier (pouces) pour la longueur donnée."""
    for size in BOARD_SIZES_IN:
        if length_in <= size + 0.001:
            return size
    return None  # pièce plus longue que 16 pi

_handlers            = []    # garde les handlers en vie tant que la palette est ouverte
_gid_to_bodies:       dict = {}  # str(gid) → [BRepBody, ...]
_body_token_to_gid:   dict = {}  # entityToken → str(gid)  (reverse mapping)
_palette_ref               = None
_ignore_selection_change: bool = False


# ── Event handlers ─────────────────────────────────────────────────────────────

class _IncomingHandler(adsk.core.HTMLEventHandler):
    """Reçoit les messages depuis la palette HTML."""
    def notify(self, args):
        global _ignore_selection_change
        try:
            ea = adsk.core.HTMLEventArgs.cast(args)
            if ea.action == 'highlight':
                gid    = (ea.data or "").strip()
                bodies = _gid_to_bodies.get(gid, [])
                if not bodies:
                    return
                app = adsk.core.Application.get()
                ui  = app.userInterface
                _ignore_selection_change = True
                try:
                    ui.activeSelections.clear()
                    for body in bodies:
                        try:
                            ui.activeSelections.add(body)
                        except Exception:
                            pass
                finally:
                    _ignore_selection_change = False
            elif ea.action == 'export':
                _handle_export(ea.data or '')
        except Exception:
            pass


class _SelectionChangedHandler(adsk.core.ActiveSelectionEventHandler):
    """Détecte la sélection d'un corps dans le viewport et highlight la ligne correspondante."""
    def notify(self, args):
        global _ignore_selection_change, _palette_ref
        if _ignore_selection_change:
            return
        try:
            app  = adsk.core.Application.get()
            ui   = app.userInterface
            sels = ui.activeSelections
            if sels.count == 0:
                return
            entity = sels.item(0).entity
            try:
                # Remonter au corps parent si face ou arête
                if isinstance(entity, adsk.fusion.BRepFace) or isinstance(entity, adsk.fusion.BRepEdge):
                    entity = entity.body
                tok = entity.entityToken
            except Exception:
                return
            gid = _body_token_to_gid.get(tok)
            if gid is None:
                return
            if _palette_ref and _palette_ref.isValid:
                _palette_ref.sendInfoToHTML('selectRow', gid)
        except Exception:
            pass


def _handle_export(data_json):
    """Reçoit le contenu depuis la palette et l'enregistre via une boîte de dialogue Fusion."""
    import json
    try:
        app = adsk.core.Application.get()
        ui  = app.userInterface
        payload = json.loads(data_json)
        fmt     = payload.get('format', 'txt')
        content = payload.get('content', '')

        ext_labels = {
            'txt': 'Fichier texte (*.txt)',
            'csv': 'Fichier CSV (*.csv)',
            'xls': 'Fichier Excel (*.xls)',
        }

        dlg = ui.createFileDialog()
        dlg.isMultiSelectEnabled = False
        dlg.title  = 'Exporter la liste de coupe'
        dlg.filter = ext_labels.get(fmt, '*.*')
        dlg.initialFilename = 'liste_coupe'
        result = dlg.showSave()
        if result == adsk.core.DialogResults.DialogOK:
            path = dlg.filename
            if not path.lower().endswith('.' + fmt):
                path += '.' + fmt
            with open(path, 'w', encoding='utf-8-sig') as f:
                f.write(content)
    except Exception:
        pass


class _ClosedHandler(adsk.core.UserInterfaceGeneralEventHandler):
    """Libère les handlers et termine le script quand la palette est fermée."""
    def notify(self, args):
        global _handlers
        _handlers.clear()
        adsk.terminate()


# ── Helpers numériques ─────────────────────────────────────────────────────────

def _fmt_num(val):
    s = f"{round(val, 2):.2f}".rstrip("0").rstrip(".")
    return s if s != "-0" else "0"


def _frac_str(val_in):
    """Convertit une valeur en pouces en chaîne fractionnaire arrondie au 1/16."""
    total_sixteenths = round(val_in * 16)
    whole = total_sixteenths // 16
    rem   = total_sixteenths % 16
    if rem == 0:
        return str(whole)
    g   = math.gcd(rem, 16)
    num = rem // g
    den = 16 // g
    if whole == 0:
        return f"{num}/{den}"
    return f"{whole} {num}/{den}"


def _fmt_in(val_cm):
    """Affichage en pouces fractionnaires, arrondi au 1/16 le plus proche."""
    return _frac_str(val_cm * CM_TO_IN) + " in"


def _iter_collection(coll):
    """Itère de façon robuste sur les collections Fusion (count/item)."""
    if not coll:
        return
    try:
        for i in range(coll.count):
            yield coll.item(i)
    except Exception:
        return


def _section_key_in_from_extents_cm(d0_cm, d1_cm):
    a = round(d0_cm * CM_TO_IN * 16) / 16
    b = round(d1_cm * CM_TO_IN * 16) / 16
    return (a, b)


def _fmt_section_title(section_key_in):
    a, b = section_key_in
    return f"{_frac_str(a)} in X {_frac_str(b)} in"


# ── Analyse géométrique ────────────────────────────────────────────────────────

def _bbox_fallback_section_cm(body):
    """Boîte AABB — dernier recours si la coupe perpendiculaire échoue."""
    bb = body.boundingBox
    if not bb:
        return 0.0, 0.0
    mn   = bb.minPoint
    mx   = bb.maxPoint
    dims = sorted((abs(mx.x - mn.x), abs(mx.y - mn.y), abs(mx.z - mn.z)))
    return dims[0], dims[1]


def _cross_section_full(body):
    """
    Calcule la coupe transversale via intersection d'un plan perpendiculaire
    à l'arête la plus longue (au point milieu).
    Retourne un dict avec :
      d_min_cm, d_max_cm : dimensions de la coupe (cm)
      Lcm               : longueur de l'arête la plus longue (cm)
      u                 : (ux, uy, uz) direction du madrier
      v_min             : direction 3D dans le plan de coupe ↔ d_min
      v_max             : direction 3D dans le plan de coupe ↔ d_max
    Retourne None en cas d'échec (bascule sur AABB dans l'appelant).
    """
    # 1. Arête la plus longue : direction et milieu
    best_L  = 0.0
    best_sp = best_ep = None
    for edge in _iter_collection(body.edges):
        try:
            L = edge.length
            if L <= best_L:
                continue
            sv   = edge.startVertex
            ev_v = edge.endVertex
            if sv is None or ev_v is None:
                continue
            best_L  = L
            best_sp = sv.geometry
            best_ep = ev_v.geometry
        except Exception:
            continue

    if best_sp is None:
        return None

    dx   = best_ep.x - best_sp.x
    dy   = best_ep.y - best_sp.y
    dz   = best_ep.z - best_sp.z
    norm = (dx**2 + dy**2 + dz**2)**0.5
    if norm < 1e-10:
        return None
    ux, uy, uz = dx/norm, dy/norm, dz/norm

    midx = (best_sp.x + best_ep.x) / 2.0
    midy = (best_sp.y + best_ep.y) / 2.0
    midz = (best_sp.z + best_ep.z) / 2.0

    # 2. Intersection de chaque arête avec le plan (normal=u, point=milieu)
    pts = []
    for edge in _iter_collection(body.edges):
        try:
            sv   = edge.startVertex
            ev_v = edge.endVertex
            if sv is None or ev_v is None:
                continue
            p0  = sv.geometry
            p1  = ev_v.geometry
            edx = p1.x - p0.x
            edy = p1.y - p0.y
            edz = p1.z - p0.z
            denom = edx*ux + edy*uy + edz*uz
            numer = (midx-p0.x)*ux + (midy-p0.y)*uy + (midz-p0.z)*uz
            if abs(denom) < 1e-10:
                if abs(numer) < 1e-4:
                    pts.append((p0.x, p0.y, p0.z))
                    pts.append((p1.x, p1.y, p1.z))
            else:
                t = numer / denom
                if -1e-4 <= t <= 1.0 + 1e-4:
                    pts.append((p0.x+t*edx, p0.y+t*edy, p0.z+t*edz))
        except Exception:
            continue

    if len(pts) < 2:
        return None

    # 3. Repère 2D aligné sur les faces latérales du madrier.
    cs_dirs = []
    for face in _iter_collection(body.faces):
        try:
            surf = face.geometry
            if not isinstance(surf, adsk.core.Plane):
                continue
            n  = surf.normal
            nn = (n.x**2 + n.y**2 + n.z**2)**0.5
            if nn < 1e-10:
                continue
            nx, ny, nz = n.x/nn, n.y/nn, n.z/nn
            if abs(nx*ux + ny*uy + nz*uz) > 0.5:
                continue
            matched = False
            for cd in cs_dirs:
                if abs(nx*cd[0] + ny*cd[1] + nz*cd[2]) > 0.9:
                    matched = True
                    break
            if not matched:
                cs_dirs.append((nx, ny, nz))
                if len(cs_dirs) == 2:
                    break
        except Exception:
            continue

    if len(cs_dirs) == 2:
        v1x, v1y, v1z = cs_dirs[0]
        v2x, v2y, v2z = cs_dirs[1]
    else:
        ax = abs(ux); ay = abs(uy); az = abs(uz)
        if ax <= ay and ax <= az:
            wx, wy, wz = 1.0, 0.0, 0.0
        elif ay <= az:
            wx, wy, wz = 0.0, 1.0, 0.0
        else:
            wx, wy, wz = 0.0, 0.0, 1.0
        dot = wx*ux + wy*uy + wz*uz
        v1x = wx - dot*ux;  v1y = wy - dot*uy;  v1z = wz - dot*uz
        v1n = (v1x**2 + v1y**2 + v1z**2)**0.5
        v1x, v1y, v1z = v1x/v1n, v1y/v1n, v1z/v1n
        v2x = uy*v1z - uz*v1y
        v2y = uz*v1x - ux*v1z
        v2z = ux*v1y - uy*v1x

    # 4. Projection et bounding box 2D
    s_vals = []
    t_vals = []
    for (px, py, pz) in pts:
        rx = px-midx;  ry = py-midy;  rz = pz-midz
        s_vals.append(rx*v1x + ry*v1y + rz*v1z)
        t_vals.append(rx*v2x + ry*v2y + rz*v2z)

    w1 = max(s_vals) - min(s_vals)
    w2 = max(t_vals) - min(t_vals)

    if w1 <= w2:
        d_min_cm, d_max_cm = w1, w2
        v_min = (v1x, v1y, v1z)
        v_max = (v2x, v2y, v2z)
    else:
        d_min_cm, d_max_cm = w2, w1
        v_min = (v2x, v2y, v2z)
        v_max = (v1x, v1y, v1z)

    return {
        "d_min_cm": d_min_cm,
        "d_max_cm": d_max_cm,
        "Lcm":      best_L,
        "u":        (ux, uy, uz),
        "v_min":    v_min,
        "v_max":    v_max,
    }


def _build_cut_note(body, u, d_min_in, d_max_in, v_min, v_max):
    """
    Analyse les faces en bout du madrier pour détecter les coupes en angle.
    Pour chaque face en bout :
      - angle = arccos(|dot(n_face, u)|)  → déviation par rapport à la coupe droite
      - axe de rotation r = normalize(u × n) :
          si r ≅ v_min → coupe sur l'arête d_min (face étroite)
          si r ≅ v_max → coupe sur l'arête d_max (face large)
    """
    ux, uy, uz = u
    end_faces  = []

    for face in _iter_collection(body.faces):
        try:
            surf = face.geometry
            if not isinstance(surf, adsk.core.Plane):
                continue
            n  = surf.normal
            nn = (n.x**2 + n.y**2 + n.z**2)**0.5
            if nn < 1e-10:
                continue
            nx, ny, nz = n.x/nn, n.y/nn, n.z/nn

            abs_dot_u = abs(nx*ux + ny*uy + nz*uz)
            if abs_dot_u < 0.5:
                continue

            angle_deg = math.degrees(math.acos(min(1.0, abs_dot_u)))
            rx = uy*nz - uz*ny
            ry = uz*nx - ux*nz
            rz = ux*ny - uy*nx
            rn = (rx**2 + ry**2 + rz**2)**0.5

            if rn < 1e-6 or angle_deg < SQUARE_TOL_DEG:
                face_dim_in = None
            else:
                rx, ry, rz = rx/rn, ry/rn, rz/rn
                dot_vmin    = abs(rx*v_min[0] + ry*v_min[1] + rz*v_min[2])
                dot_vmax    = abs(rx*v_max[0] + ry*v_max[1] + rz*v_max[2])
                face_dim_in = d_min_in if dot_vmin >= dot_vmax else d_max_in

            bb = face.boundingBox
            if bb:
                cx   = (bb.minPoint.x + bb.maxPoint.x) / 2
                cy   = (bb.minPoint.y + bb.maxPoint.y) / 2
                cz   = (bb.minPoint.z + bb.maxPoint.z) / 2
                proj = cx*ux + cy*uy + cz*uz
            else:
                proj = float(len(end_faces))

            end_faces.append({"angle": angle_deg, "dim": face_dim_in,
                               "proj": proj, "n": (nx, ny, nz)})
        except Exception:
            continue

    if not end_faces:
        return "Coupé droit"

    end_faces.sort(key=lambda f: f["proj"])
    angled = [f for f in end_faces if f["angle"] >= SQUARE_TOL_DEG]

    if not angled:
        return "Coupé droit"

    def _dim_str(dim_in):
        return f" — sur arête {_fmt_num(dim_in)} in" if dim_in is not None else ""

    if len(angled) == 1:
        f   = angled[0]
        ang = round(f["angle"], 1)
        return f"Coupé à {ang}° à un bout{_dim_str(f['dim'])}"

    if len(angled) == 2:
        f1, f2   = angled[0], angled[1]
        ang1     = round(f1["angle"], 1)
        ang2     = round(f2["angle"], 1)
        n1, n2   = f1["n"], f2["n"]
        parallel = abs(n1[0]*n2[0] + n1[1]*n2[1] + n1[2]*n2[2]) > 0.9998

        if abs(ang1 - ang2) < 0.2:
            par_str   = "parallèles" if parallel else "non parallèles"
            dim1, dim2 = f1["dim"], f2["dim"]
            if dim1 == dim2:
                dim_part = _dim_str(dim1)
            elif dim1 is not None and dim2 is not None:
                dim_part = f" — sur arêtes {_fmt_num(dim1)} in et {_fmt_num(dim2)} in"
            else:
                dim_part = _dim_str(dim1 if dim1 is not None else dim2)
            return f"Coupé à {ang1}° aux 2 bouts — {par_str}{dim_part}"

        parts = []
        if f1["dim"] is not None:
            parts.append(f"{ang1}°: {_fmt_num(f1['dim'])} in")
        if f2["dim"] is not None:
            parts.append(f"{ang2}°: {_fmt_num(f2['dim'])} in")
        detail = f" ({', '.join(parts)})" if parts else ""
        return f"Coupé à {ang1}° et {ang2}° aux 2 bouts{detail}"

    angles_str = " et ".join(f"{round(f['angle'], 1)}°" for f in angled)
    return f"Coupé à {angles_str} en bout"


def _collect_visible_bodies(root_comp):
    """
    Collecte tous les corps visibles dans le design :
    - Corps de la composante racine dont l'ampoule est allumée.
    - Corps de toutes les occurrences (hiérarchie aplatie) dont l'occurrence
      est visible ET dont l'ampoule du corps est allumée.
    Déduplication par entityToken (un même composant peut être instancié
    plusieurs fois).
    """
    seen   = set()
    bodies = []

    for body in _iter_collection(root_comp.bRepBodies):
        try:
            if not body or not body.isLightBulbOn:
                continue
            tok = body.entityToken
            if tok not in seen:
                seen.add(tok)
                bodies.append(body)
        except Exception:
            continue

    for occ in _iter_collection(root_comp.allOccurrences):
        try:
            if not occ.isVisible:
                continue
            # occ.bRepBodies retourne les proxy bodies (scopées à l'occurrence) ;
            # c'est indispensable pour que ui.activeSelections.add(body) fonctionne
            # et highlight le bon corps dans le viewport.
            for body in _iter_collection(occ.bRepBodies):
                if not body or not body.isLightBulbOn:
                    continue
                tok = body.entityToken
                if tok not in seen:
                    seen.add(tok)
                    bodies.append(body)
        except Exception:
            continue

    return bodies


# ── Liste de matériaux ─────────────────────────────────────────────────────────

def _compute_materials(sections_ordered):
    """
    Algorithme First Fit Decreasing (FFD) par section de bois.
    Étape 1 : pièces > 8 pi → madrier de la taille minimale requise.
    Étape 2 : retailles utilisées pour les pièces plus courtes.
    Étape 3 : reste placé dans des madriers de 8 pi.

    Retourne : list of (section_title, boards, summary)
      boards  : list of dict {size_in, remaining_in, pieces: [(len_in, note), ...]}
      summary : dict {board_size_in: count}
    """
    result = []
    for section_title, rows in sections_ordered:
        pieces = []
        for (len_key, note, qty, _Lcm, _bodies) in rows:
            for _ in range(qty):
                pieces.append((len_key, note))

        pieces.sort(key=lambda p: -p[0])  # plus long en premier

        boards = []
        for (length_in, note) in pieces:
            min_size = _min_board_size_in(length_in)
            if min_size is None:
                continue  # pièce > 16 pi, ignorée

            placed = False
            for board in boards:
                if board["remaining_in"] >= length_in - 0.001:
                    board["remaining_in"] -= length_in
                    board["pieces"].append((length_in, note))
                    placed = True
                    break

            if not placed:
                boards.append({
                    "size_in":      min_size,
                    "remaining_in": min_size - length_in,
                    "pieces":       [(length_in, note)],
                })

        summary = {}
        for board in boards:
            sz = board["size_in"]
            summary[sz] = summary.get(sz, 0) + 1

        result.append((section_title, boards, summary))
    return result


def _build_mat_html(materials):
    """Génère le HTML de la liste de matériaux (onglet Matériaux)."""
    parts = []
    for (section_title, boards, summary) in materials:
        sec_esc = _html_mod.escape(section_title)

        count_html = ""
        if summary:
            sum_parts = [
                f'{cnt}&times;&#160;{BOARD_LABELS.get(sz, str(sz) + " in")}'
                for sz, cnt in sorted(summary.items(), reverse=True)
            ]
            count_html = (
                f'<span class="mat-sec-count">'
                f'(Madriers&#160;: {"  +  ".join(sum_parts)})'
                f'</span>'
            )

        parts.append(
            f'<div class="mat-sec-hdr" onclick="toggleMatSec(this)">'
            f'<span>{sec_esc}</span>'
            f'{count_html}'
            f'<span class="mat-sec-arrow">&#9654;</span>'
            f'</div>'
            f'<div class="mat-sec-body" style="display:none">'
        )

        for i, board in enumerate(boards, 1):
            size_label = BOARD_LABELS.get(board["size_in"], f'{board["size_in"]} in')
            remaining  = board["remaining_in"]
            rem_str    = (_frac_str(remaining) + " in") if remaining > 0.05 else "&#8212;"

            board_parts = [
                f'<div class="mat-board-hdr">'
                f'<span class="mat-bnum">#{i}</span>'
                f'<span class="mat-bsize">Madrier {size_label}</span>'
                f'<span class="mat-rem">retaille&#160;: {rem_str}</span>'
                f'</div>'
            ]
            for (lin, note) in board["pieces"]:
                len_str  = _frac_str(lin) + " in"
                note_esc = _html_mod.escape(note)
                board_parts.append(
                    f'<div class="mat-piece">'
                    f'<span class="mat-plen">{len_str}</span>'
                    f'<span class="mat-pnote">{note_esc}</span>'
                    f'</div>'
                )
            parts.append('<div class="mat-board">' + "".join(board_parts) + '</div>')

        parts.append('</div>')  # close mat-sec-body

    return "\n".join(parts)


# ── Génération HTML ────────────────────────────────────────────────────────────

def _build_html(body_count, sections_ordered):
    """
    Construit le contenu HTML de la palette (onglets Coupes + Matériaux).

    sections_ordered : list of (section_title_str, rows)
    rows             : list of (len_key, note, qty, Lcm, bodies_list)

    Retourne html_str.
    Le gid (entier) est embarqué dans data-gid de chaque ligne ; le JS envoie
    uniquement ce gid à Python qui résout les corps via _gid_to_bodies.
    """
    # Résumé matériaux par section (calculé avant la boucle pour injecter dans l'en-tête)
    materials = _compute_materials(sections_ordered)
    sec_mat_sum: dict = {}
    for (sec_title, _boards, summary) in materials:
        if summary:
            parts = [f'{cnt}&times;&#160;{BOARD_LABELS.get(sz, str(sz)+" in")}'
                     for sz, cnt in sorted(summary.items(), reverse=True)]
            total = sum(summary.values())
            sec_mat_sum[sec_title] = (
                "  +  ".join(parts)
                + f'&#160;&#160;=&#160;&#160;{total}'
                + f'&#160;madrier{"s" if total > 1 else ""}'
            )

    # ── Onglet Coupes ──────────────────────────────────────────────────────────
    rows_html = []
    gid       = 0

    for section_title, rows in sections_ordered:
        sec_esc  = _html_mod.escape(section_title)
        sum_str  = sec_mat_sum.get(section_title, "")
        sum_html = (f'<span class="sec-mat-sum">{sum_str}</span>' if sum_str else "")
        rows_html.append(
            f'<tr class="sec-hdr"><td colspan="4">{sec_esc}{sum_html}</td></tr>'
            f'<tr class="col-hdr">'
            f'<th></th><th>Qté</th><th>Longueur</th><th>Note</th>'
            f'</tr>'
        )
        for (len_key, note, qty, Lcm, _bodies) in rows:
            len_str  = _fmt_in(Lcm)
            note_esc = _html_mod.escape(note)
            skey_esc = _html_mod.escape(
                f"{section_title}|{len_key:.4f}|{note}", quote=True
            )
            rows_html.append(
                f'<tr class="cut-row" data-gid="{gid}" data-skey="{skey_esc}"'
                f' onclick="rowClick(event,this)">'
                f'<td class="cb-cell">'
                f'<input type="checkbox" onclick="event.stopPropagation()"'
                f' onchange="cbChange(this)">'
                f'</td>'
                f'<td class="qty">{qty}</td>'
                f'<td class="len">{len_str}</td>'
                f'<td class="note">{note_esc}</td>'
                f'</tr>'
            )
            gid += 1

    rows_str = "\n".join(rows_html)

    # ── Onglet Matériaux ───────────────────────────────────────────────────────
    mat_str = _build_mat_html(materials)  # materials déjà calculé ci-dessus

    return f"""<!DOCTYPE html>
<html><head><meta charset="utf-8"><style>
  *{{box-sizing:border-box;margin:0;padding:0}}
  body{{font-family:'Courier New',monospace;font-size:13px;
       padding:10px;background:#2b2b2b;color:#e8e8e8}}
  #toolbar{{position:sticky;top:0;z-index:10;background:#2b2b2b;
            padding:6px 0 10px;border-bottom:1px solid #444;
            margin-bottom:10px;display:flex;align-items:center;gap:8px;flex-wrap:wrap}}
  .btn{{background:#3a7bd5;color:#fff;border:none;padding:5px 14px;
        font-size:12px;border-radius:4px;cursor:pointer;font-family:sans-serif}}
  .btn:hover{{background:#2a6bc5}}
  .btn-warn{{background:#8b2020}}
  .btn-warn:hover{{background:#b02828}}
  .btn-tab{{background:#3a3a3a;color:#bbb}}
  .btn-tab:hover{{background:#4a4a4a;color:#fff}}
  .btn-tab.active{{background:#3a7bd5;color:#fff}}
  .tab-sep{{width:1px;height:20px;background:#555;margin:0 4px}}
  #summary{{font-family:sans-serif;font-size:12px;color:#999;margin-bottom:10px}}
  table{{width:100%;border-collapse:collapse}}
  .sec-hdr td{{background:#1e1e1e;color:#f0a040;font-weight:bold;
               padding:8px 6px 3px;font-family:sans-serif;font-size:13px;
               border-top:2px solid #555}}
  .sec-mat-sum{{float:right;color:#aaffaa;font-size:11px;font-weight:normal;white-space:nowrap;padding-left:12px}}
  .col-hdr th{{background:#303030;color:#888;font-family:sans-serif;
               font-size:11px;font-weight:normal;padding:3px 6px;
               text-align:left;border-bottom:1px solid #3a3a3a}}
  .cut-row{{cursor:pointer;transition:background .1s}}
  .cut-row:hover{{background:#363636}}
  .cut-row.active{{background:#1a3a70}}
  .cut-row.done td:not(.cb-cell){{opacity:.4}}
  .cut-row.hidden-row{{display:none}}
  .cut-row td{{padding:3px 6px;vertical-align:middle;border-bottom:1px solid #2e2e2e}}
  .cb-cell{{width:28px;text-align:center}}
  .qty{{width:44px;text-align:right;padding-right:10px;color:#99ccff}}
  .len{{width:120px;color:#ffcc88}}
  .note{{color:#cccccc}}
  input[type=checkbox]{{cursor:pointer;accent-color:#5a9fd4}}
  /* ── Matériaux ── */
  .mat-sec-hdr{{background:#1e1e1e;color:#f0a040;font-weight:bold;
                padding:8px 6px 3px;font-family:sans-serif;font-size:13px;
                border-top:2px solid #555;margin-top:4px;
                cursor:pointer;display:flex;align-items:baseline;gap:10px;user-select:none}}
  .mat-sec-hdr:hover{{background:#252525}}
  .mat-sec-count{{color:#aaffaa;font-weight:normal;font-size:11px}}
  .mat-sec-arrow{{margin-left:auto;color:#888;font-size:11px;flex-shrink:0}}
  .mat-board{{margin:4px 0 10px 12px;border-left:2px solid #3a5070;padding-left:10px}}
  .mat-board-hdr{{display:flex;align-items:baseline;gap:10px;
                  padding:3px 0;font-family:sans-serif}}
  .mat-bnum{{color:#666;font-size:11px;min-width:22px}}
  .mat-bsize{{color:#99ccff;font-size:13px;font-weight:bold}}
  .mat-rem{{margin-left:auto;color:#888;font-size:11px}}
  .mat-piece{{display:flex;gap:14px;padding:1px 0 1px 4px;font-size:12px}}
  .mat-plen{{color:#ffcc88;min-width:90px;flex-shrink:0}}
  .mat-pnote{{color:#bbbbbb}}
</style></head><body>
  <div id="toolbar">
    <div id="tb-coupes">
      <button class="btn" id="btnToggle" onclick="toggleHide()">Masquer cochés</button>
      <button class="btn btn-warn" onclick="clearAll()">Tout décocher</button>
      <span style="color:#888;font-family:sans-serif;font-size:11px;margin-left:4px">Exporter :</span>
      <button class="btn" onclick="exportTxt()">Texte</button>
      <button class="btn" onclick="exportCsv()">CSV</button>
      <button class="btn" onclick="exportXls()">Excel</button>
    </div>
    <div class="tab-sep"></div>
    <button class="btn btn-tab active" id="tabCoupes" onclick="switchTab('coupes')">&#9998; Coupes</button>
    <button class="btn btn-tab" id="tabMat" onclick="switchTab('mat')">&#9744; Détails matériaux</button>
  </div>
  <div id="summary">Corps visibles : {body_count}</div>

  <div id="view-coupes">
    <table><tbody>
      {rows_str}
    </tbody></table>
  </div>

  <div id="view-mat" style="display:none">
    {mat_str}
  </div>

<script>
let hiding = false;

function toggleMatSec(hdr) {{
  const body = hdr.nextElementSibling;
  const arrow = hdr.querySelector('.mat-sec-arrow');
  const isOpen = body.style.display !== 'none';
  body.style.display = isOpen ? 'none' : '';
  arrow.innerHTML = isOpen ? '&#9654;' : '&#9660;';
}}

window.fusionJavaScriptHandler = {{
  handle: function(action, data) {{
    if (action === 'selectRow') {{
      const gid = String(data).trim();
      document.querySelectorAll('.cut-row.active').forEach(r => r.classList.remove('active'));
      const target = document.querySelector('.cut-row[data-gid="' + gid + '"]');
      if (target) {{
        target.classList.add('active');
        target.scrollIntoView({{ behavior: 'smooth', block: 'nearest' }});
      }}
    }}
    return 'ok';
  }}
}};

document.querySelectorAll('.cut-row').forEach(row => {{
  if (localStorage.getItem('lc_' + row.dataset.skey) === '1') {{
    row.querySelector('input').checked = true;
    row.classList.add('done');
  }}
}});

function switchTab(tab) {{
  const isCoupes = tab === 'coupes';
  document.getElementById('view-coupes').style.display = isCoupes ? '' : 'none';
  document.getElementById('view-mat').style.display    = isCoupes ? 'none' : '';
  document.getElementById('tb-coupes').style.display   = isCoupes ? '' : 'none';
  document.getElementById('tabCoupes').classList.toggle('active', isCoupes);
  document.getElementById('tabMat').classList.toggle('active', !isCoupes);
}}

function rowClick(e, row) {{
  document.querySelectorAll('.cut-row.active').forEach(r => r.classList.remove('active'));
  row.classList.add('active');
  const gid = row.dataset.gid;
  if (typeof adsk !== 'undefined' && adsk.fusionSendData) {{
    try {{ adsk.fusionSendData('highlight', gid); }} catch(err) {{}}
  }}
}}

function cbChange(cb) {{
  const row = cb.closest('tr');
  const key = 'lc_' + row.dataset.skey;
  if (cb.checked) {{
    localStorage.setItem(key, '1');
    row.classList.add('done');
    if (hiding) row.classList.add('hidden-row');
  }} else {{
    localStorage.removeItem(key);
    row.classList.remove('done', 'hidden-row');
  }}
}}

function toggleHide() {{
  hiding = !hiding;
  document.querySelectorAll('.cut-row.done').forEach(r =>
    r.classList.toggle('hidden-row', hiding));
  document.getElementById('btnToggle').textContent =
    hiding ? 'Afficher cochés' : 'Masquer cochés';
}}

function clearAll() {{
  hiding = false;
  document.getElementById('btnToggle').textContent = 'Masquer cochés';
  document.querySelectorAll('.cut-row').forEach(row => {{
    row.querySelector('input').checked = false;
    localStorage.removeItem('lc_' + row.dataset.skey);
    row.classList.remove('done', 'hidden-row', 'active');
  }});
}}

function getExportRows() {{
  const out = [];
  let section = '';
  document.querySelectorAll('table tbody tr').forEach(tr => {{
    if (tr.classList.contains('sec-hdr')) {{
      const td = tr.cells[0];
      const sumEl = td.querySelector('.sec-mat-sum');
      const sumRaw = sumEl ? sumEl.textContent.replace(/\u00a0/g, ' ').trim() : '';
      const sumText = sumRaw.replace(/\s*=\s*\d+\s*madriers?\s*$/i, '').trim();
      const titleText = sumEl
        ? td.textContent.replace(sumEl.textContent, '').trim()
        : td.textContent.trim();
      section = sumText ? titleText + ' (Madriers: ' + sumText + ')' : titleText;
    }} else if (tr.classList.contains('cut-row')) {{
      out.push({{
        section: section,
        qty: tr.querySelector('.qty').textContent.trim(),
        len: tr.querySelector('.len').textContent.trim(),
        note: tr.querySelector('.note').textContent.trim()
      }});
    }}
  }});
  return out;
}}

function exportFilename(ext) {{
  const d = new Date();
  const pad = n => String(n).padStart(2, '0');
  return 'liste_coupe_' + d.getFullYear() + '-' + pad(d.getMonth()+1) + '-' + pad(d.getDate())
    + '_' + pad(d.getHours()) + pad(d.getMinutes()) + pad(d.getSeconds()) + '.' + ext;
}}

function downloadBlob(filename, mime, parts) {{
  const blob = new Blob(parts, {{ type: mime }});
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = filename;
  a.click();
  URL.revokeObjectURL(a.href);
}}

function escapeHtml(s) {{
  return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}}

function csvCell(s) {{
  const t = String(s);
  if (/[;\\r\\n"]/.test(t)) return '"' + t.replace(/"/g,'""') + '"';
  return t;
}}

function sendExport(fmt, content) {{
  if (typeof adsk !== 'undefined' && adsk.fusionSendData) {{
    adsk.fusionSendData('export', JSON.stringify({{format: fmt, content: content}}));
  }}
}}

function exportTxt() {{
  const rows = getExportRows();
  let cur = null;
  const lines = [];
  rows.forEach(r => {{
    if (r.section !== cur) {{
      if (cur !== null) lines.push('');
      lines.push(r.section);
      cur = r.section;
    }}
    lines.push('  ' + r.qty.padStart(4, ' ') + '  ' + r.len.padEnd(14, ' ') + '  ' + r.note);
  }});
  sendExport('txt', lines.join(String.fromCharCode(10)));
}}

function exportCsv() {{
  const rows = getExportRows();
  const CRLF = String.fromCharCode(13) + String.fromCharCode(10);
  const lines = ['Section;Qté;Longueur;Note'];
  let curSection = null;
  rows.forEach(r => {{
    if (r.section !== curSection) {{
      if (curSection !== null) lines.push(';;;');
      lines.push([csvCell(r.section), csvCell(r.qty), csvCell(r.len), csvCell(r.note)].join(';'));
      curSection = r.section;
    }} else {{
      lines.push(['' , csvCell(r.qty), csvCell(r.len), csvCell(r.note)].join(';'));
    }}
  }});
  sendExport('csv', lines.join(CRLF));
}}

function exportXls() {{
  const rows = getExportRows();
  let h = '<html xmlns:o="urn:schemas-microsoft-com:office:office" '
    + 'xmlns:x="urn:schemas-microsoft-com:office:excel"><head><meta charset="utf-8">'
    + '<style>table{{border-collapse:collapse;font-family:Calibri,sans-serif}}'
    + 'td,th{{border:1px solid #ccc;padding:4px}}'
    + 'th{{background:#dde4f0;font-weight:bold}}'
    + '.sec-title{{background:#f5f0e8;font-weight:bold;color:#7a5000}}'
    + '.sep{{border:none}}'
    + '</style></head><body><table>';
  h += '<tr><th>Section</th><th>Qté</th><th>Longueur</th><th>Note</th></tr>';
  let curSection = null;
  rows.forEach(r => {{
    if (r.section !== curSection) {{
      if (curSection !== null)
        h += '<tr class="sep"><td colspan="4" style="border:none;padding:4px 0"></td></tr>';
      h += '<tr><td class="sec-title">' + escapeHtml(r.section) + '</td><td>' + escapeHtml(r.qty)
        + '</td><td>' + escapeHtml(r.len) + '</td><td>' + escapeHtml(r.note) + '</td></tr>';
      curSection = r.section;
    }} else {{
      h += '<tr><td></td><td>' + escapeHtml(r.qty)
        + '</td><td>' + escapeHtml(r.len) + '</td><td>' + escapeHtml(r.note) + '</td></tr>';
    }}
  }});
  h += '</table></body></html>';
  sendExport('xls', h);
}}
</script>
</body></html>"""


def _show_result(ui, title, body_count, sections_ordered):
    """Écrit le HTML, crée/recharge la palette et enregistre les event handlers."""
    global _handlers, _palette_ref
    _handlers.clear()

    script_dir = os.path.dirname(os.path.abspath(__file__))
    html_file  = os.path.join(script_dir, "_liste_coupe_result.html")

    html_str = _build_html(body_count, sections_ordered)

    with open(html_file, "w", encoding="utf-8") as f:
        f.write(html_str)

    # Cache-buster: force CEF à recharger le fichier plutôt que servir une copie cachée
    html_url   = Path(html_file).as_uri() + f"?t={int(time.time())}"
    palette_id = "listeCoupeResultPalette"

    existing = ui.palettes.itemById(palette_id)
    if existing:
        existing.deleteMe()

    palette = ui.palettes.add(palette_id, title, html_url, True, True, True, 760, 640)
    _palette_ref = palette

    h_in = _IncomingHandler()
    palette.incomingFromHTML.add(h_in)
    _handlers.append(h_in)

    h_cl = _ClosedHandler()
    palette.closed.add(h_cl)
    _handlers.append(h_cl)

    h_sel = _SelectionChangedHandler()
    ui.activeSelectionChanged.add(h_sel)
    _handlers.append(h_sel)


# ── Point d'entrée ─────────────────────────────────────────────────────────────

def run(_context: str):
    global _handlers, _gid_to_bodies, _body_token_to_gid, _palette_ref, _ignore_selection_change
    _handlers.clear()
    _gid_to_bodies.clear()
    _body_token_to_gid.clear()
    _palette_ref = None
    _ignore_selection_change = False

    # IMPORTANT : empêcher la fin automatique du script AVANT tout autre appel,
    # sinon le script meurt à la fin de run() et les handlers se détachent.
    try:
        adsk.autoTerminate(False)
    except Exception:
        pass

    app    = adsk.core.Application.get()
    ui     = app.userInterface
    design = adsk.fusion.Design.cast(app.activeProduct)

    try:
        if not design:
            if ui:
                ui.messageBox("Aucun design actif.", "Liste de coupe")
            return

        visible_bodies = _collect_visible_bodies(design.rootComponent)
        if not visible_bodies:
            if ui:
                ui.messageBox("Aucun corps visible trouvé.", "Liste de coupe")
            return

        # ── Phase 1 : collecte des données et comptage des sections par défaut ──
        # Nécessaire pour détecter si 2 arêtes d'un morceau correspondent à une
        # catégorie déjà créée par d'autres corps (prioritaire sur les 2 arêtes courtes).
        body_data_list = []
        sk_counts: dict = {}

        for body in visible_bodies:
            cs = _cross_section_full(body)
            if cs is None:
                d0, d1 = _bbox_fallback_section_cm(body)
                Lcm = 0.0
                for edge in _iter_collection(body.edges):
                    try:
                        L = edge.length
                        if L > Lcm:
                            Lcm = L
                    except Exception:
                        continue
                note      = "Coupé droit"
                u_vec     = None
                v_min_vec = None
                v_max_vec = None
            else:
                d0        = cs["d_min_cm"]
                d1        = cs["d_max_cm"]
                Lcm       = cs["Lcm"]
                u_vec     = cs["u"]
                v_min_vec = cs["v_min"]
                v_max_vec = cs["v_max"]
                d_min_in  = round(d0 * CM_TO_IN, 2)
                d_max_in  = round(d1 * CM_TO_IN, 2)
                note      = _build_cut_note(body, u_vec, d_min_in, d_max_in,
                                            v_min_vec, v_max_vec)

            if d0 < 1e-6 or d1 < 1e-6:
                continue

            default_sk = _section_key_in_from_extents_cm(d0, d1)
            sk_counts[default_sk] = sk_counts.get(default_sk, 0) + 1
            body_data_list.append((body, d0, d1, Lcm, note, default_sk,
                                   u_vec, v_min_vec, v_max_vec))

        # ── Phase 2 : assignation, en privilégiant les catégories existantes ──
        # Si 2 arêtes du morceau forment une clé de section déjà créée par
        # d'autres corps, on préfère cette section à la logique « 2 arêtes courtes ».
        # Quand une reclassification a lieu, la note de coupe est recalculée avec le
        # vecteur u de la nouvelle orientation pour correctement détecter les angles.
        by_section: dict = {}

        for body, d0, d1, Lcm, note, default_sk, u_vec, v_min_vec, v_max_vec \
                in body_data_list:
            d0_in = round(d0 * CM_TO_IN * 16) / 16
            d1_in = round(d1 * CM_TO_IN * 16) / 16
            L_in  = round(Lcm * CM_TO_IN * 16) / 16

            chosen_sk   = default_sk
            chosen_Lcm  = Lcm
            chosen_note = note

            # Alternatives : traiter d1 (ou d0) comme longueur et intégrer L dans la section.
            # alt A : section=(d0_in, L_in), longueur=d1, new_u=v_max_vec
            # alt B : section=(d1_in, L_in), longueur=d0, new_u=v_min_vec  (si d0 ≠ d1)
            alts = []
            if abs(d1_in - L_in) > 1e-4:
                alts.append(((min(d0_in, L_in), max(d0_in, L_in)), d1, 'A'))
            if abs(d0_in - L_in) > 1e-4 and abs(d0_in - d1_in) > 1e-4:
                alts.append(((min(d1_in, L_in), max(d1_in, L_in)), d0, 'B'))

            for alt_sk, alt_Lcm, alt_type in alts:
                if alt_sk != default_sk and sk_counts.get(alt_sk, 0) > 0:
                    chosen_sk  = alt_sk
                    chosen_Lcm = alt_Lcm
                    # Recalculer la note avec le vecteur u de la nouvelle orientation.
                    # alt A : v_max_vec devient la direction longueur ;
                    #         la nouvelle section est (d0, Lcm) avec v_min_vec et u_vec.
                    # alt B : v_min_vec devient la direction longueur ;
                    #         la nouvelle section est (d1, Lcm) avec v_max_vec et u_vec.
                    if u_vec is not None:
                        if alt_type == 'A':
                            new_u       = v_max_vec
                            new_v_min   = v_min_vec
                            new_v_max   = u_vec
                            new_dmin_in = round(d0  * CM_TO_IN, 2)
                            new_dmax_in = round(Lcm * CM_TO_IN, 2)
                        else:
                            new_u       = v_min_vec
                            new_v_min   = v_max_vec
                            new_v_max   = u_vec
                            new_dmin_in = round(d1  * CM_TO_IN, 2)
                            new_dmax_in = round(Lcm * CM_TO_IN, 2)
                        chosen_note = _build_cut_note(body, new_u,
                                                      new_dmin_in, new_dmax_in,
                                                      new_v_min, new_v_max)
                    break

            section_key = chosen_sk
            len_key     = round(chosen_Lcm * CM_TO_IN * 16) / 16
            group_key   = (len_key, chosen_note)

            if section_key not in by_section:
                by_section[section_key] = {}
            sec = by_section[section_key]
            if group_key not in sec:
                sec[group_key] = {"Lcm": chosen_Lcm, "qty": 0, "bodies": []}
            sec[group_key]["qty"]    += 1
            sec[group_key]["bodies"].append(body)

        # Construire sections_ordered (même ordre que l'affichage)
        # et peupler _gid_to_bodies avec des références directes aux corps
        sections_ordered = []
        gid = 0
        section_keys = sorted(by_section.keys(), key=lambda k: (-k[0], -k[1]))
        for section_key in section_keys:
            sec     = by_section[section_key]
            ordered = sorted(sec.items(), key=lambda kv: kv[0][0], reverse=True)
            rows    = []
            for (len_key, note), data in ordered:
                _gid_to_bodies[str(gid)] = data["bodies"]
                for body in data["bodies"]:
                    try:
                        _body_token_to_gid[body.entityToken] = str(gid)
                    except Exception:
                        pass
                rows.append((len_key, note, data["qty"], data["Lcm"], data["bodies"]))
                gid += 1
            sections_ordered.append((_fmt_section_title(section_key), rows))

        # Log console
        print(f"Corps visibles : {len(visible_bodies)}")
        for sec_title, rows in sections_ordered:
            print(f"\n{sec_title}:")
            for (len_key, note, qty, Lcm, _bodies) in rows:
                print(f"  {qty:>4}  {_fmt_in(Lcm):<14}  {note}")

    except Exception:
        tb = traceback.format_exc()
        print(tb)
        if ui:
            ui.messageBox("Erreur :\n" + tb, "Liste de coupe")
        return

    if ui:
        try:
            _show_result(ui, "Liste de coupe", len(visible_bodies), sections_ordered)
        except Exception:
            adsk.autoTerminate(True)
            tb = traceback.format_exc()
            print(tb)
            if ui:
                ui.messageBox("Erreur affichage :\n" + tb, "Liste de coupe")
