"""Generate compression self-sensing curves for ECC-style specimens.

Outputs a four-panel figure (load-displacement, resistivity-time,
fractional change in resistivity vs time, stress-time) and an Excel
workbook with the underlying data.

The electrical response is a *schematic model*: it follows the
qualitative trend reported for self-sensing cementitious composites
under monotonic compression (compaction lowers resistivity, damage near
peak raises it). It is NOT measured data and must not be presented as
experimental results.

Usage
-----
    python src/ecc_curves.py --synthetic --outdir outputs
    python src/ecc_curves.py --input examples/sample_load_displacement.csv --outdir outputs

CSV format: two columns, displacement (mm) and load (kN), optional header.
"""
from __future__ import annotations

import argparse
import csv
import math
import random
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from PIL import Image, ImageDraw, ImageFont

# --------------------------------------------------------------------------
# model
# --------------------------------------------------------------------------
# Resistivity ratio rho/rho0 as a function of d/d_peak.  The anchors encode
# the literature trend: compaction dominates before ~70% of peak
# displacement, damage takes over near and after the peak.
RHO_ANCHORS = [(0.0, 1.0), (0.35, 0.93), (0.70, 0.845),
               (1.0, 1.015), (1.4, 1.075), (2.2, 1.16)]


def _smoothstep(t: float) -> float:
    t = max(0.0, min(1.0, t))
    return t * t * (3.0 - 2.0 * t)


def rho_ratio(x: float) -> float:
    """Interpolate the resistivity ratio at normalised displacement x."""
    if x <= RHO_ANCHORS[0][0]:
        return RHO_ANCHORS[0][1]
    for (x0, y0), (x1, y1) in zip(RHO_ANCHORS, RHO_ANCHORS[1:]):
        if x <= x1:
            u = 0.5 - 0.5 * math.cos(math.pi * (x - x0) / (x1 - x0))
            return y0 + (y1 - y0) * u
    return RHO_ANCHORS[-1][1]


def synthetic_load(d: float, peak_load: float = 350.0,
                   peak_disp: float = 2.0) -> float:
    """Load-displacement shape: linear rise, rounded peak, softening tail."""
    if d <= 0:
        return 0.0
    x = d / peak_disp
    if x <= 0.75:
        return peak_load * x
    if x <= 1.0:
        h = 0.25
        t = (x - 0.75) / h
        h00 = 2 * t ** 3 - 3 * t ** 2 + 1
        h10 = t ** 3 - 2 * t ** 2 + t
        h01 = -2 * t ** 3 + 3 * t ** 2
        h11 = t ** 3 - t ** 2
        y = h00 * 0.75 + h10 * h + h01 * 1.0 + h11 * h * 0.0
        return peak_load * y
    return peak_load * (0.665 + 0.335 * math.exp(-(d - peak_disp) / 1.45))


def build_curve(displacement, load, rate=0.01, rho0=0.45, area=10000.0,
                noise=(0.15, 0.20, 0.25), seed=42, peak_disp=None):
    """Return rows of (t, d, P, sigma, rho, fcr) with optional segment noise.

    noise = (stage1, stage2, stage3) in percent of rho0, split at
    0.7*d_peak and d_peak.  Noise is AR(1)-correlated so it looks like a
    measurement trace rather than white noise.  Set all to 0 to disable.
    """
    if len(displacement) != len(load):
        raise ValueError("displacement and load must have the same length")
    if peak_disp is None:
        peak_disp = displacement[max(range(len(load)), key=load.__getitem__)]
    rng = random.Random(seed)
    state = 0.0
    rows = []
    for i, (d, p) in enumerate(zip(displacement, load)):
        x = d / peak_disp if peak_disp else 0.0
        stage = 0 if x <= 0.70 else (1 if x <= 1.0 else 2)
        sigma_pct = noise[stage] if i else 0.0
        state = 0.55 * state + math.sqrt(1 - 0.55 ** 2) * rng.gauss(0, sigma_pct)
        rho = rho0 * (1.0 + rho_ratio(x) - 1.0 + state / 100.0)
        rows.append((d / rate, d, p, p * 1000.0 / area, rho,
                     (rho / rho0 - 1.0) * 100.0))
    return rows, peak_disp


# --------------------------------------------------------------------------
# plotting
# --------------------------------------------------------------------------
def _font(size, bold=False):
    candidates = [
        r"C:\Windows\Fonts\msyhbd.ttc" if bold else r"C:\Windows\Fonts\msyh.ttc",
        r"C:\Windows\Fonts\simsun.ttc",
        "/System/Library/Fonts/PingFang.ttc",
        "/usr/share/fonts/truetype/noto/NotoSansCJK-Regular.ttc",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
    ]
    for path in candidates:
        try:
            return ImageFont.truetype(path, size)
        except Exception:
            continue
    return ImageFont.load_default()


def draw_panels(rows, peak_disp, caption, out_png):
    S = 2
    W, H = 2400, 1800
    img = Image.new("RGB", (W * S, H * S), "white")
    dr = ImageDraw.Draw(img)

    def fc(size, bold=False):
        return _font(size * S, bold)

    def fe(size):
        return _font(size * S)

    def left(x, y, s, f, fill="#222"):
        bb = dr.textbbox((0, 0), s, font=f)
        dr.text((x * S - bb[0], y * S - bb[1]), s, font=f, fill=fill)

    def right(x, y, s, f, fill="#222"):
        bb = dr.textbbox((0, 0), s, font=f)
        dr.text((x * S - (bb[2] - bb[0]) - bb[0], y * S - bb[1]), s, font=f, fill=fill)

    def center(x, y, s, f, fill="#222"):
        bb = dr.textbbox((0, 0), s, font=f)
        dr.text((x * S - (bb[2] - bb[0]) / 2 - bb[0],
                 y * S - (bb[3] - bb[1]) / 2 - bb[1]), s, font=f, fill=fill)

    def vertical(x, y, s, f, fill="#222"):
        bb = dr.textbbox((0, 0), s, font=f)
        tmp = Image.new("RGBA", (bb[2] - bb[0] + 10, bb[3] - bb[1] + 10), (255, 255, 255, 0))
        ImageDraw.Draw(tmp).text((5 - bb[0], 5 - bb[1]), s, font=f, fill=fill)
        tmp = tmp.rotate(90, expand=True)
        img.paste(tmp, (int(x * S - tmp.width / 2), int(y * S - tmp.height / 2)), tmp)

    L, R, T, B = 150, 70, 140, 190
    GX, GY = 185, 200
    PW = (W - L - R - GX) / 2
    PH = (H - T - B - GY) / 2
    colors = ["#2E5E8C", "#B03A2E", "#3C7A5A", "#B9770E"]
    tags = ["(a)", "(b)", "(c)", "(d)"]

    ts = [r[0] for r in rows]
    ds = [r[1] for r in rows]
    ps = [r[2] for r in rows]
    sg = [r[3] for r in rows]
    rho = [r[4] for r in rows]
    fcr = [r[5] for r in rows]
    t_peak = next(r[0] for r in rows if r[1] >= peak_disp)
    p_peak = max(ps)
    s_peak = max(sg)
    rho_min = min(rho)
    t_min = ts[rho.index(rho_min)]

    def frame(idx, title, xlab, ylab, xlim, ylim, xt, yt, series, shade,
              notes, stages, zero=False, dec=0):
        cx = L + (idx % 2) * (PW + GX)
        cy = T + (idx // 2) * (PH + GY)
        x0, x1, y0, y1 = cx, cx + PW, cy, cy + PH

        def X(v):
            return x0 + (v - xlim[0]) / (xlim[1] - xlim[0]) * PW

        def Y(v):
            return y1 - (v - ylim[0]) / (ylim[1] - ylim[0]) * PH

        for a, b in shade:
            dr.rectangle([X(a) * S, y0 * S, X(b) * S, y1 * S], fill="#F5F7FA")
        for v in yt:
            yy = Y(v)
            px = X(xlim[0])
            while px < X(xlim[1]):
                dr.line([px * S, yy * S, min(px + 8, X(xlim[1])) * S, yy * S],
                        fill="#DDDDDD", width=S)
                px += 16
            label = f"{v:.{dec}f}" if dec else f"{v:g}"
            right(cx - 20, yy - 9, label, fe(28), "#333")
            dr.line([(cx - 10) * S, yy * S, cx * S, yy * S], fill="#333", width=2 * S)
        for v in xt:
            xx = X(v)
            dr.line([xx * S, y1 * S, xx * S, (y1 + 9) * S], fill="#333", width=2 * S)
            center(xx, y1 + 31, f"{v:g}", fe(28), "#333")
        dr.rectangle([x0 * S, y0 * S, x1 * S, y1 * S], outline="#333", width=2 * S)
        if zero and ylim[0] < 0 < ylim[1]:
            dr.line([x0 * S, Y(0) * S, x1 * S, Y(0) * S], fill="#7F7F7F", width=2 * S)
        center((x0 + x1) / 2, y1 + 74, xlab, fc(32), "#222")
        vertical(cx - 118, (y0 + y1) / 2, ylab, fc(32), "#222")
        left(x0, y0 - 56, tags[idx], fc(33, True), "#111")
        left(x0 + 62, y0 - 53, title, fc(30), "#111")
        for xs, ys, color, width in series:
            dr.line([(X(a) * S, Y(b) * S) for a, b in zip(xs, ys)],
                    fill=color, width=int(round(width * S)), joint="curve")
        for ax, ay, txt, dx, dy in notes:
            px, py = X(ax), Y(ay)
            dr.ellipse([px * S - 4 * S, py * S - 4 * S, px * S + 4 * S, py * S + 4 * S],
                       fill="#444")
            left(px + dx, py + dy, txt, fc(25), "#333")
        for sx, txt in stages:
            center(X(sx), y0 + 22, txt, fc(24), "#8C99AA")

    x_max = max(ds) * 1.05
    p_max = p_peak * 1.15
    xt_a = [v for v in range(0, int(math.floor(x_max)) + 1)]
    step_a = 50 if p_max <= 500 else 100
    yt_a = [v for v in range(0, int(math.floor(p_max)) + 1, step_a)]
    frame(0, "荷载–位移曲线", "位移 d (mm)", "荷载 P (kN)", (0, x_max),
          (0, p_max), xt_a, yt_a,
          [(ds, ps, colors[0], 4.4)],
          [(0, 0.5), (0.5, peak_disp), (peak_disp, max(ds))],
          [(peak_disp, p_peak, f"峰值 {p_peak:.0f} kN（d = {peak_disp:g} mm）", 26, -46)],
          [(0.25, "Ⅰ 压密"), (0.5 * peak_disp, "Ⅱ 峰前"), (0.75 * max(ds), "Ⅲ 峰后")])

    t_max = max(ts) * 1.02
    step_t = 100 if t_max <= 600 else 200
    xt_t = [v for v in range(0, int(math.floor(t_max)) + 1, step_t)]
    frame(1, "电阻率–时间曲线", "时间 t (s)", "电阻率 ρ (Ω·m)", (0, t_max),
          (min(rho) * 0.97, max(rho) * 1.03), xt_t,
          [round(min(rho), 2), round((min(rho) + max(rho)) / 2, 2), round(max(rho), 2)],
          [(ts, rho, colors[1], 4.4)],
          [(0, 0.12 * max(ts)), (0.12 * max(ts), t_peak), (t_peak, max(ts))],
          [(t_min, rho_min, "最低点（压密主导）", 14, 18)],
          [(0.05 * max(ts), "Ⅰ 压密"), (0.32 * max(ts), "Ⅱ 峰前"), (0.78 * max(ts), "Ⅲ 峰后")],
          dec=2)

    fcr_lo = math.floor(min(fcr) / 5.0) * 5.0
    fcr_hi = math.ceil(max(fcr) / 5.0) * 5.0
    yt_c = [v for v in range(int(fcr_lo), int(fcr_hi) + 5, 5)]
    frame(2, "相对电阻变化率–时间曲线", "时间 t (s)", "FCR (%)", (0, t_max),
          (fcr_lo, fcr_hi), xt_t, yt_c,
          [(ts, fcr, colors[2], 4.4)],
          [(0, 0.12 * max(ts)), (0.12 * max(ts), t_peak), (t_peak, max(ts))],
          [(t_peak, 1.9, "峰值处 FCR 由负转正", 30, 112)],
          [(0.05 * max(ts), "Ⅰ 压密"), (0.32 * max(ts), "Ⅱ 峰前"), (0.78 * max(ts), "Ⅲ 峰后")],
          zero=True)

    frame(3, "应力–时间曲线", "时间 t (s)", "应力 σ (MPa)", (0, t_max),
          (0, s_peak * 1.15), xt_t, [0, 10, 20, 30, 40],
          [(ts, sg, colors[3], 4.4)],
          [(0, 0.12 * max(ts)), (0.12 * max(ts), t_peak), (t_peak, max(ts))],
          [(t_peak, s_peak, f"峰值应力 {s_peak:.0f} MPa", 26, -44)],
          [(0.05 * max(ts), "Ⅰ 压密"), (0.32 * max(ts), "Ⅱ 峰前"), (0.78 * max(ts), "Ⅲ 峰后")])

    for idx in (1, 2, 3):
        cx = L + (idx % 2) * (PW + GX)
        cy = T + (idx // 2) * (PH + GY)
        x0 = cx
        y0 = cy
        xx = x0 + (t_peak / (max(ts) * 1.02)) * PW
        dr.line([xx * S, y0 * S, xx * S, (y0 + PH) * S], fill="#AAB4C0", width=2 * S)

    center(W / 2, H - 40, caption, fc(24), "#666")
    img.resize((W, H), Image.LANCZOS).save(out_png, dpi=(300, 300))


# --------------------------------------------------------------------------
# export
# --------------------------------------------------------------------------
def export_workbook(rows, path, meta):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "data"
    ws.append(["time_s", "displacement_mm", "load_kN", "stress_MPa",
               "resistivity_ohm_m", "FCR_percent"])
    for r in rows:
        ws.append([round(v, 6) for v in r])
    for c in ws[1]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="1F3864")
        c.alignment = Alignment(horizontal="center")
    ws.freeze_panes = "A2"

    p = wb.create_sheet("parameters")
    p.append(["item", "value"])
    for k, v in meta.items():
        p.append([k, v])
    p.column_dimensions["A"].width = 28
    p.column_dimensions["B"].width = 80
    for row in p.iter_rows(min_row=2):
        row[1].alignment = Alignment(wrap_text=True, vertical="top")
    wb.save(path)


def read_csv(path):
    ds, ps = [], []
    with open(path, newline="", encoding="utf-8-sig") as fh:
        for row in csv.reader(fh):
            if not row or row[0].strip().lower().startswith(("disp", "displacement", "位移")):
                continue
            try:
                ds.append(float(row[0]))
                ps.append(float(row[1]))
            except (ValueError, IndexError):
                continue
    if len(ds) < 3:
        raise ValueError(f"need at least 3 numeric rows in {path}")
    return ds, ps


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--input", help="CSV with displacement (mm), load (kN)")
    ap.add_argument("--synthetic", action="store_true", help="use the built-in demo curve")
    ap.add_argument("--outdir", default="outputs")
    ap.add_argument("--rate", type=float, default=0.01, help="loading rate, mm/s")
    ap.add_argument("--area", type=float, default=10000.0, help="cross-section, mm^2")
    ap.add_argument("--rho0", type=float, default=0.45, help="initial resistivity, ohm*m")
    ap.add_argument("--noise", default="0.15,0.20,0.25",
                    help="stage noise, percent of rho0: s1,s2,s3")
    ap.add_argument("--seed", type=int, default=42)
    ap.add_argument("--peak-disp", type=float, default=None)
    ap.add_argument("--no-figure", action="store_true")
    args = ap.parse_args()

    outdir = Path(args.outdir)
    outdir.mkdir(parents=True, exist_ok=True)

    if args.input:
        ds, ps = read_csv(args.input)
        source = str(args.input)
    else:
        ds = [round(i * 0.02, 4) for i in range(221)]
        ps = [synthetic_load(d) for d in ds]
        source = "built-in synthetic demo curve"

    noise = tuple(float(v) for v in args.noise.split(","))
    if len(noise) != 3:
        raise SystemExit("--noise needs three comma-separated values")

    rows, peak_disp = build_curve(ds, ps, rate=args.rate, rho0=args.rho0,
                                  area=args.area, noise=noise, seed=args.seed,
                                  peak_disp=args.peak_disp)
    export_workbook(rows, outdir / "self_sensing_curves.xlsx", {
        "source": source,
        "loading rate (mm/s)": args.rate,
        "cross-section (mm^2)": args.area,
        "initial resistivity rho0 (ohm*m)": args.rho0,
        "stage noise (% of rho0)": str(noise),
        "peak displacement (mm)": round(peak_disp, 6),
        "note": "Electrical response is a schematic model, not measured data.",
    })
    if not args.no_figure:
        caption = ("注：依据文献受压自感知规律重绘的示意图，非实测数据；"
                   f"按 {args.rate:g} mm/s 等速位移换算时间；FCR=(ρ-ρ0)/ρ0×100%。")
        draw_panels(rows, peak_disp, caption, outdir / "self_sensing_panels.png")
    print(f"written to {outdir.resolve()}")


if __name__ == "__main__":
    main()
