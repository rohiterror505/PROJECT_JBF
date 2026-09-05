#!/usr/bin/env python3
"""JAI BHADRA FOUNDATION - Lucky Draw Coupon Manager (Web UI).

A small Flask app that reuses rohit.py for ALL business logic
(coupon rendering, Excel tracking, validation).  The browser handles
the UI, so there is no desktop GUI thread to freeze.

RUN:
    python server.py
Then open http://127.0.0.1:5000 in your browser.
"""

import io
import base64
import traceback
from pathlib import Path

from flask import Flask, jsonify, request, send_file, render_template, send_from_directory, session, redirect, Response

import rohit

app = Flask(__name__)
app.config["TEMPLATES_AUTO_RELOAD"] = True
app.jinja_env.auto_reload = True

# Session secret — change this to any long random string for production.
app.secret_key = "jbf-lucky-draw-secret-key-change-me-2024"

# Hardcoded admin credentials. Change these to suit your deployment.
ADMIN_USER = "admin"
ADMIN_PASS = "12qwasxz"

JOY_USER = "joy"
JOY_PASS = "joy@12qwasxz"


# ------------------------------------------------------------------
# Helpers
# ------------------------------------------------------------------

def _rows_to_dicts(rows):
    """Turn the tuple rows from list_sales into JSON-friendly dicts."""
    out = []
    for r in rows:
        if len(r) >= 11:
            sno, name, phone, address, start, end, qty, date, stype, set_sz, amount = r[:11]
        elif len(r) >= 9:
            sno, name, phone, address, start, end, qty, date, stype = r[:9]
            set_sz = qty if stype == "PHYSICAL" else None
            amount = (rohit.price_for_set_size(qty) if stype == "PHYSICAL"
                      else qty * rohit.PRICE_PER_COUPON)
        else:
            sno, name, phone, address, start, end, qty, date = r
            stype = "SALE"
            set_sz = None
            amount = qty * rohit.PRICE_PER_COUPON
        out.append({
            "sno": sno,
            "name": name or "",
            "phone": phone or "",
            "address": address or "",
            "start": start,
            "end": end,
            "qty": qty,
            "date": date or "",
            "type": stype or "SALE",
            "set_size": set_sz,
            "amount": amount,
        })
    return out


def _ok(payload=None):
    resp = {"success": True}
    if payload is not None:
        resp.update(payload)
    return jsonify(resp)


def _err(msg, code=400):
    return jsonify({"success": False, "error": str(msg)}), code


# ------------------------------------------------------------------
# Page
# ------------------------------------------------------------------

@app.route("/")
def index():
    # Serve the HTML directly so edits are picked up immediately without
    # depending on Jinja's bytecode cache.
    import os
    p = os.path.join(os.path.dirname(os.path.abspath(__file__)), "templates", "index.html")
    with open(p, "r", encoding="utf-8") as f:
        html = f.read()
    resp = Response(html, mimetype="text/html")
    resp.headers["Cache-Control"] = "no-cache, no-store, must-revalidate"
    resp.headers["Pragma"] = "no-cache"
    resp.headers["Expires"] = "0"
    return resp


@app.route("/rules")
def rules_page():
    """Public Lucky Draw rules + prizes page (no login required) so
    contestants can view the rules and print them."""
    import os
    p = os.path.join(os.path.dirname(os.path.abspath(__file__)), "templates", "rules.html")
    with open(p, "r", encoding="utf-8") as f:
        html = f.read()
    resp = Response(html, mimetype="text/html")
    resp.headers["Cache-Control"] = "no-cache, no-store, must-revalidate"
    resp.headers["Pragma"] = "no-cache"
    resp.headers["Expires"] = "0"
    return resp


# ------------------------------------------------------------------
# Auth
# ------------------------------------------------------------------

PUBLIC_PATHS = {"/login", "/api/login", "/logout", "/rules"}

@app.before_request
def _require_login():
    """Gate every page and API call behind a logged-in session, except
    the login / logout endpoints themselves."""
    path = request.path
    if path in PUBLIC_PATHS:
        return None
    if session.get("user"):
        return None
    # Static-ish asset paths (if any) — allow through.
    if path.startswith("/static/"):
        return None
    # API calls get a 401 JSON so the front-end can react; everything
    # else is redirected to the login page.
    if path.startswith("/api/"):
        return jsonify({"success": False, "error": "Authentication required. Please log in."}), 401
    return redirect("/login")


@app.route("/login", methods=["GET"])
def login_page():
    if session.get("user"):
        return redirect("/")
    import os
    p = os.path.join(os.path.dirname(os.path.abspath(__file__)), "templates", "login.html")
    with open(p, "r", encoding="utf-8") as f:
        html = f.read()
    resp = Response(html, mimetype="text/html")
    resp.headers["Cache-Control"] = "no-cache, no-store, must-revalidate"
    resp.headers["Pragma"] = "no-cache"
    resp.headers["Expires"] = "0"
    return resp


@app.route("/api/login", methods=["POST"])
def api_login():
    data = request.get_json(silent=True) or {}
    user = (data.get("username") or "").strip()
    pw = data.get("password") or ""
    if not user or not pw:
        return _err("Please enter both username and password.")
    if (user == ADMIN_USER and pw == ADMIN_PASS) or (user == JOY_USER and pw == JOY_PASS):
        session["user"] = user
        session.permanent = True
        return _ok({"redirect": "/", "message": "Login successful."})
    return _err("Invalid username or password.", 401)


@app.route("/logout", methods=["GET", "POST"])
def logout():
    session.clear()
    return redirect("/login")


# ------------------------------------------------------------------
# API
# ------------------------------------------------------------------

@app.route("/api/state")
def api_state():
    """One call that returns everything the dashboard / sales / gaps tabs
    need, so the browser can refresh all panels from a single request."""
    try:
        rohit.ensure_sales_file()
        rows = rohit.list_sales(print_it=False)
    except RuntimeError as exc:
        return _err(str(exc), 500)
    except Exception as exc:
        return _err(str(exc), 500)

    sales = _rows_to_dicts(rows)

    # Derive sold ranges from the rows we already fetched instead of making
    # a second DB round-trip via get_sold_ranges().
    ranges = [(s["start"], s["end"]) for s in sales
              if s["start"] is not None and s["end"] is not None]

    # Summary stats (what the dashboard cards show).
    sale_count = sum(1 for s in sales if s["type"] == "SALE")
    phys_count = sum(1 for s in sales if s["type"] == "PHYSICAL")
    total_qty = sum((s["qty"] or 0) for s in sales)
    # Revenue is the sum of each row's stored donation amount (set-based
    # pricing) rather than a flat qty * 100.
    revenue = sum((s["amount"] or 0) for s in sales)

    assigned = set()
    for s, e in ranges:
        for n in range(s, e + 1):
            assigned.add(n)
    gaps = 0
    gap_ranges = []
    if assigned:
        low, high = min(assigned), max(assigned)
        missing = [n for n in range(low, high + 1) if n not in assigned]
        gaps = len(missing)
        # collapse into runs
        if missing:
            run_s = run_e = missing[0]
            for n in missing[1:]:
                if n == run_e + 1:
                    run_e = n
                else:
                    gap_ranges.append((run_s, run_e))
                    run_s = run_e = n
            gap_ranges.append((run_s, run_e))
        low_str = f"{low:04d}"
        high_str = f"{high:04d}"
    else:
        low_str = high_str = "-"

    last_label = "-"
    if sales:
        s = sales[-1]
        last_label = (
            f"{s['start']:04d}" if s["start"] == s["end"]
            else f"{s['start']:04d}-{s['end']:04d}"
        )

    return _ok({
        "sales": sales,
        "max_coupon": rohit.MAX_COUPON,
        "draw_line": rohit.DRAW_LINE,
        "website": rohit.WEBSITE,
        "summary": {
            "total_sold": total_qty,
            "revenue": revenue,
            "remaining": rohit.MAX_COUPON - total_qty,
            "gaps": gaps,
            "sale_rows": sale_count,
            "phys_rows": phys_count,
            "last_coupon": last_label,
            "low": low_str,
            "high": high_str,
            "span": (high - low + 1) if assigned else 0,
            "assigned": len(assigned),
        },
        "gap_ranges": [
            {"start": s, "end": e, "count": e - s + 1} for s, e in gap_ranges
        ],
    })


@app.route("/api/preview")
def api_preview():
    """Render a coupon preview PNG in memory and return it as base64 so the
    browser can show it in an <img>.  Nothing is written to disk."""
    try:
        start = int(request.args.get("start", 1))
        end = int(request.args.get("end", start))
        name = request.args.get("name") or None
        phone = request.args.get("phone") or None
        address = request.args.get("address") or None
        set_size = request.args.get("set_size", type=int)
    except (TypeError, ValueError):
        return _err("Invalid parameters")

    # Decide the printed donation amount for the preview.
    if set_size is not None and set_size > 0:
        amount = rohit.price_for_set_size(set_size)
    else:
        amount = (end - start + 1) * rohit.PRICE_PER_COUPON

    try:
        img = rohit._render_coupon(
            start, end, buyer=name, phone=phone, address=address, amount=amount
        ).convert("RGB")
        buf = io.BytesIO()
        img.save(buf, format="PNG")
        b64 = base64.b64encode(buf.getvalue()).decode("ascii")
        return _ok({"image": "data:image/png;base64," + b64})
    except Exception as exc:
        return _err(str(exc), 500)


@app.route("/api/sample")
def api_sample():
    """Render a sample coupon PNG in memory and return it as base64.
    Nothing is written to disk or the Excel tracker — this is purely for
    showing someone what a coupon looks like."""
    try:
        num = int(request.args.get("num", 0))
    except (TypeError, ValueError):
        return _err("Invalid coupon number")
    try:
        img = rohit._render_coupon(
            num, num, amount=rohit.PRICE_PER_COUPON, sample=True
        ).convert("RGB")
        buf = io.BytesIO()
        img.save(buf, format="PNG")
        b64 = base64.b64encode(buf.getvalue()).decode("ascii")
        return _ok({"image": "data:image/png;base64," + b64})
    except Exception as exc:
        return _err(str(exc), 500)


@app.route("/api/sale", methods=["POST"])
def api_sale():
    """Record a normal sale.  In Excel mode a PNG is saved to disk; in
    Postgres (cloud) mode the coupon is rendered on demand later, so no
    file is written here."""
    data = request.get_json(force=True)
    try:
        name = (data.get("name") or "").strip()
        phone = (data.get("phone") or "").strip()
        address = (data.get("address") or "").strip()
        start = int(data.get("start"))
        qty = int(data.get("qty", 1))
    except (TypeError, ValueError):
        return _err("Invalid parameters")

    if not name:
        return _err("Buyer name is required.")
    end = start + qty - 1
    if end > rohit.MAX_COUPON:
        return _err(f"End {end:04d} exceeds max {rohit.MAX_COUPON}.")

    amount = qty * rohit.PRICE_PER_COUPON
    try:
        overlap = rohit.is_already_sold(start, end)
        if overlap:
            return _err(
                f"Coupons {overlap[0]:04d}-{overlap[1]:04d} are already sold."
            )
        filename = None
        if not rohit._USE_POSTGRES:
            filename = rohit.create_coupon(
                start, end, buyer=name or None,
                phone=phone or None, address=address or None, amount=amount,
            )
        rohit.record_sale(name, phone, address, start, end, amount=amount)
        msg = (f"Sale complete! Coupons {start:04d}-{end:04d} recorded."
               + (f" Saved as {filename}." if filename else ""))
        return _ok({"message": msg, "filename": filename})
    except RuntimeError as exc:
        return _err(str(exc), 500)
    except Exception as exc:
        return _err(str(exc), 500)


@app.route("/api/physical", methods=["POST"])
def api_physical():
    """Generate physical coupon sets (no buyer)."""
    data = request.get_json(force=True)
    try:
        set_size = int(data.get("set_size", 10))
        start = int(data.get("start"))
        num_sets = int(data.get("num_sets", 1))
    except (TypeError, ValueError):
        return _err("Invalid parameters")

    end_block = start + num_sets * set_size - 1
    if end_block > rohit.MAX_COUPON:
        return _err(f"Would end at {end_block:04d}, exceeds {rohit.MAX_COUPON}.")
    if set_size < 1 or num_sets < 1:
        return _err("Set size and number of sets must be >= 1.")

    set_amount = rohit.price_for_set_size(set_size)
    try:
        overlap = rohit.is_already_sold(start, end_block)
        if overlap:
            return _err(
                f"Coupons {overlap[0]:04d}-{overlap[1]:04d} already sold."
            )
        generated = 0
        for i in range(num_sets):
            s = start + i * set_size
            e = s + set_size - 1
            try:
                if not rohit._USE_POSTGRES:
                    rohit.create_coupon(s, e, amount=set_amount)
                rohit.record_sale(
                    None, None, None, s, e, sale_type="PHYSICAL",
                    set_size=set_size, amount=set_amount,
                )
                generated += 1
            except Exception:
                pass
        return _ok({
            "message": f"Done. {generated}/{num_sets} set(s) generated. "
                        f"Numbers {start:04d}-{end_block:04d} locked.",
            "generated": generated,
        })
    except RuntimeError as exc:
        return _err(str(exc), 500)
    except Exception as exc:
        return _err(str(exc), 500)


@app.route("/api/sale/<int:sno>", methods=["DELETE", "PUT"])
def api_sale_by_sno(sno):
    """Handle a single sale row by S.No.
    DELETE -> remove the sale (and optionally its PNG via ?png=1).
    PUT    -> update the buyer details (Name, Phone, Address) only; coupon
              numbers, qty, amount and type stay locked.  In Excel mode the
              coupon PNG is re-rendered with the new info so the QR stays correct.
    """
    if request.method == "DELETE":
        del_png = request.args.get("png", "0") in ("1", "true", "yes")
        try:
            ok = rohit.delete_sale(sno, delete_png=del_png)
            if ok:
                return _ok({"message": f"Sale #{sno} deleted."})
            return _err(f"No sale #{sno}.", 404)
        except RuntimeError as exc:
            return _err(str(exc), 500)

    # PUT
    data = request.get_json(force=True)
    try:
        name = (data.get("name") or "").strip()
        phone = (data.get("phone") or "").strip()
        address = (data.get("address") or "").strip()
    except (TypeError, ValueError):
        return _err("Invalid parameters")
    try:
        ok = rohit.update_sale(sno, name, phone, address, regen_png=True)
        if ok:
            return _ok({"message": f"Sale #{sno} updated."})
        return _err(f"No sale #{sno}.", 404)
    except RuntimeError as exc:
        return _err(str(exc), 500)
    except Exception as exc:
        return _err(str(exc), 500)


@app.route("/api/delete/physical/range", methods=["POST"])
def api_delete_physical_range():
    data = request.get_json(force=True)
    try:
        s = int(data.get("start"))
        e = int(data.get("end"))
    except (TypeError, ValueError):
        return _err("Invalid parameters")
    del_pngs = bool(data.get("delete_pngs", True))
    try:
        count = rohit.delete_physical_by_range(s, e, delete_pngs=del_pngs)
        return _ok({"message": f"Deleted {count} physical set(s) in {s:04d}-{e:04d}.",
                     "count": count})
    except RuntimeError as exc:
        return _err(str(exc), 500)


@app.route("/api/delete/physical/all", methods=["POST"])
def api_delete_physical_all():
    data = request.get_json(force=True) if request.data else {}
    del_pngs = bool(data.get("delete_pngs", True))
    try:
        count = rohit.delete_all_physical(delete_pngs=del_pngs)
        return _ok({"message": f"Deleted {count} physical set(s).", "count": count})
    except RuntimeError as exc:
        return _err(str(exc), 500)


@app.route("/api/delete/all", methods=["POST"])
def api_delete_all_sales():
    data = request.get_json(force=True) if request.data else {}
    del_pngs = bool(data.get("delete_pngs", False))
    try:
        count = rohit.delete_all_sales(delete_pngs=del_pngs)
        return _ok({"message": f"Deleted all {count} sale(s).", "count": count})
    except RuntimeError as exc:
        return _err(str(exc), 500)


@app.route("/api/coupon-image/<path:filename>")
def api_coupon_image(filename):
    """Serve a generated coupon PNG from the output folder (Excel mode)."""
    path = rohit.OUTPUT_DIR / filename
    if not path.exists():
        return _err("File not found", 404)
    return send_file(str(path), mimetype="image/png")


@app.route("/api/coupon-render/<int:sno>")
def api_coupon_render(sno):
    """Render a coupon PNG on demand from the sale row's data (works in
    both Excel and Postgres mode — no disk file needed).  Used by the web
    UI's Sales 'View' button so coupons display even on cloud hosting
    where no PNGs are saved to disk."""
    try:
        rows = rohit.list_sales(print_it=False)
    except Exception as exc:
        return _err(str(exc), 500)
    row = None
    for r in rows:
        try:
            if int(r[0]) == sno:
                row = r
                break
        except (ValueError, TypeError):
            continue
    if row is None:
        return _err("Sale not found", 404)
    try:
        start = int(row[4])
        end = int(row[5])
        name = row[1] if len(row) > 1 and row[1] else None
        phone = row[2] if len(row) > 2 and row[2] else None
        address = row[3] if len(row) > 3 and row[3] else None
        stype = row[8] if len(row) >= 9 and row[8] else "SALE"
        set_sz = row[9] if len(row) >= 10 and row[9] else None
        amount = row[10] if len(row) >= 11 and row[10] else None
        if amount is None:
            qty = end - start + 1
            amount = (rohit.price_for_set_size(qty) if stype == "PHYSICAL"
                      else qty * rohit.PRICE_PER_COUPON)
        img = rohit._render_coupon(
            start, end, buyer=name, phone=phone, address=address, amount=amount
        ).convert("RGB")
        buf = io.BytesIO()
        img.save(buf, format="PNG")
        buf.seek(0)
        return send_file(buf, mimetype="image/png")
    except Exception as exc:
        return _err(str(exc), 500)


# ------------------------------------------------------------------
# Lucky Draw
# ------------------------------------------------------------------

@app.route("/api/draw", methods=["POST"])
def api_draw():
    """Conduct the Lucky Draw: pick 50 distinct winners from ALL 9999
    coupons (1-MAX_COUPON, sold or unsold — 1-in-9999 chance each) and
    save the results.  Unsold winning coupons are marked buyer="UNSOLD".
    Returns the winners list."""
    try:
        if rohit.has_draw_results():
            return _err("Lucky Draw results already exist. Clear them first.", 409)
    except RuntimeError as exc:
        return _err(str(exc), 500)

    try:
        results = rohit.draw_winners()
    except RuntimeError as exc:
        return _err(str(exc), 500)
    except Exception as exc:
        return _err(str(exc), 500)

    return _ok({
        "message": f"Draw complete! 50 winners selected from all {rohit.MAX_COUPON} coupons (40 consolation + 10 main). Unsold winning coupons are marked UNSOLD.",
        "results": results,
    })


@app.route("/api/draw/results")
def api_draw_results():
    """Return the saved draw results, or null if no draw has been conducted."""
    try:
        results = rohit.get_draw_results()
    except RuntimeError as exc:
        return _err(str(exc), 500)
    return _ok({"results": results})


@app.route("/api/draw/clear", methods=["POST"])
def api_draw_clear():
    """Clear saved draw results so a new draw can be conducted."""
    try:
        cleared = rohit.clear_draw_results()
    except RuntimeError as exc:
        return _err(str(exc), 500)
    return _ok({
        "message": "Lucky Draw results cleared." if cleared
                   else "No results to clear.",
        "cleared": cleared,
    })


if __name__ == "__main__":
    print("=" * 60)
    print("JAI BHADRA FOUNDATION - Lucky Draw Coupon Manager (Web)")
    print("=" * 60)
    print(f"Open http://127.0.0.1:5000 (or http://<this-PC-LAN-IP>:5000) in your browser")
    print(f"Output folder: {rohit.OUTPUT_DIR}")
    print(f"Sales file  : {rohit.SALES_FILE}")
    print("=" * 60)
    app.run(host="0.0.0.0", port=5000, debug=False)