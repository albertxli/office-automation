//! Step 6: Replace literal text tokens (e.g. `[country]` → `Japan`) across the deck.
//!
//! Scope: every shape on every slide, plus every slide master and custom layout
//! (reached through `Presentation.Designs`), recursing into groups and table cells.
//! Speaker notes are not touched.
//!
//! Mechanism (GOTCHA #46): `TextRange.Replace(FindWhat, ReplaceWhat, After, MatchCase,
//! WholeWords)` — it keeps the formatting of the replaced run and finds tokens that
//! PowerPoint has split across runs. It replaces ONE occurrence per call and returns the
//! replaced `TextRange`, or Nothing when there is no further match; Nothing arrives as a
//! null dispatch, so `as_dispatch()` failing is the normal loop exit. `After` must be
//! advanced past the replaced text, otherwise a replacement that contains its own token
//! is matched again forever.
//!
//! Matching is literal and case-sensitive. A token that is never found anywhere is
//! reported with `verbose::warn` (a typo in the token or in the deck).

use std::cell::RefCell;
use std::collections::HashMap;
use std::rc::Rc;

use console::Style;

use crate::com::dispatch::Dispatch;
use crate::com::variant::Variant;
use crate::error::OaResult;
use crate::office::constants::MsoShapeType;

type DispidCache = Rc<RefCell<HashMap<String, i32>>>;

/// msoTrue / msoFalse for the `MatchCase` / `WholeWords` arguments.
const MSO_TRUE: i32 = -1;
const MSO_FALSE: i32 = 0;

/// Shared state for one presentation walk.
struct Ctx<'a> {
    pairs: &'a [(String, String)],
    /// Replacements made per pair (same index as `pairs`).
    tallies: Vec<usize>,
    // One DISPID cache per COM class (GOTCHA #31)
    shape_cache: DispidCache,
    tf_cache: DispidCache,
    tr_cache: DispidCache,
    cell_cache: DispidCache,
    found_cache: DispidCache,
}

fn new_cache() -> DispidCache {
    Rc::new(RefCell::new(HashMap::new()))
}

/// Replace every `(find, value)` pair across slides, masters and layouts.
/// Returns the total number of replacements made.
pub fn replace_text(presentation: &mut Dispatch, pairs: &[(String, String)]) -> OaResult<usize> {
    if pairs.is_empty() {
        return Ok(0);
    }
    let mut ctx = Ctx {
        pairs,
        tallies: vec![0; pairs.len()],
        shape_cache: new_cache(),
        tf_cache: new_cache(),
        tr_cache: new_cache(),
        cell_cache: new_cache(),
        found_cache: new_cache(),
    };

    // --- Slides ---
    let mut slides = Dispatch::new(presentation.get("Slides")?.as_dispatch()?);
    let slide_count = slides.get("Count")?.as_i32()?;
    for i in 1..=slide_count {
        if let Ok(v) = slides.call("Item", &[Variant::from(i)])
            && let Ok(d) = v.as_dispatch()
        {
            let mut slide = Dispatch::new(d);
            walk_shapes(&mut slide, &format!("Slide {i:>2}"), &mut ctx);
        }
    }

    // --- Masters and their layouts (every design, for multi-master decks) ---
    if let Ok(dv) = presentation.get("Designs")
        && let Ok(dd) = dv.as_dispatch()
    {
        let mut designs = Dispatch::new(dd);
        let design_count = designs.get("Count").and_then(|v| v.as_i32()).unwrap_or(0);
        for d in 1..=design_count {
            let Ok(design_v) = designs.call("Item", &[Variant::from(d)]) else { continue };
            let Ok(design_d) = design_v.as_dispatch() else { continue };
            let mut design = Dispatch::new(design_d);
            let design_name = design.get("Name").and_then(|v| v.as_string()).unwrap_or_else(|_| d.to_string());

            let Ok(mv) = design.get("SlideMaster") else { continue };
            let Ok(md) = mv.as_dispatch() else { continue };
            let mut master = Dispatch::new(md);
            walk_shapes(&mut master, &format!("Master {design_name}"), &mut ctx);

            let Ok(lv) = master.get("CustomLayouts") else { continue };
            let Ok(ld) = lv.as_dispatch() else { continue };
            let mut layouts = Dispatch::new(ld);
            let layout_count = layouts.get("Count").and_then(|v| v.as_i32()).unwrap_or(0);
            for l in 1..=layout_count {
                let Ok(layout_v) = layouts.call("Item", &[Variant::from(l)]) else { continue };
                let Ok(layout_d) = layout_v.as_dispatch() else { continue };
                let mut layout = Dispatch::new(layout_d);
                let layout_name = layout.get("Name").and_then(|v| v.as_string()).unwrap_or_else(|_| l.to_string());
                walk_shapes(&mut layout, &format!("Layout {layout_name}"), &mut ctx);
            }
        }
    }

    // --- Tokens that matched nothing are almost always a typo ---
    for ((find, _), n) in pairs.iter().zip(&ctx.tallies) {
        if *n == 0 {
            super::verbose::warn(&format!("token {find:?} was not found anywhere in the deck"));
        }
    }

    Ok(ctx.tallies.iter().sum())
}

/// Walk `container.Shapes` — `container` is a Slide, SlideMaster or CustomLayout.
fn walk_shapes(container: &mut Dispatch, label: &str, ctx: &mut Ctx) {
    let Ok(sv) = container.get("Shapes") else { return };
    let Ok(sd) = sv.as_dispatch() else { return };
    let mut shapes = Dispatch::new(sd);
    let count = shapes.get("Count").and_then(|v| v.as_i32()).unwrap_or(0);
    for i in 1..=count {
        if let Ok(v) = shapes.call("Item", &[Variant::from(i)])
            && let Ok(d) = v.as_dispatch()
        {
            replace_in_shape(Dispatch::new_with_cache(d, ctx.shape_cache.clone()), label, ctx);
        }
    }
}

/// Groups recurse; tables visit every cell; anything else with a text frame is replaced.
/// OLE objects, pictures, charts and SmartArt have no text frame and are skipped.
fn replace_in_shape(mut shape: Dispatch, label: &str, ctx: &mut Ctx) {
    let name = shape.get("Name").and_then(|v| v.as_string()).unwrap_or_default();
    let shape_type = shape.get("Type").and_then(|v| v.as_i32()).unwrap_or(0);

    if shape_type == MsoShapeType::Group as i32 {
        if let Ok(gv) = shape.get("GroupItems")
            && let Ok(gd) = gv.as_dispatch()
        {
            let mut items = Dispatch::new(gd);
            let count = items.get("Count").and_then(|v| v.as_i32()).unwrap_or(0);
            for i in 1..=count {
                if let Ok(v) = items.call("Item", &[Variant::from(i)])
                    && let Ok(d) = v.as_dispatch()
                {
                    replace_in_shape(Dispatch::new_with_cache(d, ctx.shape_cache.clone()), label, ctx);
                }
            }
        }
        return;
    }

    // Tables first: a table shape has no text frame of its own, only its cells do.
    if shape.get("HasTable").and_then(|v| v.as_i32()).unwrap_or(0) != 0 {
        let Ok(tv) = shape.get("Table") else { return };
        let Ok(td) = tv.as_dispatch() else { return };
        let mut table = Dispatch::new(td);
        let rows = table.nav("Rows").and_then(|mut r| r.get("Count")).and_then(|v| v.as_i32()).unwrap_or(0);
        let cols = table.nav("Columns").and_then(|mut c| c.get("Count")).and_then(|v| v.as_i32()).unwrap_or(0);
        for r in 1..=rows {
            for c in 1..=cols {
                let Ok(cv) = table.call("Cell", &[Variant::from(r), Variant::from(c)]) else { continue };
                let Ok(cd) = cv.as_dispatch() else { continue };
                let mut cell = Dispatch::new_with_cache(cd, ctx.cell_cache.clone());
                let Ok(csv) = cell.get("Shape") else { continue };
                let Ok(csd) = csv.as_dispatch() else { continue };
                let mut cell_shape = Dispatch::new_with_cache(csd, ctx.shape_cache.clone());
                replace_in_text_frame(&mut cell_shape, label, &format!("{name}[{r},{c}]"), ctx);
            }
        }
        return;
    }

    if shape.get("HasTextFrame").and_then(|v| v.as_i32()).unwrap_or(0) != 0 {
        replace_in_text_frame(&mut shape, label, &name, ctx);
    }
}

/// Apply every pair to one shape's `TextFrame.TextRange`.
/// Reads `.Text` once so shapes without any token cost a single COM call.
fn replace_in_text_frame(shape: &mut Dispatch, label: &str, shape_name: &str, ctx: &mut Ctx) {
    let Ok(tfv) = shape.get("TextFrame") else { return };
    let Ok(tfd) = tfv.as_dispatch() else { return };
    let mut tf = Dispatch::new_with_cache(tfd, ctx.tf_cache.clone());
    let Ok(trv) = tf.get("TextRange") else { return };
    let Ok(trd) = trv.as_dispatch() else { return };
    let mut tr = Dispatch::new_with_cache(trd, ctx.tr_cache.clone());
    let Ok(text) = tr.get("Text").and_then(|v| v.as_string()) else { return };

    for (idx, (find, value)) in ctx.pairs.iter().enumerate() {
        if !text.contains(find.as_str()) {
            continue;
        }
        let n = replace_in_range(&mut tr, find, value, ctx.found_cache.clone());
        if n > 0 {
            ctx.tallies[idx] += n;
            log_hit(label, shape_name, find, value, n);
        }
    }
}

/// Replace all occurrences of `find` in one TextRange. See module docs (GOTCHA #46).
fn replace_in_range(tr: &mut Dispatch, find: &str, value: &str, found_cache: DispidCache) -> usize {
    let mut count = 0usize;
    let mut after = 0i32;
    loop {
        let Ok(fv) = tr.call("Replace", &[
            Variant::from(find),
            Variant::from(value),
            Variant::from(after),
            Variant::from(MSO_TRUE),   // MatchCase
            Variant::from(MSO_FALSE),  // WholeWords
        ]) else { break };
        // Nothing (no further match) is a null dispatch — normal exit
        let Ok(fd) = fv.as_dispatch() else { break };
        let mut found = Dispatch::new_with_cache(fd, found_cache.clone());
        count += 1;

        let start = found.get("Start").and_then(|v| v.as_i32()).unwrap_or(0);
        let len = found.get("Length").and_then(|v| v.as_i32()).unwrap_or(0);
        // Continue searching after the replaced text (never move backwards)
        after = after.max(start + len - 1);

        if count >= 100_000 {
            break; // paranoia guard — should be unreachable
        }
    }
    count
}

/// Verbose line: `      Slide  3 │ TextBox 8                [country] → Japan (1)`
fn log_hit(label: &str, shape_name: &str, find: &str, value: &str, n: usize) {
    if !super::verbose::is_verbose() {
        return;
    }
    let s_dim = Style::new().dim();
    let s_val = Style::new().white().bold();
    println!("      {} {} {:<24} {} {} {} {}",
        s_dim.apply_to(label),
        s_dim.apply_to("│"),
        s_dim.apply_to(shape_name),
        s_dim.apply_to(find),
        s_dim.apply_to("→"),
        s_val.apply_to(if value.is_empty() { "(removed)" } else { value }),
        s_dim.apply_to(format!("({n})")));
}
