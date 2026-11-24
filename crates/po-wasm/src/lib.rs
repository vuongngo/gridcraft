//! Headless GridCraft for Po's Author spreadsheet canvas, published as `@agimon-ai/gridcraft-wasm`.
//! Every edit refuses protected sheets: GridCraft itself only guards locked cells.
use gridcraft_engine::core::{CellRef, Value};
use gridcraft_engine::{DocState, Session, Trust, display, io};
use serde_json::{Value as Json, json};
use wasm_bindgen::prelude::*;

#[wasm_bindgen]
pub struct Workbook(Session);

#[wasm_bindgen]
impl Workbook {
    #[wasm_bindgen(constructor)]
    pub fn open(bytes: &[u8]) -> Result<Workbook, JsError> {
        let (book, _) = io::open_bytes("book.xlsx", bytes).map_err(error)?;
        let mut session = Session::new();
        session.add_document(DocState::new(book, None, "book.xlsx".into(), Trust::FromFile));
        Ok(Self(session))
    }

    /// `{ sheets: [{name, protected}], sheet: {name, protected, widths, rows} }` for one sheet,
    /// covering its used range plus `pad` empty lines, at most `max_rows` x `max_cols`.
    /// An unknown sheet, such as one an undo removed, shows the first sheet.
    pub fn view(&self, sheet: &str, pad: u32, max_rows: u32, max_cols: u32) -> Result<String, JsError> {
        let doc = self.0.doc().map_err(error)?;
        let wb = &doc.wb;
        let index = wb.sheet_index(sheet).unwrap_or(0);
        let sh = wb.sheet(index).ok_or_else(|| JsError::new("Unknown sheet"))?;
        let used = sh.used_range();
        let rows = used.map_or(0, |r| r.end.row + 1).saturating_add(pad).min(max_rows);
        let cols = used.map_or(0, |r| r.end.col + 1).saturating_add(pad).min(max_cols).max(1);
        let grid: Vec<Json> = (0..rows)
            .map(|row| {
                Json::Array(
                    (0..cols)
                        .map(|col| {
                            let c = CellRef::new(row, col);
                            let value = sh.value(c);
                            if matches!(value, Value::Empty) && sh.cell(c).and_then(|x| x.formula.as_ref()).is_none() {
                                return Json::Null;
                            }
                            json!({
                                "text": display::cell_text(wb, sh, c),
                                "input": sh.input_text(c),
                                "numeric": matches!(value, Value::Number(_)),
                                "bold": wb.styles.get(sh.style_id(c)).font.bold,
                            })
                        })
                        .collect(),
                )
            })
            .collect();
        Ok(json!({
            "sheets": wb.sheets.iter().map(|s| json!({"name": s.name, "protected": s.is_protected()})).collect::<Vec<_>>(),
            "sheet": {
                "name": sh.name,
                "protected": sh.is_protected(),
                "widths": (0..cols).map(|col| sh.col_width(col).round()).collect::<Vec<_>>(),
                "heights": (0..rows).map(|row| sh.row_height(row).round()).collect::<Vec<_>>(),
                "rows": grid,
            },
        })
        .to_string())
    }

    pub fn set_cell(&mut self, sheet: &str, cell: &str, input: &str) -> Result<(), JsError> {
        self.editable(sheet)?;
        self.run("cell.set", json!({"sheet": sheet, "cell": cell, "input": input}))
    }

    /// `lines` is `"3:3"` for rows or `"C:C"` for columns, on `sheet`.
    pub fn insert_rows(&mut self, sheet: &str, lines: &str) -> Result<(), JsError> {
        self.on(sheet)?;
        self.run("home.insertRows", json!({"rows": lines}))
    }

    pub fn delete_rows(&mut self, sheet: &str, lines: &str) -> Result<(), JsError> {
        self.on(sheet)?;
        self.run("home.deleteRows", json!({"rows": lines}))
    }

    pub fn insert_columns(&mut self, sheet: &str, lines: &str) -> Result<(), JsError> {
        self.on(sheet)?;
        self.run("home.insertColumns", json!({"cols": lines}))
    }

    pub fn delete_columns(&mut self, sheet: &str, lines: &str) -> Result<(), JsError> {
        self.on(sheet)?;
        self.run("home.deleteColumns", json!({"cols": lines}))
    }

    /// Pixel width for `cols` such as `"B:D"`; 0 hides them.
    pub fn column_width(&mut self, sheet: &str, cols: &str, width: f64) -> Result<(), JsError> {
        self.editable(sheet)?;
        self.run("home.columnWidth", json!({"sheet": sheet, "cols": cols, "width": width}))
    }

    /// Pixel height for `rows` such as `"3:5"`; 0 hides them.
    pub fn row_height(&mut self, sheet: &str, rows: &str, height: f64) -> Result<(), JsError> {
        self.editable(sheet)?;
        self.run("home.rowHeight", json!({"sheet": sheet, "rows": rows, "height": height}))
    }
    /// Adds a sheet after the last one. GridCraft refuses it when the workbook structure is locked.
    pub fn add_sheet(&mut self, name: &str) -> Result<(), JsError> {
        let count = self.0.doc().map_err(error)?.wb.sheets.len();
        self.run("home.insertSheet", json!({"name": name, "before": count}))
    }

    pub fn undo(&mut self) -> Result<(), JsError> {
        self.run("edit.undo", json!({}))
    }

    pub fn save(&self) -> Result<Vec<u8>, JsError> {
        let doc = self.0.doc().map_err(error)?;
        io::save_bytes(&doc.wb, "book.xlsx").map_err(error)
    }
}

impl Workbook {
    fn editable(&self, sheet: &str) -> Result<(), JsError> {
        let doc = self.0.doc().map_err(error)?;
        let index = doc.wb.sheet_index(sheet).ok_or_else(|| JsError::new("Unknown sheet"))?;
        if doc.wb.sheet(index).is_some_and(|s| s.is_protected()) {
            return Err(JsError::new("Protected sheet is read-only"));
        }
        Ok(())
    }

    /// Structural commands act on the active sheet.
    fn on(&mut self, sheet: &str) -> Result<(), JsError> {
        self.editable(sheet)?;
        self.run("sheet.activate", json!({"sheet": sheet}))
    }

    fn run(&mut self, command: &str, args: Json) -> Result<(), JsError> {
        self.0.execute_headless(command, args).map(|_| ()).map_err(error)
    }
}

fn error(err: impl std::fmt::Display) -> JsError {
    JsError::new(&err.to_string())
}
