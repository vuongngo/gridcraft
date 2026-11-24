// Runs the wasm-bindgen build under Node: `PO_WASM_DIR=<out-dir> node --test crates/po-wasm/tests/`.
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import path from 'node:path';
import { test } from 'node:test';
import { pathToFileURL } from 'node:url';

const dir = path.resolve(process.env.PO_WASM_DIR ?? 'pkg');
const wasm = await import(pathToFileURL(path.join(dir, 'po_wasm.mjs')).href);
const { memory } = await wasm.default({ module_or_path: readFileSync(path.join(dir, 'po_wasm_bg.wasm')) });
const fixture = readFileSync(new URL('fixtures/spike.xlsx', import.meta.url));
const view = (book, sheet) => JSON.parse(book.view(sheet, 0, 20, 10));

test('a plain workbook round-trips cell-identical', () => {
  const book = new wasm.Workbook(fixture);
  const reopened = new wasm.Workbook(book.save());
  for (const sheet of ['Open', 'Locked']) assert.deepEqual(view(reopened, sheet), view(book, sheet));
});

test('edits, recalculates, inserts, undoes and keeps the edit after save and reopen', () => {
  const book = new wasm.Workbook(fixture);
  assert.equal(view(book, 'Open').sheet.rows[0][1].input, '=A1*2');
  book.set_cell('Open', 'A1', '7');
  assert.equal(view(book, 'Open').sheet.rows[0][1].text, '14');
  book.insert_rows('Open', '1:1');
  assert.equal(view(book, 'Open').sheet.rows[1][1].input, '=A2*2');
  book.undo();
  const reopened = new wasm.Workbook(book.save());
  assert.equal(view(reopened, 'Open').sheet.rows[0][0].text, '7');
  assert.equal(view(reopened, 'Open').sheet.rows[0][1].text, '14');
});

test('a protected sheet refuses every operation', () => {
  const book = new wasm.Workbook(fixture);
  assert.equal(view(book, 'Locked').sheet.protected, true);
  for (const attempt of [
    () => book.set_cell('Locked', 'A1', '9'),
    () => book.insert_rows('Locked', '1:1'),
    () => book.delete_rows('Locked', '1:1'),
    () => book.insert_columns('Locked', 'A:A'),
    () => book.delete_columns('Locked', 'A:A'),
    () => book.column_width('Locked', 'A:A', 90),
    () => book.row_height('Locked', '1:1', 30),
  ])
    assert.throws(attempt, /Protected sheet is read-only/);
  assert.equal(view(new wasm.Workbook(book.save()), 'Locked').sheet.protected, true);
});

test('linear memory stops at the 1 GiB link-time maximum', () => {
  assert.ok(memory instanceof WebAssembly.Memory);
  assert.throws(() => memory.grow(16_385));
});
