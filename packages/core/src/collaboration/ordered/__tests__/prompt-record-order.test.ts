// A control's value record keeps `w:sdtPr` in schema order (the random paste session, seed 71).
// Ana gave a control that shows its prompt three values in turn, and the third was rejected:
// the prompt shows again while her second value's record is pending. That record was made while
// her first value filled the control, so it holds the `w:placeholder` the engine names then, and
// it put `w:showingPlcHdr` before it. The saved file broke the `CT_SdtPr` sequence, and the
// reopened copy kept that order after a reject.

import { describe, expect, test } from 'bun:test';
import { serializeOoxmlPart } from '../../../store/package/ooxml-tree.ts';
import { composeOrderedPart } from '../compose/compose.ts';
import { loadOrderedPackage } from '../load/load.ts';
import { setStatus } from '../model/mutations.ts';
import { withUndoRecord } from '../model/undo-record.ts';
import { addressOf, apply, BODY, EDIT, opens } from './change-apply-fixture.ts';
import { checkSaveAndLoad } from './sessions/session.ts';
import { bodyOf, docx, MATRIX_ITEMS, packageOf } from './word-fixtures.ts';

/** The prompt control's document, after Ana's three values and the third's rejection. */
function promptShownUnderPendingValue() {
  const item = MATRIX_ITEMS.find((candidate) => candidate.key === 'placeholder');
  if (!item) throw new Error('no placeholder item');
  const base = packageOf(docx(bodyOf(item)));
  const doc = loadOrderedPackage(base);
  const value = (id: string, text: string) => {
    const control = opens(doc, 'sdt')[0];
    if (!control) throw new Error('no control');
    return apply(doc, id, {
      kind: 'control-value',
      ...EDIT,
      control: addressOf(control),
      value: { kind: 'text', text },
    }).suggestions;
  };
  value('ana-1:1', 'Party b');
  value('ana-1:2', 'Party c');
  const [third] = value('ana-1:3', 'Party a');
  if (third === undefined) throw new Error('no third value');
  withUndoRecord(doc, () => setStatus(doc, third, 'rejected'));
  return { doc, base };
}

describe("a control's value record", () => {
  test('writes the prompt flag after the placeholder, in schema order', () => {
    const { doc } = promptShownUnderPendingValue();
    const xml = serializeOoxmlPart(composeOrderedPart(doc, BODY).part);
    const sdtPr = /<w:sdtPr>.*?<\/w:sdtPr>/.exec(xml)?.[0] ?? '';
    expect(sdtPr).toContain('<w:showingPlcHdr/>');
    expect(sdtPr.indexOf('<w:placeholder>')).toBeGreaterThanOrEqual(0);
    expect(sdtPr.indexOf('<w:placeholder>')).toBeLessThan(sdtPr.indexOf('<w:showingPlcHdr/>'));
  });

  test('saves and reopens to the same view and decisions', () => {
    const { doc, base } = promptShownUnderPendingValue();
    checkSaveAndLoad(doc, 'prompt under a pending value', base);
  });
});
