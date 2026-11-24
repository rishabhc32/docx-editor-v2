// The document's settings writes (design D14): Protect Document, even and odd pages, the
// document's properties and its notes' properties, replayed on the package in the room's order.

import { withDocumentProperties } from '../../../store/package/document-property-writes.ts';
import { applyHeaderFooterLifecycleOp } from '../../../store/package/hf-lifecycle.ts';
import { writeDocumentNoteProperties } from '../../../store/package/note-lifecycle-props.ts';
import type { OoxmlPackage } from '../../../store/package/ooxml-package.ts';
import { ooxmlTreesEqual } from '../../../store/package/ooxml-serialize.ts';
import { loadedPackageOf } from '../load/load.ts';
import { statusOf } from '../model/document.ts';
import type { OrderedSettingResource } from '../model/resources.ts';
import type { OrderedDocument } from '../model/types.ts';

/** One settings write on the package; the package stays when the engine refuses it. */
function withSetting(pkg: OoxmlPackage, op: OrderedSettingResource['op']): OoxmlPackage {
  if (op.op === 'setNoteProperties') {
    return writeDocumentNoteProperties(pkg, { ...op, scope: 'document' }) ?? pkg;
  }

  if (op.op === 'setDocumentProperties') {
    return withDocumentProperties(pkg, op.values) ?? pkg;
  }

  // The engine's own settings write (Protect Document; even and odd pages). It answers a write
  // that changes nothing with a refusal
  const applied = applyHeaderFooterLifecycleOp(pkg, op);
  return applied.ok ? applied.package : pkg;
}

/** The settings writes of accepted suggestions, replayed in the order changes made them. */
export function withSettings(pkg: OoxmlPackage, doc: OrderedDocument): OoxmlPackage {
  let next = pkg;
  for (const resource of doc.resources.added.values()) {
    if (resource.kind === 'setting' && statusOf(doc, resource.suggestion) === 'accepted') {
      next = withSetting(next, resource.op);
    }
  }

  return next;
}

/**
 * Whether a settings write leaves every part as the loaded file under the accepted writes shows
 * it, so it changes nothing and makes no record. A document loaded from no package has no
 * settings to compare with: its write is made.
 */
export function settingKeeps(doc: OrderedDocument, op: OrderedSettingResource['op']): boolean {
  const loaded = loadedPackageOf(doc);
  if (loaded === null) {
    return false;
  }

  const shown = withSettings(loaded, doc);
  const next = withSetting(shown, op);
  return [...next.parts].every(([name, part]) => {
    const was = shown.parts.get(name);
    return was !== undefined && (was === part || ooxmlTreesEqual(was, part));
  });
}
