/**
 * `@docx-editor.dev/core/collaboration/ordered`: the experimental ordered suggestion model.
 *
 * One ordered list of items per story part, held in a summed B+ tree. A suggestion is a label
 * on items and a decision only changes a suggestion's status, so every item keeps its address
 * for the whole room session and any set of decisions gives one document in every order.
 *
 * This entry holds the tree, the item and suggestion types, the undo record, Word load and Word
 * compose; the edits (by item address, placed by the engine's rules), the review cards and the
 * decisions; the change format (version 1) with its validator, id minting and apply; the protocol's
 * messages and codec; the room core with the suggestions-only policy; and the client: the
 * editor's collaboration session over an ordered room, with an in-process room for tests and a
 * WebSocket transport over the host's own socket. It is EXPERIMENTAL: names and shapes can change
 * in a minor release until ordered collaboration ships.
 *
 * @packageDocumentation
 * @public
 */

export {
  SumTree,
  SumTreeInner,
  SumTreeLeaf,
  INNER_MAX,
  LEAF_MAX,
  type SumTreeItem,
  type SumTreeNode,
  type SumTreeSeek,
  type SumTreeStats,
  type SummaryOps,
} from './tree/sum-tree.ts';
export type {
  OrderedAttribution,
  OrderedContainerToken,
  OrderedCounters,
  OrderedDocument,
  OrderedElementRecord,
  OrderedFieldPiece,
  OrderedFormatRecord,
  OrderedFormatWrite,
  OrderedStyleWrite,
  OrderedItem,
  OrderedItemAddress,
  OrderedItemKind,
  OrderedItemLabels,
  OrderedItemSummary,
  OrderedNoteProperties,
  OrderedNotePropertiesOp,
  OrderedParagraphOp,
  OrderedParagraphOpRecord,
  OrderedPartRevision,
  OrderedPropertyRecord,
  OrderedPropertyTarget,
  OrderedRevisionIds,
  OrderedStory,
  OrderedSuggestion,
  OrderedSuggestionStatus,
  OrderedUndoStep,
  OrderedUnitRange,
  OrderedWithin,
} from './model/types.ts';
export { OrderedRefusal, type OrderedRefusalReason } from './model/refusal.ts';
export { orderedAuthorName, type OrderedPerson } from './model/person.ts';
export {
  createOrderedDocument,
  createOrderedItem,
  addOrderedStory,
  itemAt as orderedItemAt,
  statusOf as orderedStatusOf,
  type OrderedItemInit,
} from './model/document.ts';
export {
  insertItems as insertOrderedItems,
  isolateUnits as isolateOrderedUnits,
  newSuggestion as newOrderedSuggestion,
  relabel as relabelOrderedItem,
  setStatus as setOrderedStatus,
  splitItem as splitOrderedItem,
} from './model/mutations.ts';
export {
  takeBack as takeBackOrdered,
  withUndoRecord as withOrderedUndoRecord,
  type OrderedUndoRecord,
  type OrderedUndoRecorded,
} from './model/undo-record.ts';
export {
  isOrderedStoryPart,
  loadOrderedDocument,
  loadOrderedPackage,
  orderedStoryParts,
  type OrderedLoadOptions,
} from './load/load.ts';
export {
  composeOrderedParagraph,
  composeOrderedPart,
  composeOrderedStory,
  orderedParagraphCount,
  orderedViewParagraph,
  type OrderedComposedParagraph,
  type OrderedComposedStory,
} from './compose/compose.ts';
export {
  composeOrderedPackage,
  orderedPackageBytes,
  writeOrderedPackage,
  type OrderedPackageCommentOptions,
  type OrderedPackageWriteOptions,
} from './compose/package.ts';
export {
  orderedPackageMedia,
  writeOrderedCheckpoint,
  type OrderedPackageMedia,
} from './compose/checkpoint.ts';
export type { OrderedResourceOptions } from './compose/resources.ts';
export {
  ORDERED_MEDIA_EXTENSIONS,
  type OrderedCustomXmlResource,
  type OrderedFileMedia,
  type OrderedFileStore,
  type OrderedListBullet,
  type OrderedListLevel,
  type OrderedListLevelResource,
  type OrderedListResource,
  type OrderedMediaReference,
  type OrderedMediaResource,
  type OrderedMediaType,
  type OrderedObjectPartResource,
  type OrderedRelationshipResource,
  type OrderedRelationshipType,
  type OrderedResource,
  type OrderedResourceBase,
  type OrderedResources,
  type OrderedSettingResource,
  type OrderedStoryResource,
  type OrderedStoryRole,
  type OrderedStyleResource,
  type OrderedStyleType,
} from './model/resources.ts';
export {
  ORDERED_RESOURCE_LIMITS,
  type OrderedChangeList,
  type OrderedChangeMedia,
  type OrderedChangeSetting,
  type OrderedChangeStyle,
} from './change/resources.ts';
export { orderedBlobHash, type OrderedBlobRefusalCode } from './protocol/blobs.ts';
export type {
  OrderedComposeOptions,
  OrderedRetainedPrompt,
  OrderedViewChar,
  OrderedViewParagraph,
} from './compose/context.ts';
export type { OrderedEditContext, OrderedReplacedFormat } from './edits/context.ts';
export {
  orderedPlaceAt,
  resolveOrderedPlace,
  type OrderedPlace,
  type OrderedResolvedPlace,
} from './edits/places.ts';
export { applyOrderedEdit, type OrderedEdit, type OrderedEditOutcome } from './edits/apply.ts';
export type { OrderedTableOp } from './edits/tables.ts';
export type { OrderedBreak, OrderedBreakClear } from './edits/breaks.ts';
export type { OrderedLinkTarget } from './edits/links.ts';
export type { OrderedPageField } from './edits/fields.ts';
export type { OrderedTocBookmark, OrderedTocRow } from './edits/toc.ts';
export { decideOrdered, orderedCards, type OrderedCard } from './decide.ts';
export { isOrderedItemHidden, orderedCardOf } from './card-of.ts';
export { orderedSuggestionsOfRevision, type OrderedRevisionName } from './client/ops/review.ts';
export { orderedReplacedSuggestions, type OrderedReplacedSuggestion } from './compose/replaced.ts';
export { orderedUnitTyper, type OrderedTyper } from './model/typers.ts';
export { ORDERED_CHANGE_KINDS, type OrderedGroupEditChange } from './change/kinds/registry.ts';
export type {
  OrderedInsertObjectChange,
  OrderedObjectPropertyChange,
  OrderedObjectResourceChange,
  OrderedPictureChange,
} from './change/kinds/pictures.ts';
export { ORDERED_OBJECT_LIMITS } from './change/kinds/pictures-read.ts';
export type { OrderedNewObject, OrderedObjectChangeOp } from './edits/pictures.ts';
export type { OrderedObjectOp, OrderedObjectRecord } from './model/objects.ts';
export type {
  OrderedComment,
  OrderedCommentMarker,
  OrderedCommentRange,
} from './model/comments.ts';
export { OrderedMedia } from './client/media.ts';
export type { OrderedMediaReader } from './client/media-reads.ts';
export {
  ORDERED_MEDIA_PREFIX,
  checkOrderedMediaPrefix,
  isOrderedMediaBytes,
} from './client/media-check.ts';
export type {
  OrderedHeldBlob,
  OrderedMediaUpload,
  OrderedMediaUploader,
  OrderedMediaUploadResult,
} from './client/media-uploads.ts';
export type {
  OrderedInsertLinkChange,
  OrderedLinkFieldChange,
  OrderedRemoveLinkChange,
} from './change/kinds/links-fields.ts';
export type {
  OrderedBookmarkChange,
  OrderedFieldChange,
  OrderedFieldCodeChange,
  OrderedFieldResultChange,
  OrderedInsertFieldChange,
  OrderedInsertTocChange,
  OrderedTocPagesChange,
  OrderedTocResultChange,
} from './change/kinds/fields.ts';
export type {
  OrderedControlBinding,
  OrderedControlChange,
  OrderedControlLock,
  OrderedControlProperties,
  OrderedControlPropertyChange,
  OrderedControlType,
  OrderedCustomControl,
  OrderedCustomPayload,
  OrderedDeleteControlChange,
  OrderedInsertControlChange,
  OrderedRemoveControlChange,
  OrderedRepeatingItemChange,
  OrderedBuildingBlockChange,
  OrderedWrapControlChange,
} from './change/kinds/controls.ts';
export type {
  OrderedFormFieldChange,
  OrderedFormTextChange,
} from './change/kinds/control-form-fields.ts';
export type { OrderedFormFieldOp } from './edits/form-fields.ts';
export type { OrderedTableChange } from './change/kinds/tables.ts';
export type {
  OrderedChangeFurnitureOptions,
  OrderedChangePageSetup,
  OrderedChangePictureWatermark,
  OrderedChangeTextWatermark,
  OrderedChangeWatermark,
  OrderedChangeWatermarkBox,
  OrderedChangeWatermarkSize,
  OrderedHeaderFooterChange,
  OrderedSectionBreakChange,
  OrderedSectionChange,
  OrderedSectionPropertyChange,
  OrderedWatermarkChange,
} from './change/kinds/sections.ts';
export type {
  OrderedFurnitureKind,
  OrderedFurnitureVariant,
  OrderedSectionOp,
  OrderedSectionRecord,
} from './model/sections.ts';
export type {
  OrderedConvertNoteChange,
  OrderedEquationChange,
  OrderedNoteChange,
  OrderedNotePropertiesChange,
  OrderedNoteSection,
} from './change/kinds/notes.ts';
export type {
  OrderedAttributionChange,
  OrderedDeleteParagraphChange,
  OrderedInsertFragmentChange,
  OrderedListLevelChange,
  OrderedNewListChange,
  OrderedNewListParagraph,
  OrderedPasteBlockChange,
  OrderedReplaceStoryChange,
  OrderedSettingChange,
} from './change/kinds/paste-blocks.ts';
export {
  ORDERED_FRAGMENT_LIMITS,
  type OrderedFragmentElement,
  type OrderedFragmentList,
  type OrderedFragmentNote,
  type OrderedFragmentPart,
  type OrderedFragmentPartRelationship,
} from './change/fragment.ts';
export type { OrderedFragmentPartType } from './change/fragment-names-objects.ts';
export {
  ORDERED_CHANGE_LIMITS,
  ORDERED_CHANGE_VERSION,
  isOrderedCommentChange,
  isOrderedEditChange,
  type OrderedChange,
  type OrderedChangeAddress,
  type OrderedChangeBase,
  type OrderedChangeKind,
  type OrderedChangeMarkPlace,
  type OrderedChangeMode,
  type OrderedChangePlace,
  type OrderedChangeProperty,
  type OrderedChangePropertyChild,
  type OrderedCommentChange,
  type OrderedCommentReplyChange,
  type OrderedCommentResolveChange,
  type OrderedCommentThreadChange,
  type OrderedControlValue,
  type OrderedControlValueChange,
  type OrderedCoreEditChange,
  type OrderedDecideChange,
  type OrderedDeleteChange,
  type OrderedDeleteColumnChange,
  type OrderedDeleteRowChange,
  type OrderedEditBase,
  type OrderedEditChange,
  type OrderedFormatChange,
  type OrderedMarkFormatChange,
  type OrderedBreakKind,
  type OrderedInsertBreakChange,
  type OrderedInsertChange,
  type OrderedInsertTabChange,
  type OrderedInsertColumnChange,
  type OrderedInsertNoteChange,
  type OrderedInsertRowChange,
  type OrderedInsertTableChange,
  type OrderedJoinChange,
  type OrderedLinkTargetChange,
  type OrderedMoveChange,
  type OrderedParagraphPropertyChange,
  type OrderedRetractChange,
  type OrderedRevertChange,
  type OrderedSeenCard,
  type OrderedSeenSuggestion,
  type OrderedServerFields,
  type OrderedSplitChange,
  type OrderedStampedChange,
  type OrderedTablePropertyChange,
} from './change/types.ts';
export { isOrderedPropertyAllowed, type OrderedPropertyScope } from './change/properties.ts';
export {
  ORDERED_NAMESPACE_PATTERN,
  isOrderedChangeId,
  isOrderedSuggestionId,
  orderedChangeId,
  orderedIdMinter,
  parseOrderedChangeId,
  type OrderedIdMinter,
  type OrderedMintedParagraph,
} from './change/ids.ts';
export type {
  OrderedChangeError,
  OrderedChangeErrorCode,
  OrderedChangeResult,
} from './change/read.ts';
export { applyOrderedChange, type OrderedChangeOutcome } from './change/apply.ts';
export { orderedCardSeen, orderedDecidedCards } from './change/apply-decide.ts';
export {
  validateOrderedBatch,
  validateOrderedChange,
  validateOrderedStampedChange,
  type OrderedValidateOptions,
} from './change/validate.ts';
export {
  orderedChangeTargets,
  type OrderedCaret,
  type OrderedTargetErrorCode,
  type OrderedTouched,
  type OrderedTouchedResult,
} from './change/targets.ts';
export {
  ORDERED_PROTOCOL_LIMITS,
  ORDERED_PROTOCOL_VERSION,
  type OrderedAckMessage,
  type OrderedClientChangesMessage,
  type OrderedClientDecideMessage,
  type OrderedClientMessage,
  type OrderedClientPresenceMessage,
  type OrderedEventMessage,
  type OrderedEventValue,
  type OrderedHelloMessage,
  type OrderedPresencePoint,
  type OrderedPresenceSelection,
  type OrderedPresenceState,
  type OrderedRefuseMessage,
  type OrderedReloadMessage,
  type OrderedReloadReason,
  type OrderedSavedMessage,
  type OrderedServerChangesMessage,
  type OrderedServerMessage,
  type OrderedServerPresenceMessage,
  type OrderedWelcomeMessage,
} from './protocol/messages.ts';
export {
  decodeOrderedClientMessage,
  decodeOrderedServerMessage,
  encodeOrderedClientMessage,
  encodeOrderedServerMessage,
  type OrderedDecodeOptions,
} from './protocol/codec.ts';
export {
  allowWritersPolicy,
  type OrderedActor,
  type OrderedPolicyCode,
  type OrderedPolicyInput,
  type OrderedPolicyReason,
  type OrderedPolicyVerdict,
  type OrderedRoomPolicy,
} from './room/policy.ts';
export { suggestionsOnlyPolicy } from './room/suggestions-only.ts';
export {
  OrderedRoom,
  OrderedRoomOpenError,
  orderedStampDate,
  type OrderedAck,
  type OrderedApplyChange,
  type OrderedApplyContext,
  type OrderedRoomOpenCode,
  type OrderedRoomOptions,
  type OrderedRoomRefusal,
  type OrderedRoomRefusalCode,
  type OrderedSubmitResult,
} from './room/room.ts';
export {
  admitOrderedPresence,
  type OrderedPresenceOptions,
  type OrderedPresenceRefusalCode,
  type OrderedPresenceVerdict,
} from './room/presence.ts';
export {
  createOrderedCollaboration,
  type OrderedCollaborationHandle,
  type OrderedCollaborationOptions,
  type OrderedCollaborationSession,
} from './client/create.ts';
export type { OrderedClientRefusalCode, OrderedClientRefusalReport } from './client/refusal.ts';
export {
  openOrderedCheckpoint,
  openOrderedWordFile,
  OrderedWordFileError,
  type OrderedWordFile,
  type OrderedWordFileOptions,
} from './client/open.ts';
export {
  orderedWebSocketTransport,
  type OrderedConnection,
  type OrderedConnectionEvents,
  type OrderedSocket,
  type OrderedTransport,
  type OrderedWebSocketOptions,
} from './client/transport.ts';
export { ORDERED_RECONNECT_DELAYS, type OrderedSchedule } from './client/link.ts';
export {
  OrderedLocalRoom,
  type OrderedLocalLogin,
  type OrderedLocalRoomOptions,
} from './client/local-room.ts';
