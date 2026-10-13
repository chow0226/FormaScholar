/** Returns true only when a delete confirmation still targets the current import. */
export function isDeleteConfirmationCurrent(
  pendingStrategyId: string | null,
  currentImportedStrategyId: string | null
): boolean {
  return Boolean(
    pendingStrategyId
    && currentImportedStrategyId
    && pendingStrategyId === currentImportedStrategyId
  );
}
