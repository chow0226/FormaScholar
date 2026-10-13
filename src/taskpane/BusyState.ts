export function nextBusyDepth(currentDepth: number, entering: boolean): number {
  return entering
    ? currentDepth + 1
    : Math.max(0, currentDepth - 1);
}

export function canStartCorrection(
  operationBusy: boolean,
  hasSummary: boolean,
  selectedFixCount: number
): boolean {
  return !operationBusy && hasSummary && selectedFixCount > 0;
}
